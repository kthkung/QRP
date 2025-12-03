import { NextRequest, NextResponse } from 'next/server';
import * as XLSX from 'xlsx';

export async function POST(request: NextRequest) {
  try {
    const formData = await request.formData();
    const file = formData.get('file') as File | null;

    if (!file) {
      return NextResponse.json({ error: 'No file uploaded' }, { status: 400 });
    }

    const arrayBuffer = await file.arrayBuffer();
    const buffer = Buffer.from(arrayBuffer);

    // Extract text from EMF records in QRP file
    const extractedData = extractEMFTextRecords(buffer);

    // Prepare data for Excel - organize into rows
    const rows: string[][] = [];
    rows.push(['ข้อมูลที่ดึงได้จากไฟล์ QRP']);
    rows.push([]); // Empty row

    // Group text by Y position to form rows
    const groupedRows = groupTextByPosition(extractedData);
    groupedRows.forEach(row => {
      rows.push(row);
    });

    // Create Excel
    const workbook = XLSX.utils.book_new();
    const worksheet = XLSX.utils.aoa_to_sheet(rows);
    
    // Auto-size columns
    const colWidths = rows.reduce((widths, row) => {
      row.forEach((cell, i) => {
        const len = cell ? cell.toString().length : 0;
        widths[i] = Math.max(widths[i] || 10, len + 2);
      });
      return widths;
    }, [] as number[]);
    worksheet['!cols'] = colWidths.map(w => ({ wch: Math.min(w, 50) }));

    XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');

    const excelBuffer = XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' });

    const filename = file.name.replace(/\.qrp$/i, '');
    const encodedFilename = encodeURIComponent(filename + '.xlsx');

    return new NextResponse(excelBuffer, {
      status: 200,
      headers: {
        'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'Content-Disposition': `attachment; filename="converted.xlsx"; filename*=UTF-8''${encodedFilename}`,
      },
    });
  } catch (error) {
    console.error('Conversion error:', error);
    return NextResponse.json({ error: 'Internal server error' }, { status: 500 });
  }
}

interface TextRecord {
  text: string;
  x: number;
  y: number;
}

// EMF Record Types
const EMR_HEADER = 0x01;
const EMR_EXTTEXTOUTW = 0x54; // Extended text output (Unicode)
const EMR_EXTTEXTOUTA = 0x53; // Extended text output (ANSI)
const EMR_EOF = 0x0E;

function extractEMFTextRecords(buffer: Buffer): TextRecord[] {
  const records: TextRecord[] = [];
  let offset = 0;
  
  // Find EMF signature " EMF" in the file
  const emfSignature = Buffer.from([0x20, 0x45, 0x4D, 0x46]); // " EMF"
  let emfStart = -1;
  
  for (let i = 0; i < buffer.length - 4; i++) {
    if (buffer[i] === 0x20 && buffer[i+1] === 0x45 && buffer[i+2] === 0x4D && buffer[i+3] === 0x46) {
      // EMF header starts 40 bytes before " EMF" signature
      emfStart = i - 40;
      if (emfStart < 0) emfStart = 0;
      break;
    }
  }
  
  if (emfStart === -1) {
    // No EMF found, try alternate extraction
    return extractUTF16Strings(buffer);
  }
  
  offset = emfStart;
  
  // Parse EMF records
  while (offset + 8 <= buffer.length) {
    const recordType = buffer.readUInt32LE(offset);
    const recordSize = buffer.readUInt32LE(offset + 4);
    
    if (recordSize < 8 || recordSize > buffer.length - offset) {
      // Invalid record, try to find next valid record or use fallback
      break;
    }
    
    // EMR_EXTTEXTOUTW (0x54) - Unicode text
    if (recordType === EMR_EXTTEXTOUTW && recordSize > 76) {
      try {
        const textRecord = parseExtTextOutW(buffer, offset);
        if (textRecord && textRecord.text.trim().length > 0) {
          records.push(textRecord);
        }
      } catch (e) {
        // Skip invalid record
      }
    }
    
    // EMR_EXTTEXTOUTA (0x53) - ANSI text  
    if (recordType === EMR_EXTTEXTOUTA && recordSize > 76) {
      try {
        const textRecord = parseExtTextOutA(buffer, offset);
        if (textRecord && textRecord.text.trim().length > 0) {
          records.push(textRecord);
        }
      } catch (e) {
        // Skip invalid record
      }
    }
    
    if (recordType === EMR_EOF) {
      // End of EMF, check if there's more EMF data after
      const nextEmfStart = findNextEMF(buffer, offset + recordSize);
      if (nextEmfStart > 0) {
        offset = nextEmfStart;
        continue;
      }
      break;
    }
    
    offset += recordSize;
  }
  
  // If no records found from EMF parsing, use fallback
  if (records.length === 0) {
    return extractUTF16Strings(buffer);
  }
  
  return records;
}

function findNextEMF(buffer: Buffer, startOffset: number): number {
  for (let i = startOffset; i < buffer.length - 4; i++) {
    if (buffer[i] === 0x20 && buffer[i+1] === 0x45 && buffer[i+2] === 0x4D && buffer[i+3] === 0x46) {
      const emfStart = i - 40;
      return emfStart > startOffset ? emfStart : -1;
    }
  }
  return -1;
}

function parseExtTextOutW(buffer: Buffer, offset: number): TextRecord | null {
  // EMR_EXTTEXTOUTW structure:
  // 0-3: Type (0x54)
  // 4-7: Size
  // 8-23: Bounds (RECTL)
  // 24-27: iGraphicsMode
  // 28-31: exScale
  // 32-35: eyScale
  // 36-39: Reference X
  // 40-43: Reference Y
  // 44-47: nChars (number of characters)
  // 48-51: offString (offset to string from start of record)
  // 52-55: Options
  // ... more fields
  // String data at offString
  
  const recordSize = buffer.readUInt32LE(offset + 4);
  const refX = buffer.readInt32LE(offset + 36);
  const refY = buffer.readInt32LE(offset + 40);
  const nChars = buffer.readUInt32LE(offset + 44);
  const offString = buffer.readUInt32LE(offset + 48);
  
  if (nChars === 0 || nChars > 1000 || offString + nChars * 2 > recordSize) {
    return null;
  }
  
  const stringStart = offset + offString;
  const stringEnd = stringStart + nChars * 2;
  
  if (stringEnd > buffer.length) {
    return null;
  }
  
  const textBuffer = buffer.subarray(stringStart, stringEnd);
  const text = textBuffer.toString('utf16le');
  
  // Filter out font names and garbage
  if (isIgnoredText(text)) {
    return null;
  }
  
  return { text: text.trim(), x: refX, y: refY };
}

function parseExtTextOutA(buffer: Buffer, offset: number): TextRecord | null {
  const recordSize = buffer.readUInt32LE(offset + 4);
  const refX = buffer.readInt32LE(offset + 36);
  const refY = buffer.readInt32LE(offset + 40);
  const nChars = buffer.readUInt32LE(offset + 44);
  const offString = buffer.readUInt32LE(offset + 48);
  
  if (nChars === 0 || nChars > 1000 || offString + nChars > recordSize) {
    return null;
  }
  
  const stringStart = offset + offString;
  const stringEnd = stringStart + nChars;
  
  if (stringEnd > buffer.length) {
    return null;
  }
  
  const textBuffer = buffer.subarray(stringStart, stringEnd);
  // Try Thai encoding (Windows-874/TIS-620)
  let text = '';
  for (let i = 0; i < textBuffer.length; i++) {
    const b = textBuffer[i];
    if (b >= 0xA1 && b <= 0xFB) {
      // Thai character range in Windows-874
      text += String.fromCharCode(0x0E00 + (b - 0xA0));
    } else if (b >= 0x20 && b <= 0x7E) {
      text += String.fromCharCode(b);
    }
  }
  
  if (isIgnoredText(text)) {
    return null;
  }
  
  return { text: text.trim(), x: refX, y: refY };
}

// Fallback: Extract UTF-16 strings directly from buffer
function extractUTF16Strings(buffer: Buffer): TextRecord[] {
  const records: TextRecord[] = [];
  let i = 0;
  
  while (i < buffer.length - 4) {
    // Look for pattern: [length 4 bytes] [UTF-16LE chars with Thai/ASCII]
    // Or look for 0x54 0x00 0x00 0x00 pattern (EMR_EXTTEXTOUTW marker)
    
    // Check for direct Thai Unicode pattern (0x0Exx range)
    if (buffer[i+1] === 0x0E && buffer[i] >= 0x01 && buffer[i] <= 0x7F) {
      // Found Thai character, try to extract the string
      let start = i;
      let text = '';
      
      // Look backwards to find start of string
      while (start > 0 && isValidUTF16Char(buffer, start - 2)) {
        start -= 2;
      }
      
      // Read the string
      let pos = start;
      while (pos < buffer.length - 1) {
        const code = buffer.readUInt16LE(pos);
        if (code === 0) break;
        if (!isValidChar(code)) break;
        text += String.fromCharCode(code);
        pos += 2;
      }
      
      if (text.trim().length > 0 && !isIgnoredText(text)) {
        records.push({ text: text.trim(), x: 0, y: records.length * 20 });
      }
      
      i = pos + 2;
      continue;
    }
    
    i++;
  }
  
  // Deduplicate
  const seen = new Set<string>();
  return records.filter(r => {
    if (seen.has(r.text)) return false;
    seen.add(r.text);
    return true;
  });
}

function isValidUTF16Char(buffer: Buffer, offset: number): boolean {
  if (offset < 0 || offset + 1 >= buffer.length) return false;
  const code = buffer.readUInt16LE(offset);
  return isValidChar(code);
}

function isValidChar(code: number): boolean {
  // ASCII printable
  if (code >= 0x20 && code <= 0x7E) return true;
  // Thai Unicode range
  if (code >= 0x0E01 && code <= 0x0E5B) return true;
  // Common punctuation and numbers
  if (code === 0x0A || code === 0x0D || code === 0x09) return true;
  return false;
}

function isIgnoredText(text: string): boolean {
  const ignored = [
    'Arial', 'Times New Roman', 'Courier New', 'Tahoma', 'Verdana',
    'Angsana New', 'AngsanaUPC', 'CordiaUPC', 'Cordia New', 'Browallia New',
    'BrowalliaUPC', 'Standard', 'Text', 'Page', 'Title', 'EucrosiaUPC',
    'FreesiaUPC', 'IrisUPC', 'JasmineUPC', 'KodchiangUPC', 'LilyUPC'
  ];
  
  const trimmed = text.trim();
  if (ignored.some(ig => trimmed === ig || trimmed.toLowerCase() === ig.toLowerCase())) {
    return true;
  }
  
  // Ignore very short strings that are just numbers or symbols
  if (trimmed.length <= 1) return true;
  
  return false;
}

function groupTextByPosition(records: TextRecord[]): string[][] {
  if (records.length === 0) return [];
  
  // Sort by Y position first, then X position
  const sorted = [...records].sort((a, b) => {
    const yDiff = a.y - b.y;
    if (Math.abs(yDiff) > 10) return yDiff; // Different rows
    return a.x - b.x; // Same row, sort by X
  });
  
  const rows: string[][] = [];
  let currentRow: TextRecord[] = [];
  let lastY = sorted[0]?.y ?? 0;
  
  for (const record of sorted) {
    // If Y position differs significantly, start new row
    if (Math.abs(record.y - lastY) > 15) {
      if (currentRow.length > 0) {
        // Sort current row by X and extract text
        currentRow.sort((a, b) => a.x - b.x);
        rows.push(currentRow.map(r => r.text));
      }
      currentRow = [];
      lastY = record.y;
    }
    currentRow.push(record);
  }
  
  // Don't forget the last row
  if (currentRow.length > 0) {
    currentRow.sort((a, b) => a.x - b.x);
    rows.push(currentRow.map(r => r.text));
  }
  
  return rows;
}
