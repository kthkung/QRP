import { NextRequest, NextResponse } from 'next/server';

interface TextRecord {
  text: string;
  x: number;
  y: number;
}

// EMF Record Types
const EMR_EXTTEXTOUTW = 0x54;
const EMR_EXTTEXTOUTA = 0x53;
const EMR_EOF = 0x0E;

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
    const groupedRows = groupTextByPosition(extractedData);

    // Return preview data (limit to 100 rows for preview)
    const previewRows = groupedRows.slice(0, 100);

    return NextResponse.json({
      rows: previewRows,
      totalRows: groupedRows.length,
    });
  } catch (error) {
    console.error('Preview error:', error);
    return NextResponse.json({ error: 'Internal server error' }, { status: 500 });
  }
}

function extractEMFTextRecords(buffer: Buffer): TextRecord[] {
  const records: TextRecord[] = [];
  let offset = 0;
  
  // Find EMF signature " EMF" in the file
  let emfStart = -1;
  
  for (let i = 0; i < buffer.length - 4; i++) {
    if (buffer[i] === 0x20 && buffer[i+1] === 0x45 && buffer[i+2] === 0x4D && buffer[i+3] === 0x46) {
      emfStart = i - 40;
      if (emfStart < 0) emfStart = 0;
      break;
    }
  }
  
  if (emfStart === -1) {
    return extractUTF16Strings(buffer);
  }
  
  offset = emfStart;
  
  while (offset + 8 <= buffer.length) {
    const recordType = buffer.readUInt32LE(offset);
    const recordSize = buffer.readUInt32LE(offset + 4);
    
    if (recordSize < 8 || recordSize > buffer.length - offset) {
      break;
    }
    
    if (recordType === EMR_EXTTEXTOUTW && recordSize > 76) {
      try {
        const textRecord = parseExtTextOutW(buffer, offset);
        if (textRecord && textRecord.text.trim().length > 0) {
          records.push(textRecord);
        }
      } catch (e) {}
    }
    
    if (recordType === EMR_EXTTEXTOUTA && recordSize > 76) {
      try {
        const textRecord = parseExtTextOutA(buffer, offset);
        if (textRecord && textRecord.text.trim().length > 0) {
          records.push(textRecord);
        }
      } catch (e) {}
    }
    
    if (recordType === EMR_EOF) {
      const nextEmfStart = findNextEMF(buffer, offset + recordSize);
      if (nextEmfStart > 0) {
        offset = nextEmfStart;
        continue;
      }
      break;
    }
    
    offset += recordSize;
  }
  
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
  let text = '';
  for (let i = 0; i < textBuffer.length; i++) {
    const b = textBuffer[i];
    if (b >= 0xA1 && b <= 0xFB) {
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

function extractUTF16Strings(buffer: Buffer): TextRecord[] {
  const records: TextRecord[] = [];
  let i = 0;
  
  while (i < buffer.length - 4) {
    if (buffer[i+1] === 0x0E && buffer[i] >= 0x01 && buffer[i] <= 0x7F) {
      let start = i;
      let text = '';
      
      while (start > 0 && isValidUTF16Char(buffer, start - 2)) {
        start -= 2;
      }
      
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
  if (code >= 0x20 && code <= 0x7E) return true;
  if (code >= 0x0E01 && code <= 0x0E5B) return true;
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
  
  if (trimmed.length <= 1) return true;
  
  return false;
}

function groupTextByPosition(records: TextRecord[]): string[][] {
  if (records.length === 0) return [];
  
  const sorted = [...records].sort((a, b) => {
    const yDiff = a.y - b.y;
    if (Math.abs(yDiff) > 10) return yDiff;
    return a.x - b.x;
  });
  
  const rows: string[][] = [];
  let currentRow: TextRecord[] = [];
  let lastY = sorted[0]?.y ?? 0;
  
  for (const record of sorted) {
    if (Math.abs(record.y - lastY) > 15) {
      if (currentRow.length > 0) {
        currentRow.sort((a, b) => a.x - b.x);
        rows.push(currentRow.map(r => r.text));
      }
      currentRow = [];
      lastY = record.y;
    }
    currentRow.push(record);
  }
  
  if (currentRow.length > 0) {
    currentRow.sort((a, b) => a.x - b.x);
    rows.push(currentRow.map(r => r.text));
  }
  
  return rows;
}
