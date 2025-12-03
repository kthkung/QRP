'use client';

import { useState } from 'react';

interface PreviewData {
  rows: string[][];
  totalRows: number;
}

export default function Home() {
  const [file, setFile] = useState<File | null>(null);
  const [status, setStatus] = useState<'idle' | 'uploading' | 'preview' | 'success' | 'error'>('idle');
  const [downloadUrl, setDownloadUrl] = useState<string | null>(null);
  const [errorMessage, setErrorMessage] = useState<string | null>(null);
  const [previewData, setPreviewData] = useState<PreviewData | null>(null);

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files[0]) {
      setFile(e.target.files[0]);
      setStatus('idle');
      setDownloadUrl(null);
      setErrorMessage(null);
      setPreviewData(null);
    }
  };

  const handlePreview = async () => {
    if (!file) return;

    setStatus('uploading');
    const formData = new FormData();
    formData.append('file', file);

    try {
      const response = await fetch('/api/preview', {
        method: 'POST',
        body: formData,
      });

      if (!response.ok) {
        throw new Error('Preview failed');
      }

      const data = await response.json();
      setPreviewData(data);
      setStatus('preview');
    } catch (error) {
      console.error(error);
      setStatus('error');
      setErrorMessage('Failed to preview file. Please try again.');
    }
  };

  const handleConvert = async () => {
    if (!file) return;

    setStatus('uploading');
    const formData = new FormData();
    formData.append('file', file);

    try {
      const response = await fetch('/api/convert', {
        method: 'POST',
        body: formData,
      });

      if (!response.ok) {
        throw new Error('Conversion failed');
      }

      const blob = await response.blob();
      const url = window.URL.createObjectURL(blob);
      setDownloadUrl(url);
      setStatus('success');
    } catch (error) {
      console.error(error);
      setStatus('error');
      setErrorMessage('Failed to convert file. Please try again.');
    }
  };

  return (
    <main className="flex min-h-screen flex-col items-center justify-center p-8 bg-gray-900">
      <div className="z-10 w-full max-w-6xl items-center justify-between font-mono text-sm">
        <h1 className="text-4xl font-bold mb-8 text-white text-center">QRP to Excel Converter For P' Pete</h1>
        
        <div className="bg-gray-800 p-8 rounded-xl shadow-lg w-full border border-gray-700">
          <div className="mb-6">
            <label className="block text-gray-200 text-sm font-bold mb-2" htmlFor="file-upload">
              Upload .qrp File
            </label>
            <input
              id="file-upload"
              type="file"
              accept=".qrp"
              onChange={handleFileChange}
              className="block w-full text-sm text-gray-400
                file:mr-4 file:py-2 file:px-4
                file:rounded-full file:border-0
                file:text-sm file:font-semibold
                file:bg-blue-900 file:text-blue-200
                hover:file:bg-blue-800
              "
            />
          </div>

          <div className="flex gap-4">
            <button
              onClick={handlePreview}
              disabled={!file || status === 'uploading'}
              className={`flex-1 py-3 px-4 rounded-lg font-bold text-white transition-colors ${
                !file || status === 'uploading'
                  ? 'bg-gray-600 cursor-not-allowed'
                  : 'bg-purple-600 hover:bg-purple-700'
              }`}
            >
              {status === 'uploading' ? 'Processing...' : 'Preview Data'}
            </button>

            <button
              onClick={handleConvert}
              disabled={!file || status === 'uploading'}
              className={`flex-1 py-3 px-4 rounded-lg font-bold text-white transition-colors ${
                !file || status === 'uploading'
                  ? 'bg-gray-600 cursor-not-allowed'
                  : 'bg-blue-600 hover:bg-blue-700'
              }`}
            >
              {status === 'uploading' ? 'Converting...' : 'Convert to Excel'}
            </button>
          </div>

          {status === 'error' && (
            <div className="mt-4 p-3 bg-red-900 text-red-200 rounded-lg text-center">
              {errorMessage}
            </div>
          )}

          {/* Preview Section */}
          {status === 'preview' && previewData && (
            <div className="mt-6">
              <div className="flex justify-between items-center mb-3">
                <h2 className="text-lg font-semibold text-white">
                  Preview ({previewData.totalRows} rows)
                </h2>
                <button
                  onClick={handleConvert}
                  className="bg-green-600 hover:bg-green-700 text-white font-bold py-2 px-4 rounded-lg transition-colors"
                >
                  Download Excel
                </button>
              </div>
              <div className="overflow-x-auto max-h-96 overflow-y-auto border border-gray-600 rounded-lg">
                <table className="min-w-full bg-gray-900">
                  <tbody>
                    {previewData.rows.map((row, rowIndex) => (
                      <tr key={rowIndex} className={rowIndex % 2 === 0 ? 'bg-gray-800' : 'bg-gray-850'}>
                        {row.map((cell, cellIndex) => (
                          <td
                            key={cellIndex}
                            className="px-4 py-2 text-sm text-gray-200 border-b border-gray-700 whitespace-nowrap"
                          >
                            {cell}
                          </td>
                        ))}
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
              {previewData.totalRows > previewData.rows.length && (
                <p className="text-gray-400 text-sm mt-2 text-center">
                  Showing first {previewData.rows.length} of {previewData.totalRows} rows
                </p>
              )}
            </div>
          )}

          {status === 'success' && downloadUrl && (
            <div className="mt-6 text-center">
              <p className="text-green-400 font-semibold mb-2">Conversion Successful!</p>
              <a
                href={downloadUrl}
                download={`${file?.name.replace('.qrp', '').replace('.QRP', '')}.xlsx`}
                className="inline-block bg-green-500 hover:bg-green-600 text-white font-bold py-2 px-6 rounded-lg transition-colors"
              >
                Download Excel
              </a>
            </div>
          )}
          
          <div className="mt-8 text-xs text-gray-400 text-center">
            <p>Note: .qrp is a proprietary format (QuickReport/EMF).</p>
            <p>This tool extracts text data from EMF records in the file.</p>
          </div>
        </div>
      </div>
    </main>
  );
}
