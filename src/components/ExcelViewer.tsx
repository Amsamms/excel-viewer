import React, { useEffect, useRef, useState } from 'react';
import * as XLSX from 'xlsx';

declare global {
  interface Window {
    luckysheet: any;
  }
}

interface ExcelViewerProps {
  file: File | null;
}

const ExcelViewer: React.FC<ExcelViewerProps> = ({ file }) => {
  const luckysheetRef = useRef<HTMLDivElement>(null);
  const [isLoading, setIsLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [debugInfo, setDebugInfo] = useState<string>('');

  useEffect(() => {
    if (file) {
      loadExcelFile(file);
    }
  }, [file]);

  const loadExcelFile = async (file: File) => {
    setIsLoading(true);
    setError(null);
    setDebugInfo('Starting file processing...');

    try {
      // Check if Luckysheet is available
      if (!window.luckysheet) {
        setError('Luckysheet library is not loaded. Please refresh the page.');
        setIsLoading(false);
        return;
      }

      const arrayBuffer = await file.arrayBuffer();
      setDebugInfo('File read successfully, parsing Excel...');
      
      const workbook = XLSX.read(arrayBuffer, { type: 'array' });
      setDebugInfo(`Found ${workbook.SheetNames.length} sheets: ${workbook.SheetNames.join(', ')}`);
      
      const sheets = workbook.SheetNames.map((sheetName, index) => {
        const worksheet = workbook.Sheets[sheetName];
        const data = XLSX.utils.sheet_to_json(worksheet, { 
          header: 1, 
          defval: '',
          raw: false 
        });
        
        setDebugInfo(`Processing sheet "${sheetName}" with ${data.length} rows`);
        
        return {
          name: sheetName,
          data: convertToLuckysheetData(data as any[][]),
          config: {},
          status: index === 0 ? 1 : 0, // Only first sheet active
          order: index,
          hide: 0,
          row: Math.max(data.length, 20),
          column: Math.max(getMaxColumns(data as any[][]), 10),
          defaultRowHeight: 19,
          defaultColWidth: 73,
        };
      });

      setDebugInfo('Initializing Luckysheet...');
      initializeLuckysheet(sheets);
      
    } catch (err) {
      console.error('Error loading Excel file:', err);
      setError('Error loading Excel file: ' + (err as Error).message);
    } finally {
      setIsLoading(false);
    }
  };

  const getMaxColumns = (data: any[][]): number => {
    return Math.max(...data.map(row => row.length), 0);
  };

  const convertToLuckysheetData = (data: any[][]): any[] => {
    const result: any[] = [];
    
    data.forEach((row, rowIndex) => {
      row.forEach((cell, colIndex) => {
        if (cell !== null && cell !== undefined && cell !== '') {
          result.push({
            r: rowIndex,
            c: colIndex,
            v: {
              v: cell,
              ct: { fa: 'General', t: 'g' },
              m: String(cell),
              bg: null,
              bl: 0,
              it: 0,
              ff: 'Arial',
              fs: 10,
              fc: '#000000',
              ht: 1,
              vt: 1
            }
          });
        }
      });
    });
    
    return result;
  };

  const initializeLuckysheet = (sheets: any[]) => {
    if (luckysheetRef.current && window.luckysheet) {
      // Clear previous instance
      luckysheetRef.current.innerHTML = '';
      
      // Wait a bit for DOM to update
      setTimeout(() => {
        try {
          window.luckysheet.create({
            container: luckysheetRef.current,
            showtoolbar: true,
            showinfobar: true,
            showsheetbar: true,
            showstatisticBar: true,
            allowCopy: false,
            allowEdit: true,
            allowUpdate: true,
            showConfigWindowResize: true,
            enableAddRow: true,
            enableAddCol: true,
            data: sheets,
            title: file?.name || 'Excel Viewer',
            lang: 'en',
          });
          setDebugInfo('Luckysheet initialized successfully!');
        } catch (err) {
          console.error('Luckysheet initialization error:', err);
          setError('Failed to initialize spreadsheet viewer: ' + (err as Error).message);
        }
      }, 100);
    } else {
      setError('Luckysheet container or library not available');
    }
  };

  if (isLoading) {
    return (
      <div className="flex flex-col items-center justify-center h-96">
        <div className="text-xl mb-4">Loading Excel file...</div>
        <div className="text-sm text-gray-600">{debugInfo}</div>
      </div>
    );
  }

  if (error) {
    return (
      <div className="flex flex-col items-center justify-center h-96">
        <div className="text-red-500 text-xl mb-4">{error}</div>
        <div className="text-sm text-gray-600">{debugInfo}</div>
        <button 
          onClick={() => window.location.reload()} 
          className="mt-4 px-4 py-2 bg-blue-500 text-white rounded hover:bg-blue-600"
        >
          Refresh Page
        </button>
      </div>
    );
  }

  return (
    <div className="w-full h-full">
      <div className="text-sm text-gray-600 mb-2">{debugInfo}</div>
      <div
        ref={luckysheetRef}
        style={{ width: '100%', height: '600px' }}
        className="border border-gray-300 rounded-lg"
      />
    </div>
  );
};

export default ExcelViewer;