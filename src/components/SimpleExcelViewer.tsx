import React, { useState, useEffect } from 'react';
import * as XLSX from 'xlsx';

interface SimpleExcelViewerProps {
  file: File | null;
}

const SimpleExcelViewer: React.FC<SimpleExcelViewerProps> = ({ file }) => {
  const [sheets, setSheets] = useState<any[]>([]);
  const [activeSheetIndex, setActiveSheetIndex] = useState(0);
  const [isLoading, setIsLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [showFormulas, setShowFormulas] = useState(false);

  useEffect(() => {
    if (file) {
      loadExcelFile(file);
    }
  }, [file]);

  const loadExcelFile = async (file: File) => {
    setIsLoading(true);
    setError(null);

    try {
      const arrayBuffer = await file.arrayBuffer();
      const workbook = XLSX.read(arrayBuffer, { 
        type: 'array',
        cellStyles: true,
        cellNF: true,
        cellHTML: true
      });
      
      const sheetsData = workbook.SheetNames.map((sheetName) => {
        const worksheet = workbook.Sheets[sheetName];
        
        // Get the range to understand the dimensions
        const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1:A1');
        const rows = range.e.r + 1;
        const cols = range.e.c + 1;
        
        // Create a matrix to store cell data with formatting
        const cellMatrix: any[][] = [];
        
        for (let row = 0; row < rows; row++) {
          cellMatrix[row] = [];
          for (let col = 0; col < cols; col++) {
            const cellAddress = XLSX.utils.encode_cell({ r: row, c: col });
            const cell = worksheet[cellAddress];
            
            if (cell) {
              const backgroundColor = getCellBackgroundColor(cell);
              const textColor = getCellTextColor(cell);
              const isBold = isCellBold(cell);
              
              cellMatrix[row][col] = {
                value: cell.w || cell.v || '', // formatted value or raw value
                formula: cell.f || null, // formula if exists
                type: cell.t || 'n', // cell type
                style: cell.s || null, // style information
                backgroundColor: backgroundColor,
                textColor: textColor,
                isBold: isBold,
                isItalic: isCellItalic(cell),
                fontSize: getCellFontSize(cell),
                fontFamily: getCellFontFamily(cell),
                alignment: getCellAlignment(cell),
                border: getCellBorder(cell),
                originalCell: cell, // store original cell for debugging
                debugInfo: `BG: ${backgroundColor}, Text: ${textColor}, Bold: ${isBold}, Style: ${JSON.stringify(cell.s)}`
              };
              
              // Log first few cells for debugging
              if (row < 3 && col < 3) {
                console.log(`Cell [${row},${col}]:`, {
                  value: cell.w || cell.v,
                  style: cell.s,
                  backgroundColor,
                  textColor,
                  isBold,
                  rawCell: cell
                });
              }
            } else {
              cellMatrix[row][col] = {
                value: '',
                formula: null,
                type: 'n',
                style: null,
                backgroundColor: null,
                textColor: null,
                isBold: false,
                isItalic: false,
                fontSize: null,
                fontFamily: null,
                alignment: null,
                border: null,
                originalCell: null
              };
            }
          }
        }
        
        return {
          name: sheetName,
          data: cellMatrix,
          rawData: XLSX.utils.sheet_to_json(worksheet, { 
            header: 1, 
            defval: '',
            raw: false 
          }),
        };
      });

      setSheets(sheetsData);
      setActiveSheetIndex(0);
    } catch (err) {
      setError('Error loading Excel file: ' + (err as Error).message);
    } finally {
      setIsLoading(false);
    }
  };

  // Helper functions to extract formatting
  const getCellBackgroundColor = (cell: any) => {
    if (cell.s && cell.s.fill && cell.s.fill.bgColor) {
      const color = cell.s.fill.bgColor;
      if (color.rgb) {
        return `#${color.rgb.slice(2)}`; // Remove alpha channel
      }
      if (color.indexed) {
        // Map common indexed colors
        const indexedColors: { [key: number]: string } = {
          64: '#000000', // black
          9: '#ffffff',  // white
          10: '#ff0000', // red
          11: '#00ff00', // green
          12: '#0000ff', // blue
          13: '#ffff00', // yellow
          14: '#ff00ff', // magenta
          15: '#00ffff', // cyan
          43: '#92d050', // light green
          44: '#00b0f0', // light blue
          45: '#0070c0', // blue
          46: '#002060', // dark blue
        };
        return indexedColors[color.indexed] || null;
      }
    }
    return null;
  };

  const getCellTextColor = (cell: any) => {
    if (cell.s && cell.s.font && cell.s.font.color) {
      const color = cell.s.font.color;
      if (color.rgb) {
        return `#${color.rgb.slice(2)}`;
      }
    }
    return null;
  };

  const isCellBold = (cell: any) => {
    return cell.s && cell.s.font && cell.s.font.bold;
  };

  const isCellItalic = (cell: any) => {
    return cell.s && cell.s.font && cell.s.font.italic;
  };

  const getCellFontSize = (cell: any) => {
    return cell.s && cell.s.font && cell.s.font.sz ? `${cell.s.font.sz}px` : null;
  };

  const getCellFontFamily = (cell: any) => {
    return cell.s && cell.s.font && cell.s.font.name ? cell.s.font.name : null;
  };

  const getCellAlignment = (cell: any) => {
    if (cell.s && cell.s.alignment) {
      return {
        horizontal: cell.s.alignment.horizontal || 'left',
        vertical: cell.s.alignment.vertical || 'middle'
      };
    }
    return null;
  };

  const getCellBorder = (cell: any) => {
    if (cell.s && cell.s.border) {
      return cell.s.border;
    }
    return null;
  };

  const handleCellChange = (rowIndex: number, colIndex: number, value: string) => {
    const newSheets = [...sheets];
    if (!newSheets[activeSheetIndex].data[rowIndex]) {
      newSheets[activeSheetIndex].data[rowIndex] = [];
    }
    
    // Preserve the formatting while updating the value
    const existingCell = newSheets[activeSheetIndex].data[rowIndex][colIndex];
    if (existingCell && typeof existingCell === 'object') {
      newSheets[activeSheetIndex].data[rowIndex][colIndex] = {
        ...existingCell,
        value: value
      };
    } else {
      newSheets[activeSheetIndex].data[rowIndex][colIndex] = {
        value: value,
        formula: null,
        type: 'n',
        style: null,
        backgroundColor: null,
        textColor: null,
        isBold: false,
        isItalic: false,
        fontSize: null,
        fontFamily: null,
        alignment: null,
        border: null,
        originalCell: null
      };
    }
    
    setSheets(newSheets);
  };

  const getColumnLabel = (index: number): string => {
    let label = '';
    let num = index;
    while (num >= 0) {
      label = String.fromCharCode(65 + (num % 26)) + label;
      num = Math.floor(num / 26) - 1;
    }
    return label;
  };

  if (isLoading) {
    return (
      <div className="flex items-center justify-center h-96">
        <div className="text-xl">Loading Excel file...</div>
      </div>
    );
  }

  if (error) {
    return (
      <div className="flex items-center justify-center h-96">
        <div className="text-red-500 text-xl">{error}</div>
      </div>
    );
  }

  if (sheets.length === 0) {
    return null;
  }

  const activeSheet = sheets[activeSheetIndex];
  const maxRows = Math.max(activeSheet.data.length, 20);
  const maxCols = Math.max(
    ...activeSheet.data.map((row: any[]) => row.length),
    10
  );

  const getCellStyle = (cell: any) => {
    const style: React.CSSProperties = {};
    
    if (cell && typeof cell === 'object') {
      if (cell.backgroundColor) {
        style.backgroundColor = cell.backgroundColor;
      }
      if (cell.textColor) {
        style.color = cell.textColor;
      }
      if (cell.isBold) {
        style.fontWeight = 'bold';
      }
      if (cell.isItalic) {
        style.fontStyle = 'italic';
      }
      if (cell.fontSize) {
        style.fontSize = cell.fontSize;
      }
      if (cell.fontFamily) {
        style.fontFamily = cell.fontFamily;
      }
      if (cell.alignment) {
        style.textAlign = cell.alignment.horizontal;
        style.verticalAlign = cell.alignment.vertical;
      }
    }
    
    return style;
  };

  const getCellDisplayValue = (cell: any) => {
    if (!cell || typeof cell !== 'object') {
      return String(cell || '');
    }
    
    if (showFormulas && cell.formula) {
      return `=${cell.formula}`;
    }
    
    return String(cell.value || '');
  };

  const getCellTitle = (cell: any) => {
    if (!cell || typeof cell !== 'object') {
      return '';
    }
    
    let title = '';
    if (cell.formula) {
      title += `Formula: =${cell.formula}\n`;
    }
    if (cell.value !== undefined) {
      title += `Value: ${cell.value}\n`;
    }
    if (cell.type) {
      title += `Type: ${cell.type}`;
    }
    
    return title.trim();
  };

  return (
    <div className="w-full">
      {/* Controls */}
      <div className="flex items-center justify-between mb-4">
        {/* Sheet tabs */}
        {sheets.length > 1 && (
          <div className="flex border-b border-gray-300">
            {sheets.map((sheet, index) => (
              <button
                key={index}
                onClick={() => setActiveSheetIndex(index)}
                className={`px-4 py-2 text-sm font-medium ${
                  index === activeSheetIndex
                    ? 'text-blue-600 border-b-2 border-blue-600'
                    : 'text-gray-500 hover:text-gray-700'
                }`}
              >
                {sheet.name}
              </button>
            ))}
          </div>
        )}
        
        {/* Formula toggle */}
        <button
          onClick={() => setShowFormulas(!showFormulas)}
          className={`px-3 py-1 text-xs font-medium rounded ${
            showFormulas
              ? 'bg-blue-100 text-blue-800 border border-blue-300'
              : 'bg-gray-100 text-gray-700 border border-gray-300'
          }`}
        >
          {showFormulas ? 'Hide Formulas' : 'Show Formulas'}
        </button>
      </div>

      {/* Spreadsheet grid */}
      <div className="overflow-auto border border-gray-300 rounded-lg">
        <table className="min-w-full">
          <thead>
            <tr className="bg-gray-100">
              <th className="w-12 p-2 text-xs font-medium text-gray-500 border-r border-gray-300">
                #
              </th>
              {Array.from({ length: maxCols }, (_, colIndex) => (
                <th
                  key={colIndex}
                  className="p-2 text-xs font-medium text-gray-500 border-r border-gray-300 min-w-20"
                >
                  {getColumnLabel(colIndex)}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {Array.from({ length: maxRows }, (_, rowIndex) => (
              <tr key={rowIndex} className="hover:bg-gray-50">
                <td className="p-2 text-xs text-gray-500 bg-gray-100 border-r border-gray-300 text-center">
                  {rowIndex + 1}
                </td>
                {Array.from({ length: maxCols }, (_, colIndex) => {
                  const cell = activeSheet.data[rowIndex] && activeSheet.data[rowIndex][colIndex];
                  return (
                    <td
                      key={colIndex}
                      className="p-0 border-r border-b border-gray-300"
                    >
                      <input
                        type="text"
                        value={getCellDisplayValue(cell)}
                        onChange={(e) =>
                          handleCellChange(rowIndex, colIndex, e.target.value)
                        }
                        title={getCellTitle(cell)}
                        className="w-full h-8 px-2 text-sm border-none outline-none focus:bg-blue-50 focus:ring-2 focus:ring-blue-500 focus:ring-inset"
                        style={{
                          minWidth: '80px',
                          ...getCellStyle(cell)
                        }}
                      />
                    </td>
                  );
                })}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
};

export default SimpleExcelViewer;