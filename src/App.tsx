import React, { useState } from 'react';
import FileUpload from './components/FileUpload';
import SimpleExcelViewer from './components/SimpleExcelViewer';
import './App.css';

function App() {
  const [selectedFile, setSelectedFile] = useState<File | null>(null);

  const handleFileSelect = (file: File) => {
    setSelectedFile(file);
  };

  return (
    <div className="min-h-screen bg-gray-100">
      <div className="container mx-auto px-4 py-8">
        <header className="text-center mb-8">
          <h1 className="text-4xl font-bold text-gray-800 mb-2">
            Excel Viewer
          </h1>
          <p className="text-gray-600">
            Upload and view Excel files directly in your browser
          </p>
        </header>

        <div className="bg-white rounded-lg shadow-lg p-6">
          <FileUpload onFileSelect={handleFileSelect} />
          
          {selectedFile && (
            <div className="mt-6">
              <div className="flex items-center justify-between mb-4">
                <h2 className="text-xl font-semibold text-gray-800">
                  {selectedFile.name}
                </h2>
                <span className="text-sm text-gray-500">
                  {(selectedFile.size / (1024 * 1024)).toFixed(2)} MB
                </span>
              </div>
              <SimpleExcelViewer file={selectedFile} />
            </div>
          )}
        </div>
      </div>
    </div>
  );
}

export default App;