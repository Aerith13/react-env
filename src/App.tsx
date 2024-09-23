import React, { useState } from 'react';

interface OcrComponentProps {}

const OcrComponent: React.FC<OcrComponentProps> = () => {
  const [image, setImage] = useState(null);
  const [extractedText, setExtractedText] = useState('');
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState('');
  const [draggedOver, setDraggedOver] = useState(false);

  const handleImageChange = (event: React.ChangeEvent<HTMLInputElement>) => {
    if (event.target.files && event.target.files.length > 0) {
      setImage(event.target.files[0]);
    }
  };

  const handleDragOver = (event: React.DragEvent<HTMLDivElement>) => {
    event.preventDefault();
    setDraggedOver(true);
  };

  const handleDragLeave = () => {
    setDraggedOver(false);
  };

  const handleDrop = (event: React.DragEvent<HTMLDivElement>) => {
    event.preventDefault();
    setDraggedOver(false);
    if (event.dataTransfer.files && event.dataTransfer.files.length > 0) {
      setImage(event.dataTransfer.files[0]);
    }
  };

  const handleExtractText = async () => {
    if (image) {
      setLoading(true);
      setError('');
      try {
        const formData = new FormData();
        formData.append('language', 'eng');
        formData.append('isOverlayRequired', 'false');
        formData.append('iscreatesearchablepdf', 'false');
        formData.append('issearchablepdfhidetextlayer', 'false');
        formData.append('filetype', 'image/jpeg');
        formData.append('base64Image', '');
        formData.append('file', image);

        const response = await fetch('https://api.ocr.space/parse/image', {
          method: 'POST',
          headers: {
            apikey: 'K85034884388957',
          },
          body: formData,
        });

        if (response.ok) {
          const data = await response.json();
          setExtractedText(data.ParsedResults[0].ParsedText);
        } else {
          setError('Failed to extract text from image');
        }
      } catch (error) {
        setError('An error occurred while extracting text from image');
      } finally {
        setLoading(false);
      }
    }
  };

  const handlePostToExcel = async () => {
    if (extractedText) {
      try {
        await Office.context.document.load('Selection');
        await Office.context.sync();
        const selection = Office.context.document.Selection;
        selection.insertText(extractedText, 'replace');
      } catch (error) {
        setError('Failed to post extracted text to Excel');
      }
    }
  };

  const handleAddToExcel = async () => {
    if (extractedText) {
      try {
        await Office.context.document.load('Selection');
        await Office.context.sync();
        const selection = Office.context.document.Selection;
        selection.insertText(extractedText, 'after');
      } catch (error) {
        setError('Failed to add extracted text to Excel');
      }
    }
  };

  const handleClear = () => {
    setImage(null);
    setExtractedText('');
    setError('');
  };

  return (
    <div className="max-w-3xl mx-auto p-4 bg-#faf7f5 rounded-md shadow-md">
      <h1 className="text-3xl font-bold text-center mb-4 text-#432c47">OCR Text Extractor</h1>
      <div
        className={`flex flex-col items-center justify-center mb-4 p-4 border-2 border-dashed border-gray-500 rounded-md ${draggedOver ? 'bg-gray-200' : 'bg-gray-100'}`}
        onDragOver={handleDragOver}
        onDragLeave={handleDragLeave}
        onDrop={handleDrop}
      >
        <p className="text-gray-600">Drag and drop an image file here</p>
      </div>
      <input
        type="file"
        onChange={handleImageChange}
        className="block w-full text-sm text-gray-500 file:mr-4 file:py-2 file:px-4 file:rounded-full file:border-0 file:text-sm file:font-semibold file:bg-#e5cac8 file:text-#4f4f4f hover:file:bg-#c2a15a"
      />
      {image && (
        <img src={URL.createObjectURL(image)} alt="Uploaded Image" className="w-full h-auto mb-4" />
      )}
      {image && (
        <button
          onClick={handleExtractText}
          className="mt-4 py-2 px-4 bg-#432c47 text-#faf7f5 rounded-lg hover:bg-#c2a15a"
        >
          {loading ? 'Extracting...' : 'Extract Text'}
        </button>
      )}
      {extractedText && (
        <div className="bg-#e5cac8 p-4 rounded-md">
          <h2 className="text-2xl font-bold mb-2 text-#432c47">Extracted Text:</h2>
          <p className="text-#4f4f4f">{extractedText}</p>
          <button
            onClick={handlePostToExcel}
            className="mt-4 py-2 px-4 bg-#432c47 text-#faf7f5 rounded-lg hover:bg-#c2a15a"
          >
            Post to Excel
          </button>
          <button
            onClick={handleAddToExcel}
            className="mt-4 py-2 px-4 bg-#432c47 text-#faf7f5 rounded-lg hover:bg-#c2a15a"
          >
            Add to Excel
          </button>
        </div>
      )}
      {error && (
        <div className="bg-red-100 p-4 rounded-md">
          <h2 className="text-2xl font-bold mb-2">Error:</h2>
          <p className="text-red-600">{error}</p>
        </div>
      )}
      <button
        onClick={handleClear}
        className="mt-4 py-2 px-4 bg-#432c47 text-#faf7f5 rounded-lg hover:bg-#c2a15a"
      >
        Clear
      </button>
    </div>
  );
};

export default OcrComponent;