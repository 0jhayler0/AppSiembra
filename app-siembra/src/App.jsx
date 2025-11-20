import React, { useState, useRef } from "react";
import axios from "axios";
import "./App.css";

function App() {
  const [file, setFile] = useState(null);
  const [isDragging, setIsDragging] = useState(false);
  const [isDropping, setIsDropping] = useState(false);
  const [isProcessing, setIsProcessing] = useState(false);
  const [showSuccess, setShowSuccess] = useState(false);
  const [reportData, setReportData] = useState(null);
  const inputRef = useRef(null);

  const downloadBase64 = (base64String, filename) => {
    try {
      const byteCharacters = atob(base64String);
      const byteNumbers = new Array(byteCharacters.length);
      for (let i = 0; i < byteCharacters.length; i++) {
        byteNumbers[i] = byteCharacters.charCodeAt(i);
      }
      const byteArray = new Uint8Array(byteNumbers);
      
      const mimeType = filename.endsWith('.pdf') 
        ? 'application/pdf' 
        : 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
      
      const blob = new Blob([byteArray], { type: mimeType });
      const url = window.URL.createObjectURL(blob);
      const link = document.createElement("a");
      link.href = url;
      link.setAttribute("download", filename);
      document.body.appendChild(link);
      link.click();
      link.remove();
      window.URL.revokeObjectURL(url);
      
      setShowSuccess(true);
      setTimeout(() => setShowSuccess(false), 3000);
    } catch (error) {
      console.error("Error descargando archivo:", error);
      alert("❌ Error al descargar el archivo");
    }
  };

  const handleUpload = async () => {
    if (!file) return alert("Selecciona o arrastra un archivo Excel primero");
    
    const formData = new FormData();
    formData.append("file", file);

    try {
      setIsProcessing(true);

      const res = await axios.post(
        `https://app-siembra-backend.onrender.com/upload-excel`, 
        formData, 
        {
          headers: { "Content-Type": "multipart/form-data" },
        }
      );

      setReportData(res.data);
      setIsProcessing(false);
      
    } catch (error) {
      console.error("Error subiendo el archivo:", error);
      alert("❌ Error al generar el reporte. Revisa la consola para más detalles.");
      setIsProcessing(false);
    }
  };

  const handleDownloadExcel = () => {
    if (!reportData || !reportData.excel) {
      return alert("Primero debes procesar un archivo");
    }
    downloadBase64(reportData.excel.data, reportData.excel.filename);
  };

  const handleDownloadPDF = () => {
    if (!reportData || !reportData.pdf) {
      return alert("El PDF no está disponible para este reporte");
    }
    downloadBase64(reportData.pdf.data, reportData.pdf.filename);
  };

  const handleDragOver = (e) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragging(true);
  };

  const handleDragEnter = (e) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragging(true);
  };

  const handleDragLeave = (e) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragging(false);
  };

  const handleDrop = (e) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragging(false);
    setIsDropping(true);

    const droppedFile = e.dataTransfer?.files?.[0];
    if (!droppedFile) return;

    const name = String(droppedFile.name).toLowerCase();

    if (name.endsWith(".xlsx") || name.endsWith(".xls") || name.endsWith(".pdf")) {
      setFile(droppedFile);
      setReportData(null);
    } else {
      alert("Solo se permiten archivos Excel o PDF(.xlsx, .xls, .pdf)");
    }

    setTimeout(() => setIsDropping(false), 300);
  };

  const handleContainerClick = (e) => {
    if (isDragging || isDropping) {
      e.preventDefault();
      e.stopPropagation();
      return;
    }
    if (inputRef.current) inputRef.current.click();
  };

  return (
    <section className="mainContainer">
      <div
        className={`principalContainer ${isDragging ? "dragging" : ""} ${isProcessing ? "rotating" : ""}`}
        onDragOver={handleDragOver}
        onDragEnter={handleDragEnter}
        onDragLeave={handleDragLeave}
        onDrop={handleDrop}
        onClick={handleContainerClick}
        role="button"
        tabIndex={0}
      >
        <h1 className="title">Reportes de siembra</h1>
        <h2 className="text">
          Haz click para seleccionar o arrastra un archivo de Excel o PDF para generar los reportes automáticamente
        </h2>

        <input
          ref={inputRef}
          style={{ display: "none" }}
          type="file"
          accept=".xlsx, .xls, .pdf"
          onChange={(e) => {
            const f = e.target.files && e.target.files[0];
            if (!f) return;
            const name = String(f.name).toLowerCase();
            if (name.endsWith(".xlsx") || name.endsWith(".xls") || name.endsWith(".pdf")) {
              setFile(f);
              setReportData(null);
            } else {
              alert("Solo se permiten archivos Excel o PDF (.xlsx, .xls, .pdf)");
            }
          }}
        />

        <div style={{ minHeight: 40 }}>
          {file ? (
            <p style={{ color: "#23a523", margin: 0 }}>
              📄 Archivo seleccionado: {file.name}
            </p>
          ) : (
            <p className="pointsAnimation" style={{ color: "#666", margin: 0, fontSize: 15, fontFamily: "monospace" }}>
              Esperando el archivo<span>.</span><span>.</span><span>.</span>
            </p>
          )}
        </div>
      </div>

      <div className="buttons">
        {!reportData && (
          <button 
            className="btn-excel" 
            onClick={handleUpload} 
            disabled={!file || isProcessing}
            style={{ width: '100%', maxWidth: '400px' }}
          >
            {isProcessing ? 'Procesando...' : 'Procesar Archivo'}
          </button>
        )}

        {reportData && (
          <>
            <button className="btn-excel" onClick={handleDownloadExcel}>
              Descargar Excel
            </button>
            <button 
              className="btn-pdf" 
              onClick={handleDownloadPDF}
              disabled={!reportData.pdf}
            >
              Descargar PDF
            </button>
            <button 
              className="btn-excel" 
              onClick={() => {
                setFile(null);
                setReportData(null);
              }}
              style={{ background: '#666' }}
            >
              Procesar otro archivo
            </button>
          </>
        )}
      </div>

      {showSuccess && <div className="toast">✅ Reporte descargado con éxito</div>}
    </section>
  );
}

export default App;