import React, { useState, useRef } from "react";
import axios from "axios";
import "./App.css";

function App() {
  const [file, setFile] = useState(null);
  const [isDragging, setIsDragging] = useState(false);
  const [isDropping, setIsDropping] = useState(false);
  const [isProcessing, setIsProcessing] = useState(false);
  const [showSuccess, setShowSuccess] = useState(false);
  const inputRef = useRef(null);

  const downloadBlob = (blob, dispositionFallback) => {
    const disposition = blob.headers ? (blob.headers["content-disposition"] || blob.headers["Content-Disposition"]) : null;
    let filename = dispositionFallback || "Reporte_Siembra";
    if (disposition) {
      const match = /filename\*?=(?:UTF-8''?)?["']?([^;"']+)["']?/i.exec(disposition);
      if (match && match[1]) {
        try { filename = decodeURIComponent(match[1]); } catch (e) { filename = match[1]; }
      }
    }
    const url = window.URL.createObjectURL(new Blob([blob.data] || [blob]));
    const link = document.createElement("a");
    link.href = url;
    link.setAttribute("download", filename);
    document.body.appendChild(link);
    link.click();
    link.remove();
    window.URL.revokeObjectURL(url);
    setFile(null);
    setIsProcessing(false);
    setShowSuccess(true);
    setTimeout(() => setShowSuccess(false), 3000);
  };

  const handleUpload = async (format = 'xlsx') => {
    if (!file) return alert("Selecciona o arrastra un archivo Excel primero");
    const formData = new FormData();
    formData.append("file", file);

    try {
      setIsProcessing(true);

    const res = await axios.post(`https://app-siembra-backend.onrender.com/upload-excel`, formData, {
      responseType: "blob",
      headers: { "Content-Type": "multipart/form-data" },
});


      downloadBlob({ data: res.data, headers: res.headers }, `Reporte_Siembra_.${format === 'pdf' ? 'pdf' : 'xlsx'}`);
    } catch (error) {
      console.error("Error subiendo el archivo:", error);
      alert("❌ Error al generar el reporte. Revisa la consola para más detalles.");
    }
  };


  // --- DRAG & DROP ---
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
        <h2 className="text">Haz click para seleccionar o arrastra un archivo de Excel o PDF para generar los reportes automáticamente</h2>

        <input
          ref={inputRef}
          style={{ display: "none" }}
          type="file"
          accept=".xlsx, .xls, .pdf"
          onChange={(e) => {
            const f = e.target.files && e.target.files[0];
            if (!f) return;
            const name = String(f.name).toLowerCase();
            if (name.endsWith(".xlsx") || name.endsWith(".xls") || name.endsWith(".pdf")) setFile(f);
            else alert("Solo se permiten archivos Excel o PDF (.xlsx, .xls, .pdf)");
          }}
        />

        <div style={{ minHeight: 40 }}>
          {file ? (
            <p style={{ color: "#23a523", margin: 0 }}>📄 Archivo seleccionado: {file.name}</p>
          ) : (
            <p className="pointsAnimation" style={{ color: "#666", margin: 0, fontSize: 15, fontFamily: "monospace"}}>Esperando el archivo<span>.</span><span>.</span><span>.</span></p>
          )}
        </div>
      </div>
          <div className="buttons"></div>
        <div className="buttons">
          <button className="btn-excel" onClick={() => handleUpload('xlsx')} disabled={!file}>
            Descargar en Excel
          </button>
          <button className="btn-pdf" onClick={() => handleUpload('pdf')} disabled={!file}>
            Descargar en PDF
          </button>
        </div>
        {showSuccess && <div className="toast">✅ Reporte realizado con éxito</div>} 
    </section>
  );
}

export default App;
