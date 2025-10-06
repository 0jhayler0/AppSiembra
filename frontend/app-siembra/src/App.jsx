import React, { useState, useRef } from "react";
import axios from "axios";
import "./App.css";

function App() {
  const [file, setFile] = useState(null);
  const [isDragging, setIsDragging] = useState(false);
  const [isDropping, setIsDropping] = useState(false);
  const inputRef = useRef(null);

  const handleUpload = async () => {
  if (!file) return alert("Selecciona o arrastra un archivo Excel primero");

  const formData = new FormData();
  formData.append("file", file);

  try {
    const res = await axios.post("http://localhost:5000/upload-excel", formData, {
      responseType: "blob", // 👈 Muy importante para recibir el Excel
      headers: { "Content-Type": "multipart/form-data" },
    });

    // Crear un objeto URL temporal
    const url = window.URL.createObjectURL(new Blob([res.data]));
    const link = document.createElement("a");
    link.href = url;
    link.setAttribute("download", "Reporte_Siembra.xlsx"); // 👈 Nombre del archivo
    document.body.appendChild(link);
    link.click();

    // Limpieza
    link.remove();
    window.URL.revokeObjectURL(url);
    setFile(null);
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
    if (name.endsWith(".xlsx") || name.endsWith(".xls")) {
      setFile(droppedFile);
    } else {
      alert("Solo se permiten archivos Excel (.xlsx, .xls)");
    }

    // Desactiva el estado "isDropping" después de un momento
    setTimeout(() => setIsDropping(false), 300);
  };

  const handleContainerClick = (e) => {
    // Bloquea el click mientras arrastras 
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
        className={`principalContainer ${isDragging ? "dragging" : ""}`}
        onDragOver={handleDragOver}
        onDragEnter={handleDragEnter}
        onDragLeave={handleDragLeave}
        onDrop={handleDrop}
        onClick={handleContainerClick}
        role="button"
        tabIndex={0}
      >
        <h1 className="title">Reportes de siembra</h1>
        <h2 className="text">Selecciona o arrastra un archivo Excel para generar los reportes automáticamente</h2>

        <input
          ref={inputRef}
          style={{ display: "none" }}
          type="file"
          accept=".xlsx, .xls"
          onChange={(e) => {
            const f = e.target.files && e.target.files[0];
            if (!f) return;
            const name = String(f.name).toLowerCase();
            if (name.endsWith(".xlsx") || name.endsWith(".xls")) setFile(f);
            else alert("Solo se permiten archivos Excel (.xlsx, .xls)");
          }}
        />

        <div style={{ minHeight: 40 }}>
          {file ? (
            <p style={{ color: "#23a523", margin: 0 }}>📄 Archivo seleccionado: {file.name}</p>
          ) : (
            <p style={{ color: "#666", margin: 0 }}>O haz clic aquí o arrastra un .xlsx/.xls</p>
          )}
        </div>
      </div>

        <button onClick={handleUpload} disabled={!file}>
          Crear reportes
        </button>
    </section>
  );
}

export default App;
