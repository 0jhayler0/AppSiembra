import React, { useState } from 'react';
import "../styles/Home.css";

const Home = () => {
  const [dragging, setDragging] = useState(false);
  const [file, setFile] = useState(null);

  const handleDragOver = (e) => {
    e.preventDefault(); 
    setDragging(true);
  };

  const handleDragLeave = () => {
    setDragging(false);
  };

  const handleDrop = (e) => {
    e.preventDefault();
    setDragging(false);

    if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
      setFile(e.dataTransfer.files[0]); 
      console.log("Archivo recibido:", e.dataTransfer.files[0]);
    }
  };

  return (
    <div className="homeContainer">
      <div className={`centralDiv ${dragging ? "dragging" : ""}`} 
        onDragOver={handleDragOver}
        onDragLeave={handleDragLeave}
        onDrop={handleDrop}>
        <h1>Crear Reportes</h1>
        {file ? (
          <p>Archivo cargado: {file.name}</p>
        ) : (
        <p>
          Arrasta y suelta un archivo PDF para crear reportes de desbotonado, distribucion de los productos, pruebas de floracion y girasol sembrado semanalmente
        </p>
        )}
        <p>o</p>
        <input type="file"
          id="fileImput"
          accept="aplication/pdf"
          onChange={(e) => setFile(e.target.files[0])}
        />
      </div>
    </div>
  )
}

export default Home
