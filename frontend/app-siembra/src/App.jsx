import React, { useState } from "react";
import axios from "axios";

function App() {
  const [file, setFile] = useState(null);

  const handleUpload = async () => {
    if (!file) return alert("Selecciona un archivo Excel primero");
    const formData = new FormData();
    formData.append("file", file);

    try {
      const res = await axios.post("http://localhost:5000/upload-excel", formData, {
        headers: { "Content-Type": "multipart/form-data" },
      });
      console.log("Respuesta del backend:", res.data);
    } catch (error) {
      console.error("Error subiendo el archivo:", error);
    }
  };

  return (
    <div style={{ padding: "20px" }}>
      <h1>Subir archivo Excel</h1>
      <input
        type="file"
        accept=".xlsx, .xls"
        onChange={(e) => setFile(e.target.files[0])}
      />
      <button onClick={handleUpload} disabled={!file}>
        Enviar al backend
      </button>
    </div>
  );
}

export default App;

