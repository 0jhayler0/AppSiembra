import { useState } from "react";
import axios from "axios";

function App() {
  const [file, setFile] = useState(null);

  const handleupload = async () => {
    const formData = new formData();
    formData.append("file", file);

    const response = await axios.post("http://localhost:5000/upload", formData, {
      responseType: "blob",
  });

    const url = window.URL.createObjectURL(new Blob([response.data]));
    const link = document.createElement("a");
    link.href = url;
    link.setAttribute("download", "datos.xlsx");
    document.body.appendChild(link);
    link.click();
  };
};

return (
  <div>
    <h1>TEXTO</h1>
    <input type="file" onChange={(e) => setFile(e.target.files[0])} />
    <button onClick={handleUpload} disabled={!file}>
      Convertir
    </button>
  </div>
)

export default App
