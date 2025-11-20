// pdfshift.js
const axios = require("axios");

async function generarPDF(html) {
  const apiKey = process.env.PDFSHIFT_API_KEY;
  if (!apiKey) throw new Error("Falta PDFSHIFT_API_KEY");

  const response = await axios.post(
    "https://api.pdfshift.io/v3/convert/pdf",
    {
      source: html,       // HTML completo que le mandamos
      landscape: false,   // Si quieres modo horizontal, cámbialo a true
      use_print: true     // Activa estilos @media print (más bonito para PDF)
    },
    {
      auth: {
        username: apiKey,
        password: ""      // PDFShift requiere username=APIKEY, password vacío
      },
      responseType: "arraybuffer",
      headers: {
        "Content-Type": "application/json"
      }
    }
  );

  // Retorna el PDF como Buffer
  return response.data;
}

module.exports = { generarPDF };
