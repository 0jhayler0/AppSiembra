const express = require("express");
const multer = require("multer");
const XLSX = require("xlsx");
const ExcelJS = require("exceljs");
const fs = require("fs");
const path = require("path");
const variedades = require("./variedades");
const sheets = require("./sheets");

const app = express();
const upload = multer({ dest: path.join(__dirname, 'output', 'uploads') });
const cors = require('cors');

// Permitir peticiones desde el frontend (Vite por defecto corre en 5173)
app.use(cors({
  origin: ['http://localhost:5173', 'http://127.0.0.1:5173'],
  methods: ['GET','POST','OPTIONS'],
  exposedHeaders: ['Content-Disposition']
}));

// puppeteer is optional; if not installed we'll return a helpful error when PDF is requested
let puppeteer = null;
try {
  puppeteer = require('puppeteer');
} catch (e) {
  // ignore; we'll notify later if PDF requested
}

// === 🔹 FUNCIONES AUXILIARES =======================

const limpiarDatos = (data) => {
  return data
    .filter(row => row.some(cell => cell !== "" && cell != null))
    .map(row => row.map(cell => (typeof cell === "string" ? cell.trim() : cell)));
};

const rellenarColumna = (data, indexCol = 0) => {
  if (data.length === 0) return data;
  let lastValue = null;
  return data.map(row => {
    const valorActual = row[indexCol];
    if (valorActual && String(valorActual).trim() !== "") {
      lastValue = valorActual;
    } else {
      row[indexCol] = lastValue;
    }
    return row;
  });
};

const extraerSeccion = (row) => {
  const textoSeccion = row.find(cell => typeof cell === "string" && cell.includes("Seccion:"));
  if (textoSeccion) {
    const match = textoSeccion.match(/Seccion:\s*(\d+)/);
    if (match && match[1]) return match[1];
  }
  return null;
};

function expandirVariedades(row) {
  const multiplesVariedades = row.Variedad;
  const regex = /(.+?)\s*\(([\d\.]+)\)/g;
  let match;
  const nuevasFilas = [];
  while ((match = regex.exec(multiplesVariedades)) !== null) {
    const nombreVariedad = match[1].trim();
    const largoVariedad = match[2];
    nuevasFilas.push({ Seccion: row.Seccion, Lado: row.Lado, Nave: row.Nave, Era: row.Era, Fecha_Siembra: row.Fecha_Siembra, Inicio_Corte: row.Inicio_Corte, Variedad: nombreVariedad, Largo: largoVariedad });
  }
  if (nuevasFilas.length === 0) return [{ ...row }];
  return nuevasFilas;
}



// ========== RUTA PRINCIPAL ====================================================================

app.post("/upload-excel", upload.single("file"), async (req, res) => {
    try {
        if (!req.file) return res.status(400).json({ error: "No se envió ningún archivo" });

        const filePath = req.file.path;
        const workbook = XLSX.readFile(filePath);
        const sheetName = workbook.SheetNames[0];
        const data = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1 });

        let datosLimpios = limpiarDatos(data);
        const datosCrudos = []; 
        let seccionActual = "N/A";
        let semanaActual = "";

const extraerSemana = (row) => {
  if (!Array.isArray(row)) return null;

  for (const cell of row) {
    if (!cell) continue;
    const texto = String(cell).trim();

    const match = texto.match(/Semana\s+Siembra\s+(2\d{5})/i);
    if (match) {
      console.log("🗓 Semana encontrada:", match[1], "en texto:", texto);
      return match[1];
    }
  }

  return null;
};


        for (let i = 0; i < datosLimpios.length; i++) {
            const row = datosLimpios[i];
            const nuevaSemana = extraerSemana(row);
            const nuevaSeccion = extraerSeccion(row);

            if (nuevaSemana) {
                semanaActual = nuevaSemana;
            }

            if (nuevaSeccion) {
                seccionActual = nuevaSeccion;
                continue;
            }

            if (row[0] === "Nave" && row[6] === "Nave") {
                const bloqueDatos = [];
                let j = i + 1;
                while (j < datosLimpios.length) {
                    const currentRow = datosLimpios[j];
                    if (extraerSeccion(currentRow) !== null || (currentRow[0] === "Nave" && currentRow[6] === "Nave")) break;
                    if (currentRow.some(cell => cell !== "" && cell != null)) bloqueDatos.push(currentRow);
                    j++;
                }
                i = j - 1;

                if (bloqueDatos.length > 0) {
                    let datosCompletos = rellenarColumna(bloqueDatos, 0);
                    datosCompletos = rellenarColumna(datosCompletos, 6);

                   let filaId = 0;
                    datosCompletos.forEach(r => {
                        datosCrudos.push(
                            { Seccion: seccionActual, Lado: "A", FilaId: filaId, Nave: r[0] || "", Era: r[1] || "", Variedad: r[2] || "", Largo: r[3] || "", Fecha_Siembra: r[4] || "", Inicio_Corte: r[5] || "" },
                            { Seccion: seccionActual, Lado: "B", FilaId: filaId, Nave: r[6] || "", Era: r[7] || "", Variedad: r[8] || "", Largo: r[9] || "", Fecha_Siembra: r[10] || "", Inicio_Corte: r[11] || "" }
                        );
                        filaId++;
                    });
                }
            }
        }


        
        const datosFinales = datosCrudos.flatMap(expandirVariedades);

        const wbFinal = new ExcelJS.Workbook();

        
  // Creacion de las hojas .xlsx
  sheets.crearHojaDistribucionProductos(wbFinal, datosFinales);
  sheets.crearHojaDisbud(wbFinal, datosFinales);
  sheets.crearHojaGirasol(wbFinal, datosFinales);
  sheets.crearHojaPruebaFloracion(wbFinal, datosFinales);
  sheets.crearHojaNochesLuz(wbFinal, datosCrudos, { variedades });

        // Guardar archivo final o generar PDF según query param
        const wantPdf = String(req.query.format || '').toLowerCase() === 'pdf';

        if (!wantPdf) {
          const outputPath = `Reporte_Siembra_${semanaActual}_${Date.now()}.xlsx`;
          await wbFinal.xlsx.writeFile(outputPath);
          console.log("Reporte completo generado:", outputPath);
          res.download(outputPath, `Reporte_Siembra_${semanaActual}.xlsx`, (err) => {
              if (err) console.error("Error enviando el archivo:", err);
              try { fs.unlinkSync(outputPath); } catch(e){}
          });
        } else {
          if (!puppeteer) {
            return res.status(400).json({ error: 'PDF generation not available: puppeteer not installed on the server.' });
          }

          // Convert each worksheet to a simple HTML table and concatenate
          function sheetToHTML(sheet) {
            const rows = [];
            sheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
              const cells = [];
              row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
                const text = cell.value == null ? '' : String(cell.value);
                cells.push(`<td>${text.replace(/</g,'&lt;').replace(/>/g,'&gt;')}</td>`);
              });
              rows.push(`<tr>${cells.join('')}</tr>`);
            });
            return `<h2>${sheet.name}</h2><table border="1" style="border-collapse:collapse; width:100%">${rows.join('')}</table><div style="page-break-after:always"></div>`;
          }

          // build HTML
          const htmlParts = ['<html><head><meta charset="utf-8"><style>table,td{font-family:Arial,sans-serif;font-size:10px;padding:4px}</style></head><body>'];
          wbFinal.eachSheet(sheet => {
            htmlParts.push(sheetToHTML(sheet));
          });
          htmlParts.push('</body></html>');
          const finalHtml = htmlParts.join('\n');

          // render PDF with puppeteer
          const browser = await puppeteer.launch({ args: ['--no-sandbox', '--disable-setuid-sandbox'] });
          const page = await browser.newPage();
          await page.setContent(finalHtml, { waitUntil: 'networkidle0' });
          const pdfBuffer = await page.pdf({ format: 'A4', printBackground: true });
          await browser.close();

          const outputPdf = `Reporte_Siembra_${semanaActual}_${Date.now()}.pdf`;
          fs.writeFileSync(outputPdf, pdfBuffer);
          res.setHeader('Content-Type', 'application/pdf');
          res.setHeader('Content-Disposition', `attachment; filename="Reporte_Siembra_${semanaActual}.pdf"`);
          res.sendFile(path.resolve(outputPdf), (err) => {
            try { fs.unlinkSync(outputPdf); } catch(e){}
          });
        }

        // borrar archivo subido
        try { fs.unlinkSync(filePath); } catch(e){}

    } catch (error) {
        console.error("Error procesando Excel:", error);
        if (req.file && fs.existsSync(req.file.path)) fs.unlinkSync(req.file.path);
        res.status(500).json({ error: "Error procesando el archivo", detalle: error.message });
    }
});

app.listen(5000, () =>
  console.log("🚀 Servidor corriendo en http://localhost:5000")
);
