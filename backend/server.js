const express = require("express");
const multer = require("multer");
const ExcelJS = require("exceljs");
const fs = require("fs");
const path = require("path");
const variedades = require("./variedades");
const sheets = require("./sheets");
const cors = require('cors');
const axios = require('axios');
const FormData = require('form-data');

const app = express();
// Guardar uploads en /tmp (mejor para Docker / Render)
const upload = multer({ dest: path.join('/tmp') });

app.use(cors({
  origin: [
    'http://localhost:5173',
    'http://127.0.0.1:5173',
    'https://appsiembralavictoria.web.app',
    process.env.FRONTEND_URL || ''
  ],
  methods: ['GET','POST','OPTIONS'],
  exposedHeaders: ['Content-Disposition']
}));

let puppeteer = null;
try {
  puppeteer = require('puppeteer');
} catch (e) {
  // no hago nada; si intentan generar PDF y no está instalado lanzamos error más adelante
}

// === FUNCIONES AUXILIARES =======================

const getTextFromCell = (cell) => {
  if (cell == null) return "";
  if (typeof cell === 'string' || typeof cell === 'number') return String(cell).trim();
  if (cell.text) return String(cell.text).trim(); // some cell shapes
  if (cell.richText && Array.isArray(cell.richText)) return cell.richText.map(p => p.text).join('').trim();
  return String(cell).trim ? String(cell).trim() : "";
};

const limpiarDatos = (data) => {
  return data
    .filter(row => Array.isArray(row) && row.some(cell => cell !== "" && cell != null))
    .map(row => row.map(cell => (typeof cell === "string" ? cell.trim() : cell)));
};

const rellenarColumna = (data, indexCol = 0) => {
  if (!Array.isArray(data) || data.length === 0) return data;
  let lastValue = null;
  return data.map(row => {
    const valorActual = row[indexCol];
    if (valorActual && String(valorActual).trim() !== "") {
      lastValue = valorActual;
    } else {
      // Si no hay valor actual, rellenar con el último
      row[indexCol] = lastValue;
    }
    return row;
  });
};

const extraerSeccion = (row) => {
  if (!Array.isArray(row)) return null;
  for (const cell of row) {
    if (!cell) continue;
    const texto = String(cell);
    if (texto.includes("Seccion:")) {
      const match = texto.match(/Seccion:\s*(\d+)/i);
      if (match && match[1]) return match[1];
    }
  }
  return null;
};

function expandirVariedades(row) {
  const multiplesVariedades = row.Variedad || "";
  const regex = /(.+?)\s*\(([\d\.]+)\)/g;
  let match;
  const nuevasFilas = [];
  while ((match = regex.exec(multiplesVariedades)) !== null) {
    const nombreVariedad = match[1].trim();
    const largoVariedad = match[2];
    nuevasFilas.push({
      Seccion: row.Seccion,
      Lado: row.Lado,
      Nave: row.Nave,
      Era: row.Era,
      Fecha_Siembra: row.Fecha_Siembra,
      Inicio_Corte: row.Inicio_Corte,
      Variedad: nombreVariedad,
      Largo: largoVariedad
    });
  }
  if (nuevasFilas.length === 0) return [{ ...row }];
  return nuevasFilas;
}

// ========== RUTA PRINCIPAL ====================================================================

app.post("/upload-excel", upload.single("file"), async (req, res) => {
  let originalUploadPath = null;
  let convertedXlsxPath = null;
  let pdfPath = null;
  let finalReportPath = null;

  try {
    if (!req.file) return res.status(400).json({ error: "No se envió ningún archivo" });

    originalUploadPath = req.file.path;
    console.log("Archivo recibido:", { path: originalUploadPath, originalname: req.file.originalname });

    let filePath = originalUploadPath;
    const originalName = req.file.originalname || '';
    const ext = path.extname(originalName).toLowerCase();

    // Si viene PDF, renombrar y llamar al servicio python para convertir
    if (ext === '.pdf' || req.file.mimetype === 'application/pdf') {
      pdfPath = filePath + '.pdf';
      fs.renameSync(filePath, pdfPath);
      filePath = pdfPath;
      const outXlsx = filePath + '.converted.xlsx';

      try {
        const pythonServiceUrl = process.env.PYTHON_SERVICE_URL || 'http://localhost:5001';
        const formData = new FormData();
        formData.append('file', fs.createReadStream(pdfPath));

        const response = await axios.post(`${pythonServiceUrl}/upload-excel`, formData, {
          responseType: 'stream',
          headers: formData.getHeaders ? formData.getHeaders() : { 'Content-Type': 'multipart/form-data' }
        });

        const writer = fs.createWriteStream(outXlsx);
        response.data.pipe(writer);

        await new Promise((resolve, reject) => {
          writer.on('finish', resolve);
          writer.on('error', reject);
        });

        filePath = outXlsx;
        convertedXlsxPath = outXlsx;
        console.log("PDF convertido a XLSX:", convertedXlsxPath);
      } catch (err) {
        console.error('Error convirtiendo PDF a XLSX:', err);
        // limpieza parcial
        try { if (fs.existsSync(pdfPath)) fs.unlinkSync(pdfPath); } catch(e){}
        return res.status(500).json({ error: 'Error convirtiendo PDF a XLSX', detalle: String(err) });
      }
    }

    // Leer Excel y construir matriz 0-based (igual que la versión A que funcionaba)
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    const worksheet = workbook.getWorksheet(1);
    const data = [];
    worksheet.eachRow((row, rowNumber) => {
      const rowData = [];
      row.eachCell((cell, colNumber) => {
        rowData.push(cell.value);
      });
      data.push(rowData);
    });

    let datosLimpios = limpiarDatos(data);
    console.log("Filas leídas:", datosLimpiasCount=datosLimpios.length ? datosLimpios.length : 0);

    const datosCrudos = [];
    let seccionActual = "NN";
    let semanaActual = "NN";

    const extraerSemana = (row) => {
  if (!Array.isArray(row)) return null;

  for (const cell of row) {
    if (!cell) continue;

    const texto = getTextFromCell(cell);

    if (!texto) continue;

    // Caso real del archivo: "Flores de la Victoria S.A.S Semana Siembra 202545 Seccion: 06"
    let m = texto.match(/Semana\s+Siembra\s+(2\d{5})/i);
    if (m) return m[1];

    // Otros posibles formatos
    m = texto.match(/Semana(?:\s+Siembra)?\s+(\d{1,2})\b/i);
    if (m) return m[1];

    m = texto.match(/\bSem\s+(\d{1,2})\b/i);
    if (m) return m[1];
  }

  return null;
};

    // LOG: inicio parse
    console.log("Iniciando parseo de datos (método A)...");

    for (let i = 0; i < datosLimpios.length; i++) {
      const row = datosLimpios[i];

      // Extraer semana si aparece en cualquier celda de la fila
      const nuevaSemana = extraerSemana(row);
      if (nuevaSemana) {
        semanaActual = nuevaSemana;
        continue;
      }

      // Extraer seccion
      const nuevaSeccion = extraerSeccion(row);
      if (nuevaSeccion) {
        seccionActual = nuevaSeccion;
        continue;
      }

      // Detectar encabezado de bloque (igual que en la versión A)
      const cell0 = getTextFromCell(row[0]).toLowerCase();
      const cell6 = getTextFromCell(row[6]).toLowerCase();
      if (cell0 === "nave" && cell6 === "nave") {
        // recolectar bloque hasta próximo encabezado o seccion
        const bloqueDatos = [];
        let j = i + 1;
        while (j < datosLimpios.length) {
          const currentRow = datosLimpios[j];
          // parar si llegamos a otra Seccion o a otro encabezado
          const nextSeccion = extraerSeccion(currentRow);
          const nextCell0 = getTextFromCell(currentRow[0]).toLowerCase();
          const nextCell6 = getTextFromCell(currentRow[6]).toLowerCase();
          if (nextSeccion !== null || (nextCell0 === "nave" && nextCell6 === "nave")) break;

          if (currentRow.some(cell => cell !== "" && cell != null)) bloqueDatos.push(currentRow);
          j++;
        }
        // avanzar el índice principal
        i = j - 1;

        if (bloqueDatos.length > 0) {
          // rellenar columnas 0 y 6 como en la versión A
          let datosCompletos = rellenarColumna(bloqueDatos, 0);
          datosCompletos = rellenarColumna(datosCompletos, 6);

          let filaId = 0;
          datosCompletos.forEach(r => {
            // Asegurarnos de leer cada índice según la estructura:
            // [0]=Nave, [1]=Era, [2]=Variedad, [3]=Largo, [4]=Fecha_Siembra, [5]=Inicio_Corte
            // [6]=Nave (B), [7]=Era, [8]=Variedad, [9]=Largo, [10]=Fecha_Siembra, [11]=Inicio_Corte
            const aNave = getTextFromCell(r[0]) || "";
            const aEra = getTextFromCell(r[1]) || "";
            const aVar = getTextFromCell(r[2]) || "";
            const aLargo = getTextFromCell(r[3]) || "";
            const aFecha = getTextFromCell(r[4]) || "";
            const aInicio = getTextFromCell(r[5]) || "";

            const bNave = getTextFromCell(r[6]) || "";
            const bEra = getTextFromCell(r[7]) || "";
            const bVar = getTextFromCell(r[8]) || "";
            const bLargo = getTextFromCell(r[9]) || "";
            const bFecha = getTextFromCell(r[10]) || "";
            const bInicio = getTextFromCell(r[11]) || "";

            // Si la nave A está vacía, intentar rellenarla con la anterior (misma lógica que usabas)
            const lastNave = datosCrudos.length > 0 ? datosCrudos[datosCrudos.length - 1].Nave : "";
            const finalANave = aNave && aNave.trim() !== "" ? aNave : (lastNave || "NN");

            datosCrudos.push(
              { Seccion: seccionActual, Lado: "A", FilaId: filaId, Nave: finalANave, Era: aEra, Variedad: aVar, Largo: aLargo, Fecha_Siembra: aFecha, Inicio_Corte: aInicio },
              { Seccion: seccionActual, Lado: "B", FilaId: filaId, Nave: bNave, Era: bEra, Variedad: bVar, Largo: bLargo, Fecha_Siembra: bFecha, Inicio_Corte: bInicio }
            );
            filaId++;
          });
        }
      }
    } // fin for

    console.log("Parseo completado. Registros crudos:", datosCrudos.length);

    const datosFinales = datosCrudos.flatMap(expandirVariedades);
    console.log("Registros finales tras expandir variedades:", datosFinales.length);

    // Crear workbook final y hojas (usa tu módulo sheets)
    const wbFinal = new ExcelJS.Workbook();
    sheets.crearHojaDistribucionProductos(wbFinal, datosFinales);
    sheets.crearHojaDisbud(wbFinal, datosFinales);
    sheets.crearHojaGirasol(wbFinal, datosFinales);
    sheets.crearHojaPruebaFloracion(wbFinal, datosFinales);
    sheets.crearHojaNochesLuz(wbFinal, datosCrudos, { variedades });

    const wantPdf = String(req.query.format || '').toLowerCase() === 'pdf';

    if (!wantPdf) {
      finalReportPath = path.join('/tmp', `Reporte_Siembra_${semanaActual}_${Date.now()}.xlsx`);
      await wbFinal.xlsx.writeFile(finalReportPath);
      console.log("Reporte XLSX generado en:", finalReportPath);

      res.download(finalReportPath, `Reporte_Siembra_${semanaActual}.xlsx`, (err) => {
        if (err) {
          console.error("Error enviando el archivo:", err);
        }
        // borrar el reporte final después de enviar
        try { if (fs.existsSync(finalReportPath)) fs.unlinkSync(finalReportPath); } catch(e){ console.error(e); }
      });
    } else {
      if (!puppeteer) return res.status(400).json({ error: 'PDF generation not available: puppeteer not installed on the server.' });

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

      const htmlParts = ['<html><head><meta charset="utf-8"><style>table,td{font-family:Arial,sans-serif;font-size:10px;padding:4px}</style></head><body>'];
      wbFinal.eachSheet(sheet => {
        htmlParts.push(sheetToHTML(sheet));
      });
      htmlParts.push('</body></html>');
      const finalHtml = htmlParts.join('\n');

      const browser = await puppeteer.launch({ args: ['--no-sandbox', '--disable-setuid-sandbox'] });
      const page = await browser.newPage();
      await page.setContent(finalHtml, { waitUntil: 'networkidle0' });
      const pdfBuffer = await page.pdf({ format: 'A4', printBackground: true });
      await browser.close();

      const outputPdf = path.join('/tmp', `Reporte_Siembra_${semanaActual}_${Date.now()}.pdf`);
      fs.writeFileSync(outputPdf, pdfBuffer);
      res.setHeader('Content-Type', 'application/pdf');
      res.setHeader('Content-Disposition', `attachment; filename="Reporte_Siembra_${semanaActual}.pdf"`);
      res.sendFile(path.resolve(outputPdf), (err) => {
        try { if (fs.existsSync(outputPdf)) fs.unlinkSync(outputPdf); } catch(e){}
      });
    }

  } catch (error) {
    console.error("Error general procesando la solicitud:", error);
    res.status(500).json({ error: "Error procesando el archivo", detalle: error.message });
  } finally {
    // limpieza de temporales
    console.log("Limpieza de temporales...");
    try { if (originalUploadPath && fs.existsSync(originalUploadPath)) fs.unlinkSync(originalUploadPath); } catch(e){}
    try { if (convertedXlsxPath && fs.existsSync(convertedXlsxPath)) fs.unlinkSync(convertedXlsxPath); } catch(e){}
    try { if (pdfPath && fs.existsSync(pdfPath)) fs.unlinkSync(pdfPath); } catch(e){}
    console.log("Limpieza completada.");
  }
});

// SPA fallback (si usas frontend estático en dist)
app.use((req, res, next) => {
  if (req.path.startsWith('/api')) return next();
  const indexPath = path.join(__dirname, 'dist', 'index.html');
  if (fs.existsSync(indexPath)) return res.sendFile(indexPath);
  next();
});

const PORT = process.env.PORT || 5000;
app.listen(PORT, () => console.log(`🚀 Servidor corriendo en puerto ${PORT}`));
