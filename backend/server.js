const express = require("express");
const multer = require("multer");
const ExcelJS = require("exceljs");
const fs = require("fs");
const path = require("path");
const variedades = require("./variedades");
const sheets = require("./sheets");

const app = express();
const upload = multer({ dest: '/tmp' });
const cors = require('cors');
const axios = require('axios');

app.use(cors({
  origin: ['http://localhost:5173', 'http://127.0.0.1:5173', 'https://appsiembralavictoria.web.app', process.env.FRONTEND_URL || ''], 
  methods: ['GET','POST','OPTIONS'],
  exposedHeaders: ['Content-Disposition']
}));


let puppeteer = null;
try {
  puppeteer = require('puppeteer');
} catch (e) {
}

// === FUNCIONES AUXILIARES =======================

const getTextFromCell = (cell) => {
  if (!cell) return "";
  if (typeof cell === 'string' || typeof cell === 'number') {
    return String(cell);
  }
  if (cell.richText && Array.isArray(cell.richText)) {
    return cell.richText.map(part => part.text).join('');
  }
  return "";
};

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



// ========== RUTA PRINCIPAL (CORREGIDA CON LOGS) ====================================================================

app.post("/upload-excel", upload.single("file"), async (req, res) => {
  // Variables para limpiar archivos al final
  let originalUploadPath = null;
  let convertedXlsxPath = null;
  let pdfPath = null;
  let finalReportPath = null;

  try {
    if (!req.file) return res.status(400).json({ error: "No se envió ningún archivo" });

    originalUploadPath = req.file.path;
    console.log("PASO 1: Archivo recibido.");
    console.log(" - Ruta temporal:", originalUploadPath);
    console.log(" - Nombre original:", req.file.originalname);

    let filePath = originalUploadPath;
    const originalName = req.file.originalname || '';
    const ext = path.extname(originalName).toLowerCase();

    if (ext === '.pdf' || req.file.mimetype === 'application/pdf') {
      console.log("PASO 2: Detectado archivo PDF, iniciando conversión...");
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
          headers: { 'Content-Type': 'multipart/form-data' }
        });

        const writer = fs.createWriteStream(outXlsx);
        response.data.pipe(writer);

        await new Promise((resolve, reject) => {
          writer.on('finish', resolve);
          writer.on('error', reject);
        });

        filePath = outXlsx;
        convertedXlsxPath = outXlsx;
        console.log("PASO 2b: PDF convertido a XLSX exitosamente en:", convertedXlsxPath);
      } catch (err) {
        console.error('Error convirtiendo PDF a XLSX:', err);
        return res.status(500).json({ error: 'Error convirtiendo PDF a XLSX', detalle: String(err) });
      }
    }

    console.log("PASO 3: Leyendo archivo Excel desde:", filePath);
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    const worksheet = workbook.getWorksheet(1);
    const data = [];
    worksheet.eachRow((row, rowNumber) => {
      data.push(row.values); // row.values es más directo que iterar celda por celda
    });
    
    let datosLimpios = limpiarDatos(data);
    console.log("PASO 4: Datos limpios (primeras 5 filas):", JSON.stringify(datosLimpios.slice(0, 5), null, 2));

    const datosCrudos = [];
    let seccionActual = "N/A";
    let semanaActual = "N/A";

    const extraerSemana = (row) => {
      if (!Array.isArray(row)) return null;
      for (const cell of row) {
        const texto = getTextFromCell(cell); // <-- USA LA NUEVA FUNCIÓN AQUÍ
        if (!texto) continue;
        
        const match = texto.match(/Semana\s+Siembra\s+(2\d{5})/i);
        if (match) {
          console.log("🗓 Semana encontrada:", match[1]);
          return match[1];
        }
      }
  return null;
};

    console.log("PASO 5: Iniciando parseo de datos crudos...");
    for (let i = 0; i < datosLimpios.length; i++) {
    const row = datosLimpios[i];

    // USAMOS LA NUEVA FUNCIÓN PARA EXTRAER TEXTO DE LAS CELDAS
    const cell0_text = getTextFromCell(row[0]);
    const cell6_text = getTextFromCell(row[6]);

    const nuevaSemana = extraerSemana(row);
    const nuevaSeccion = extraerSeccion(row);

    if (nuevaSemana) {
        semanaActual = nuevaSemana;
    }

    if (nuevaSeccion) {
        seccionActual = nuevaSeccion;
        console.log(`Sección actualizada a: ${seccionActual}`);
        continue;
    }

    // AHORA LA COMPARACIÓN FUNCIONARÁ
    if (cell0_text === "Nave" && cell6_text === "Nave") {
        console.log(`🎯 BLOQUE 'Nave' ENCONTRADO en la fila ${i + 1}. Procesando...`);
        
        const bloqueDatos = [];
        let j = i + 1;
        while (j < datosLimpios.length) {
            const currentRow = datosLimpios[j];
            const currentCell0_text = getTextFromCell(currentRow[0]);
            const currentCell6_text = getTextFromCell(currentRow[6]);

            if (extraerSeccion(currentRow) !== null || (currentCell0_text === "Nave" && currentCell6_text === "Nave")) break;
            if (currentRow.some(cell => cell !== "" && cell != null)) bloqueDatos.push(currentRow);
            j++;
        }
        i = j - 1;

        if (bloqueDatos.length > 0) {
            console.log(`   -> Se encontraron ${bloqueDatos.length} filas de datos en este bloque.`);
            let datosCompletos = rellenarColumna(bloqueDatos, 0);
            datosCompletos = rellenarColumna(datosCompletos, 6);

            let filaId = 0;
            datosCompletos.forEach(r => {
                // AQUÍ TAMBIÉN USAMOS LA FUNCIÓN POR SI HAY CELDAS CON FORMATO
                datosCrudos.push(
                    { Seccion: seccionActual, Lado: "A", FilaId: filaId, Nave: getTextFromCell(r[0]), Era: getTextFromCell(r[1]), Variedad: getTextFromCell(r[2]), Largo: getTextFromCell(r[3]), Fecha_Siembra: getTextFromCell(r[4]), Inicio_Corte: getTextFromCell(r[5]) },
                    { Seccion: seccionActual, Lado: "B", FilaId: filaId, Nave: getTextFromCell(r[6]), Era: getTextFromCell(r[7]), Variedad: getTextFromCell(r[8]), Largo: getTextFromCell(r[9]), Fecha_Siembra: getTextFromCell(r[10]), Inicio_Corte: getTextFromCell(r[11]) }
                );
                filaId++;
            });
        }
    }
}
    
    console.log(`PASO 6: Parseo completado. Total de datos crudos extraídos: ${datosCrudos.length}`);
    if (datosCrudos.length === 0) {
        console.warn("ADVERTENCIA: No se extrajeron datos crudos. El reporte estará vacío. Revisa el formato de tu Excel y los logs de PASO 5.");
    }

    const datosFinales = datosCrudos.flatMap(expandirVariedades);
    console.log(`PASO 7: Datos finales después de expandir variedades: ${datosFinales.length}`);

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
      console.log("PASO 8: Reporte XLSX generado en:", finalReportPath);
      res.download(finalReportPath, `Reporte_Siembra_${semanaActual}.xlsx`, (err) => {
        if (err) console.error("Error enviando el archivo:", err);
      });
    } else {
      // ... (la lógica de PDF también usaría path.join('/tmp', ...)) ...
      // Para no alargar más, asumimos que ya sabes cómo cambiar la ruta aquí también.
      if (!puppeteer) return res.status(400).json({ error: 'PDF generation not available' });
      // ... resto del código de PDF ...
    }

  } catch (error) {
    console.error("Error general procesando la solicitud:", error);
    res.status(500).json({ error: "Error procesando el archivo", detalle: error.message });
  } finally {
    // Limpieza de archivos temporales
    console.log("PASO 9: Limpiando archivos temporales...");
    if (originalUploadPath && fs.existsSync(originalUploadPath)) fs.unlinkSync(originalUploadPath);
    if (convertedXlsxPath && fs.existsSync(convertedXlsxPath)) fs.unlinkSync(convertedXlsxPath);
    if (pdfPath && fs.existsSync(pdfPath)) fs.unlinkSync(pdfPath);
    if (finalReportPath && fs.existsSync(finalReportPath)) fs.unlinkSync(finalReportPath);
    console.log("Limpieza completada.");
  }
});

const PORT = process.env.PORT || 5000;
app.listen(PORT, () => console.log(`🚀 Servidor corriendo en puerto ${PORT}`));

