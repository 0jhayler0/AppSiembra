//--------------------------------------
// IMPORTS
//--------------------------------------
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

let puppeteer = null;
try { puppeteer = require("puppeteer"); } catch (e) {}

const app = express();

//--------------------------------------
// MULTER → Guardar SIEMPRE en /tmp
//--------------------------------------
const upload = multer({ dest: "/tmp" });

app.use(cors({
  origin: [
    "http://localhost:5173",
    "http://127.0.0.1:5173",
    "https://appsiembralavictoria.web.app",
    process.env.FRONTEND_URL || ""
  ],
  methods: ["GET","POST","OPTIONS"],
  exposedHeaders: ["Content-Disposition"]
}));

//--------------------------------------
// UTILS
//--------------------------------------
function getTextFromCell(cell) {
  if (cell == null) return "";
  if (typeof cell === "string" || typeof cell === "number") return String(cell).trim();
  if (cell.text) return String(cell.text).trim();
  if (cell.richText) return cell.richText.map(p => p.text).join("").trim();
  return "";
}

function limpiarDatos(data) {
  return data
    .filter(row => row.some(c => c !== null && c !== ""))
    .map(row => row.map(c => (typeof c === "string" ? c.trim() : c)));
}

function rellenarColumna(rows, indexCol) {
  let last = null;
  return rows.map(r => {
    if (r[indexCol] && String(r[indexCol]).trim() !== "") {
      last = r[indexCol];
    } else {
      r[indexCol] = last;
    }
    return r;
  });
}

function extraerSeccion(row) {
  for (const cell of row) {
    const txt = String(cell || "");
    const m = txt.match(/Seccion:\s*(\d+)/i);
    if (m) return m[1];
  }
  return null;
}

//--- FIX REAL: semana sí se detecta usando getTextFromCell -----
function extraerSemana(row) {
  for (const cell of row) {
    const texto = getTextFromCell(cell);
    if (!texto) continue;

    let m = texto.match(/Semana\s+Siembra\s+(2\d{5})/i);
    if (m) return m[1];

    m = texto.match(/Semana(?:\s+Siembra)?\s+(\d{1,2})\b/i);
    if (m) return m[1];

    m = texto.match(/\bSem\s+(\d{1,2})\b/i);
    if (m) return m[1];
  }
  return null;
}

//--------------------------------------
// EXPANDIR VARIEDADES
//--------------------------------------
function expandirVariedades(row) {
  const varText = row.Variedad || "";
  const regex = /(.+?)\s*\(([\d\.]+)\)/g;

  let match;
  const list = [];

  while ((match = regex.exec(varText)) !== null) {
    list.push({
      Seccion: row.Seccion,
      Lado: row.Lado,
      Nave: row.Nave,
      Era: row.Era,
      Variedad: match[1].trim(),
      Largo: match[2],
      Fecha_Siembra: row.Fecha_Siembra,
      Inicio_Corte: row.Inicio_Corte
    });
  }

  return list.length === 0 ? [row] : list;
}

//--------------------------------------
// RUTA: /upload-excel
//--------------------------------------
app.post("/upload-excel", upload.single("file"), async (req, res) => {
  let originalUploadPath = null;
  let convertedXlsxPath = null;
  let pdfPath = null;
  let finalReportPath = null;

  try {
    if (!req.file) return res.status(400).json({ error: "No file" });

    originalUploadPath = req.file.path;
    let filePath = originalUploadPath;

    const originalName = req.file.originalname || "";
    const ext = path.extname(originalName).toLowerCase();

    //--------------------------------------------
    // SI ES PDF → convertir con el microservicio
    //--------------------------------------------
    if (ext === ".pdf") {
      pdfPath = filePath + ".pdf";
      fs.renameSync(filePath, pdfPath);
      filePath = pdfPath;

      const out = filePath + ".converted.xlsx";
      const pythonURL = process.env.PYTHON_SERVICE_URL || "http://localhost:5001";

      const formData = new FormData();
      formData.append("file", fs.createReadStream(pdfPath));

      const response = await axios.post(`${pythonURL}/upload-excel`, formData, {
        responseType: "stream",
        headers: formData.getHeaders()
      });

      const writer = fs.createWriteStream(out);
      response.data.pipe(writer);
      await new Promise(res => writer.on("finish", res));

      filePath = out;
      convertedXlsxPath = out;
    }

    //--------------------------------------------
    // LEER XLSX Y ARMAR MATRIZ
    //--------------------------------------------
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);

    const ws = workbook.getWorksheet(1);
    const data = [];

    ws.eachRow((row) => {
      const r = [];
      row.eachCell((cell) => r.push(cell.value));
      data.push(r);
    });

    let datosLimpios = limpiarDatos(data);

    //--------------------------------------------
    // PARSEO DE BLOQUES (Método A, original)
    //--------------------------------------------
    const datosCrudos = [];
    let seccionActual = "N/A";
    let semanaActual = "N/A";

    for (let i = 0; i < datosLimpios.length; i++) {
      const row = datosLimpios[i];

      const sem = extraerSemana(row);
      if (sem) {
        semanaActual = sem;
        continue;
      }

      const sec = extraerSeccion(row);
      if (sec) {
        seccionActual = sec;
        continue;
      }

      const c0 = getTextFromCell(row[0]).toLowerCase();
      const c6 = getTextFromCell(row[6]).toLowerCase();

      if (c0 === "nave" && c6 === "nave") {
        const bloque = [];

        let j = i + 1;
        while (j < datosLimpios.length) {
          const r = datosLimpios[j];

          const nextSec = extraerSeccion(r);
          const n0 = getTextFromCell(r[0]).toLowerCase();
          const n6 = getTextFromCell(r[6]).toLowerCase();

          if (nextSec || (n0 === "nave" && n6 === "nave")) break;

          if (r.some(x => x !== "" && x != null)) bloque.push(r);
          j++;
        }

        i = j - 1;

        if (bloque.length > 0) {
          let full = rellenarColumna(bloque, 0);
          full = rellenarColumna(full, 6);

          //--------------------------------------------
          // FIX MÁS IMPORTANTE:
          // GARANTIZAR 12 COLUMNAS POR FILA
          //--------------------------------------------
          full = full.map(r => {
            const row = [...r];
            while (row.length < 12) row.push("");
            return row;
          });

          let filaId = 0;
          for (const r of full) {
            const A = {
              Seccion: seccionActual,
              Lado: "A",
              FilaId: filaId,
              Nave: getTextFromCell(r[0]),
              Era: getTextFromCell(r[1]),
              Variedad: getTextFromCell(r[2]),
              Largo: getTextFromCell(r[3]),
              Fecha_Siembra: getTextFromCell(r[4]),
              Inicio_Corte: getTextFromCell(r[5]),
            };
            const B = {
              Seccion: seccionActual,
              Lado: "B",
              FilaId: filaId,
              Nave: getTextFromCell(r[6]),
              Era: getTextFromCell(r[7]),
              Variedad: getTextFromCell(r[8]),
              Largo: getTextFromCell(r[9]),
              Fecha_Siembra: getTextFromCell(r[10]),
              Inicio_Corte: getTextFromCell(r[11]),
            };

            // FIX: si A.Nave viene vacío, usar la última nave válida
            if (!A.Nave) {
              const last = datosCrudos.length > 0 ? datosCrudos[datosCrudos.length - 1].Nave : "N/A";
              A.Nave = last;
            }

            datosCrudos.push(A, B);
            filaId++;
          }
        }
      }
    }

    //--------------------------------------------
    // EXPANDIR VARIEDADES
    //--------------------------------------------
    const datosFinales = datosCrudos.flatMap(expandirVariedades);

    //--------------------------------------------
    // CREAR REPORTE
    //--------------------------------------------
    const wbFinal = new ExcelJS.Workbook();
    sheets.crearHojaDistribucionProductos(wbFinal, datosFinales);
    sheets.crearHojaDisbud(wbFinal, datosFinales);
    sheets.crearHojaGirasol(wbFinal, datosFinales);
    sheets.crearHojaPruebaFloracion(wbFinal, datosFinales);
    sheets.crearHojaNochesLuz(wbFinal, datosCrudos, { variedades });

    const wantPdf = String(req.query.format || "").toLowerCase() === "pdf";

    // Normalizar semana
    const safeSemana = String(semanaActual || "NA").replace(/[\/\\]/g, "_");

    if (!wantPdf) {
      finalReportPath = `/tmp/Reporte_Siembra_${safeSemana}_${Date.now()}.xlsx`;
      await wbFinal.xlsx.writeFile(finalReportPath);

      return res.download(
        finalReportPath,
        `Reporte_Siembra_${safeSemana}.xlsx`,
        () => fs.unlinkSync(finalReportPath)
      );
    }

    //--------------------------------------------
    // GENERAR PDF
    //--------------------------------------------
    if (!puppeteer) {
      return res.status(400).json({ error: "Puppeteer no disponible." });
    }

    function sheetToHTML(sheet) {
      const rows = [];
      sheet.eachRow({ includeEmpty: true }, (row) => {
        const cells = [];
        row.eachCell({ includeEmpty: true }, (cell) => {
          const txt = cell.value == null ? "" : String(cell.value);
          cells.push(`<td>${txt}</td>`);
        });
        rows.push(`<tr>${cells.join("")}</tr>`);
      });
      return `<h2>${sheet.name}</h2><table border="1">${rows.join("")}</table><div style="page-break-after:always"></div>`;
    }

    let html = "<html><body>";
    wbFinal.eachSheet(sheet => html += sheetToHTML(sheet));
    html += "</body></html>";

    const browser = await puppeteer.launch({ args: ["--no-sandbox"] });
    const page = await browser.newPage();
    await page.setContent(html, { waitUntil: "networkidle0" });

    const pdf = await page.pdf({ format: "A4", printBackground: true });
    await browser.close();

    const outPDF = `/tmp/Reporte_Siembra_${safeSemana}_${Date.now()}.pdf`;
    fs.writeFileSync(outPDF, pdf);

    return res.sendFile(outPDF, () => fs.unlinkSync(outPDF));

  } catch (err) {
    console.error("ERROR:", err);
    return res.status(500).json({ error: "Error procesando", detalle: String(err) });

  } finally {
    try { if (originalUploadPath) fs.unlinkSync(originalUploadPath); } catch {}
    try { if (convertedXlsxPath) fs.unlinkSync(convertedXlsxPath); } catch {}
    try { if (pdfPath) fs.unlinkSync(pdfPath); } catch {}
  }
});

//--------------------------------------
// SPA fallback
//--------------------------------------
app.use((req, res, next) => {
  const index = path.join(__dirname, "dist", "index.html");
  if (fs.existsSync(index)) return res.sendFile(index);
  next();
});

//--------------------------------------
// START SERVER
//--------------------------------------
const PORT = process.env.PORT || 5000;
app.listen(PORT, () => console.log("🔥 Servidor corriendo en", PORT));
