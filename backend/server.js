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
const cors = require("cors");
const axios = require("axios");
const FormData = require("form-data");

let puppeteer = null;
try { puppeteer = require("puppeteer"); } catch (e) {}

const app = express();

//--------------------------------------
// MULTER: Guardar siempre en /tmp
//--------------------------------------
const upload = multer({ dest: "/tmp" });

app.use(cors({
  origin: [
    "http://localhost:5173",
    "http://127.0.0.1:5173",
    "https://appsiembralavictoria.web.app",
    process.env.FRONTEND_URL || ""
  ],
  methods: ["GET", "POST", "OPTIONS"],
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
// EXPANSIÓN DE VARIEDADES
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
// RUTA PRINCIPAL: /upload-excel
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
    // SI ES PDF → CONVERTIR CON EL MICRO SERVICIO
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
    // LEER XLSX Y MATRIZ COMPLETA
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
    // PARSEO: MÉTODO B → "LADO A" / "LADO B"
    //--------------------------------------------
    const datosCrudos = [];
    let seccionActual = "N/A";
    let semanaActual = "N/A";

    for (let i = 0; i < datosLimpios.length; i++) {
      const row = datosLimpios[i];

      // Semana
      const sem = extraerSemana(row);
      if (sem) {
        semanaActual = sem;
        continue;
      }

      // Sección
      const sec = extraerSeccion(row);
      if (sec) {
        seccionActual = sec;
        continue;
      }

      // Encabezado de bloques A/B
      const c0 = getTextFromCell(row[0]).toLowerCase();
      const c6 = getTextFromCell(row[6]).toLowerCase();

      if (c0 === "lado a" && c6 === "lado b") {

        const bloqueDatos = [];
        let j = i + 1;

        while (j < datosLimpios.length) {
          const r = datosLimpios[j];

          const r0 = getTextFromCell(r[0]).toLowerCase();
          const r6 = getTextFromCell(r[6]).toLowerCase();
          const sec2 = extraerSeccion(r);

          // Nuevo bloque o nueva sección → parar
          if (sec2 || (r0 === "lado a" && r6 === "lado b")) break;

          if (r.some(x => x !== "" && x != null)) {
            bloqueDatos.push(r);
          }
          j++;
        }

        i = j - 1;

        if (bloqueDatos.length > 0) {
          let datosCompletos = rellenarColumna(bloqueDatos, 0);
          datosCompletos = rellenarColumna(datosCompletos, 6);

          // FIX: normalizar 12 columnas
          datosCompletos = datosCompletos.map(r => {
            const row = [...r];
            while (row.length < 12) row.push("");
            return row;
          });

          // Procesar A/B
          let filaId = 0;
          for (const r of datosCompletos) {
            const A = {
              Seccion: seccionActual,
              Lado: "A",
              FilaId: filaId,
              Nave: getTextFromCell(r[0]),
              Era: getTextFromCell(r[1]),
              Variedad: getTextFromCell(r[2]),
              Largo: getTextFromCell(r[3]),
              Fecha_Siembra: getTextFromCell(r[4]),
              Inicio_Corte: getTextFromCell(r[5])
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
              Inicio_Corte: getTextFromCell(r[11])
            };

            // FIX: Nave A vacía → copiar anterior
            if (!A.Nave) {
              const prev = datosCrudos.length > 0 ? datosCrudos[datosCrudos.length - 1].Nave : "";
              A.Nave = prev || "N/A";
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
    // GENERAR ARCHIVO XLSX FINAL
    //--------------------------------------------
    const reporteNombre = `Reporte_Siembra_${semanaActual}_${Date.now()}.xlsx`;
    finalReportPath = `/tmp/${reporteNombre}`;

    const wbOut = new ExcelJS.Workbook();
    const wsOut = wbOut.addWorksheet("Reporte Siembra");

    wsOut.columns = [
      { header: "Sección", key: "Seccion", width: 10 },
      { header: "Lado", key: "Lado", width: 6 },
      { header: "Nave", key: "Nave", width: 10 },
      { header: "Era", key: "Era", width: 8 },
      { header: "Variedad", key: "Variedad", width: 25 },
      { header: "Largo", key: "Largo", width: 8 },
      { header: "Fecha_Siembra", key: "Fecha_Siembra", width: 15 },
      { header: "Inicio_Corte", key: "Inicio_Corte", width: 15 }
    ];

    datosFinales.forEach(r => wsOut.addRow(r));

    await wbOut.xlsx.writeFile(finalReportPath);

    //--------------------------------------------
    // GENERAR PDF (si puppeteer está disponible)
    //--------------------------------------------
    let pdfBuffer = null;

    if (puppeteer) {
      const html = sheets.generarHTML(datosFinales, semanaActual);

      const browser = await puppeteer.launch({
        headless: "new",
        args: ["--no-sandbox", "--disable-setuid-sandbox"]
      });

      const page = await browser.newPage();
      await page.setContent(html, { waitUntil: "networkidle0" });

      pdfBuffer = await page.pdf({
        format: "A4",
        printBackground: true,
        margin: { top: "20px", bottom: "20px" }
      });

      await browser.close();
    }

    //--------------------------------------------
    // RESPONDER AL FRONTEND
    //--------------------------------------------
    res.json({
      ok: true,
      semana: semanaActual,
      registros: datosFinales.length,
      excel: {
        filename: reporteNombre,
        url: finalReportPath
      },
      pdf: pdfBuffer ? pdfBuffer.toString("base64") : null
    });

  } catch (error) {
    console.error("Error general procesando la solicitud:", error);
    res.status(500).json({ error: error.message || "Error procesando archivo" });

  } finally {
    //--------------------------------------------
    // LIMPIEZA DE TEMPORALES
    //--------------------------------------------
    try {
      if (originalUploadPath && fs.existsSync(originalUploadPath)) fs.unlinkSync(originalUploadPath);
      if (convertedXlsxPath && fs.existsSync(convertedXlsxPath)) fs.unlinkSync(convertedXlsxPath);
      if (pdfPath && fs.existsSync(pdfPath)) fs.unlinkSync(pdfPath);
    } catch (e) {
      console.log("Error limpiando temporales:", e);
    }
  }
});

//--------------------------------------
// START SERVER
//--------------------------------------
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Servidor corriendo en puerto ${PORT}`));
