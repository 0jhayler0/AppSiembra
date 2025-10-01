import express from "express";
import multer from "multer";
import cors from "cors";
import fs from "fs";
import pdf from "pdf-parse";
import ExcelJS from "exceljs";

const app = express();
const upload = multer({ dest: "uploads/" });

app.use(cors());

app.post("/upload", upload.single("file"), async (req, res) => {
  try {
    const dataBuffer = fs.readFileSync(req.file.path);
    const data = await pdf(dataBuffer);

    const lines = data.text.split("\n").map(l => l.trim()).filter(l => l);

    const workbook = new ExcelJS.Workbook();
    let worksheet = null;
    let currentLado = null;

    lines.forEach(line => {
      // 📌 Detectamos inicio de sección
      if (line.startsWith("Flores de la Victoria")) {
        const match = line.match(/Seccion:\s*(\d+)/);
        const sectionName = match ? `Sección ${match[1]}` : `Sección`;
        worksheet = workbook.addWorksheet(sectionName);

        // Encabezados
        worksheet.addRow([
          "Lado",
          "Nº",
          "Cama",
          "Variedad",
          "LargoNave",
          "Era",
          "Fecha Siembra",
          "Inicio Corte"
        ]);
      } else if (line.startsWith("Lado A")) {
        currentLado = "A";
      } else if (line.startsWith("Lado B")) {
        currentLado = "B";
      } else if (worksheet && currentLado) {
        // 📌 Parsear filas normales (ejemplo básico separando por espacios)
        const parts = line.split(/\s+/);

        // si la fila parece válida, la guardamos
        if (parts.length >= 4) {
          worksheet.addRow([
            currentLado,
            ...parts
          ]);
        }
      }
    });

    const outputPath = `uploads/output-${Date.now()}.xlsx`;
    await workbook.xlsx.writeFile(outputPath);

    res.download(outputPath, "resultado.xlsx", (err) => {
      if (err) console.error(err);
      fs.unlinkSync(req.file.path);
      fs.unlinkSync(outputPath);
    });
  } catch (error) {
    console.error(error);
    res.status(500).json({ error: "Error procesando el PDF" });
  }
});

app.listen(5000, () => console.log("✅ Servidor corriendo en http://localhost:5000"));
