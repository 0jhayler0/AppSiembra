import express from "express";
import multer from "multer";
import pdf from "pdf-parse";
import XLSX from "xlsx";
import fs from "fs";
import cors from "cors";

const app = express();
const upload = multer({ dest: "uploads/" });
app.use(cors());

function parsePDFText(text) {
  const lines = text.split("\n").map(l => l.trim()).filter(l => l !== "");
  const data = [];
  let currentSection = null;
  let currentLado = null;

  for (let line of lines) {
    if (line.includes("Seccion:")) {
      const secMatch = line.match(/Seccion:\s*(\d+)/);
      if (secMatch) currentSection = secMatch[1];
    }

    if (line.startsWith("Lado A")) {
      currentLado = "A";
      continue;
    }
    if (line.startsWith("Lado B")) {
      currentLado = "B";
      continue;
    }

    if (/^\d+\s/.test(line)) {
      const parts = line.split(/\s+/);

      if (currentLado === "A") {
        const [nave, era, ...rest] = parts;
        const largo = rest.pop();
        const variedad = rest.join(" ");

        data.push({
          seccion: currentSection,
          lado: "A",
          nave,
          era,
          variedad,
          largo
        });
      }

      if (currentLado === "B") {
        let nave = parts[0];
        let era = parts[1];
        let largo = parts[parts.length - 3];
        let fechaSiembra = parts[parts.length - 2];
        let inicioCorte = parts[parts.length - 1];
        let variedad = parts.slice(2, parts.length - 3).join(" ");

        data.push({
          seccion: currentSection,
          lado: "B",
          nave,
          era,
          variedad,
          largo,
          fechaSiembra,
          inicioCorte
        });
      }
    }
  }

  return data;
}

app.post("/upload", upload.single("file"), async (req, res) => {
    try {
        const filePath = req.file.path;
        const dataBuffer = fs.readFileSync(filePath);

        const pdfData = await pdf(dataBuffer);
        const text = pdfData.text;

        const rows = text
            .split("\n")
            .filter(line => line.trim() !== "")
            .map(line => ({ raw: line }));

        const ws = XLSX.utils.json_to_sheet(rows);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, "Datos");

        const excelBuffer = XLSX.write(wb, { type: "buffer", bookType: "xlsx" });

        res.setHeader("Content-Disposition", "attachment; filename=datos.xlsx");
        res.type(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        );
        res.send(excelBuffer);

        fs.unlinkSync(filePath);
    } catch (err) {
        console.error(err);
        res.status(500).json({ error: "Error procesando el PDF"});
    }
});

app.listen(5000, () => console.log("Servidor corriendo en http://localhost:5000"));