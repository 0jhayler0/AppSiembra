const express = require("express");
const multer = require("multer");
const XLSX = require("xlsx");
const fs = require("fs");
const cors = require("cors");
const ExcelJS = require("exceljs");

const app = express();
app.use(cors());

const upload = multer({ dest: "uploads/" });

// === 🔹 FUNCIONES DE LIMPIEZA Y CLASIFICACIÓN =======================
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
// === 🔹 FUNCIÓN: EXPANDIR VARIEDADES MULTIPLES =======================
function expandirVariedades(row) {
    const multiplesVariedades = row.Variedad;
    const multiplesLargos = row.Largo;

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

    if (nuevasFilas.length === 0) {
        return [{ ...row }]; 
    }
    
    return nuevasFilas;
}

function clasificarVariedad(nombre) {
  if (!nombre || typeof nombre !== "string") return "Desconocido";
  const lower = nombre.toLowerCase();
  if (lower.includes("prueba de floracion")) return "Prueba de Floracion";
  if (lower.includes("cremon") || lower.includes("spider") || lower.includes("anastasia") || lower.includes("towi"))
    return "Disbud";
  if (lower.includes("vincent choice") || lower.includes("girasol")) return "Girasol";
  return "Normal";
}

// === 🔹 FUNCIÓN: CREA HOJA ESPECIAL DISBUD =======================
function crearHojaDisbud(workbook, datos) {
  const gruposTemporales = {};
  let granTotalLargo = 0;
  datos.forEach(row => {
    const key = `${row.Seccion}_${row.Nave}_${row.Lado}_${row.Era}`;
    if (!gruposTemporales[key]) gruposTemporales[key] = [];
    gruposTemporales[key].push(row);
  });

  const gruposFinales = {};
  Object.entries(gruposTemporales).forEach(([key, rows]) => {
    const base = rows[0];
    let totalLargoDisbud = 0; 
    let tieneDisbud = false;
    const variedades = [];

    rows.forEach(row => {
      const tipo = clasificarVariedad(row.Variedad);
      const esDisbud = tipo === "Disbud";
      
      const largo = parseFloat(row.Largo);
      
      if (esDisbud && !isNaN(largo) && largo > 0) {
        totalLargoDisbud += largo;
        tieneDisbud = true;
      }
      
      variedades.push({
          nombre: row.Variedad,
          esDisbud: esDisbud,
          largo: !isNaN(largo) ? largo.toFixed(2) : '0'
      });
    });

    if (tieneDisbud) {
      gruposFinales[key] = {
        ...base,
        Variedades: variedades,
        TotalLargo: totalLargoDisbud 
      };

      granTotalLargo += totalLargoDisbud;
      
    }
  });
  let totalEnEras = granTotalLargo / 30;

  if (Object.keys(gruposFinales).length === 0) return;

// ----------------------------------------------------------------------------------
  // Crear la hoja "Disbud" en el workbook
  let sheet = workbook.getWorksheet("Disbud");
  if (sheet) workbook.removeWorksheet(sheet.id);
  sheet = workbook.addWorksheet("Disbud");

  sheet.columns = [
    { header: "Sección", key: "seccion", width: 10 },
    { header: "Nave", key: "nave", width: 8 },
    { header: "Lado", key: "lado", width: 8 },
    { header: "Era", key: "era", width: 8 },
    { header: "Variedades", key: "variedades", width: 60 },
    { header: "Total Largo", key: "total", width: 25 },
    { header: "Fecha Siembra", key: "siembra", width: 15 },
    { header: "Inicio Corte", key: "corte", width: 15 }
  ];

  const headerRow = sheet.getRow(1);
    headerRow.font = { bold: true };
    headerRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFCCCCCC' } };
    headerRow.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };

  Object.values(gruposFinales).forEach(g => {
    const fila = sheet.addRow({
      seccion: g.Seccion,
      nave: g.Nave,
      lado: g.Lado,
      era: g.Era,
      variedades: "", 
      total: g.TotalLargo.toFixed(2),
      siembra: g.Fecha_Siembra,
      corte: g.Inicio_Corte
    });

    const mostrarLargo = g.Variedades.length > 1;

    const celda = fila.getCell("variedades");
    celda.value = {
            richText: g.Variedades.flatMap((v, i) => {
                const textoVariedad = mostrarLargo 
                    ? `${v.nombre} (${v.largo})`
                    : v.nombre;
                
                return [
                    ...(i > 0 ? [{ text: " " }] : []),
                    { 
                        text: textoVariedad, 
                        // El color rojo solo se aplica si es Disbud
                        font: v.esDisbud ? { color: { argb: "FFFF0000" } } : {} 
                    }
                ];
            })
        };
    });
   
    const totalRow = sheet.addRow({}); // Fila vacía inicial
    
    // Unir las primeras 4 celdas (Sección, Nave, Lado, Era) y poner el texto "TOTAL"
    sheet.mergeCells(totalRow.number, 1, totalRow.number, 5);
    const totalCell = totalRow.getCell(1);
    totalCell.value = "TOTAL";
    totalCell.alignment = { horizontal: 'right' };
    
    // Poner el Gran Total en la columna "Total Largo (solo Disbud)"
    totalRow.getCell("total").value = granTotalLargo.toFixed(2);
    
    // Aplicar estilos a la fila de totales
    totalRow.font = { bold: true, size: 12 };
    totalRow.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFD9E1F2' } // Color de fondo azul muy claro para diferenciar
    };
    
    const totalErasRow = sheet.addRow({}); 
    
    sheet.mergeCells(totalErasRow.number, 1, totalErasRow.number, 5);
    const totalErasCell = totalErasRow.getCell(1);
    totalErasCell.value = "TOTAL ERAS";
    totalErasCell.alignment = { horizontal: 'right' };
    
    // Poner el Gran Total en la columna "Total Largo (solo Disbud)"
    totalErasRow.getCell("total").value = totalEnEras.toFixed(2);
    
    // Aplicar estilos a la fila de totales
    totalErasRow.font = { bold: true, size: 12 };
    totalErasRow.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFD9E1F2' } // Color de fondo azul muy claro para diferenciar
    };

}


// === 🔹 RUTA PRINCIPAL ===========================================
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

        //  Lógica para extraer datos crudos del Excel 
        for (let i = 0; i < datosLimpios.length; i++) {
            const row = datosLimpios[i];
            const nuevaSeccion = extraerSeccion(row);
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

                    datosCompletos.forEach(r => {
                        datosCrudos.push(
                            { Seccion: seccionActual, Lado: "A", Nave: r[0] || "", Era: r[1] || "", Variedad: r[2] || "", Largo: r[3] || "", Fecha_Siembra: r[4] || "", Inicio_Corte: r[5] || "" },
                            { Seccion: seccionActual, Lado: "B", Nave: r[6] || "", Era: r[7] || "", Variedad: r[8] || "", Largo: r[9] || "", Fecha_Siembra: r[10] || "", Inicio_Corte: r[11] || "" }
                        );
                    });
                }
            }
        }
        
        const datosFinales = datosCrudos.flatMap(expandirVariedades);
        // ----------------------------------------------------

        // === Crear libro combinado con ExcelJS ===
        const wbFinal = new ExcelJS.Workbook();

        // Crear hoja Disbud con formato
        crearHojaDisbud(wbFinal, datosFinales);

        // Guardar archivo final
        const outputPath = `Reporte_Siembra_${Date.now()}.xlsx`;
        await wbFinal.xlsx.writeFile(outputPath);

        console.log("✅ Reporte completo generado:", outputPath);

        res.download(outputPath, "Reporte_Siembra.xlsx", (err) => {
            if (err) console.error("Error enviando el archivo:", err);
            fs.unlinkSync(outputPath);
        });

        fs.unlinkSync(filePath);

    } catch (error) {
        console.error("❌ Error procesando Excel:", error);
        if (req.file && fs.existsSync(req.file.path)) fs.unlinkSync(req.file.path);
        res.status(500).json({ error: "Error procesando el archivo", detalle: error.message });
    }
});

app.listen(5000, () =>
  console.log("🚀 Servidor corriendo en http://localhost:5000")
);
