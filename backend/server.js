const express = require("express");
const multer = require("multer");
const XLSX = require("xlsx");
const fs = require("fs");
const cors = require("cors");

const app = express();
app.use(cors());

// Configuración de multer
const upload = multer({ dest: "uploads/" });

// Función para limpiar filas vacías y espacios
const limpiarDatos = (data) => {
    return data
        .filter(row => row.some(cell => cell !== "" && cell != null))
        .map(row => row.map(cell => (typeof cell === "string" ? cell.trim() : cell)));
};

// Función para rellenar la columna "Nave" con el valor de arriba
const rellenarColumnaNave = (data, indexCol = 0) => {
    if (data.length === 0) return data;

    let lastValue = null;

    return data.map(row => {
        const valorActual = row[indexCol];
        if (valorActual !== null && valorActual !== undefined && String(valorActual).trim() !== "") {
            lastValue = valorActual;
        } else {
            row[indexCol] = lastValue;
        }
        return row;
    });
};

// Función para extraer número de sección de una fila
const extraerSeccion = (row) => {
    const textoSeccion = row.find(cell => typeof cell === 'string' && cell.includes('Seccion:'));
    if (textoSeccion) {
        const match = textoSeccion.match(/Seccion:\s*(\d+)/);
        if (match && match[1]) return match[1];
    }
    return null;
};

// Endpoint para subir y procesar Excel
app.post("/upload-excel", upload.single("file"), (req, res) => {
    try {
        if (!req.file) {
            return res.status(400).json({ error: "No se envió ningún archivo" });
        }

        const filePath = req.file.path;
        const workbook = XLSX.readFile(filePath);
        const sheetName = workbook.SheetNames[0];
        const data = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1 });

        let datosLimpios = limpiarDatos(data);
        const datosFinales = [];
        let seccionActual = "N/A";

        for (let i = 0; i < datosLimpios.length; i++) {
            const row = datosLimpios[i];

            const nuevaSeccion = extraerSeccion(row);
            if (nuevaSeccion) {
                seccionActual = nuevaSeccion;
                continue;
            }

            // Detecta inicio de bloque
            if (row[0] === 'Nave' && row[6] === 'Nave') {

                const bloqueDatos = [];
                let j = i + 1;

                while (j < datosLimpios.length) {
                    const currentRow = datosLimpios[j];
                    if (extraerSeccion(currentRow) !== null || (currentRow[0] === 'Nave' && currentRow[6] === 'Nave')) {
                        break;
                    }
                    if (currentRow.some(cell => cell !== "" && cell != null)) {
                        bloqueDatos.push(currentRow);
                    }
                    j++;
                }

                i = j - 1;

                if (bloqueDatos.length > 0) {
                    // Rellenamos la columna "Nave" del lado A (0) y B (6)
                    let datosCompletos = rellenarColumnaNave(bloqueDatos, 0); 
                    datosCompletos = rellenarColumnaNave(datosCompletos, 6);

                    // Convertimos cada fila en objetos lado A y B
                    datosCompletos.forEach(dataRow => {
                        // Lado A
                        datosFinales.push({
                            Seccion: seccionActual,
                            Lado: "A",
                            Nave: dataRow[0] || "",
                            Era: dataRow[1] || "",
                            Variedad: dataRow[2] || "",
                            Largo: dataRow[3] || "",
                            Fecha_Siembra: dataRow[4] || "",
                            Inicio_Corte: dataRow[5] || "",
                        });
                        // Lado B
                        datosFinales.push({
                            Seccion: seccionActual,
                            Lado: "B",
                            Nave: dataRow[6] || "",
                            Era: dataRow[7] || "",
                            Variedad: dataRow[8] || "",
                            Largo: dataRow[9] || "",
                            Fecha_Siembra: dataRow[10] || "",
                            Inicio_Corte: dataRow[11] || "",
                        });
                    });
                }
            }
        }

        // Mostrar solo primeras 50 filas para no saturar consola
        console.log(`\n=== DATOS FINALES (Total ${datosFinales.length} filas) ===`);
        console.table(datosFinales);

        // Contar secciones detectadas
        const secciones = new Set(datosFinales.map(row => row.Seccion));

        // Borrar archivo temporal
        fs.unlinkSync(filePath);

        res.json({
            message: "Archivo procesado y reestructurado correctamente",
            filas_total: datosFinales.length,
            secciones_detectadas: secciones.size,
            preview: datosFinales.slice(0, 5)
        });

    } catch (error) {
        console.error("Error procesando el Excel:", error);
        if (req.file && fs.existsSync(req.file.path)) {
            fs.unlinkSync(req.file.path);
        }
        res.status(500).json({ error: "Error procesando el archivo", detalle: error.message });
    }
});

app.listen(5000, () =>
    console.log("Servidor corriendo en http://localhost:5000")
);
