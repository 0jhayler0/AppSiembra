const express = require("express");
const multer = require("multer");
const XLSX = require("xlsx");
const fs = require("fs");
const cors = require("cors");
const ExcelJS = require("exceljs");

const app = express();
app.use(cors());
app.use((req, res, next) => {
  res.header("Access-Control-Expose-Headers", "Content-Disposition");
  next();
});

const upload = multer({ dest: "uploads/" });

const variedades = require("./variedades");

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
  if (lower.includes("cremon") || lower.includes("spider") || lower.includes("anastasia") || lower.includes("tutu"))
    return "Disbud";
  if (lower.includes("vincent choice") || lower.includes("girasol")) return "Girasol";
  return "Normal";
}

// ========== FUNCIÓN: HOJA ESPECIAL DISBUD ================================================

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
                        font: v.esDisbud ? { color: { argb: "FFFF0000" } } : {} 
                    }
                ];
            })
        };
    });
   
    const totalRow = sheet.addRow({}); 
    
    sheet.mergeCells(totalRow.number, 1, totalRow.number, 5);
    const totalCell = totalRow.getCell(1);
    totalCell.value = "TOTAL";
    totalCell.alignment = { horizontal: 'right' };
    
    totalRow.getCell("total").value = granTotalLargo.toFixed(2);
    
    totalRow.font = { bold: true, size: 12 };
    totalRow.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFD9E1F2' } 
    };
    
    const totalErasRow = sheet.addRow({}); 
    
    sheet.mergeCells(totalErasRow.number, 1, totalErasRow.number, 5);
    const totalErasCell = totalErasRow.getCell(1);
    totalErasCell.value = "TOTAL ERAS";
    totalErasCell.alignment = { horizontal: 'right' };
    
    totalErasRow.getCell("total").value = totalEnEras.toFixed(2);
    
    totalErasRow.font = { bold: true, size: 12 };
    totalErasRow.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFD9E9F2' } 
    };

}


// ========== FUNCIÓN: HOJA ESPECIAL DISTRUBUCION DE LOS PRODUCTOS ============================

function crearHojaDistribucionProductos(workbook, datos) {
  const metrosPorSeccion = {};
  const navesPorSeccion = {}; 

  datos.forEach(row => {
    const seccion = row.Seccion || "Sin Sección";
    const tipo = clasificarVariedad(row.Variedad);
    const largo = parseFloat(row.Largo) || 0;
    const nave = row.Nave?.toString().trim() || "Sin Nave";

  
    if (!metrosPorSeccion[seccion]) {
      metrosPorSeccion[seccion] = {
        Disbud: 0,
        Girasol: 0,
        Normal: 0,
        "Prueba de Floracion": 0
      };
    }
    if (!navesPorSeccion[seccion]) {
      navesPorSeccion[seccion] = new Set(); 
    }

    if (metrosPorSeccion[seccion][tipo] !== undefined) {
      metrosPorSeccion[seccion][tipo] += largo;
    }

    if (nave !== "Sin Nave") {
      navesPorSeccion[seccion].add(nave);
    }
  });

  const metrosPorSeccionEnEras = {};
  for (const seccion in metrosPorSeccion) {
    metrosPorSeccionEnEras[seccion] = {};
    for (const tipo in metrosPorSeccion[seccion]) {
      metrosPorSeccionEnEras[seccion][tipo] = (metrosPorSeccion[seccion][tipo] / 30).toFixed(2);
    }
  }

  let sheet = workbook.getWorksheet("Distribución Productos");
  if (sheet) workbook.removeWorksheet(sheet.id);
  sheet = workbook.addWorksheet("Distribución Productos");

  sheet.columns = [
    { header: "Sección", key: "seccion", width: 15 },
    { header: "Naves", key: "naves", width: 25 }, 
    { header: "Eras", key: "eras", width: 15 },
    { header: "Pompon", key: "pompon", width: 15 },
    { header: "Disbud", key: "disbud", width: 15 },
    { header: "Girasol", key: "girasol", width: 15 },
    { header: "Prueba de Floración", key: "floracion", width: 20 },
    { header: "Total", key: "total", width: 15 },
  ];

  
  Object.entries(metrosPorSeccionEnEras).forEach(([seccion, valores]) => {
    const naves = Array.from(navesPorSeccion[seccion] || []).sort().join(", "); 

    sheet.addRow({
      seccion: seccion,
      naves: naves,
      eras: (
        parseFloat(valores.Disbud || 0) +
        parseFloat(valores.Girasol || 0) +
        parseFloat(valores.Normal || 0) +
        parseFloat(valores["Prueba de Floracion"] || 0)
      ).toFixed(2),
      pompon: valores.Normal,
      disbud: valores.Disbud,
      girasol: valores.Girasol,
      floracion: valores["Prueba de Floracion"],
      total: (
        parseFloat(valores.Disbud || 0) +
        parseFloat(valores.Girasol || 0) +
        parseFloat(valores.Normal || 0) +
        parseFloat(valores["Prueba de Floracion"] || 0)
      ).toFixed(2),
    });
  });

  sheet.getRow(1).font = { bold: true };

  const totalDisbud = sheet.getColumn("disbud").values.slice(2).reduce((a, b) => a + (parseFloat(b) || 0), 0);
  const totalGirasol = sheet.getColumn("girasol").values.slice(2).reduce((a, b) => a + (parseFloat(b) || 0), 0);
  const totalPompon = sheet.getColumn("pompon").values.slice(2).reduce((a, b) => a + (parseFloat(b) || 0), 0);
  const totalFloracion = sheet.getColumn("floracion").values.slice(2).reduce((a, b) => a + (parseFloat(b) || 0), 0);
  const totalGlobal = totalDisbud + totalGirasol + totalPompon + totalFloracion;

  const totalRow = sheet.addRow({
    seccion: "TOTAL GENERAL",
    naves: "",
    eras: totalGlobal.toFixed(2),
    pompon: totalPompon.toFixed(2),
    disbud: totalDisbud.toFixed(2),
    girasol: totalGirasol.toFixed(2),
    floracion: totalFloracion.toFixed(2),
    total: totalGlobal.toFixed(2),
  });

  const headerRow = sheet.getRow(1);
    headerRow.font = { bold: true };
    headerRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFCCCCCC' } };
    headerRow.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    
  totalRow.font = { bold: true, size: 12 };
  totalRow.fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FFCCE5FF" },
  };
  totalRow.alignment = { horizontal: "center" };
}


// ========== FUNCIÓN: HOJA ESPECIAL GIRASOL ====================================================

function crearHojaGirasol(workbook, datos) {
  const girasolPorSeccion = {};
  const navesPorSeccion = {};
  const fechasPorSeccion = {};

  datos.forEach(row => {
    const tipo = clasificarVariedad(row.Variedad);
    if (tipo !== "Girasol") return;

    const seccion = row.Seccion || "Sin Sección";
    const nave = row.Nave?.toString().trim() || "Sin Nave";
    const largo = parseFloat(row.Largo) || 0;
    const fechaSiembra = row.Fecha_Siembra || "";
    const inicioCorte = row.Inicio_Corte || "";

    if (!girasolPorSeccion[seccion]) {
      girasolPorSeccion[seccion] = { metros: 0 };
      navesPorSeccion[seccion] = new Set();
      fechasPorSeccion[seccion] = { siembra: new Set(), corte: new Set() };
    }

    girasolPorSeccion[seccion].metros += largo;
    if (nave !== "Sin Nave") navesPorSeccion[seccion].add(nave);
    if (fechaSiembra) fechasPorSeccion[seccion].siembra.add(fechaSiembra);
    if (inicioCorte) fechasPorSeccion[seccion].corte.add(inicioCorte);
  });

  const monthMap = {
    ene:0, "ene.":0, enero:0, jan:0, january:0,
    feb:1, "feb.":1, febrero:1, february:1,
    mar:2, "mar.":2, marzo:2, march:2,
    abr:3, "abr.":3, abril:3, apr:3, april:3,
    may:4, mayo:4,
    jun:5, "jun.":5, junio:5, june:5,
    jul:6, julio:6, july:6,
    ago:7, "ago.":7, agosto:7, aug:7, august:7,
    sep:8, "sep.":8, septiembre:8, sept:8, september:8,
    oct:9, "oct.":9, octubre:9, october:9,
    nov:10, "nov.":10, noviembre:10, november:10,
    dic:11, "dic.":11, diciembre:11, dec:11, december:11
  };

  function parseInicioCorteToDate(token, fallbackYear) {
    if (!token && token !== 0) return null;
    if (token instanceof Date && !isNaN(token)) return token;
    const s = String(token).trim().toLowerCase();
    const m = s.match(/^(\d{1,2})\s*[-\/\s\.]?\s*([a-zñ\.]+)\b/);
    if (m) {
      const day = parseInt(m[1], 10);
      const monToken = m[2].replace(/\.$/, "");
      const monthIndex = monthMap[monToken];
      if (monthIndex !== undefined) {
        const year = Number(fallbackYear) || (new Date()).getFullYear();
        const d = new Date(year, monthIndex, day);
        if (!isNaN(d)) return d;
      }
    }
    const iso = s.match(/(\d{4}-\d{2}-\d{2})/);
    if (iso) {
      const d = new Date(iso[1]);
      if (!isNaN(d)) return d;
    }
    const dmy = s.match(/(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})/);
    if (dmy) {
      let day = parseInt(dmy[1],10), month = parseInt(dmy[2],10)-1, year = parseInt(dmy[3],10);
      if (year < 100) year += 2000;
      const d = new Date(year, month, day);
      if (!isNaN(d)) return d;
    }
    const dFallback = new Date(s);
    if (!isNaN(dFallback)) return dFallback;
    return null;
  }

  function getISOWeek(date) {
    const d = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
    const dayNum = d.getUTCDay() || 7;
    d.setUTCDate(d.getUTCDate() + 4 - dayNum);
    const yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
    return Math.ceil((((d - yearStart) / 86400000) + 1) / 7);
  }

  let sheet = workbook.getWorksheet("Girasol");
  if (sheet) workbook.removeWorksheet(sheet.id);
  sheet = workbook.addWorksheet("Girasol");

  sheet.columns = [
    { header: "Sección", key: "seccion", width: 15 },
    { header: "Nave", key: "naves", width: 25 },
    { header: "Eras", key: "eras", width: 15 },
    { header: "Fecha Siembra", key: "fechaSiembra", width: 25 },
    { header: "Inicio Corte", key: "inicioCorte", width: 25 },
    { header: "Semana Corte", key: "semanaCorte", width: 25 },
    { header: "Estimado Producción", key: "estimado", width: 20 },
  ];

  Object.entries(girasolPorSeccion).forEach(([seccion, info]) => {
    const metros = info.metros || 0;
    const eras = (metros / 30).toFixed(2);
    const estimado = Math.round(eras * 850);

    const naves = Array.from(navesPorSeccion[seccion] || []).sort().join(", ");
    const fechasSiembraArr = Array.from(fechasPorSeccion[seccion].siembra).filter(Boolean);
    const fechasCorteArr = Array.from(fechasPorSeccion[seccion].corte).filter(Boolean);

    let defaultYear = (new Date()).getFullYear();
    if (fechasSiembraArr.length > 0) {
      const tryDate = new Date(fechasSiembraArr[0]);
      if (!isNaN(tryDate)) defaultYear = tryDate.getFullYear();
      else {
        const m = String(fechasSiembraArr[0]).match(/(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})/);
        if (m) { let y = parseInt(m[3],10); if (y < 100) y += 2000; defaultYear = y; }
      }
    }

    const semanas = fechasCorteArr
      .map(fc => parseInicioCorteToDate(fc, defaultYear))
      .filter(d => d instanceof Date && !isNaN(d))
      .map(d => getISOWeek(d));

    const semanasUnicas = Array.from(new Set(semanas)).sort((a,b) => a-b);

    const semanaCorte = semanasUnicas.length > 0 ? String(Math.max(...semanasUnicas)).padStart(2, "0") : "";

    sheet.addRow({
      seccion,
      naves,
      eras,
      fechaSiembra: fechasSiembraArr.join(", "),
      inicioCorte: fechasCorteArr.join(", "),
      semanaCorte,
      estimado
    });
  });

  // totales
  const totalMetros = Object.values(girasolPorSeccion).reduce((acc, v) => acc + v.metros, 0);
  const totalEras = (totalMetros / 30).toFixed(2);
  const totalEstimado = Math.round(totalEras * 850);

  const headerRow = sheet.getRow(1);
    headerRow.font = { bold: true };
    headerRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFCCCCCC' } };
    headerRow.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
  

  const totalRow = sheet.addRow({
    seccion: "TOTAL GENERAL",
    naves: "",
    eras: totalEras,
    semanaCorte: "",
    estimado: totalEstimado
  });

  totalRow.font = { bold: true, size: 12 };
  totalRow.fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FFFFE699" }
  };
  totalRow.alignment = { horizontal: "center" };

  sheet.getRow(1).font = { bold: true };
}

// ========== FUNCIÓN: HOJA ESPECIAL PRUEBA DE FLORACION =======================================

function crearHojaPruebaFloracion(workbook, datos) {
  const floracionData = datos.filter(row => clasificarVariedad(row.Variedad) === "Prueba de Floracion");
  if (floracionData.length === 0) return;

  const grupos = {};
  floracionData.forEach(row => {
    const seccion = row.Seccion || "Sin Sección";
    const nave = row.Nave?.toString().trim() || "Sin Nave";
    const era = row.Era || "Sin Era";
    const lado = row.Lado || ""; // A o B
    const key = `${seccion}__${nave}__${era}__${lado}`;

    if (!grupos[key]) {
      grupos[key] = {
        seccion,
        nave,
        era,
        lado,
        metros: 0,
        fechaSiembra: row.Fecha_Siembra || "",
        inicioCorte: row.Inicio_Corte || "",
        variedades: new Set()
      };
    }

    grupos[key].metros += parseFloat(row.Largo) || 0;
    if (row.Variedad) grupos[key].variedades.add(row.Variedad);
    if (!grupos[key].fechaSiembra && row.Fecha_Siembra) grupos[key].fechaSiembra = row.Fecha_Siembra;
    if (!grupos[key].inicioCorte && row.Inicio_Corte) grupos[key].inicioCorte = row.Inicio_Corte;
  });

  let sheet = workbook.getWorksheet("Prueba de Floración");
  if (sheet) workbook.removeWorksheet(sheet.id);
  sheet = workbook.addWorksheet("Prueba de Floración");

  sheet.columns = [
    { header: "Sección", key: "seccion", width: 15 },
    { header: "Nave", key: "nave", width: 12 },
    { header: "Lado", key: "lado", width: 8 },
    { header: "Era", key: "era", width: 14 },
    { header: "Variedades", key: "variedades", width: 40 },
    { header: "Metros", key: "metros", width: 12 },
    { header: "Eras", key: "eras", width: 12 },
    { header: "Fecha Siembra", key: "fechaSiembra", width: 18 },
    { header: "Inicio Corte", key: "inicioCorte", width: 18 },
  ];

  Object.values(grupos).forEach(info => {
    const variedades = Array.from(info.variedades).sort().join(", ");
    const metros = info.metros;
    const eras = (metros / 30).toFixed(2);

    sheet.addRow({
      seccion: info.seccion,
      nave: info.nave,
      lado: info.lado,
      era: info.era,
      variedades,
      metros: metros.toFixed(2),
      eras,
      fechaSiembra: info.fechaSiembra,
      inicioCorte: info.inicioCorte,
    });
  });

  const totalMetros = Object.values(grupos).reduce((acc, g) => acc + g.metros, 0);
  const totalEras = (totalMetros / 30).toFixed(2);

  const headerRow = sheet.getRow(1);
    headerRow.font = { bold: true };
    headerRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFCCCCCC' } };
    headerRow.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };

  const totalRow = sheet.addRow({
    seccion: "TOTAL GENERAL",
    metros: totalMetros.toFixed(2),
    eras: totalEras,
  });

  totalRow.font = { bold: true, size: 12 };
  totalRow.fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FFFFE699" },
  };
  totalRow.alignment = { horizontal: "center" };

  sheet.getRow(1).font = { bold: true };
}

// ========== FUNCIÓN: HOJA ESPECIAL NOCHES DE LUZ ==============================================

function crearHojaNochesLuz(workbook, datos) {
  if (!Array.isArray(datos) || datos.length === 0) return;

  const nochesMap = new Map(variedades.map(v => [String(v.nombre).trim().toLowerCase(), Number(v.nochesLuz) || 17]));
  const defecto = 17;

  const obtenerNochesDeCelda = (variedadCell) => {
    if (!variedadCell || String(variedadCell).trim() === "") return defecto;
    const partes = String(variedadCell)
      .split(/[,\/\\\|\;\+\&]| y |;/i)
      .map(p => p.trim())
      .filter(Boolean);
    const candidatos = partes.length ? partes : [String(variedadCell).trim()];
    const nochesEncontradas = candidatos.map(p => {
      const key = p.toLowerCase();
      return nochesMap.has(key) ? nochesMap.get(key) : defecto;
    });
    return Math.max(...nochesEncontradas);
  };

  function parseFechaSiembra(value) {
    if (value == null || value === "") return null;
    if (value instanceof Date && !isNaN(value)) return value;

    if (typeof value === "number") {
      try {
        const utcDays = Math.floor(value - 25569);
        const utcValue = utcDays * 86400;
        const dateInfo = new Date(utcValue * 1000);
        const fractionalDay = value - Math.floor(value);
        const totalSeconds = Math.round(86400 * fractionalDay);
        dateInfo.setSeconds(dateInfo.getSeconds() + totalSeconds);
        if (!isNaN(dateInfo)) return dateInfo;
      } catch (e) {}
    }
    

    const s = String(value).trim();
    const dIso = new Date(s);
    if (!isNaN(dIso)) return dIso;

    const m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/);
    if (m) {
      let day = parseInt(m[1], 10);
      let month = parseInt(m[2], 10) - 1;
      let year = parseInt(m[3], 10);
      if (year < 100) year += 2000;
      const d2 = new Date(year, month, day);
      if (!isNaN(d2)) return d2;
    }

    const monthMap = {
      ene:0, enero:0, feb:1, febrero:1, mar:2, marzo:2, abr:3, abril:3, may:4, mayo:4,
      jun:5, junio:5, jul:6, julio:6, ago:7, agosto:7, sep:8, septiembre:8, oct:9, octubre:9,
      nov:10, noviembre:10, dic:11, diciembre:11
    };

    const m2 = s.toLowerCase().match(/^(\d{1,2})\s*[-\/\s\.]?\s*([a-zñ\.]+)/);
    if (m2) {
      const day = parseInt(m2[1], 10);
      const monToken = m2[2].replace(/\.$/, "");
      const monthIndex = monthMap[monToken];
      if (monthIndex !== undefined) {
        const year = (new Date()).getFullYear();
        const d3 = new Date(year, monthIndex, day);
        if (!isNaN(d3)) return d3;
      }
    }

    return null;
  }

  const formatDateDDMMYYYY = (d) => {
    if (!d || !(d instanceof Date) || isNaN(d)) return "";
    const dd = String(d.getDate()).padStart(2, "0");
    const mm = String(d.getMonth() + 1).padStart(2, "0");
    const yyyy = d.getFullYear();
    return `${dd}-${mm}-${yyyy}`;
  };

  const grupos = {};
  datos.forEach(r => {
    const key = `${r.Seccion || ""}__${r.FilaId ?? ""}`;
    if (!grupos[key]) grupos[key] = { A: null, B: null, Seccion: r.Seccion || "" };
    if (r.Lado === "A") grupos[key].A = r;
    if (r.Lado === "B") grupos[key].B = r;
  });

  let sheet = workbook.getWorksheet("Noches de Luz");
  if (sheet) workbook.removeWorksheet(sheet.id);
  sheet = workbook.addWorksheet("Noches de Luz");

  sheet.columns = [
    { header: "Sección", key: "Seccion", width: 8 },
    { header: "Nave", key: "NaveA", width: 8 },
    { header: "Era", key: "EraA", width: 8 },
    { header: "Variedad", key: "VarA", width: 55 },
    { header: "Fecha Siembra", key: "SiembraA", width: 15 },
    { header: "Noche Final", key: "NochesA", width: 18 },
    { header: "Inicio Corte", key: "CorteA", width: 15 },
    { header: " ", key: "espacioEnBlanco", width: 5 },
    { header: "Nave", key: "NaveB", width: 8 },
    { header: "Era", key: "EraB", width: 8 },
    { header: "Variedad", key: "VarB", width: 55 },
    { header: "Fecha Siembra", key: "SiembraB", width: 15 },
    { header: "Noche Final", key: "NochesB", width: 18 },
    { header: "Inicio Corte", key: "CorteB", width: 15 }
  ];

  sheet.mergeCells('F');
  const nochesLadoACell = sheet.getColumn('F');
  nochesLadoACell.font = { color: { argb: 'FFFF0000' }};
  nochesLadoACell.alignment = { horizontal: 'center', vertical: 'middle' };
  
  sheet.mergeCells('M');
  const nochesLadoBCell = sheet.getColumn('M');
  nochesLadoBCell.font = { color: { argb: 'FFFF0000' }};
  nochesLadoBCell.alignment = { horizontal: 'center', vertical: 'middle' };

  const headerRow = sheet.getRow(1);
  headerRow.font = { bold: true };
  headerRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF87CEEB' } };
  headerRow.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };

  sheet.mergeCells('B2:G2');
  const ladoACell = sheet.getCell('B2');
  ladoACell.value = 'Lado A';
  ladoACell.alignment = { horizontal: 'center', vertical: 'middle' };
  ladoACell.font = { bold: true };
  ladoACell.fill = { type: 'pattern', pattern:'solid', fgColor: { argb:'FFEEEEEE' } };

  sheet.mergeCells('I2:N2');
  const ladoBCell = sheet.getCell('I2');
  ladoBCell.value = 'Lado B';
  ladoBCell.alignment = { horizontal: 'center', vertical: 'middle' };
  ladoBCell.font = { bold: true };
  ladoBCell.fill = { type: 'pattern', pattern:'solid', fgColor: { argb:'FFEEEEEE' } };

  sheet.getRow(2).height = 18;

  let eraAnterior = null;
  Object.values(grupos).forEach(g => {
    const A = g.A || {};
    const B = g.B || {};

    const nochesA = obtenerNochesDeCelda(A.Variedad);
    const nochesB = obtenerNochesDeCelda(B.Variedad);
    const fechaA = parseFechaSiembra(A.Fecha_Siembra);
    const fechaB = parseFechaSiembra(B.Fecha_Siembra);
    const fechaMasNochesA = fechaA ? new Date(fechaA.getTime() + nochesA * 86400000) : null;
    const fechaMasNochesB = fechaB ? new Date(fechaB.getTime() + nochesB * 86400000) : null;

    const nuevaFila = sheet.addRow({
      Seccion: g.Seccion,
      NaveA: A.Nave || "",
      EraA: A.Era || "",
      VarA: A.Variedad || "",
      SiembraA: fechaA ? formatDateDDMMYYYY(fechaA) : (A.Fecha_Siembra || ""),
      NochesA: fechaMasNochesA ? formatDateDDMMYYYY(fechaMasNochesA) : "",
      CorteA: A.Inicio_Corte || "",
      NaveB: B.Nave || "",
      EraB: B.Era || "",
      VarB: B.Variedad || "",
      SiembraB: fechaB ? formatDateDDMMYYYY(fechaB) : (B.Fecha_Siembra || ""),
      NochesB: fechaMasNochesB ? formatDateDDMMYYYY(fechaMasNochesB) : "",
      CorteB: B.Inicio_Corte || ""
    });

    const eraActual = A.Era || B.Era || "";
    if (eraAnterior <= 3 && eraActual >= 4 || eraAnterior >=4 && eraActual <=3) {
      nuevaFila.eachCell(cell => {
        cell.border = {
          top: { style: 'thick', color: { argb: 'FF000000' } }
        };
      });
    }

    eraAnterior = eraActual;

    const navesPares = A.Nave || B.Nave || "";
    if (navesPares % 2 == 0){
      nuevaFila.eachCell(cell =>{
        cell.fill = {type: 'pattern', pattern:'solid', fgColor: { argb:'FFCCCCCC' }}
      })
    }
  });
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
        crearHojaDistribucionProductos(wbFinal, datosFinales);
        crearHojaDisbud(wbFinal, datosFinales);
        crearHojaGirasol(wbFinal, datosFinales);
        crearHojaPruebaFloracion(wbFinal, datosFinales);
        crearHojaNochesLuz(wbFinal, datosCrudos);

        // Guardar archivo final
        const outputPath = `Reporte_Siembra_${semanaActual}_${Date.now()}.xlsx`;
        await wbFinal.xlsx.writeFile(outputPath);

        console.log("Reporte completo generado:", outputPath);

        res.download(outputPath, `Reporte_Siembra_${semanaActual}.xlsx`, (err) => {
            if (err) console.error("Error enviando el archivo:", err);
            fs.unlinkSync(outputPath);
        });

        fs.unlinkSync(filePath);

    } catch (error) {
        console.error("Error procesando Excel:", error);
        if (req.file && fs.existsSync(req.file.path)) fs.unlinkSync(req.file.path);
        res.status(500).json({ error: "Error procesando el archivo", detalle: error.message });
    }
});

app.listen(5000, () =>
  console.log("🚀 Servidor corriendo en http://localhost:5000")
);
