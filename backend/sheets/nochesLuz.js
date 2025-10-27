const { parseFechaFlexible, formatDateDDMMYYYY } = require('./helpers');

module.exports = function crearHojaNochesLuz(workbook, datos, opts = {}) {

  if (!Array.isArray(datos) || datos.length === 0) return;
  // aceptar variedades inyectadas vía opts (mejor para tests) o requerir el módulo por defecto
  const variedades = opts.variedades || require('../variedades');
  const nochesMap = new Map((variedades || []).map(v => [String(v.nombre).trim().toLowerCase(), Number(v.nochesLuz) || 17]));
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

  // usamos formatDateDDMMYYYY importado desde helpers

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
    if (eraAnterior >=4 && eraActual <=3) {
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
};
