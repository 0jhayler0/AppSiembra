const { clasificarVariedad } = require('./helpers');

module.exports = function crearHojaDisbud(workbook, datos) {
  // Copied logic from server.js crearHojaDisbud
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
      variedades.push({ nombre: row.Variedad, esDisbud: esDisbud, largo: !isNaN(largo) ? largo.toFixed(2) : '0' });
    });

    if (tieneDisbud) {
      gruposFinales[key] = { ...base, Variedades: variedades, TotalLargo: totalLargoDisbud };
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
    const fila = sheet.addRow({ seccion: g.Seccion, nave: g.Nave, lado: g.Lado, era: g.Era, variedades: "", total: g.TotalLargo.toFixed(2), siembra: g.Fecha_Siembra, corte: g.Inicio_Corte });
    const mostrarLargo = g.Variedades.length > 1;
    const celda = fila.getCell("variedades");
    celda.value = {
      richText: g.Variedades.flatMap((v, i) => {
        const textoVariedad = mostrarLargo ? `${v.nombre} (${v.largo})` : v.nombre;
        return [ ...(i > 0 ? [{ text: " " }] : []), { text: textoVariedad, font: v.esDisbud ? { color: { argb: "FFFF0000" } } : {} } ];
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
  totalRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD9E1F2' } };

  const totalErasRow = sheet.addRow({});
  sheet.mergeCells(totalErasRow.number, 1, totalErasRow.number, 5);
  const totalErasCell = totalErasRow.getCell(1);
  totalErasCell.value = "TOTAL ERAS";
  totalErasCell.alignment = { horizontal: 'right' };
  totalErasRow.getCell("total").value = totalEnEras.toFixed(2);
  totalErasRow.font = { bold: true, size: 12 };
  totalErasRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD9E9F2' } };
};
