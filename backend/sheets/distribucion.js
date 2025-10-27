module.exports = function crearHojaDistribucionProductos(workbook, datos) {
  const metrosPorSeccion = {};
  const navesPorSeccion = {};

  datos.forEach(row => {
    const seccion = row.Seccion || "Sin Sección";
    const tipo = (row.Variedad && typeof row.Variedad === 'string') ? (row.Variedad.toLowerCase().includes('vincent choice') ? 'Girasol' : (row.Variedad.toLowerCase().includes('prueba de floracion') ? 'Prueba de Floracion' : (row.Variedad.toLowerCase().includes('cremon')||row.Variedad.toLowerCase().includes('spider')||row.Variedad.toLowerCase().includes('towi')||row.Variedad.toLowerCase().includes('tutu') ? 'Disbud' : 'Normal'))) : 'Normal';
    const largo = parseFloat(row.Largo) || 0;
    const nave = row.Nave?.toString().trim() || "Sin Nave";

    if (!metrosPorSeccion[seccion]) {
      metrosPorSeccion[seccion] = { Disbud: 0, Girasol: 0, Normal: 0, "Prueba de Floracion": 0 };
    }
    if (!navesPorSeccion[seccion]) navesPorSeccion[seccion] = new Set();
    if (metrosPorSeccion[seccion][tipo] !== undefined) metrosPorSeccion[seccion][tipo] += largo;
    if (nave !== "Sin Nave") navesPorSeccion[seccion].add(nave);
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

  const borderSmall = { top: { style: 'thin', color: { argb: 'FF1E90FF' } }, left: { style: 'thin', color: { argb: 'FF1E90FF' } }, bottom: { style: 'thin', color: { argb: 'FF1E90FF' } }, right: { style: 'thin', color: { argb: 'FF1E90FF' } } };
  const borderLarge = { top: { style: 'thin', color: { argb: 'FFFF8C00' } }, left: { style: 'thin', color: { argb: 'FFFF8C00' } }, bottom: { style: 'thin', color: { argb: 'FFFF8C00' } }, right: { style: 'thin', color: { argb: 'FFFF8C00' } } };

  Object.entries(metrosPorSeccionEnEras).forEach(([seccion, valores]) => {
    const naves = Array.from(navesPorSeccion[seccion] || []).sort().join(", ");
    const row = sheet.addRow({ seccion, naves, eras: (parseFloat(valores.Disbud||0)+parseFloat(valores.Girasol||0)+parseFloat(valores.Normal||0)+parseFloat(valores["Prueba de Floracion"]||0)).toFixed(2), pompon: valores.Normal, disbud: valores.Disbud, girasol: valores.Girasol, floracion: valores["Prueba de Floracion"], total: (parseFloat(valores.Disbud||0)+parseFloat(valores.Girasol||0)+parseFloat(valores.Normal||0)+parseFloat(valores["Prueba de Floracion"]||0)).toFixed(2) });
    const navesCount = naves ? naves.split(",").map(s => s.trim()).filter(Boolean).length : 0;
    const chosenBorder = navesCount <= 3 ? borderSmall : borderLarge;
    row.eachCell(cell => { cell.border = chosenBorder; });
  });

  sheet.getRow(1).font = { bold: true };

  const totalDisbud = sheet.getColumn("disbud").values.slice(2).reduce((a,b)=>a+(parseFloat(b)||0),0);
  const totalGirasol = sheet.getColumn("girasol").values.slice(2).reduce((a,b)=>a+(parseFloat(b)||0),0);
  const totalPompon = sheet.getColumn("pompon").values.slice(2).reduce((a,b)=>a+(parseFloat(b)||0),0);
  const totalFloracion = sheet.getColumn("floracion").values.slice(2).reduce((a,b)=>a+(parseFloat(b)||0),0);
  const totalGlobal = totalDisbud + totalGirasol + totalPompon + totalFloracion;

  const totalRow = sheet.addRow({ seccion: "TOTAL GENERAL", naves: "", eras: totalGlobal.toFixed(2), pompon: totalPompon.toFixed(2), disbud: totalDisbud.toFixed(2), girasol: totalGirasol.toFixed(2), floracion: totalFloracion.toFixed(2), total: totalGlobal.toFixed(2) });

  const headerRow = sheet.getRow(1);
  headerRow.font = { bold: true };
  headerRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFCCCCCC' } };
  headerRow.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
  totalRow.font = { bold: true, size: 12 };
  totalRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFCCE5FF' } };
  totalRow.alignment = { horizontal: 'center' };
};
