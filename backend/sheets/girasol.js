const { parseInicioCorteToDate, getISOWeek, clasificarVariedad } = require('./helpers');

module.exports = function crearHojaGirasol(workbook, datos) {
  const girasolPorSeccion = {};
  const navesPorSeccion = {};
  const fechasPorSeccion = {};

  datos.forEach(row => {
    const tipo = clasificarVariedad(row.Variedad);
    if (tipo !== 'Girasol') return;
    const seccion = row.Seccion || 'Sin Sección';
    const nave = row.Nave?.toString().trim() || 'Sin Nave';
    const largo = parseFloat(row.Largo) || 0;
    const fechaSiembra = row.Fecha_Siembra || '';
    const inicioCorte = row.Inicio_Corte || '';
    if (!girasolPorSeccion[seccion]) { girasolPorSeccion[seccion] = { metros: 0 }; navesPorSeccion[seccion] = new Set(); fechasPorSeccion[seccion] = { siembra: new Set(), corte: new Set() }; }
    girasolPorSeccion[seccion].metros += largo;
    if (nave !== 'Sin Nave') navesPorSeccion[seccion].add(nave);
    if (fechaSiembra) fechasPorSeccion[seccion].siembra.add(fechaSiembra);
    if (inicioCorte) fechasPorSeccion[seccion].corte.add(inicioCorte);
  });

  let sheet = workbook.getWorksheet('Girasol');
  if (sheet) workbook.removeWorksheet(sheet.id);
  sheet = workbook.addWorksheet('Girasol');

  sheet.columns = [
    { header: 'Sección', key: 'seccion', width: 15 },
    { header: 'Nave', key: 'naves', width: 25 },
    { header: 'Eras', key: 'eras', width: 15 },
    { header: 'Fecha Siembra', key: 'fechaSiembra', width: 25 },
    { header: 'Inicio Corte', key: 'inicioCorte', width: 25 },
    { header: 'Semana Corte', key: 'semanaCorte', width: 25 },
    { header: 'Estimado Producción', key: 'estimado', width: 20 }
  ];

  Object.entries(girasolPorSeccion).forEach(([seccion, info]) => {
    const metros = info.metros || 0;
    const eras = (metros / 30).toFixed(2);
    const estimado = Math.round(eras * 850);
    const naves = Array.from(navesPorSeccion[seccion] || []).sort().join(', ');
    const fechasSiembraArr = Array.from(fechasPorSeccion[seccion].siembra).filter(Boolean);
    const fechasCorteArr = Array.from(fechasPorSeccion[seccion].corte).filter(Boolean);
    let defaultYear = (new Date()).getFullYear();
    if (fechasSiembraArr.length > 0) {
      const tryDate = new Date(fechasSiembraArr[0]);
      if (!isNaN(tryDate)) defaultYear = tryDate.getFullYear();
      else { const m = String(fechasSiembraArr[0]).match(/(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})/); if (m) { let y = parseInt(m[3],10); if (y<100) y+=2000; defaultYear = y; } }
    }
    const semanas = fechasCorteArr.map(fc => parseInicioCorteToDate(fc, defaultYear)).filter(d => d instanceof Date && !isNaN(d)).map(d => getISOWeek(d));
    const semanasUnicas = Array.from(new Set(semanas)).sort((a,b)=>a-b);
    const semanaCorte = semanasUnicas.length > 0 ? String(Math.max(...semanasUnicas)).padStart(2, '0') : '';
    sheet.addRow({ seccion, naves, eras, fechaSiembra: fechasSiembraArr.join(', '), inicioCorte: fechasCorteArr.join(', '), semanaCorte, estimado });
  });

  const totalMetros = Object.values(girasolPorSeccion).reduce((acc,v)=>acc+(v.metros||0),0);
  const totalEras = (totalMetros/30).toFixed(2);
  const totalEstimado = Math.round(totalEras * 850);
  const totalRow = sheet.addRow({ seccion: 'TOTAL GENERAL', naves: '', eras: totalEras, semanaCorte: '', estimado: totalEstimado });
  totalRow.font = { bold: true, size: 12 };
  totalRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFE699' } };
  totalRow.alignment = { horizontal: 'center' };
  sheet.getRow(1).font = { bold: true };
};
