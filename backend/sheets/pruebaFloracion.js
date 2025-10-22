const { clasificarVariedad } = require('./helpers');

module.exports = function crearHojaPruebaFloracion(workbook, datos) {
  const floracionData = datos.filter(row => clasificarVariedad(row.Variedad) === 'Prueba de Floracion');
  if (floracionData.length === 0) return;
  const grupos = {};
  floracionData.forEach(row => {
    const seccion = row.Seccion || 'Sin Sección';
    const nave = row.Nave?.toString().trim() || 'Sin Nave';
    const era = row.Era || 'Sin Era';
    const lado = row.Lado || '';
    const key = `${seccion}__${nave}__${era}__${lado}`;
    if (!grupos[key]) { grupos[key] = { seccion, nave, era, lado, metros: 0, fechaSiembra: row.Fecha_Siembra || '', inicioCorte: row.Inicio_Corte || '', variedades: new Set() }; }
    grupos[key].metros += parseFloat(row.Largo) || 0;
    if (row.Variedad) grupos[key].variedades.add(row.Variedad);
    if (!grupos[key].fechaSiembra && row.Fecha_Siembra) grupos[key].fechaSiembra = row.Fecha_Siembra;
    if (!grupos[key].inicioCorte && row.Inicio_Corte) grupos[key].inicioCorte = row.Inicio_Corte;
  });
  let sheet = workbook.getWorksheet('Prueba de Floración');
  if (sheet) workbook.removeWorksheet(sheet.id);
  sheet = workbook.addWorksheet('Prueba de Floración');
  sheet.columns = [ { header: 'Sección', key: 'seccion', width: 15 }, { header: 'Nave', key: 'nave', width: 12 }, { header: 'Lado', key: 'lado', width: 8 }, { header: 'Era', key: 'era', width: 14 }, { header: 'Variedades', key: 'variedades', width: 40 }, { header: 'Metros', key: 'metros', width: 12 }, { header: 'Eras', key: 'eras', width: 12 }, { header: 'Fecha Siembra', key: 'fechaSiembra', width: 18 }, { header: 'Inicio Corte', key: 'inicioCorte', width: 18 } ];
  Object.values(grupos).forEach(info => {
    const variedades = Array.from(info.variedades).sort().join(', ');
    const metros = info.metros;
    const eras = (metros / 30).toFixed(2);
    sheet.addRow({ seccion: info.seccion, nave: info.nave, lado: info.lado, era: info.era, variedades, metros: metros.toFixed(2), eras, fechaSiembra: info.fechaSiembra, inicioCorte: info.inicioCorte });
  });
  const totalMetros = Object.values(grupos).reduce((acc,g)=>acc+g.metros,0);
  const totalEras = (totalMetros/30).toFixed(2);
  const headerRow = sheet.getRow(1); headerRow.font = { bold: true }; headerRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFCCCCCC' } }; headerRow.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
  const totalRow = sheet.addRow({ seccion: 'TOTAL GENERAL', metros: totalMetros.toFixed(2), eras: totalEras });
  totalRow.font = { bold: true, size: 12 }; totalRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFE699' } }; totalRow.alignment = { horizontal: 'center' };
  sheet.getRow(1).font = { bold: true };
};
