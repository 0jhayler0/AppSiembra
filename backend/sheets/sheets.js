// sheets.js

function generarHTML(rows, semana) {
  const htmlRows = rows.map(r => `
    <tr>
      <td>${r.Seccion}</td>
      <td>${r.Lado}</td>
      <td>${r.Nave}</td>
      <td>${r.Era}</td>
      <td>${r.Variedad}</td>
      <td>${r.Largo}</td>
      <td>${r.Fecha_Siembra}</td>
      <td>${r.Inicio_Corte}</td>
    </tr>
  `).join("");

  return `
  <html>
  <head>
    <style>
      table { width: 100%; border-collapse: collapse; font-size: 12px; }
      th, td { border: 1px solid #000; padding: 4px; text-align: left; }
      th { background: #eee; }
      h2 { margin-bottom: 12px; }
    </style>
  </head>
  <body>
    <h2>Reporte Siembra - Semana ${semana}</h2>
    <table>
      <thead>
        <tr>
          <th>Sección</th>
          <th>Lado</th>
          <th>Nave</th>
          <th>Era</th>
          <th>Variedad</th>
          <th>Largo</th>
          <th>Fecha Siembra</th>
          <th>Inicio Corte</th>
        </tr>
      </thead>
      <tbody>
        ${htmlRows}
      </tbody>
    </table>
  </body>
  </html>
  `;
}

module.exports = { generarHTML };
