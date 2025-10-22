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

function getISOWeek(date) {
  const d = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
  const dayNum = d.getUTCDay() || 7;
  d.setUTCDate(d.getUTCDate() + 4 - dayNum);
  const yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
  return Math.ceil((((d - yearStart) / 86400000) + 1) / 7);
}

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

function parseFechaFlexible(value) {
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

function formatDateDDMMYYYY(d) {
  if (!d || !(d instanceof Date) || isNaN(d)) return "";
  const dd = String(d.getDate()).padStart(2, "0");
  const mm = String(d.getMonth() + 1).padStart(2, "0");
  const yyyy = d.getFullYear();
  return `${dd}-${mm}-${yyyy}`;
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

module.exports = { getISOWeek, parseInicioCorteToDate, parseFechaFlexible, formatDateDDMMYYYY, clasificarVariedad };
