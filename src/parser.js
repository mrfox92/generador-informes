const XLSX = require('xlsx');

const EXCLUDED_FROM_SEMESTER_AVG = ['RELIGION', 'RELIGIÓN', 'ORIENTACION', 'ORIENTACIÓN'];

function norm(s) {
  return String(s || '')
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .trim().toUpperCase();
}

function toNumber(value) {
  if (value === undefined || value === null || value === '') return null;
  if (typeof value === 'number' && Number.isFinite(value)) return value;
  const clean = String(value).replace(',', '.').trim();
  const n = Number(clean);
  return Number.isFinite(n) ? n : null;
}

function reorderStudentName(original) {
  const raw = String(original || '').replace(/\s+/g, ' ').trim();
  if (!raw) return '';
  const parts = raw.split(' ');
  if (parts.length === 1) return raw;
  // Las planillas vienen generalmente como: APELLIDO1 APELLIDO2 NOMBRE1 NOMBRE2...
  // Se invierte a: NOMBRE1 NOMBRE2... APELLIDO1 APELLIDO2.
  if (parts.length === 2) return `${parts[1]} ${parts[0]}`;
  if (parts.length === 3) return `${parts[2]} ${parts[0]} ${parts[1]}`;
  return `${parts.slice(2).join(' ')} ${parts[0]} ${parts[1]}`;
}


function keyWords(s) {
  return norm(s)
    .replace(/[^A-Z0-9Ñ]+/g, ' ')
    .split(/\s+/)
    .filter(w => w.length >= 3 && !['LAS','LOS','PARA','CON','DEL'].includes(w));
}

function sheetMatchScore(subjectName, sheetName) {
  const a = norm(subjectName);
  const b = norm(sheetName);
  if (!a || !b) return 0;
  if (a === b) return 100;
  if (a.includes(b) || b.includes(a)) return 90;
  const aw = keyWords(a);
  const bw = keyWords(b);
  if (!aw.length || !bw.length) return 0;
  let hits = 0;
  for (const w of aw) {
    if (bw.some(x => x === w || x.startsWith(w.slice(0, 6)) || w.startsWith(x.slice(0, 6)))) hits++;
  }
  return Math.round((hits / Math.max(aw.length, bw.length)) * 80);
}

function findSubjectMap(subjectMaps, subjectName) {
  const exact = subjectMaps.get(norm(subjectName));
  if (exact) return exact;
  let best = null;
  let bestScore = 0;
  for (const [sheetKey, data] of subjectMaps.entries()) {
    const score = sheetMatchScore(subjectName, sheetKey);
    if (score > bestScore) {
      bestScore = score;
      best = data;
    }
  }
  // 45 permite emparejar hojas truncadas o con pequeñas diferencias:
  // "Programación y base de datos" vs "Programación y bases de datos",
  // "Instalación y configuracion..." vs hoja truncada de Excel, etc.
  return bestScore >= 35 ? best : null;
}

function display(value) {
  if (value === undefined || value === null) return '';
  if (typeof value === 'number') return String(Math.trunc(value));
  return String(value).trim();
}

function semesterAverage(subjects) {
  const values = subjects
    .filter(s => !EXCLUDED_FROM_SEMESTER_AVG.includes(norm(s.name)))
    .map(s => toNumber(s.average))
    .filter(n => n !== null);
  if (!values.length) return '';
  return Math.trunc(values.reduce((a, b) => a + b, 0) / values.length);
}

function parseSubjectSheet(sheet, subjectName) {
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
  const data = new Map();
  for (let r = 1; r < rows.length; r++) {
    const row = rows[r];
    const rut = String(row[1] || '').trim();
    const nombre = String(row[2] || '').trim();
    if (!rut || !nombre) continue;
    const grades = [];
    for (let c = 3; c <= 8; c++) grades.push(display(row[c]));
    const avg = display(row[9] || row[8] || row[7] || row[6]);
    data.set(norm(rut), { subjectName, grades, average: avg });
  }
  return data;
}

function parseWorkbook(filePath, meta = {}) {
  const wb = XLSX.readFile(filePath, { cellDates: false });
  if (!wb.SheetNames.length) throw new Error('La planilla no contiene hojas.');
  const consolidadoName = wb.SheetNames.find(n => norm(n).includes('CONSOLIDADO')) || wb.SheetNames[0];
  const ws = wb.Sheets[consolidadoName];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
  if (rows.length < 3) throw new Error('No se pudo leer el consolidado general.');

  const header = rows[1];
  const subjectNames = [];
  for (let c = 3; c < header.length; c++) {
    const name = String(header[c] || '').trim();
    if (name && !norm(name).includes('PROMEDIO')) subjectNames.push({ col: c, name });
  }

  const subjectMaps = new Map();
  for (const sheetName of wb.SheetNames) {
    if (sheetName === consolidadoName) continue;
    subjectMaps.set(norm(sheetName), parseSubjectSheet(wb.Sheets[sheetName], sheetName));
  }

  const students = [];
  for (let r = 2; r < rows.length; r++) {
    const row = rows[r];
    const rut = String(row[1] || '').trim();
    const nombreOriginal = String(row[2] || '').trim();
    const nombre = reorderStudentName(nombreOriginal);
    if (!rut || !nombreOriginal) continue;

    const subjects = subjectNames.map(({ col, name }) => {
      const fromSheet = findSubjectMap(subjectMaps, name)?.get(norm(rut));
      const grades = fromSheet?.grades || ['', '', '', '', '', ''];
      const average = fromSheet?.average || display(row[col]);
      return { name, grades, average };
    });

    students.push({ rut, nombre, nombreOriginal, subjects, promedioSemestral: semesterAverage(subjects) });
  }

  return {
    course: meta.curso || '',
    consolidadoName,
    subjects: subjectNames.map(s => s.name),
    students
  };
}

module.exports = { parseWorkbook, semesterAverage, EXCLUDED_FROM_SEMESTER_AVG, reorderStudentName };
