const fs = require('fs');
const path = require('path');
const PDFDocument = require('pdfkit');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  WidthType, AlignmentType, BorderStyle, VerticalAlign, ImageRun, PageOrientation, TableLayoutType
} = require('docx');

function thinBorders() {
  return ['top', 'bottom', 'left', 'right'].reduce((a, k) => {
    a[k] = { style: BorderStyle.SINGLE, size: 4, color: '000000' };
    return a;
  }, {});
}

function noBorders() {
  return ['top', 'bottom', 'left', 'right'].reduce((a, k) => {
    a[k] = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };
    return a;
  }, {});
}

function run(text = '', opts = {}) {
  return new TextRun({
    text: String(text ?? ''),
    bold: !!opts.bold,
    size: opts.size || 14,
    font: 'Arial'
  });
}

function paragraph(text = '', opts = {}) {
  return new Paragraph({
    alignment: opts.align || AlignmentType.LEFT,
    spacing: { before: opts.before || 0, after: opts.after || 0 },
    children: [run(text, opts)]
  });
}

function cell(text = '', opts = {}) {
  const lines = String(text ?? '').split('\n');
  return new TableCell({
    width: { size: opts.width || 1000, type: WidthType.DXA },
    verticalAlign: VerticalAlign.CENTER,
    margins: { top: 35, bottom: 35, left: 45, right: 45 },
    borders: opts.noBorder ? noBorders() : thinBorders(),
    children: lines.map((line) => new Paragraph({
      alignment: opts.align || AlignmentType.CENTER,
      spacing: { before: 0, after: 0 },
      children: [run(line, { bold: opts.bold, size: opts.size || 14 })]
    }))
  });
}

function logoCell(file, width, height, align) {
  const children = [];
  if (file && fs.existsSync(file)) {
    children.push(new Paragraph({
      alignment: align || AlignmentType.LEFT,
      children: [new ImageRun({ type: 'png', data: fs.readFileSync(file), transformation: { width, height } })]
    }));
  } else {
    children.push(paragraph(''));
  }
  return new TableCell({ borders: noBorders(), children });
}

function normalizePath(p) { return p ? path.resolve(p) : ''; }
function cleanGrade(v) { return v === undefined || v === null ? '' : String(v); }

function buildRows(subjects, promedioSemestral) {
  // Anchos tomados desde el formato institucional original TEMPLATE_INFORME.docx.
  // Total aproximado: 10783 dxa. No usar 100% porque Word expande la tabla y rompe el aspecto institucional.
  const W_SUBJECT = 5102;
  const W_GRADE = 749;
  const W_AVG = 1188;
  const rows = [];
  rows.push(new TableRow({ children: [
    cell('ASIGNATURAS', { width: W_SUBJECT, bold: true, align: AlignmentType.LEFT }),
    ...['N1', 'N2', 'N3', 'N4', 'N5', 'N6'].map(h => cell(h, { bold: true, width: W_GRADE })),
    cell('Promedio', { bold: true, width: W_AVG })
  ] }));

  for (const s of subjects) {
    rows.push(new TableRow({ children: [
      cell(String(s.name || '').toUpperCase(), { width: W_SUBJECT, align: AlignmentType.LEFT }),
      ...[0, 1, 2, 3, 4, 5].map(i => cell(cleanGrade(s.grades?.[i]), { width: W_GRADE })),
      cell(cleanGrade(s.average), { width: W_AVG })
    ] }));
  }

  rows.push(new TableRow({ children: [
    cell('Promedio Semestral', { width: W_SUBJECT, align: AlignmentType.LEFT }),
    ...Array(6).fill(0).map(() => cell('', { width: W_GRADE })),
    cell(cleanGrade(promedioSemestral), { width: W_AVG, bold: true })
  ] }));
  return rows;
}

async function generateDocxReport({ student, meta, outPath }) {
  // LOGOS WORD: deje sus imágenes finales con estos nombres exactos.
  // Se insertan con tamaño fijo para evitar deformación del documento.
  // Izquierda: assets/logo-bicentenario.png | Derecha: assets/logo-liceo.png
  const logoBicentenario = normalizePath(meta.logoBicentenario || path.join(__dirname, '..', 'assets', 'logo-bicentenario.png'));
  const logoLiceo = normalizePath(meta.logoLiceo || path.join(__dirname, '..', 'assets', 'logo-liceo.png'));
  // Tamaños Word en pixeles visuales: bicentenario 110x25, liceo 45x56.
  // Las imágenes se leen desde assets/ y se insertan con tamaño fijo para evitar deformaciones.
  const subjects = student.subjects || [];

  const headerTable = new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    borders: noBorders(),
    rows: [new TableRow({ children: [
      logoCell(logoBicentenario, 110, 25, AlignmentType.LEFT),
      cell('', { noBorder: true, width: 5000 }),
      logoCell(logoLiceo, 45, 56, AlignmentType.RIGHT)
    ] })]
  });

  const doc = new Document({
    styles: {
      default: { document: { run: { font: 'Arial', size: 14 } } }
    },
    sections: [{
      properties: {
        page: {
          size: { orientation: PageOrientation.PORTRAIT },
          margin: { top: 560, right: 560, bottom: 560, left: 560 }
        }
      },
      children: [
        headerTable,
        paragraph('LICEO BICENTENARIO GABRIELA MISTRAL', { align: AlignmentType.CENTER, bold: true, size: 16, before: 150 }),
        paragraph('INFORME  DE NOTAS', { align: AlignmentType.CENTER, bold: true, size: 16 }),
        paragraph(`Año Escolar: ${meta.anio || '2026'}`, { align: AlignmentType.CENTER, bold: true, size: 16, after: 220 }),
        paragraph(`Nombre estudiante: ${student.nombre || ''}`, { size: 14 }),
        paragraph(`Rut: ${student.rut || ''}`, { size: 14 }),
        paragraph(`Curso: ${meta.curso || ''}`, { size: 14 }),
        paragraph(`Periodo: ${meta.periodo || 'Primer Semestre'}`, { size: 14 }),
        paragraph(`Profesor jefe: ${meta.profesorJefe || ''}`, { size: 14, after: 160 }),
        new Table({
          width: { size: 10783, type: WidthType.DXA },
          columnWidths: [5102, 748, 749, 749, 749, 749, 749, 1188],
          layout: TableLayoutType.FIXED,
          alignment: AlignmentType.CENTER,
          rows: buildRows(subjects, student.promedioSemestral)
        }),
        paragraph('', { after: 720 }),
        new Table({ width: { size: 100, type: WidthType.PERCENTAGE }, borders: noBorders(), rows: [
          new TableRow({ children: [
            cell('___________________________\nNombre y firma profesor(a) jefe', { noBorder: true, width: 4200, size: 14 }),
            cell('', { noBorder: true, width: 1600 }),
            cell(`___________________________\n${meta.director || 'Arturo Alvear Avendaño'}\nDirector(a)`, { noBorder: true, width: 4200, size: 14 })
          ] })
        ] })
      ]
    }]
  });

  const buffer = await Packer.toBuffer(doc);
  fs.writeFileSync(outPath, buffer);
  return outPath;
}

function drawPdfTable(doc, x, y, tableWidth, rowHeight, subjects, promedioSemestral) {
  const colSubject = 250;
  const colGrade = 35;
  const colAvg = tableWidth - colSubject - colGrade * 6;
  const cols = [colSubject, colGrade, colGrade, colGrade, colGrade, colGrade, colGrade, colAvg];
  const headers = ['ASIGNATURAS', 'N1', 'N2', 'N3', 'N4', 'N5', 'N6', 'Promedio'];
  const rows = [headers, ...subjects.map(s => [String(s.name || '').toUpperCase(), ...[0,1,2,3,4,5].map(i => cleanGrade(s.grades?.[i])), cleanGrade(s.average)]), ['Promedio Semestral', '', '', '', '', '', '', cleanGrade(promedioSemestral)]];

  doc.lineWidth(0.5).font('Helvetica').fontSize(7);
  rows.forEach((row, r) => {
    let cx = x;
    const h = r === 0 ? rowHeight + 5 : rowHeight;
    row.forEach((txt, c) => {
      doc.rect(cx, y, cols[c], h).stroke();
      doc.font(r === 0 ? 'Helvetica-Bold' : 'Helvetica').fontSize(7);
      const tx = cx + 3;
      const ty = y + 4;
      const align = c === 0 ? 'left' : 'center';
      doc.text(String(txt || ''), tx, ty, { width: cols[c] - 6, height: h - 4, align, ellipsis: true });
      cx += cols[c];
    });
    y += h;
  });
}

async function generatePdfReport({ student, meta, outPath }) {
  return new Promise((resolve, reject) => {
    const doc = new PDFDocument({ size: 'LETTER', margin: 36 });
    const stream = fs.createWriteStream(outPath);
    doc.pipe(stream);

    // LOGOS PDF: mismas rutas que Word.
    // Se usan anchos fijos; PDFKit respeta proporción si solo se indica width.
    const logoBicentenario = normalizePath(meta.logoBicentenario || path.join(__dirname, '..', 'assets', 'logo-bicentenario.png'));
    const logoLiceo = normalizePath(meta.logoLiceo || path.join(__dirname, '..', 'assets', 'logo-liceo.png'));
    try { if (fs.existsSync(logoBicentenario)) doc.image(logoBicentenario, 55, 30, { width: 110 }); } catch (_) {}
    try { if (fs.existsSync(logoLiceo)) doc.image(logoLiceo, 505, 25, { width: 45 }); } catch (_) {}

    doc.font('Helvetica-Bold').fontSize(9).text('LICEO BICENTENARIO GABRIELA MISTRAL', 36, 72, { align: 'center' });
    doc.text('INFORME  DE NOTAS', { align: 'center' });
    doc.text(`Año Escolar: ${meta.anio || '2026'}`, { align: 'center' });

    doc.font('Helvetica').fontSize(7.5);
    let y = 125;
    doc.text(`Nombre estudiante: ${student.nombre || ''}`, 55, y); y += 12;
    doc.text(`Rut: ${student.rut || ''}`, 55, y); y += 12;
    doc.text(`Curso: ${meta.curso || ''}`, 55, y); y += 12;
    doc.text(`Periodo: ${meta.periodo || 'Primer Semestre'}`, 55, y); y += 12;
    doc.text(`Profesor jefe: ${meta.profesorJefe || ''}`, 55, y); y += 20;

    drawPdfTable(doc, 55, y, 500, 13, student.subjects || [], student.promedioSemestral);

    doc.font('Helvetica').fontSize(7.5);
    doc.text('___________________________', 70, 675, { width: 170, align: 'center' });
    doc.text('Nombre y firma profesor(a) jefe', 70, 688, { width: 170, align: 'center' });
    doc.text('___________________________', 380, 675, { width: 170, align: 'center' });
    doc.text(meta.director || 'Arturo Alvear Avendaño', 380, 688, { width: 170, align: 'center' });
    doc.font('Helvetica-Bold').text('Director(a)', 380, 700, { width: 170, align: 'center' });

    doc.end();
    stream.on('finish', () => resolve(outPath));
    stream.on('error', reject);
  });
}

module.exports = { generateDocxReport, generatePdfReport };
