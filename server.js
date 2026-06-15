const express = require('express');
const multer = require('multer');
const path = require('path');
const fs = require('fs');
const archiver = require('archiver');
const sanitize = require('sanitize-filename');
const { parseWorkbook } = require('./src/parser');
const { generateDocxReport, generatePdfReport } = require('./src/reportGenerator');

const app = express();
const PORT = process.env.PORT || 3000;
const upload = multer({ dest: path.join(__dirname, 'uploads') });
const sessions = new Map();

app.use(express.json({ limit: '10mb' }));
app.use(express.static(path.join(__dirname, 'public')));
app.use('/assets', express.static(path.join(__dirname, 'assets')));

function ensureDir(dir) { fs.mkdirSync(dir, { recursive: true }); }
function slug(text) {
  return sanitize(String(text || 'SIN_CURSO')
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .replace(/°/g, '')
    .replace(/[^a-zA-Z0-9]+/g, '_')
    .replace(/^_+|_+$/g, '')
    .toUpperCase());
}

app.post('/api/upload', upload.single('excel'), (req, res) => {
  try {
    const meta = JSON.parse(req.body.meta || '{}');
    if (!req.file) return res.status(400).json({ error: 'Debe subir una planilla Excel.' });
    const parsed = parseWorkbook(req.file.path, meta);
    const sessionId = Date.now().toString(36) + Math.random().toString(36).slice(2);
    sessions.set(sessionId, { filePath: req.file.path, meta, parsed });
    res.json({ sessionId, ...parsed });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.post('/api/report/:sessionId/:studentIndex', async (req, res) => {
  try {
    const session = sessions.get(req.params.sessionId);
    if (!session) return res.status(404).json({ error: 'Sesión no encontrada. Vuelva a cargar la planilla.' });
    const student = session.parsed.students[Number(req.params.studentIndex)];
    if (!student) return res.status(404).json({ error: 'Estudiante no encontrado.' });

    const format = req.body.format || 'docx';
    const outDir = path.join(__dirname, 'reports', req.params.sessionId);
    ensureDir(outDir);
    const base = slug(`${student.nombre}_${student.rut}`);
    const docxPath = path.join(outDir, `${base}.docx`);
    await generateDocxReport({ student, meta: session.meta, outPath: docxPath });

    if (format === 'pdf') {
      const pdfPath = path.join(outDir, `${base}.pdf`);
      await generatePdfReport({ student, meta: session.meta, outPath: pdfPath });
      res.download(pdfPath, path.basename(pdfPath));
    } else {
      res.download(docxPath, path.basename(docxPath));
    }
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.post('/api/report-all/:sessionId', async (req, res) => {
  try {
    const session = sessions.get(req.params.sessionId);
    if (!session) return res.status(404).json({ error: 'Sesión no encontrada. Vuelva a cargar la planilla.' });
    const selected = req.body.selected || session.parsed.students.map((_, i) => i);
    const format = req.body.format || 'docx';
    const cursoSlug = slug(session.meta.curso || session.parsed.course || 'CURSO');
    const outRoot = path.join(__dirname, 'reports', req.params.sessionId, cursoSlug);
    ensureDir(outRoot);

    for (const idx of selected) {
      const student = session.parsed.students[Number(idx)];
      if (!student) continue;
      const base = slug(`${student.nombre}_${student.rut}`);
      const docxPath = path.join(outRoot, `${base}.docx`);
      await generateDocxReport({ student, meta: session.meta, outPath: docxPath });
      if (format === 'pdf') {
        const pdfPath = path.join(outRoot, `${base}.pdf`);
        await generatePdfReport({ student, meta: session.meta, outPath: pdfPath });
        if (fs.existsSync(docxPath)) fs.unlinkSync(docxPath);
      }
    }

    const zipPath = path.join(__dirname, 'reports', `${cursoSlug}.zip`);
    await new Promise((resolve, reject) => {
      const output = fs.createWriteStream(zipPath);
      const archive = archiver('zip', { zlib: { level: 9 } });
      output.on('close', resolve);
      archive.on('error', reject);
      archive.pipe(output);
      archive.directory(outRoot, cursoSlug);
      archive.finalize();
    });
    res.download(zipPath, `${cursoSlug}.zip`);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.listen(PORT, () => console.log(`App informes de notas: http://localhost:${PORT}`));
