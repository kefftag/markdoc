const express = require('express');
const multer = require('multer');
const { marked } = require('marked');
const HTMLtoDOCX = require('html-to-docx');
const path = require('path');

const app = express();
const PORT = process.env.PORT || 3000;

// In-memory storage — no files written to disk
const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 10 * 1024 * 1024 }, // 10 MB
  fileFilter: (_req, file, cb) => {
    const allowed = ['.md', '.markdown', '.txt'];
    const ext = path.extname(file.originalname).toLowerCase();
    if (allowed.includes(ext) || file.mimetype === 'text/markdown' || file.mimetype === 'text/plain') {
      cb(null, true);
    } else {
      cb(new Error('Only .md, .markdown, and .txt files are accepted'));
    }
  },
});

app.use(express.static(path.join(__dirname, 'public')));

app.post('/convert', upload.single('file'), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ error: 'No file uploaded' });
    }

    const markdown = req.file.buffer.toString('utf-8');

    // Parse markdown → HTML
    const html = `<!DOCTYPE html><html><body>${marked(markdown)}</body></html>`;

    // Convert HTML → docx buffer
    const docxBuffer = await HTMLtoDOCX(html, null, {
      table: { row: { cantSplit: true } },
      footer: false,
      pageNumber: false,
    });

    // Derive output filename from uploaded filename
    const baseName = path.basename(req.file.originalname, path.extname(req.file.originalname));
    const outputName = `${baseName}.docx`;

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', `attachment; filename="${outputName}"`);
    res.send(docxBuffer);
  } catch (err) {
    console.error('Conversion error:', err);
    res.status(500).json({ error: 'Conversion failed', detail: err.message });
  }
});

// Simple error handler for multer and other middleware errors
app.use((err, _req, res, _next) => {
  res.status(400).json({ error: err.message });
});

app.listen(PORT, () => {
  console.log(`markdoc running at http://localhost:${PORT}`);
});
