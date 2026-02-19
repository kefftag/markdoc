const express = require('express');
const multer = require('multer');
const { marked } = require('marked');
const {
  Document, Paragraph, TextRun, HeadingLevel, Packer,
  UnderlineType, Table, TableRow, TableCell, WidthType, BorderStyle,
} = require('docx');
const path = require('path');

const app = express();
const PORT = process.env.PORT || 3000;

const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 10 * 1024 * 1024 },
  fileFilter: (_req, file, cb) => {
    const ext = path.extname(file.originalname).toLowerCase();
    if (['.md', '.markdown', '.txt'].includes(ext)) cb(null, true);
    else cb(new Error('Only .md, .markdown, and .txt files are accepted'));
  },
});

app.use(express.static(path.join(__dirname, 'public')));

// ── Inline tokens → TextRun[] ────────────────────────────────────────────────
function inlineToRuns(tokens = [], opts = {}) {
  const runs = [];
  for (const tok of tokens) {
    switch (tok.type) {
      case 'text':
      case 'escape':
        // text tokens can have child tokens (e.g. bold inside a paragraph)
        if (tok.tokens) {
          runs.push(...inlineToRuns(tok.tokens, opts));
        } else {
          runs.push(new TextRun({ text: tok.text ?? tok.raw ?? '', ...opts }));
        }
        break;
      case 'strong':
        runs.push(...inlineToRuns(tok.tokens, { ...opts, bold: true }));
        break;
      case 'em':
        runs.push(...inlineToRuns(tok.tokens, { ...opts, italics: true }));
        break;
      case 'codespan':
        runs.push(new TextRun({ text: tok.text, ...opts, font: 'Courier New' }));
        break;
      case 'link':
        runs.push(...inlineToRuns(tok.tokens, {
          ...opts,
          underline: { type: UnderlineType.SINGLE },
          color: '0563C1',
        }));
        break;
      case 'del':
        runs.push(...inlineToRuns(tok.tokens, { ...opts, strike: true }));
        break;
      case 'br':
        runs.push(new TextRun({ break: 1 }));
        break;
      default:
        if (tok.tokens) {
          runs.push(...inlineToRuns(tok.tokens, opts));
        } else if (tok.raw || tok.text) {
          runs.push(new TextRun({ text: tok.raw ?? tok.text ?? '', ...opts }));
        }
    }
  }
  return runs;
}

const HEADING_LEVELS = [
  HeadingLevel.HEADING_1,
  HeadingLevel.HEADING_2,
  HeadingLevel.HEADING_3,
  HeadingLevel.HEADING_4,
  HeadingLevel.HEADING_5,
  HeadingLevel.HEADING_6,
];

// ── Block tokens → docx elements ────────────────────────────────────────────
function blockToElements(tokens, listLevel = 0) {
  const elements = [];

  for (const tok of tokens) {
    switch (tok.type) {

      case 'heading':
        elements.push(new Paragraph({
          children: inlineToRuns(tok.tokens),
          heading: HEADING_LEVELS[Math.min(tok.depth - 1, 5)],
        }));
        break;

      case 'paragraph':
        elements.push(new Paragraph({
          children: inlineToRuns(tok.tokens),
          spacing: { after: 120 },
        }));
        break;

      case 'list': {
        let num = typeof tok.start === 'number' ? tok.start : 1;
        for (const item of tok.items) {
          const firstToken = item.tokens?.[0];
          const inlineTokens = firstToken?.tokens ?? (firstToken ? [firstToken] : []);

          if (tok.ordered) {
            elements.push(new Paragraph({
              children: [
                new TextRun({ text: `${num}.\t` }),
                ...inlineToRuns(inlineTokens),
              ],
              indent: { left: (listLevel + 1) * 360, hanging: 360 },
              spacing: { after: 60 },
            }));
            num++;
          } else {
            elements.push(new Paragraph({
              children: inlineToRuns(inlineTokens),
              bullet: { level: listLevel },
              spacing: { after: 60 },
            }));
          }

          // Recurse into nested list
          const nested = item.tokens?.find(t => t.type === 'list');
          if (nested) {
            elements.push(...blockToElements([nested], listLevel + 1));
          }
        }
        break;
      }

      case 'code':
        tok.text.split('\n').forEach(line => {
          elements.push(new Paragraph({
            children: [new TextRun({ text: line || ' ', font: 'Courier New', size: 18 })],
            indent: { left: 360 },
            spacing: { after: 0, before: 0 },
          }));
        });
        elements.push(new Paragraph({ children: [], spacing: { after: 120 } }));
        break;

      case 'blockquote':
        for (const inner of tok.tokens) {
          if (inner.type === 'paragraph') {
            elements.push(new Paragraph({
              children: inlineToRuns(inner.tokens),
              indent: { left: 720 },
              border: {
                left: { style: BorderStyle.THICK, size: 8, color: 'AAAAAA', space: 1 },
              },
              spacing: { after: 120 },
            }));
          } else {
            elements.push(...blockToElements([inner]));
          }
        }
        break;

      case 'hr':
        elements.push(new Paragraph({
          children: [],
          border: {
            bottom: { style: BorderStyle.SINGLE, size: 6, color: 'CCCCCC', space: 1 },
          },
          spacing: { after: 200 },
        }));
        break;

      case 'table': {
        const rows = [];
        const headerCells = tok.header.map(cell =>
          new TableCell({ children: [new Paragraph({ children: inlineToRuns(cell.tokens) })] })
        );
        rows.push(new TableRow({ children: headerCells, tableHeader: true }));

        for (const row of tok.rows) {
          rows.push(new TableRow({
            children: row.map(cell =>
              new TableCell({ children: [new Paragraph({ children: inlineToRuns(cell.tokens) })] })
            ),
          }));
        }

        elements.push(new Table({ rows, width: { size: 100, type: WidthType.PERCENTAGE } }));
        elements.push(new Paragraph({ children: [], spacing: { after: 200 } }));
        break;
      }

      case 'space':
        break;

      default:
        if (tok.tokens) elements.push(...blockToElements(tok.tokens));
    }
  }

  return elements;
}

// ── Convert endpoint ─────────────────────────────────────────────────────────
app.post('/convert', upload.single('file'), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: 'No file uploaded' });

    const markdown = req.file.buffer.toString('utf-8');
    const tokens = marked.lexer(markdown);

    const doc = new Document({
      sections: [{ children: blockToElements(tokens) }],
    });

    const buffer = await Packer.toBuffer(doc);
    const baseName = path.basename(req.file.originalname, path.extname(req.file.originalname));

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', `attachment; filename="${baseName}.docx"`);
    res.send(buffer);
  } catch (err) {
    console.error('Conversion error:', err);
    res.status(500).json({ error: 'Conversion failed', detail: err.message });
  }
});

app.use((err, _req, res, _next) => {
  res.status(400).json({ error: err.message });
});

app.listen(PORT, () => {
  console.log(`markdoc running at http://localhost:${PORT}`);
});
