const express = require('express');
const cors = require('cors');
const multer = require('multer');
const crypto = require('crypto');
const ExcelJS = require('exceljs');
const { buildReport } = require('./lib/reportBuilder');

const app = express();
const PORT = process.env.PORT || 4000;

// Comma-separated list of allowed frontend origins, e.g.
// "https://your-app.vercel.app,http://localhost:5173"
// Falls back to allowing all origins if not set (fine for local dev).
const allowedOrigins = (process.env.CORS_ORIGIN || '')
  .split(',')
  .map((o) => o.trim())
  .filter(Boolean);

app.use(
  cors(
    allowedOrigins.length
      ? {
          origin: (origin, callback) => {
            // allow same-origin / server-to-server requests with no origin header
            if (!origin || allowedOrigins.includes(origin)) return callback(null, true);
            callback(new Error(`Origin ${origin} not allowed by CORS`));
          },
        }
      : {}
  )
);
app.use(express.json());

const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 60 * 1024 * 1024 }, // 60MB
});

// uploadId -> { workbook, fileName, uploadedAt }
const store = new Map();
const MAX_AGE_MS = 2 * 60 * 60 * 1000; // 2 hours

setInterval(() => {
  const now = Date.now();
  for (const [id, entry] of store.entries()) {
    if (now - entry.uploadedAt > MAX_AGE_MS) store.delete(id);
  }
}, 15 * 60 * 1000).unref();

app.post('/api/upload', upload.single('file'), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: 'No file uploaded.' });

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(req.file.buffer);

    const uploadId = crypto.randomUUID();
    store.set(uploadId, {
      workbook,
      fileName: req.file.originalname,
      uploadedAt: Date.now(),
    });

    const sheetNames = [];
    workbook.eachSheet((ws) => sheetNames.push(ws.name));

    res.json({ uploadId, fileName: req.file.originalname, sheetCount: sheetNames.length, sheetNames });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: 'Could not read that file. Is it a valid .xlsx workbook?' });
  }
});

function normalizeRange(body) {
  // Accepts either { fromDate, toDate } or a single { date } for backward
  // compatibility; a single date is treated as a one-day range.
  const fromDate = body.fromDate || body.date;
  const toDate = body.toDate || body.date;
  if (!fromDate || !toDate) return null;
  if (fromDate > toDate) return { fromDate: toDate, toDate: fromDate }; // tolerate swapped inputs
  return { fromDate, toDate };
}

app.post('/api/report', (req, res) => {
  const { uploadId } = req.body || {};
  const range = normalizeRange(req.body || {});
  if (!uploadId || !range) {
    return res.status(400).json({ error: 'uploadId, fromDate and toDate are required.' });
  }

  const entry = store.get(uploadId);
  if (!entry) return res.status(404).json({ error: 'Upload not found. Please re-upload the file.' });

  const { rows, warnings } = buildReport(entry.workbook, range.fromDate, range.toDate);
  res.json({ rows, warnings, count: rows.length });
});

app.post('/api/download', async (req, res) => {
  const { uploadId } = req.body || {};
  const range = normalizeRange(req.body || {});
  if (!uploadId || !range) {
    return res.status(400).json({ error: 'uploadId, fromDate and toDate are required.' });
  }

  const entry = store.get(uploadId);
  if (!entry) return res.status(404).json({ error: 'Upload not found. Please re-upload the file.' });

  const { rows } = buildReport(entry.workbook, range.fromDate, range.toDate);

  const outWorkbook = new ExcelJS.Workbook();
  const label = range.fromDate === range.toDate ? range.fromDate : `${range.fromDate}_to_${range.toDate}`;
  const sheet = outWorkbook.addWorksheet(`Report ${label}`.slice(0, 31));

  sheet.columns = [
    { header: 'Date', key: 'date', width: 14 },
    { header: 'Sheet Name', key: 'sheet', width: 32 },
    { header: 'Receiver / Issuer', key: 'person', width: 28 },
    { header: 'Receive Quantity', key: 'receiveQty', width: 16 },
    { header: 'Receive Net Weight (kg)', key: 'receiveNetWeightKg', width: 20 },
    { header: 'Issue Quantity', key: 'issueQty', width: 16 },
    { header: 'Issue Net Weight (kg)', key: 'issueNetWeightKg', width: 20 },
    { header: 'Total Receive Quantity (Sheet)', key: 'totalReceiveQty', width: 24 },
    { header: 'Total Receive Net Weight (kg) (Sheet)', key: 'totalReceiveNetWeightKg', width: 26 },
    { header: 'Total Issue Quantity (Sheet)', key: 'totalIssueQty', width: 24 },
    { header: 'Total Issue Net Weight (kg) (Sheet)', key: 'totalIssueNetWeightKg', width: 26 },
  ];
  sheet.getRow(1).font = { bold: true };
  sheet.autoFilter = { from: 'A1', to: 'K1' };

  rows.forEach((r) => sheet.addRow(r));

  const buffer = await outWorkbook.xlsx.writeBuffer();

  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', `attachment; filename="report-${label}.xlsx"`);
  res.send(Buffer.from(buffer));
});

// Health check — used by Render for zero-downtime deploys, and by an
// external uptime pinger (see README) to keep the free-tier instance awake.
app.get('/api/health', (req, res) => {
  res.json({ ok: true, uptimeSeconds: Math.round(process.uptime()), timestamp: new Date().toISOString() });
});

app.listen(PORT, () => {
  console.log(`Backend listening on http://localhost:${PORT}`);
});
