const ExcelJS = require('exceljs');
const { normalizeHeader, mapColumns } = require('./parser');

const HEADER_SCAN_ROWS = 10;

/** Normalize a cell's date value (Date object, Excel serial, or a few common
 * string formats) down to a plain 'YYYY-MM-DD' string, or null if it can't
 * be read as a date at all. */
function cellToIsoDate(value) {
  if (value === null || value === undefined || value === '') return null;

  if (value instanceof Date && !isNaN(value.getTime())) {
    return value.toISOString().slice(0, 10);
  }

  if (typeof value === 'object' && value.result instanceof Date) {
    return value.result.toISOString().slice(0, 10);
  }

  // exceljs's streaming reader (with styles ignored, for memory) can't
  // auto-convert date-formatted cells, so dates often arrive as raw Excel
  // serial numbers instead of Date objects. Excel's epoch is 1899-12-30
  // (intentionally reusing Excel's own 1900-leap-year quirk so the numbers
  // round-trip correctly). Only treat this as a date for a plausible range
  // (Excel serials from roughly 1950 to 2100) to avoid misreading an
  // unrelated small integer (like a lot number) as a date.
  const numeric = typeof value === 'number' ? value : typeof value === 'object' && typeof value.result === 'number' ? value.result : null;
  if (numeric !== null && numeric > 18000 && numeric < 73000) {
    const d = new Date(Date.UTC(1899, 11, 30) + numeric * 86400000);
    if (!isNaN(d.getTime())) return d.toISOString().slice(0, 10);
  }

  if (typeof value === 'string') {
    const s = value.trim();

    let m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
    if (m) {
      const [, d, mo, y] = m;
      return `${y}-${mo.padStart(2, '0')}-${d.padStart(2, '0')}`;
    }

    m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
    if (m) return m[0];
  }

  return null;
}

function cellText(value) {
  if (value === null || value === undefined) return '';
  if (typeof value === 'object') {
    if (value.richText) return value.richText.map((r) => r.text).join('').trim();
    if (value.text) return String(value.text).trim();
    if (value.result !== undefined) return String(value.result).trim();
  }
  return String(value).trim();
}

function cellNumber(value) {
  if (value === null || value === undefined || value === '') return null;
  if (typeof value === 'object' && value.result !== undefined) value = value.result;
  const n = Number(value);
  return Number.isFinite(n) ? n : null;
}

/** Build {col, text} headers from a streamed row's `.values` array
 * (1-indexed; index 0 is unused, matching exceljs's convention). */
function headersFromRowValues(rowValues) {
  const headers = [];
  for (let col = 1; col < rowValues.length; col++) {
    const v = rowValues[col];
    if (v === null || v === undefined || v === '') continue;
    const text = normalizeHeader(cellText(v));
    if (text) headers.push({ col, text });
  }
  return headers;
}

function round2(n) {
  return Math.round((n + Number.EPSILON) * 100) / 100;
}

function newWorkbookReader(filePath) {
  return new ExcelJS.stream.xlsx.WorkbookReader(filePath, {
    entries: 'emit',
    sharedStrings: 'cache', // needed for correct string values
    styles: 'ignore', // we don't need formatting, saves memory
    hyperlinks: 'ignore',
    worksheets: 'emit',
  });
}

/**
 * Stream every worksheet in the file at filePath, scanning for rows whose
 * DATE column falls within [fromDate, toDate] inclusive. Never holds more
 * than one row (or a small buffer) in memory at a time, so it scales to
 * workbooks with hundreds/thousands of sheets without loading the whole
 * file into RAM.
 */
async function buildReport(filePath, fromDate, toDate) {
  const rows = [];
  const warnings = [];

  const workbookReader = newWorkbookReader(filePath);

  for await (const worksheetReader of workbookReader) {
    const sheetName = worksheetReader.name;
    let map = null;
    let rowIndex = 0;
    let foundHeader = false;

    for await (const row of worksheetReader) {
      rowIndex++;
      const values = row.values;

      if (!map) {
        if (rowIndex > HEADER_SCAN_ROWS) continue; // give up looking, but keep draining the sheet
        const headers = headersFromRowValues(values);
        if (headers.some((h) => h.text === 'DATE')) {
          map = mapColumns(headers);
          foundHeader = true;
        }
        continue;
      }

      const dateCellValue = values[map.date];
      const isoDate = cellToIsoDate(dateCellValue);
      if (!isoDate || isoDate < fromDate || isoDate > toDate) continue;

      const receiveQty = map.receiveQty ? cellNumber(values[map.receiveQty]) : null;
      const issueQty = map.issueQty ? cellNumber(values[map.issueQty]) : null;
      const receiveWeight = map.receiveWeight ? cellNumber(values[map.receiveWeight]) : null;
      const issueWeight = map.issueWeight ? cellNumber(values[map.issueWeight]) : null;
      const personName = map.personName ? cellText(values[map.personName]) : '';

      const hasAnyValue = [receiveQty, issueQty, receiveWeight, issueWeight].some(
        (v) => v !== null && v !== 0
      );
      if (!hasAnyValue) continue;

      rows.push({
        date: isoDate,
        sheet: sheetName,
        person: personName || '(not found)',
        receiveQty: receiveQty ?? null,
        receiveNetWeightKg: receiveWeight ?? null,
        issueQty: issueQty ?? null,
        issueNetWeightKg: issueWeight ?? null,
      });
    }

    if (!foundHeader) {
      warnings.push(`"${sheetName}": couldn't find a header row (no DATE column) — skipped.`);
    }
  }

  rows.sort((a, b) => (a.date < b.date ? -1 : a.date > b.date ? 1 : 0));

  // Per-sheet aggregate totals across the matched date range, stamped onto
  // every row belonging to that sheet.
  const totalsBySheet = new Map();
  for (const r of rows) {
    if (!totalsBySheet.has(r.sheet)) {
      totalsBySheet.set(r.sheet, {
        totalReceiveQty: 0,
        totalReceiveNetWeightKg: 0,
        totalIssueQty: 0,
        totalIssueNetWeightKg: 0,
      });
    }
    const t = totalsBySheet.get(r.sheet);
    t.totalReceiveQty += r.receiveQty ?? 0;
    t.totalReceiveNetWeightKg += r.receiveNetWeightKg ?? 0;
    t.totalIssueQty += r.issueQty ?? 0;
    t.totalIssueNetWeightKg += r.issueNetWeightKg ?? 0;
  }

  const roundedRows = rows.map((r) => {
    const t = totalsBySheet.get(r.sheet);
    return {
      ...r,
      totalReceiveQty: round2(t.totalReceiveQty),
      totalReceiveNetWeightKg: round2(t.totalReceiveNetWeightKg),
      totalIssueQty: round2(t.totalIssueQty),
      totalIssueNetWeightKg: round2(t.totalIssueNetWeightKg),
    };
  });

  return { rows: roundedRows, warnings };
}

/** Quick pass to list sheet names (used right after upload, just so the UI
 * can show "N sheets loaded"). Streams the whole file like buildReport does
 * — no shortcut through it — since exceljs's streaming reader expects each
 * worksheet to be fully drained before moving to the next one. */
async function listSheetNames(filePath) {
  const sheetNames = [];
  const workbookReader = newWorkbookReader(filePath);
  for await (const worksheetReader of workbookReader) {
    sheetNames.push(worksheetReader.name);
    for await (const _row of worksheetReader) {
      // drain without retaining anything
    }
  }
  return sheetNames;
}

module.exports = { buildReport, listSheetNames, cellToIsoDate };