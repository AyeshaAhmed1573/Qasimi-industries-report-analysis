const { findHeaderRow, collectHeaders, mapColumns } = require('./parser');

/** Normalize a cell's date value (Date object, Excel serial, or a few common
 * string formats) down to a plain 'YYYY-MM-DD' string, or null if it can't
 * be read as a date at all. */
function cellToIsoDate(value) {
  if (value === null || value === undefined || value === '') return null;

  if (value instanceof Date && !isNaN(value.getTime())) {
    return value.toISOString().slice(0, 10);
  }

  if (typeof value === 'object' && value.result instanceof Date) {
    // formula result
    return value.result.toISOString().slice(0, 10);
  }

  if (typeof value === 'string') {
    const s = value.trim();

    // DD/MM/YYYY or D/M/YYYY
    let m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
    if (m) {
      const [, d, mo, y] = m;
      return `${y}-${mo.padStart(2, '0')}-${d.padStart(2, '0')}`;
    }

    // already ISO-ish
    m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
    if (m) return m[0];
  }

  return null;
}

function cellText(cell) {
  if (cell === null || cell === undefined) return '';
  if (typeof cell === 'object') {
    if (cell.text) return String(cell.text).trim();
    if (cell.result !== undefined) return String(cell.result).trim();
  }
  return String(cell).trim();
}

function cellNumber(cell) {
  if (cell === null || cell === undefined || cell === '') return null;
  if (typeof cell === 'object' && cell.result !== undefined) cell = cell.result;
  const n = Number(cell);
  return Number.isFinite(n) ? n : null;
}

/**
 * Scan every worksheet in the workbook for rows whose DATE column falls
 * within [fromDate, toDate] inclusive (both 'YYYY-MM-DD'; pass the same
 * value for both to match a single day). Returns { rows, warnings }.
 *
 * Each matched sheet row becomes exactly one output row, with the
 * receive/issue quantity and net weight kept in their own columns — no
 * splitting into separate "Receive"/"Issue" rows, so a row where a single
 * physical entry carries both (rare, but happens) doesn't get duplicated.
 */
function buildReport(workbook, fromDate, toDate) {
  const rows = [];
  const warnings = [];

  workbook.eachSheet((worksheet) => {
    const sheetName = worksheet.name;
    const headerRowNum = findHeaderRow(worksheet);

    if (!headerRowNum) {
      warnings.push(`"${sheetName}": couldn't find a header row (no DATE column) — skipped.`);
      return;
    }

    const headers = collectHeaders(worksheet, headerRowNum);
    const map = mapColumns(headers);

    if (!map.date) {
      warnings.push(`"${sheetName}": couldn't identify the DATE column — skipped.`);
      return;
    }

    worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
      if (rowNumber <= headerRowNum) return;

      const dateCellValue = row.getCell(map.date).value;
      const isoDate = cellToIsoDate(dateCellValue);
      if (!isoDate || isoDate < fromDate || isoDate > toDate) return;

      const receiveQty = map.receiveQty ? cellNumber(row.getCell(map.receiveQty).value) : null;
      const issueQty = map.issueQty ? cellNumber(row.getCell(map.issueQty).value) : null;
      const receiveWeight = map.receiveWeight ? cellNumber(row.getCell(map.receiveWeight).value) : null;
      const issueWeight = map.issueWeight ? cellNumber(row.getCell(map.issueWeight).value) : null;
      const personName = map.personName ? cellText(row.getCell(map.personName).value) : '';

      // Skip completely empty rows that merely matched on a blank/leftover
      // date cell with no actual quantity data.
      const hasAnyValue = [receiveQty, issueQty, receiveWeight, issueWeight].some(
        (v) => v !== null && v !== 0
      );
      if (!hasAnyValue) return;

      rows.push({
        date: isoDate,
        sheet: sheetName,
        person: personName || '(not found)',
        receiveQty: receiveQty ?? null,
        receiveNetWeightKg: receiveWeight ?? null,
        issueQty: issueQty ?? null,
        issueNetWeightKg: issueWeight ?? null,
      });
    });
  });

  rows.sort((a, b) => (a.date < b.date ? -1 : a.date > b.date ? 1 : 0));

  // Aggregate totals per sheet name, across the whole matched date range,
  // then stamp each row with its sheet's totals so the totals are visible
  // alongside every individual entry (not just as a separate summary).
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

function round2(n) {
  return Math.round((n + Number.EPSILON) * 100) / 100;
}

module.exports = { buildReport, cellToIsoDate };
