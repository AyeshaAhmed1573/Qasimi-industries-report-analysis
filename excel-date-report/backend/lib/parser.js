/**
 * Handles the messy reality of this workbook: every sheet is a "daily
 * balance sheet" for one yarn lot, but the exact header text/order drifts
 * from sheet to sheet (RECEIVER NAME vs RECEIVER/ISSUE NAME vs missing,
 * CONES vs BAL. CONES, a repeated "TOTAL NET WEIGHT / KG" label used for
 * both receive and issue, etc). Instead of hardcoding column positions,
 * we read the header row on each sheet and fuzzy-match it every time.
 */

const HEADER_SCAN_ROWS = 10;

function normalizeHeader(value) {
  if (value === null || value === undefined) return '';
  return String(value)
    .toUpperCase()
    .replace(/\n/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

/** Find the row that contains the column headers (the row with "DATE" in it). */
function findHeaderRow(worksheet) {
  for (let r = 1; r <= HEADER_SCAN_ROWS; r++) {
    const row = worksheet.getRow(r);
    let hasDate = false;
    row.eachCell({ includeEmpty: false }, (cell) => {
      if (normalizeHeader(cell.value) === 'DATE') hasDate = true;
    });
    if (hasDate) return r;
  }
  return null;
}

/** Collect {col, text} for every non-empty header cell in the header row. */
function collectHeaders(worksheet, headerRowNum) {
  const row = worksheet.getRow(headerRowNum);
  const headers = [];
  row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
    headers.push({ col: colNumber, text: normalizeHeader(cell.value) });
  });
  return headers;
}

function findFirst(headers, predicate) {
  const hit = headers.find(predicate);
  return hit ? hit.col : null;
}

function findAll(headers, predicate) {
  return headers.filter(predicate).sort((a, b) => a.col - b.col);
}

/**
 * Build a field -> column-number map for one sheet's header row.
 * Every lookup is done by matching header TEXT, never by fixed position.
 */
function mapColumns(headers) {
  const map = {};

  map.date = findFirst(headers, (h) => h.text === 'DATE');

  map.receiveQty = findFirst(
    headers,
    (h) => h.text.includes('RECEIVE') && !h.text.includes('WEIGHT') && !h.text.includes('NAME')
  );

  map.issueQty = findFirst(
    headers,
    (h) =>
      h.text.includes('ISSUE') &&
      !h.text.includes('WEIGHT') &&
      !h.text.includes('NAME') &&
      !h.text.includes('RECEIVE')
  );

  map.receiveWeight = findFirst(
    headers,
    (h) => h.text.includes('RECEIVE') && h.text.includes('WEIGHT')
  );
  map.issueWeight = findFirst(
    headers,
    (h) => h.text.includes('ISSUE') && h.text.includes('WEIGHT')
  );

  // Some sheets reuse the exact same generic "TOTAL NET WEIGHT / KG" label
  // for both the receive and issue columns. If we couldn't find a clearly
  // labeled receive/issue weight column, fall back to picking generic
  // weight columns in left-to-right order.
  if (!map.receiveWeight || !map.issueWeight) {
    const generic = findAll(
      headers,
      (h) => h.text.includes('WEIGHT') && h.text.includes('NET')
    );
    if (generic.length >= 2) {
      if (!map.receiveWeight) map.receiveWeight = generic[0].col;
      if (!map.issueWeight) map.issueWeight = generic[1].col;
    } else if (generic.length === 1) {
      if (!map.receiveWeight) map.receiveWeight = generic[0].col;
      if (!map.issueWeight) map.issueWeight = generic[0].col;
    }
  }

  // Person name: prefer a column explicitly about receiver/issuer, else
  // any column with "NAME" in it, else a couple of known fallback labels
  // seen on older sheets that don't have a dedicated name column.
  map.personName =
    findFirst(headers, (h) => h.text.includes('NAME') && (h.text.includes('RECEIVER') || h.text.includes('ISSUE'))) ||
    findFirst(headers, (h) => h.text.includes('NAME')) ||
    // Some sheets (older "Galaxy"-style layouts) put the party name under a
    // header like "ISSUE / RECEIVE" instead of anything with "NAME" in it —
    // but only when that column isn't already the issue-quantity column.
    findFirst(
      headers,
      (h) => h.text.includes('ISSUE') && h.text.includes('RECEIVE') && h.col !== map.issueQty && h.col !== map.receiveQty
    ) ||
    findFirst(headers, (h) => h.text.includes('RETURN')) ||
    null;

  return map;
}

module.exports = { normalizeHeader, findHeaderRow, collectHeaders, mapColumns };
