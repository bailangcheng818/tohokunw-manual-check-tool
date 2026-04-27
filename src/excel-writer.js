'use strict';

const ExcelJS = require('exceljs');
const fs      = require('fs');
const path    = require('path');

/**
 * Normalize a cell value: objects and arrays are JSON-stringified so they
 * are stored as readable text in Excel instead of "[object Object]".
 */
function normalizeCell(v) {
  if (v !== null && typeof v === 'object') return JSON.stringify(v, null, 0);
  return v;
}

/**
 * When a caller (e.g. Dify) spreads a JSON array into separate row elements,
 * the row arrives longer than the header.  Collapse every element from
 * headerLength-1 onwards into a single value so it lands in one cell.
 *
 * Example: header has 11 cols, row arrives with 15 cols because a 5-item
 * JSON array was spread → combine r[10..14] into one array, store at col 11.
 */
function collapseToHeader(r, headerLength) {
  if (!headerLength || r.length <= headerLength) return r;
  const tail     = r.slice(headerLength - 1);          // everything from last col on
  const combined = tail.length === 1 ? tail[0] : tail; // keep scalar, wrap multiples
  return [...r.slice(0, headerLength - 1), combined];
}

/**
 * Find the last row that contains at least one non-empty, non-whitespace cell value.
 * Uses strict check so that rows with only empty strings or null values are ignored,
 * preventing new data from being inserted far below the real data in template files.
 */
function findLastDataRow(ws) {
  let lastRow = 1; // fallback to row 1 (header) if no data rows found
  ws.eachRow({ includeEmpty: true }, (row, rowNumber) => {
    let hasContent = false;
    row.eachCell({ includeEmpty: false }, (cell) => {
      const v = cell.value;
      if (v !== null && v !== undefined && String(v).trim() !== '') {
        hasContent = true;
      }
    });
    if (hasContent) lastRow = rowNumber;
  });
  return lastRow;
}

/**
 * Append one or more rows to an existing Excel file.
 * If the file doesn't exist it will be created with a header row (if provided).
 *
 * @param {object} opts
 * @param {string}   opts.file_path   - Absolute path to the .xlsx file
 * @param {string}   [opts.sheet]     - Sheet name (default: first sheet)
 * @param {Array}    opts.row         - Single row: array of cell values
 * @param {Array}    [opts.rows]      - Multiple rows: array of row arrays (alternative to row)
 * @param {Array}    [opts.header]    - Column headers written only when creating a new file
 * @returns {Promise<{ file_path, sheet, appended_rows, total_rows }>}
 */
async function appendRows({ file_path, sheet, row, rows, data, header }) {
  if (!file_path) throw new Error('file_path is required');

  // ── key-value mode: resolve data object against the file's actual header row ──
  if (data && typeof data === 'object' && !Array.isArray(data)) {
    if (!fs.existsSync(file_path)) {
      throw new Error(`key-value data mode requires an existing file with a header row: ${file_path}`);
    }
    const headers = await getHeaders(file_path, sheet);
    row = resolveByHeader(data, headers);
  }

  const rawRows = rows || (row ? [row] : null);
  if (!rawRows || rawRows.length === 0) {
    throw new Error('Provide row, rows, or data (key-value object)');
  }
  // Collapse rows that are longer than the header (e.g. Dify spreading JSON arrays).
  const headerLen = header && header.length > 0 ? header.length : 0;
  const rowsToAppend = headerLen
    ? rawRows.map(r => collapseToHeader(r, headerLen))
    : rawRows;

  const wb        = new ExcelJS.Workbook();
  const fileExists = fs.existsSync(file_path);

  // Ensure parent directory exists
  fs.mkdirSync(path.dirname(file_path), { recursive: true });

  if (fileExists) {
    await wb.xlsx.readFile(file_path);
  }

  // Resolve target worksheet
  let ws;
  if (sheet) {
    ws = wb.getWorksheet(sheet) || wb.addWorksheet(sheet);
  } else {
    ws = wb.worksheets[0] || wb.addWorksheet('Sheet1');
  }

  // Write header only when creating the file fresh
  if (!fileExists && header && header.length > 0) {
    ws.addRow(header);
  }

  if (fileExists) {
    // When appending to an existing file, find the actual last data row
    // (ws.rowCount may include formatted-but-empty rows, causing rows to be
    // inserted far below the real data)
    let insertAt = findLastDataRow(ws) + 1;
    for (const r of rowsToAppend) {
      const wsRow = ws.getRow(insertAt++);
      // Set each cell individually to prevent ExcelJS from spreading arrays
      // across multiple columns when using bulk wsRow.values assignment.
      for (let i = 0; i < r.length; i++) {
        wsRow.getCell(i + 1).value = normalizeCell(r[i]);
      }
      wsRow.commit();
    }
  } else {
    for (const r of rowsToAppend) {
      const wsRow = ws.getRow(ws.rowCount + 1);
      for (let i = 0; i < r.length; i++) {
        wsRow.getCell(i + 1).value = normalizeCell(r[i]);
      }
      wsRow.commit();
    }
  }

  await wb.xlsx.writeFile(file_path);

  return {
    file_path,
    sheet:         ws.name,
    appended_rows: rowsToAppend.length,
    total_rows:    findLastDataRow(ws),
  };
}

/**
 * Update specific cells in an existing row without touching other cells.
 *
 * @param {object} opts
 * @param {string}   opts.file_path   - Absolute path to the .xlsx file
 * @param {string}   [opts.sheet]     - Sheet name (default: first sheet)
 * @param {number}   [opts.row_number]- Ignored. Always targets the last data row.
 * @param {string|number} [opts.column] - Column name (matched against header row) or 1-based column index.
 * @param {*}        [opts.value]     - New cell value. Used when column is specified.
 * @param {Array}    [opts.header]    - Column names for multi-cell update mode.
 * @param {Array}    [opts.row]       - Values parallel to header; empty strings are skipped.
 * @returns {Promise<{ file_path, sheet, updated_cells, excel_row }>}
 */
async function updateCell({ file_path, sheet, column, value, header, row }) {
  if (!file_path) throw new Error('file_path is required');

  const wb = new ExcelJS.Workbook();
  if (!fs.existsSync(file_path)) throw new Error(`File not found: ${file_path}`);
  await wb.xlsx.readFile(file_path);

  let ws;
  if (sheet) {
    ws = wb.getWorksheet(sheet);
    if (!ws) throw new Error(`Sheet "${sheet}" not found`);
  } else {
    ws = wb.worksheets[0];
    if (!ws) throw new Error('No worksheets found');
  }

  // Find the actual last row with data (ignores formatted-but-empty rows)
  const lastDataRow = findLastDataRow(ws);

  // Always target the last row that contains data, ignoring row_number.
  const excelRow = lastDataRow;

  if (excelRow < 2) {
    throw new Error('No data rows found (only a header row or empty sheet)');
  }

  const targetRow = ws.getRow(excelRow);
  let updatedCells = 0;

  if (column !== undefined) {
    // Single-cell mode: find column index
    let colIndex;
    if (typeof column === 'number') {
      colIndex = column;
    } else {
      // Find column by name in header row (row 1)
      const headerRow = ws.getRow(1);
      colIndex = null;
      headerRow.eachCell({ includeEmpty: true }, (cell, idx) => {
        if (cell.value === column) colIndex = idx;
      });
      if (!colIndex) throw new Error(`Column "${column}" not found in header row`);
    }
    targetRow.getCell(colIndex).value = normalizeCell(value);
    updatedCells = 1;
  } else if (header && row) {
    // Multi-cell mode: update only non-empty cells
    const headerRow = ws.getRow(1);
    const colMap = {};
    headerRow.eachCell({ includeEmpty: true }, (cell, idx) => {
      if (cell.value) colMap[cell.value] = idx;
    });

    for (let i = 0; i < header.length; i++) {
      const cellValue = row[i];
      if (cellValue === '' || cellValue === null || cellValue === undefined) continue;
      const colName = header[i];
      const colIndex = colMap[colName];
      if (!colIndex) throw new Error(`Column "${colName}" not found in header row`);
      targetRow.getCell(colIndex).value = normalizeCell(cellValue);
      updatedCells++;
    }
  } else {
    throw new Error('Provide either (column + value) or (header + row)');
  }

  targetRow.commit();
  await wb.xlsx.writeFile(file_path);

  return { file_path, sheet: ws.name, updated_cells: updatedCells, excel_row: excelRow };
}

/**
 * Read the header row (row 1) of an Excel file and return an array of
 * { index, name } objects (1-based column index).
 *
 * @param {string} file_path - Absolute path to the .xlsx file
 * @param {string} [sheet]   - Sheet name (default: first sheet)
 * @returns {Promise<Array<{ index: number, name: string }>>}
 */
async function getHeaders(file_path, sheet) {
  if (!fs.existsSync(file_path)) throw new Error(`File not found: ${file_path}`);
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(file_path);
  let ws;
  if (sheet) {
    ws = wb.getWorksheet(sheet);
    if (!ws) throw new Error(`Sheet "${sheet}" not found`);
  } else {
    ws = wb.worksheets[0];
    if (!ws) throw new Error('No worksheets found');
  }
  const headers = [];
  ws.getRow(1).eachCell({ includeEmpty: true }, (cell, idx) => {
    headers.push({ index: idx, name: String(cell.value ?? '') });
  });
  return headers;
}

/**
 * Map a key-value data object to a positional row array using a headers list
 * returned by getHeaders().  Unknown keys are silently ignored; missing columns
 * default to empty string.
 *
 * @param {object} data    - e.g. { "入力月日": "2025-06-15", "氏名": "東北 太郎" }
 * @param {Array}  headers - output of getHeaders()
 * @returns {Array} positional row array aligned to header columns
 */
function resolveByHeader(data, headers) {
  const colMap = {};
  for (const h of headers) {
    if (h.name) colMap[h.name] = h.index;
  }
  const maxCol = headers.length > 0 ? Math.max(...headers.map(h => h.index)) : 0;
  const row = new Array(maxCol).fill('');
  for (const [key, value] of Object.entries(data)) {
    const colIndex = colMap[key];
    if (colIndex !== undefined) {
      row[colIndex - 1] = value;
    }
    // keys not found in header are silently ignored
  }
  return row;
}

/**
 * Search a sheet for the first row where a named column matches a value.
 *
 * @param {object} opts
 * @param {string} opts.file_path - Absolute path to the .xlsx file
 * @param {string} [opts.sheet]   - Sheet name (default: first sheet)
 * @param {string} opts.column    - Header name to match against
 * @param {string} opts.value     - Value to search for (compared as string)
 * @param {number} [opts.header_row=1] - Row number containing headers
 * @returns {Promise<{ found: boolean, row_number?: number, data?: object }>}
 */
async function findRow({ file_path, sheet, column, value, header_row = 1 }) {
  if (!fs.existsSync(file_path)) throw new Error(`File not found: ${file_path}`);
  const headers = await getHeaders(file_path, sheet, header_row);
  const colDef = headers.find(h => h.name === column);
  if (!colDef) return { found: false };

  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(file_path);
  const ws = sheet ? wb.getWorksheet(sheet) : wb.worksheets[0];
  if (!ws) throw new Error('Sheet not found');

  const lastRow = findLastDataRow(ws);
  for (let r = header_row + 1; r <= lastRow; r++) {
    const wsRow = ws.getRow(r);
    const cellVal = String(wsRow.getCell(colDef.index).value ?? '');
    if (cellVal === String(value)) {
      const data = {};
      headers.forEach(h => { data[h.name] = wsRow.getCell(h.index).value ?? null; });
      return { found: true, row_number: r, data };
    }
  }
  return { found: false };
}

/**
 * Update multiple cells in a specific row using a key-value data object.
 *
 * @param {object} opts
 * @param {string} opts.file_path  - Absolute path to the .xlsx file
 * @param {string} [opts.sheet]    - Sheet name (default: first sheet)
 * @param {number} opts.row_number - Target row (1-based, must be >= 2)
 * @param {object} opts.data       - Key-value pairs keyed by header name
 * @param {number} [opts.header_row=1] - Row number containing headers
 * @returns {Promise<{ file_path, sheet, row_number, updated_columns, unmatched_keys }>}
 */
async function updateRow({ file_path, sheet, row_number, data, header_row = 1 }) {
  if (!file_path) throw new Error('file_path is required');
  if (!row_number || row_number < 2) throw new Error('row_number >= 2 is required');
  if (!data || typeof data !== 'object') throw new Error('data object is required');
  if (!fs.existsSync(file_path)) throw new Error(`File not found: ${file_path}`);

  const headers = await getHeaders(file_path, sheet, header_row);
  const colMap = {};
  for (const h of headers) { if (h.name) colMap[h.name] = h.index; }

  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(file_path);
  const ws = sheet ? wb.getWorksheet(sheet) : wb.worksheets[0];
  if (!ws) throw new Error('Sheet not found');

  const wsRow = ws.getRow(row_number);
  const updatedColumns = [];
  const unmatchedKeys = [];
  for (const [key, val] of Object.entries(data)) {
    const colIndex = colMap[key];
    if (!colIndex) { unmatchedKeys.push(key); continue; }
    wsRow.getCell(colIndex).value = normalizeCell(val);
    updatedColumns.push(key);
  }
  wsRow.commit();
  await wb.xlsx.writeFile(file_path);

  return { file_path, sheet: ws.name, row_number, updated_columns: updatedColumns, unmatched_keys: unmatchedKeys };
}

module.exports = { appendRows, updateCell, findRow, updateRow, getHeaders, resolveByHeader };
