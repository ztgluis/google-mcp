import { google } from 'googleapis';
import { extractId } from '../google.js';
import { extractSheetName, getSheetId, columnToIndex, colIndexToLetter } from '../utils/range-helpers.js';

function sheets(auth) { return google.sheets({ version: 'v4', auth }); }
function drive(auth)  { return google.drive({ version: 'v3', auth }); }

// ── Data ─────────────────────────────────────────────────────────────────────

export async function readSheet(auth, { spreadsheetId, range }) {
  const id = extractId(spreadsheetId);
  const res = await sheets(auth).spreadsheets.values.get({ spreadsheetId: id, range: range || 'A1:Z1000' });
  const rows = res.data.values || [];
  if (!rows.length) return 'No data found.';
  return rows.map(r => r.join('\t')).join('\n');
}

export async function batchReadSheet(auth, { spreadsheetId, ranges, majorDimension, valueRenderOption }) {
  const id = extractId(spreadsheetId);
  const res = await sheets(auth).spreadsheets.values.batchGet({
    spreadsheetId: id, ranges,
    majorDimension: majorDimension || 'ROWS',
    valueRenderOption: valueRenderOption || 'FORMATTED_VALUE',
  });
  const vr = res.data.valueRanges || [];
  if (!vr.length) return 'No data found.';
  return vr.map(r => `Range: ${r.range}\n${(r.values || []).map(row => row.join('\t')).join('\n')}`).join('\n\n');
}

export async function appendSheetRow(auth, { spreadsheetId, sheetName, dataRange, values }) {
  const id = extractId(spreadsheetId);
  const api = sheets(auth);

  // Reliably append a row by finding the last occupied row and writing at lastRow + 1.
  // The Sheets API values.append has a table-detection bug on complex layouts (instructions
  // area + data area in overlapping columns), so we bypass it entirely.
  //
  // dataRange (optional): A1 range scoping the data area, e.g. "C8:G". When provided,
  // only that range is scanned and the row is written within those columns.
  // When omitted, frozen rows are used to infer where data starts.

  const meta = await api.spreadsheets.get({ spreadsheetId: id, fields: 'sheets.properties' });
  const allSheets = meta.data.sheets || [];
  const target = sheetName
    ? allSheets.find(s => s.properties?.title === sheetName)
    : allSheets[0];
  if (!target) throw new Error(`Sheet "${sheetName || '(first)'}" not found.`);
  const tabName = target.properties.title;

  let scanRange, startRow, startCol;

  if (dataRange) {
    // Caller specified the data area — use it directly
    scanRange = dataRange.includes('!') ? dataRange : `'${tabName}'!${dataRange}`;
    const match = dataRange.replace(/^[^!]*!/, '').match(/^([A-Z]+)(\d+)/i);
    startRow = match ? parseInt(match[2]) : 1;
    startCol = match ? match[1].toUpperCase() : 'A';
  } else {
    // Auto-detect: use frozen rows as the header boundary, then inspect the header row
    // to find the first column with data (handles sheets where data starts in col C, not A).
    const frozenRows = target.properties.gridProperties?.frozenRowCount || 0;
    startRow = frozenRows > 0 ? frozenRows + 1 : 1;
    startCol = 'A';
    if (frozenRows > 0) {
      const headerScan = await api.spreadsheets.values.get({
        spreadsheetId: id, range: `'${tabName}'!A${startRow}:Z${startRow}`,
      });
      const headerRow = headerScan.data.values?.[0] || [];
      const firstDataCol = headerRow.findIndex(c => c && String(c).trim() !== '');
      if (firstDataCol >= 0) startCol = colIndexToLetter(firstDataCol);
    }
    scanRange = `'${tabName}'!${startCol}${startRow}:Z`;
  }

  const scan = await api.spreadsheets.values.get({ spreadsheetId: id, range: scanRange });
  const rows = scan.data.values || [];

  // Find the last row with a non-empty value in the first column of the scan range.
  // Using the first column (typically the date) avoids being fooled by reference content
  // that lives in other columns below the data area.
  let lastOffset = -1;
  for (let i = 0; i < rows.length; i++) {
    if (rows[i] && rows[i][0] && String(rows[i][0]).trim() !== '') {
      lastOffset = i;
    }
  }

  const nextRow = startRow + lastOffset + 1;
  const endCol = colIndexToLetter(columnToIndex(startCol) + values.length - 1);
  const writeRange = `'${tabName}'!${startCol}${nextRow}:${endCol}${nextRow}`;

  await api.spreadsheets.values.update({
    spreadsheetId: id, range: writeRange, valueInputOption: 'USER_ENTERED',
    requestBody: { values: [values] },
  });
  return `Row appended at ${writeRange} (row ${nextRow}).`;
}

export async function updateSheetCell(auth, { spreadsheetId, range, values }) {
  const id = extractId(spreadsheetId);
  await sheets(auth).spreadsheets.values.update({
    spreadsheetId: id, range, valueInputOption: 'USER_ENTERED', requestBody: { values },
  });
  return `Updated ${range}.`;
}

export async function batchUpdateSheet(auth, { spreadsheetId, data, valueInputOption }) {
  const id = extractId(spreadsheetId);
  const res = await sheets(auth).spreadsheets.values.batchUpdate({
    spreadsheetId: id,
    requestBody: {
      valueInputOption: valueInputOption || 'USER_ENTERED',
      data: data.map(item => ({ range: item.range, values: item.values })),
    },
  });
  const total = (res.data.responses || []).reduce((s, r) => s + (r.updatedCells || 0), 0);
  return `Batch update complete. ${total} cells updated across ${data.length} ranges.`;
}

export async function clearSheet(auth, { spreadsheetId, range }) {
  const id = extractId(spreadsheetId);
  const res = await sheets(auth).spreadsheets.values.clear({ spreadsheetId: id, range });
  return `Cleared ${res.data.clearedRange || range}.`;
}

export async function insertRows(auth, { spreadsheetId, range, rows = 1, position = 'BEFORE', inheritFromBefore = false, values, valueInputOption }) {
  const id = extractId(spreadsheetId);
  const api = sheets(auth);
  const { sheetName, range: cellRange } = extractSheetName(range);
  const sheetId = await getSheetId(api, id, sheetName);

  const colMatch = cellRange.match(/^([A-Z]+)(\d+)/i);
  const anchorRow = colMatch ? parseInt(colMatch[2]) - 1 : 0;
  const anchorCol = colMatch ? columnToIndex(colMatch[1].toUpperCase()) : 0;
  const startIndex = position === 'AFTER' ? anchorRow + 1 : anchorRow;

  await api.spreadsheets.batchUpdate({
    spreadsheetId: id,
    requestBody: { requests: [{ insertDimension: { range: { sheetId, dimension: 'ROWS', startIndex, endIndex: startIndex + rows }, inheritFromBefore } }] },
  });

  if (values?.length) {
    const startRow = startIndex + 1;
    const endRow = startRow + values.length - 1;
    const startCol = colIndexToLetter(anchorCol);
    const endCol = colIndexToLetter(anchorCol + Math.max(...values.map(r => r.length)) - 1);
    const updateRange = sheetName ? `'${sheetName}'!${startCol}${startRow}:${endCol}${endRow}` : `${startCol}${startRow}:${endCol}${endRow}`;
    await api.spreadsheets.values.update({ spreadsheetId: id, range: updateRange, valueInputOption: valueInputOption || 'USER_ENTERED', requestBody: { values } });
    return `Inserted ${rows} rows and filled ${values.reduce((s, r) => s + r.length, 0)} cells.`;
  }
  return `Inserted ${rows} row(s) ${position} row ${anchorRow + 1}.`;
}

export async function deleteRows(auth, { spreadsheetId, range }) {
  const id = extractId(spreadsheetId);
  const api = sheets(auth);
  const { sheetName, range: rowRange } = extractSheetName(range);
  const sheetId = await getSheetId(api, id, sheetName);

  const match = rowRange.match(/^(\d+):(\d+)$/);
  if (!match) throw new Error('Use full-row notation e.g. "Sheet1!2:4"');
  const startIndex = parseInt(match[1]) - 1;
  const endIndex = parseInt(match[2]);

  await api.spreadsheets.batchUpdate({
    spreadsheetId: id,
    requestBody: { requests: [{ deleteDimension: { range: { sheetId, dimension: 'ROWS', startIndex, endIndex } } }] },
  });
  return `Deleted rows ${match[1]}–${match[2]}.`;
}

export async function deleteColumns(auth, { spreadsheetId, range }) {
  const id = extractId(spreadsheetId);
  const api = sheets(auth);
  const { sheetName, range: colRange } = extractSheetName(range);
  const sheetId = await getSheetId(api, id, sheetName);

  const match = colRange.match(/^([A-Z]+):([A-Z]+)$/i);
  if (!match) throw new Error('Use full-column notation e.g. "Sheet1!B:D"');
  const startIndex = columnToIndex(match[1].toUpperCase());
  const endIndex = columnToIndex(match[2].toUpperCase()) + 1;

  await api.spreadsheets.batchUpdate({
    spreadsheetId: id,
    requestBody: { requests: [{ deleteDimension: { range: { sheetId, dimension: 'COLUMNS', startIndex, endIndex } } }] },
  });
  return `Deleted columns ${match[1]}–${match[2]}.`;
}

export async function findReplace(auth, { spreadsheetId, find, replacement, sheetName, matchCase = false, matchEntireCell = false, searchByRegex = false }) {
  const id = extractId(spreadsheetId);
  const api = sheets(auth);

  const findReplace = { find, replacement, matchCase, matchEntireCell, searchByRegex };
  if (sheetName) {
    const sheetId = await getSheetId(api, id, sheetName);
    findReplace.sheetId = sheetId;
  }

  const res = await api.spreadsheets.batchUpdate({
    spreadsheetId: id,
    requestBody: { requests: [{ findReplace }] },
  });
  const r = res.data.replies?.[0]?.findReplace;
  return `Replaced ${r?.occurrencesChanged || 0} occurrence(s) of "${find}" with "${replacement}".`;
}

export async function sortRange(auth, { spreadsheetId, range, sortSpecs }) {
  const id = extractId(spreadsheetId);
  const api = sheets(auth);
  const { sheetName, range: cleanRange } = extractSheetName(range);
  const sheetId = await getSheetId(api, id, sheetName);

  const colMatch = cleanRange.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/i);
  if (!colMatch) throw new Error('Range must be in A1:B10 format');

  const gridRange = {
    sheetId,
    startRowIndex: parseInt(colMatch[2]) - 1,
    endRowIndex: parseInt(colMatch[4]),
    startColumnIndex: columnToIndex(colMatch[1].toUpperCase()),
    endColumnIndex: columnToIndex(colMatch[3].toUpperCase()) + 1,
  };

  const specs = sortSpecs.map(s => ({
    dimensionIndex: typeof s.column === 'string' ? columnToIndex(s.column.toUpperCase()) : s.column,
    sortOrder: s.order === 'DESC' ? 'DESCENDING' : 'ASCENDING',
  }));

  await api.spreadsheets.batchUpdate({
    spreadsheetId: id,
    requestBody: { requests: [{ sortRange: { range: gridRange, sortSpecs: specs } }] },
  });
  return `Sorted ${range} by ${sortSpecs.length} column(s).`;
}

export async function exportSheet(auth, { spreadsheetId, format, sheetId: gid }) {
  const id = extractId(spreadsheetId);
  const mimeTypes = { csv: 'text/csv', tsv: 'text/tab-separated-values', xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', pdf: 'application/pdf', html: 'application/zip' };
  const mimeType = mimeTypes[format?.toLowerCase()];
  if (!mimeType) throw new Error(`Unsupported format. Use: ${Object.keys(mimeTypes).join(', ')}`);

  const params = { fileId: id, mimeType };
  if (gid !== undefined) params.params = { gid };

  const res = await drive(auth).files.export({ ...params, alt: 'media' }, { responseType: 'arraybuffer' });
  const bytes = Buffer.from(res.data);
  return `Exported ${bytes.length} bytes as ${format.toUpperCase()}. (Use download_drive_file to save locally.)`;
}

// ── Tool definitions ──────────────────────────────────────────────────────────

export const TOOLS = [
  { name: 'read_sheet', fn: readSheet, description: 'Read rows from a Google Sheet', inputSchema: { type: 'object', properties: { spreadsheetId: { type: 'string', description: 'Spreadsheet ID or URL' }, range: { type: 'string', description: 'A1 range, e.g. Sheet1!A1:Z100. Defaults to A1:Z1000.' } }, required: ['spreadsheetId'] } },
  { name: 'batch_read_sheet', fn: batchReadSheet, description: 'Read multiple ranges from a Google Sheet in one call', inputSchema: { type: 'object', properties: { spreadsheetId: { type: 'string' }, ranges: { type: 'array', items: { type: 'string' }, description: 'Array of A1 ranges' }, majorDimension: { type: 'string', enum: ['ROWS', 'COLUMNS'] }, valueRenderOption: { type: 'string', enum: ['FORMATTED_VALUE', 'UNFORMATTED_VALUE', 'FORMULA'] } }, required: ['spreadsheetId', 'ranges'] } },
  { name: 'append_sheet_row', fn: appendSheetRow, description: 'Append a new row after the last data row in a Google Sheet. Uses dataRange to scope the data area (recommended for sheets with instructions/reference content below the data).', inputSchema: { type: 'object', properties: { spreadsheetId: { type: 'string' }, sheetName: { type: 'string', description: 'Tab name (defaults to first sheet)' }, dataRange: { type: 'string', description: 'A1 range scoping the data area, e.g. "C8:G". Scans first column for last occupied row. Omit to auto-detect via frozen rows.' }, values: { type: 'array', items: { type: 'string' }, description: 'Cell values in column order' } }, required: ['spreadsheetId', 'values'] } },
  { name: 'update_sheet_cell', fn: updateSheetCell, description: 'Update a specific cell or range in a Google Sheet', inputSchema: { type: 'object', properties: { spreadsheetId: { type: 'string' }, range: { type: 'string', description: 'A1 range, e.g. Sheet1!B5' }, values: { type: 'array', description: '2D array of values', items: { type: 'array', items: {} } } }, required: ['spreadsheetId', 'range', 'values'] } },
  { name: 'batch_update_sheet', fn: batchUpdateSheet, description: 'Write to multiple ranges in a Google Sheet in one call', inputSchema: { type: 'object', properties: { spreadsheetId: { type: 'string' }, data: { type: 'array', items: { type: 'object', properties: { range: { type: 'string' }, values: { type: 'array', items: { type: 'array', items: {} } } }, required: ['range', 'values'] } }, valueInputOption: { type: 'string', enum: ['RAW', 'USER_ENTERED'] } }, required: ['spreadsheetId', 'data'] } },
  { name: 'clear_sheet', fn: clearSheet, description: 'Clear all values in a range of a Google Sheet', inputSchema: { type: 'object', properties: { spreadsheetId: { type: 'string' }, range: { type: 'string', description: 'A1 range to clear, e.g. Sheet1!A1:Z100' } }, required: ['spreadsheetId', 'range'] } },
  { name: 'insert_rows', fn: insertRows, description: 'Insert new rows at a specific position, with optional data', inputSchema: { type: 'object', properties: { spreadsheetId: { type: 'string' }, range: { type: 'string', description: 'Anchor cell, e.g. Sheet1!A5' }, rows: { type: 'number', description: 'Number of rows to insert (default 1)' }, position: { type: 'string', enum: ['BEFORE', 'AFTER'], description: 'Insert before or after anchor row (default BEFORE)' }, inheritFromBefore: { type: 'boolean' }, values: { type: 'array', items: { type: 'array', items: {} }, description: 'Optional values to fill the new rows' }, valueInputOption: { type: 'string', enum: ['RAW', 'USER_ENTERED'] } }, required: ['spreadsheetId', 'range'] } },
  { name: 'delete_rows', fn: deleteRows, description: 'Delete rows from a Google Sheet', inputSchema: { type: 'object', properties: { spreadsheetId: { type: 'string' }, range: { type: 'string', description: 'Full-row range, e.g. Sheet1!2:4' } }, required: ['spreadsheetId', 'range'] } },
  { name: 'delete_columns', fn: deleteColumns, description: 'Delete columns from a Google Sheet', inputSchema: { type: 'object', properties: { spreadsheetId: { type: 'string' }, range: { type: 'string', description: 'Full-column range, e.g. Sheet1!B:D' } }, required: ['spreadsheetId', 'range'] } },
  { name: 'find_replace', fn: findReplace, description: 'Find and replace text in a Google Sheet', inputSchema: { type: 'object', properties: { spreadsheetId: { type: 'string' }, find: { type: 'string' }, replacement: { type: 'string' }, sheetName: { type: 'string', description: 'Limit to one sheet (omit to search all sheets)' }, matchCase: { type: 'boolean' }, matchEntireCell: { type: 'boolean' }, searchByRegex: { type: 'boolean' } }, required: ['spreadsheetId', 'find', 'replacement'] } },
  { name: 'sort_range', fn: sortRange, description: 'Sort a range of cells in a Google Sheet', inputSchema: { type: 'object', properties: { spreadsheetId: { type: 'string' }, range: { type: 'string', description: 'Range to sort, e.g. Sheet1!A2:D100' }, sortSpecs: { type: 'array', items: { type: 'object', properties: { column: { description: 'Column letter (e.g. "A") or 0-based index' }, order: { type: 'string', enum: ['ASC', 'DESC'] } }, required: ['column'] }, description: 'Sort criteria in priority order' } }, required: ['spreadsheetId', 'range', 'sortSpecs'] } },
  { name: 'export_sheet', fn: exportSheet, description: 'Export a Google Sheet to csv, tsv, xlsx, pdf, or html', inputSchema: { type: 'object', properties: { spreadsheetId: { type: 'string' }, format: { type: 'string', enum: ['csv', 'tsv', 'xlsx', 'pdf', 'html'] }, sheetId: { type: 'number', description: 'Numeric sheet ID for CSV/TSV (exports first sheet if omitted)' } }, required: ['spreadsheetId', 'format'] } },
];
