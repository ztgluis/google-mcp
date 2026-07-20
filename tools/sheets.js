import { google } from 'googleapis';
import { extractId } from '../google.js';
import { extractSheetName, getSheetId, parseRange, columnToIndex, colIndexToLetter } from '../utils/range-helpers.js';

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

export async function setDataValidation(auth, { spreadsheetId, range, type, values, strict = true, showCustomUi = true, inputMessage }) {
  const id = extractId(spreadsheetId);
  const api = sheets(auth);
  const { sheetName, range: cleanRange } = extractSheetName(range);
  const sheetId = await getSheetId(api, id, sheetName);
  const gridRange = parseRange(cleanRange, sheetId);

  const condition = { type };
  if (values?.length) {
    condition.values = values.map(v => ({ userEnteredValue: String(v) }));
  }

  const dataValidation = { condition, strict, showCustomUi };
  if (inputMessage) dataValidation.inputMessage = inputMessage;

  await api.spreadsheets.batchUpdate({
    spreadsheetId: id,
    requestBody: { requests: [{ repeatCell: { range: gridRange, cell: { dataValidation }, fields: 'dataValidation' } }] },
  });
  return `Set data validation (${type}) on ${range}.`;
}

export async function batchClearSheet(auth, { spreadsheetId, ranges }) {
  const id = extractId(spreadsheetId);
  const res = await sheets(auth).spreadsheets.values.batchClear({
    spreadsheetId: id,
    requestBody: { ranges },
  });
  const cleared = res.data.clearedRanges || ranges;
  return `Cleared ${cleared.length} range(s): ${cleared.join(', ')}.`;
}

export async function addChart(auth, { spreadsheetId, sheetName, chartType, title, sourceRange, position }) {
  const id = extractId(spreadsheetId);
  const api = sheets(auth);
  const { sheetName: srcSheet, range: cleanRange } = extractSheetName(sourceRange);
  const srcSheetId = await getSheetId(api, id, srcSheet || sheetName);
  const gridRange = parseRange(cleanRange, srcSheetId);

  const anchorSheetId = position?.sheetName
    ? await getSheetId(api, id, position.sheetName)
    : srcSheetId;

  const chart = {
    spec: {
      title,
      basicChart: {
        chartType,
        legendPosition: 'BOTTOM_LEGEND',
        domains: [{ domain: { sourceRange: { sources: [{ sheetId: gridRange.sheetId, startRowIndex: gridRange.startRowIndex, endRowIndex: gridRange.endRowIndex, startColumnIndex: gridRange.startColumnIndex, endColumnIndex: gridRange.startColumnIndex + 1 }] } } }],
        series: [{ series: { sourceRange: { sources: [{ sheetId: gridRange.sheetId, startRowIndex: gridRange.startRowIndex, endRowIndex: gridRange.endRowIndex, startColumnIndex: gridRange.startColumnIndex + 1, endColumnIndex: gridRange.endColumnIndex }] } } }],
      },
    },
    position: {
      overlayPosition: {
        anchorCell: {
          sheetId: anchorSheetId,
          rowIndex: position?.rowIndex || 0,
          columnIndex: position?.columnIndex || 0,
        },
      },
    },
  };

  const res = await api.spreadsheets.batchUpdate({
    spreadsheetId: id,
    requestBody: { requests: [{ addChart: { chart } }] },
  });
  const chartId = res.data.replies?.[0]?.addChart?.chart?.chartId;
  return `Added ${chartType} chart "${title}" (chartId=${chartId}).`;
}

export async function deleteEmbeddedObject(auth, { spreadsheetId, objectId }) {
  const id = extractId(spreadsheetId);
  await sheets(auth).spreadsheets.batchUpdate({
    spreadsheetId: id,
    requestBody: { requests: [{ deleteEmbeddedObject: { objectId } }] },
  });
  return `Deleted embedded object (id=${objectId}).`;
}

export async function setBasicFilter(auth, { spreadsheetId, range }) {
  const id = extractId(spreadsheetId);
  const api = sheets(auth);
  const { sheetName, range: cleanRange } = extractSheetName(range);
  const sheetId = await getSheetId(api, id, sheetName);
  const gridRange = parseRange(cleanRange, sheetId);

  await api.spreadsheets.batchUpdate({
    spreadsheetId: id,
    requestBody: { requests: [{ setBasicFilter: { filter: { range: gridRange } } }] },
  });
  return `Set basic filter on ${range}.`;
}

export async function clearBasicFilter(auth, { spreadsheetId, sheetName }) {
  const id = extractId(spreadsheetId);
  const api = sheets(auth);
  const sheetId = await getSheetId(api, id, sheetName);

  await api.spreadsheets.batchUpdate({
    spreadsheetId: id,
    requestBody: { requests: [{ clearBasicFilter: { sheetId } }] },
  });
  return `Cleared basic filter on "${sheetName || 'first sheet'}".`;
}

export async function addFilterView(auth, { spreadsheetId, range, title }) {
  const id = extractId(spreadsheetId);
  const api = sheets(auth);
  const { sheetName, range: cleanRange } = extractSheetName(range);
  const sheetId = await getSheetId(api, id, sheetName);
  const gridRange = parseRange(cleanRange, sheetId);

  const res = await api.spreadsheets.batchUpdate({
    spreadsheetId: id,
    requestBody: { requests: [{ addFilterView: { filter: { title, range: gridRange } } }] },
  });
  const fvId = res.data.replies?.[0]?.addFilterView?.filter?.filterViewId;
  return `Added filter view "${title}" (filterViewId=${fvId}).`;
}

export async function setCellNote(auth, { spreadsheetId, range, note }) {
  const id = extractId(spreadsheetId);
  const api = sheets(auth);
  const { sheetName, range: cleanRange } = extractSheetName(range);
  const sheetId = await getSheetId(api, id, sheetName);
  const gridRange = parseRange(cleanRange, sheetId);
  await api.spreadsheets.batchUpdate({
    spreadsheetId: id,
    requestBody: { requests: [{ repeatCell: { range: gridRange, cell: { note: note || '' }, fields: 'note' } }] },
  });
  return note ? `Set note on ${range}.` : `Cleared note on ${range}.`;
}

export async function setCellRichText(auth, { spreadsheetId, range, text, runs }) {
  const id = extractId(spreadsheetId);
  const api = sheets(auth);
  const { sheetName, range: cleanRange } = extractSheetName(range);
  const sheetId = await getSheetId(api, id, sheetName);
  const gridRange = parseRange(cleanRange, sheetId);

  const cellData = { userEnteredValue: { stringValue: text } };
  if (runs?.length) {
    cellData.textFormatRuns = runs.map(r => {
      const run = { startIndex: r.startIndex };
      if (r.format) run.format = r.format;
      return run;
    });
  }

  await api.spreadsheets.batchUpdate({
    spreadsheetId: id,
    requestBody: { requests: [{ updateCells: { rows: [{ values: [cellData] }], range: gridRange, fields: 'userEnteredValue,textFormatRuns' } }] },
  });
  return `Set rich text on ${range}.`;
}

export async function updateChart(auth, { spreadsheetId, chartId, spec }) {
  const id = extractId(spreadsheetId);
  await sheets(auth).spreadsheets.batchUpdate({
    spreadsheetId: id,
    requestBody: { requests: [{ updateChartSpec: { chartId, spec } }] },
  });
  return `Updated chart (chartId=${chartId}).`;
}

export async function copyPaste(auth, { spreadsheetId, source, destination, pasteType = 'PASTE_NORMAL', pasteOrientation = 'NORMAL' }) {
  const id = extractId(spreadsheetId);
  const api = sheets(auth);
  const src = extractSheetName(source);
  const srcSheetId = await getSheetId(api, id, src.sheetName);
  const srcRange = parseRange(src.range, srcSheetId);
  const dst = extractSheetName(destination);
  const dstSheetId = await getSheetId(api, id, dst.sheetName);
  const dstRange = parseRange(dst.range, dstSheetId);

  await api.spreadsheets.batchUpdate({
    spreadsheetId: id,
    requestBody: { requests: [{ copyPaste: { source: srcRange, destination: dstRange, pasteType, pasteOrientation } }] },
  });
  return `Copied ${source} to ${destination} (${pasteType}).`;
}

export async function cutPaste(auth, { spreadsheetId, source, destination, pasteType = 'PASTE_NORMAL' }) {
  const id = extractId(spreadsheetId);
  const api = sheets(auth);
  const src = extractSheetName(source);
  const srcSheetId = await getSheetId(api, id, src.sheetName);
  const srcRange = parseRange(src.range, srcSheetId);
  const dst = extractSheetName(destination);
  const dstSheetId = await getSheetId(api, id, dst.sheetName);
  const dstCoord = parseRange(dst.range, dstSheetId);

  await api.spreadsheets.batchUpdate({
    spreadsheetId: id,
    requestBody: { requests: [{ cutPaste: { source: srcRange, destination: { sheetId: dstCoord.sheetId, rowIndex: dstCoord.startRowIndex, columnIndex: dstCoord.startColumnIndex }, pasteType } }] },
  });
  return `Cut ${source} and pasted to ${destination} (${pasteType}).`;
}

export async function autoFill(auth, { spreadsheetId, source, destination, useAlternateSeries = false }) {
  const id = extractId(spreadsheetId);
  const api = sheets(auth);
  const src = extractSheetName(source);
  const srcSheetId = await getSheetId(api, id, src.sheetName);
  const srcRange = parseRange(src.range, srcSheetId);
  const dst = extractSheetName(destination);
  const dstRange = parseRange(dst.range, srcSheetId);

  await api.spreadsheets.batchUpdate({
    spreadsheetId: id,
    requestBody: { requests: [{ autoFill: { useAlternateSeries, range: dstRange, sourceAndDestination: { source: srcRange, dimension: 'ROWS', fillLength: (dstRange.endRowIndex || 0) - (srcRange.endRowIndex || 0) } } }] },
  });
  return `Auto-filled from ${source} to ${destination}.`;
}

export async function trimWhitespace(auth, { spreadsheetId, range }) {
  const id = extractId(spreadsheetId);
  const api = sheets(auth);
  const { sheetName, range: cleanRange } = extractSheetName(range);
  const sheetId = await getSheetId(api, id, sheetName);
  const gridRange = parseRange(cleanRange, sheetId);

  await api.spreadsheets.batchUpdate({
    spreadsheetId: id,
    requestBody: { requests: [{ trimWhitespace: { range: gridRange } }] },
  });
  return `Trimmed whitespace in ${range}.`;
}

export async function clearDataValidation(auth, { spreadsheetId, range }) {
  const id = extractId(spreadsheetId);
  const api = sheets(auth);
  const { sheetName, range: cleanRange } = extractSheetName(range);
  const sheetId = await getSheetId(api, id, sheetName);
  const gridRange = parseRange(cleanRange, sheetId);

  await api.spreadsheets.batchUpdate({
    spreadsheetId: id,
    requestBody: { requests: [{ repeatCell: { range: gridRange, cell: { dataValidation: null }, fields: 'dataValidation' } }] },
  });
  return `Cleared data validation on ${range}.`;
}

export async function updateSpreadsheetProperties(auth, { spreadsheetId, title, locale, autoRecalc, timeZone }) {
  const id = extractId(spreadsheetId);
  const properties = {};
  const fields = [];
  if (title !== undefined) { properties.title = title; fields.push('title'); }
  if (locale !== undefined) { properties.locale = locale; fields.push('locale'); }
  if (autoRecalc !== undefined) { properties.autoRecalc = autoRecalc; fields.push('autoRecalc'); }
  if (timeZone !== undefined) { properties.timeZone = timeZone; fields.push('timeZone'); }

  if (!fields.length) return 'No properties to update.';

  await sheets(auth).spreadsheets.batchUpdate({
    spreadsheetId: id,
    requestBody: { requests: [{ updateSpreadsheetProperties: { properties, fields: fields.join(',') } }] },
  });
  return `Updated spreadsheet properties: ${fields.join(', ')}.`;
}

export async function appendDimension(auth, { spreadsheetId, sheetName, dimension = 'ROWS', length = 1 }) {
  const id = extractId(spreadsheetId);
  const api = sheets(auth);
  const sheetId = await getSheetId(api, id, sheetName);

  await api.spreadsheets.batchUpdate({
    spreadsheetId: id,
    requestBody: { requests: [{ appendDimension: { sheetId, dimension, length } }] },
  });
  return `Appended ${length} ${dimension.toLowerCase()} to "${sheetName || 'first sheet'}".`;
}

export async function setDimensionSize(auth, { spreadsheetId, sheetName, dimension = 'ROWS', startIndex, endIndex, pixelSize }) {
  const id = extractId(spreadsheetId);
  const api = sheets(auth);
  const sheetId = await getSheetId(api, id, sheetName);

  if (dimension === 'COLUMNS' && typeof startIndex === 'string') {
    startIndex = columnToIndex(startIndex);
    if (typeof endIndex === 'string') endIndex = columnToIndex(endIndex) + 1;
    else if (endIndex === undefined) endIndex = startIndex + 1;
  } else {
    if (endIndex === undefined) endIndex = startIndex + 1;
  }

  await api.spreadsheets.batchUpdate({
    spreadsheetId: id,
    requestBody: { requests: [{ updateDimensionProperties: { range: { sheetId, dimension, startIndex, endIndex }, properties: { pixelSize }, fields: 'pixelSize' } }] },
  });
  return `Set ${dimension.toLowerCase()} ${startIndex}-${endIndex - 1} to ${pixelSize}px.`;
}

export async function addPivotTable(auth, { spreadsheetId, sourceRange, targetCell, rows, columns, values, filterCriteria }) {
  const id = extractId(spreadsheetId);
  const api = sheets(auth);
  const src = extractSheetName(sourceRange);
  const srcSheetId = await getSheetId(api, id, src.sheetName);
  const srcRange = parseRange(src.range, srcSheetId);

  const tgt = extractSheetName(targetCell);
  const tgtSheetId = await getSheetId(api, id, tgt.sheetName);
  const tgtRange = parseRange(tgt.range, tgtSheetId);

  const pivotTable = {
    source: srcRange,
    rows: (rows || []).map(r => ({
      sourceColumnOffset: r.sourceColumnOffset,
      showTotals: r.showTotals !== false,
      sortOrder: r.sortOrder || 'ASCENDING',
      ...(r.label ? { label: r.label } : {}),
    })),
    columns: (columns || []).map(c => ({
      sourceColumnOffset: c.sourceColumnOffset,
      showTotals: c.showTotals !== false,
      sortOrder: c.sortOrder || 'ASCENDING',
      ...(c.label ? { label: c.label } : {}),
    })),
    values: (values || []).map(v => ({
      sourceColumnOffset: v.sourceColumnOffset,
      summarizeFunction: v.summarizeFunction || 'SUM',
      ...(v.name ? { name: v.name } : {}),
    })),
    ...(filterCriteria ? { criteria: filterCriteria } : {}),
  };

  await api.spreadsheets.batchUpdate({
    spreadsheetId: id,
    requestBody: { requests: [{ updateCells: { rows: [{ values: [{ pivotTable }] }], start: { sheetId: tgtRange.sheetId, rowIndex: tgtRange.startRowIndex, columnIndex: tgtRange.startColumnIndex }, fields: 'pivotTable' } }] },
  });
  return `Created pivot table at ${targetCell} from ${sourceRange}.`;
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
  { name: 'set_data_validation', fn: setDataValidation, description: 'Set data validation rules on a range (dropdowns, number constraints, custom formulas)', inputSchema: { type: 'object', properties: { spreadsheetId: { type: 'string' }, range: { type: 'string', description: 'A1 range, e.g. Sheet1!B2:B100' }, type: { type: 'string', enum: ['ONE_OF_LIST', 'ONE_OF_RANGE', 'NUMBER_BETWEEN', 'NUMBER_GREATER', 'TEXT_CONTAINS', 'TEXT_NOT_CONTAINS', 'CUSTOM_FORMULA', 'BOOLEAN'], description: 'Validation type' }, values: { type: 'array', items: { type: 'string' }, description: 'Condition values (e.g. list items, range ref, min/max numbers, formula)' }, strict: { type: 'boolean', description: 'Reject invalid input (default true)' }, showCustomUi: { type: 'boolean', description: 'Show dropdown for list validations (default true)' }, inputMessage: { type: 'string', description: 'Help text shown when editing the cell' } }, required: ['spreadsheetId', 'range', 'type'] } },
  { name: 'batch_clear_sheet', fn: batchClearSheet, description: 'Clear values from multiple ranges in one call', inputSchema: { type: 'object', properties: { spreadsheetId: { type: 'string' }, ranges: { type: 'array', items: { type: 'string' }, description: 'Array of A1 ranges to clear' } }, required: ['spreadsheetId', 'ranges'] } },
  { name: 'add_chart', fn: addChart, description: 'Add an embedded chart to a Google Sheet', inputSchema: { type: 'object', properties: { spreadsheetId: { type: 'string' }, sheetName: { type: 'string', description: 'Fallback sheet name if sourceRange has no sheet prefix' }, chartType: { type: 'string', enum: ['BAR', 'LINE', 'PIE', 'COLUMN', 'AREA', 'SCATTER', 'COMBO'] }, title: { type: 'string', description: 'Chart title' }, sourceRange: { type: 'string', description: 'A1 range for chart data, e.g. Sheet1!A1:C10. First column is domain, remaining are series.' }, position: { type: 'object', properties: { sheetName: { type: 'string', description: 'Sheet to place chart on (defaults to source sheet)' }, rowIndex: { type: 'number', description: '0-based row for anchor cell' }, columnIndex: { type: 'number', description: '0-based column for anchor cell' } }, description: 'Where to place the chart overlay' } }, required: ['spreadsheetId', 'chartType', 'title', 'sourceRange'] } },
  { name: 'delete_embedded_object', fn: deleteEmbeddedObject, description: 'Delete an embedded object (chart, image) by its ID', inputSchema: { type: 'object', properties: { spreadsheetId: { type: 'string' }, objectId: { type: 'number', description: 'Embedded object ID (chartId)' } }, required: ['spreadsheetId', 'objectId'] } },
  { name: 'set_basic_filter', fn: setBasicFilter, description: 'Set a basic filter (auto-filter) on a range', inputSchema: { type: 'object', properties: { spreadsheetId: { type: 'string' }, range: { type: 'string', description: 'A1 range, e.g. Sheet1!A1:D100' } }, required: ['spreadsheetId', 'range'] } },
  { name: 'clear_basic_filter', fn: clearBasicFilter, description: 'Clear the basic filter from a sheet', inputSchema: { type: 'object', properties: { spreadsheetId: { type: 'string' }, sheetName: { type: 'string', description: 'Tab name (defaults to first sheet)' } }, required: ['spreadsheetId'] } },
  { name: 'add_filter_view', fn: addFilterView, description: 'Add a named filter view to a range', inputSchema: { type: 'object', properties: { spreadsheetId: { type: 'string' }, range: { type: 'string', description: 'A1 range, e.g. Sheet1!A1:D100' }, title: { type: 'string', description: 'Name for the filter view' } }, required: ['spreadsheetId', 'range', 'title'] } },
  { name: 'set_cell_note', fn: setCellNote, description: 'Set, update, or clear a note on a cell or range', inputSchema: { type: 'object', properties: { spreadsheetId: { type: 'string' }, range: { type: 'string', description: 'A1 range, e.g. Sheet1!B5' }, note: { type: 'string', description: 'Note text (empty string to clear)' } }, required: ['spreadsheetId', 'range'] } },
  { name: 'set_cell_rich_text', fn: setCellRichText, description: 'Set rich text with per-character formatting in a cell (e.g. part bold, part colored)', inputSchema: { type: 'object', properties: { spreadsheetId: { type: 'string' }, range: { type: 'string', description: 'Single cell A1 reference, e.g. Sheet1!B5' }, text: { type: 'string', description: 'The full cell text' }, runs: { type: 'array', description: 'Array of format runs, each with startIndex and format', items: { type: 'object', properties: { startIndex: { type: 'number', description: '0-based character index where this format starts' }, format: { type: 'object', description: 'Text format: bold, italic, underline, strikethrough, fontSize, fontFamily, foregroundColor, link' } }, required: ['startIndex'] } } }, required: ['spreadsheetId', 'range', 'text', 'runs'] } },
  { name: 'update_chart', fn: updateChart, description: 'Update an existing chart spec (title, chart type, data ranges, series, axes, colors)', inputSchema: { type: 'object', properties: { spreadsheetId: { type: 'string' }, chartId: { type: 'number', description: 'Chart ID (from list_sheet_objects)' }, spec: { type: 'object', description: 'Full chart spec object with basicChart, title, etc. Use list_sheet_objects to read the current spec first.' } }, required: ['spreadsheetId', 'chartId', 'spec'] } },
  { name: 'copy_paste', fn: copyPaste, description: 'Copy a range to another location with formatting options (paste values only, formatting only, etc.)', inputSchema: { type: 'object', properties: { spreadsheetId: { type: 'string' }, source: { type: 'string', description: 'Source A1 range, e.g. Sheet1!A1:D10' }, destination: { type: 'string', description: 'Destination A1 range, e.g. Sheet1!F1:I10' }, pasteType: { type: 'string', enum: ['PASTE_NORMAL', 'PASTE_VALUES', 'PASTE_FORMAT', 'PASTE_NO_BORDERS', 'PASTE_FORMULA', 'PASTE_DATA_VALIDATION', 'PASTE_CONDITIONAL_FORMATTING'], description: 'Default: PASTE_NORMAL (everything)' }, pasteOrientation: { type: 'string', enum: ['NORMAL', 'TRANSPOSE'], description: 'Default: NORMAL' } }, required: ['spreadsheetId', 'source', 'destination'] } },
  { name: 'cut_paste', fn: cutPaste, description: 'Cut a range and paste to a new location (moves data and formatting)', inputSchema: { type: 'object', properties: { spreadsheetId: { type: 'string' }, source: { type: 'string', description: 'Source A1 range to cut, e.g. Sheet1!A1:D10' }, destination: { type: 'string', description: 'Destination anchor cell, e.g. Sheet1!F1' }, pasteType: { type: 'string', enum: ['PASTE_NORMAL', 'PASTE_VALUES', 'PASTE_FORMAT', 'PASTE_NO_BORDERS', 'PASTE_FORMULA', 'PASTE_DATA_VALIDATION', 'PASTE_CONDITIONAL_FORMATTING'], description: 'Default: PASTE_NORMAL' } }, required: ['spreadsheetId', 'source', 'destination'] } },
  { name: 'auto_fill', fn: autoFill, description: 'Auto-fill a range by extending a pattern or formula (like dragging the fill handle)', inputSchema: { type: 'object', properties: { spreadsheetId: { type: 'string' }, source: { type: 'string', description: 'Source range containing the pattern, e.g. Sheet1!A1:A3' }, destination: { type: 'string', description: 'Full range to fill into (must include source), e.g. Sheet1!A1:A20' }, useAlternateSeries: { type: 'boolean', description: 'Use alternate series (e.g. 1,3,5 instead of 1,2,3). Default: false' } }, required: ['spreadsheetId', 'source', 'destination'] } },
  { name: 'trim_whitespace', fn: trimWhitespace, description: 'Trim leading and trailing whitespace from all cells in a range', inputSchema: { type: 'object', properties: { spreadsheetId: { type: 'string' }, range: { type: 'string', description: 'A1 range, e.g. Sheet1!A1:D100' } }, required: ['spreadsheetId', 'range'] } },
  { name: 'clear_data_validation', fn: clearDataValidation, description: 'Clear data validation rules from a range', inputSchema: { type: 'object', properties: { spreadsheetId: { type: 'string' }, range: { type: 'string', description: 'A1 range, e.g. Sheet1!B2:B100' } }, required: ['spreadsheetId', 'range'] } },
  { name: 'update_spreadsheet_properties', fn: updateSpreadsheetProperties, description: 'Update spreadsheet-level properties (title, locale, timezone, recalculation)', inputSchema: { type: 'object', properties: { spreadsheetId: { type: 'string' }, title: { type: 'string', description: 'New spreadsheet title' }, locale: { type: 'string', description: 'Locale (e.g. en_US)' }, autoRecalc: { type: 'string', enum: ['ON_CHANGE', 'MINUTE', 'HOUR'], description: 'Recalculation interval' }, timeZone: { type: 'string', description: 'Timezone (e.g. America/New_York)' } }, required: ['spreadsheetId'] } },
  { name: 'append_dimension', fn: appendDimension, description: 'Append empty rows or columns at the end of a sheet', inputSchema: { type: 'object', properties: { spreadsheetId: { type: 'string' }, sheetName: { type: 'string', description: 'Tab name (defaults to first sheet)' }, dimension: { type: 'string', enum: ['ROWS', 'COLUMNS'], description: 'Default: ROWS' }, length: { type: 'number', description: 'Number of rows/columns to add (default: 1)' } }, required: ['spreadsheetId'] } },
  { name: 'set_dimension_size', fn: setDimensionSize, description: 'Set specific row height or column width in pixels', inputSchema: { type: 'object', properties: { spreadsheetId: { type: 'string' }, sheetName: { type: 'string', description: 'Tab name (defaults to first sheet)' }, dimension: { type: 'string', enum: ['ROWS', 'COLUMNS'], description: 'Default: ROWS' }, startIndex: { description: '0-based row index or column letter (e.g. "A")' }, endIndex: { description: '0-based end index (exclusive) or column letter. Defaults to startIndex + 1 (single row/col).' }, pixelSize: { type: 'number', description: 'Size in pixels' } }, required: ['spreadsheetId', 'startIndex', 'pixelSize'] } },
  { name: 'add_pivot_table', fn: addPivotTable, description: 'Create a pivot table from a data range', inputSchema: { type: 'object', properties: { spreadsheetId: { type: 'string' }, sourceRange: { type: 'string', description: 'A1 range of the source data, e.g. Sheet1!A1:E100' }, targetCell: { type: 'string', description: 'Cell where the pivot table is placed, e.g. Sheet2!A1' }, rows: { type: 'array', description: 'Row grouping fields', items: { type: 'object', properties: { sourceColumnOffset: { type: 'number', description: '0-based column offset in the source range' }, showTotals: { type: 'boolean' }, sortOrder: { type: 'string', enum: ['ASCENDING', 'DESCENDING'] }, label: { type: 'string' } }, required: ['sourceColumnOffset'] } }, columns: { type: 'array', description: 'Column grouping fields', items: { type: 'object', properties: { sourceColumnOffset: { type: 'number' }, showTotals: { type: 'boolean' }, sortOrder: { type: 'string', enum: ['ASCENDING', 'DESCENDING'] }, label: { type: 'string' } }, required: ['sourceColumnOffset'] } }, values: { type: 'array', description: 'Value aggregation fields', items: { type: 'object', properties: { sourceColumnOffset: { type: 'number' }, summarizeFunction: { type: 'string', enum: ['SUM', 'COUNTA', 'COUNT', 'COUNTUNIQUE', 'AVERAGE', 'MAX', 'MIN', 'MEDIAN', 'PRODUCT', 'STDEV', 'STDEVP', 'VAR', 'VARP', 'CUSTOM'] }, name: { type: 'string', description: 'Display name for this value' } }, required: ['sourceColumnOffset'] } }, filterCriteria: { type: 'object', description: 'Filter criteria keyed by source column offset' } }, required: ['spreadsheetId', 'sourceRange', 'targetCell', 'values'] } },
];
