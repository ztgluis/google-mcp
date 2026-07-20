import { google } from 'googleapis';
import { extractId } from '../google.js';
import { extractSheetName, getSheetId, parseRange, columnToIndex } from '../utils/range-helpers.js';

function sheets(auth) { return google.sheets({ version: 'v4', auth }); }

export async function getMetadata(auth, { spreadsheetId }) {
  const id = extractId(spreadsheetId);
  const res = await sheets(auth).spreadsheets.get({ spreadsheetId: id, includeGridData: false });
  const d = res.data;
  const sheetList = (d.sheets || []).map(s => {
    const p = s.properties;
    return `  • ${p.title} (sheetId=${p.sheetId}, index=${p.index}, rows=${p.gridProperties?.rowCount}, cols=${p.gridProperties?.columnCount}, frozen rows=${p.gridProperties?.frozenRowCount || 0}, frozen cols=${p.gridProperties?.frozenColumnCount || 0})`;
  }).join('\n');
  return `Title: ${d.properties?.title}\nURL: ${d.spreadsheetUrl}\nSheets:\n${sheetList}`;
}

export async function createSpreadsheet(auth, { title, sheets: sheetDefs }) {
  const requestBody = { properties: { title } };
  if (sheetDefs?.length) {
    requestBody.sheets = sheetDefs.map((s, i) => ({
      properties: { title: s.title || `Sheet${i + 1}`, gridProperties: { rowCount: s.rowCount || 1000, columnCount: s.columnCount || 26 } },
    }));
  }
  const res = await sheets(auth).spreadsheets.create({ requestBody });
  return `Created spreadsheet "${res.data.properties?.title}".\nID: ${res.data.spreadsheetId}\nURL: ${res.data.spreadsheetUrl}`;
}

export async function insertSheet(auth, { spreadsheetId, title, index, rowCount, columnCount }) {
  const id = extractId(spreadsheetId);
  const res = await sheets(auth).spreadsheets.batchUpdate({
    spreadsheetId: id,
    requestBody: { requests: [{ addSheet: { properties: { title, index, gridProperties: { rowCount: rowCount || 1000, columnCount: columnCount || 26 } } } }] },
  });
  const added = res.data.replies?.[0]?.addSheet?.properties;
  return `Added sheet "${added?.title}" (sheetId=${added?.sheetId}, index=${added?.index}).`;
}

export async function deleteSheet(auth, { spreadsheetId, sheetId }) {
  const id = extractId(spreadsheetId);
  await sheets(auth).spreadsheets.batchUpdate({
    spreadsheetId: id,
    requestBody: { requests: [{ deleteSheet: { sheetId } }] },
  });
  return `Deleted sheet (sheetId=${sheetId}).`;
}

export async function renameSheet(auth, { spreadsheetId, sheetName, newName }) {
  const id = extractId(spreadsheetId);
  const api = sheets(auth);
  const sheetId = await getSheetId(api, id, sheetName);
  await api.spreadsheets.batchUpdate({
    spreadsheetId: id,
    requestBody: { requests: [{ updateSheetProperties: { properties: { sheetId, title: newName }, fields: 'title' } }] },
  });
  return `Renamed "${sheetName}" → "${newName}".`;
}

export async function duplicateSheet(auth, { spreadsheetId, sheetId, insertSheetIndex, newSheetName }) {
  const id = extractId(spreadsheetId);
  const res = await sheets(auth).spreadsheets.batchUpdate({
    spreadsheetId: id,
    requestBody: { requests: [{ duplicateSheet: { sourceSheetId: sheetId, insertSheetIndex, newSheetName } }] },
  });
  const dup = res.data.replies?.[0]?.duplicateSheet?.properties;
  return `Duplicated sheet → "${dup?.title}" (sheetId=${dup?.sheetId}, index=${dup?.index}).`;
}

export async function copySheetTo(auth, { spreadsheetId, sheetId, destinationSpreadsheetId }) {
  const id = extractId(spreadsheetId);
  const destId = extractId(destinationSpreadsheetId);
  const res = await sheets(auth).spreadsheets.sheets.copyTo({
    spreadsheetId: id, sheetId, requestBody: { destinationSpreadsheetId: destId },
  });
  return `Copied sheet to destination spreadsheet as "${res.data.title}" (sheetId=${res.data.sheetId}).`;
}

export async function updateSheetProperties(auth, { spreadsheetId, sheetId, title, gridProperties, tabColor }) {
  const id = extractId(spreadsheetId);
  const properties = { sheetId };
  const fields = [];
  if (title !== undefined) { properties.title = title; fields.push('title'); }
  if (gridProperties) {
    properties.gridProperties = gridProperties;
    ['rowCount', 'columnCount', 'frozenRowCount', 'frozenColumnCount'].forEach(k => { if (gridProperties[k] !== undefined) fields.push(`gridProperties.${k}`); });
  }
  if (tabColor) { properties.tabColor = tabColor; fields.push('tabColor'); }
  if (!fields.length) throw new Error('No properties to update.');

  await sheets(auth).spreadsheets.batchUpdate({
    spreadsheetId: id,
    requestBody: { requests: [{ updateSheetProperties: { properties, fields: fields.join(',') } }] },
  });
  return `Updated sheet properties: ${fields.join(', ')}.`;
}

export async function addProtectedRange(auth, { spreadsheetId, range, description, warningOnly = false }) {
  const id = extractId(spreadsheetId);
  const api = sheets(auth);
  const { sheetName, range: cleanRange } = extractSheetName(range);
  const sheetId = await getSheetId(api, id, sheetName);
  const gridRange = parseRange(cleanRange, sheetId);

  const protectedRange = { range: gridRange, warningOnly };
  if (description) protectedRange.description = description;

  const res = await api.spreadsheets.batchUpdate({
    spreadsheetId: id,
    requestBody: { requests: [{ addProtectedRange: { protectedRange } }] },
  });
  const prId = res.data.replies?.[0]?.addProtectedRange?.protectedRange?.protectedRangeId;
  return `Added protected range on ${range} (protectedRangeId=${prId}).`;
}

export async function updateProtectedRange(auth, { spreadsheetId, protectedRangeId, range, description, warningOnly }) {
  const id = extractId(spreadsheetId);
  const api = sheets(auth);
  const protectedRange = { protectedRangeId };
  const fields = [];

  if (range) {
    const { sheetName, range: cleanRange } = extractSheetName(range);
    const sheetId = await getSheetId(api, id, sheetName);
    protectedRange.range = parseRange(cleanRange, sheetId);
    fields.push('range');
  }
  if (description !== undefined) { protectedRange.description = description; fields.push('description'); }
  if (warningOnly !== undefined) { protectedRange.warningOnly = warningOnly; fields.push('warningOnly'); }

  await api.spreadsheets.batchUpdate({
    spreadsheetId: id,
    requestBody: { requests: [{ updateProtectedRange: { protectedRange, fields: fields.join(',') } }] },
  });
  return `Updated protected range (id=${protectedRangeId}).`;
}

export async function deleteProtectedRange(auth, { spreadsheetId, protectedRangeId }) {
  const id = extractId(spreadsheetId);
  await sheets(auth).spreadsheets.batchUpdate({
    spreadsheetId: id,
    requestBody: { requests: [{ deleteProtectedRange: { protectedRangeId } }] },
  });
  return `Deleted protected range (id=${protectedRangeId}).`;
}

export async function moveDimension(auth, { spreadsheetId, sheetName, dimension, sourceStart, sourceEnd, destinationIndex }) {
  const id = extractId(spreadsheetId);
  const api = sheets(auth);
  const sheetId = await getSheetId(api, id, sheetName);

  await api.spreadsheets.batchUpdate({
    spreadsheetId: id,
    requestBody: { requests: [{ moveDimension: { source: { sheetId, dimension, startIndex: sourceStart, endIndex: sourceEnd }, destinationIndex } }] },
  });
  return `Moved ${dimension.toLowerCase()} ${sourceStart}–${sourceEnd} to index ${destinationIndex}.`;
}

export async function addDimensionGroup(auth, { spreadsheetId, sheetName, dimension, startIndex, endIndex }) {
  const id = extractId(spreadsheetId);
  const api = sheets(auth);
  const sheetId = await getSheetId(api, id, sheetName);

  await api.spreadsheets.batchUpdate({
    spreadsheetId: id,
    requestBody: { requests: [{ addDimensionGroup: { range: { sheetId, dimension, startIndex, endIndex } } }] },
  });
  return `Added ${dimension.toLowerCase()} group ${startIndex}–${endIndex}.`;
}

export async function deleteDimensionGroup(auth, { spreadsheetId, sheetName, dimension, startIndex, endIndex }) {
  const id = extractId(spreadsheetId);
  const api = sheets(auth);
  const sheetId = await getSheetId(api, id, sheetName);

  await api.spreadsheets.batchUpdate({
    spreadsheetId: id,
    requestBody: { requests: [{ deleteDimensionGroup: { range: { sheetId, dimension, startIndex, endIndex } } }] },
  });
  return `Deleted ${dimension.toLowerCase()} group ${startIndex}–${endIndex}.`;
}

export async function textToColumns(auth, { spreadsheetId, range, delimiter, delimiterType = 'AUTODETECT' }) {
  const id = extractId(spreadsheetId);
  const api = sheets(auth);
  const { sheetName, range: cleanRange } = extractSheetName(range);
  const sheetId = await getSheetId(api, id, sheetName);
  const gridRange = parseRange(cleanRange, sheetId);

  const req = { source: gridRange, delimiterType };
  if (delimiterType === 'CUSTOM' && delimiter) req.delimiter = delimiter;

  await api.spreadsheets.batchUpdate({
    spreadsheetId: id,
    requestBody: { requests: [{ textToColumns: req }] },
  });
  return `Split text to columns in ${range} using ${delimiterType}${delimiterType === 'CUSTOM' ? ` ("${delimiter}")` : ''}.`;
}

export async function listSheetObjects(auth, { spreadsheetId, sheetName }) {
  const id = extractId(spreadsheetId);
  const api = sheets(auth);

  const res = await api.spreadsheets.get({
    spreadsheetId: id,
    fields: 'sheets(properties(sheetId,title),charts,conditionalFormats,filterViews,basicFilter,bandedRanges,protectedRanges),namedRanges',
  });

  const targetSheets = sheetName
    ? (res.data.sheets || []).filter(s => s.properties?.title === sheetName)
    : res.data.sheets || [];

  if (!targetSheets.length) return sheetName ? `Sheet "${sheetName}" not found.` : 'No sheets found.';

  const lines = [];

  for (const sheet of targetSheets) {
    lines.push(`=== ${sheet.properties.title} (sheetId=${sheet.properties.sheetId}) ===`);

    // Charts
    const charts = sheet.charts || [];
    if (charts.length) {
      lines.push(`\nCharts (${charts.length}):`);
      for (const chart of charts) {
        const spec = chart.spec || {};
        const title = spec.title || '(untitled)';
        lines.push(`  chartId=${chart.chartId} | "${title}"`);
        lines.push(`    spec: ${JSON.stringify(chart.spec, null, 2).split('\n').join('\n    ')}`);
        if (chart.position) lines.push(`    position: ${JSON.stringify(chart.position)}`);
      }
    }

    // Conditional formats
    const cfs = sheet.conditionalFormats || [];
    if (cfs.length) {
      lines.push(`\nConditional formats (${cfs.length}):`);
      for (let i = 0; i < cfs.length; i++) {
        const cf = cfs[i];
        const ranges = (cf.ranges || []).map(r => {
          const parts = [];
          if (r.startRowIndex !== undefined) parts.push(`rows ${r.startRowIndex}-${r.endRowIndex}`);
          if (r.startColumnIndex !== undefined) parts.push(`cols ${r.startColumnIndex}-${r.endColumnIndex}`);
          return parts.join(', ');
        }).join('; ');
        const type = cf.booleanRule ? `boolean (${cf.booleanRule.condition?.type || '?'})` : cf.gradientRule ? 'gradient' : '?';
        lines.push(`  [${i}] ${type} | ranges: ${ranges}`);
      }
    }

    // Filter views
    const fvs = sheet.filterViews || [];
    if (fvs.length) {
      lines.push(`\nFilter views (${fvs.length}):`);
      for (const fv of fvs) {
        lines.push(`  filterViewId=${fv.filterViewId} | "${fv.title || '(untitled)'}"`);
        if (fv.criteria) lines.push(`    criteria: ${JSON.stringify(fv.criteria)}`);
        if (fv.sortSpecs) lines.push(`    sortSpecs: ${JSON.stringify(fv.sortSpecs)}`);
        const r = fv.range || {};
        if (r.startRowIndex !== undefined) lines.push(`    range: rows ${r.startRowIndex}-${r.endRowIndex}, cols ${r.startColumnIndex}-${r.endColumnIndex}`);
      }
    }

    // Basic filter
    if (sheet.basicFilter) {
      const bf = sheet.basicFilter;
      const r = bf.range || {};
      lines.push(`\nBasic filter: rows ${r.startRowIndex}-${r.endRowIndex}, cols ${r.startColumnIndex}-${r.endColumnIndex}`);
    }

    // Banded ranges
    const brs = sheet.bandedRanges || [];
    if (brs.length) {
      lines.push(`\nBanded ranges (${brs.length}):`);
      for (const br of brs) {
        lines.push(`  bandedRangeId=${br.bandedRangeId}`);
      }
    }

    // Protected ranges
    const prs = sheet.protectedRanges || [];
    if (prs.length) {
      lines.push(`\nProtected ranges (${prs.length}):`);
      for (const pr of prs) {
        const desc = pr.description ? ` "${pr.description}"` : '';
        const warn = pr.warningOnly ? ' [warning only]' : ' [locked]';
        lines.push(`  protectedRangeId=${pr.protectedRangeId}${desc}${warn}`);
      }
    }

    if (!charts.length && !cfs.length && !fvs.length && !sheet.basicFilter && !brs.length && !prs.length) {
      lines.push('  (no objects)');
    }
  }

  const namedRanges = res.data.namedRanges || [];
  if (namedRanges.length) {
    lines.push(`\n=== Named Ranges (${namedRanges.length}) ===`);
    for (const nr of namedRanges) {
      const r = nr.range || {};
      const rangeStr = r.startRowIndex !== undefined
        ? `sheetId=${r.sheetId}, rows ${r.startRowIndex}-${r.endRowIndex}, cols ${r.startColumnIndex}-${r.endColumnIndex}`
        : JSON.stringify(r);
      lines.push(`  "${nr.name}" (namedRangeId=${nr.namedRangeId}): ${rangeStr}`);
    }
  }

  return lines.join('\n');
}

export async function addSheetNamedRange(auth, { spreadsheetId, name, range }) {
  const id = extractId(spreadsheetId);
  const api = sheets(auth);
  const { sheetName, range: cleanRange } = extractSheetName(range);
  const sheetId = await getSheetId(api, id, sheetName);
  const gridRange = parseRange(cleanRange, sheetId);

  const res = await api.spreadsheets.batchUpdate({
    spreadsheetId: id,
    requestBody: { requests: [{ addNamedRange: { namedRange: { name, range: gridRange } } }] },
  });
  const nrId = res.data.replies?.[0]?.addNamedRange?.namedRange?.namedRangeId;
  return `Created named range "${name}" (namedRangeId=${nrId}).`;
}

export async function updateSheetNamedRange(auth, { spreadsheetId, namedRangeId, name, range }) {
  const id = extractId(spreadsheetId);
  const api = sheets(auth);

  const namedRange = { namedRangeId };
  const fields = [];
  if (name !== undefined) { namedRange.name = name; fields.push('name'); }
  if (range !== undefined) {
    const { sheetName, range: cleanRange } = extractSheetName(range);
    const sheetId = await getSheetId(api, id, sheetName);
    namedRange.range = parseRange(cleanRange, sheetId);
    fields.push('range');
  }

  await api.spreadsheets.batchUpdate({
    spreadsheetId: id,
    requestBody: { requests: [{ updateNamedRange: { namedRange, fields: fields.join(',') } }] },
  });
  return `Updated named range (namedRangeId=${namedRangeId}).`;
}

export async function deleteSheetNamedRange(auth, { spreadsheetId, namedRangeId }) {
  const id = extractId(spreadsheetId);
  await sheets(auth).spreadsheets.batchUpdate({
    spreadsheetId: id,
    requestBody: { requests: [{ deleteNamedRange: { namedRangeId } }] },
  });
  return `Deleted named range (namedRangeId=${namedRangeId}).`;
}

export async function updateFilterView(auth, { spreadsheetId, filterViewId, title, range, sortSpecs, criteria }) {
  const id = extractId(spreadsheetId);
  const api = sheets(auth);
  const filter = { filterViewId };
  const fields = [];

  if (title !== undefined) { filter.title = title; fields.push('title'); }
  if (range !== undefined) {
    const { sheetName, range: cleanRange } = extractSheetName(range);
    const sheetId = await getSheetId(api, id, sheetName);
    filter.range = parseRange(cleanRange, sheetId);
    fields.push('range');
  }
  if (sortSpecs !== undefined) { filter.sortSpecs = sortSpecs; fields.push('sortSpecs'); }
  if (criteria !== undefined) { filter.criteria = criteria; fields.push('criteria'); }

  await api.spreadsheets.batchUpdate({
    spreadsheetId: id,
    requestBody: { requests: [{ updateFilterView: { filter, fields: { paths: fields } } }] },
  });
  return `Updated filter view (filterViewId=${filterViewId}).`;
}

export async function deleteFilterView(auth, { spreadsheetId, filterViewId }) {
  const id = extractId(spreadsheetId);
  await sheets(auth).spreadsheets.batchUpdate({
    spreadsheetId: id,
    requestBody: { requests: [{ deleteFilterView: { filterViewId } }] },
  });
  return `Deleted filter view (filterViewId=${filterViewId}).`;
}

export async function addSlicer(auth, { spreadsheetId, range, filterColumnIndex, title, position }) {
  const id = extractId(spreadsheetId);
  const api = sheets(auth);
  const { sheetName, range: cleanRange } = extractSheetName(range);
  const sheetId = await getSheetId(api, id, sheetName);
  const gridRange = parseRange(cleanRange, sheetId);

  const slicer = {
    spec: { dataRange: gridRange, filterCriteria: {}, columnIndex: filterColumnIndex, title },
    position: {
      overlayPosition: {
        anchorCell: { sheetId, rowIndex: position?.rowIndex || 0, columnIndex: position?.columnIndex || 0 },
      },
    },
  };

  const res = await api.spreadsheets.batchUpdate({
    spreadsheetId: id,
    requestBody: { requests: [{ addSlicer: { slicer } }] },
  });
  const slicerId = res.data.replies?.[0]?.addSlicer?.slicer?.slicerId;
  return `Added slicer "${title || ''}" (slicerId=${slicerId}).`;
}

// ── Tool definitions ──────────────────────────────────────────────────────────

export const TOOLS = [
  { name: 'get_sheet_metadata', fn: getMetadata, description: 'Get metadata about a spreadsheet: title, URL, all sheet names, IDs, dimensions, and frozen rows/cols', inputSchema: { type: 'object', properties: { spreadsheetId: { type: 'string', description: 'Spreadsheet ID or URL' } }, required: ['spreadsheetId'] } },
  { name: 'create_spreadsheet', fn: createSpreadsheet, description: 'Create a new Google Sheets spreadsheet', inputSchema: { type: 'object', properties: { title: { type: 'string' }, sheets: { type: 'array', items: { type: 'object', properties: { title: { type: 'string' }, rowCount: { type: 'number' }, columnCount: { type: 'number' } } }, description: 'Optional sheet tabs to create (defaults to one sheet)' } }, required: ['title'] } },
  { name: 'insert_sheet', fn: insertSheet, description: 'Add a new tab to an existing spreadsheet', inputSchema: { type: 'object', properties: { spreadsheetId: { type: 'string' }, title: { type: 'string', description: 'Name for the new tab' }, index: { type: 'number', description: '0-based position (appends at end if omitted)' }, rowCount: { type: 'number' }, columnCount: { type: 'number' } }, required: ['spreadsheetId', 'title'] } },
  { name: 'delete_sheet', fn: deleteSheet, description: 'Delete a tab from a spreadsheet by its numeric sheetId (use get_sheet_metadata to find sheetIds)', inputSchema: { type: 'object', properties: { spreadsheetId: { type: 'string' }, sheetId: { type: 'number', description: 'Numeric sheet ID (not the tab name)' } }, required: ['spreadsheetId', 'sheetId'] } },
  { name: 'rename_sheet', fn: renameSheet, description: 'Rename a tab in a spreadsheet', inputSchema: { type: 'object', properties: { spreadsheetId: { type: 'string' }, sheetName: { type: 'string', description: 'Current tab name' }, newName: { type: 'string', description: 'New tab name' } }, required: ['spreadsheetId', 'sheetName', 'newName'] } },
  { name: 'duplicate_sheet', fn: duplicateSheet, description: 'Duplicate a tab within the same spreadsheet', inputSchema: { type: 'object', properties: { spreadsheetId: { type: 'string' }, sheetId: { type: 'number', description: 'Numeric sheet ID of the tab to duplicate' }, insertSheetIndex: { type: 'number', description: '0-based index where the copy is inserted' }, newSheetName: { type: 'string', description: 'Name for the duplicated tab' } }, required: ['spreadsheetId', 'sheetId'] } },
  { name: 'copy_sheet_to', fn: copySheetTo, description: 'Copy a tab to a different spreadsheet', inputSchema: { type: 'object', properties: { spreadsheetId: { type: 'string', description: 'Source spreadsheet ID or URL' }, sheetId: { type: 'number', description: 'Numeric sheet ID of the tab to copy' }, destinationSpreadsheetId: { type: 'string', description: 'Destination spreadsheet ID or URL' } }, required: ['spreadsheetId', 'sheetId', 'destinationSpreadsheetId'] } },
  { name: 'update_sheet_properties', fn: updateSheetProperties, description: 'Update a sheet tab properties: title, grid size, frozen rows/cols, tab color', inputSchema: { type: 'object', properties: { spreadsheetId: { type: 'string' }, sheetId: { type: 'number', description: 'Numeric sheet ID' }, title: { type: 'string' }, gridProperties: { type: 'object', properties: { rowCount: { type: 'number' }, columnCount: { type: 'number' }, frozenRowCount: { type: 'number' }, frozenColumnCount: { type: 'number' } } }, tabColor: { type: 'object', properties: { red: { type: 'number' }, green: { type: 'number' }, blue: { type: 'number' } }, description: 'RGB values 0.0–1.0' } }, required: ['spreadsheetId', 'sheetId'] } },
  { name: 'add_protected_range', fn: addProtectedRange, description: 'Protect a range from editing (or show a warning)', inputSchema: { type: 'object', properties: { spreadsheetId: { type: 'string' }, range: { type: 'string', description: 'A1 range, e.g. Sheet1!A1:D10' }, description: { type: 'string', description: 'Description of why the range is protected' }, warningOnly: { type: 'boolean', description: 'If true, show a warning instead of blocking edits (default false)' } }, required: ['spreadsheetId', 'range'] } },
  { name: 'update_protected_range', fn: updateProtectedRange, description: 'Update an existing protected range', inputSchema: { type: 'object', properties: { spreadsheetId: { type: 'string' }, protectedRangeId: { type: 'number', description: 'Protected range ID' }, range: { type: 'string', description: 'New A1 range' }, description: { type: 'string' }, warningOnly: { type: 'boolean' } }, required: ['spreadsheetId', 'protectedRangeId'] } },
  { name: 'delete_protected_range', fn: deleteProtectedRange, description: 'Remove protection from a range', inputSchema: { type: 'object', properties: { spreadsheetId: { type: 'string' }, protectedRangeId: { type: 'number', description: 'Protected range ID' } }, required: ['spreadsheetId', 'protectedRangeId'] } },
  { name: 'move_dimension', fn: moveDimension, description: 'Move rows or columns from one position to another', inputSchema: { type: 'object', properties: { spreadsheetId: { type: 'string' }, sheetName: { type: 'string', description: 'Tab name' }, dimension: { type: 'string', enum: ['ROWS', 'COLUMNS'] }, sourceStart: { type: 'number', description: '0-based start index of source range' }, sourceEnd: { type: 'number', description: '0-based end index (exclusive) of source range' }, destinationIndex: { type: 'number', description: '0-based destination index' } }, required: ['spreadsheetId', 'dimension', 'sourceStart', 'sourceEnd', 'destinationIndex'] } },
  { name: 'add_dimension_group', fn: addDimensionGroup, description: 'Group rows or columns (creates a collapsible outline)', inputSchema: { type: 'object', properties: { spreadsheetId: { type: 'string' }, sheetName: { type: 'string', description: 'Tab name' }, dimension: { type: 'string', enum: ['ROWS', 'COLUMNS'] }, startIndex: { type: 'number', description: '0-based start index' }, endIndex: { type: 'number', description: '0-based end index (exclusive)' } }, required: ['spreadsheetId', 'dimension', 'startIndex', 'endIndex'] } },
  { name: 'delete_dimension_group', fn: deleteDimensionGroup, description: 'Remove a row or column group', inputSchema: { type: 'object', properties: { spreadsheetId: { type: 'string' }, sheetName: { type: 'string', description: 'Tab name' }, dimension: { type: 'string', enum: ['ROWS', 'COLUMNS'] }, startIndex: { type: 'number', description: '0-based start index' }, endIndex: { type: 'number', description: '0-based end index (exclusive)' } }, required: ['spreadsheetId', 'dimension', 'startIndex', 'endIndex'] } },
  { name: 'list_sheet_objects', fn: listSheetObjects, description: 'List all objects in a spreadsheet: charts (with full spec), conditional formats, filter views (with criteria), basic filters, banded ranges, protected ranges, and named ranges. Optionally filter to a specific sheet tab.', inputSchema: { type: 'object', properties: { spreadsheetId: { type: 'string' }, sheetName: { type: 'string', description: 'Optional tab name to filter to' } }, required: ['spreadsheetId'] } },
  { name: 'text_to_columns', fn: textToColumns, description: 'Split a column of text into multiple columns by delimiter', inputSchema: { type: 'object', properties: { spreadsheetId: { type: 'string' }, range: { type: 'string', description: 'A1 range of the source column, e.g. Sheet1!A1:A100' }, delimiterType: { type: 'string', enum: ['COMMA', 'SEMICOLON', 'PERIOD', 'SPACE', 'CUSTOM', 'AUTODETECT'], description: 'Default: AUTODETECT' }, delimiter: { type: 'string', description: 'Custom delimiter character (only used when delimiterType is CUSTOM)' } }, required: ['spreadsheetId', 'range'] } },
  { name: 'add_sheet_named_range', fn: addSheetNamedRange, description: 'Create a named range in a spreadsheet (e.g. "SalesData" → Sheet1!A1:D100)', inputSchema: { type: 'object', properties: { spreadsheetId: { type: 'string' }, name: { type: 'string', description: 'Name for the range (e.g. SalesData)' }, range: { type: 'string', description: 'A1 range, e.g. Sheet1!A1:D100' } }, required: ['spreadsheetId', 'name', 'range'] } },
  { name: 'update_sheet_named_range', fn: updateSheetNamedRange, description: 'Update a named range (rename or change its range reference)', inputSchema: { type: 'object', properties: { spreadsheetId: { type: 'string' }, namedRangeId: { type: 'string', description: 'Named range ID (from list_sheet_objects)' }, name: { type: 'string', description: 'New name' }, range: { type: 'string', description: 'New A1 range' } }, required: ['spreadsheetId', 'namedRangeId'] } },
  { name: 'delete_sheet_named_range', fn: deleteSheetNamedRange, description: 'Delete a named range from a spreadsheet', inputSchema: { type: 'object', properties: { spreadsheetId: { type: 'string' }, namedRangeId: { type: 'string', description: 'Named range ID' } }, required: ['spreadsheetId', 'namedRangeId'] } },
  { name: 'update_filter_view', fn: updateFilterView, description: 'Update a filter view (title, range, sort specs, filter criteria)', inputSchema: { type: 'object', properties: { spreadsheetId: { type: 'string' }, filterViewId: { type: 'number', description: 'Filter view ID (from list_sheet_objects)' }, title: { type: 'string' }, range: { type: 'string', description: 'New A1 range' }, sortSpecs: { type: 'array', items: { type: 'object' }, description: 'Sort specifications' }, criteria: { type: 'object', description: 'Filter criteria keyed by column index' } }, required: ['spreadsheetId', 'filterViewId'] } },
  { name: 'delete_filter_view', fn: deleteFilterView, description: 'Delete a filter view', inputSchema: { type: 'object', properties: { spreadsheetId: { type: 'string' }, filterViewId: { type: 'number', description: 'Filter view ID' } }, required: ['spreadsheetId', 'filterViewId'] } },
  { name: 'add_slicer', fn: addSlicer, description: 'Add an interactive slicer (filter control) to a sheet', inputSchema: { type: 'object', properties: { spreadsheetId: { type: 'string' }, range: { type: 'string', description: 'Data range the slicer filters, e.g. Sheet1!A1:E100' }, filterColumnIndex: { type: 'number', description: '0-based column index within the range to filter on' }, title: { type: 'string', description: 'Slicer title' }, position: { type: 'object', properties: { rowIndex: { type: 'number' }, columnIndex: { type: 'number' } }, description: 'Anchor cell position' } }, required: ['spreadsheetId', 'range', 'filterColumnIndex'] } },
];
