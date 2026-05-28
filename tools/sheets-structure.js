import { google } from 'googleapis';
import { extractId } from '../google.js';
import { getSheetId } from '../utils/range-helpers.js';

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
];
