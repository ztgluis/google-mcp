import { google } from 'googleapis';
import { extractId } from '../google.js';
import { extractSheetName, getSheetId, parseRange, columnToIndex } from '../utils/range-helpers.js';

function sheets(auth) { return google.sheets({ version: 'v4', auth }); }

export async function formatCells(auth, { spreadsheetId, range, format }) {
  const id = extractId(spreadsheetId);
  const api = sheets(auth);
  const { sheetName, range: cleanRange } = extractSheetName(range);
  const sheetId = await getSheetId(api, id, sheetName);
  const gridRange = parseRange(cleanRange, sheetId);

  const cellFormat = {};
  if (format.backgroundColor) cellFormat.backgroundColor = format.backgroundColor;
  if (format.textFormat) {
    cellFormat.textFormat = {};
    const tf = format.textFormat;
    if (tf.foregroundColor !== undefined) cellFormat.textFormat.foregroundColor = tf.foregroundColor;
    if (tf.fontFamily !== undefined) cellFormat.textFormat.fontFamily = tf.fontFamily;
    if (tf.fontSize !== undefined) cellFormat.textFormat.fontSize = tf.fontSize;
    if (tf.bold !== undefined) cellFormat.textFormat.bold = tf.bold;
    if (tf.italic !== undefined) cellFormat.textFormat.italic = tf.italic;
    if (tf.strikethrough !== undefined) cellFormat.textFormat.strikethrough = tf.strikethrough;
    if (tf.underline !== undefined) cellFormat.textFormat.underline = tf.underline;
  }
  if (format.horizontalAlignment) cellFormat.horizontalAlignment = format.horizontalAlignment;
  if (format.verticalAlignment) cellFormat.verticalAlignment = format.verticalAlignment;
  if (format.wrapStrategy) cellFormat.wrapStrategy = format.wrapStrategy;
  if (format.numberFormat) cellFormat.numberFormat = format.numberFormat;
  if (format.padding) cellFormat.padding = format.padding;

  await api.spreadsheets.batchUpdate({
    spreadsheetId: id,
    requestBody: { requests: [{ repeatCell: { range: gridRange, cell: { userEnteredFormat: cellFormat }, fields: 'userEnteredFormat' } }] },
  });
  return `Formatted cells in ${range}.`;
}

export async function mergeCells(auth, { spreadsheetId, range, mergeType = 'MERGE_ALL' }) {
  const id = extractId(spreadsheetId);
  const api = sheets(auth);
  const { sheetName, range: cleanRange } = extractSheetName(range);
  const sheetId = await getSheetId(api, id, sheetName);
  const gridRange = parseRange(cleanRange, sheetId);

  await api.spreadsheets.batchUpdate({
    spreadsheetId: id,
    requestBody: { requests: [{ mergeCells: { range: gridRange, mergeType } }] },
  });
  return `Merged cells in ${range} (${mergeType}).`;
}

export async function unmergeCells(auth, { spreadsheetId, range }) {
  const id = extractId(spreadsheetId);
  const api = sheets(auth);
  const { sheetName, range: cleanRange } = extractSheetName(range);
  const sheetId = await getSheetId(api, id, sheetName);
  const gridRange = parseRange(cleanRange, sheetId);

  await api.spreadsheets.batchUpdate({
    spreadsheetId: id,
    requestBody: { requests: [{ unmergeCells: { range: gridRange } }] },
  });
  return `Unmerged cells in ${range}.`;
}

export async function addConditionalFormatting(auth, { spreadsheetId, rules }) {
  const id = extractId(spreadsheetId);
  const api = sheets(auth);
  const requests = [];

  for (const rule of rules) {
    const gridRanges = [];
    for (const r of rule.ranges) {
      const { sheetName, range: cleanRange } = extractSheetName(r);
      const sheetId = await getSheetId(api, id, sheetName);
      gridRanges.push(parseRange(cleanRange, sheetId));
    }

    const cfRule = { ranges: gridRanges };

    if (rule.booleanRule) {
      const fmt = {};
      const f = rule.booleanRule.format;
      if (f.backgroundColor) fmt.backgroundColor = f.backgroundColor;
      if (f.textFormat) fmt.textFormat = f.textFormat;
      if (f.horizontalAlignment) fmt.horizontalAlignment = f.horizontalAlignment;
      if (f.numberFormat) fmt.numberFormat = f.numberFormat;
      cfRule.booleanRule = { condition: rule.booleanRule.condition, format: fmt };
    } else if (rule.gradientRule) {
      cfRule.gradientRule = rule.gradientRule;
    }

    requests.push({ addConditionalFormatRule: { rule: cfRule } });
  }

  await api.spreadsheets.batchUpdate({ spreadsheetId: id, requestBody: { requests } });
  return `Added ${rules.length} conditional formatting rule(s).`;
}

export async function freezeRowsColumns(auth, { spreadsheetId, sheetName, frozenRowCount, frozenColumnCount }) {
  const id = extractId(spreadsheetId);
  const api = sheets(auth);

  // Fetch full sheet data to get sheetId and check for merged cell conflicts
  const meta = await api.spreadsheets.get({
    spreadsheetId: id,
    fields: 'sheets.properties,sheets.merges',
  });
  const allSheets = meta.data.sheets || [];
  const target = sheetName
    ? allSheets.find(s => s.properties?.title === sheetName)
    : allSheets[0];
  if (!target) throw new Error(`Sheet "${sheetName || '(first)'}" not found.`);
  const sheetId = target.properties.sheetId;
  const merges = target.merges || [];

  // Check if freezing columns would split any merged cell
  const conflicts = [];
  if (frozenColumnCount !== undefined && frozenColumnCount > 0) {
    for (const m of merges) {
      if (m.startColumnIndex < frozenColumnCount && m.endColumnIndex > frozenColumnCount) {
        const startCol = String.fromCharCode(65 + m.startColumnIndex);
        const endCol = String.fromCharCode(65 + m.endColumnIndex - 1);
        conflicts.push(`${startCol}${m.startRowIndex + 1}:${endCol}${m.endRowIndex}`);
      }
    }
  }
  if (frozenRowCount !== undefined && frozenRowCount > 0) {
    for (const m of merges) {
      if (m.startRowIndex < frozenRowCount && m.endRowIndex > frozenRowCount) {
        const startCol = String.fromCharCode(65 + m.startColumnIndex);
        const endCol = String.fromCharCode(65 + m.endColumnIndex - 1);
        conflicts.push(`${startCol}${m.startRowIndex + 1}:${endCol}${m.endRowIndex}`);
      }
    }
  }

  if (conflicts.length) {
    return `Cannot freeze: merged cells cross the freeze boundary. Conflicting merges: ${conflicts.join(', ')}. Unmerge them first or adjust the freeze count to include the full merged range (e.g. freeze ${merges.reduce((max, m) => Math.max(max, m.endColumnIndex), 0)} columns instead).`;
  }

  const gridProperties = {};
  const fields = [];
  if (frozenRowCount !== undefined) { gridProperties.frozenRowCount = frozenRowCount; fields.push('gridProperties.frozenRowCount'); }
  if (frozenColumnCount !== undefined) { gridProperties.frozenColumnCount = frozenColumnCount; fields.push('gridProperties.frozenColumnCount'); }
  if (!fields.length) throw new Error('Provide frozenRowCount or frozenColumnCount.');

  await api.spreadsheets.batchUpdate({
    spreadsheetId: id,
    requestBody: { requests: [{ updateSheetProperties: { properties: { sheetId, gridProperties }, fields: fields.join(',') } }] },
  });
  return `Freeze updated on "${sheetName || 'first sheet'}": rows=${frozenRowCount ?? 'unchanged'}, cols=${frozenColumnCount ?? 'unchanged'}.`;
}

export async function autoResize(auth, { spreadsheetId, sheetName, dimension = 'COLUMNS', startIndex, endIndex }) {
  const id = extractId(spreadsheetId);
  const api = sheets(auth);
  const sheetId = await getSheetId(api, id, sheetName);

  const start = typeof startIndex === 'string' ? columnToIndex(startIndex.toUpperCase()) : (startIndex ?? 0);
  const end = typeof endIndex === 'string' ? columnToIndex(endIndex.toUpperCase()) + 1 : (endIndex ?? start + 1);

  await api.spreadsheets.batchUpdate({
    spreadsheetId: id,
    requestBody: { requests: [{ autoResizeDimensions: { dimensions: { sheetId, dimension: dimension.toUpperCase(), startIndex: start, endIndex: end } } }] },
  });
  return `Auto-resized ${dimension.toLowerCase()}s ${startIndex}–${endIndex} on "${sheetName || 'first sheet'}".`;
}

// ── Tool definitions ──────────────────────────────────────────────────────────

export const TOOLS = [
  { name: 'format_cells', fn: formatCells, description: 'Format cells in a Google Sheet (colors, fonts, alignment, number formats)', inputSchema: { type: 'object', properties: { spreadsheetId: { type: 'string' }, range: { type: 'string', description: 'A1 range, e.g. Sheet1!A1:C10' }, format: { type: 'object', properties: { backgroundColor: { type: 'object', description: 'RGBA object with red/green/blue/alpha keys (0.0–1.0)' }, textFormat: { type: 'object', properties: { bold: { type: 'boolean' }, italic: { type: 'boolean' }, underline: { type: 'boolean' }, strikethrough: { type: 'boolean' }, fontSize: { type: 'number' }, fontFamily: { type: 'string' }, foregroundColor: { type: 'object' } } }, horizontalAlignment: { type: 'string', enum: ['LEFT', 'CENTER', 'RIGHT'] }, verticalAlignment: { type: 'string', enum: ['TOP', 'MIDDLE', 'BOTTOM'] }, wrapStrategy: { type: 'string', enum: ['OVERFLOW_CELL', 'LEGACY_WRAP', 'CLIP', 'WRAP'] }, numberFormat: { type: 'object', properties: { type: { type: 'string', enum: ['TEXT', 'NUMBER', 'PERCENT', 'CURRENCY', 'DATE', 'TIME', 'DATE_TIME', 'SCIENTIFIC'] }, pattern: { type: 'string' } } }, padding: { type: 'object', properties: { top: { type: 'number' }, right: { type: 'number' }, bottom: { type: 'number' }, left: { type: 'number' } } } } } }, required: ['spreadsheetId', 'range', 'format'] } },
  { name: 'merge_cells', fn: mergeCells, description: 'Merge cells in a Google Sheet', inputSchema: { type: 'object', properties: { spreadsheetId: { type: 'string' }, range: { type: 'string', description: 'A1 range, e.g. Sheet1!A1:C3' }, mergeType: { type: 'string', enum: ['MERGE_ALL', 'MERGE_COLUMNS', 'MERGE_ROWS'], description: 'Default: MERGE_ALL' } }, required: ['spreadsheetId', 'range'] } },
  { name: 'unmerge_cells', fn: unmergeCells, description: 'Unmerge previously merged cells in a Google Sheet', inputSchema: { type: 'object', properties: { spreadsheetId: { type: 'string' }, range: { type: 'string' } }, required: ['spreadsheetId', 'range'] } },
  { name: 'add_conditional_formatting', fn: addConditionalFormatting, description: 'Add conditional formatting rules to a Google Sheet', inputSchema: { type: 'object', properties: { spreadsheetId: { type: 'string' }, rules: { type: 'array', description: 'Array of conditional format rules', items: { type: 'object', properties: { ranges: { type: 'array', items: { type: 'string' } }, booleanRule: { type: 'object', properties: { condition: { type: 'object', properties: { type: { type: 'string' }, values: { type: 'array', items: { type: 'object' } } }, required: ['type'] }, format: { type: 'object' } } }, gradientRule: { type: 'object' } }, required: ['ranges'] } } }, required: ['spreadsheetId', 'rules'] } },
  { name: 'freeze', fn: freezeRowsColumns, description: 'Freeze rows and/or columns in a Google Sheet', inputSchema: { type: 'object', properties: { spreadsheetId: { type: 'string' }, sheetName: { type: 'string', description: 'Tab name (defaults to first sheet)' }, frozenRowCount: { type: 'number', description: 'Number of rows to freeze (0 to unfreeze)' }, frozenColumnCount: { type: 'number', description: 'Number of columns to freeze (0 to unfreeze)' } }, required: ['spreadsheetId'] } },
  { name: 'auto_resize', fn: autoResize, description: 'Auto-resize columns or rows to fit their content', inputSchema: { type: 'object', properties: { spreadsheetId: { type: 'string' }, sheetName: { type: 'string' }, dimension: { type: 'string', enum: ['COLUMNS', 'ROWS'], description: 'Default: COLUMNS' }, startIndex: { description: 'Start column letter (e.g. "A") or 0-based row index' }, endIndex: { description: 'End column letter (e.g. "D") or 0-based row index (inclusive)' } }, required: ['spreadsheetId'] } },
];
