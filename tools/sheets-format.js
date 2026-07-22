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
  const fieldPaths = [];

  if (format.backgroundColor) {
    cellFormat.backgroundColor = format.backgroundColor;
    fieldPaths.push('userEnteredFormat.backgroundColor');
  }
  if (format.textFormat) {
    cellFormat.textFormat = {};
    const tf = format.textFormat;
    if (tf.foregroundColor !== undefined) { cellFormat.textFormat.foregroundColor = tf.foregroundColor; fieldPaths.push('userEnteredFormat.textFormat.foregroundColor'); }
    if (tf.fontFamily !== undefined) { cellFormat.textFormat.fontFamily = tf.fontFamily; fieldPaths.push('userEnteredFormat.textFormat.fontFamily'); }
    if (tf.fontSize !== undefined) { cellFormat.textFormat.fontSize = tf.fontSize; fieldPaths.push('userEnteredFormat.textFormat.fontSize'); }
    if (tf.bold !== undefined) { cellFormat.textFormat.bold = tf.bold; fieldPaths.push('userEnteredFormat.textFormat.bold'); }
    if (tf.italic !== undefined) { cellFormat.textFormat.italic = tf.italic; fieldPaths.push('userEnteredFormat.textFormat.italic'); }
    if (tf.strikethrough !== undefined) { cellFormat.textFormat.strikethrough = tf.strikethrough; fieldPaths.push('userEnteredFormat.textFormat.strikethrough'); }
    if (tf.underline !== undefined) { cellFormat.textFormat.underline = tf.underline; fieldPaths.push('userEnteredFormat.textFormat.underline'); }
  }
  if (format.horizontalAlignment) { cellFormat.horizontalAlignment = format.horizontalAlignment; fieldPaths.push('userEnteredFormat.horizontalAlignment'); }
  if (format.verticalAlignment) { cellFormat.verticalAlignment = format.verticalAlignment; fieldPaths.push('userEnteredFormat.verticalAlignment'); }
  if (format.wrapStrategy) { cellFormat.wrapStrategy = format.wrapStrategy; fieldPaths.push('userEnteredFormat.wrapStrategy'); }
  if (format.numberFormat) { cellFormat.numberFormat = format.numberFormat; fieldPaths.push('userEnteredFormat.numberFormat'); }
  if (format.padding) { cellFormat.padding = format.padding; fieldPaths.push('userEnteredFormat.padding'); }

  await api.spreadsheets.batchUpdate({
    spreadsheetId: id,
    requestBody: { requests: [{ repeatCell: { range: gridRange, cell: { userEnteredFormat: cellFormat }, fields: fieldPaths.join(',') } }] },
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

export async function updateBorders(auth, { spreadsheetId, range, top, bottom, left, right, innerHorizontal, innerVertical }) {
  const id = extractId(spreadsheetId);
  const api = sheets(auth);
  const { sheetName, range: cleanRange } = extractSheetName(range);
  const sheetId = await getSheetId(api, id, sheetName);
  const gridRange = parseRange(cleanRange, sheetId);

  function makeBorder(b) {
    if (!b) return undefined;
    const border = {};
    if (b.style) border.style = b.style;
    if (b.width !== undefined) border.width = b.width;
    if (b.color) border.color = b.color;
    return border;
  }

  const updateBordersReq = { range: gridRange };
  if (top) updateBordersReq.top = makeBorder(top);
  if (bottom) updateBordersReq.bottom = makeBorder(bottom);
  if (left) updateBordersReq.left = makeBorder(left);
  if (right) updateBordersReq.right = makeBorder(right);
  if (innerHorizontal) updateBordersReq.innerHorizontal = makeBorder(innerHorizontal);
  if (innerVertical) updateBordersReq.innerVertical = makeBorder(innerVertical);

  await api.spreadsheets.batchUpdate({
    spreadsheetId: id,
    requestBody: { requests: [{ updateBorders: updateBordersReq }] },
  });
  return `Updated borders on ${range}.`;
}

export async function addBanding(auth, { spreadsheetId, range, headerColor, firstBandColor, secondBandColor, footerColor }) {
  const id = extractId(spreadsheetId);
  const api = sheets(auth);
  const { sheetName, range: cleanRange } = extractSheetName(range);
  const sheetId = await getSheetId(api, id, sheetName);
  const gridRange = parseRange(cleanRange, sheetId);

  const rowProperties = {};
  if (headerColor) rowProperties.headerColor = headerColor;
  if (firstBandColor) rowProperties.firstBandColor = firstBandColor;
  if (secondBandColor) rowProperties.secondBandColor = secondBandColor;
  if (footerColor) rowProperties.footerColor = footerColor;

  await api.spreadsheets.batchUpdate({
    spreadsheetId: id,
    requestBody: { requests: [{ addBanding: { bandedRange: { range: gridRange, rowProperties } } }] },
  });
  return `Added banding to ${range}.`;
}

export async function deleteBanding(auth, { spreadsheetId, bandedRangeId }) {
  const id = extractId(spreadsheetId);
  await sheets(auth).spreadsheets.batchUpdate({
    spreadsheetId: id,
    requestBody: { requests: [{ deleteBanding: { bandedRangeId } }] },
  });
  return `Deleted banded range (id=${bandedRangeId}).`;
}

export async function updateBanding(auth, { spreadsheetId, bandedRangeId, headerColor, firstBandColor, secondBandColor, footerColor }) {
  const id = extractId(spreadsheetId);
  const rowProperties = {};
  const fields = [];
  if (headerColor) { rowProperties.headerColor = headerColor; fields.push('rowProperties.headerColor'); }
  if (firstBandColor) { rowProperties.firstBandColor = firstBandColor; fields.push('rowProperties.firstBandColor'); }
  if (secondBandColor) { rowProperties.secondBandColor = secondBandColor; fields.push('rowProperties.secondBandColor'); }
  if (footerColor) { rowProperties.footerColor = footerColor; fields.push('rowProperties.footerColor'); }

  await sheets(auth).spreadsheets.batchUpdate({
    spreadsheetId: id,
    requestBody: { requests: [{ updateBanding: { bandedRange: { bandedRangeId, rowProperties }, fields: fields.join(',') } }] },
  });
  return `Updated banded range (id=${bandedRangeId}).`;
}

function formatGridRange(r) {
  const parts = [];
  if (r.startColumnIndex !== undefined && r.endColumnIndex !== undefined) {
    const sc = String.fromCharCode(65 + r.startColumnIndex);
    const ec = String.fromCharCode(65 + r.endColumnIndex - 1);
    parts.push(`${sc}${(r.startRowIndex || 0) + 1}:${ec}${r.endRowIndex || ''}`);
  } else {
    if (r.startRowIndex !== undefined) parts.push(`rows ${r.startRowIndex + 1}-${r.endRowIndex}`);
    if (r.startColumnIndex !== undefined) parts.push(`cols ${r.startColumnIndex}-${r.endColumnIndex}`);
  }
  return parts.join(', ') || '(whole sheet)';
}

function describeCellFormat(fmt) {
  if (!fmt) return '(none)';
  const parts = [];
  if (fmt.backgroundColor) {
    const c = fmt.backgroundColor;
    parts.push(`bg(${Math.round((c.red || 0) * 255)},${Math.round((c.green || 0) * 255)},${Math.round((c.blue || 0) * 255)})`);
  }
  if (fmt.textFormat) {
    const tf = fmt.textFormat;
    if (tf.bold) parts.push('bold');
    if (tf.italic) parts.push('italic');
    if (tf.underline) parts.push('underline');
    if (tf.strikethrough) parts.push('strikethrough');
    if (tf.foregroundColor) {
      const c = tf.foregroundColor;
      parts.push(`color(${Math.round((c.red || 0) * 255)},${Math.round((c.green || 0) * 255)},${Math.round((c.blue || 0) * 255)})`);
    }
    if (tf.fontSize) parts.push(`${tf.fontSize}pt`);
    if (tf.fontFamily) parts.push(tf.fontFamily);
  }
  return parts.join(', ') || '(default)';
}

function describeInterpolationPoint(p) {
  if (!p) return '?';
  const parts = [p.type || '?'];
  if (p.value) parts.push(`val=${p.value}`);
  if (p.color) {
    const c = p.color;
    parts.push(`rgb(${Math.round((c.red || 0) * 255)},${Math.round((c.green || 0) * 255)},${Math.round((c.blue || 0) * 255)})`);
  }
  return parts.join(' ');
}

export async function readConditionalFormats(auth, { spreadsheetId, sheetName }) {
  const id = extractId(spreadsheetId);
  const api = sheets(auth);

  const res = await api.spreadsheets.get({
    spreadsheetId: id,
    fields: 'sheets(properties(sheetId,title),conditionalFormats)',
  });

  const targetSheets = sheetName
    ? (res.data.sheets || []).filter(s => s.properties?.title === sheetName)
    : res.data.sheets || [];

  if (!targetSheets.length) return sheetName ? `Sheet "${sheetName}" not found.` : 'No sheets found.';

  const lines = [];

  for (const sheet of targetSheets) {
    const cfs = sheet.conditionalFormats || [];
    if (!cfs.length) {
      lines.push(`${sheet.properties.title}: no conditional formats`);
      continue;
    }

    lines.push(`=== ${sheet.properties.title} (${cfs.length} rules) ===`);

    for (let i = 0; i < cfs.length; i++) {
      const cf = cfs[i];
      const ranges = (cf.ranges || []).map(formatGridRange).join('; ');
      lines.push(`\n[${i}] Ranges: ${ranges}`);

      if (cf.booleanRule) {
        const br = cf.booleanRule;
        const cond = br.condition || {};
        const condValues = (cond.values || []).map(v => v.userEnteredValue || v.relativeDate || '').join(', ');
        lines.push(`  Type: boolean`);
        lines.push(`  Condition: ${cond.type || '?'}${condValues ? ' → ' + condValues : ''}`);
        lines.push(`  Format: ${describeCellFormat(br.format)}`);
      } else if (cf.gradientRule) {
        const gr = cf.gradientRule;
        lines.push(`  Type: gradient`);
        lines.push(`  Min: ${describeInterpolationPoint(gr.minpoint)}`);
        if (gr.midpoint) lines.push(`  Mid: ${describeInterpolationPoint(gr.midpoint)}`);
        lines.push(`  Max: ${describeInterpolationPoint(gr.maxpoint)}`);
      }
    }
  }

  return lines.join('\n');
}

export async function updateConditionalFormatting(auth, { spreadsheetId, sheetName, index, rule }) {
  const id = extractId(spreadsheetId);
  const api = sheets(auth);
  const sId = await getSheetId(api, id, sheetName);

  const newRule = {};

  if (rule.ranges) {
    newRule.ranges = rule.ranges.map(r => {
      if (typeof r === 'string') {
        const { range: cleanRange } = extractSheetName(r);
        return parseRange(cleanRange, sId);
      }
      return { ...r, sheetId: sId };
    });
  }

  if (rule.booleanRule) {
    newRule.booleanRule = rule.booleanRule;
  }
  if (rule.gradientRule) {
    newRule.gradientRule = rule.gradientRule;
  }

  const fields = Object.keys(newRule).join(',');

  await api.spreadsheets.batchUpdate({
    spreadsheetId: id,
    requestBody: {
      requests: [{
        updateConditionalFormatRule: {
          index,
          rule: newRule,
          sheetId: sId,
          fields,
        },
      }],
    },
  });
  return `Updated conditional format rule at index ${index}.`;
}

export async function deleteConditionalFormatting(auth, { spreadsheetId, sheetName, index }) {
  const id = extractId(spreadsheetId);
  const api = sheets(auth);
  const sId = await getSheetId(api, id, sheetName);

  await api.spreadsheets.batchUpdate({
    spreadsheetId: id,
    requestBody: {
      requests: [{
        deleteConditionalFormatRule: { index, sheetId: sId },
      }],
    },
  });
  return `Deleted conditional format rule at index ${index}.`;
}

function colLetter(idx) {
  let s = '';
  let i = idx;
  while (i >= 0) { s = String.fromCharCode(65 + (i % 26)) + s; i = Math.floor(i / 26) - 1; }
  return s;
}

function formatColor(c) {
  if (!c) return null;
  const parts = [];
  if (c.red !== undefined) parts.push(`r=${Math.round(c.red * 255)}`);
  if (c.green !== undefined) parts.push(`g=${Math.round(c.green * 255)}`);
  if (c.blue !== undefined) parts.push(`b=${Math.round(c.blue * 255)}`);
  if (c.alpha !== undefined) parts.push(`a=${c.alpha}`);
  return parts.join(',');
}

function formatBorder(b) {
  if (!b) return null;
  const parts = [];
  if (b.style) parts.push(b.style);
  if (b.width) parts.push(`w=${b.width}`);
  if (b.color) parts.push(formatColor(b.color));
  return parts.join(' ');
}

function formatCellDetails(fmt, label) {
  if (!fmt) return [];
  const parts = [];
  if (fmt.numberFormat) parts.push(`numFmt=${fmt.numberFormat.type}${fmt.numberFormat.pattern ? '(' + fmt.numberFormat.pattern + ')' : ''}`);
  if (fmt.backgroundColor) parts.push(`bg=${formatColor(fmt.backgroundColor)}`);
  if (fmt.backgroundColorStyle?.rgbColor) parts.push(`bgStyle=${formatColor(fmt.backgroundColorStyle.rgbColor)}`);
  if (fmt.textFormat) {
    const tf = fmt.textFormat;
    const tParts = [];
    if (tf.bold) tParts.push('bold');
    if (tf.italic) tParts.push('italic');
    if (tf.underline) tParts.push('underline');
    if (tf.strikethrough) tParts.push('strikethrough');
    if (tf.fontSize) tParts.push(`${tf.fontSize}pt`);
    if (tf.fontFamily) tParts.push(tf.fontFamily);
    if (tf.foregroundColor) tParts.push(`color=${formatColor(tf.foregroundColor)}`);
    if (tParts.length) parts.push(`text=[${tParts.join(', ')}]`);
  }
  if (fmt.horizontalAlignment) parts.push(`hAlign=${fmt.horizontalAlignment}`);
  if (fmt.verticalAlignment) parts.push(`vAlign=${fmt.verticalAlignment}`);
  if (fmt.wrapStrategy) parts.push(`wrap=${fmt.wrapStrategy}`);
  if (fmt.textRotation) {
    const tr = fmt.textRotation;
    parts.push(`rotation=${tr.angle || 0}${tr.vertical ? ' vertical' : ''}`);
  }
  if (fmt.borders) {
    const b = fmt.borders;
    const bParts = [];
    if (b.top) bParts.push(`top:${formatBorder(b.top)}`);
    if (b.bottom) bParts.push(`bottom:${formatBorder(b.bottom)}`);
    if (b.left) bParts.push(`left:${formatBorder(b.left)}`);
    if (b.right) bParts.push(`right:${formatBorder(b.right)}`);
    if (bParts.length) parts.push(`borders=[${bParts.join('; ')}]`);
  }
  if (fmt.padding) {
    const pad = fmt.padding;
    parts.push(`padding=${pad.top || 0},${pad.right || 0},${pad.bottom || 0},${pad.left || 0}`);
  }
  return parts.length ? [`${label}: ${parts.join(', ')}`] : [];
}

export async function readCellFormat(auth, { spreadsheetId, range }) {
  const id = extractId(spreadsheetId);
  const api = sheets(auth);
  const { sheetName } = extractSheetName(range);
  const sheetId = await getSheetId(api, id, sheetName);

  const res = await api.spreadsheets.get({
    spreadsheetId: id,
    ranges: [range],
    includeGridData: true,
    fields: 'sheets(properties(sheetId,title),merges,data(startRow,startColumn,rowMetadata.pixelSize,columnMetadata.pixelSize,rowData.values(userEnteredFormat,effectiveFormat,dataValidation,note,userEnteredValue,textFormatRuns,hyperlink)))',
  });

  const sheet = res.data.sheets?.[0];
  const data = sheet?.data?.[0];
  if (!data || !data.rowData?.length) return 'No data in range.';

  const startRow = data.startRow || 0;
  const startCol = data.startColumn || 0;
  const lines = [];

  // Row heights
  if (data.rowMetadata?.length) {
    const heights = data.rowMetadata.map((m, i) => `row ${startRow + i + 1}: ${m.pixelSize}px`);
    lines.push(`Row heights: ${heights.join(', ')}`);
  }

  // Column widths
  if (data.columnMetadata?.length) {
    const widths = data.columnMetadata.map((m, i) => `${colLetter(startCol + i)}: ${m.pixelSize}px`);
    lines.push(`Column widths: ${widths.join(', ')}`);
  }

  // Merges in range
  const merges = (sheet.merges || []).filter(m =>
    m.sheetId === sheetId &&
    m.startRowIndex < startRow + (data.rowData?.length || 0) &&
    m.endRowIndex > startRow &&
    m.startColumnIndex < startCol + (data.rowData?.[0]?.values?.length || 0) &&
    m.endColumnIndex > startCol
  );
  if (merges.length) {
    lines.push(`\nMerged cells:`);
    for (const m of merges) {
      lines.push(`  ${colLetter(m.startColumnIndex)}${m.startRowIndex + 1}:${colLetter(m.endColumnIndex - 1)}${m.endRowIndex}`);
    }
  }

  lines.push('');

  // Per-cell formatting
  for (let ri = 0; ri < data.rowData.length; ri++) {
    const row = data.rowData[ri];
    if (!row.values) continue;
    for (let ci = 0; ci < row.values.length; ci++) {
      const cell = row.values[ci];
      const uFmt = cell.userEnteredFormat;
      const eFmt = cell.effectiveFormat;
      const dv = cell.dataValidation;
      const note = cell.note;
      const val = cell.userEnteredValue;

      if (!uFmt && !eFmt && !dv && !note) continue;

      const cellRef = `${colLetter(startCol + ci)}${startRow + ri + 1}`;
      const cellLines = [cellRef];

      if (val) {
        const v = val.stringValue ?? val.numberValue ?? val.boolValue ?? val.formulaValue ?? '';
        cellLines.push(`  value: "${String(v).substring(0, 80)}"`);
      }

      const uParts = formatCellDetails(uFmt, 'set');
      const eParts = formatCellDetails(eFmt, 'effective');
      if (uParts.length) cellLines.push(`  ${uParts[0]}`);
      if (eParts.length) cellLines.push(`  ${eParts[0]}`);

      if (dv) {
        const cond = dv.condition || {};
        const dvParts = [cond.type || '?'];
        if (cond.values?.length) dvParts.push(cond.values.map(v => v.userEnteredValue).join(','));
        if (dv.strict) dvParts.push('strict');
        if (dv.inputMessage) dvParts.push(`msg="${dv.inputMessage}"`);
        cellLines.push(`  validation: ${dvParts.join('; ')}`);
      }

      if (note) cellLines.push(`  note: "${note.substring(0, 80)}"`);

      const hyperlink = cell.hyperlink;
      if (hyperlink) cellLines.push(`  hyperlink: ${hyperlink}`);

      const runs = cell.textFormatRuns;
      if (runs?.length) {
        cellLines.push(`  richText (${runs.length} runs):`);
        for (const run of runs) {
          const fmt = run.format || {};
          const parts = [];
          if (fmt.bold) parts.push('bold');
          if (fmt.italic) parts.push('italic');
          if (fmt.underline) parts.push('underline');
          if (fmt.strikethrough) parts.push('strikethrough');
          if (fmt.fontSize) parts.push(`${fmt.fontSize}pt`);
          if (fmt.fontFamily) parts.push(fmt.fontFamily);
          if (fmt.foregroundColorStyle?.rgbColor || fmt.foregroundColor?.red !== undefined) {
            const c = fmt.foregroundColorStyle?.rgbColor || fmt.foregroundColor || {};
            parts.push(`color(${c.red || 0},${c.green || 0},${c.blue || 0})`);
          }
          if (fmt.link?.uri) parts.push(`link=${fmt.link.uri}`);
          cellLines.push(`    [${run.startIndex}+] ${parts.join(', ') || '(default)'}`);
        }
      }

      lines.push(cellLines.join('\n'));
    }
  }

  if (lines.every(l => !l.trim())) return 'No formatting found in range.';
  return lines.join('\n');
}

// ── Tool definitions ──────────────────────────────────────────────────────────

export const TOOLS = [
  { name: 'format_cells', fn: formatCells, description: 'Format cells in a Google Sheet (colors, fonts, alignment, number formats)', inputSchema: { type: 'object', properties: { spreadsheetId: { type: 'string' }, range: { type: 'string', description: 'A1 range, e.g. Sheet1!A1:C10' }, format: { type: 'object', properties: { backgroundColor: { type: 'object', description: 'RGBA object with red/green/blue/alpha keys (0.0–1.0)' }, textFormat: { type: 'object', properties: { bold: { type: 'boolean' }, italic: { type: 'boolean' }, underline: { type: 'boolean' }, strikethrough: { type: 'boolean' }, fontSize: { type: 'number' }, fontFamily: { type: 'string' }, foregroundColor: { type: 'object' } } }, horizontalAlignment: { type: 'string', enum: ['LEFT', 'CENTER', 'RIGHT'] }, verticalAlignment: { type: 'string', enum: ['TOP', 'MIDDLE', 'BOTTOM'] }, wrapStrategy: { type: 'string', enum: ['OVERFLOW_CELL', 'LEGACY_WRAP', 'CLIP', 'WRAP'] }, numberFormat: { type: 'object', properties: { type: { type: 'string', enum: ['TEXT', 'NUMBER', 'PERCENT', 'CURRENCY', 'DATE', 'TIME', 'DATE_TIME', 'SCIENTIFIC'] }, pattern: { type: 'string' } } }, padding: { type: 'object', properties: { top: { type: 'number' }, right: { type: 'number' }, bottom: { type: 'number' }, left: { type: 'number' } } } } } }, required: ['spreadsheetId', 'range', 'format'] } },
  { name: 'merge_cells', fn: mergeCells, description: 'Merge cells in a Google Sheet', inputSchema: { type: 'object', properties: { spreadsheetId: { type: 'string' }, range: { type: 'string', description: 'A1 range, e.g. Sheet1!A1:C3' }, mergeType: { type: 'string', enum: ['MERGE_ALL', 'MERGE_COLUMNS', 'MERGE_ROWS'], description: 'Default: MERGE_ALL' } }, required: ['spreadsheetId', 'range'] } },
  { name: 'unmerge_cells', fn: unmergeCells, description: 'Unmerge previously merged cells in a Google Sheet', inputSchema: { type: 'object', properties: { spreadsheetId: { type: 'string' }, range: { type: 'string' } }, required: ['spreadsheetId', 'range'] } },
  { name: 'add_conditional_formatting', fn: addConditionalFormatting, description: 'Add conditional formatting rules to a Google Sheet', inputSchema: { type: 'object', properties: { spreadsheetId: { type: 'string' }, rules: { type: 'array', description: 'Array of conditional format rules', items: { type: 'object', properties: { ranges: { type: 'array', items: { type: 'string' } }, booleanRule: { type: 'object', properties: { condition: { type: 'object', properties: { type: { type: 'string' }, values: { type: 'array', items: { type: 'object' } } }, required: ['type'] }, format: { type: 'object' } } }, gradientRule: { type: 'object' } }, required: ['ranges'] } } }, required: ['spreadsheetId', 'rules'] } },
  { name: 'freeze', fn: freezeRowsColumns, description: 'Freeze rows and/or columns in a Google Sheet', inputSchema: { type: 'object', properties: { spreadsheetId: { type: 'string' }, sheetName: { type: 'string', description: 'Tab name (defaults to first sheet)' }, frozenRowCount: { type: 'number', description: 'Number of rows to freeze (0 to unfreeze)' }, frozenColumnCount: { type: 'number', description: 'Number of columns to freeze (0 to unfreeze)' } }, required: ['spreadsheetId'] } },
  { name: 'auto_resize', fn: autoResize, description: 'Auto-resize columns or rows to fit their content', inputSchema: { type: 'object', properties: { spreadsheetId: { type: 'string' }, sheetName: { type: 'string' }, dimension: { type: 'string', enum: ['COLUMNS', 'ROWS'], description: 'Default: COLUMNS' }, startIndex: { description: 'Start column letter (e.g. "A") or 0-based row index' }, endIndex: { description: 'End column letter (e.g. "D") or 0-based row index (inclusive)' } }, required: ['spreadsheetId'] } },
  { name: 'update_borders', fn: updateBorders, description: 'Update borders on a range of cells. Each side (top, bottom, left, right, innerHorizontal, innerVertical) is optional with style, width, and color.', inputSchema: { type: 'object', properties: { spreadsheetId: { type: 'string' }, range: { type: 'string', description: 'A1 range, e.g. Sheet1!A1:D10' }, top: { type: 'object', properties: { style: { type: 'string', enum: ['SOLID', 'DASHED', 'DOTTED', 'SOLID_MEDIUM', 'SOLID_THICK', 'DOUBLE', 'NONE'] }, width: { type: 'number' }, color: { type: 'object', description: 'RGBA object' } } }, bottom: { type: 'object', properties: { style: { type: 'string', enum: ['SOLID', 'DASHED', 'DOTTED', 'SOLID_MEDIUM', 'SOLID_THICK', 'DOUBLE', 'NONE'] }, width: { type: 'number' }, color: { type: 'object' } } }, left: { type: 'object', properties: { style: { type: 'string', enum: ['SOLID', 'DASHED', 'DOTTED', 'SOLID_MEDIUM', 'SOLID_THICK', 'DOUBLE', 'NONE'] }, width: { type: 'number' }, color: { type: 'object' } } }, right: { type: 'object', properties: { style: { type: 'string', enum: ['SOLID', 'DASHED', 'DOTTED', 'SOLID_MEDIUM', 'SOLID_THICK', 'DOUBLE', 'NONE'] }, width: { type: 'number' }, color: { type: 'object' } } }, innerHorizontal: { type: 'object', properties: { style: { type: 'string', enum: ['SOLID', 'DASHED', 'DOTTED', 'SOLID_MEDIUM', 'SOLID_THICK', 'DOUBLE', 'NONE'] }, width: { type: 'number' }, color: { type: 'object' } } }, innerVertical: { type: 'object', properties: { style: { type: 'string', enum: ['SOLID', 'DASHED', 'DOTTED', 'SOLID_MEDIUM', 'SOLID_THICK', 'DOUBLE', 'NONE'] }, width: { type: 'number' }, color: { type: 'object' } } } }, required: ['spreadsheetId', 'range'] } },
  { name: 'add_banding', fn: addBanding, description: 'Add alternating row color banding to a range', inputSchema: { type: 'object', properties: { spreadsheetId: { type: 'string' }, range: { type: 'string', description: 'A1 range, e.g. Sheet1!A1:D20' }, headerColor: { type: 'object', description: 'RGBA color for the header row' }, firstBandColor: { type: 'object', description: 'RGBA color for odd rows' }, secondBandColor: { type: 'object', description: 'RGBA color for even rows' }, footerColor: { type: 'object', description: 'RGBA color for the footer row' } }, required: ['spreadsheetId', 'range'] } },
  { name: 'delete_banding', fn: deleteBanding, description: 'Delete a banded range by its ID', inputSchema: { type: 'object', properties: { spreadsheetId: { type: 'string' }, bandedRangeId: { type: 'number', description: 'Banded range ID (from get_sheet_metadata)' } }, required: ['spreadsheetId', 'bandedRangeId'] } },
  { name: 'update_banding', fn: updateBanding, description: 'Update colors on an existing banded range', inputSchema: { type: 'object', properties: { spreadsheetId: { type: 'string' }, bandedRangeId: { type: 'number', description: 'Banded range ID' }, headerColor: { type: 'object', description: 'RGBA color for the header row' }, firstBandColor: { type: 'object', description: 'RGBA color for odd rows' }, secondBandColor: { type: 'object', description: 'RGBA color for even rows' }, footerColor: { type: 'object', description: 'RGBA color for the footer row' } }, required: ['spreadsheetId', 'bandedRangeId'] } },
  { name: 'read_conditional_formats', fn: readConditionalFormats, description: 'Read full details of all conditional formatting rules on a sheet: condition types, comparison values, applied formats (colors, bold, etc.), gradient color stops, and ranges. Use this to understand or replicate existing rules.', inputSchema: { type: 'object', properties: { spreadsheetId: { type: 'string' }, sheetName: { type: 'string', description: 'Tab name (omit to read all sheets)' } }, required: ['spreadsheetId'] } },
  { name: 'update_conditional_formatting', fn: updateConditionalFormatting, description: 'Update an existing conditional formatting rule by index. Can change the condition, format, or ranges.', inputSchema: { type: 'object', properties: { spreadsheetId: { type: 'string' }, sheetName: { type: 'string', description: 'Tab name' }, index: { type: 'number', description: 'Rule index (from read_conditional_formats)' }, rule: { type: 'object', description: 'Updated rule properties', properties: { ranges: { type: 'array', items: { type: 'string' }, description: 'A1 ranges' }, booleanRule: { type: 'object', description: '{ condition: { type, values }, format: { backgroundColor, textFormat } }' }, gradientRule: { type: 'object', description: '{ minpoint, midpoint, maxpoint } with { type, value, color }' } } } }, required: ['spreadsheetId', 'sheetName', 'index', 'rule'] } },
  { name: 'delete_conditional_formatting', fn: deleteConditionalFormatting, description: 'Delete a conditional formatting rule by index.', inputSchema: { type: 'object', properties: { spreadsheetId: { type: 'string' }, sheetName: { type: 'string', description: 'Tab name' }, index: { type: 'number', description: 'Rule index (from read_conditional_formats)' } }, required: ['spreadsheetId', 'sheetName', 'index'] } },
  { name: 'read_cell_format', fn: readCellFormat, description: 'Read the complete visual state of cells in a range: both user-set and effective formatting (colors, fonts, borders, number formats, alignment, rotation), row heights, column widths, merged cells, data validation rules, and notes. Use this to understand existing formatting before replicating it.', inputSchema: { type: 'object', properties: { spreadsheetId: { type: 'string' }, range: { type: 'string', description: 'A1 range, e.g. Sheet1!A1:D10' } }, required: ['spreadsheetId', 'range'] } },
];
