// Ported from freema/mcp-gsheets (MIT) — TypeScript types stripped, no auth dependency

export function columnToIndex(column) {
  let index = 0;
  for (let i = 0; i < column.length; i++) {
    index = index * 26 + (column.charCodeAt(i) - 'A'.charCodeAt(0) + 1);
  }
  return index - 1;
}

export function colIndexToLetter(index) {
  let result = '';
  let n = index + 1;
  while (n > 0) {
    const rem = (n - 1) % 26;
    result = String.fromCharCode(65 + rem) + result;
    n = Math.floor((n - 1) / 26);
  }
  return result;
}

export function parseRange(range, sheetId) {
  const rangePart = range.includes('!') ? range.split('!')[1] : range;
  if (!rangePart) throw new Error(`Invalid range format: ${range}`);

  const singleCell = rangePart.match(/^([A-Z]+)(\d+)$/i);
  if (singleCell) {
    const col = columnToIndex(singleCell[1].toUpperCase());
    const row = parseInt(singleCell[2]) - 1;
    return { sheetId: sheetId ?? null, startRowIndex: row, endRowIndex: row + 1, startColumnIndex: col, endColumnIndex: col + 1 };
  }

  const rangeMatch = rangePart.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/i);
  if (!rangeMatch) throw new Error(`Invalid range format: ${range}`);
  return {
    sheetId: sheetId ?? null,
    startRowIndex: parseInt(rangeMatch[2]) - 1,
    endRowIndex: parseInt(rangeMatch[4]),
    startColumnIndex: columnToIndex(rangeMatch[1].toUpperCase()),
    endColumnIndex: columnToIndex(rangeMatch[3].toUpperCase()) + 1,
  };
}

export async function getSheetId(sheets, spreadsheetId, sheetName) {
  const res = await sheets.spreadsheets.get({ spreadsheetId, fields: 'sheets.properties' });
  const all = res.data.sheets || [];

  if (sheetName) {
    const sheet = all.find(s => s.properties?.title === sheetName);
    if (!sheet) {
      const available = all.map(s => s.properties?.title).filter(Boolean).join(', ');
      throw new Error(`Sheet "${sheetName}" not found. Available: ${available}`);
    }
    return sheet.properties.sheetId;
  }

  if (all.length > 0 && all[0]?.properties?.sheetId !== undefined) return all[0].properties.sheetId;
  throw new Error('No sheets found in spreadsheet');
}

export function extractSheetName(range) {
  if (range.includes('!')) {
    const idx = range.indexOf('!');
    let sheetName = range.slice(0, idx);
    const rangePart = range.slice(idx + 1);
    if ((sheetName.startsWith('"') && sheetName.endsWith('"')) ||
        (sheetName.startsWith("'") && sheetName.endsWith("'"))) {
      sheetName = sheetName.slice(1, -1);
    }
    return { sheetName, range: rangePart };
  }
  return { range };
}
