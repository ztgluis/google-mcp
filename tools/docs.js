import { google } from 'googleapis';
import { extractId } from '../google.js';

function docToText(doc) {
  const lines = [];

  for (const element of doc.body?.content || []) {
    if (element.paragraph) {
      const style = element.paragraph.paragraphStyle?.namedStyleType || '';
      const text = (element.paragraph.elements || [])
        .map(e => e.textRun?.content || '')
        .join('');

      if (style.startsWith('HEADING_')) {
        const level = parseInt(style.replace('HEADING_', ''), 10);
        lines.push('#'.repeat(level) + ' ' + text.trimEnd());
      } else {
        lines.push(text);
      }
    } else if (element.table) {
      for (const row of element.table.tableRows || []) {
        const cells = (row.tableCells || []).map(cell =>
          (cell.content || [])
            .flatMap(c => (c.paragraph?.elements || []).map(e => e.textRun?.content || ''))
            .join('')
            .trim()
        );
        lines.push('| ' + cells.join(' | ') + ' |');
      }
    }
  }

  return lines.join('');
}

export async function createDoc(auth, { title, folderId, body }) {
  const docs = google.docs({ version: 'v1', auth });
  const drive = google.drive({ version: 'v3', auth });

  const doc = await docs.documents.create({ requestBody: { title } });
  const docId = doc.data.documentId;

  if (folderId) {
    const file = await drive.files.get({ fileId: docId, fields: 'parents' });
    await drive.files.update({
      fileId: docId,
      addParents: folderId,
      removeParents: file.data.parents.join(','),
    });
  }

  if (body) {
    await docs.documents.batchUpdate({
      documentId: docId,
      requestBody: {
        requests: [{ insertText: { location: { index: 1 }, text: body } }],
      },
    });
  }

  return `Created: ${title}\n  ID: ${docId}\n  URL: https://docs.google.com/document/d/${docId}/edit`;
}

export async function copyDoc(auth, { fileId, title, folderId }) {
  const drive = google.drive({ version: 'v3', auth });
  const id = extractId(fileId);
  const requestBody = {};
  if (title) requestBody.name = title;
  if (folderId) requestBody.parents = [folderId];
  const res = await drive.files.copy({ fileId: id, requestBody });
  return `Copied: ${res.data.name}\n  ID: ${res.data.id}\n  URL: https://docs.google.com/document/d/${res.data.id}/edit`;
}

export async function formatDoc(auth, { fileId, operations }) {
  const docs = google.docs({ version: 'v1', auth });
  const id = extractId(fileId);

  const requests = operations.map(op => {
    if (op.type === 'textStyle') {
      const textStyle = {};
      const fields = [];
      if (op.bold !== undefined) { textStyle.bold = op.bold; fields.push('bold'); }
      if (op.italic !== undefined) { textStyle.italic = op.italic; fields.push('italic'); }
      if (op.underline !== undefined) { textStyle.underline = op.underline; fields.push('underline'); }
      if (op.strikethrough !== undefined) { textStyle.strikethrough = op.strikethrough; fields.push('strikethrough'); }
      if (op.fontSize) { textStyle.fontSize = { magnitude: op.fontSize, unit: 'PT' }; fields.push('fontSize'); }
      if (op.fontFamily) { textStyle.weightedFontFamily = { fontFamily: op.fontFamily }; fields.push('weightedFontFamily'); }
      if (op.foregroundColor) { textStyle.foregroundColor = { color: { rgbColor: op.foregroundColor } }; fields.push('foregroundColor'); }
      if (op.backgroundColor) { textStyle.backgroundColor = { color: { rgbColor: op.backgroundColor } }; fields.push('backgroundColor'); }
      if (op.link) { textStyle.link = { url: op.link }; fields.push('link'); }
      if (op.headingLink) { textStyle.link = { headingId: op.headingLink }; fields.push('link'); }
      if (op.bookmarkLink) { textStyle.link = { bookmarkId: op.bookmarkLink }; fields.push('link'); }
      if (op.removeLink) { textStyle.link = null; fields.push('link'); }
      return { updateTextStyle: { range: { startIndex: op.startIndex, endIndex: op.endIndex }, textStyle, fields: fields.join(',') } };
    }
    if (op.type === 'paragraphStyle') {
      const paragraphStyle = {};
      const fields = [];
      if (op.namedStyleType) { paragraphStyle.namedStyleType = op.namedStyleType; fields.push('namedStyleType'); }
      if (op.alignment) { paragraphStyle.alignment = op.alignment; fields.push('alignment'); }
      if (op.lineSpacing) { paragraphStyle.lineSpacing = op.lineSpacing; fields.push('lineSpacing'); }
      if (op.spaceAbove) { paragraphStyle.spaceAbove = { magnitude: op.spaceAbove, unit: 'PT' }; fields.push('spaceAbove'); }
      if (op.spaceBelow) { paragraphStyle.spaceBelow = { magnitude: op.spaceBelow, unit: 'PT' }; fields.push('spaceBelow'); }
      return { updateParagraphStyle: { range: { startIndex: op.startIndex, endIndex: op.endIndex }, paragraphStyle, fields: fields.join(',') } };
    }
    return null;
  }).filter(Boolean);

  if (!requests.length) return 'No valid operations provided.';
  await docs.documents.batchUpdate({ documentId: id, requestBody: { requests } });
  return `Applied ${requests.length} formatting operation(s).`;
}

export async function exportDoc(auth, { fileId, format }) {
  const drive = google.drive({ version: 'v3', auth });
  const id = extractId(fileId);
  const mimeTypes = {
    pdf: 'application/pdf',
    docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    txt: 'text/plain',
    html: 'text/html',
    rtf: 'application/rtf',
    epub: 'application/epub+zip',
  };
  const mimeType = mimeTypes[format];
  if (!mimeType) throw new Error(`Unsupported format: ${format}. Use: ${Object.keys(mimeTypes).join(', ')}`);
  const res = await drive.files.export({ fileId: id, mimeType }, { responseType: 'arraybuffer' });
  const buffer = Buffer.from(res.data);
  if (format === 'txt' || format === 'html') return buffer.toString('utf-8');
  return `data:${mimeType};base64,${buffer.toString('base64')}`;
}

export async function listDocTabs(auth, { fileId }) {
  const docs = google.docs({ version: 'v1', auth });
  const id = extractId(fileId);
  const res = await docs.documents.get({ documentId: id, includeTabsContent: true });
  const tabs = res.data.tabs || [];
  if (!tabs.length) return 'This document has no tabs (or uses the default single-tab layout).';
  return tabs.map(tab => {
    const props = tab.tabProperties;
    return `${props.title}\n  ID: ${props.tabId}\n  Index: ${props.index}`;
  }).join('\n\n');
}

export async function updateHeaderFooter(auth, { fileId, type, text }) {
  const docs = google.docs({ version: 'v1', auth });
  const id = extractId(fileId);
  const isHeader = type === 'header';

  const doc = await docs.documents.get({ documentId: id });
  let segmentId = isHeader
    ? doc.data.documentStyle?.defaultHeaderId
    : doc.data.documentStyle?.defaultFooterId;

  if (!segmentId) {
    const createReq = isHeader
      ? { createHeader: { type: 'DEFAULT', sectionBreakLocation: { index: 0 } } }
      : { createFooter: { type: 'DEFAULT', sectionBreakLocation: { index: 0 } } };
    const res = await docs.documents.batchUpdate({ documentId: id, requestBody: { requests: [createReq] } });
    segmentId = isHeader
      ? res.data.replies[0].createHeader.headerId
      : res.data.replies[0].createFooter.footerId;
  } else {
    const segment = isHeader ? doc.data.headers?.[segmentId] : doc.data.footers?.[segmentId];
    if (segment?.content) {
      const endIdx = segment.content.at(-1)?.endIndex || 1;
      if (endIdx > 1) {
        await docs.documents.batchUpdate({
          documentId: id,
          requestBody: { requests: [{ deleteContentRange: { range: { segmentId, startIndex: 0, endIndex: endIdx - 1 } } }] },
        });
      }
    }
  }

  await docs.documents.batchUpdate({
    documentId: id,
    requestBody: { requests: [{ insertText: { location: { segmentId, index: 0 }, text } }] },
  });
  return `Updated ${type}: "${text}"`;
}

export async function insertImage(auth, { fileId, imageUrl, index, width, height }) {
  const docs = google.docs({ version: 'v1', auth });
  const id = extractId(fileId);
  const req = { insertInlineImage: { uri: imageUrl, location: { index: index || 1 } } };
  if (width || height) {
    req.insertInlineImage.objectSize = {};
    if (width) req.insertInlineImage.objectSize.width = { magnitude: width, unit: 'PT' };
    if (height) req.insertInlineImage.objectSize.height = { magnitude: height, unit: 'PT' };
  }
  await docs.documents.batchUpdate({ documentId: id, requestBody: { requests: [req] } });
  return `Inserted image at index ${index || 1}.`;
}

export async function renameDoc(auth, { fileId, title }) {
  const drive = google.drive({ version: 'v3', auth });
  const id = extractId(fileId);
  const res = await drive.files.update({ fileId: id, requestBody: { name: title } });
  return `Renamed to: ${res.data.name}`;
}

export async function renameDocTab(auth, { fileId, tabId, title }) {
  const docs = google.docs({ version: 'v1', auth });
  const id = extractId(fileId);
  await docs.documents.batchUpdate({
    documentId: id,
    requestBody: {
      requests: [{ updateDocumentTab: { documentTab: { tabProperties: { tabId, title } }, fields: 'title' } }],
    },
  });
  return `Renamed tab ${tabId} to: "${title}"`;
}

export async function readDoc(auth, { fileId }) {
  const docs = google.docs({ version: 'v1', auth });
  const res = await docs.documents.get({ documentId: extractId(fileId) });
  return docToText(res.data);
}

export async function editDoc(auth, { fileId, operations }) {
  const docs = google.docs({ version: 'v1', auth });
  const id = extractId(fileId);

  const requests = operations.flatMap(op => {
    if (op.type === 'insert') {
      return [{ insertText: { location: { index: op.index }, text: op.text } }];
    }
    if (op.type === 'delete') {
      return [{ deleteContentRange: { range: { startIndex: op.startIndex, endIndex: op.endIndex } } }];
    }
    if (op.type === 'replace') {
      return [{ replaceAllText: { containsText: { text: op.find, matchCase: true }, replaceText: op.replaceWith } }];
    }
    if (op.type === 'append') {
      // Get end index by reading doc first — handled below
      return [{ _append: op.text }];
    }
    return [];
  });

  // Handle append: need doc end index
  const appendOps = requests.filter(r => r._append);
  const normalOps = requests.filter(r => !r._append);

  if (appendOps.length) {
    const doc = await docs.documents.get({ documentId: id });
    const endIndex = doc.data.body.content.at(-1).endIndex - 1;
    for (const op of appendOps) {
      normalOps.push({ insertText: { location: { index: endIndex }, text: op._append } });
    }
  }

  if (!normalOps.length) return 'No valid operations provided.';

  await docs.documents.batchUpdate({ documentId: id, requestBody: { requests: normalOps } });
  return `Applied ${normalOps.length} operation(s) to document.`;
}

export async function insertTable(auth, { fileId, rows, columns, index }) {
  const docs = google.docs({ version: 'v1', auth });
  const id = extractId(fileId);
  await docs.documents.batchUpdate({
    documentId: id,
    requestBody: {
      requests: [{ insertTable: { rows, columns, location: { index: index || 1 } } }],
    },
  });
  return `Inserted ${rows}x${columns} table at index ${index || 1}.`;
}

export async function modifyTable(auth, { fileId, operations }) {
  const docs = google.docs({ version: 'v1', auth });
  const id = extractId(fileId);

  const requests = operations.map(op => {
    const tableCellLocation = {
      tableStartLocation: { index: op.tableStartIndex },
      rowIndex: op.rowIndex,
      columnIndex: op.columnIndex,
    };
    if (op.type === 'insertRow') return { insertTableRow: { tableCellLocation, insertBelow: op.insertBelow !== false } };
    if (op.type === 'insertColumn') return { insertTableColumn: { tableCellLocation, insertRight: op.insertRight !== false } };
    if (op.type === 'deleteRow') return { deleteTableRow: { tableCellLocation } };
    if (op.type === 'deleteColumn') return { deleteTableColumn: { tableCellLocation } };
    return null;
  }).filter(Boolean);

  if (!requests.length) return 'No valid operations provided.';
  await docs.documents.batchUpdate({ documentId: id, requestBody: { requests } });
  return `Applied ${requests.length} table operation(s).`;
}

export async function updateList(auth, { fileId, startIndex, endIndex, type }) {
  const docs = google.docs({ version: 'v1', auth });
  const id = extractId(fileId);

  const requests = [];
  if (type === 'remove') {
    requests.push({ deleteParagraphBullets: { range: { startIndex, endIndex } } });
  } else {
    const bulletPreset = type === 'numbered' ? 'NUMBERED_DECIMAL_NESTED' : 'BULLET_DISC_CIRCLE_SQUARE';
    requests.push({ createParagraphBullets: { range: { startIndex, endIndex }, bulletPreset } });
  }

  await docs.documents.batchUpdate({ documentId: id, requestBody: { requests } });
  return type === 'remove' ? 'Removed list formatting.' : `Applied ${type} list formatting.`;
}

export async function insertPageBreak(auth, { fileId, index }) {
  const docs = google.docs({ version: 'v1', auth });
  const id = extractId(fileId);
  await docs.documents.batchUpdate({
    documentId: id,
    requestBody: { requests: [{ insertPageBreak: { location: { index } } }] },
  });
  return `Inserted page break at index ${index}.`;
}

export async function createNamedRange(auth, { fileId, name, startIndex, endIndex }) {
  const docs = google.docs({ version: 'v1', auth });
  const id = extractId(fileId);
  const res = await docs.documents.batchUpdate({
    documentId: id,
    requestBody: {
      requests: [{ createNamedRange: { name, range: { startIndex, endIndex } } }],
    },
  });
  const namedRangeId = res.data.replies[0].createNamedRange.namedRangeId;
  return `Created named range "${name}" (ID: ${namedRangeId}).`;
}

export async function deleteNamedRange(auth, { fileId, name, namedRangeId }) {
  const docs = google.docs({ version: 'v1', auth });
  const id = extractId(fileId);
  const req = namedRangeId
    ? { deleteNamedRange: { namedRangeId } }
    : { deleteNamedRange: { name } };
  await docs.documents.batchUpdate({ documentId: id, requestBody: { requests: [req] } });
  return `Deleted named range ${namedRangeId || name}.`;
}

export async function listNamedRanges(auth, { fileId }) {
  const docs = google.docs({ version: 'v1', auth });
  const id = extractId(fileId);
  const doc = await docs.documents.get({ documentId: id });
  const ranges = doc.data.namedRanges || {};
  if (!Object.keys(ranges).length) return 'No named ranges found.';
  return Object.entries(ranges).map(([name, data]) => {
    const nr = data.namedRanges[0];
    const rangeInfo = nr.ranges.map(r => `${r.startIndex}-${r.endIndex}`).join(', ');
    return `${name}\n  ID: ${nr.namedRangeId}\n  Ranges: ${rangeInfo}`;
  }).join('\n\n');
}
