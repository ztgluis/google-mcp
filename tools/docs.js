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
