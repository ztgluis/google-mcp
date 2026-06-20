import { google } from 'googleapis';
import { extractId } from '../google.js';

export async function searchDrive(auth, { query, mimeType, maxResults = 10 }) {
  const drive = google.drive({ version: 'v3', auth });
  const safe = query.replace(/\\/g, '\\\\').replace(/'/g, "\\'");

  let q = `(name contains '${safe}' or fullText contains '${safe}') and trashed=false`;
  if (mimeType) q += ` and mimeType='${mimeType}'`;

  const res = await drive.files.list({
    q,
    pageSize: Math.min(maxResults, 50),
    fields: 'files(id,name,mimeType,modifiedTime,webViewLink)',
    orderBy: 'modifiedTime desc',
  });

  const files = res.data.files || [];
  if (!files.length) return 'No files found.';

  return files
    .map(f => `${f.name}\n  ID: ${f.id}\n  Type: ${f.mimeType}\n  Modified: ${f.modifiedTime}\n  URL: ${f.webViewLink}`)
    .join('\n\n');
}

export async function trashFile(auth, { fileId, permanent = false }) {
  const drive = google.drive({ version: 'v3', auth });

  if (permanent) {
    await drive.files.delete({ fileId, supportsAllDrives: true });
    return `Permanently deleted file ${fileId}.`;
  }

  const res = await drive.files.update({
    fileId,
    supportsAllDrives: true,
    requestBody: { trashed: true },
    fields: 'id,name,trashed',
  });

  return `Trashed: ${res.data.name} (${res.data.id})`;
}

export async function renameFile(auth, { fileId, name }) {
  const drive = google.drive({ version: 'v3', auth });
  const id = extractId(fileId);
  const res = await drive.files.update({
    fileId: id,
    requestBody: { name },
    fields: 'id,name,webViewLink',
  });
  return `Renamed to: ${res.data.name}\n  ID: ${res.data.id}\n  URL: ${res.data.webViewLink}`;
}

export async function moveFile(auth, { fileId, folderId }) {
  const drive = google.drive({ version: 'v3', auth });
  const id = extractId(fileId);
  const file = await drive.files.get({ fileId: id, fields: 'parents' });
  const previousParents = (file.data.parents || []).join(',');
  const res = await drive.files.update({
    fileId: id,
    addParents: folderId,
    removeParents: previousParents,
    fields: 'id,name,parents,webViewLink',
  });
  return `Moved ${res.data.name} to folder ${folderId}\n  ID: ${res.data.id}\n  URL: ${res.data.webViewLink}`;
}

export async function copyFile(auth, { fileId, name, folderId }) {
  const drive = google.drive({ version: 'v3', auth });
  const id = extractId(fileId);
  const requestBody = {};
  if (name) requestBody.name = name;
  if (folderId) requestBody.parents = [folderId];
  const res = await drive.files.copy({
    fileId: id,
    requestBody,
    fields: 'id,name,mimeType,webViewLink',
  });
  return `Copied: ${res.data.name}\n  ID: ${res.data.id}\n  Type: ${res.data.mimeType}\n  URL: ${res.data.webViewLink}`;
}

export async function listFolder(auth, { folderId, maxResults = 20, orderBy }) {
  const drive = google.drive({ version: 'v3', auth });
  const q = `'${folderId}' in parents and trashed=false`;
  const res = await drive.files.list({
    q,
    pageSize: Math.min(maxResults, 100),
    fields: 'files(id,name,mimeType,modifiedTime,size,webViewLink)',
    orderBy: orderBy || 'name',
  });
  const files = res.data.files || [];
  if (!files.length) return 'Folder is empty.';
  return files.map(f => {
    const lines = [f.name, `  ID: ${f.id}`, `  Type: ${f.mimeType}`, `  Modified: ${f.modifiedTime}`];
    if (f.size) lines.push(`  Size: ${f.size} bytes`);
    if (f.webViewLink) lines.push(`  URL: ${f.webViewLink}`);
    return lines.join('\n');
  }).join('\n\n');
}

export async function createFolder(auth, { name, parentFolderId }) {
  const drive = google.drive({ version: 'v3', auth });
  const requestBody = { name, mimeType: 'application/vnd.google-apps.folder' };
  if (parentFolderId) requestBody.parents = [parentFolderId];
  const res = await drive.files.create({
    requestBody,
    fields: 'id,name,webViewLink',
  });
  return `Created folder: ${res.data.name}\n  ID: ${res.data.id}\n  URL: ${res.data.webViewLink}`;
}

export async function getFileMetadata(auth, { fileId }) {
  const drive = google.drive({ version: 'v3', auth });
  const id = extractId(fileId);
  const res = await drive.files.get({
    fileId: id,
    fields: 'id,name,mimeType,modifiedTime,createdTime,size,parents,webViewLink,owners,shared,trashed',
  });
  const f = res.data;
  const lines = [f.name, `  ID: ${f.id}`, `  Type: ${f.mimeType}`, `  Created: ${f.createdTime}`, `  Modified: ${f.modifiedTime}`];
  if (f.size) lines.push(`  Size: ${f.size} bytes`);
  if (f.parents) lines.push(`  Parent folders: ${f.parents.join(', ')}`);
  if (f.owners) lines.push(`  Owner: ${f.owners.map(o => o.emailAddress).join(', ')}`);
  lines.push(`  Shared: ${f.shared}`, `  Trashed: ${f.trashed}`);
  if (f.webViewLink) lines.push(`  URL: ${f.webViewLink}`);
  return lines.join('\n');
}

export async function addComment(auth, { fileId, content, quotedText }) {
  const drive = google.drive({ version: 'v3', auth });
  const id = extractId(fileId);
  const requestBody = { content };
  if (quotedText) requestBody.quotedFileContent = { value: quotedText };
  const res = await drive.comments.create({
    fileId: id,
    requestBody,
    fields: 'id,content,author,createdTime',
  });
  return `Comment added (ID: ${res.data.id}): "${res.data.content}"`;
}

export async function listComments(auth, { fileId, includeResolved = false }) {
  const drive = google.drive({ version: 'v3', auth });
  const id = extractId(fileId);
  const res = await drive.comments.list({
    fileId: id,
    fields: 'comments(id,content,author,createdTime,resolved,quotedFileContent,replies)',
    includeDeleted: false,
  });
  let comments = res.data.comments || [];
  if (!includeResolved) comments = comments.filter(c => !c.resolved);
  if (!comments.length) return 'No comments found.';
  return comments.map(c => {
    const lines = [`${c.author?.displayName || 'Unknown'} (${c.createdTime})`, `  "${c.content}"`];
    if (c.quotedFileContent) lines.push(`  Quoted: "${c.quotedFileContent.value}"`);
    if (c.resolved) lines.push('  [RESOLVED]');
    lines.push(`  ID: ${c.id}`);
    if (c.replies?.length) {
      for (const r of c.replies) lines.push(`  Reply by ${r.author?.displayName}: "${r.content}"`);
    }
    return lines.join('\n');
  }).join('\n\n');
}

export async function resolveComment(auth, { fileId, commentId }) {
  const drive = google.drive({ version: 'v3', auth });
  const id = extractId(fileId);
  await drive.replies.create({
    fileId: id,
    commentId,
    requestBody: { content: '', action: 'resolve' },
    fields: 'id',
  });
  return `Resolved comment ${commentId}.`;
}
