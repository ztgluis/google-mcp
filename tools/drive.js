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

export async function listPermissions(auth, { fileId }) {
  const drive = google.drive({ version: 'v3', auth });
  const id = extractId(fileId);
  const res = await drive.permissions.list({
    fileId: id,
    fields: 'permissions(id,type,role,emailAddress,domain,displayName)',
  });
  const perms = res.data.permissions || [];
  if (!perms.length) return 'No permissions found.';
  return perms.map(p => {
    const lines = [`${p.displayName || p.emailAddress || p.domain || p.type}`, `  Role: ${p.role}`, `  Type: ${p.type}`, `  ID: ${p.id}`];
    if (p.emailAddress) lines.push(`  Email: ${p.emailAddress}`);
    if (p.domain) lines.push(`  Domain: ${p.domain}`);
    return lines.join('\n');
  }).join('\n\n');
}

export async function createPermission(auth, { fileId, role, type, emailAddress, domain, sendNotificationEmail = false }) {
  const drive = google.drive({ version: 'v3', auth });
  const id = extractId(fileId);
  const requestBody = { role, type };
  if (emailAddress) requestBody.emailAddress = emailAddress;
  if (domain) requestBody.domain = domain;
  const res = await drive.permissions.create({
    fileId: id,
    requestBody,
    sendNotificationEmail,
    fields: 'id,type,role,emailAddress,domain',
  });
  const p = res.data;
  return `Permission created (ID: ${p.id}): ${p.role} for ${p.emailAddress || p.domain || p.type}`;
}

export async function updatePermission(auth, { fileId, permissionId, role }) {
  const drive = google.drive({ version: 'v3', auth });
  const id = extractId(fileId);
  const res = await drive.permissions.update({
    fileId: id,
    permissionId,
    requestBody: { role },
    fields: 'id,role,type,emailAddress',
  });
  const p = res.data;
  return `Permission ${p.id} updated to role: ${p.role}`;
}

export async function deletePermission(auth, { fileId, permissionId }) {
  const drive = google.drive({ version: 'v3', auth });
  const id = extractId(fileId);
  await drive.permissions.delete({ fileId: id, permissionId });
  return `Deleted permission ${permissionId}.`;
}

export async function updateComment(auth, { fileId, commentId, content }) {
  const drive = google.drive({ version: 'v3', auth });
  const id = extractId(fileId);
  const res = await drive.comments.update({
    fileId: id,
    commentId,
    requestBody: { content },
    fields: 'id,content',
  });
  return `Comment ${res.data.id} updated: "${res.data.content}"`;
}

export async function deleteComment(auth, { fileId, commentId }) {
  const drive = google.drive({ version: 'v3', auth });
  const id = extractId(fileId);
  await drive.comments.delete({ fileId: id, commentId });
  return `Deleted comment ${commentId}.`;
}

export async function listReplies(auth, { fileId, commentId }) {
  const drive = google.drive({ version: 'v3', auth });
  const id = extractId(fileId);
  const res = await drive.replies.list({
    fileId: id,
    commentId,
    fields: 'replies(id,content,author,createdTime,action)',
  });
  const replies = res.data.replies || [];
  if (!replies.length) return 'No replies found.';
  return replies.map(r => {
    const lines = [`${r.author?.displayName || 'Unknown'} (${r.createdTime})`, `  "${r.content}"`, `  ID: ${r.id}`];
    if (r.action) lines.push(`  Action: ${r.action}`);
    return lines.join('\n');
  }).join('\n\n');
}

export async function createReply(auth, { fileId, commentId, content, action }) {
  const drive = google.drive({ version: 'v3', auth });
  const id = extractId(fileId);
  const requestBody = { content };
  if (action) requestBody.action = action;
  const res = await drive.replies.create({
    fileId: id,
    commentId,
    requestBody,
    fields: 'id,content,action',
  });
  return `Reply added (ID: ${res.data.id}): "${res.data.content}"${res.data.action ? ` [${res.data.action}]` : ''}`;
}

export async function deleteReply(auth, { fileId, commentId, replyId }) {
  const drive = google.drive({ version: 'v3', auth });
  const id = extractId(fileId);
  await drive.replies.delete({ fileId: id, commentId, replyId });
  return `Deleted reply ${replyId}.`;
}

export async function listRevisions(auth, { fileId }) {
  const drive = google.drive({ version: 'v3', auth });
  const id = extractId(fileId);
  const res = await drive.revisions.list({
    fileId: id,
    fields: 'revisions(id,modifiedTime,lastModifyingUser,size)',
  });
  const revisions = res.data.revisions || [];
  if (!revisions.length) return 'No revisions found.';
  return revisions.map(r => {
    const lines = [`Revision ${r.id}`, `  Modified: ${r.modifiedTime}`];
    if (r.lastModifyingUser) lines.push(`  By: ${r.lastModifyingUser.displayName || r.lastModifyingUser.emailAddress}`);
    if (r.size) lines.push(`  Size: ${r.size} bytes`);
    return lines.join('\n');
  }).join('\n\n');
}

export async function getRevision(auth, { fileId, revisionId }) {
  const drive = google.drive({ version: 'v3', auth });
  const id = extractId(fileId);
  const res = await drive.revisions.get({
    fileId: id,
    revisionId,
    fields: 'id,modifiedTime,lastModifyingUser,size,mimeType',
  });
  const r = res.data;
  const lines = [`Revision ${r.id}`, `  Modified: ${r.modifiedTime}`, `  MIME Type: ${r.mimeType}`];
  if (r.lastModifyingUser) lines.push(`  By: ${r.lastModifyingUser.displayName || r.lastModifyingUser.emailAddress}`);
  if (r.size) lines.push(`  Size: ${r.size} bytes`);
  return lines.join('\n');
}

export async function aboutGet(auth) {
  const drive = google.drive({ version: 'v3', auth });
  const res = await drive.about.get({
    fields: 'user,storageQuota',
  });
  const { user, storageQuota } = res.data;
  const lines = [
    `User: ${user.displayName}`,
    `  Email: ${user.emailAddress}`,
  ];
  if (storageQuota) {
    if (storageQuota.limit) lines.push(`  Storage Limit: ${storageQuota.limit} bytes`);
    lines.push(`  Usage: ${storageQuota.usage} bytes`);
    if (storageQuota.usageInDrive) lines.push(`  Drive Usage: ${storageQuota.usageInDrive} bytes`);
    if (storageQuota.usageInDriveTrash) lines.push(`  Trash Usage: ${storageQuota.usageInDriveTrash} bytes`);
  }
  return lines.join('\n');
}
