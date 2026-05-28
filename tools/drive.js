import { google } from 'googleapis';

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
