import { google } from 'googleapis';
import { readFileSync, writeFileSync, existsSync } from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));

export const CREDENTIALS_PATH =
  process.env.GOOGLE_CREDENTIALS_PATH || path.join(__dirname, 'credentials.json');
export const TOKEN_PATH = path.join(__dirname, 'token.json');

export const SCOPES = [
  'https://www.googleapis.com/auth/drive',
  'https://www.googleapis.com/auth/documents',
  'https://www.googleapis.com/auth/spreadsheets',
];

export function getOAuthClient() {
  // Prefer env vars; fall back to credentials.json
  const clientId = process.env.GOOGLE_CLIENT_ID;
  const clientSecret = process.env.GOOGLE_CLIENT_SECRET;

  if (clientId && clientSecret) {
    return new google.auth.OAuth2(clientId, clientSecret, 'http://localhost:3000');
  }

  if (existsSync(CREDENTIALS_PATH)) {
    const raw = JSON.parse(readFileSync(CREDENTIALS_PATH, 'utf8'));
    const { client_id, client_secret } = raw.installed || raw.web;
    return new google.auth.OAuth2(client_id, client_secret, 'http://localhost:3000');
  }

  throw new Error('No credentials found. Set GOOGLE_CLIENT_ID + GOOGLE_CLIENT_SECRET env vars, or place credentials.json in the project folder.');
}

export async function getAuth() {
  if (!existsSync(TOKEN_PATH)) {
    throw new Error('Not authenticated. Run: npm run auth');
  }

  const client = getOAuthClient();
  const token = JSON.parse(readFileSync(TOKEN_PATH, 'utf8'));
  client.setCredentials(token);

  // Refresh proactively if expiring within 60 seconds
  if (token.expiry_date && token.expiry_date < Date.now() + 60_000) {
    const { credentials } = await client.refreshAccessToken();
    writeFileSync(TOKEN_PATH, JSON.stringify(credentials, null, 2));
    client.setCredentials(credentials);
  }

  return client;
}

export function extractId(idOrUrl) {
  const patterns = [
    /\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/,
    /\/document\/d\/([a-zA-Z0-9-_]+)/,
    /\/file\/d\/([a-zA-Z0-9-_]+)/,
    /[?&]id=([a-zA-Z0-9-_]+)/,
  ];
  for (const pattern of patterns) {
    const match = idOrUrl.match(pattern);
    if (match) return match[1];
  }
  return idOrUrl; // already a bare ID
}
