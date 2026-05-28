// Run once to authenticate: npm run auth
import { createServer } from 'http';
import { writeFileSync } from 'fs';
import { exec } from 'child_process';
import path from 'path';
import { fileURLToPath } from 'url';
import { getOAuthClient, SCOPES, TOKEN_PATH } from './google.js';

const client = getOAuthClient();

const authUrl = client.generateAuthUrl({
  access_type: 'offline',
  scope: SCOPES,
  prompt: 'consent', // force refresh_token to be returned
});

console.log('\nOpening browser for Google sign-in...');
console.log('If the browser does not open, paste this URL manually:\n');
console.log(authUrl + '\n');

const openCmd = process.platform === 'darwin' ? 'open' : process.platform === 'win32' ? 'start ""' : 'xdg-open';
exec(`${openCmd} "${authUrl}"`);

const server = createServer(async (req, res) => {
  const url = new URL(req.url, 'http://localhost:3000');
  const code = url.searchParams.get('code');
  const error = url.searchParams.get('error');

  if (error) {
    res.writeHead(400);
    res.end(`<h2>Auth denied: ${error}</h2>`);
    console.error('Auth denied:', error);
    server.close();
    process.exit(1);
  }

  if (!code) {
    res.writeHead(400);
    res.end('<h2>No code received.</h2>');
    return;
  }

  try {
    const { tokens } = await client.getToken(code);
    writeFileSync(TOKEN_PATH, JSON.stringify(tokens, null, 2));
    res.writeHead(200, { 'Content-Type': 'text/html' });
    res.end('<h2>Authenticated. You can close this tab.</h2>');
    console.log('Token saved. Authentication complete.\n');
    server.close();
    process.exit(0);
  } catch (err) {
    res.writeHead(500);
    res.end(`<h2>Error: ${err.message}</h2>`);
    console.error('Token exchange failed:', err.message);
    server.close();
    process.exit(1);
  }
});

server.listen(3000, () => {
  console.log('Waiting for Google callback on http://localhost:3000 ...\n');
});
