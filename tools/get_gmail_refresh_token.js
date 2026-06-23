#!/usr/bin/env node
require('dotenv').config();
const { google } = require('googleapis');
const http = require('http');
const open = require('open');

async function obtainRefreshToken({ clientId, clientSecret, scope = ['https://www.googleapis.com/auth/gmail.readonly'], timeoutMs = 120000 } = {}) {
  if (!clientId || !clientSecret) throw new Error('Missing clientId or clientSecret');

  return new Promise((resolve, reject) => {
    let port;
    const server = http.createServer(async (req, res) => {
      try {
        const url = new URL(req.url, `http://127.0.0.1`);
        const code = url.searchParams.get('code');
        if (!code) {
          res.writeHead(400, { 'Content-Type': 'text/plain; charset=utf-8' });
          res.end('Missing code');
          return;
        }

        res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' });
        res.end('<h3>Autorização recebida. Pode fechar esta janela.</h3>');

        server.close();

        const redirectUri = `http://127.0.0.1:${port}/oauth2callback`;
        const oauth2Client = new google.auth.OAuth2(clientId, clientSecret, redirectUri);
        try {
          const r = await oauth2Client.getToken(code);
          const tokens = r && r.tokens ? r.tokens : r;
          resolve(tokens);
        } catch (e) {
          reject(e);
        }
      } catch (e) {
        server.close();
        reject(e);
      }
    });

    server.on('error', (err) => reject(err));
    server.listen(0, '127.0.0.1', () => {
      port = server.address().port;
      const redirectUri = `http://127.0.0.1:${port}/oauth2callback`;
      const oauth2Client = new google.auth.OAuth2(clientId, clientSecret, redirectUri);
      const authUrl = oauth2Client.generateAuthUrl({ access_type: 'offline', scope, prompt: 'consent' });
      console.log('\nAbrindo navegador para autorização (Desktop app / loopback)...\n');
      console.log(authUrl + '\n');
      try { open(authUrl); } catch (e) { console.log('Abra a URL acima manualmente.'); }

      // fallback timeout
      setTimeout(() => {
        try { server.close(); } catch (e) {}
        reject(new Error('Timeout waiting for OAuth2 code'));
      }, timeoutMs);
    });
  });
}

async function main() {
  try {
    const clientId = process.env.GOOGLE_CLIENT_ID || '';
    const clientSecret = process.env.GOOGLE_CLIENT_SECRET || '';
    if (!clientId || !clientSecret) {
      console.error('Defina GOOGLE_CLIENT_ID e GOOGLE_CLIENT_SECRET no .env ou variáveis de ambiente.');
      process.exit(2);
    }

    const tokens = await obtainRefreshToken({ clientId, clientSecret });
    console.log('\nTokens obtidos:\n', JSON.stringify(tokens, null, 2));
    if (tokens.refresh_token) {
      console.log('\nCopie estas linhas para o seu arquivo .env:');
      console.log(`GOOGLE_CLIENT_ID=${clientId}`);
      console.log(`GOOGLE_CLIENT_SECRET=${clientSecret}`);
      console.log(`GOOGLE_REFRESH_TOKEN=${tokens.refresh_token}`);
    } else {
      console.log('\nNenhum refresh_token retornado. Tente executar novamente com prompt=consent e verifique as configurações do OAuth consent screen.');
    }
  } catch (e) {
    console.error('Erro:', e && e.message || e);
    process.exit(1);
  }
}

main();
