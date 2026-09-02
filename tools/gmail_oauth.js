const { google } = require('googleapis');
require('dotenv').config();

function makeOAuthClient() {
  const clientId = process.env.GOOGLE_CLIENT_ID;
  const clientSecret = process.env.GOOGLE_CLIENT_SECRET;
  const redirectUri = process.env.GOOGLE_REDIRECT_URI || 'urn:ietf:wg:oauth:2.0:oob';
  const refreshToken = process.env.GOOGLE_REFRESH_TOKEN;
  if (!clientId || !clientSecret || !refreshToken) {
    throw new Error('Missing GOOGLE_CLIENT_ID, GOOGLE_CLIENT_SECRET or GOOGLE_REFRESH_TOKEN');
  }
  const oAuth2Client = new google.auth.OAuth2(clientId, clientSecret, redirectUri);
  oAuth2Client.setCredentials({ refresh_token: refreshToken });
  return oAuth2Client;
}

function decodeBase64Url(str) {
  // Gmail API uses base64url
  str = str.replace(/-/g, '+').replace(/_/g, '/');
  while (str.length % 4) str += '=';
  return Buffer.from(str, 'base64').toString('utf8');
}

async function fetchMfaCodeWithGmailApi({query, maxResults = 10, userId = 'me'}) {
  let auth;
  try {
    auth = makeOAuthClient();
  } catch (e) {
    throw Object.assign(new Error('makeOAuthClient failed: ' + String(e && e.message || e)), { cause: e });
  }
  const gmail = google.gmail({ version: 'v1', auth });

  const q = String(query || process.env.MFA_GMRAW_QUERY || '"Código de Verificação" sienge');
  let res;
  try {
    res = await gmail.users.messages.list({ userId, q, maxResults });
  } catch (e) {
    const err = new Error('gmail.users.messages.list failed: ' + String(e && e.message || e));
    err.original = e;
    throw err;
  }
  const msgs = (res && res.data && res.data.messages) || [];
  for (const item of msgs) {
    try {
      const m = await gmail.users.messages.get({ userId, id: item.id, format: 'full' });
      const payload = m && m.data && m.data.payload;
      let text = '';
      if (m.data.snippet) text += m.data.snippet + '\n';
      if (payload) {
        const parts = payload.parts || [payload];
        for (const p of parts) {
          if (p.mimeType && p.mimeType.includes('text')) {
            if (p.body && p.body.data) text += decodeBase64Url(p.body.data) + '\n';
          } else if (p.parts) {
            for (const sub of p.parts) {
              if (sub.body && sub.body.data) text += decodeBase64Url(sub.body.data) + '\n';
            }
          }
        }
      }
      // search for 6-digit code
      const m6 = text.match(/(\d{6})/);
      if (m6) return { code: m6[1], raw: text, messageId: item.id };
    } catch (e) {
      // bubble up non-transient errors where appropriate
      // but continue to try other messages
      // attach original error for debugging
      // console.debug('gmail message fetch error', e && e.message || e);
      continue;
    }
  }
  return null;
}

module.exports = { fetchMfaCodeWithGmailApi };