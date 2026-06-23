#!/usr/bin/env node
const imaps = (() => { try { return require('imap-simple'); } catch (e) { console.error('imap-simple not installed'); process.exit(1); } })();
require('dotenv').config();

const host = process.env.SOLO_MFA_HOST || process.env.MFA_IMAP_HOST || 'imap.gmail.com';
const port = Number(process.env.SOLO_MFA_PORT || process.env.MFA_IMAP_PORT || 993);
const user = process.env.SOLO_MFA_USER || process.env.MFA_IMAP_USER || process.env.SIENGE_USERNAME;
const password = process.env.SOLO_MFA_PASS || process.env.MFA_IMAP_PASS;
const tls = String(process.env.SOLO_MFA_TLS || process.env.MFA_IMAP_TLS || 'true').toLowerCase() !== 'false';

if (!user || !password) {
  console.error('Missing SOLO_MFA_USER/SOLO_MFA_PASS or MFA_IMAP_USER/MFA_IMAP_PASS in env');
  process.exit(2);
}

const config = { imap: { user, password, host, port, tls, authTimeout: 30000 } };

(async () => {
  console.log('Connecting to', host, 'as', user);
  let connection;
  try {
    connection = await imaps.connect(config);
    console.log('Connected. Opening INBOX...');
    const box = await connection.openBox('INBOX');
    console.log('INBOX opened:', box && (box.name || ''), 'messages:', box.messages && box.messages.total || box.messages);

    console.log('Listing boxes...');
    let boxes = null;
    if (typeof connection.getBoxes === 'function') boxes = await connection.getBoxes();
    else if (connection.imap && typeof connection.imap.getBoxes === 'function') boxes = await new Promise((res, rej) => connection.imap.getBoxes((err, b) => err ? rej(err) : res(b)));
    console.log('Boxes:', boxes ? Object.keys(boxes).slice(0,50) : 'none');

    const fetchOptions = { bodies: ['HEADER.FIELDS (FROM SUBJECT DATE)', 'TEXT'], struct: true };
    console.log('Searching UNSEEN...');
    const unseen = await connection.search(['UNSEEN'], fetchOptions).catch(e => { console.error('UNSEEN search err', e && e.message); return []; });
    console.log('UNSEEN count:', unseen.length);

    console.log('Searching ALL...');
    const all = await connection.search(['ALL'], fetchOptions).catch(e => { console.error('ALL search err', e && e.message); return []; });
    console.log('ALL count:', all.length);

    // show up to 10 headers
    const show = all.slice(0, 10);
    for (let i = 0; i < show.length; i++) {
      const msg = show[i];
      const parts = Array.isArray(msg.parts) ? msg.parts : [msg];
      const header = parts.find(p => /HEADER.FIELDS/i.test(p.which)) || parts[0];
      const text = parts.find(p => String(p.which).toUpperCase().includes('TEXT')) || parts[0];
      console.log('MSG', i, 'seqno', msg.seqno, 'attrs', msg.attributes || {});
      console.log('HEADER_SNIPPET', String(header.body || '').replace(/\r?\n/g,' ').slice(0,300));
      console.log('TEXT_SNIPPET', String(text.body || '').replace(/\s+/g,' ').slice(0,300));
    }

  } catch (err) {
    console.error('IMAP test error:', err && err.message, err && err.stack);
  } finally {
    try { if (connection) await connection.end(); } catch (e) { console.error('Error closing connection', e && e.message); }
  }
})();
