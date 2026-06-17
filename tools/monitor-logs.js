#!/usr/bin/env node
const fs = require('fs');
const path = require('path');

const files = process.argv.slice(2).length ? process.argv.slice(2) : [
  'sienge-authorize-log-1.json',
  'sienge-authorize-log-2.json',
  'sienge-authorize-log-3.json',
  'sienge-authorize-log-4.json',
  'sienge-authorize-log-5.json',
  'sienge-authorize-log.json',
].map(f => path.resolve(process.cwd(), f));

const pollMs = Number(process.env.POLL_MS) || 2000;

const state = {};
for (const f of files) state[f] = { lastLen: 0, lastTs: null };

function readJsonSafe(p) {
  try {
    const s = fs.readFileSync(p, 'utf8');
    return JSON.parse(s);
  } catch (e) { return null; }
}

function pretty(e) {
  try { return JSON.stringify(e, null, 2); } catch (_) { return String(e); }
}

console.log('Monitoring log files:');
for (const f of files) console.log(' -', f);
console.log('Poll interval:', pollMs, 'ms');
console.log('Press Ctrl+C to stop.\n');

setInterval(() => {
  for (const f of files) {
    try {
      if (!fs.existsSync(f)) continue;
      const arr = readJsonSafe(f);
      if (!Array.isArray(arr)) continue;
      const last = state[f].lastLen || 0;
      if (arr.length > last) {
        const newItems = arr.slice(last);
        for (const it of newItems) {
          const ts = it.ts || it.timestamp || (new Date()).toISOString();
          const lvl = (it.level || 'info').toLowerCase();
          const pre = `[${path.basename(f)}] [${ts}] [${lvl.toUpperCase()}]`;
          if (lvl === 'error' || lvl === 'fatal' || it.level && /error/i.test(it.level)) {
            console.error(pre, it.message || '', '\n', pretty(it));
          } else {
            console.log(pre, it.message || '');
          }
        }
        state[f].lastLen = arr.length;
      }
    } catch (e) {
      console.error('Monitor error for', f, e && e.message || e);
    }
  }
}, pollMs);
