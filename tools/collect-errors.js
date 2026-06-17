#!/usr/bin/env node
const fs = require('fs');
const path = require('path');

const inputs = process.argv.slice(2).length ? process.argv.slice(2) : [
  'sienge-authorize-log-1.json',
  'sienge-authorize-log-2.json',
  'sienge-authorize-log-3.json',
  'sienge-authorize-log-4.json',
  'sienge-authorize-log-5.json',
  'sienge-authorize-log.json',
].map(f => path.resolve(process.cwd(), f));

const outDir = path.resolve(process.cwd(), 'profile-logs');
const outFile = path.join(outDir, 'captured-errors.json');

try { fs.mkdirSync(outDir, { recursive: true }); } catch (e) {}

let collected = [];

for (const f of inputs) {
  try {
    if (!fs.existsSync(f)) continue;
    const raw = fs.readFileSync(f, 'utf8');
    const arr = JSON.parse(raw);
    if (!Array.isArray(arr)) continue;
    for (const it of arr) {
      const lvl = String(it.level || '').toLowerCase();
      const msg = String(it.message || '');
      if (lvl === 'error' || lvl === 'fatal' || /error/i.test(lvl) || /error/i.test(msg) ) {
        collected.push(Object.assign({ _source: path.basename(f) }, it));
      }
    }
  } catch (e) {
    // ignore per-file errors
    console.error('failed read', f, e && e.message);
  }
}

try {
  fs.writeFileSync(outFile, JSON.stringify({ generatedAt: new Date().toISOString(), count: collected.length, errors: collected }, null, 2), 'utf8');
  console.log('Captured', collected.length, 'error(s) ->', outFile);
} catch (e) {
  console.error('failed write', e && e.message);
  process.exit(2);
}
