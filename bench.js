#!/usr/bin/env node
const { spawnSync } = require('child_process');
const fs = require('fs');
const path = require('path');

const STATUS_FILE = '/tmp/report-status.json';
const SCRIPT = path.resolve(__dirname, 'script.js');

const count = Number(process.env.BENCH_COUNT || 2);
const indices = [];
for (let i = 1; i <= count; i++) indices.push(i);

function readStatus() {
  try {
    const txt = fs.readFileSync(STATUS_FILE, 'utf8');
    return JSON.parse(txt);
  } catch (e) { return null; }
}

for (const idx of indices) {
  console.log(`\n=== Bench: running report #${idx} ===`);
  try { if (fs.existsSync(STATUS_FILE)) fs.unlinkSync(STATUS_FILE); } catch {}

  const res = spawnSync('node', [SCRIPT, '--single', String(idx)], {
    stdio: 'inherit',
    env: process.env,
    cwd: process.cwd(),
    timeout: 0,
  });

  console.log(`Process exited with code ${res.status}`);

  const status = readStatus();
  if (!status) {
    console.log('No status file found or failed to parse.');
    continue;
  }

  console.log('--- Overall ---');
  console.log(JSON.stringify(status.overall || {}, null, 2));

  console.log('--- Per reports ---');
  try {
    const reps = status.reports || {};
    for (const [name, info] of Object.entries(reps)) {
      console.log(`${name}: ${JSON.stringify(info, null, 2)}`);
    }
  } catch (e) { console.log('failed to read per-report info'); }
}

console.log('\nBenchmark finished.');
