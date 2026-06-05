#!/usr/bin/env node
const fs = require('fs');
const path = require('path');
const readline = require('readline');

const file = process.argv[2] || '/tmp/report-status.json';
const pollInterval = Number(process.env.INTERVAL_MS) || 1000;
const width = process.stdout.columns || 80;

function readJson(p) {
  try {
    return JSON.parse(fs.readFileSync(p, 'utf8'));
  } catch (e) { return null; }
}

function fmt(s) { if (!s) return '-'; try { return new Date(s).toLocaleString(); } catch (_) { return String(s); } }

// Simple ANSI styles
const ansi = {
  reset: '\u001b[0m',
  bold: '\u001b[1m',
  dim: '\u001b[2m',
  red: '\u001b[31m',
  green: '\u001b[32m',
  yellow: '\u001b[33m',
  cyan: '\u001b[36m',
  blue: '\u001b[34m',
};
function color(text, c) { return (c || '') + text + ansi.reset; }

function pad(s, n) { s = String(s || ''); return s.length >= n ? s.slice(0, n) : s + ' '.repeat(n - s.length); }

function renderBar(pct, w) {
  const p = Math.max(0, Math.min(100, Number(pct) || 0));
  const filled = Math.round((p / 100) * w);
  const empty = w - filled;
  return '[' + '█'.repeat(filled) + ' '.repeat(empty) + `] ${p}%`;
}

function clearScreen() {
  readline.cursorTo(process.stdout, 0, 0);
  readline.clearScreenDown(process.stdout);
}

function render(status) {
  clearScreen();
  const now = new Date().toLocaleString();
  process.stdout.write(`${ansi.bold}${ansi.cyan}Report Status Watcher${ansi.reset} — ${file} — ${now}\n\n`);

  if (!status) {
    process.stdout.write(color('Arquivo não encontrado ou inválido. Aguardando...', ansi.yellow) + '\n');
    return;
  }

  const overall = status.overall || {};
  const overallProg = typeof overall.progress !== 'undefined' ? overall.progress : '-';
  const running = overall.running ? color('RUNNING', ansi.green) : color('IDLE', ansi.dim);
  process.stdout.write(`${ansi.bold}Overall:${ansi.reset} ${running}  progress: ${color(overallProg + '%', ansi.blue)}  current: ${overall.current ?? '-'} / ${overall.total ?? '-'}\n`);
  process.stdout.write(`  startedAt: ${fmt(overall.startedAt)}  finishedAt: ${fmt(overall.finishedAt)}  lastUpdated: ${fmt(overall.lastUpdated)}\n\n`);

  const reports = status.reports || {};
  const names = Object.keys(reports);
  if (!names.length) {
    process.stdout.write(color('Nenhuma entrada por relatório ainda.', ansi.dim) + '\n');
    return;
  }

  process.stdout.write(`${ansi.bold}Reports:${ansi.reset}\n`);
  const nameCol = Math.min(40, Math.max(20, Math.floor(width * 0.35)));
  const barCol = Math.min(30, Math.max(10, Math.floor(width * 0.25)));

  for (const name of names) {
    const r = reports[name] || {};
    const statusStr = (r.status || '-').toLowerCase();
    let statusCol = ansi.dim;
    if (statusStr === 'running' || statusStr === 'processing') statusCol = ansi.green;
    else if (statusStr === 'pending') statusCol = ansi.yellow;
    else if (statusStr === 'error') statusCol = ansi.red;
    else if (statusStr === 'completed' || statusStr === 'done') statusCol = ansi.blue;

    const prog = typeof r.progress !== 'undefined' ? r.progress : 0;
    const bar = renderBar(prog, Math.max(10, Math.min(barCol, 20)));
    const processed = r.processedItems != null ? r.processedItems : '-';
    const totalItems = r.totalItems != null ? r.totalItems : '-';
    const step = r.step || '-';
    const lastMsg = r.lastMessage || r.lastError || '-';

    process.stdout.write(` ${pad(name, nameCol)} ${color(pad(statusStr.toUpperCase(), 10), statusCol)} ${pad(bar, barCol + 8)} ${pad(`${processed}/${totalItems}`, 12)} step:${pad(step, 12)} ${ansi.dim}${truncate(lastMsg, 60)}${ansi.reset}\n`);
  }
}

function truncate(s, n) { s = String(s || ''); return s.length > n ? s.slice(0, n - 1) + '…' : s; }

let lastSerialized = null;

function updateFromFile() {
  const status = readJson(file);
  const serialized = JSON.stringify(status || {});
  if (serialized !== lastSerialized) {
    lastSerialized = serialized;
    try { render(status); } catch (e) { /* ignore render errors */ }
  }
}

// Initial
updateFromFile();

// Watch for file changes
try {
  fs.watch(path.dirname(file) || '.', (eventType, filename) => {
    if (!filename) { updateFromFile(); return; }
    if (filename === path.basename(file)) updateFromFile();
  });
} catch (_) {
  // fallback to polling
}

// Poll as backup for environments where fs.watch is unreliable
const poll = setInterval(updateFromFile, pollInterval);

process.on('SIGINT', () => {
  clearInterval(poll);
  console.log('\nWatcher terminado.');
  process.exit(0);
});
