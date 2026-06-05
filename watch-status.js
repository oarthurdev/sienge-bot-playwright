#!/usr/bin/env node
const fs = require('fs');
const path = require('path');

const file = process.argv[2] || '/tmp/report-status.json';
const interval = Number(process.env.INTERVAL_MS) || 1000;

if (process.argv.includes('--help') || process.argv.includes('-h')) {
  console.log('Usage: node watch-status.js [status-file]');
  process.exit(0);
}

function readJson(p) {
  try {
    return JSON.parse(fs.readFileSync(p, 'utf8'));
  } catch (e) { return null; }
}

function formatDate(s) {
  if (!s) return '-';
  try { return new Date(s).toLocaleString(); } catch (_) { return String(s); }
}

function render(status) {
  console.clear();
  const now = new Date().toLocaleString();
  console.log(`Status watcher — ${file} — ${now}\n`);
  if (!status) {
    console.log('Arquivo não encontrado ou inválido. Aguardando...');
    return;
  }

  const overall = status.overall || {};
  console.log('Overall:');
  console.log(`  running: ${overall.running ? 'yes' : 'no'}  progress: ${typeof overall.progress !== 'undefined' ? overall.progress + '%' : '-'}  current: ${overall.current ?? '-'} / ${overall.total ?? '-'}  currentReport: ${overall.currentReport ?? '-'}  error: ${overall.error || '-'}'`);
  console.log(`  startedAt: ${formatDate(overall.startedAt)}  finishedAt: ${formatDate(overall.finishedAt)}  lastUpdated: ${formatDate(overall.lastUpdated)}\n`);

  const reports = status.reports || {};
  const names = Object.keys(reports);
  if (!names.length) {
    console.log('Nenhuma entrada por relatório ainda.');
    return;
  }

  console.log('Reports:');
  for (const name of names) {
    const r = reports[name] || {};
    const prog = typeof r.progress !== 'undefined' ? `${r.progress}%` : '-';
    const processed = r.processedItems != null ? r.processedItems : '-';
    const totalItems = r.totalItems != null ? r.totalItems : '-';
    console.log(`- ${name}: ${r.status || '-'}  progress:${prog}  ${processed}/${totalItems}  step:${r.step || '-'}  lastMsg:${r.lastMessage || '-'}  lastErr:${r.lastError || '-'} `);
  }
}

let last = null;

// initial render
render(readJson(file));

const timer = setInterval(() => {
  const status = readJson(file);
  const serialized = JSON.stringify(status || {});
  if (serialized !== last) {
    last = serialized;
    render(status);
  }
}, interval);

process.on('SIGINT', () => {
  clearInterval(timer);
  console.log('\nWatcher terminado.');
  process.exit(0);
});
