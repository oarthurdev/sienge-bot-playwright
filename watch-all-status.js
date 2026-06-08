#!/usr/bin/env node
const fs = require('fs');
const path = require('path');
const readline = require('readline');

const dirsToWatch = [process.argv[2] || '/tmp', process.cwd()];
const pattern = process.env.STATUS_GLOB_PATTERN || /^report-status(?:-(.+?))?\.json$/i;
const pollInterval = Number(process.env.INTERVAL_MS) || 1000;
const width = process.stdout.columns || 120;

const ansi = {
  reset: '\u001b[0m', bold: '\u001b[1m', dim: '\u001b[2m',
  red: '\u001b[31m', green: '\u001b[32m', yellow: '\u001b[33m', blue: '\u001b[34m', cyan: '\u001b[36m'
};
function color(s, c){ return (c||'')+s+ansi.reset; }
function clear(){ readline.cursorTo(process.stdout, 0, 0); readline.clearScreenDown(process.stdout); }
function pad(s,n){ s=String(s||''); return s.length>=n? s.slice(0,n): s+ ' '.repeat(n-s.length); }
function bar(p,w){ const pct = Math.max(0, Math.min(100, Number(p)||0)); const f = Math.round((pct/100)*w); return '['+'█'.repeat(f)+' '.repeat(w-f)+'] '+pct+'%'; }
function miniBar(p,w){ const pct = Math.max(0, Math.min(100, Number(p)||0)); const f = Math.round((pct/100)*w); return '█'.repeat(f) + ' '.repeat(w-f) + ` ${pct}%`; }

function findStatusFiles(){
  const out = new Map();
  for (const d of dirsToWatch){
    let files = [];
    try{ files = fs.readdirSync(d); }catch(e){ continue; }
    for (const f of files){
      const m = f.match(pattern);
      if (m){ out.set(path.resolve(d,f), m[1]||f); }
    }
  }
  return out; // map filepath -> instanceId
}

function readJson(p){ try{ return JSON.parse(fs.readFileSync(p,'utf8')); }catch(e){ return null; } }

function aggregate(statusMap){
  // statusMap: filepath -> parsed json
  const instances = {};
  const reportsByName = {};
  for (const [file, obj] of Object.entries(statusMap)){
    if (!obj) continue;
    const inst = obj.instanceId || obj.overall && obj.overall.instanceId || path.basename(file).replace(/^report-status-/, '').replace(/\.json$/,'');
    instances[inst] = obj;
    const reps = obj.reports || {};
    for (const [rname, r] of Object.entries(reps)){
      reportsByName[rname] = reportsByName[rname] || {};
      reportsByName[rname][inst] = r;
    }
  }
  return { instances, reportsByName };
}

function renderAll(statusFilesMap){
  clear();
  const now = new Date().toLocaleString();
  process.stdout.write(`${ansi.bold}${ansi.cyan}Multi-Instance Report Watcher${ansi.reset} — ${now}\n\n`);

  const entries = Array.from(statusFilesMap.entries());
  if (!entries.length){ process.stdout.write(color('Nenhum arquivo report-status*.json encontrado nos diretórios configurados.\n', ansi.yellow)); return; }

  const statuses = {};
  for (const [file, instId] of entries){ statuses[instId] = readJson(file); }

  // Overall summary across instances
  const totals = { instances: Object.keys(statuses).length, reports: 0, completed: 0 };
  let sumProg = 0; let countProg = 0;
  for (const [inst, s] of Object.entries(statuses)){
    if (!s) continue;
    const overall = s.overall || {};
    if (typeof overall.progress === 'number'){ sumProg += overall.progress; countProg++; }
    if (Array.isArray(overall.completedReports)) totals.completed += overall.completedReports.length;
  }
  const avgProg = countProg ? Math.round(sumProg / countProg) : 0;
  process.stdout.write(`${ansi.bold}Instances:${ansi.reset} ${totals.instances}   ${ansi.bold}Avg progress:${ansi.reset} ${color(avgProg+'%', ansi.blue)}   ${ansi.bold}Completed reports total:${ansi.reset} ${totals.completed}\n\n`);

  // Per-instance mini-summary
  process.stdout.write(`${ansi.bold}Instances:${ansi.reset}\n`);
  const instCol = Math.max(12, Math.min(24, Math.floor(width*0.18)));
  const barCol = Math.max(20, Math.min(40, Math.floor(width*0.25)));
  for (const [inst, s] of Object.entries(statuses)){
    const overall = s && s.overall ? s.overall : {};
    const rprog = typeof overall.progress === 'number' ? overall.progress : '-';
    const state = overall.running ? color('RUNNING', ansi.green) : color('IDLE', ansi.dim);
    process.stdout.write(` ${pad(inst, instCol)} ${pad(state,10)} ${pad(bar(rprog, Math.min(barCol,20)), barCol+8)} started:${pad(overall.startedAt?new Date(overall.startedAt).toLocaleTimeString():'-',8)}\n`);
  }

  // Per-report aggregated table
  process.stdout.write('\n' + ansi.bold + 'Reports (aggregated across instances):' + ansi.reset + '\n');
  const allReports = new Set();
  for (const s of Object.values(statuses)){
    if (!s || !s.reports) continue;
    for (const r of Object.keys(s.reports)) allReports.add(r);
  }
  const reportNames = Array.from(allReports).sort();
  const nameCol = Math.min(48, Math.max(20, Math.floor(width*0.40)));
  const instancesOrdered = Object.keys(statuses);
  const perInstCol = Math.max(12, Math.min(20, Math.floor((width - nameCol - 4) / Math.max(1, instancesOrdered.length))));

  // header
  let hdr = pad('Report', nameCol) + ' ';
  for (const inst of instancesOrdered){ hdr += pad(inst, perInstCol) + ' '; }
  process.stdout.write(ansi.dim + hdr + ansi.reset + '\n');
  // underline
  process.stdout.write(pad('', nameCol).replace(/ /g,'-') + ' ' + instancesOrdered.map(()=>('-'.repeat(perInstCol))).join(' ') + '\n');

  for (const rname of reportNames){
    let line = pad(rname, nameCol) + ' ';
    for (const inst of instancesOrdered){
      const rep = statuses[inst] && statuses[inst].reports && statuses[inst].reports[rname];
      if (!rep){ line += pad('-', perInstCol) + ' '; continue; }

      const prog = typeof rep.progress === 'number' ? rep.progress : null;
      const status = (rep.status || '').toLowerCase();
      let cell = '';

      if (status === 'completed' || prog === 100){ cell = color('DONE', ansi.green); }
      else if (status === 'running' || (typeof prog === 'number' && prog > 0 && prog < 100)) { cell = color('RUN', ansi.yellow); }
      else if (status === 'error' || rep.lastError) { cell = color('ERR', ansi.red); }
      else { cell = color('IDLE', ansi.dim); }

      // compact bar width based on column
      const barWidth = Math.max(4, perInstCol - 6);
      const mb = prog === null ? pad('-', barWidth) : miniBar(prog, barWidth);
      const right = `${mb}`;

      // combine and pad
      const combined = `${cell} ${right}`.slice(0, perInstCol);
      line += pad(combined, perInstCol) + ' ';
    }
    process.stdout.write(line + '\n');
  }
}

// initial discovery
let filesMap = findStatusFiles();
renderAll(filesMap);

// watch directories
for (const d of dirsToWatch){
  try{
    fs.watch(d, { persistent: true }, (ev, filename) => {
      if (!filename) { filesMap = findStatusFiles(); renderAll(filesMap); return; }
      if (pattern.test(filename)) { filesMap = findStatusFiles(); renderAll(filesMap); }
    });
  }catch(e){ /* ignore */ }
}

// poll fallback
setInterval(()=>{ const newMap = findStatusFiles(); const a = Array.from(newMap.keys()).join('|'); const b = Array.from(filesMap.keys()).join('|'); if (a !== b) { filesMap = newMap; renderAll(filesMap); return; } // if same list, still refresh content
  renderAll(filesMap);
}, pollInterval);

process.on('SIGINT', ()=>{ console.log('\nWatcher terminado.'); process.exit(0); });
