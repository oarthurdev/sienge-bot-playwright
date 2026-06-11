#!/usr/bin/env node
const { spawn } = require('child_process');
const fs = require('fs');
const path = require('path');

function usage() {
  console.error('Usage: node tools/profile_reports.js -- <command> [args...]');
  console.error('Example: node tools/profile_reports.js -- node script.js --single 1');
  process.exit(1);
}

const sep = process.argv.indexOf('--');
if (sep === -1 || process.argv.length <= sep + 1) usage();

const cmd = process.argv[sep + 1];
const args = process.argv.slice(sep + 2);

const outDir = path.resolve(process.cwd(), 'profile-logs');
if (!fs.existsSync(outDir)) fs.mkdirSync(outDir, { recursive: true });

const ts = new Date().toISOString().replace(/[:.]/g, '-');
const stdoutFile = path.join(outDir, `stdout-${ts}.log`);
const stderrFile = path.join(outDir, `stderr-${ts}.log`);
const summaryFile = path.join(outDir, `summary-${ts}.json`);

const outStream = fs.createWriteStream(stdoutFile);
const errStream = fs.createWriteStream(stderrFile);

console.log(`Profiling command: ${cmd} ${args.join(' ')}`);
console.log(`Writing logs to ${outDir}`);

const startHr = process.hrtime.bigint();
const startTime = Date.now();

const child = spawn(cmd, args, { shell: false });

// simple step parser: look for lines like "STEP START name" and "STEP END name"
const steps = {};
function parseLine(line) {
  const m = line.toString().trim().match(/^STEP\s+(START|END)\s+(.+)$/i);
  if (!m) return;
  const when = m[1].toUpperCase();
  const name = m[2].trim();
  const t = Date.now();
  if (when === 'START') steps[name] = steps[name] || { start: t, end: null };
  if (when === 'END') {
    steps[name] = steps[name] || { start: null, end: t };
    steps[name].end = t;
  }
}

child.stdout.on('data', chunk => {
  outStream.write(chunk);
  const lines = chunk.toString().split(/\r?\n/);
  lines.forEach(l => parseLine(l));
});
child.stderr.on('data', chunk => {
  errStream.write(chunk);
  const lines = chunk.toString().split(/\r?\n/);
  lines.forEach(l => parseLine(l));
});

child.on('close', code => {
  const endHr = process.hrtime.bigint();
  const endTime = Date.now();
  const durationMs = Number(endHr - startHr) / 1e6;

  outStream.end();
  errStream.end();

  // convert steps to durations
  const stepsSummary = Object.entries(steps).map(([k, v]) => ({
    name: k,
    start: v.start,
    end: v.end,
    durationMs: v.start && v.end ? v.end - v.start : null,
  }));

  const summary = {
    command: `${cmd} ${args.join(' ')}`,
    exitCode: code,
    startedAt: new Date(startTime).toISOString(),
    endedAt: new Date(endTime).toISOString(),
    durationMs,
    stdoutFile: path.relative(process.cwd(), stdoutFile),
    stderrFile: path.relative(process.cwd(), stderrFile),
    steps: stepsSummary,
  };

  fs.writeFileSync(summaryFile, JSON.stringify(summary, null, 2));
  console.log('Profile summary written to', summaryFile);
  console.log('Duration (ms):', Math.round(durationMs));
  if (stepsSummary.length) console.log('Detected steps:', stepsSummary.map(s => `${s.name} ${s.durationMs}ms`).join(', '));
  process.exit(code);
});

child.on('error', err => {
  console.error('Failed to start command:', err.message);
  process.exit(2);
});
