const fs = require('fs');
const path = require('path');

const dir = '/tmp';
const files = fs.readdirSync(dir).filter(f => f.match(/^report-status(?:-.+?)?\.json$/)).map(f => path.join(dir,f));

for (const f of files) {
  try {
    const raw = fs.readFileSync(f, 'utf8');
    const obj = JSON.parse(raw);

    // Ensure instanceId
    obj.instanceId = obj.instanceId || (obj.overall && obj.overall.instanceId) || path.basename(f).replace(/^report-status-/, '').replace(/\.json$/, '');

    // Normalize reports
    if (obj.reports && typeof obj.reports === 'object') {
      for (const [k,v] of Object.entries(obj.reports)) {
        if (!v || typeof v !== 'object') continue;
        if (Object.prototype.hasOwnProperty.call(v, 'progress')) {
          if (v.progress === null) {
            // leave null
          } else if (typeof v.progress === 'string') {
            const n = Number(v.progress);
            v.progress = Number.isNaN(n) ? 0 : n;
          } else if (typeof v.progress !== 'number') {
            v.progress = 0;
          }
        } else {
          // try to infer from processedItems/totalItems
          if (typeof v.processedItems === 'number' && typeof v.totalItems === 'number' && v.totalItems > 0) {
            v.progress = Math.round((v.processedItems / v.totalItems) * 100);
          } else {
            v.progress = (v.status === 'completed' || v.status === 'done') ? 100 : (v.status === 'running' ? 1 : 0);
          }
        }
        // ensure status exists
        v.status = v.status || (v.progress === 100 ? 'completed' : (v.progress > 0 ? 'running' : 'pending'));
      }
    }

    // Recalculate overall.progress as average of numeric per-report progress
    try {
      const reps = obj.reports ? Object.values(obj.reports) : [];
      const progVals = reps.map(r => (typeof r.progress === 'number' ? r.progress : null)).filter(x => x !== null);
      if (!obj.overall) obj.overall = {};
      if (progVals.length) obj.overall.progress = Math.round(progVals.reduce((a,b)=>a+b,0)/progVals.length);
      else obj.overall.progress = typeof obj.overall.progress === 'number' ? obj.overall.progress : 0;

      // ensure running/startedAt etc exist
      obj.overall.total = obj.overall.total || (obj.reports ? Object.keys(obj.reports).length : 0);
      obj.overall.lastUpdated = new Date().toISOString();
    } catch (e){ /* ignore */ }

    fs.writeFileSync(f, JSON.stringify(obj, null, 2), 'utf8');
    console.log('Fixed', f);
  } catch (e) {
    console.error('Failed', f, e && e.message);
  }
}
