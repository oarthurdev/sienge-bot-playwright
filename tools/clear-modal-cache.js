#!/usr/bin/env node
const fs = require('fs');
const path = require('path');

const cwd = process.cwd();
const pattern = /^\.modal-cache-(.+?)\.json$/i;

const files = fs.readdirSync(cwd).filter(f => pattern.test(f));

if (!files.length) {
  console.log('Nenhum arquivo .modal-cache-*.json encontrado.');
  process.exit(0);
}

for (const f of files) {
  const p = path.join(cwd, f);
  try {
    fs.unlinkSync(p);
    console.log('Removido:', p);
  } catch (e) {
    console.error('Falha ao remover', p, e && e.message);
  }
}
console.log('Limpeza de cache de modais concluída.');
