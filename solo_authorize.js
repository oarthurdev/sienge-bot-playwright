#!/usr/bin/env node
const fs = require('fs');
const path = require('path');
const readline = require('readline/promises');
const { stdin: input, stdout: output } = require('process');
const { chromium } = require('playwright');
require('dotenv').config();

// IMAP e o parser de e-mail só são necessários se uma sessão expirar e exigir
// MFA. Carregá-los sob demanda reduz o heap do processo nos ciclos normais.
let imaps = null;
let mailparser = null;
let mailDependenciesLoaded = false;
function loadMailDependencies() {
  if (mailDependenciesLoaded) return;
  mailDependenciesLoaded = true;
  try { imaps = require('imap-simple'); } catch {}
  try { mailparser = require('mailparser'); } catch {}
}

const BASE_URL = process.env.SIENGE_BASE_URL;
const USERNAME = process.env.SIENGE_USERNAME;
const PASSWORD = process.env.SIENGE_PASSWORD;
const HEADLESS = (process.env.HEADLESS ?? 'true').toLowerCase() !== 'false';
const STATE_PATH = process.env.STATE_PATH || path.resolve(process.cwd(), 'solo_sienge-storage-state.json');
const SCREENSHOT_DIR = process.env.SCREENSHOT_DIR || path.resolve(process.cwd(), 'screenshots');
const LOG_PATH = process.env.LOG_PATH || path.resolve(process.cwd(), 'solo_sienge-authorize-log.json');
const DEBUG_HTML = (process.env.DEBUG_HTML ?? 'false').toLowerCase() === 'true';
// Capturas de tela completas fazem o Chromium rasterizar toda a página e podem
// consumir bastante CPU e RAM. Em produção, mantenha-as apenas para falhas.
const CAPTURE_PAGE_STATE = (process.env.CAPTURE_PAGE_STATE ?? 'false').toLowerCase() === 'true';
const CAPTURE_SUCCESS_SCREENSHOTS = (process.env.CAPTURE_SUCCESS_SCREENSHOTS ?? 'false').toLowerCase() === 'true';
const DEBUG_PAGE_EVENTS = (process.env.DEBUG_PAGE_EVENTS ?? 'false').toLowerCase() === 'true';
// Alguns provedores de SSO dependem de recursos visuais durante a transição de
// MFA. Por compatibilidade, o bloqueio é opt-in.
const BLOCK_NON_ESSENTIAL_RESOURCES = (process.env.BLOCK_NON_ESSENTIAL_RESOURCES ?? 'false').toLowerCase() === 'true';
const LOG_MAX_EVENTS = Math.max(100, Number(process.env.LOG_MAX_EVENTS || 500) || 500);
const LOG_FLUSH_EVERY = Math.max(1, Number(process.env.LOG_FLUSH_EVERY || 25) || 25);
// Uma viewport menor reduz a área de pintura/composição do renderer. Ainda é
// larga o suficiente para manter o layout desktop do Sienge.
const VIEWPORT_WIDTH = Math.max(800, Number(process.env.BROWSER_VIEWPORT_WIDTH || 1024) || 1024);
const VIEWPORT_HEIGHT = Math.max(600, Number(process.env.BROWSER_VIEWPORT_HEIGHT || 768) || 768);
const TASK_MODE = (process.env.TASK_MODE || 'authorize').toLowerCase();
const REPORT_OUTPUT_DIR = process.env.REPORT_OUTPUT_DIR || path.resolve(process.cwd(), 'reports');
const TARGET_PAGE_URL = `${BASE_URL}/sienge/8/index.html#/common/page/1777`;
const REPORT_FILTER_PAGE_URL = `${BASE_URL}/sienge/CRC/filterContasRecebidas.do`;
const TARGET_END_DATE = '31/12/2040';
const REPORT_PERIOD_START = '01/04/2026';
const REPORT_PERIOD_END = '30/04/2026';
const MFA_PREWAIT_MS = Number(process.env.MFA_PREWAIT_MS || 20000);
const ZAPI_SEND_TEXT_URL = process.env.ZAPI_SEND_TEXT_URL || 'https://api.z-api.io/instances/3E3A6DFFDC3AD016E1A29E80A122AB54/token/82CAC4DF8A0B69FC6B3D230F/send-text';
const ZAPI_CLIENT_TOKEN = process.env.ZAPI_CLIENT_TOKEN || 'Fa09c1cac17974bb2a3b812fc21c54e21S';
const ZAPI_ALERT_PHONE = process.env.ZAPI_ALERT_PHONE || '554799688517';

if (!BASE_URL || !USERNAME || !PASSWORD) {
  console.error('Faltam variáveis: SIENGE_BASE_URL, SIENGE_USERNAME, SIENGE_PASSWORD');
  process.exit(1);
}

function nowIso() { return new Date().toISOString(); }
function todayBr() {
  const d = new Date();
  const inicio = new Date(d.getFullYear(), 0, 1); // 0 = janeiro
  return `${String(inicio.getDate()).padStart(2, '0')}/${String(inicio.getMonth() + 1).padStart(2, '0')}/${inicio.getFullYear()}`;
}
function escRe(s) { return String(s).replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); }
function sanitizeFileName(name) {
  return String(name)
    .trim()
    .replace(/[<>:"/\\|?*\u0000-\u001F]/g, '-')
    .replace(/\s+/g, ' ')
    .replace(/\s-\s/g, ' - ');
}
function formatBrDate(date) {
  const d = date instanceof Date ? date : new Date(date);
  return `${String(d.getDate()).padStart(2, '0')}/${String(d.getMonth() + 1).padStart(2, '0')}/${d.getFullYear()}`;
}
function excelSerialToBrDate(serial) {
  const base = new Date(1899, 11, 30);
  base.setDate(base.getDate() + Number(serial));
  return formatBrDate(base);
}

const REPORT_DEFINITIONS = [
  {
    sheetName: 'Receita Parcelas Mensal',
    pdfName: 'Receita Parcelas Mensal',
    planoFinanceiro: [''],
    planoFinanceiroExcecao: ['Receita de Estoque de Terrenos'],
    contasCorrente: [],
    documentos: ['CONTRATO'],
    documentosExcecao: [''],
    condicoesPagamento: ['Parcelas Mensais'],
    condicoesPagamentoExcecao: [''],
    flags: {
    imprimirParcelasReparceladas: false,
    },

    ordem: 'Valor',

    processarLancamentos: 'Contas a receber',
  },
  {
    sheetName: 'Receita Reforços',
    pdfName: 'Receita Reforços',
    planoFinanceiro: [''],
    documentos: ['CONTRATO'],
    condicoesPagamento: [
      'Parcelas Semestrais',
      'Parcelas Bimestrais',
      'Parcela Única',
      'Parcela Anual',
      'Novo Parcelas Semestrais',
      'Novo Parcela Anual',
      'Novo Parcelas Bimestrais',
      'Novo Parcela Única',
    ],
  },
  {
    sheetName: 'Receita Aluguel',
    pdfName: 'Receita Aluguel',
    planoFinanceiro: [''],
    documentos: [],
    condicoesPagamento: ['Novo Parcelas Mensais Aluguel', 'Parcela Locação Aluguel Mensal'],
  },
  {
    sheetName: 'Ato + PE',
    pdfName: 'Ato + PE',
    planoFinanceiro: [''],
    documentos: ['CONTRATO'],
    condicoesPagamento: ['Parcela na Escritura', 'Novo Parcela na Escritura', 'Ato', 'Novo Ato'],
  },
  {
    sheetName: 'Venda a Vista',
    pdfName: 'Venda a Vista',
    planoFinanceiro: [''],
    documentos: ['CONTRATO'],
    condicoesPagamento: ['Venda a Vista', 'Novo Venda a Vista'],
  },
  {
    sheetName: 'Financiamento',
    pdfName: 'Financiamento',
    planoFinanceiro: [''],
    documentos: ['CONTRATO'],
    condicoesPagamento: ['Novo Financiamento', 'Financiamento'],
  },
  {
    sheetName: 'Venda Lote',
    pdfName: 'Venda Lote',
    planoFinanceiro: ['Receita de Estoque de Terrenos'],
    documentos: [],
    condicoesPagamento: [],
  },
  {
    sheetName: 'Empréstimo',
    pdfName: 'Empréstimo',
    planoFinanceiro: ['Receita de Empréstimos'],
    documentos: [],
    condicoesPagamento: [],
  },
  {
    sheetName: 'Venda de Passivo',
    pdfName: 'Venda de Passivo',
    planoFinanceiro: ['Venda de Passivos'],
    documentos: [],
    condicoesPagamento: [],
  },
  {
    sheetName: 'Receitas Diversas',
    pdfName: 'Receitas Diversas',
    planoFinanceiro: [
      ''
    ],
    documentos: [''],
    condicoesPagamento: [
      ''
    ],
  },
];

function getAllPlanoFinanceiroValues() {

  const values = [];

  for (const report of REPORT_DEFINITIONS) {

    for (const value of (report.planoFinanceiro || [])) {

      if (!value?.trim()) {
        continue;
      }

      if (!values.includes(value)) {
        values.push(value);
      }
    }
  }

  return values;
}

async function clearLegacyFilters(page) {

  // limpa inputs
  const textInputs = page.locator(
    'input[type="text"]'
  );

  const count = await textInputs.count();

  for (let i = 0; i < count; i++) {

    const input = textInputs.nth(i);

    try {
      await input.fill('');
    } catch {}
  }

  // desmarca checkboxes
  const checks = page.locator(
    'input[type="checkbox"]'
  );

  const checksCount = await checks.count();

  for (let i = 0; i < checksCount; i++) {

    const check = checks.nth(i);

    try {

      if (await check.isChecked()) {
        await check.uncheck();
      }

    } catch {}
  }
}

let logEntries;
let pendingLogEvents = 0;

function readLog() {
  if (logEntries) return logEntries;
  if (!fs.existsSync(LOG_PATH)) return (logEntries = []);
  try {
    const arr = JSON.parse(fs.readFileSync(LOG_PATH, 'utf8'));
    // Não mantenha um histórico ilimitado no heap nem o regrave a cada evento.
    return (logEntries = Array.isArray(arr) ? arr.slice(-LOG_MAX_EVENTS) : []);
  } catch {
    return (logEntries = []);
  }
}

function flushLog() {
  if (!logEntries || !pendingLogEvents) return;
  fs.writeFileSync(LOG_PATH, JSON.stringify(logEntries, null, 2), 'utf8');
  pendingLogEvents = 0;
}

function logEvent(event) {
  const payload = { ts: nowIso(), ...event };
  const arr = readLog();
  arr.push(payload);
  if (arr.length > LOG_MAX_EVENTS) arr.splice(0, arr.length - LOG_MAX_EVENTS);
  pendingLogEvents += 1;
  // Eventos comuns são agrupados; avisos e erros continuam persistidos na hora.
  if (pendingLogEvents >= LOG_FLUSH_EVERY || ['warning', 'error', 'fatal'].includes(payload.level)) flushLog();
  console.log(`[${payload.ts}] [${String(payload.level || 'info').toUpperCase()}] ${payload.message || ''}`);
}

function truncateForLog(value, maxLen = 500) {
  const s = String(value == null ? '' : value);
  return s.length > maxLen ? s.slice(0, maxLen) + '…' : s;
}

function truncateForAlert(value, maxLen = 1200) {
  const s = String(value == null ? '' : value);
  return s.length > maxLen ? `${s.slice(0, maxLen)}…` : s;
}

async function sendZapiAlert({ title, detail, stack, url, pageTitle }) {
  if (!ZAPI_SEND_TEXT_URL || !ZAPI_CLIENT_TOKEN || !ZAPI_ALERT_PHONE) {
    return false;
  }

  const lines = [
    `Bot falhou: ${title || 'falha no robô'}`,
  ];

  if (detail) lines.push(`Erro: ${truncateForAlert(detail, 400)}`);
  if (pageTitle) lines.push(`Tela: ${truncateForAlert(pageTitle, 180)}`);
  if (url) lines.push(`URL: ${truncateForAlert(url, 220)}`);
  if (stack) lines.push(`Stack: ${truncateForAlert(stack, 700)}`);

  const payload = {
    phone: ZAPI_ALERT_PHONE,
    message: lines.join('\n'),
  };

  try {
    const response = await fetch(ZAPI_SEND_TEXT_URL, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Client-Token': ZAPI_CLIENT_TOKEN,
      },
      body: JSON.stringify(payload),
    });

    if (!response.ok) {
      const responseText = await response.text().catch(() => '');
      throw new Error(`HTTP ${response.status}${responseText ? ` - ${truncateForAlert(responseText, 300)}` : ''}`);
    }

    return true;
  } catch (err) {
    logEvent({
      level: 'warning',
      message: 'Falha ao enviar alerta via Z-API.',
      detail: String(err && err.message || err),
    });
    return false;
  }
}

async function debugPageSnapshot(page, label) {
  if (!DEBUG_PAGE_EVENTS) return;
  try {
    const summary = await pageSummary(page);
    logEvent({
      level: 'debug',
      message: `Snapshot: ${label}`,
      url: summary.url,
      title: summary.title,
      bodySnippet: truncateForLog(summary.bodySnippet, 1200),
    });
  } catch (err) {
    logEvent({
      level: 'debug',
      message: `Snapshot falhou: ${label}`,
      detail: String(err && err.message || err),
    });
  }
}

async function debugLocatorState(page, label, selectors) {
  if (!DEBUG_PAGE_EVENTS) return;
  try {
    const result = {};
    for (const [name, locator] of Object.entries(selectors)) {
      try {
        const count = await locator.count();
        const firstVisible = count ? await locator.first().isVisible().catch(() => false) : false;
        result[name] = { count, firstVisible };
      } catch (err) {
        result[name] = { error: String(err && err.message || err) };
      }
    }
    logEvent({
      level: 'debug',
      message: `Estado de locators: ${label}`,
      locators: result,
      currentUrl: page.url(),
    });
  } catch (err) {
    logEvent({
      level: 'debug',
      message: `Falha ao inspecionar locators: ${label}`,
      detail: String(err && err.message || err),
    });
  }
}

async function debugFrames(page, label) {
  if (!DEBUG_PAGE_EVENTS) return;
  try {
    const frames = page.frames().map((f, i) => ({
      index: i,
      url: f.url(),
      name: f.name() || '',
    }));
    logEvent({
      level: 'debug',
      message: `Frames disponíveis: ${label}`,
      frames,
      currentUrl: page.url(),
    });
  } catch (err) {
    logEvent({
      level: 'debug',
      message: `Falha ao inspecionar frames: ${label}`,
      detail: String(err && err.message || err),
    });
  }
}


async function ensureDir(dir) { await fs.promises.mkdir(dir, { recursive: true }); }
async function saveShot(page, name) {
  await ensureDir(SCREENSHOT_DIR);
  const file = path.join(SCREENSHOT_DIR, `${Date.now()}-${name}.png`);
  // A viewport contém o estado necessário para diagnóstico sem rasterizar uma
  // página inteira (que pode conter milhares de linhas).
  await page.screenshot({ path: file });
  return file;
}
async function saveHtml(page, name) {
  if (!DEBUG_HTML) return null;
  await ensureDir(SCREENSHOT_DIR);
  const file = path.join(SCREENSHOT_DIR, `${Date.now()}-${name}.html`);
  await fs.promises.writeFile(file, await page.content(), 'utf8');
  return file;
}
async function pageSummary(page, { includeBody = true } = {}) {
  let bodyText = '';
  // innerText força o navegador a calcular o layout de toda a árvore. Só o
  // consulte quando o texto for necessário para diagnosticar ou decidir login.
  if (includeBody) {
    try { bodyText = await page.locator('body').innerText({ timeout: 6000 }); } catch {}
  }
  bodyText = String(bodyText || '').replace(/\s+/g, ' ').trim().slice(0, 3000);
  return { url: page.url(), title: await page.title().catch(() => ''), bodySnippet: bodyText };
}
async function logPageState(page, message, extra = {}) {
  const shouldCapture = CAPTURE_PAGE_STATE || ['warning', 'error'].includes(extra.level);
  const shot = shouldCapture
    ? await saveShot(page, (extra.shotName || 'state').replace(/[^a-z0-9_-]+/gi, '-'))
    : null;
  const html = shouldCapture ? await saveHtml(page, (extra.shotName || 'state').replace(/[^a-z0-9_-]+/gi, '-')) : null;
  const summary = await pageSummary(page, { includeBody: shouldCapture });
  logEvent({
    level: extra.level || 'info',
    message,
    screenshot: shot,
    html,
    ...summary,
    ...Object.fromEntries(Object.entries(extra).filter(([k]) => !['shotName', 'level'].includes(k))),
  });
}
async function waitForAppReady(
  page,
  timeout = 15000
) {

  try {

    await page.waitForLoadState(
      'domcontentloaded',
      { timeout }
    );

  } catch {}

  // espera body
  await page.waitForSelector('body', {
    timeout,
  }).catch(() => {});

  // espera JS legado
  await page.waitForTimeout(1500);

  // espera jquery se existir
  await page.waitForFunction(() => {

    if (!window.jQuery) {
      return true;
    }

    return jQuery.active === 0;

  }, {
    timeout: 5000,
  }).catch(() => {});
}

async function waitForReportFilterPage(
  page,
  timeout = 15000
) {

  try {

    await page.waitForFunction(() => {

      const hasButton =
        !!document.querySelector('#btFiltrar');

      const hasForm =
        !!document.querySelector(
          'form[action*="findContasRecebidas.do"]'
        );

      return hasButton && hasForm;

    }, {
      timeout,
    });

    return true;

  } catch {

    return false;
  }
}

async function clickFirstVisible(label, locators) {
  for (const locator of locators) {
    try {
      const count = await locator.count();
      if (!count) continue;
      const first = locator.first();
      if (!(await first.isVisible({ timeout: 1500 }).catch(() => false))) continue;
      await first.click({ timeout: 8000 });
      logEvent({ level: 'info', message: `${label}: clique realizado.` });
      return true;
    } catch (err) {
      logEvent({ level: 'debug', message: `${label}: tentativa falhou.`, detail: String(err.message || err) });
    }
  }
  return false;
}
async function fillInputHuman(locator, value) {
  await locator.click({ timeout: 6000 }).catch(() => {});
  await locator.fill('').catch(() => {});
  await locator.type(String(value), { delay: 20 }).catch(async () => { await locator.fill(String(value)); });
  await locator.press('Tab').catch(() => {});
}

async function setSelectOptionByLabel(page, selector, labelOrValue) {
  const locator = page.locator(selector).first();
  const desired = String(labelOrValue);
  try {
    await locator.selectOption({ label: desired });
    return true;
  } catch {}
  try {
    await locator.selectOption({ value: desired });
    return true;
  } catch {}
  try {
    await locator.selectOption({ label: new RegExp(`^${escRe(desired)}$`, 'i') });
    return true;
  } catch {}
  return false;
}

async function setCheckboxState(page, selector, checked) {
  const locator = page.locator(selector).first();
  const isChecked = await locator.isChecked().catch(() => false);
  if (isChecked !== checked) {
    await locator.click({ timeout: 6000 });
  }
}

// =========================================================
// ORDEM
// =========================================================
async function configureOrdenacao(page, ordem) {

  if (!ordem) {
    return;
  }

  logEvent({
    level: 'info',
    message: 'Configurando ordenação.',
    ordem,
  });

  // select legado
  const select = page.locator(
    'select[name*="ordem"], select[id*="ordem"]'
  );

  if (await select.count()) {

    await select.selectOption({
      label: ordem,
    }).catch(async () => {

      await select.selectOption({
        value: ordem,
      });

    });

    return;
  }

  // radio fallback
  const radio = page.locator(
    `input[type="radio"][value="${ordem}"]`
  );

  if (await radio.count()) {
    await radio.check();
  }
}

// =========================================================
// PROCESSAR LANÇAMENTOS
// =========================================================
async function configureTipoLancamento(
  page,
  tipo
) {

  if (!tipo) {
    return;
  }

  logEvent({
    level: 'info',
    message: 'Configurando tipo de lançamento.',
    tipo,
  });

  // select
  const select = page.locator(
    'select[name*="processar"], select[id*="processar"]'
  );

  if (await select.count()) {

    await select.selectOption({
      label: tipo,
    }).catch(async () => {

      await select.selectOption({
        value: tipo,
      });

    });

    return;
  }

  // radio fallback
  const radio = page.locator(
    `input[type="radio"][value="${tipo}"]`
  );

  if (await radio.count()) {
    await radio.check();
  }
}

// =========================================================
// FLAGS
// =========================================================
async function configureFlags(page, flags = {}) {

  // =====================================================
  // IMPRIMIR PARCELAS REPARCELADAS
  // =====================================================
  if (
    typeof flags.imprimirParcelasReparceladas
    !== 'undefined'
  ) {

    const checkbox = page.locator(
      'input[type="checkbox"][name*="reparcel"]'
    );

    if (await checkbox.count()) {

      const checked =
        await checkbox.first().isChecked();

      if (
        flags.imprimirParcelasReparceladas
        && !checked
      ) {
        await checkbox.first().check();
      }

      if (
        !flags.imprimirParcelasReparceladas
        && checked
      ) {
        await checkbox.first().uncheck();
      }

      logEvent({
        level: 'info',
        message:
          'Flag imprimir parcelas reparceladas configurada.',
        enabled:
          flags.imprimirParcelasReparceladas,
      });
    }
  }
}

// =========================================================
// CONTAS CORRENTE
// =========================================================
async function selectContasCorrente(
  page,
  values = []
) {

  if (!values?.length) {
    return;
  }

  logEvent({
    level: 'info',
    message: 'Selecionando contas corrente.',
    values,
  });

  // abre popup consulta
  const searchButton = page.locator(
    'img[onclick*="ContaCorrente"], img[onclick*="contaCorrente"]'
  );

  if (!(await searchButton.count())) {

    logEvent({
      level: 'warning',
      message:
        'Botão de conta corrente não encontrado.',
    });

    return;
  }

  await searchButton.first().click();

  await page.waitForTimeout(3000);

  for (const value of values) {

    // pesquisa
    const searchField = page.locator(
      'input[type="text"]'
    ).last();

    if (await searchField.count()) {

      await searchField.fill('');

      await searchField.fill(value);

      await page.keyboard.press('Enter');

      await page.waitForTimeout(1500);
    }

    // marca checkbox
    const row = page.locator('tr', {
      hasText: value,
    });

    if (await row.count()) {

      const checkbox = row.locator(
        'input[name="rowSelect"]'
      ).first();

      await checkbox.check().catch(() => {});
    }
  }

  // selecionar
  const selectBtn = page.locator(
    '#pbSelecionar'
  );

  if (await selectBtn.count()) {
    await selectBtn.click();
  }

  await page.waitForTimeout(2000);
}

// =========================================================
// PLANO FINANCEIRO EXCEÇÃO
// =========================================================
async function selectPlanoFinanceiroExcecao(
  page,
  values = []
) {

  if (!values?.length) {
    return;
  }

  logEvent({
    level: 'info',
    message:
      'Configurando exceção plano financeiro.',
    values,
  });

  await selectPlanoFinanceiro(
    page,
    values
  );

  // marca exceção
  const checkbox = page.locator(
    'input[type="checkbox"][name*="Exc"], input[type="checkbox"][name*="exc"]'
  );

  if (await checkbox.count()) {

    const checked =
      await checkbox.first().isChecked();

    if (!checked) {

      await checkbox.first().check()
        .catch(async () => {

          await checkbox.first().click({
            force: true
          });

        });
    }

    logEvent({
      level: 'info',
      message:
        'Flag de exceção marcada.',
    });
  }
}

// =========================================================
// CONFIGURADOR PRINCIPAL
// =========================================================
async function configureReportFilters(
  page,
  report
) {

  logEvent({
    level: 'info',
    message:
      `[${report.sheetName}] Iniciando configureReportFilters`,
  });

  // =====================================================
  // ORDEM
  // =====================================================
  logEvent({
    level: 'info',
    message:
      `[${report.sheetName}] Configurando ordenação`,
  });

  await configureOrdenacao(
    page,
    report.ordem
  );

  // =====================================================
  // PROCESSAR
  // =====================================================
  logEvent({
    level: 'info',
    message:
      `[${report.sheetName}] Configurando tipo lançamento`,
  });

  await configureTipoLancamento(
    page,
    report.processarLancamentos
  );

  // =====================================================
  // FLAGS
  // =====================================================
  logEvent({
    level: 'info',
    message:
      `[${report.sheetName}] Configurando flags`,
  });

  await configureFlags(
    page,
    report.flags
  );

  // =====================================================
  // PLANO FINANCEIRO
  // =====================================================
  logEvent({
    level: 'info',
    message:
      `[${report.sheetName}] Configurando plano financeiro`,
  });

  let planosParaSelecionar =
    report.planoFinanceiro || [];

  // vazio = selecionar TODOS
  if (
    planosParaSelecionar.length === 1
    && !planosParaSelecionar[0]
  ) {

    planosParaSelecionar =
      getAllPlanoFinanceiroValues();
  }

  // remove exceções
  if (report.planoFinanceiroExcecao?.length) {

    planosParaSelecionar =
      planosParaSelecionar.filter(
        x => !report.planoFinanceiroExcecao.includes(x)
      );
  }

  await selectPlanoFinanceiro(
    page,
    planosParaSelecionar
  );

  // =====================================================
  // CONTAS
  // =====================================================
  logEvent({
    level: 'info',
    message:
      `[${report.sheetName}] Configurando contas corrente`,
  });

  // =====================================================
// CONTAS CORRENTE
// =====================================================
logEvent({
  level: 'info',
  message:
    `[${report.sheetName}] Configurando contas corrente`,
});

let contasCorrente =
  report.contasCorrente || [];

// vazio = todas
if (
  contasCorrente.length === 1
  && !contasCorrente[0]
) {

  contasCorrente =
    getAllContasCorrenteValues();
}

// remove exceções
if (report.contasCorrenteExcecao?.length) {

  contasCorrente =
    contasCorrente.filter(
      x => !report.contasCorrenteExcecao.includes(x)
    );
}

await selectContasCorrente(
    page,
    contasCorrente
  );

  // =====================================================
  // DOCUMENTOS
  // =====================================================
  logEvent({
    level: 'info',
    message:
      `[${report.sheetName}] Configurando documentos`,
  });

  let documentos =
    report.documentos || [];

  // vazio = todos
  if (
    documentos.length === 1
    && !documentos[0]
  ) {

    documentos =
      getAllDocumentosValues();
  }

  // remove exceções
  if (report.documentosExcecao?.length) {

    documentos =
      documentos.filter(
        x => !report.documentosExcecao.includes(x)
      );
  }

  await selectDocumentos(
    page,
    documentos
  );

  // =====================================================
  // CONDIÇÕES
  // =====================================================
  logEvent({
    level: 'info',
    message:
      `[${report.sheetName}] Configurando condições`,
  });

  let condicoes =
    report.condicoesPagamento || [];

  // vazio = todas
  if (
    condicoes.length === 1
    && !condicoes[0]
  ) {

    condicoes =
      getAllCondicoesPagamentoValues();
  }

  // remove exceções
  if (report.condicoesPagamentoExcecao?.length) {

    condicoes =
      condicoes.filter(
        x =>
          !report.condicoesPagamentoExcecao.includes(x)
      );
  }

  await selectCondicoesPagamento(
    page,
    condicoes
  );

  logEvent({
    level: 'info',
    message:
      `[${report.sheetName}] Filtros finalizados`,
  });

  await page.waitForTimeout(1500);
}


function uniqueNonEmpty(values = []) {

  return [...new Set(
    values.filter(v => v?.trim())
  )];
}

// =====================================================
// CONTAS CORRENTE
// =====================================================
function getAllContasCorrenteValues() {

  return uniqueNonEmpty(
    REPORT_DEFINITIONS.flatMap(
      r => r.contasCorrente || []
    )
  );
}

// =====================================================
// DOCUMENTOS
// =====================================================
function getAllDocumentosValues() {

  return uniqueNonEmpty(
    REPORT_DEFINITIONS.flatMap(
      r => r.documentos || []
    )
  );
}

// =====================================================
// CONDIÇÕES PAGAMENTO
// =====================================================
function getAllCondicoesPagamentoValues() {

  return uniqueNonEmpty(
    REPORT_DEFINITIONS.flatMap(
      r => r.condicoesPagamento || []
    )
  );
}

async function ensureSubmitEnabled(page) {
  const state = await page.evaluate(() => ({
    hasForm: !!document.forms[0],
    isEnableSubmit: typeof window.IS_enableSubmit === 'undefined' ? null : window.IS_enableSubmit,
  }));

  if (!state.hasForm) {
    throw new Error('Não encontrei o formulário do relatório para habilitar o submit.');
  }

  if (state.isEnableSubmit !== true) {
    logEvent({
      level: 'debug',
      message: 'IS_enableSubmit estava desativado; reabilitando antes do submit do relatório.',
      currentState: state.isEnableSubmit,
    });
    await page.evaluate(() => {
      window.IS_enableSubmit = true;
    });
  }
}

async function setTextValue(page, selector, value) {
  const locator = page.locator(selector).first();
  await fillInputHuman(locator, value);
}

async function setValueFast(page, selector, value) {
  const locator = page.locator(selector).first();
  const count = await locator.count().catch(() => 0);
  if (!count) throw new Error(`Não encontrei o campo ${selector}.`);
  await locator.evaluate((el, nextValue) => {
    const input = el;
    input.focus();
    input.value = String(nextValue);
    input.dispatchEvent(new Event('input', { bubbles: true, cancelable: true }));
    input.dispatchEvent(new Event('change', { bubbles: true, cancelable: true }));
    input.dispatchEvent(new Event('blur', { bubbles: true, cancelable: true }));
  }, value);
}

async function setSelectFast(page, selector, labelOrValue) {
  const locator = page.locator(selector).first();
  const count = await locator.count().catch(() => 0);
  if (!count) throw new Error(`Não encontrei o select ${selector}.`);
  const desired = String(labelOrValue);
  const valueMap = {
    '#flOrdem': {
      'Data de recebimento': 'D',
      'Cliente': 'CLI',
      'Valor': 'VL',
      'Centro de custo': 'E',
      'Apropriação financeira': 'A',
      'Apropriaçao financeira': 'A',
      'Portador': 'P',
      'Conta corrente': 'CCO',
    },
    '#flTipoSelecao': {
      'Ambos': 'A',
      'Contas a receber': 'CR',
      'Caixa e bancos': 'CX',
    },
    '#flColuna': {
      'Padrão': 'P',
      'Variação monetária': 'V',
      'Juros embutidos': 'J',
    },
  };
  const mappedValue = valueMap[selector]?.[desired] || desired;
  const ok = await locator.evaluate((el, nextValue) => {
    const select = el;
    const target = String(nextValue);
    const exists = Array.from(select.options).some(opt => String(opt.value) === target);
    if (!exists) return false;
    select.value = target;
    return select.value === target;
  }, mappedValue);
  if (!ok) throw new Error(`Não consegui selecionar "${desired}" em ${selector}.`);
}

async function addMultiSelectValue(page, selector, value) {
  const locator = page.locator(selector).first();
  await locator.click({ timeout: 6000 }).catch(() => {});
  await locator.fill('').catch(() => {});
  await locator.type(String(value), { delay: 15 }).catch(async () => {
    await locator.fill(String(value));
  });
  await locator.press('Enter').catch(() => {});
  await locator.press('Tab').catch(() => {});
  await page.waitForTimeout(400);
}

async function addMultiSelectValues(page, selector, values) {
  for (const value of values || []) {
    if (!value) continue;
    await addMultiSelectValue(page, selector, value);
  }
}

function reportBySheetName(sheetName) {
  return REPORT_DEFINITIONS.find(report => report.sheetName === sheetName);
}

async function waitForContextResponse(context, predicate, timeout = 60000) {
  return await new Promise((resolve, reject) => {
    let finished = false;
    const timer = setTimeout(() => {
      if (finished) return;
      finished = true;
      context.off('response', onResponse);
      reject(new Error('Timeout aguardando resposta da requisição do relatório.'));
    }, timeout);

    const onResponse = (response) => {
      if (finished) return;
      try {
        if (predicate(response)) {
          finished = true;
          clearTimeout(timer);
          context.off('response', onResponse);
          resolve(response);
        }
      } catch {}
    };

    context.on('response', onResponse);
  });
}

async function saveReportFromContext(
  context,
  reportUrl,
  filePath
) {

  await ensureDir(path.dirname(filePath));

  const response = await context.request.get(
    reportUrl,
    {
      timeout: 120000,
    }
  );

  if (!response.ok()) {

    throw new Error(
      `Falha ao baixar relatório: HTTP ${response.status()}`
    );
  }

  const body = await response.body();

  if (fs.existsSync(filePath)) {
    await fs.promises.unlink(filePath).catch(() => {});
  }

  await fs.promises.writeFile(
    filePath,
    body
  );

  return filePath;
}

async function waitLegacyAjax(page, timeout = 15000) {

  await page.waitForFunction(() => {

    if (typeof Ajax !== 'undefined') {
      return Ajax.activeRequestCount === 0;
    }

    return true;

  }, {
    timeout,
  }).catch(() => {});
}

async function ensureExpandedFilters(page) {

  const field = page.locator('#deTipoCondicao');

  if (await field.isVisible().catch(() => false)) {
    return;
  }

  await page.locator('[name="toggleFiltro"]').click();

  await page.waitForTimeout(1200);

  await field.waitFor({
    state: 'visible',
    timeout: 10000,
  });
}

function resolveUrlAgainstBase(rawUrl) {
  if (!rawUrl) return null;
  try {
    return new URL(rawUrl, BASE_URL).toString();
  } catch {
    return null;
  }
}

function extractPopupReportUrl(rawUrl) {
  if (!rawUrl) return null;

  try {
    const parsed = new URL(rawUrl, BASE_URL);

    if (/viewReportSPW\.do/i.test(parsed.href)) {
      return parsed.href;
    }

    const embedded = parsed.searchParams.get('url');
    if (embedded) {
      const resolved = resolveUrlAgainstBase(decodeURIComponent(embedded));
      if (resolved && /viewReportSPW\.do/i.test(resolved)) {
        return resolved;
      }
    }
  } catch {
    const resolved = resolveUrlAgainstBase(rawUrl);
    if (resolved && /viewReportSPW\.do/i.test(resolved)) {
      return resolved;
    }
  }

  return null;
}

function isReportResponse(response) {
  const url = response.url();
  const headers = response.headers();
  const contentType = String(headers['content-type'] || '').toLowerCase();
  const contentDisposition = String(headers['content-disposition'] || '').toLowerCase();

  if (/viewReportSPW\.do/i.test(url)) {
    return !contentType.includes('text/html');
  }

  return (
    contentType.includes('application/pdf') ||
    contentType.includes('application/octet-stream') ||
    contentDisposition.includes('attachment') ||
    contentDisposition.includes('.pdf')
  );
}

async function waitForPopupReportUrl(popup, timeout = 30000) {
  const start = Date.now();
  while (Date.now() - start < timeout) {
    const candidates = [popup.url(), ...popup.frames().map(frame => frame.url())];
    for (const candidate of candidates) {
      const resolved = extractPopupReportUrl(candidate);
      if (resolved) {
        return resolved;
      }
    }
    await popup.waitForTimeout(500).catch(() => {});
  }
  return null;
}

async function openReportsPage(page) {

  logEvent({
    level: 'info',
    message: 'Validando tela de Relatório de Contas Recebidas.',
  });

  // já está na página?
  const alreadyThere =
    await waitForReportFilterPage(page, 5000);

  if (alreadyThere) {

    logEvent({
      level: 'info',
      message:
        'Tela de relatório já estava aberta.',
      url: page.url(),
    });

    return true;
  }

  // fallback
  logEvent({
    level: 'warning',
    message:
      'Tela não detectada. Tentando abrir novamente.',
  });

  await page.goto(
    REPORT_FILTER_PAGE_URL,
    {
      waitUntil: 'domcontentloaded',
    }
  );

  await waitForAppReady(page, 15000);

  const ok =
    await waitForReportFilterPage(page, 12000);

  if (!ok) {

    throw new Error(
      'Não foi possível abrir a tela de filtros do relatório.'
    );
  }

  return true;
}

async function goDirectToReportsPageAfterLogin(page) {
  logEvent({ level: 'info', message: 'Após login válido, tentando ir direto para a tela de relatório de contas recebidas.' });

  await page.waitForTimeout(2500);
  await dismissHomeOverlays(page);
  await page.waitForTimeout(1200);

  const attempts = [
    async () => {
      await page.goto(REPORT_FILTER_PAGE_URL, { waitUntil: 'domcontentloaded' });
      await waitForAppReady(page, 12000);
      return await waitForReportFilterPage(page, 10000);
    },
    async () => {
      await page.goto(`${BASE_URL}/sienge/8/index.html`, { waitUntil: 'domcontentloaded' });
      await waitForAppReady(page, 12000);
      await page.goto(REPORT_FILTER_PAGE_URL, { waitUntil: 'domcontentloaded' });
      await waitForAppReady(page, 12000);
      return await waitForReportFilterPage(page, 10000);
    },
  ];

  for (let i = 0; i < attempts.length; i++) {
    try {
      const ok = await attempts[i]();
      if (ok) {
        const summary = await pageSummary(page);
        logEvent({
          level: 'info',
          message: 'Redirecionamento pós-login para a tela de relatório funcionou.',
          strategy: i + 1,
          url: page.url(),
          title: summary.title,
          bodySnippet: summary.bodySnippet,
        });
        return true;
      }
      logEvent({
        level: 'warning',
        message: 'Tentativa de redirecionamento pós-login não abriu a tela real de relatório.',
        strategy: i + 1,
        actualUrl: page.url(),
      });
    } catch (err) {
      logEvent({
        level: 'warning',
        message: 'Tentativa de redirecionamento pós-login do relatório falhou.',
        strategy: i + 1,
        detail: String(err.message || err),
        actualUrl: page.url(),
      });
    }
  }

  return false;
}

async function selectPlanoFinanceiro(
  page,
  values = []
) {

  if (!values?.length || !values[0]) {
    return;
  }

  logEvent({
    level: 'info',
    message: 'Selecionando Plano Financeiro.',
    values,
  });

  const selector = '#nmConta';

  await page.waitForSelector(
    selector,
    {
      timeout: 15000,
    }
  );

  for (const value of values) {

    logEvent({
      level: 'info',
      message: 'Pesquisando plano financeiro.',
      value,
    });

    // limpa
    await page.fill(
      selector,
      ''
    );

    // dispara limpeza legacy
    await page.$eval(
      selector,
      el => {

        el.dispatchEvent(
          new Event('input', {
            bubbles: true
          })
        );

        el.dispatchEvent(
          new Event('change', {
            bubbles: true
          })
        );

      }
    );

    await page.waitForTimeout(300);

    // preenche
    await page.fill(
      selector,
      value
    );

    // dispara onchange legacy
    await page.$eval(
      selector,
      el => {

        el.dispatchEvent(
          new Event('input', {
            bubbles: true
          })
        );

        el.dispatchEvent(
          new Event('change', {
            bubbles: true
          })
        );

        el.dispatchEvent(
          new Event('blur', {
            bubbles: true
          })
        );

      }
    );

    // espera ajax autocomplete
    await page.waitForTimeout(2500);

    // autocomplete legacy
    await page.press(
      selector,
      'ArrowDown'
    ).catch(() => {});

    await page.waitForTimeout(500);

    await page.press(
      selector,
      'Enter'
    ).catch(() => {});

    await page.waitForTimeout(1500);
  }

  logEvent({
    level: 'info',
    message: 'Plano Financeiro configurado.',
  });
}

async function selectDocumentos(
  page,
  values = []
) {

  if (!values?.length || !values[0]) {
    return;
  }

  logEvent({
    level: 'info',
    message: 'Selecionando documentos.',
    values,
  });

  // garante filtros expandidos
  const toggle = page.locator(
    '[name="toggleFiltro"]'
  );

  if (await toggle.count()) {

    await toggle.click()
      .catch(() => {});
  }

  await page.waitForTimeout(1000);

  // campo descrição documento
  const input = page.locator('#nmDocumento');

  for (const value of values) {

    logEvent({
      level: 'info',
      message: 'Selecionando documento',
      value,
    });

    await fillLegacyAjaxInput(
      page,
      '#nmDocumento',
      value
    );
  }

  logEvent({
    level: 'info',
    message: 'Documentos configurados.',
  });
}

async function selectCondicoesPagamento(
  page,
  values = []
) {

  if (!values?.length || !values[0]) {
    return;
  }

  logEvent({
    level: 'info',
    message:
      'Selecionando condições de pagamento.',
    values,
  });

  // garante filtros expandidos
  const toggle = page.locator(
    '[name="toggleFiltro"]'
  );

  if (await toggle.count()) {

    await toggle.click()
      .catch(() => {});
  }

  await page.waitForTimeout(1000);

  // campo descrição condição
  const input =
  await page.locator('#deTipoCondicao')
    .isHidden()
    .catch(() => true);

  if (input) {
    await toggle.click();
  }

  for (const value of values) {

    logEvent({
      level: 'info',
      message:
        'Selecionando condição de pagamento',
      value,
    });

    await fillLegacyAjaxInput(
      page,
      '#deTipoCondicao',
      value
    );
  }

  logEvent({
    level: 'info',
    message:
      'Condições de pagamento configuradas.',
  });
}

async function fillLegacyAjaxInput(
  page,
  selector,
  value
) {

  await page.waitForSelector(selector, {
    timeout: 15000,
  });

  // limpa
  await page.fill(selector, '');

  // dispara change vazio
  await page.$eval(selector, el => {

    el.dispatchEvent(
      new Event('change', {
        bubbles: true
      })
    );

  });

  // preenche
  await page.fill(selector, value);

  // eventos legacy
  await page.$eval(selector, el => {

    el.dispatchEvent(
      new Event('input', {
        bubbles: true
      })
    );

    el.dispatchEvent(
      new Event('change', {
        bubbles: true
      })
    );

    el.dispatchEvent(
      new Event('blur', {
        bubbles: true
      })
    );

  });

  // espera ajax
  await page.waitForTimeout(2500);

  // autocomplete legacy
  await page.press(selector, 'ArrowDown')
    .catch(() => {});

  await page.waitForTimeout(500);

  await page.press(selector, 'Enter')
    .catch(() => {});

  await page.waitForTimeout(1200);
}

async function setInputSelectState(page, config) {
  const {
    inputName,
    entityName,
    values,
    hiddenMappings = {},
  } = config;

  const entries = Array.isArray(values) ? values : [values];

  const selectedList = entries
    .map((entry, idx) => {
      const parts = Object.entries(entry)
        .map(([k, v]) => `${entityName}[${idx}].${k}=${v}`);

      return `DivID=div${inputName}${idx + 1};^${parts.join('^')}$`;
    })
    .join('|');

  await page.evaluate((payload) => {
    const setValue = (id, value) => {
      const el = document.getElementById(id);

      if (!el) return;

      el.value = String(value);

      el.dispatchEvent(new Event('input', { bubbles: true }));
      el.dispatchEvent(new Event('change', { bubbles: true }));
      el.dispatchEvent(new Event('blur', { bubbles: true }));
    };

    for (const [field, value] of Object.entries(payload.hiddenMappings)) {
      setValue(field, value);
    }

    setValue(`${payload.inputName}SelectedEntitiesList`, payload.selectedList);
    setValue(`contador${payload.inputName}`, payload.entries.length);
    setValue(`contadorMaior${payload.inputName}`, payload.entries.length - 1);

  }, {
    inputName,
    selectedList,
    entries,
    hiddenMappings,
  });
}

function extractFinalReportUrl(reportUrl) {

  if (!reportUrl) {
    return null;
  }

  try {

    let current = reportUrl;

    // resolve até 5 níveis
    for (let i = 0; i < 5; i++) {

      const parsed = new URL(current);

      // =====================================================
      // URL INTERNA
      // =====================================================
      const innerUrl =
        parsed.searchParams.get('url');

      if (innerUrl) {

        current = new URL(
          decodeURIComponent(innerUrl),
          parsed.origin
        ).toString();

        continue;
      }

      // =====================================================
      // VIEW REPORT
      // =====================================================
      if (
        current.includes('viewReportSPW.do')
      ) {

        return current;
      }

      // =====================================================
      // PDF DIRETO
      // =====================================================
      if (
        current.includes('.pdf')
      ) {

        return current;
      }

      // =====================================================
      // REPORT SERVLET
      // =====================================================
      if (
        current.includes('gerarRelatorio') ||
        current.includes('report') ||
        current.includes('viewer')
      ) {

        return current;
      }

      return current;
    }

    return current;

  } catch (err) {

    logEvent({
      level: 'warning',
      message:
        'Falha ao extrair URL final do relatório.',
      reportUrl,
      error: String(err),
    });

    return reportUrl;
  }
}

async function waitForEmbeddedReport(page, timeout = 30000) {
  const start = Date.now();

  while (Date.now() - start < timeout) {

    // iframe do layer
    const iframe = page.locator('iframe');

    if (await iframe.count().catch(() => 0)) {
      const src = await iframe.first().getAttribute('src').catch(() => '');

      if (src && /viewReportSPW|please_wait_frame/i.test(src)) {
        return {
          type: 'iframe',
          url: resolveUrlAgainstBase(src),
        };
      }
    }

    // embed PDF
    const embed = page.locator('embed, object');

    if (await embed.count().catch(() => 0)) {
      const src =
        await embed.first().getAttribute('src').catch(() => '') ||
        await embed.first().getAttribute('data').catch(() => '');

      if (src) {
        return {
          type: 'embed',
          url: resolveUrlAgainstBase(src),
        };
      }
    }

    // frame runtime
    for (const frame of page.frames()) {
      const url = frame.url();

      if (/viewReportSPW|please_wait_frame/i.test(url)) {
        return {
          type: 'frame-runtime',
          url,
        };
      }
    }

    await page.waitForTimeout(500);
  }

  return null;
}

async function generateSingleReport(context, page, report) {

  logEvent({
    level: 'info',
    message: 'Entrou em generateSingleReport',
    report: report.sheetName,
  });

  await openReportsPage(page);

  logEvent({
    level: 'info',
    message: 'Depois openReportsPage',
  });

  await configureReportFilters(
    page,
    report
  );

  logEvent({
    level: 'info',
    message: 'Depois configureReportFilters',
  });

  logEvent({
    level: 'info',
    message: `Relatório "${report.sheetName}": filtros configurados.`,
  });

  await ensureSubmitEnabled(page);

  const finalPdfPath = path.join(
    REPORT_OUTPUT_DIR,
    `${sanitizeFileName(report.pdfName)}.pdf`
  );

  // =========================================================
  // LISTENERS ANTES DO CLICK
  // =========================================================
  const popupPromise = context.waitForEvent('page', {
    timeout: 30000,
  }).catch(() => null);

  const embeddedPromise = waitForEmbeddedReport(
    page,
    30000
  ).catch(() => null);

  // =========================================================
  // BOTÃO
  // =========================================================
  const button = page.locator('#btFiltrar');

  await button.waitFor({
    state: 'attached',
    timeout: 30000,
  });

  // remove disabled legado
  await page.evaluate(() => {
    const btn = document.getElementById('btFiltrar');

    if (btn) {
      btn.disabled = false;
      btn.removeAttribute('disabled');
    }
  });

  // =========================================================
  // CLICK
  // =========================================================
  try {

    // scroll
    await button.scrollIntoViewIfNeeded()
      .catch(() => {});

    // garante habilitado
    await page.evaluate(() => {

      const btn =
        document.getElementById('btFiltrar');

      if (!btn) {
        return;
      }

      btn.disabled = false;
      btn.removeAttribute('disabled');

    });

    // click REAL
    await Promise.race([

      button.click({
        force: true,
        timeout: 10000,
      }),

      page.keyboard.press('Enter')
        .catch(() => {}),

    ]);

  } catch (err) {

    logEvent({
      level: 'warning',
      message:
        'Falha click Playwright. Tentando DOM click.',
      error: String(err),
    });

    // fallback REAL
    await page.evaluate(() => {

      const btn =
        document.getElementById('btFiltrar');

      if (!btn) {
        throw new Error(
          'Botão btFiltrar não encontrado.'
        );
      }

      // click REAL DOM
      btn.click();

    });
  }

  // =========================================================
  // ESPERA NAVEGAÇÃO LEGADA
  // =========================================================
  await Promise.race([

    page.waitForNavigation({
      waitUntil: 'domcontentloaded',
      timeout: 15000,
    }).catch(() => null),

    page.waitForLoadState(
      'networkidle',
      { timeout: 15000 }
    ).catch(() => null),

    page.waitForTimeout(5000),

  ]);

  logEvent({
    level: 'info',
    message: `Relatório "${report.sheetName}": aguardando relatório.`,
  });

  // =========================================================
  // POPUP / EMBED
  // =========================================================
  const [popup, embedded] = await Promise.all([
    popupPromise,
    embeddedPromise,
  ]);

  let reportUrl = null;

  // =====================================================
  // POPUP NOVO
  // =====================================================
  if (popup) {

    try {

      await popup.waitForLoadState(
        'domcontentloaded',
        {
          timeout: 30000,
        }
      );

      reportUrl =
        await waitForPopupReportUrl(
          popup,
          30000
        );

      logEvent({
        level: 'info',
        message:
          `Relatório "${report.sheetName}": popup detectado.`,
        popupUrl: popup.url(),
        resolvedUrl: reportUrl || '',
      });

    } catch (err) {

      logEvent({
        level: 'warning',
        message:
          `Erro popup relatório "${report.sheetName}".`,
        error: String(err),
      });
    }
  }

  // =====================================================
  // EMBEDDED
  // =====================================================
  if (!reportUrl && embedded?.url) {

    reportUrl = embedded.url;

    logEvent({
      level: 'info',
      message:
        `Relatório "${report.sheetName}": embedded detectado.`,
      embeddedUrl: embedded.url,
    });
  }

  // =====================================================
  // PROCURA EM TODAS AS PÁGINAS
  // =====================================================
  if (!reportUrl) {

    const allPages = context.pages();

    for (const p of allPages) {

      try {

        const url = p.url();

        if (
          url.includes('viewReport') ||
          url.includes('.pdf') ||
          url.includes('gerarRelatorio') ||
          url.includes('report')
        ) {

          reportUrl = url;

          logEvent({
            level: 'info',
            message:
              `URL de relatório encontrada em página existente.`,
            url,
          });

          break;
        }

        // frames
        for (const frame of p.frames()) {

          const frameUrl = frame.url();

          if (
            frameUrl.includes('viewReport') ||
            frameUrl.includes('.pdf') ||
            frameUrl.includes('report')
          ) {

            reportUrl = frameUrl;

            logEvent({
              level: 'info',
              message:
                `URL de relatório encontrada em frame.`,
              frameUrl,
            });

            break;
          }
        }

        if (reportUrl) {
          break;
        }

      } catch {}
    }
  }

  // =========================================================
  // POPUP
  // =========================================================
  if (popup) {

    try {

      await popup.waitForLoadState('domcontentloaded', {
        timeout: 30000,
      });

      reportUrl = await waitForPopupReportUrl(
        popup,
        30000
      );

      logEvent({
        level: 'info',
        message: `Relatório "${report.sheetName}": popup detectado.`,
        popupUrl: popup.url(),
        popupFrames: popup.frames().map(f => f.url()),
        resolvedUrl: reportUrl || '',
      });

    } catch (err) {

      logEvent({
        level: 'warning',
        message: `Erro ao processar popup do relatório "${report.sheetName}".`,
        error: String(err),
      });
    }
  }

  // =========================================================
  // EMBEDDED
  // =========================================================
  if (!reportUrl && embedded?.url) {

    reportUrl = embedded.url;

    logEvent({
      level: 'info',
      message: `Relatório "${report.sheetName}": relatório embedded detectado.`,
      embeddedUrl: embedded.url,
    });
  }

  // =========================================================
  // URL FINAL
  // =========================================================
  if (reportUrl) {
    reportUrl = extractFinalReportUrl(reportUrl);
  }

  console.log(
    'Popup detectado:',
    !!popup,
    'Report URL:',
    reportUrl || 'N/A'
  );

  // =========================================================
  // DOWNLOAD
  // =========================================================
  if (!reportUrl) {

    // fallback:
    // às vezes o relatório abriu na mesma página
    // sem popup nem iframe

    const currentUrl = page.url();

    if (
      currentUrl.includes('viewReport') ||
      currentUrl.includes('report') ||
      currentUrl.includes('viewer')
    ) {

      reportUrl = currentUrl;

      logEvent({
        level: 'warning',
        message:
          'Usando URL atual da página como fallback de relatório.',
        currentUrl,
      });
    }
  }

  if (!reportUrl) {

    // tenta iframe final
    const iframeSrc = await page.evaluate(() => {

      const iframe =
        document.querySelector('iframe');

      return iframe?.src || null;

    }).catch(() => null);

    if (iframeSrc) {

      reportUrl = iframeSrc;

      logEvent({
        level: 'warning',
        message:
          'Usando iframe src como fallback de relatório.',
        iframeSrc,
      });
    }
  }

  if (!reportUrl) {

    throw new Error(
      `Não foi possível localizar a URL do relatório "${report.sheetName}".`
    );
  }

  logEvent({
    level: 'info',
    message: `Baixando relatório "${report.sheetName}".`,
    reportUrl,
  });

  await saveReportFromContext(
    context,
    reportUrl,
    finalPdfPath
  );

  logEvent({
    level: 'info',
    message: `Relatório "${report.sheetName}" salvo com sucesso.`,
    pdfPath: finalPdfPath,
  });

  // =========================================================
  // FINALIZAÇÃO
  // =========================================================
  await page.waitForTimeout(2000);

  await logPageState(
    page,
    `Relatório "${report.sheetName}" processado.`,
    {
      shotName: `report-${sanitizeFileName(report.sheetName)}`,
    }
  );
}

async function runReports(context, page) {
  await ensureDir(REPORT_OUTPUT_DIR);
  page.setDefaultTimeout(15000);
  page.setDefaultNavigationTimeout(30000);
  const reports = REPORT_DEFINITIONS.map(report => ({ ...report }));
  logEvent({
    level: 'info',
    message: 'Iniciando geração de relatórios.',
    reports: reports.map(report => report.sheetName),
    outputDir: REPORT_OUTPUT_DIR,
  });

  await goDirectToReportsPageAfterLogin(page);

  for (const report of reports) {

    logEvent({
      level: 'info',
      message: `Iniciando relatório`,
      report: report.sheetName,
    });

    try {

      logEvent({
        level: 'info',
        message: 'Antes generateSingleReport',
      });

      await generateSingleReport(
        context,
        page,
        report
      );

      logEvent({
        level: 'info',
        message: 'Depois generateSingleReport',
      });

    } catch (err) {

      console.log(err);

      await logPageState(
        page,
        `Falha ao gerar o relatório "${report.sheetName}".`,
        {
          level: 'error',
          shotName: `report-error-${sanitizeFileName(report.sheetName)}`,
          report: report.sheetName,
          detail: String(err.message || err),
        }
      );

      throw err;
    }
  }

  logEvent({ level: 'info', message: 'Geração de relatórios concluída.' });
}

function isLoggedArea(summary) {
  const t = `${summary.title} ${summary.bodySnippet}`;
  return (
    /sienge plataforma|base:\s|código:\s7453|codigo:\s7453|pesquise uma funcionalidade|financeiro|contas a pagar|autorização de pagamento|autorizacao de pagamento/i.test(t)
    && !/entrar com sienge id|verifique o código no seu e-mail|selecione o método de autenticação em duas etapas|escolha uma conta/i.test(t)
  );
}
function isWelcomeGate(summary) {
  const t = `${summary.title} ${summary.bodySnippet}`;
  return /bem-vindo!|entrar com sienge id/i.test(t) && !/digite a sua senha|verifique o código no seu e-mail|selecione o método de autenticação em duas etapas/i.test(t);
}

function isUserAlreadyLoggedScreen(summary) {
  const t = `${summary.title} ${summary.bodySnippet}`;
  return /usuário logado|usuario logado|já está conectado ao sistema|ja esta conectado ao sistema|prosseguir cancelar/i.test(t);
}


async function waitForAuthPageDom(page, timeout = 12000) {
  const start = Date.now();
  while (Date.now() - start < timeout) {
    const hasDtInicio = await page.locator('input#dtInicio, input[name="dtInicio"]').count().catch(() => 0);
    const hasDtFim = await page.locator('input#dtFim, input[name="dtFim"]').count().catch(() => 0);
    const hasHeading = await page.getByText(/AUTORIZAÇÃO DE PAGAMENTO|AUTORIZACAO DE PAGAMENTO/i).count().catch(() => 0);
    if ((hasDtInicio && hasDtFim) || hasHeading) return true;
    await page.waitForTimeout(500);
  }
  return false;
}

async function tryOpenAuthFromHome(page) {
  return await clickRecentAuthorizationLink(page);
}

async function handleUserAlreadyLoggedScreen(page) {
  const summary = await pageSummary(page);
  if (!isUserAlreadyLoggedScreen(summary)) return false;

  logEvent({
    level: 'warning',
    message: 'Tela de usuário já logado detectada. Clicando em "Prosseguir".',
    url: summary.url,
    title: summary.title,
  });

  const clicked = await clickFirstVisible('Botão Prosseguir', [
    page.getByRole('link', { name: /prosseguir/i }),
    page.getByText(/^Prosseguir$/i),
    page.locator('a.Button-prim[href*="acao=S"]'),
    page.locator('a').filter({ hasText: /^Prosseguir$/i }),
  ]);

  if (!clicked) {
    await logPageState(page, 'Não consegui clicar em "Prosseguir" na tela de usuário já logado.', {
      level: 'error',
      shotName: 'already-logged-proceed-failed',
    });
    throw new Error('Não consegui clicar em "Prosseguir" na tela de usuário já logado.');
  }

  await waitForAppReady(page, 15000);
  await logPageState(page, 'Clique em "Prosseguir" executado na tela de usuário já logado.', {
    shotName: 'after-prosseguir-already-logged',
  });
  return true;
}
function isAccountChooser(summary) {
  const t = `${summary.title} ${summary.bodySnippet}`;
  return /escolha uma conta|usar outra conta|conectado/i.test(t);
}
function isPasswordStep(summary) {
  const t = `${summary.title} ${summary.bodySnippet}`;
  return /digite a sua senha|esqueci a minha senha|email pr[eé]-preenchido|email pré-preenchido/i.test(t);
}
function isEmailStep(summary) {
  const t = `${summary.title} ${summary.bodySnippet}`;
  return /seu e-mail|continuar/i.test(t) && !/digite a sua senha|verifique o código no seu e-mail|selecione o método/i.test(t);
}
function isMfaMethodStep(summary) {
  const t = `${summary.title} ${summary.bodySnippet}`;
  return /selecione o método de autenticação em duas etapas|autenticação por e-mail|autenticacao por e-mail|aplicativo de autenticação|aplicativo de autenticacao/i.test(t)
    && !/verifique o código no seu e-mail/i.test(t);
}
function isMfaCodeStep(summary) {
  const t = `${summary.title} ${summary.bodySnippet}`;
  return /verifique o código no seu e-mail|enviamos um código de 6 dígitos|enviamos um código de 6 digitos|insira-o para continuar/i.test(t);
}
async function fetchMfaCodeFromEmail(timeoutMs = 120000, pollMs = 5000) {
  // Prefer Gmail API OAuth2 if configured
  try {
    const useOauth = String(process.env.SOLO_MFA_OAUTH2 || process.env.MFA_USE_OAUTH2 || 'false').toLowerCase() === 'true' || Boolean(process.env.GOOGLE_REFRESH_TOKEN);
    if (useOauth) {
      try {
        const gmailHelper = require('./tools/gmail_oauth');
        const query = process.env.MFA_GMRAW_QUERY || '"Código de Verificação" sienge';
        logEvent({ level: 'info', message: 'Tentando buscar código MFA via Gmail API OAuth2', query });
        const found = await gmailHelper.fetchMfaCodeWithGmailApi({ query, maxResults: Number(process.env.MFA_OAUTH_MAX || 10) });
        if (found && found.code) {
          logEvent({ level: 'info', message: 'Código MFA obtido via Gmail API', code: found.code, messageId: found.messageId });
          return found.code;
        }
        logEvent({ level: 'info', message: 'Nenhum código encontrado via Gmail API OAuth2' });
      } catch (e) {
        try {
          const detail = String(e && e.message || e);
          const stack = e && e.stack ? String(e.stack) : undefined;
          const props = {};
          try { Object.getOwnPropertyNames(e || {}).forEach(k => { props[k] = e[k]; }); } catch (er) {}
          logEvent({ level: 'warning', message: 'Falha ao usar Gmail API OAuth2', detail, stack, errorProps: props });
        } catch (er) {
          logEvent({ level: 'warning', message: 'Falha ao usar Gmail API OAuth2', detail: String(e && e.message || e) });
        }
      }
    }
  } catch (e) { }

  const host = process.env.MFA_IMAP_HOST;
  const port = Number(process.env.MFA_IMAP_PORT || 993);
  // allow explicit solo credentials for the MFA-reading mailbox
  const user = process.env.SOLO_MFA_USER || process.env.MFA_IMAP_USER || process.env.SIENGE_USERNAME;
  const password = process.env.SOLO_MFA_PASS || process.env.MFA_IMAP_PASS;
  const tls = String(process.env.MFA_IMAP_TLS ?? 'true').toLowerCase() !== 'false';

  if (!host || !user || !password) return null;

  loadMailDependencies();
  if (!imaps) return null;

  const config = {
    imap: {
      user,
      password,
      host,
      port,
      tls,
      authTimeout: 30000,
    },
  };

  const senderRegex = process.env.MFA_EMAIL_FROM_REGEX ? new RegExp(process.env.MFA_EMAIL_FROM_REGEX, 'i') : null;
  const subjectRegex = process.env.MFA_EMAIL_SUBJECT_REGEX ? new RegExp(process.env.MFA_EMAIL_SUBJECT_REGEX, 'i') : /c[oó]digo|codigo|verifica|verificação|verificacao|autentica/i;

  const start = Date.now();
  let connection;
  try {
    logEvent({ level: 'info', message: 'Conectando IMAP para busca de MFA.', host, port, tls });
    connection = await imaps.connect(config);
    logEvent({ level: 'info', message: 'Conexão IMAP estabelecida.' });
    const box = await connection.openBox('INBOX');
    try {
      logEvent({ level: 'info', message: 'Caixa INBOX aberta.', box: { name: box.name, messages: box.messages && box.messages.total || box.messages } });

      // Debug dump messages if requested
      if (String(process.env.MFA_DUMP_MESSAGES || 'false').toLowerCase() === 'true') {
        try {
          const fetchOptionsDump = { bodies: ['HEADER.FIELDS (FROM TO SUBJECT DATE)', 'TEXT'], struct: true, markSeen: false };
          const allMsgs = await connection.search(['ALL'], fetchOptionsDump).catch(() => []);
          logEvent({ level: 'info', message: `MFA_DUMP_MESSAGES enabled — total messages: ${allMsgs.length}` });
          const limit = Number(process.env.MFA_DUMP_LIMIT || 50);
          for (let i = 0; i < Math.min(allMsgs.length, limit); i++) {
            const msg = allMsgs[i];
            const parts = Array.isArray(msg.parts) ? msg.parts : [msg];
            const headerPart = parts.find(p => /HEADER.FIELDS/i.test(p.which)) || parts[0];
            const textPart = parts.find(p => String(p.which).toUpperCase().includes('TEXT')) || parts[0];
            const header = String(headerPart.body || '').slice(0, 1200).replace(/\r?\n/g, ' ');
            const snippet = String(textPart.body || '').replace(/\s+/g, ' ').slice(0, 1200);
            logEvent({ level: 'debug', message: 'DUMP_MSG', index: i, seqno: msg.seqno, attributes: msg.attributes, header, snippet });
          }
        } catch (e) {
          logEvent({ level: 'warning', message: 'Erro ao dump de mensagens IMAP', detail: String(e && e.message || e) });
        }
      }
    } catch (e) { logEvent({ level: 'debug', message: 'INBOX opened (could not stringify box)', detail: String(e && e.message || e) }); }
    // Try to list available mailboxes and their counts to debug servers that use non-standard INBOX names
    try {
      let boxes = null;
      if (typeof connection.getBoxes === 'function') {
        boxes = await connection.getBoxes();
      } else if (connection.imap && typeof connection.imap.getBoxes === 'function') {
        boxes = await new Promise((res, rej) => connection.imap.getBoxes((err, b) => err ? rej(err) : res(b)));
      }
      if (boxes) {
        const names = [];
        function walk(obj, prefix = '') {
          for (const k of Object.keys(obj || {})) {
            const v = obj[k] || {};
            const name = prefix ? `${prefix}${k}` : k;
            names.push(name);
            if (v.children) walk(v.children, `${name}${v.delimiter || '.'}`);
          }
        }
        walk(boxes, '');
        logEvent({ level: 'info', message: 'Mailboxes detected', count: names.length, names: names.slice(0, 50) });

        // Try opening first few mailboxes to inspect message counts
        const tryLimit = Math.min(names.length, Number(process.env.MFA_MAILBOX_PROBE_LIMIT || 10));
        for (let i = 0; i < tryLimit; i++) {
          const nm = names[i];
          try {
            const b = await connection.openBox(nm);
            logEvent({ level: 'info', message: 'Probe openBox', mailbox: nm, messages: b.messages && b.messages.total || b.messages });
          } catch (err) {
            logEvent({ level: 'debug', message: 'Probe openBox failed', mailbox: nm, detail: String(err && err.message || err) });
          }
        }
      }
    } catch (e) {
      logEvent({ level: 'debug', message: 'Falha ao listar mailboxes', detail: String(e && e.message || e) });
    }

    // If no messages found in INBOX, also try common spam/junk folders
    const spamCandidates = [
      '[Gmail]/Spam',
      'Spam',
      'Junk',
      'Spam/Recebidos',
      'INBOX.Spam',
      'Caixa de spam',
      'Caixa de entrada/Spam',
    ];
    async function searchMailboxForMfa(boxName, fetchOptionsLocal) {
      try {
        const b = await connection.openBox(boxName);
        logEvent({ level: 'debug', message: 'Opened mailbox for spam probe', mailbox: boxName, messages: b.messages && b.messages.total || b.messages });
        const msgs = await connection.search(['ALL'], fetchOptionsLocal).catch(() => []);
        logEvent({ level: 'debug', message: `search in ${boxName} returned ${msgs.length}` });
        return msgs || [];
      } catch (err) {
        logEvent({ level: 'debug', message: 'searchMailboxForMfa failed', mailbox: boxName, detail: String(err && err.message || err) });
        return [];
      }
    }

    const sinceTime = Date.now() - Math.max(timeoutMs, 5 * 60 * 1000);
    while (Date.now() - start < timeoutMs) {
      const fetchOptions = { bodies: ['HEADER.FIELDS (FROM SUBJECT)', 'TEXT'], struct: true, markSeen: false };
      logEvent({ level: 'debug', message: 'Buscando mensagens IMAP.' , fetchOptions });
      let messages = [];

      // Gmail-optimized search using X-GM-RAW if applicable
      const isGmail = Boolean((user || '').toLowerCase().endsWith('@gmail.com') || (host || '').toLowerCase().includes('gmail'));
      const useGmRaw = String(process.env.MFA_GMAIL_USE_GMRAW ?? 'true').toLowerCase() !== 'false';
      if (isGmail && useGmRaw) {
        const gmrawQuery = process.env.MFA_GMRAW_QUERY || (process.env.MFA_EMAIL_SUBJECT_REGEX ? process.env.MFA_EMAIL_SUBJECT_REGEX : '"Código de Verificação" sienge');
        try {
          logEvent({ level: 'debug', message: 'Tentando busca Gmail X-GM-RAW', query: String(gmrawQuery).slice(0,200) });
          messages = await connection.search([['X-GM-RAW', String(gmrawQuery)]], fetchOptions).catch((e) => { logEvent({ level: 'debug', message: 'X-GM-RAW search falhou', detail: String(e && e.message || e) }); return []; });
          logEvent({ level: 'debug', message: `connection.search X-GM-RAW retornou ${messages.length}` });
        } catch (e) {
          logEvent({ level: 'debug', message: 'Erro ao executar X-GM-RAW', detail: String(e && e.message || e) });
          messages = [];
        }
      }

      // fallback to generic searches if no messages found
      if (!messages.length) {
        try {
          messages = await connection.search(['ALL'], fetchOptions);
          logEvent({ level: 'debug', message: `connection.search ALL retornou ${messages.length}` });
        } catch (e) { logEvent({ level: 'warning', message: 'Erro ao buscar ALL', detail: String(e && e.message || e) }); }
      }

      // also try common spam/junk folders if still empty
      if (!messages.length) {
        for (const spamBox of spamCandidates) {
          // for Gmail, try GM-RAW in the spam box first
          let found = [];
          if (isGmail && useGmRaw) {
            try {
              const gmrawQuery = process.env.MFA_GMRAW_QUERY || (process.env.MFA_EMAIL_SUBJECT_REGEX ? process.env.MFA_EMAIL_SUBJECT_REGEX : '"Código de Verificação" sienge');
              logEvent({ level: 'debug', message: 'Tentando X-GM-RAW dentro do spam box', mailbox: spamBox, query: String(gmrawQuery).slice(0,200) });
              // need to open box then search
              await connection.openBox(spamBox).catch(() => null);
              found = await connection.search([['X-GM-RAW', String(gmrawQuery)]], fetchOptions).catch(() => []);
              logEvent({ level: 'debug', message: `X-GM-RAW in ${spamBox} returned ${found.length}` });
            } catch (e) { logEvent({ level: 'debug', message: 'X-GM-RAW in spam failed', mailbox: spamBox, detail: String(e && e.message || e) }); found = []; }
          }

          if (!found.length) {
            found = await searchMailboxForMfa(spamBox, fetchOptions).catch(() => []);
          }

          if (found && found.length) {
            messages = found;
            logEvent({ level: 'info', message: `Mensagens encontradas em pasta de spam: ${spamBox}`, count: messages.length });
            break;
          }
        }
      }

      try {
        const unseen = await connection.search(['UNSEEN'], fetchOptions).catch(() => []);
        logEvent({ level: 'debug', message: `connection.search UNSEEN retornou ${unseen.length}` });
      } catch (e) { logEvent({ level: 'debug', message: 'Erro ao buscar UNSEEN', detail: String(e && e.message || e) }); }

      if (!messages.length) {
        // try SINCE last 3 days
        const days = Number(process.env.MFA_IMAP_SINCE_DAYS || 3);
        const since = new Date(Date.now() - days * 24 * 60 * 60 * 1000);
        const monthNames = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
        const sinceStr = `${since.getDate()}-${monthNames[since.getMonth()]}-${since.getFullYear()}`;
        try {
          const sinceMsgs = await connection.search(['SINCE', sinceStr], fetchOptions).catch(() => []);
          logEvent({ level: 'debug', message: `connection.search SINCE ${sinceStr} retornou ${sinceMsgs.length}` });
          if (sinceMsgs.length) messages = sinceMsgs;
        } catch (e) { logEvent({ level: 'debug', message: 'Erro ao buscar SINCE', detail: String(e && e.message || e) }); }
      }
      // sort messages by date descending so newest are processed first
      function getMsgTime(msg) {
        try {
          const d = msg.attributes && (msg.attributes.internalDate || msg.attributes.date) || (msg.attrs && msg.attrs.date) || msg.date || msg.internalDate;
          const t = d ? (new Date(d)).getTime() : (msg.seqno || 0);
          return isFinite(t) ? t : 0;
        } catch (e) { return 0; }
      }
      try {
        messages = Array.isArray(messages) ? messages.slice().sort((a, b) => getMsgTime(b) - getMsgTime(a)) : messages;
      } catch (e) {}
      logEvent({ level: 'debug', message: `Mensagens a processar: ${messages.length}` });
      for (const msg of messages) {
        // filter by date if available
        try {
          const msgDate = msg.attributes && (msg.attributes.date || msg.attributes.internalDate || (msg.attrs && msg.attrs.date));
          if (msgDate) {
            const d = new Date(msgDate);
            if (isFinite(d) && d.getTime() < sinceTime) {
              logEvent({ level: 'debug', message: 'Ignorando mensagem pela data', date: String(msgDate) });
              continue;
            }
          }
        } catch (e) { logEvent({ level: 'debug', message: 'Falha ao verificar data da mensagem', detail: String(e && e.message || e) }); }

        // combine possible text bodies
        let raw = '';
        try {
          if (Array.isArray(msg.parts)) {
            raw = msg.parts.map(p => (p.body || '')).join('\n');
          } else if (msg.body) {
            raw = String(msg.body || '');
          } else {
            raw = JSON.stringify(msg).slice(0, 2000);
          }
        } catch (e) {
          raw = JSON.stringify(msg).slice(0, 2000);
        }
        let parsed;
        if (mailparser && mailparser.simpleParser) {
          try { parsed = await mailparser.simpleParser(raw); } catch (e) { parsed = { text: String(raw || ''), from: {} }; }
        } else {
          parsed = { text: String(raw || ''), from: {} };
        }

        const fromText = (parsed.from && parsed.from.text) || '';
        const subject = parsed.subject || '';
        if (senderRegex && !senderRegex.test(fromText)) { logEvent({ level: 'debug', message: 'Remetente não corresponde ao filtro', from: fromText.slice(0,200) }); continue; }
        if (subjectRegex && !subjectRegex.test(subject + ' ' + (parsed.text || ''))) { logEvent({ level: 'debug', message: 'Assunto/texto não corresponde ao filtro', subject: String(subject).slice(0,200) }); continue; }

        const text = (parsed && (parsed.text || parsed.html) ? (parsed.text || parsed.html) : raw || '').toString();

        // debug log candidate
        try {
          logEvent({ level: 'debug', message: 'Candidate email for MFA', from: fromText.slice(0, 200), subject: String(subject).slice(0,200), snippet: text.replace(/\s+/g,' ').slice(0,300) });
        } catch (e) {}
        const m = text.match(/(\d{6})/);
        if (m) {
          logEvent({ level: 'info', message: 'Código MFA encontrado no e-mail', code: m[1] });
          try { await connection.addFlags(msg.attributes && msg.attributes.uid || msg.attributes && msg.attributes['UID'] || msg.seqno, '\\Seen'); } catch (e) { logEvent({ level: 'debug', message: 'Falha ao marcar mensagem como lida', detail: String(e && e.message || e) }); }
          return m[1];
        } else {
          logEvent({ level: 'debug', message: 'Nenhum código 6 dígitos encontrado neste candidato.' });
        }
      }

      await new Promise(r => setTimeout(r, pollMs));
    }
  } catch (err) {
    // log full error details for debugging
    try {
      const detail = String(err && err.message || err);
      const stack = err && err.stack ? String(err.stack) : undefined;
      const picked = {};
      try { Object.getOwnPropertyNames(err || {}).forEach(k => { picked[k] = err[k]; }); } catch (e) {}
      logEvent({ level: 'warning', message: 'Erro na rotina IMAP de busca MFA', detail, stack, errorProps: picked });
    } catch (e) {
      logEvent({ level: 'warning', message: 'Erro na rotina IMAP de busca MFA (falha ao formatar erro)', detail: String(err && err.message || err) });
    }
    // ignore and return null
  } finally {
    try { if (connection) await connection.end(); } catch (e) { logEvent({ level: 'debug', message: 'Falha ao encerrar conexão IMAP', detail: String(e && e.message || e) }); }
  }

  return null;
}

async function promptUserForCode() {
  const autoEnabled = String(process.env.MFA_AUTO_FETCH ?? 'true').toLowerCase() !== 'false';
  if (autoEnabled) {
    const timeoutMs = Number(process.env.MFA_AUTO_TIMEOUT_MS || 120000);
    logEvent({ level: 'info', message: `Tentando obter código MFA do e-mail por até ${timeoutMs}ms.` });
    // allow a small pre-wait to let the most recent MFA email arrive
    try {
      const prewait = Number(process.env.MFA_PREWAIT_MS || MFA_PREWAIT_MS || 0);
      if (prewait > 0) {
        logEvent({ level: 'info', message: `Aguardando ${prewait}ms para chegada do e-mail MFA antes de buscar.` });
        await new Promise(r => setTimeout(r, prewait));
      }
    } catch (e) {}
    try {
      const code = await fetchMfaCodeFromEmail(timeoutMs, Number(process.env.MFA_AUTO_POLL_MS || 5000));
      if (code) {
        logEvent({ level: 'info', message: 'Código MFA obtido automaticamente via e-mail.' });
        return code;
      }
      logEvent({ level: 'info', message: 'Nenhum código MFA encontrado no e-mail dentro do tempo limite.' });
    } catch (err) {
      logEvent({ level: 'warning', message: 'Falha ao tentar obter MFA por e-mail.', detail: String(err && err.message || err) });
    }
  }

  const rl = readline.createInterface({ input, output });
  try {
    while (true) {
      const ans = (await rl.question('Digite o código MFA enviado por e-mail e pressione Enter: ')).trim();
      const digits = ans.replace(/\D/g, '');
      if (digits.length >= 6) return digits.slice(0, 6);
      console.log('Código inválido. Informe 6 dígitos.');
    }
  } finally { rl.close(); }
}

async function dismissHomeOverlays(page) {
  logEvent({ level: 'debug', message: 'Iniciando dismissHomeOverlays.', currentUrl: page.url() });
  let changed = false;

  const actions = [
    ['Popup push NÃO, OBRIGADO', [
      page.getByRole('button', { name: /não,\s*obrigado/i }),
      page.locator('button').filter({ hasText: /não,\s*obrigado/i }),
      page.locator('[role="dialog"], .modal, .popup, .popover').locator('button').filter({ hasText: /não,\s*obrigado/i }),
    ]],
    ['Ajuda contextual ENTENDI', [
      page.getByRole('button', { name: /^entendi$/i }),
      page.locator('button').filter({ hasText: /^ENTENDI$/i }),
      page.locator('[class*="chameleon"], [class*="beamer"], [class*="fresh"], [role="dialog"], aside').locator('button').filter({ hasText: /^ENTENDI$/i }),
    ]],
    ['Fechar painel de ajuda contextual', [
      page.locator('aside button[aria-label*="close" i], aside button[title*="close" i]'),
      page.locator('[class*="chameleon"] button[aria-label*="close" i], [class*="beamer"] button[aria-label*="close" i]'),
      page.locator('aside svg').locator('xpath=ancestor::button[1]'),
    ]],
  ];

  for (const [label, locs] of actions) {
    const ok = await clickFirstVisible(label, locs);
    if (ok) {
      changed = true;
      await page.waitForTimeout(500);
    }
  }

  return changed;
}

async function ensurePassWelcomeGate(page) {
  const summary = await pageSummary(page);
  if (!isWelcomeGate(summary)) return false;
  logEvent({ level: 'info', message: 'Tela de boas-vindas detectada. Clicando em "ENTRAR COM SIENGE ID".' });
  const clicked = await clickFirstVisible('Botão Entrar com Sienge ID', [
    page.getByRole('button', { name: /entrar com sienge id/i }),
    page.getByText(/entrar com sienge id/i),
    page.locator('button, a, div, span').filter({ hasText: /entrar com sienge id/i }),
  ]);
  if (!clicked) throw new Error('Não consegui clicar em "ENTRAR COM SIENGE ID".');
  await waitForAppReady(page, 20000);
  await logPageState(page, 'Clique em "ENTRAR COM SIENGE ID" executado.', { shotName: 'after-enter-sienge-id' });
  return true;
}
async function tryAccountChooser(page) {
  const summary = await pageSummary(page);
  if (!isAccountChooser(summary)) return false;
  logEvent({ level: 'info', message: 'Tela de seleção de conta detectada. Tentando clicar na conta salva.' });
  const clicked = await clickFirstVisible('Conta salva Sienge ID', [
    page.locator('button').filter({ hasText: new RegExp(escRe(USERNAME), 'i') }),
    page.locator('[role="button"]').filter({ hasText: new RegExp(escRe(USERNAME), 'i') }),
    page.locator('button').filter({ hasText: /conectado/i }),
    page.getByText(/Tayane Granemann/i).locator('xpath=ancestor::button[1]'),
  ]);
  if (!clicked) throw new Error('Não consegui clicar na conta salva do Sienge ID.');
  await waitForAppReady(page, 20000);
  await logPageState(page, 'Conta salva clicada com sucesso.', { shotName: 'after-account-click' });
  return true;
}
async function detectMfaPinInputs(page) {
  const locators = [
    page.locator('input.grua-pin-input-textfield'),
    page.locator('input[maxlength="1"]'),
    page.locator('.sc-fqkvVR input'),
    // A versão atual do Sienge ID não expõe classe nem maxlength nos PINs;
    // nesta etapa os seis inputs visíveis são exatamente o código MFA.
    page.locator('input:not([type="hidden"]):not([disabled])'),
  ];
  for (const loc of locators) {
    const count = await loc.count().catch(() => 0);
    if (count >= 6) {
      const visible = [];
      for (let i = 0; i < count; i++) {
        const item = loc.nth(i);
        if (await item.isVisible({ timeout: 500 }).catch(() => false)) visible.push(item);
      }
      if (visible.length >= 6) return visible.slice(0, 6);
    }
  }
  return [];
}
async function fillMfaCode(page, code) {
  const digits = code.split('').slice(0, 6);
  const inputs = await detectMfaPinInputs(page);
  if (inputs.length >= 6) {
    logEvent({ level: 'info', message: 'Preenchendo código MFA em 6 campos separados.' });
    for (let i = 0; i < 6; i++) {
      // `type()` espera o timeout padrão do Playwright nessa implementação do
      // OTP; `fill()` dispara o mesmo evento de input e é imediato.
      await inputs[i].fill(digits[i], { timeout: 3000 });
      logEvent({ level: 'info', message: `Dígito MFA ${i + 1}/6 preenchido.` });
    }
    return true;
  }
  return false;
}

function mfaEmailMethodLocators(page) {
  return [
    page.locator('[data-testid="mfa-choose-method-email"]'),
    page.locator('[data-testid*="email" i]'),
    page.getByRole('button', { name: /autenticação por e-mail|autenticacao por e-mail/i }),
    page.getByText(/autenticação por e-mail|autenticacao por e-mail/i).locator('xpath=ancestor::*[@role="button" or self::button][1]'),
    page.locator('button, [role="button"]').filter({ hasText: /autenticação por e-mail|autenticacao por e-mail/i }),
  ];
}

async function waitForMfaStep(page, timeout = 30000) {
  const deadline = Date.now() + timeout;
  while (Date.now() < deadline) {
    const pins = await detectMfaPinInputs(page);
    if (pins.length >= 6) return { type: 'code', pins };

    for (const locator of mfaEmailMethodLocators(page)) {
      const candidate = locator.first();
      if (
        await candidate.isVisible({ timeout: 250 }).catch(() => false)
        && await candidate.isEnabled().catch(() => true)
      ) {
        return { type: 'method', locator: candidate };
      }
    }
    await page.waitForTimeout(300);
  }
  return null;
}

async function waitForMfaCodeInputs(page, timeout = 30000) {
  const deadline = Date.now() + timeout;
  while (Date.now() < deadline) {
    const pins = await detectMfaPinInputs(page);
    if (pins.length >= 6) return pins;
    await page.waitForTimeout(300);
  }
  return [];
}

async function handleMfa(page) {
  let summary = await pageSummary(page);
  if (!isMfaMethodStep(summary) && !isMfaCodeStep(summary)) return false;

  if (isMfaMethodStep(summary)) {
    // Após senha/conta, o SSO pode exibir "Verificando..." antes de montar a
    // escolha de método. Espere a opção clicável ou os PINs, sem clicar em uma
    // tela transitória.
    const step = await waitForMfaStep(page);
    if (!step) throw new Error('A tela de MFA não ficou pronta para seleção do método ou inserção do código.');

    if (step.type === 'method') {
      logEvent({ level: 'info', message: 'Tela de escolha do método MFA pronta. Selecionando "Autenticação por e-mail".' });
      await step.locator.click({ timeout: 8000 });
      const pins = await waitForMfaCodeInputs(page);
      if (!pins.length) throw new Error('A tela para inserir o código MFA não apareceu após selecionar autenticação por e-mail.');
      await logPageState(page, 'Método MFA por e-mail selecionado.', { shotName: 'after-mfa-email-method' });
      summary = { ...summary, bodySnippet: 'verifique o código no seu e-mail' };
    } else {
      summary = { ...summary, bodySnippet: 'verifique o código no seu e-mail' };
    }
  }

  if (!isMfaCodeStep(summary)) {
    await page.waitForTimeout(1500);
    summary = await pageSummary(page);
  }
  if (!isMfaCodeStep(summary)) throw new Error('A tela para inserir o código MFA não apareceu.');

  const readyPins = await waitForMfaCodeInputs(page);
  if (!readyPins.length) throw new Error('Os campos do código MFA não ficaram disponíveis.');

  logEvent({ level: 'info', message: 'Tela de código MFA detectada. Aguardando o código digitado no terminal.' });
  const code = await promptUserForCode();
  const ok = await fillMfaCode(page, code);
  if (!ok) throw new Error('Não encontrei o input para inserir o código MFA.');

  // O Sienge ID inicia a validação automaticamente no sexto dígito. Não há
  // botão a clicar e os inputs podem permanecer no DOM durante a transição.
  logEvent({ level: 'info', message: 'Código MFA preenchido. Aguardando redirecionamento automático do Sienge ID.' });
  await page.waitForTimeout(800);
  await waitForAppReady(page, 25000);
  logEvent({ level: 'info', message: 'Etapa automática do código MFA concluída; continuando o login.' });
  await logPageState(page, 'Código MFA informado.', { shotName: 'after-mfa-code-submit' });
  return true;
}

async function trySsoLogin(page) {
  let summary = await pageSummary(page);

  if (isAccountChooser(summary)) {
    await tryAccountChooser(page);
    summary = await pageSummary(page);
  }
  if (isEmailStep(summary)) {
    logEvent({ level: 'info', message: 'Etapa de e-mail do SSO detectada.' });
    const candidates = [
      page.locator('input[name="username"]'),
      page.locator('input[name="email"]'),
      page.locator('input[type="email"]'),
      page.locator('input[placeholder*="e-mail" i]'),
      page.locator('input[placeholder*="email" i]'),
      page.locator('input[type="text"]'),
    ];
    let userInput = null;
    for (const loc of candidates) if (await loc.count()) { userInput = loc.first(); break; }
    if (userInput) {
      await fillInputHuman(userInput, USERNAME);
      const clicked = await clickFirstVisible('Continuar SSO', [
        page.getByRole('button', { name: /continuar|próximo|proximo|entrar/i }),
        page.getByText(/continuar/i),
        page.locator('button[type="submit"]'),
      ]);
      if (!clicked) await userInput.press('Enter').catch(() => {});
      await waitForAppReady(page, 20000);
      await logPageState(page, 'Etapa de e-mail do SSO executada.', { shotName: 'after-sso-email' });
      summary = await pageSummary(page);
    }
  }
  if (isAccountChooser(summary)) {
    await tryAccountChooser(page);
    summary = await pageSummary(page);
  }
  if (isPasswordStep(summary)) {
    logEvent({ level: 'info', message: 'Etapa de senha do SSO detectada.' });
    const candidates = [
      page.locator('input[name="password"]'),
      page.locator('input[type="password"]'),
      page.locator('input[placeholder*="senha" i]'),
    ];
    let passInput = null;
    for (const loc of candidates) if (await loc.count()) { passInput = loc.first(); break; }
    if (passInput) {
      await fillInputHuman(passInput, PASSWORD);
      const clicked = await clickFirstVisible('Entrar SSO', [
        page.getByRole('button', { name: /^entrar$/i }),
        page.getByText(/^ENTRAR$/i),
        page.locator('button[type="submit"]'),
      ]);
      if (!clicked) await passInput.press('Enter').catch(() => {});
      await waitForAppReady(page, 25000);
      await logPageState(page, 'Etapa de senha do SSO executada.', { shotName: 'after-sso-password' });
      summary = await pageSummary(page);
    }
  }
  if (isMfaMethodStep(summary) || isMfaCodeStep(summary)) {
    await handleMfa(page);
    return true;
  }
  return false;
}

async function login(page) {
  logEvent({ level: 'info', message: 'Iniciando login.' });
  await page.goto(`${BASE_URL}/sienge/`, { waitUntil: 'domcontentloaded' });
  await waitForAppReady(page);
  let summary = await pageSummary(page);
  if (isUserAlreadyLoggedScreen(summary)) {
    await debugPageSnapshot(page, 'antes de handleUserAlreadyLoggedScreen');
        await handleUserAlreadyLoggedScreen(page);
    summary = await pageSummary(page);
  }

  await debugPageSnapshot(page, 'antes de ensurePassWelcomeGate');
  await ensurePassWelcomeGate(page);
  await debugPageSnapshot(page, 'antes de trySsoLogin');
  await trySsoLogin(page);

  summary = await pageSummary(page);
  if (isUserAlreadyLoggedScreen(summary)) {
    await debugPageSnapshot(page, 'antes de handleUserAlreadyLoggedScreen');
        await handleUserAlreadyLoggedScreen(page);
    summary = await pageSummary(page);
  }

  if (isLoggedArea(summary)) {
    logEvent({ level: 'info', message: 'Login concluído com sucesso.' });
    await goDirectToAuthorizationPageAfterLogin(page);
    return;
  }
  throw new Error('Login não concluiu na área autenticada.');
}

async function ensureLoggedIn(context, page) {
  await page.goto(`${BASE_URL}/sienge/8/index.html`, { waitUntil: 'domcontentloaded' });
  await waitForAppReady(page, 12000);
  let summary = await pageSummary(page);

  if (isUserAlreadyLoggedScreen(summary)) {
    await debugPageSnapshot(page, 'antes de handleUserAlreadyLoggedScreen');
        await handleUserAlreadyLoggedScreen(page);
    summary = await pageSummary(page);
  }

  if (isWelcomeGate(summary)) {
    logEvent({ level: 'info', message: 'Sessão caiu na tela de boas-vindas; isso não é login válido.' });
    await debugPageSnapshot(page, 'antes de ensurePassWelcomeGate');
  await ensurePassWelcomeGate(page);
    await debugPageSnapshot(page, 'antes de trySsoLogin');
  await trySsoLogin(page);
    summary = await pageSummary(page);
  }
  if (isUserAlreadyLoggedScreen(summary)) {
    await debugPageSnapshot(page, 'antes de handleUserAlreadyLoggedScreen');
        await handleUserAlreadyLoggedScreen(page);
    summary = await pageSummary(page);
  }
  if (isAccountChooser(summary) || isEmailStep(summary) || isPasswordStep(summary) || isMfaMethodStep(summary) || isMfaCodeStep(summary)) {
    logEvent({ level: 'info', message: 'Fluxo SSO detectado durante verificação de sessão.' });
    await debugPageSnapshot(page, 'antes de trySsoLogin');
  await trySsoLogin(page);
    summary = await pageSummary(page);
  }
    if (isLoggedArea(summary)) {
      logEvent({ level: 'info', message: 'Sessão existente válida.', url: summary.url, title: summary.title });
      if (TASK_MODE === 'reports') {
        await goDirectToReportsPageAfterLogin(page);
      } else {
        await goDirectToAuthorizationPageAfterLogin(page);
      }
      return;
    }

  logEvent({ level: 'info', message: 'Sessão inválida ou ausente. Fazendo login completo.' });
  await login(page);
  await context.storageState({ path: STATE_PATH });
}



async function goDirectToAuthorizationPageAfterLogin(page) {
  logEvent({ level: 'info', message: 'Após login válido, tentando ir direto para a URL da tela 1777.' });

  await page.waitForTimeout(2500);
  await dismissHomeOverlays(page);
  await page.waitForTimeout(1200);

  const attempts = [
    async () => {
      await page.goto(TARGET_PAGE_URL, { waitUntil: 'domcontentloaded' });
      await waitForAppReady(page, 12000);
      return await waitForRealAuthorizationPage(page, 10000);
    },
    async () => {
      await page.goto(`${BASE_URL}/sienge/8/index.html`, { waitUntil: 'domcontentloaded' });
      await waitForAppReady(page, 12000);
      await page.evaluate(() => {
        try { window.location.hash = '#/common/page/1777'; } catch {}
      });
      await page.waitForTimeout(3000);
      return await waitForRealAuthorizationPage(page, 10000);
    },
    async () => {
      await page.goto(`${BASE_URL}/sienge/8/index.html`, { waitUntil: 'domcontentloaded' });
      await waitForAppReady(page, 12000);
      await clickRecentAuthorizationLink(page);
      return await waitForRealAuthorizationPage(page, 10000);
    },
    async () => {
      await page.goto(`${BASE_URL}/sienge/8/index.html`, { waitUntil: 'domcontentloaded' });
      await waitForAppReady(page, 12000);
      await openAuthViaMenu(page);
      return await waitForRealAuthorizationPage(page, 10000);
    },
  ];

  for (let i = 0; i < attempts.length; i++) {
    try {
      const ok = await attempts[i]();
      if (ok) {
        const summary = await pageSummary(page);
        logEvent({
          level: 'info',
          message: 'Redirecionamento pós-login para a tela de Autorização funcionou.',
          strategy: i + 1,
          url: page.url(),
          title: summary.title,
          bodySnippet: summary.bodySnippet,
        });
        return true;
      }
      logEvent({
        level: 'warning',
        message: 'Tentativa de redirecionamento pós-login não abriu o formulário real da 1777.',
        strategy: i + 1,
        actualUrl: page.url(),
      });
    } catch (err) {
      logEvent({
        level: 'warning',
        message: 'Tentativa de redirecionamento pós-login falhou.',
        strategy: i + 1,
        detail: String(err.message || err),
        actualUrl: page.url(),
      });
    }
  }

  return false;
}

async function hasVisible(locator) {
  const count = await locator.count().catch(() => 0);
  if (!count) return false;
  for (let i = 0; i < count; i++) {
    if (await locator.nth(i).isVisible({ timeout: 500 }).catch(() => false)) return true;
  }
  return false;
}


function surfaceName_(surface) {
  try {
    if (typeof surface.url === 'function') return surface.url() || 'frame';
  } catch {}
  return 'page';
}


function getAuthSurfaces_(page) {
  const surfaces = [];
  const seen = new Set();

  const push = (surface) => {
    if (!surface) return;
    const key = surfaceName_(surface);
    if (seen.has(key)) return;
    seen.add(key);
    surfaces.push(surface);
  };

  push(page);

  try {
    const authFrame = page.frame({ url: /\/sienge\/CPG\/listAutorizacaoPagamento\.do/i });
    push(authFrame);
  } catch {}

  try {
    for (const f of page.frames()) {
      if (!f || f === page.mainFrame()) continue;
      if (/\/sienge\/CPG\/listAutorizacaoPagamento\.do/i.test(f.url() || '')) {
        push(f);
      }
    }
  } catch {}

  try {
    for (const f of page.frames()) {
      if (!f || f === page.mainFrame()) continue;
      push(f);
    }
  } catch {}

  return surfaces;
}

async function waitForAuthorizationIframe(page, timeout = 20000) {
  const start = Date.now();
  while (Date.now() - start < timeout) {
    try {
      const frameEl = page.locator('iframe#iFramePage[src*="/sienge/CPG/listAutorizacaoPagamento.do"], iframe[src*="/sienge/CPG/listAutorizacaoPagamento.do"]');
      const count = await frameEl.count().catch(() => 0);
      if (count) {
        const visible = await frameEl.first().isVisible().catch(() => false);
        const handle = await frameEl.first().elementHandle().catch(() => null);
        const frame = handle ? await handle.contentFrame().catch(() => null) : null;
        if (visible && frame) {
          logEvent({
            level: 'info',
            message: 'Iframe da Autorização de Pagamento localizado.',
            iframeSrc: await frameEl.first().getAttribute('src').catch(() => ''),
            frameUrl: frame.url(),
          });
          return frame;
        }
      }
    } catch {}

    try {
      const byUrl = page.frame({ url: /\/sienge\/CPG\/listAutorizacaoPagamento\.do/i });
      if (byUrl) {
        logEvent({ level: 'info', message: 'Frame da Autorização localizado pela URL.', frameUrl: byUrl.url() });
        return byUrl;
      }
    } catch {}

    await page.waitForTimeout(500);
  }
  return null;
}

async function findAuthorizationSurface(page) {
  const iframeSurface = await waitForAuthorizationIframe(page, 6000).catch(() => null);
  if (iframeSurface) {
    try {
      const dtInicio = iframeSurface.locator('input#dtInicio, input[name="dtInicio"]');
      const dtFim = iframeSurface.locator('input#dtFim, input[name="dtFim"]');
      const consultar = iframeSurface.locator('input[type="submit"][name="btFiltrar"][value="Consultar"], input[type="submit"][value*="Consultar" i]');
      const hasIni = await dtInicio.count().catch(() => 0);
      const hasFim = await dtFim.count().catch(() => 0);
      const hasConsultar = await consultar.count().catch(() => 0);
      logEvent({ level: 'debug', message: 'Checagem do iframe da Autorização.', hasIni, hasFim, hasConsultar, frameUrl: iframeSurface.url() });
      if (hasIni && hasFim && hasConsultar) return iframeSurface;
    } catch {}
  }

  for (const surface of getAuthSurfaces_(page)) {
    try {
      const dtInicio = surface.locator('input#dtInicio, input[name="dtInicio"]');
      const dtFim = surface.locator('input#dtFim, input[name="dtFim"]');
      const consultar = surface.locator('input[type="submit"][name="btFiltrar"][value="Consultar"], input[type="submit"][value*="Consultar" i]');

      const hasIni = await dtInicio.count().catch(() => 0);
      const hasFim = await dtFim.count().catch(() => 0);
      const hasConsultar = await consultar.count().catch(() => 0);
      if (hasIni && hasFim && hasConsultar) return surface;

      const dateInputs = surface.locator('input[formattype="DATE"]');
      const dateCount = await dateInputs.count().catch(() => 0);
      if (dateCount >= 2 && hasConsultar) return surface;
    } catch {}
  }
  return null;
}

async function clickFirstVisibleOnSurface(label, surface, locators) {

  for (const build of locators) {
    try {
      const locator = build(surface);
      const count = await locator.count().catch(() => 0);
      if (!count) continue;
      const first = locator.first();
      if (!(await first.isVisible({ timeout: 1500 }).catch(() => false))) continue;
      await first.click({ timeout: 8000 });
      logEvent({ level: 'info', message: `${label}: clique realizado.`, surface: surfaceName_(surface) });
      return true;
    } catch (err) {
      logEvent({ level: 'debug', message: `${label}: tentativa falhou.`, surface: surfaceName_(surface), detail: String(err.message || err) });
    }
  }
  return false;
}

async function isRealAuthorizationPage(page) {
  const surface = await findAuthorizationSurface(page);
  if (!surface) return false;

  const hasDtInicio = await surface.locator('input#dtInicio, input[name="dtInicio"]').count().catch(() => 0);
  const hasDtFim = await surface.locator('input#dtFim, input[name="dtFim"]').count().catch(() => 0);
  const hasConsultar = await surface.locator('input[type="submit"][name="btFiltrar"][value="Consultar"], input[type="submit"][value*="Consultar" i]').count().catch(() => 0);
  return !!(hasDtInicio && hasDtFim && hasConsultar);
}

async function waitForRealAuthorizationPage(page, timeout = 15000) {
  const start = Date.now();
  while (Date.now() - start < timeout) {
    if (await isRealAuthorizationPage(page)) return true;
    await page.waitForTimeout(500);
  }
  return false;
}

async function navigateByHash(page) {
  logEvent({ level: 'info', message: 'Tentando navegar via hash para a página 1777.' });
  await page.evaluate((hash) => {
    try {
      window.location.hash = hash;
      window.dispatchEvent(new HashChangeEvent('hashchange'));
    } catch {}
  }, '#/common/page/1777');
  await waitForAppReady(page, 12000);
  return await waitForRealAuthorizationPage(page, 10000);
}

async function clickRecentAuthorizationLink(page) {
  logEvent({ level: 'info', message: 'Tentando abrir Autorização de Pagamento pelo card de acessos recentes.' });
  const clicked = await clickFirstVisible('Acesso recente Autorização de Pagamento', [
    page.locator('a').filter({ hasText: /Autorização de Pagamento/i }).filter({ hasText: /Financeiro\s*\/\s*Contas a Pagar/i }),
    page.locator('a,div,span,p').filter({ hasText: /^Autorização de Pagamento$/i }).locator('xpath=ancestor::*[self::a or self::div][1]'),
    page.locator('text=/Autorização de Pagamento/i').locator('xpath=ancestor::*[contains(@class,"card") or contains(@class,"Card") or self::a][1]'),
  ]);
  if (!clicked) return false;
  await waitForAppReady(page, 12000);
  return await waitForRealAuthorizationPage(page, 10000);
}

async function openAuthViaMenu(page) {
  logEvent({ level: 'info', message: 'Tentando abrir Autorização de Pagamento pelo menu lateral.' });
  const steps = [
    /Financeiro/i,
    /Contas a Pagar/i,
    /Autorização de Pagamento|Autorizacao de Pagamento/i,
  ];
  for (const step of steps) {
    const ok = await clickFirstVisible(`Menu ${step}`, [
      page.getByRole('link', { name: step }),
      page.getByRole('button', { name: step }),
      page.locator('a,button,span,div').filter({ hasText: step }),
    ]);
    if (!ok) return false;
    await page.waitForTimeout(900);
  }
  await waitForAppReady(page, 12000);
  return await waitForRealAuthorizationPage(page, 10000);
}

async function waitForAuthorizationShell(page) {
  return await waitForRealAuthorizationPage(page, 12000);
}

async function openAuthorizationPage(page) {
  logEvent({ level: 'info', message: 'Abrindo tela de Autorização de Pagamento.' });
  await debugPageSnapshot(page, 'início de openAuthorizationPage');
  await debugFrames(page, 'início de openAuthorizationPage');

  let surface = await findAuthorizationSurface(page);
  if (surface) {
    logEvent({ level: 'info', message: 'A superfície real da Autorização já está carregada.', surface: surfaceName_(surface) });
    return surface;
  }

  const targetUrls = [TARGET_PAGE_URL, `${TARGET_PAGE_URL}?tab=0`];

  for (const url of targetUrls) {
    logEvent({ level: 'info', message: `Tentando abrir URL direta: ${url}` });
    await page.goto(url, { waitUntil: 'domcontentloaded' });
    await waitForAppReady(page, 15000);
    await dismissHomeOverlays(page);

    surface = await waitForAuthorizationIframe(page, 12000);
    if (!surface) {
      try {
        await page.evaluate(() => {
          if (window.location.hash !== '#/common/page/1777') window.location.hash = '#/common/page/1777';
        });
      } catch {}
      await waitForAppReady(page, 8000);
      surface = await waitForAuthorizationIframe(page, 12000);
    }

    if (surface && await isRealAuthorizationPage(page)) {
      const summary = await pageSummary(page);
      logEvent({
        level: 'info',
        message: 'Tela de Autorização confirmada pelo iframe real.',
        url: page.url(),
        title: summary.title,
        iframeUrl: surface.url(),
      });
      return surface;
    }

    await debugPageSnapshot(page, `falha para abrir 1777 via ${url}`);
    await debugFrames(page, `falha para abrir 1777 via ${url}`);
  }

  surface = await clickRecentAuthorizationLink(page) ? await findAuthorizationSurface(page) : null;
  if (surface) return surface;

  surface = await openAuthViaMenu(page) ? await findAuthorizationSurface(page) : null;
  if (surface) return surface;

  await logPageState(page, 'Não consegui abrir a superfície real da tela de Autorização de Pagamento.', {
    level: 'error',
    shotName: 'open-auth-page-failed',
  });
  throw new Error('Não consegui abrir a superfície real da tela de Autorização de Pagamento.');
}

async function detectDateInputs(page) {
  const surface = await findAuthorizationSurface(page);
  if (!surface) return [];

  const dtInicio = surface.locator('input#dtInicio, input[name="dtInicio"]');
  const dtFim = surface.locator('input#dtFim, input[name="dtFim"]');

  if (await dtInicio.count().catch(() => 0) && await dtFim.count().catch(() => 0)) {
    const a = dtInicio.first();
    const b = dtFim.first();
    const aVisible = await a.isVisible({ timeout: 1000 }).catch(() => false);
    const bVisible = await b.isVisible({ timeout: 1000 }).catch(() => false);
    if (aVisible && bVisible) {
      logEvent({ level: 'info', message: 'Campos de período localizados por ID/name: dtInicio e dtFim.', surface: surfaceName_(surface) });
      return [a, b, surface];
    }
  }

  const dateInputs = surface.locator('input[formattype="DATE"]');
  const dateCount = await dateInputs.count().catch(() => 0);
  if (dateCount >= 2) {
    const visible = [];
    for (let i = 0; i < dateCount; i++) {
      const item = dateInputs.nth(i);
      const isVisible = await item.isVisible({ timeout: 1000 }).catch(() => false);
      if (isVisible) visible.push(item);
    }
    if (visible.length >= 2) {
      logEvent({ level: 'info', message: 'Campos de período localizados por formattype=DATE.', surface: surfaceName_(surface) });
      return [visible[0], visible[1], surface];
    }
  }

  return [];
}



async function detectNoAuthorizationResults(page, authSurface) {
  const surface = authSurface || await findAuthorizationSurface(page);
  const candidates = [];

  if (surface) candidates.push(surface);
  candidates.push(page);

  for (const ctx of candidates) {
    try {
      const noRowsText = ctx.getByText(/não há registros para os parâmetros informados|nao ha registros para os parametros informados/i);
      if (await noRowsText.count().catch(() => 0)) {
        const visible = await noRowsText.first().isVisible().catch(() => false);
        if (visible) {
          return {
            found: true,
            reason: 'message',
            text: 'Não há registros para os parâmetros informados.',
            surface: surfaceName_(ctx),
          };
        }
      }

      const qtyText = ctx.getByText(/quantidade de registros:\s*0/i);
      if (await qtyText.count().catch(() => 0)) {
        const visible = await qtyText.first().isVisible().catch(() => false);
        if (visible) {
          return {
            found: true,
            reason: 'count_zero',
            text: 'Quantidade de registros: 0',
            surface: surfaceName_(ctx),
          };
        }
      }

      const markAllBtn = ctx.locator('input[type="button"][value*="Marcar todos" i], input[type="submit"][value*="Marcar todos" i]');
      const saveBtn = ctx.locator('input[type="button"][value*="Salvar" i], input[type="submit"][value*="Salvar" i]');
      const markCount = await markAllBtn.count().catch(() => 0);
      const saveCount = await saveBtn.count().catch(() => 0);

      if (!markCount && !saveCount) {
        const infoBanner = ctx.locator('div,span,font,td').filter({ hasText: /não há registros para os parâmetros informados|nao ha registros para os parametros informados/i });
        if (await infoBanner.count().catch(() => 0)) {
          return {
            found: true,
            reason: 'banner_without_actions',
            text: 'Nenhum resultado após consulta.',
            surface: surfaceName_(ctx),
          };
        }
      }
    } catch {}
  }

  return { found: false };
}

async function configureFilters(page, surface) {
  const startDate = todayBr();
  const endDate = TARGET_END_DATE;
  logEvent({ level: 'info', message: `Configurando período ${startDate} até ${endDate}.` });
  await debugPageSnapshot(page, 'início de configureFilters');

  const authSurface = surface || await findAuthorizationSurface(page);
  if (!authSurface) {
    await logPageState(page, 'A página atual não contém o iframe real da tela de Autorização de Pagamento.', {
      level: 'error',
      shotName: 'wrong-page-before-date-detection',
      currentUrl: page.url(),
    });
    throw new Error('A página atual não contém o iframe real da Autorização de Pagamento.');
  }

  const dateA = authSurface.locator('input#dtInicio, input[name="dtInicio"]').first();
  const dateB = authSurface.locator('input#dtFim, input[name="dtFim"]').first();
  const hasA = await dateA.count().catch(() => 0);
  const hasB = await dateB.count().catch(() => 0);
  if (!hasA || !hasB) {
    await debugLocatorState(page, 'date-fields-missing-iframe', {
      dtInicio: authSurface.locator('input#dtInicio, input[name="dtInicio"]'),
      dtFim: authSurface.locator('input#dtFim, input[name="dtFim"]'),
      consultar: authSurface.locator('input[type="submit"][name="btFiltrar"][value="Consultar"]'),
    });
    throw new Error('Não consegui localizar os campos dtInicio/dtFim dentro do iframe da Autorização.');
  }

  logEvent({ level: 'info', message: 'Campos de período localizados no iframe da Autorização.', surface: surfaceName_(authSurface) });
  await fillInputHuman(dateA, startDate);
  await fillInputHuman(dateB, endDate);

  const radioClicked = await clickFirstVisibleOnSurface('Radio Somente não autorizados', authSurface, [
    s => s.getByLabel(/somente não autorizados/i),
    s => s.getByText(/somente não autorizados/i),
    s => s.locator('label').filter({ hasText: /somente não autorizados/i }),
    s => s.locator('input[type="radio"]').nth(0),
  ]);
  if (!radioClicked) throw new Error('Não consegui marcar o radio "Somente não autorizados" no iframe.');

  const consultClicked = await clickFirstVisibleOnSurface('Botão Consultar', authSurface, [
    s => s.locator('input[type="submit"][name="btFiltrar"][value="Consultar"]'),
    s => s.locator('input[type="submit"][value*="Consultar" i], input[type="button"][value*="Consultar" i]'),
    s => s.getByRole('button', { name: /consultar/i }),
    s => s.getByText(/^CONSULTAR$/i),
  ]);
  if (!consultClicked) throw new Error('Não consegui clicar em Consultar dentro do iframe.');

  await page.waitForTimeout(2500);
  await logPageState(page, 'Consulta executada.', { shotName: 'after-consultar' });

  const noResults = await detectNoAuthorizationResults(page, authSurface);
  if (noResults.found) {
    logEvent({
      level: 'info',
      message: 'Consulta realizada, mas não há parcelas pendentes para autorizar nesta execução.',
      reason: noResults.reason,
      detail: noResults.text,
      surface: noResults.surface || '',
      currentUrl: page.url(),
    });
    return { hasResults: false, reason: noResults.reason };
  }

  return { hasResults: true };
}

async function markAllAndSave(page, surface) {
  const authSurface = surface || await findAuthorizationSurface(page);
  if (!authSurface) throw new Error('Não encontrei o iframe real da tela de Autorização.');

  const noResults = await detectNoAuthorizationResults(page, authSurface);
  if (noResults.found) {
    logEvent({
      level: 'info',
      message: 'Nenhum resultado disponível para marcar/salvar. Pulando autorização nesta execução.',
      reason: noResults.reason,
      detail: noResults.text,
      surface: noResults.surface || '',
      currentUrl: page.url(),
    });
    return { skipped: true, reason: noResults.reason };
  }

  const markAllClicked = await clickFirstVisibleOnSurface('Botão Marcar todos', authSurface, [
    s => s.locator('input[type="button"][value*="Marcar todos" i], input[type="submit"][value*="Marcar todos" i]'),
    s => s.getByRole('button', { name: /marcar todos/i }),
    s => s.getByText(/marcar todos/i),
  ]);
  if (!markAllClicked) throw new Error('Não encontrei o botão "Marcar todos" dentro do iframe.');

  await page.waitForTimeout(1200);

  const saveClicked = await clickFirstVisibleOnSurface('Botão Salvar', authSurface, [
    s => s.locator('input[type="button"][value*="Salvar" i], input[type="submit"][value*="Salvar" i]'),
    s => s.getByRole('button', { name: /^salvar$/i }),
    s => s.getByText(/^SALVAR$/i),
  ]);
  if (!saveClicked) throw new Error('Não encontrei o botão "Salvar" dentro do iframe.');

  await page.waitForTimeout(2500);
  await logPageState(page, 'Salvar executado.', { shotName: 'after-salvar' });
  return { skipped: false };
}

async function attachPageDebug(page) {
  if (!DEBUG_PAGE_EVENTS) return;
  page.on('pageerror', (err) => logEvent({ level: 'debug', message: 'Erro JS na página.', detail: String(err.message || err) }));
  page.on('console', (msg) => {
    if (['error', 'warning'].includes(msg.type())) {
      logEvent({ level: 'debug', message: `Console ${msg.type()} da página.`, detail: msg.text() });
    }
  });
}

async function runAuthorization(context, page) {
  const authSurface = await openAuthorizationPage(page);
  const filterResult = await configureFilters(page, authSurface);

  if (filterResult && filterResult.hasResults === false) {
    const idleShot = CAPTURE_SUCCESS_SCREENSHOTS ? await saveShot(page, 'final-no-results') : null;
    const idleSummary = await pageSummary(page, { includeBody: false });
    logEvent({
      level: 'info',
      message: 'Execução concluída sem parcelas pendentes para autorizar.',
      screenshot: idleShot,
      ...idleSummary
    });
  } else {
    await markAllAndSave(page, authSurface);
    const finalShot = CAPTURE_SUCCESS_SCREENSHOTS ? await saveShot(page, 'final-success') : null;
    const summary = await pageSummary(page, { includeBody: false });
    logEvent({ level: 'info', message: 'Processo de autorização concluído com sucesso.', screenshot: finalShot, ...summary });
  }

}

async function run() {
  // Cria diretórios somente quando aquele tipo de artefato será realmente usado.
  if (TASK_MODE === 'reports' || TASK_MODE === 'both') await ensureDir(REPORT_OUTPUT_DIR);
  const browser = await chromium.launch({
    headless: HEADLESS,
    args: [
      '--disable-gpu',
      '--disable-extensions',
      '--disable-component-update',
      '--disable-background-networking',
      '--disable-sync',
      '--metrics-recording-only',
      '--mute-audio',
    ],
  });
  const contextOptions = {
    viewport: { width: VIEWPORT_WIDTH, height: VIEWPORT_HEIGHT },
    deviceScaleFactor: 1,
  };
  if (fs.existsSync(STATE_PATH)) contextOptions.storageState = STATE_PATH;
  const context = await browser.newContext(contextOptions);
  if (BLOCK_NON_ESSENTIAL_RESOURCES) {
    await context.route('**/*', route => {
      const type = route.request().resourceType();
      return ['image', 'media', 'font'].includes(type) ? route.abort() : route.continue();
    });
  }
  const page = await context.newPage();
  await attachPageDebug(page);

  try {
    logEvent({
      level: 'info',
      message: 'Início da execução do robô.',
      baseUrl: BASE_URL,
      headless: HEADLESS,
      mode: TASK_MODE,
      targetPage: TARGET_PAGE_URL,
      reportPage: REPORT_FILTER_PAGE_URL,
      logPath: LOG_PATH
    });

    await ensureLoggedIn(context, page);

    if (TASK_MODE === 'reports') {
      await runReports(context, page);
    } else if (TASK_MODE === 'both') {
      await runAuthorization(context, page);
      const reportPage = await context.newPage();
      await attachPageDebug(reportPage);
      await runReports(context, reportPage);
    } else {
      await runAuthorization(context, page);
    }

    await context.storageState({ path: STATE_PATH });
    logEvent({ level: 'info', message: 'Estado da sessão salvo com sucesso.', statePath: STATE_PATH });
  } catch (err) {
    // PM2 usa SIGINT em stop/reload. Nesse caso o Chromium pode já ter sido
    // encerrado pelo supervisor, portanto não tente capturar uma página morta.
    if (__stopping || page.isClosed() || !browser.isConnected()) {
      logEvent({
        level: 'warning',
        message: 'Execução interrompida por sinal do processo; diagnóstico visual ignorado.',
        detail: String(err && err.message || err),
      });
      return;
    }

    let errorShot = null;
    let html = null;
    let summary = { url: page.url(), title: '', bodySnippet: '' };
    try {
      errorShot = await saveShot(page, 'error');
      html = await saveHtml(page, 'error');
      summary = await pageSummary(page);
    } catch (captureErr) {
      logEvent({
        level: 'warning',
        message: 'Não foi possível capturar o diagnóstico visual da falha.',
        detail: String(captureErr && captureErr.message || captureErr),
      });
    }
    await sendZapiAlert({
      title: err.message || 'Falha na execução do robô',
      detail: err.message || String(err),
      stack: String(err.stack || ''),
      url: summary.url,
      pageTitle: summary.title,
    });
    logEvent({ level: 'error', message: err.message, screenshot: errorShot, html, stack: String(err.stack || ''), ...summary });
    console.error(err);
    process.exitCode = 1;
  } finally {
    await browser.close().catch(() => {});
    logEvent({ level: 'info', message: 'Execução finalizada.' });
    flushLog();
  }
}



const PM2_LOOP = String(process.env.PM2_LOOP || 'true').toLowerCase() !== 'false';
const PM2_INTERVAL_MS = Number(process.env.PM2_INTERVAL_MS || process.env.POLL_INTERVAL_MS || 60000);
const PM2_STOP_ON_FATAL = String(process.env.PM2_STOP_ON_FATAL || 'false').toLowerCase() === 'true';

let __running = false;
let __stopping = false;

async function sleepMs(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

async function runOnceSafely() {
  if (__running) {
    logEvent({ level: 'warning', message: 'Execução anterior ainda está em andamento. Pulando este ciclo.' });
    return;
  }

  __running = true;
  try {
    await run();
  } catch (err) {
    logEvent({
      level: 'error',
      message: 'Falha não tratada no ciclo principal.',
      detail: String(err && err.message || err),
      stack: String(err && err.stack || '')
    });

    if (PM2_STOP_ON_FATAL) {
      process.exitCode = 1;
      __stopping = true;
    }
  } finally {
    __running = false;
  }
}

async function startPm2Loop() {
  logEvent({
    level: 'info',
    message: 'Inicializando modo PM2.',
    pm2Loop: PM2_LOOP,
    intervalMs: PM2_INTERVAL_MS,
    stopOnFatal: PM2_STOP_ON_FATAL
  });

  if (!PM2_LOOP) {
    await runOnceSafely();
    return;
  }

  while (!__stopping) {
    const cycleStartedAt = Date.now();
    await runOnceSafely();

    if (__stopping) break;

    const elapsed = Date.now() - cycleStartedAt;
    const waitMs = Math.max(1000, PM2_INTERVAL_MS - elapsed);

    logEvent({
      level: 'info',
      message: 'Aguardando próximo ciclo do PM2.',
      waitMs
    });

    await sleepMs(waitMs);
  }

  logEvent({ level: 'info', message: 'Loop PM2 finalizado.' });
}

process.on('SIGINT', async () => {
  logEvent({ level: 'info', message: 'SIGINT recebido. Encerrando loop PM2.' });
  __stopping = true;
  flushLog();
});

process.on('SIGTERM', async () => {
  logEvent({ level: 'info', message: 'SIGTERM recebido. Encerrando loop PM2.' });
  __stopping = true;
  flushLog();
});

startPm2Loop().catch(err => {
  logEvent({
    level: 'error',
    message: 'Falha fatal ao inicializar o modo PM2.',
    detail: String(err && err.message || err),
    stack: String(err && err.stack || '')
  });
  sendZapiAlert({
    title: 'Falha fatal ao inicializar o modo PM2',
    detail: String(err && err.message || err),
    stack: String(err && err.stack || ''),
  }).catch(() => {});
  flushLog();
  process.exit(1);
});
