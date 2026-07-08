#!/usr/bin/env node
const fs = require('fs');
const path = require('path');
const readline = require('readline/promises');
const { stdin: input, stdout: output } = require('process');
const { chromium } = require('playwright');
require('dotenv').config();

// Instance identifier to allow multiple parallel runs. Defaults to process.pid
const INSTANCE_ID = process.env.INSTANCE_ID;

// Use a single global status file by default to aggregate progress across runs.
// Can be overridden via the `STATUS_FILE` env var when needed.
const STATUS_FILE = process.env.STATUS_FILE || '/tmp/report-status.json';
const BASE_URL = process.env.SIENGE_BASE_URL;
const USERNAME = process.env.SIENGE_USERNAME;
const PASSWORD = process.env.SIENGE_PASSWORD;
const HEADLESS = (process.env.HEADLESS ?? 'true').toLowerCase() !== 'false';
const STATE_PATH = process.env.STATE_PATH || path.resolve(process.cwd(), `solo_sienge-storage-state-${INSTANCE_ID}.json`);
const SCREENSHOT_DIR = process.env.SCREENSHOT_DIR || path.resolve(process.cwd(), `screenshots-${INSTANCE_ID}`);
const LOG_PATH = process.env.LOG_PATH || path.resolve(process.cwd(), `solo_sienge-authorize-log-${INSTANCE_ID}.json`);
const DEBUG_HTML = (process.env.DEBUG_HTML ?? 'false').toLowerCase() === 'true';
const BLOCK_RESOURCES = (process.env.BLOCK_RESOURCES ?? 'true').toLowerCase() !== 'false';
const REPORT_OUTPUT_DIR = process.env.REPORT_OUTPUT_DIR || path.resolve(process.cwd(), 'reports');
const TASK_MODE = (process.env.TASK_MODE || 'authorize').toLowerCase();
const TARGET_PAGE_URL = `${BASE_URL}/sienge/8/index.html#/common/page/1777`;
const REPORT_FILTER_PAGE_URL = `${BASE_URL}/sienge/8/index.html#/common/page/4929`;
const SAVE_SHOTS = (process.env.SAVE_SHOTS ?? 'false').toLowerCase() === 'true' || DEBUG_HTML;
const MODAL_CACHE_PATH = process.env.MODAL_CACHE_PATH || path.resolve(process.cwd(), `.modal-cache-${INSTANCE_ID}.json`);

const TARGET_END_DATE = '31/12/2040';
const REPORT_PERIOD_START = '01/04/2026';
const REPORT_PERIOD_END = '30/04/2026';


const STATE_REVALIDATE_TIMEOUT_MS = Number(process.env.STATE_REVALIDATE_TIMEOUT_MS || 15000);
const STATE_REVALIDATE_POLL_MS = Number(process.env.STATE_REVALIDATE_POLL_MS || 2000);

const SINGLE_REPORT_ARG = (process.env.SINGLE_REPORT || getCliArg('--single') || getCliArg('-s')) || null;
let MODAL_CACHE = {};

function getCliArg(name) {
  const idx = process.argv.indexOf(name);
  if (idx >= 0 && process.argv.length > idx + 1) return process.argv[idx + 1];
  return null;
}

// Global helper: retry actions that fail due to transient frame detach / execution context errors
async function withFrameRetry(actionFn, { retries = 2, name = 'action' } = {}) {
  for (let attempt = 1; attempt <= retries; attempt++) {
    try {
      return await actionFn();
    } catch (err) {
      const msg = String(err && (err.message || err));
      if (/Frame was detached|Execution context was destroyed|Cannot find context|Cannot find object/i.test(msg) && attempt < retries) {
        logEvent({ level: 'warn', message: `${name}: quadro foi recriado; refazendo tentativa - detail=${msg}` });
        await new Promise(r => setTimeout(r, 250 * attempt));
        continue;
      }
      throw err;
    }
  }
        // Optional: post-modal cleanup of SelectedEntitiesList can interfere with app behavior
        // Enable explicitly via env var ENABLE_POST_MODAL_CLEANUP=1 when needed.
        try {
          if (process.env.ENABLE_POST_MODAL_CLEANUP === '1') {
            try {
              const iframeLocator2 = page.locator('iframe#iFramePage');
              const mainFrame2 = await iframeLocator2.contentFrame().catch(() => null);
              if (mainFrame2) {
                await mainFrame2.evaluate(() => {
                  try {
                    Array.from(document.querySelectorAll('[id$="SelectedEntitiesList"], [class*="SelectedEntitiesList"]')).forEach(el => {
                      if (el.tagName === 'INPUT' || el.tagName === 'TEXTAREA') el.value = '';
                      else el.innerHTML = '';
                    });
                    Array.from(document.querySelectorAll('[id*="contador"], [name*="contador"]')).forEach(el => {
                      if (el.tagName === 'INPUT' || el.tagName === 'TEXTAREA') el.value = '0';
                      else el.textContent = '0';
                    });
                  } catch (e) {}
                }).catch(() => {});
              }
            } catch (_) {}
          } else {
            logEvent({ level: 'debug', message: 'Skipping post-modal cleanup (ENABLE_POST_MODAL_CLEANUP!=1)', modalTitle });
          }
        } catch (_) {}
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

const DOCUMENTOS_NEXT_PAGE_SELECTORS = [
  'input[type="button"].spwBotaoProximo[title="Próxima página"]',
  'input.spwBotaoProximo[title="Próxima página"]',
  '.spwBotaoProximo',
  '.ui-paginator-next',
  'a.ui-paginator-next',
  'button.ui-paginator-next',
  'a[aria-label="Proxima pagina"]',
  'a[aria-label="Próxima página"]',
  'button[title*="Próxima"]',
  'a[title*="Próxima"]',
  'a[aria-label*="Next"]',
];

const REPORT_DEFINITIONS = [
  {
    sheetName: 'Receita Parcelas Mensal',
    pdfName: 'Receita Parcelas Mensal',
    planoFinanceiro: [
    "Receita de Incorporação de Imóveis",
    "Receita de Bens/Imóveis Adquiridos de Terceiros",
    "Descontos Concedidos (Unidades Imobiliárias)",
    "xxxxxReceita de Aluguel (Locação)",
    "Permuta Externa",
    "Receita de Cliente Desconhecido",
    "Receita de Administração (Taxa Administrativa)",
    "Receita Empreitada Só Mão de Obra",
    "Receita de Projeto",
    "Receita de Locação (Aluguel)",
    "Entrada de Financiamento para Construção",
    "Entrada de Capital de Giro",
    "Receita de Aplicações Financeiras",
    "Receita de Empréstimos",
    "Aporte Capital de Sócio",
    "Descontos Obtidos",
    "Receita de Venda de Imobilizado",
    "Receita Aluguel PCS 35",
    "Estoque de Material",
    "Receita Energia Solar Locatários",
    "Venda de Passivos",
    "(-) Anulação de Custos",
    "(-) Anulação de Custos Mão de Obra",
    "(-) Retenção Caução/Sinal",
    "(-) Retenção de INSS",
    "(-) Retenção de ISS",
    "(-) Retenção de PIS",
    "(-) Retenção de CSLL",
    "(-) Retenção de COFINS",
    "(-) Retenção de IR",
    "(-) Retenção (PIS/COFINS/CSLL - 4,65%)",
    "(-) Retenção de Permuta",
    "(-) Transf. Materiais Estoque Laurentino P/ Finan",
    "(-) Transf. Materiais Estoque Laurentino P/ Bom J.",
    "(-) Transf. Mat. Estoque P/ Obras da Laurentino",
    "(-) Reembolso de Despesas Gerais (Obra)",
    "(-) Divisão de muro",
    "(-) Retenção Contribuição Sindical (Escritório)",
    "(-) Reembolso de Despesas Administrativas",
    "(-) Restituição IPTU Imóveis Alugados (Escritório)",
    "(-) Reembolso de Despesas Financeiras",
    "(-) Reembolso de Despesas Tributárias",
    "(-) Abatimento de Adiantamento",
    "(-) Restituição Condomínio",
    "(-) Restituição água"
],
    contasCorrente: [
      "Viacredi Laurentino",
      "CEF Laurentino",
      "Sicredi Doca",
      "Sicredi Laurentino",
      "Aplicação Viacredi Laurentino",
      "ADIANTAMENTO CASA DA LAJE 150 MIL",
      "ADIANTAMENTO 250 MIL PRA CASA DA LAJE",
      "Aplicação CEF Laurentino",
      "Aplicação Doca Sicredi",
      "Aplicação Sicredi Laurentino",
      "ADIANTAMENTO CASA DA LAJE 175K",
      "ADIANTAMENTO CASA DA LAJE 200K",
      "ADIANTAMENTO CASA DA LAJE 250K",
      "ADIANTAMENTO 95K PORTARIS",
      "CAIXA Laurentino",
      "Conta Babi - Não Conciliar",
      "Conta Doca",
      "Conta Equilíbrio",
      "ESTOQUE",
      "PERMUTA INDAPAV BP1 LOTE 13 CASA 1",
      "Lançamentos Passados",
      "PERMUTA DISBRACON LOTE 10 E 14 BV3",
      "PERMUTA AB3 MR VIDROS",
      "PERMUTA AB 4 E 5 MONDINI",
      "PERMUTA DISBRACON ACACIA 26",
      "PERMUTA ACÁCIAS POMERODE LOTE 8 MINI CAR",
      "ADIANTAMENTO 115K VIDRAÇARIA ITOUP",
      "PERMUTA AMS LAJES GOLDEN PARK",
      "PERMUTA AMS LAJES LOTE 18 ACACIA POMEROD",
      "PERMUTA GAIO SERPA",
      "PERMUTA BERTELLI GREEN GARDEN LOTE 36",
      "PERMUTA BERTELLI KW 3 SALA 1",
      "PERMUTA BERTELLI KW 3 SALA 2",
      "PERMUTA BERTELLI SEVEN POINT SALA 9",
      "PERMUTA BONGO 2020 EBERT",
      "PERMUTA MONDINI LOTE 09 CASA 2 BP1",
      "PERMUTA BERTELLI VILA DAS NAÇÕES LT 15",
      "PERMUTA BERTELLI VILA DAS NAÇÕES LT 14",
      "PERMUTA PORTARIS CA 90 CS 1",
      "PERMUTA CALHAS INDAIAL LT NOVA INDAIAL 4",
      "PERMUTA CIVIC EBERT",
      "PERMUTA CASA MIA SEVEN POINT SALA 6",
      "PERMUTA CASA MIA VILA DAS NAÇÕES LT 18",
      "PERMUTA CONCRETOREI GREEN GARDEN LOTE 45",
      "PERMUTA DAINOX BP1 TERRENO MAT. 21884",
      "PERMUTA DAINOX SALA 7 SEVEN POINT",
      "PERMUTA DAINOX GOLDEN PARK LOT A DEF.",
      "PERMUTA DAMACENO",
      "PERMUTA DAMACENO LAND DO ALMIR",
      "PERMUTA DISBRACON ACÁCIA POMERODE LOTE16",
      "PERMUTA DISBRACON ACÁCIA POMERODE LOTE17",
      "PERMUTA DISBRACON SEVEN POINT SALA 4",
      "PERMUTA DISBRACON SEVEN POINT SALA 5",
      "PERMUTA EBERT ONIX",
      "PERMUTA ECO SPORT DAMACENO",
      "PERMUTA FCF RENEGADE (TROCOU DO PA 28)",
      "PERMUTA FIESTA (PEGEOUT) EBERT",
      "PERMUTA CASA MIA GARTEN DORF LOTE 1 E 19",
      "PERMUTA GG 21 DAINOX",
      "PERMUTA CASA DA LAJE GREEN GARDEN LT 30",
      "PERMUTA PORTARIS GG 31",
      "PERMUTA GREEN GARDEN GRAMEIRA FELIPPI 27",
      "PERMUTA GREEN GARDEN MULDE JANDT",
      "Permuta Gol Kosmos",
      "PERMUTA GRAMEIRA FELIPPI GOLDEN PARK",
      "PERMUTA GRAMEIRA FELIPPI GOLDEN PARK 2",
      "PERMUTA GRAMEIRA F. TOLEDO (TROCOU RY12)",
      "PERMUTA INDAPAV SALA 13 SEVEN POINT",
      "PERMUTA INDUMADER GREEN GARDEN 38",
      "PERMUTA INDUMADER GOLDEN PARK 2 LOTES",
      "PERMUTA INDUMADER GREEN VILLAGE 2LT 540M",
      "PERMUTA INDUMADER NOVA A DEFINIR",
      "PERMUTA INDUMADER GOLDEN PARK 3 LOTES",
      "PERMUTA INDUMADER ROYAL PARK 4/5/6",
      "PERMUTA INDUMADER ROYAL PARK 7/8/9",
      "PERMUTA INDAPAV WARNOW 12X80",
      "PERMUTA DAINOX GG LOTE 19",
      "PERMUTA INDAPAV VOLVO GARTEN DORF",
      "PERMUTA VIDRAÇARIA ITOUPAVAZINHA GP1 CS1",
      "PERMUTA JOURNEY EBERT",
      "PERMUTA KOSMOS GREEN GARDEN LOTE 23",
      "PERMUTAS KOSMOS SEVEN POINT SALA 10",
      "CASA DA LAJE ADIANTAMENTO 357K",
      "PERMUTA CASA DA LAJE KW 4 SALA 02",
      "PERMUTA CASA DA LAJE ACACIA POMEROD LT20",
      "PERMUTA CASA DA LAJE ROYAL PARK LT 13/14",
      "PERMUTA CASA DA LAJE SP SALA 8",
      "PERMUTA LF CALHAS GOLDEN PARK A DEF.",
      "PERMUTA LF CALHAS GREEN VILLAGE A DEF.",
      "PERMUTA LF CALHAS SEVEN POINT SALA 1",
      "PERMUTA LF CALHAS RUA TOLEDO",
      "PERMUTA CASA DA LAJE LOTE 03 BV3",
      "PERMUTA CASA DA LAJE LOTE 13 BV",
      "PERMUTA INDAPAV PC 35",
      "PERMUTA CASA DA LAJE LOTE 49 CA",
      "PERMUTA CALHAS INDAIAL LOTE 72 AA",
      "PERMUTA BERTELLI LOTE 10 BP1",
      "PERMUTA VIDRAÇARIA ITOUPAVAZINHA LT11BP1",
      "PERMUTA VIDR. ITOUPAZAVINHA 55 CA",
      "PERMUTA VIDR. ITOUPAVAZINHA 7/8 B. VISTA",
      "PERMUTA LOT. ROYAL PARK SCHROEDER E SCHM",
      "PERMUTA LIVINA EBERT",
      "PERMUTA MINI CARREGADEIRA VILA DAS NAÇÕE",
      "PERMUTA CASA MIA AA 62 CS2",
      "PERMUTA MICHELSON GOLDEN PARK",
      "PERMUTA MOTO ELÉTRICA GANHA SORTEIO DISB",
      "PERMUTA MR VIDROS GOLDEN PARK",
      "PERMUTA PARQUE DAS AREIAS LT 34 CASA MIA",
      "PERMUTA PARQUE DAS AREIAS LOTE 8 MÁRCIO",
      "PERMUTA PB SALA 5 DISBRACON",
      "PERMUTA PB SALA 6 WJ",
      "PERMUTA CONCRETOREI PB SALA A DEFINIR",
      "PERMUTA CASA MIA PB SALA 4",
      "PERMUTA PORTARIS KW4 SALA 1",
      "PERMUTA PORTARIS SEVEN POINT SALA 11",
      "PERMUTA RANGER OPA WJ",
      "PERMUTA GREEN GARDEN ROKA LOTE 29",
      "PERMUTA ROKA PARQUE DAS AREIAS LOTE 25",
      "PERMUTA S10 CHEVROLET LOT. ROYAL PARK",
      "PERMUTA RUI ELÉTRICA SEVEN POINT SALA 12",
      "PERMUTA RUI ELETRICISTA GG 33, 34 E 35",
      "PERMUTA TOMIO SEVEN POINT SALA 3",
      "PERMUTA TOLEDO 2 CONCRETO DJF CONCRETOS",
      "PERMUTA TOYOTA EBERT",
      "PERMUTA WJ RB CS3",
      "PERMUTA WJ SEVEN POINT SALA 2",
      "PERMUTA WARNOW PARK 1 INDAPAV",
      "PERMUTA WARNOW PARK 2 BERTELLI",
      "SALDO DEVOLULÇAO G LIGHT",
      "TROCA DO KW 5 CASA 3 PELO TERRENO SARTOR",
      "PERMUTA TERRENO CARIJÓS NO BP1 08",
      "PERMUTA TERRENO HENRIQUE KUNEN EST. AREA",
      "PERMUTA TOMIO GREEN GARDEN LOTE 46"
    ],
    documentos: ['CONTRATO'],
    
    condicoesPagamento: ['Parcelas Mensais', 'Novo Parcelas Mensais'],
    
    flags: {
      imprimirParcelasReparceladas: false,
    },
    
    ordem: 'VL',
    
    processarLancamentos: 'CR',
  },
  {
    sheetName: 'Receita Reforços',
    pdfName: 'Receita Reforços',
    contasCorrente: [
      "Viacredi Laurentino",
      "CEF Laurentino",
      "Sicredi Doca",
      "Sicredi Laurentino",
      "Aplicação Viacredi Laurentino",
      "ADIANTAMENTO CASA DA LAJE 150 MIL",
      "ADIANTAMENTO 250 MIL PRA CASA DA LAJE",
      "Aplicação CEF Laurentino",
      "Aplicação Doca Sicredi",
      "Aplicação Sicredi Laurentino",
      "ADIANTAMENTO CASA DA LAJE 175K",
      "ADIANTAMENTO CASA DA LAJE 200K",
      "ADIANTAMENTO CASA DA LAJE 250K",
      "ADIANTAMENTO 95K PORTARIS",
      "CAIXA Laurentino",
      "Conta Babi - Não Conciliar",
      "Conta Doca",
      "Conta Equilíbrio",
      "ESTOQUE",
      "PERMUTA INDAPAV BP1 LOTE 13 CASA 1",
      "Lançamentos Passados",
      "PERMUTA DISBRACON LOTE 10 E 14 BV3",
      "PERMUTA AB3 MR VIDROS",
      "PERMUTA AB 4 E 5 MONDINI",
      "PERMUTA DISBRACON ACACIA 26",
      "PERMUTA ACÁCIAS POMERODE LOTE 8 MINI CAR",
      "ADIANTAMENTO 115K VIDRAÇARIA ITOUP",
      "PERMUTA AMS LAJES GOLDEN PARK",
      "PERMUTA AMS LAJES LOTE 18 ACACIA POMEROD",
      "PERMUTA GAIO SERPA",
      "PERMUTA BERTELLI GREEN GARDEN LOTE 36",
      "PERMUTA BERTELLI KW 3 SALA 1",
      "PERMUTA BERTELLI KW 3 SALA 2",
      "PERMUTA BERTELLI SEVEN POINT SALA 9",
      "PERMUTA BONGO 2020 EBERT",
      "PERMUTA MONDINI LOTE 09 CASA 2 BP1",
      "PERMUTA BERTELLI VILA DAS NAÇÕES LT 15",
      "PERMUTA BERTELLI VILA DAS NAÇÕES LT 14",
      "PERMUTA PORTARIS CA 90 CS 1",
      "PERMUTA CALHAS INDAIAL LT NOVA INDAIAL 4",
      "PERMUTA CIVIC EBERT",
      "PERMUTA CASA MIA SEVEN POINT SALA 6",
      "PERMUTA CASA MIA VILA DAS NAÇÕES LT 18",
      "PERMUTA CONCRETOREI GREEN GARDEN LOTE 45",
      "PERMUTA DAINOX BP1 TERRENO MAT. 21884",
      "PERMUTA DAINOX SALA 7 SEVEN POINT",
      "PERMUTA DAINOX GOLDEN PARK LOT A DEF.",
      "PERMUTA DAMACENO",
      "PERMUTA DAMACENO LAND DO ALMIR",
      "PERMUTA DISBRACON ACÁCIA POMERODE LOTE16",
      "PERMUTA DISBRACON ACÁCIA POMERODE LOTE17",
      "PERMUTA DISBRACON SEVEN POINT SALA 4",
      "PERMUTA DISBRACON SEVEN POINT SALA 5",
      "PERMUTA EBERT ONIX",
      "PERMUTA ECO SPORT DAMACENO",
      "PERMUTA FCF RENEGADE (TROCOU DO PA 28)",
      "PERMUTA FIESTA (PEGEOUT) EBERT",
      "PERMUTA CASA MIA GARTEN DORF LOTE 1 E 19",
      "PERMUTA GG 21 DAINOX",
      "PERMUTA CASA DA LAJE GREEN GARDEN LT 30",
      "PERMUTA PORTARIS GG 31",
      "PERMUTA GREEN GARDEN GRAMEIRA FELIPPI 27",
      "PERMUTA GREEN GARDEN MULDE JANDT",
      "Permuta Gol Kosmos",
      "PERMUTA GRAMEIRA FELIPPI GOLDEN PARK",
      "PERMUTA GRAMEIRA FELIPPI GOLDEN PARK 2",
      "PERMUTA GRAMEIRA F. TOLEDO (TROCOU RY12)",
      "PERMUTA INDAPAV SALA 13 SEVEN POINT",
      "PERMUTA INDUMADER GREEN GARDEN 38",
      "PERMUTA INDUMADER GOLDEN PARK 2 LOTES",
      "PERMUTA INDUMADER GREEN VILLAGE 2LT 540M",
      "PERMUTA INDUMADER NOVA A DEFINIR",
      "PERMUTA INDUMADER GOLDEN PARK 3 LOTES",
      "PERMUTA INDUMADER ROYAL PARK 4/5/6",
      "PERMUTA INDUMADER ROYAL PARK 7/8/9",
      "PERMUTA INDAPAV WARNOW 12X80",
      "PERMUTA DAINOX GG LOTE 19",
      "PERMUTA INDAPAV VOLVO GARTEN DORF",
      "PERMUTA VIDRAÇARIA ITOUPAVAZINHA GP1 CS1",
      "PERMUTA JOURNEY EBERT",
      "PERMUTA KOSMOS GREEN GARDEN LOTE 23",
      "PERMUTAS KOSMOS SEVEN POINT SALA 10",
      "CASA DA LAJE ADIANTAMENTO 357K",
      "PERMUTA CASA DA LAJE KW 4 SALA 02",
      "PERMUTA CASA DA LAJE ACACIA POMEROD LT20",
      "PERMUTA CASA DA LAJE ROYAL PARK LT 13/14",
      "PERMUTA CASA DA LAJE SP SALA 8",
      "PERMUTA LF CALHAS GOLDEN PARK A DEF.",
      "PERMUTA LF CALHAS GREEN VILLAGE A DEF.",
      "PERMUTA LF CALHAS SEVEN POINT SALA 1",
      "PERMUTA LF CALHAS RUA TOLEDO",
      "PERMUTA CASA DA LAJE LOTE 03 BV3",
      "PERMUTA CASA DA LAJE LOTE 13 BV",
      "PERMUTA INDAPAV PC 35",
      "PERMUTA CASA DA LAJE LOTE 49 CA",
      "PERMUTA CALHAS INDAIAL LOTE 72 AA",
      "PERMUTA BERTELLI LOTE 10 BP1",
      "PERMUTA VIDRAÇARIA ITOUPAVAZINHA LT11BP1",
      "PERMUTA VIDR. ITOUPAZAVINHA 55 CA",
      "PERMUTA VIDR. ITOUPAVAZINHA 7/8 B. VISTA",
      "PERMUTA LOT. ROYAL PARK SCHROEDER E SCHM",
      "PERMUTA LIVINA EBERT",
      "PERMUTA MINI CARREGADEIRA VILA DAS NAÇÕE",
      "PERMUTA CASA MIA AA 62 CS2",
      "PERMUTA MICHELSON GOLDEN PARK",
      "PERMUTA MOTO ELÉTRICA GANHA SORTEIO DISB",
      "PERMUTA MR VIDROS GOLDEN PARK",
      "PERMUTA PARQUE DAS AREIAS LT 34 CASA MIA",
      "PERMUTA PARQUE DAS AREIAS LOTE 8 MÁRCIO",
      "PERMUTA PB SALA 5 DISBRACON",
      "PERMUTA PB SALA 6 WJ",
      "PERMUTA CONCRETOREI PB SALA A DEFINIR",
      "PERMUTA CASA MIA PB SALA 4",
      "PERMUTA PORTARIS KW4 SALA 1",
      "PERMUTA PORTARIS SEVEN POINT SALA 11",
      "PERMUTA RANGER OPA WJ",
      "PERMUTA GREEN GARDEN ROKA LOTE 29",
      "PERMUTA ROKA PARQUE DAS AREIAS LOTE 25",
      "PERMUTA S10 CHEVROLET LOT. ROYAL PARK",
      "PERMUTA RUI ELÉTRICA SEVEN POINT SALA 12",
      "PERMUTA RUI ELETRICISTA GG 33, 34 E 35",
      "PERMUTA TOMIO SEVEN POINT SALA 3",
      "PERMUTA TOLEDO 2 CONCRETO DJF CONCRETOS",
      "PERMUTA TOYOTA EBERT",
      "PERMUTA WJ RB CS3",
      "PERMUTA WJ SEVEN POINT SALA 2",
      "PERMUTA WARNOW PARK 1 INDAPAV",
      "PERMUTA WARNOW PARK 2 BERTELLI",
      "SALDO DEVOLULÇAO G LIGHT",
      "TROCA DO KW 5 CASA 3 PELO TERRENO SARTOR",
      "PERMUTA TERRENO CARIJÓS NO BP1 08",
      "PERMUTA TERRENO HENRIQUE KUNEN EST. AREA",
      "PERMUTA TOMIO GREEN GARDEN LOTE 46"
    ],
    planoFinanceiro: [
    "Receita de Incorporação de Imóveis",
    "Receita de Bens/Imóveis Adquiridos de Terceiros",
    "Descontos Concedidos (Unidades Imobiliárias)",
    "xxxxxReceita de Aluguel (Locação)",
    "Permuta Externa",
    "Receita de Cliente Desconhecido",
    "Receita de Administração (Taxa Administrativa)",
    "Receita Empreitada Só Mão de Obra",
    "Receita de Projeto",
    "Receita de Locação (Aluguel)",
    "Entrada de Financiamento para Construção",
    "Entrada de Capital de Giro",
    "Receita de Aplicações Financeiras",
    "Receita de Empréstimos",
    "Aporte Capital de Sócio",
    "Descontos Obtidos",
    "Receita de Venda de Imobilizado",
    "Receita Aluguel PCS 35",
    "Estoque de Material",
    "Receita Energia Solar Locatários",
    "Venda de Passivos",
    "(-) Anulação de Custos",
    "(-) Anulação de Custos Mão de Obra",
    "(-) Retenção Caução/Sinal",
    "(-) Retenção de INSS",
    "(-) Retenção de ISS",
    "(-) Retenção de PIS",
    "(-) Retenção de CSLL",
    "(-) Retenção de COFINS",
    "(-) Retenção de IR",
    "(-) Retenção (PIS/COFINS/CSLL - 4,65%)",
    "(-) Retenção de Permuta",
    "(-) Transf. Materiais Estoque Laurentino P/ Finan",
    "(-) Transf. Materiais Estoque Laurentino P/ Bom J.",
    "(-) Transf. Mat. Estoque P/ Obras da Laurentino",
    "(-) Reembolso de Despesas Gerais (Obra)",
    "(-) Divisão de muro",
    "(-) Retenção Contribuição Sindical (Escritório)",
    "(-) Reembolso de Despesas Administrativas",
    "(-) Restituição IPTU Imóveis Alugados (Escritório)",
    "(-) Reembolso de Despesas Financeiras",
    "(-) Reembolso de Despesas Tributárias",
    "(-) Abatimento de Adiantamento",
    "(-) Restituição Condomínio",
    "(-) Restituição água"
],
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
    flags: {
      imprimirParcelasReparceladas: false,
    },
    
    ordem: 'VL',
    
    processarLancamentos: 'CR',
  },
  {
    sheetName: 'Receita Aluguel',
    pdfName: 'Receita Aluguel',
    planoFinanceiro: [
    "Receita de Incorporação de Imóveis",
    "Receita de Bens/Imóveis Adquiridos de Terceiros",
    "Descontos Concedidos (Unidades Imobiliárias)",
    "xxxxxReceita de Aluguel (Locação)",
    "Permuta Externa",
    "Receita de Cliente Desconhecido",
    "Receita de Administração (Taxa Administrativa)",
    "Receita Empreitada Só Mão de Obra",
    "Receita de Projeto",
    "Receita de Locação (Aluguel)",
    "Entrada de Financiamento para Construção",
    "Entrada de Capital de Giro",
    "Receita de Aplicações Financeiras",
    "Receita de Empréstimos",
    "Aporte Capital de Sócio",
    "Descontos Obtidos",
    "Receita de Venda de Imobilizado",
    "Receita Aluguel PCS 35",
    "Estoque de Material",
    "Receita Energia Solar Locatários",
    "Venda de Passivos",
    "(-) Anulação de Custos",
    "(-) Anulação de Custos Mão de Obra",
    "(-) Retenção Caução/Sinal",
    "(-) Retenção de INSS",
    "(-) Retenção de ISS",
    "(-) Retenção de PIS",
    "(-) Retenção de CSLL",
    "(-) Retenção de COFINS",
    "(-) Retenção de IR",
    "(-) Retenção (PIS/COFINS/CSLL - 4,65%)",
    "(-) Retenção de Permuta",
    "(-) Transf. Materiais Estoque Laurentino P/ Finan",
    "(-) Transf. Materiais Estoque Laurentino P/ Bom J.",
    "(-) Transf. Mat. Estoque P/ Obras da Laurentino",
    "(-) Reembolso de Despesas Gerais (Obra)",
    "(-) Divisão de muro",
    "(-) Retenção Contribuição Sindical (Escritório)",
    "(-) Reembolso de Despesas Administrativas",
    "(-) Restituição IPTU Imóveis Alugados (Escritório)",
    "(-) Reembolso de Despesas Financeiras",
    "(-) Reembolso de Despesas Tributárias",
    "(-) Abatimento de Adiantamento",
    "(-) Restituição Condomínio",
    "(-) Restituição água"
],
    documentos: [
      "Adiantamento",
      "APORTE DE CAPITAL",
      "AVISO DE LANÇAMENTO",
      "CAUÇÃO",
      "CONHECIMENTO DE FRETE",
      "CONTRATO",
      "CUPOM FISCAL",
      "Documento de Arrecadação de Receita Est",
      "Documento de arrecadação federal",
      "Documento de Arrecadação Municipal",
      "Doc. de Arrecadação do Simples Nacional",
      "Documento de arrecadação simples naciona",
      "Distrato",
      "FATURA",
      "FOLHA DE PAGAMENTO",
      "FATURAMENTO DIRETO",
      "Guia do FGTS Digital",
      "Guia de Previdência Social",
      "GUIA RECOLHIMENTO DE COMISSÃO SINDICAL",
      "Guia de recolhimento de FGTS",
      "GUIA PARA RECOLHIMENTO DE IMPOSTOS",
      "INVISÍVEL",
      "INICIALIZAÇÃO DE ESTOQUE",
      "DOC. CONTAB. SISTEMA CUSTO ORÇ. E INCOR",
      "Movimento de Frotas e Equipamento",
      "Movimento de Locação de Equipamentos",
      "Movimento de transf. de equipamento",
      "MOVIMENTO SEGURANÇA SAUDE NO TRABALHO",
      "NOTA FISCAL",
      "NOTA FISCAL ELETRÔNICA",
      "Nota Fiscal de Serviço",
      "PREVISÃO FINANCEIRA DE CONTR. DE MED.",
      "PERMUTA",
      "PREVISÃO FINANCEIRA DE PEDIDOS DE COMPRA",
      "PREVISÃO DE COMISSÃO",
      "PREVISÃO DE DISTRATO",
      "PREVISÃO DE PAGAMENTO/RECEBIMENTO",
      "REALOCAÇÃO DE MATERIAIS",
      "RECIBO",
      "REQUISIÇÃO DE MATERIAIS",
      "SN",
      "DOCUMENTO DE TRANSFERÊNCIA DE MATERIAIS",
      "TRANSF./RESERVA DE ESTOQUE P/SOLICITAÇÃO",
      "TRANSF. DE MATERIAIS POR SOLICITAÇÃO",
      "TAXAS"
    ],
    condicoesPagamento: [
      'Novo Parcelas Mensais Aluguel',
      'Parcela Locação Aluguel Mensal',
    ],
    contasCorrente: [
      "Viacredi Laurentino",
      "CEF Laurentino",
      "Sicredi Doca",
      "Sicredi Laurentino",
      "Aplicação Viacredi Laurentino",
      "ADIANTAMENTO CASA DA LAJE 150 MIL",
      "ADIANTAMENTO 250 MIL PRA CASA DA LAJE",
      "Aplicação CEF Laurentino",
      "Aplicação Doca Sicredi",
      "Aplicação Sicredi Laurentino",
      "ADIANTAMENTO CASA DA LAJE 175K",
      "ADIANTAMENTO CASA DA LAJE 200K",
      "ADIANTAMENTO CASA DA LAJE 250K",
      "ADIANTAMENTO 95K PORTARIS",
      "CAIXA Laurentino",
      "Conta Babi - Não Conciliar",
      "Conta Doca",
      "Conta Equilíbrio",
      "ESTOQUE",
      "PERMUTA INDAPAV BP1 LOTE 13 CASA 1",
      "Lançamentos Passados",
      "PERMUTA DISBRACON LOTE 10 E 14 BV3",
      "PERMUTA AB3 MR VIDROS",
      "PERMUTA AB 4 E 5 MONDINI",
      "PERMUTA DISBRACON ACACIA 26",
      "PERMUTA ACÁCIAS POMERODE LOTE 8 MINI CAR",
      "ADIANTAMENTO 115K VIDRAÇARIA ITOUP",
      "PERMUTA AMS LAJES GOLDEN PARK",
      "PERMUTA AMS LAJES LOTE 18 ACACIA POMEROD",
      "PERMUTA GAIO SERPA",
      "PERMUTA BERTELLI GREEN GARDEN LOTE 36",
      "PERMUTA BERTELLI KW 3 SALA 1",
      "PERMUTA BERTELLI KW 3 SALA 2",
      "PERMUTA BERTELLI SEVEN POINT SALA 9",
      "PERMUTA BONGO 2020 EBERT",
      "PERMUTA MONDINI LOTE 09 CASA 2 BP1",
      "PERMUTA BERTELLI VILA DAS NAÇÕES LT 15",
      "PERMUTA BERTELLI VILA DAS NAÇÕES LT 14",
      "PERMUTA PORTARIS CA 90 CS 1",
      "PERMUTA CALHAS INDAIAL LT NOVA INDAIAL 4",
      "PERMUTA CIVIC EBERT",
      "PERMUTA CASA MIA SEVEN POINT SALA 6",
      "PERMUTA CASA MIA VILA DAS NAÇÕES LT 18",
      "PERMUTA CONCRETOREI GREEN GARDEN LOTE 45",
      "PERMUTA DAINOX BP1 TERRENO MAT. 21884",
      "PERMUTA DAINOX SALA 7 SEVEN POINT",
      "PERMUTA DAINOX GOLDEN PARK LOT A DEF.",
      "PERMUTA DAMACENO",
      "PERMUTA DAMACENO LAND DO ALMIR",
      "PERMUTA DISBRACON ACÁCIA POMERODE LOTE16",
      "PERMUTA DISBRACON ACÁCIA POMERODE LOTE17",
      "PERMUTA DISBRACON SEVEN POINT SALA 4",
      "PERMUTA DISBRACON SEVEN POINT SALA 5",
      "PERMUTA EBERT ONIX",
      "PERMUTA ECO SPORT DAMACENO",
      "PERMUTA FCF RENEGADE (TROCOU DO PA 28)",
      "PERMUTA FIESTA (PEGEOUT) EBERT",
      "PERMUTA CASA MIA GARTEN DORF LOTE 1 E 19",
      "PERMUTA GG 21 DAINOX",
      "PERMUTA CASA DA LAJE GREEN GARDEN LT 30",
      "PERMUTA PORTARIS GG 31",
      "PERMUTA GREEN GARDEN GRAMEIRA FELIPPI 27",
      "PERMUTA GREEN GARDEN MULDE JANDT",
      "Permuta Gol Kosmos",
      "PERMUTA GRAMEIRA FELIPPI GOLDEN PARK",
      "PERMUTA GRAMEIRA FELIPPI GOLDEN PARK 2",
      "PERMUTA GRAMEIRA F. TOLEDO (TROCOU RY12)",
      "PERMUTA INDAPAV SALA 13 SEVEN POINT",
      "PERMUTA INDUMADER GREEN GARDEN 38",
      "PERMUTA INDUMADER GOLDEN PARK 2 LOTES",
      "PERMUTA INDUMADER GREEN VILLAGE 2LT 540M",
      "PERMUTA INDUMADER NOVA A DEFINIR",
      "PERMUTA INDUMADER GOLDEN PARK 3 LOTES",
      "PERMUTA INDUMADER ROYAL PARK 4/5/6",
      "PERMUTA INDUMADER ROYAL PARK 7/8/9",
      "PERMUTA INDAPAV WARNOW 12X80",
      "PERMUTA DAINOX GG LOTE 19",
      "PERMUTA INDAPAV VOLVO GARTEN DORF",
      "PERMUTA VIDRAÇARIA ITOUPAVAZINHA GP1 CS1",
      "PERMUTA JOURNEY EBERT",
      "PERMUTA KOSMOS GREEN GARDEN LOTE 23",
      "PERMUTAS KOSMOS SEVEN POINT SALA 10",
      "CASA DA LAJE ADIANTAMENTO 357K",
      "PERMUTA CASA DA LAJE KW 4 SALA 02",
      "PERMUTA CASA DA LAJE ACACIA POMEROD LT20",
      "PERMUTA CASA DA LAJE ROYAL PARK LT 13/14",
      "PERMUTA CASA DA LAJE SP SALA 8",
      "PERMUTA LF CALHAS GOLDEN PARK A DEF.",
      "PERMUTA LF CALHAS GREEN VILLAGE A DEF.",
      "PERMUTA LF CALHAS SEVEN POINT SALA 1",
      "PERMUTA LF CALHAS RUA TOLEDO",
      "PERMUTA CASA DA LAJE LOTE 03 BV3",
      "PERMUTA CASA DA LAJE LOTE 13 BV",
      "PERMUTA INDAPAV PC 35",
      "PERMUTA CASA DA LAJE LOTE 49 CA",
      "PERMUTA CALHAS INDAIAL LOTE 72 AA",
      "PERMUTA BERTELLI LOTE 10 BP1",
      "PERMUTA VIDRAÇARIA ITOUPAVAZINHA LT11BP1",
      "PERMUTA VIDR. ITOUPAZAVINHA 55 CA",
      "PERMUTA VIDR. ITOUPAVAZINHA 7/8 B. VISTA",
      "PERMUTA LOT. ROYAL PARK SCHROEDER E SCHM",
      "PERMUTA LIVINA EBERT",
      "PERMUTA MINI CARREGADEIRA VILA DAS NAÇÕE",
      "PERMUTA CASA MIA AA 62 CS2",
      "PERMUTA MICHELSON GOLDEN PARK",
      "PERMUTA MOTO ELÉTRICA GANHA SORTEIO DISB",
      "PERMUTA MR VIDROS GOLDEN PARK",
      "PERMUTA PARQUE DAS AREIAS LT 34 CASA MIA",
      "PERMUTA PARQUE DAS AREIAS LOTE 8 MÁRCIO",
      "PERMUTA PB SALA 5 DISBRACON",
      "PERMUTA PB SALA 6 WJ",
      "PERMUTA CONCRETOREI PB SALA A DEFINIR",
      "PERMUTA CASA MIA PB SALA 4",
      "PERMUTA PORTARIS KW4 SALA 1",
      "PERMUTA PORTARIS SEVEN POINT SALA 11",
      "PERMUTA RANGER OPA WJ",
      "PERMUTA GREEN GARDEN ROKA LOTE 29",
      "PERMUTA ROKA PARQUE DAS AREIAS LOTE 25",
      "PERMUTA S10 CHEVROLET LOT. ROYAL PARK",
      "PERMUTA RUI ELÉTRICA SEVEN POINT SALA 12",
      "PERMUTA RUI ELETRICISTA GG 33, 34 E 35",
      "PERMUTA TOMIO SEVEN POINT SALA 3",
      "PERMUTA TOLEDO 2 CONCRETO DJF CONCRETOS",
      "PERMUTA TOYOTA EBERT",
      "PERMUTA WJ RB CS3",
      "PERMUTA WJ SEVEN POINT SALA 2",
      "PERMUTA WARNOW PARK 1 INDAPAV",
      "PERMUTA WARNOW PARK 2 BERTELLI",
      "SALDO DEVOLULÇAO G LIGHT",
      "TROCA DO KW 5 CASA 3 PELO TERRENO SARTOR",
      "PERMUTA TERRENO CARIJÓS NO BP1 08",
      "PERMUTA TERRENO HENRIQUE KUNEN EST. AREA",
      "PERMUTA TOMIO GREEN GARDEN LOTE 46"
    ],
    flags: {
      imprimirParcelasReparceladas: false,
    },
    
    ordem: 'VL',
    
    processarLancamentos: 'CR',
  },
  {
    sheetName: 'Ato + PE',
    pdfName: 'Ato + PE',
    contasCorrente: [
      "Viacredi Laurentino",
      "CEF Laurentino",
      "Sicredi Doca",
      "Sicredi Laurentino",
      "Aplicação Viacredi Laurentino",
      "ADIANTAMENTO CASA DA LAJE 150 MIL",
      "ADIANTAMENTO 250 MIL PRA CASA DA LAJE",
      "Aplicação CEF Laurentino",
      "Aplicação Doca Sicredi",
      "Aplicação Sicredi Laurentino",
      "ADIANTAMENTO CASA DA LAJE 175K",
      "ADIANTAMENTO CASA DA LAJE 200K",
      "ADIANTAMENTO CASA DA LAJE 250K",
      "ADIANTAMENTO 95K PORTARIS",
      "CAIXA Laurentino",
      "Conta Babi - Não Conciliar",
      "Conta Doca",
      "Conta Equilíbrio",
      "ESTOQUE",
      "PERMUTA INDAPAV BP1 LOTE 13 CASA 1",
      "Lançamentos Passados",
      "PERMUTA DISBRACON LOTE 10 E 14 BV3",
      "PERMUTA AB3 MR VIDROS",
      "PERMUTA AB 4 E 5 MONDINI",
      "PERMUTA DISBRACON ACACIA 26",
      "PERMUTA ACÁCIAS POMERODE LOTE 8 MINI CAR",
      "ADIANTAMENTO 115K VIDRAÇARIA ITOUP",
      "PERMUTA AMS LAJES GOLDEN PARK",
      "PERMUTA AMS LAJES LOTE 18 ACACIA POMEROD",
      "PERMUTA GAIO SERPA",
      "PERMUTA BERTELLI GREEN GARDEN LOTE 36",
      "PERMUTA BERTELLI KW 3 SALA 1",
      "PERMUTA BERTELLI KW 3 SALA 2",
      "PERMUTA BERTELLI SEVEN POINT SALA 9",
      "PERMUTA BONGO 2020 EBERT",
      "PERMUTA MONDINI LOTE 09 CASA 2 BP1",
      "PERMUTA BERTELLI VILA DAS NAÇÕES LT 15",
      "PERMUTA BERTELLI VILA DAS NAÇÕES LT 14",
      "PERMUTA PORTARIS CA 90 CS 1",
      "PERMUTA CALHAS INDAIAL LT NOVA INDAIAL 4",
      "PERMUTA CIVIC EBERT",
      "PERMUTA CASA MIA SEVEN POINT SALA 6",
      "PERMUTA CASA MIA VILA DAS NAÇÕES LT 18",
      "PERMUTA CONCRETOREI GREEN GARDEN LOTE 45",
      "PERMUTA DAINOX BP1 TERRENO MAT. 21884",
      "PERMUTA DAINOX SALA 7 SEVEN POINT",
      "PERMUTA DAINOX GOLDEN PARK LOT A DEF.",
      "PERMUTA DAMACENO",
      "PERMUTA DAMACENO LAND DO ALMIR",
      "PERMUTA DISBRACON ACÁCIA POMERODE LOTE16",
      "PERMUTA DISBRACON ACÁCIA POMERODE LOTE17",
      "PERMUTA DISBRACON SEVEN POINT SALA 4",
      "PERMUTA DISBRACON SEVEN POINT SALA 5",
      "PERMUTA EBERT ONIX",
      "PERMUTA ECO SPORT DAMACENO",
      "PERMUTA FCF RENEGADE (TROCOU DO PA 28)",
      "PERMUTA FIESTA (PEGEOUT) EBERT",
      "PERMUTA CASA MIA GARTEN DORF LOTE 1 E 19",
      "PERMUTA GG 21 DAINOX",
      "PERMUTA CASA DA LAJE GREEN GARDEN LT 30",
      "PERMUTA PORTARIS GG 31",
      "PERMUTA GREEN GARDEN GRAMEIRA FELIPPI 27",
      "PERMUTA GREEN GARDEN MULDE JANDT",
      "Permuta Gol Kosmos",
      "PERMUTA GRAMEIRA FELIPPI GOLDEN PARK",
      "PERMUTA GRAMEIRA FELIPPI GOLDEN PARK 2",
      "PERMUTA GRAMEIRA F. TOLEDO (TROCOU RY12)",
      "PERMUTA INDAPAV SALA 13 SEVEN POINT",
      "PERMUTA INDUMADER GREEN GARDEN 38",
      "PERMUTA INDUMADER GOLDEN PARK 2 LOTES",
      "PERMUTA INDUMADER GREEN VILLAGE 2LT 540M",
      "PERMUTA INDUMADER NOVA A DEFINIR",
      "PERMUTA INDUMADER GOLDEN PARK 3 LOTES",
      "PERMUTA INDUMADER ROYAL PARK 4/5/6",
      "PERMUTA INDUMADER ROYAL PARK 7/8/9",
      "PERMUTA INDAPAV WARNOW 12X80",
      "PERMUTA DAINOX GG LOTE 19",
      "PERMUTA INDAPAV VOLVO GARTEN DORF",
      "PERMUTA VIDRAÇARIA ITOUPAVAZINHA GP1 CS1",
      "PERMUTA JOURNEY EBERT",
      "PERMUTA KOSMOS GREEN GARDEN LOTE 23",
      "PERMUTAS KOSMOS SEVEN POINT SALA 10",
      "CASA DA LAJE ADIANTAMENTO 357K",
      "PERMUTA CASA DA LAJE KW 4 SALA 02",
      "PERMUTA CASA DA LAJE ACACIA POMEROD LT20",
      "PERMUTA CASA DA LAJE ROYAL PARK LT 13/14",
      "PERMUTA CASA DA LAJE SP SALA 8",
      "PERMUTA LF CALHAS GOLDEN PARK A DEF.",
      "PERMUTA LF CALHAS GREEN VILLAGE A DEF.",
      "PERMUTA LF CALHAS SEVEN POINT SALA 1",
      "PERMUTA LF CALHAS RUA TOLEDO",
      "PERMUTA CASA DA LAJE LOTE 03 BV3",
      "PERMUTA CASA DA LAJE LOTE 13 BV",
      "PERMUTA INDAPAV PC 35",
      "PERMUTA CASA DA LAJE LOTE 49 CA",
      "PERMUTA CALHAS INDAIAL LOTE 72 AA",
      "PERMUTA BERTELLI LOTE 10 BP1",
      "PERMUTA VIDRAÇARIA ITOUPAVAZINHA LT11BP1",
      "PERMUTA VIDR. ITOUPAZAVINHA 55 CA",
      "PERMUTA VIDR. ITOUPAVAZINHA 7/8 B. VISTA",
      "PERMUTA LOT. ROYAL PARK SCHROEDER E SCHM",
      "PERMUTA LIVINA EBERT",
      "PERMUTA MINI CARREGADEIRA VILA DAS NAÇÕE",
      "PERMUTA CASA MIA AA 62 CS2",
      "PERMUTA MICHELSON GOLDEN PARK",
      "PERMUTA MOTO ELÉTRICA GANHA SORTEIO DISB",
      "PERMUTA MR VIDROS GOLDEN PARK",
      "PERMUTA PARQUE DAS AREIAS LT 34 CASA MIA",
      "PERMUTA PARQUE DAS AREIAS LOTE 8 MÁRCIO",
      "PERMUTA PB SALA 5 DISBRACON",
      "PERMUTA PB SALA 6 WJ",
      "PERMUTA CONCRETOREI PB SALA A DEFINIR",
      "PERMUTA CASA MIA PB SALA 4",
      "PERMUTA PORTARIS KW4 SALA 1",
      "PERMUTA PORTARIS SEVEN POINT SALA 11",
      "PERMUTA RANGER OPA WJ",
      "PERMUTA GREEN GARDEN ROKA LOTE 29",
      "PERMUTA ROKA PARQUE DAS AREIAS LOTE 25",
      "PERMUTA S10 CHEVROLET LOT. ROYAL PARK",
      "PERMUTA RUI ELÉTRICA SEVEN POINT SALA 12",
      "PERMUTA RUI ELETRICISTA GG 33, 34 E 35",
      "PERMUTA TOMIO SEVEN POINT SALA 3",
      "PERMUTA TOLEDO 2 CONCRETO DJF CONCRETOS",
      "PERMUTA TOYOTA EBERT",
      "PERMUTA WJ RB CS3",
      "PERMUTA WJ SEVEN POINT SALA 2",
      "PERMUTA WARNOW PARK 1 INDAPAV",
      "PERMUTA WARNOW PARK 2 BERTELLI",
      "SALDO DEVOLULÇAO G LIGHT",
      "TROCA DO KW 5 CASA 3 PELO TERRENO SARTOR",
      "PERMUTA TERRENO CARIJÓS NO BP1 08",
      "PERMUTA TERRENO HENRIQUE KUNEN EST. AREA",
      "PERMUTA TOMIO GREEN GARDEN LOTE 46"
    ],
    planoFinanceiro: [
    "Receita de Incorporação de Imóveis",
    "Receita de Bens/Imóveis Adquiridos de Terceiros",
    "Descontos Concedidos (Unidades Imobiliárias)",
    "xxxxxReceita de Aluguel (Locação)",
    "Permuta Externa",
    "Receita de Cliente Desconhecido",
    "Receita de Administração (Taxa Administrativa)",
    "Receita Empreitada Só Mão de Obra",
    "Receita de Projeto",
    "Receita de Locação (Aluguel)",
    "Entrada de Financiamento para Construção",
    "Entrada de Capital de Giro",
    "Receita de Aplicações Financeiras",
    "Receita de Empréstimos",
    "Aporte Capital de Sócio",
    "Descontos Obtidos",
    "Receita de Venda de Imobilizado",
    "Receita Aluguel PCS 35",
    "Estoque de Material",
    "Receita Energia Solar Locatários",
    "Venda de Passivos",
    "(-) Anulação de Custos",
    "(-) Anulação de Custos Mão de Obra",
    "(-) Retenção Caução/Sinal",
    "(-) Retenção de INSS",
    "(-) Retenção de ISS",
    "(-) Retenção de PIS",
    "(-) Retenção de CSLL",
    "(-) Retenção de COFINS",
    "(-) Retenção de IR",
    "(-) Retenção (PIS/COFINS/CSLL - 4,65%)",
    "(-) Retenção de Permuta",
    "(-) Transf. Materiais Estoque Laurentino P/ Finan",
    "(-) Transf. Materiais Estoque Laurentino P/ Bom J.",
    "(-) Transf. Mat. Estoque P/ Obras da Laurentino",
    "(-) Reembolso de Despesas Gerais (Obra)",
    "(-) Divisão de muro",
    "(-) Retenção Contribuição Sindical (Escritório)",
    "(-) Reembolso de Despesas Administrativas",
    "(-) Restituição IPTU Imóveis Alugados (Escritório)",
    "(-) Reembolso de Despesas Financeiras",
    "(-) Reembolso de Despesas Tributárias",
    "(-) Abatimento de Adiantamento",
    "(-) Restituição Condomínio",
    "(-) Restituição água"
],
    documentos: ['CONTRATO'],
    condicoesPagamento: [
      'Parcela na Escritura', 
      'Novo Parcela na Escritura', 
      'Ato', 
      'Novo Ato'
    ],
    flags: {
      imprimirParcelasReparceladas: false,
    },
    
    ordem: 'VL',
    
    processarLancamentos: 'CR',
  },
  {
    sheetName: 'Venda a Vista',
    pdfName: 'Venda a Vista',
    planoFinanceiro: [
    "Receita de Incorporação de Imóveis",
    "Receita de Bens/Imóveis Adquiridos de Terceiros",
    "Descontos Concedidos (Unidades Imobiliárias)",
    "xxxxxReceita de Aluguel (Locação)",
    "Permuta Externa",
    "Receita de Cliente Desconhecido",
    "Receita de Administração (Taxa Administrativa)",
    "Receita Empreitada Só Mão de Obra",
    "Receita de Projeto",
    "Receita de Locação (Aluguel)",
    "Entrada de Financiamento para Construção",
    "Entrada de Capital de Giro",
    "Receita de Aplicações Financeiras",
    "Receita de Empréstimos",
    "Aporte Capital de Sócio",
    "Descontos Obtidos",
    "Receita de Venda de Imobilizado",
    "Receita Aluguel PCS 35",
    "Estoque de Material",
    "Receita Energia Solar Locatários",
    "Venda de Passivos",
    "(-) Anulação de Custos",
    "(-) Anulação de Custos Mão de Obra",
    "(-) Retenção Caução/Sinal",
    "(-) Retenção de INSS",
    "(-) Retenção de ISS",
    "(-) Retenção de PIS",
    "(-) Retenção de CSLL",
    "(-) Retenção de COFINS",
    "(-) Retenção de IR",
    "(-) Retenção (PIS/COFINS/CSLL - 4,65%)",
    "(-) Retenção de Permuta",
    "(-) Transf. Materiais Estoque Laurentino P/ Finan",
    "(-) Transf. Materiais Estoque Laurentino P/ Bom J.",
    "(-) Transf. Mat. Estoque P/ Obras da Laurentino",
    "(-) Reembolso de Despesas Gerais (Obra)",
    "(-) Divisão de muro",
    "(-) Retenção Contribuição Sindical (Escritório)",
    "(-) Reembolso de Despesas Administrativas",
    "(-) Restituição IPTU Imóveis Alugados (Escritório)",
    "(-) Reembolso de Despesas Financeiras",
    "(-) Reembolso de Despesas Tributárias",
    "(-) Abatimento de Adiantamento",
    "(-) Restituição Condomínio",
    "(-) Restituição água"
],
    documentos: ['CONTRATO'],
    condicoesPagamento: [
      'Venda a Vista', 
      'Novo Venda a Vista'
    ],
    contasCorrente: [
      "Viacredi Laurentino",
      "CEF Laurentino",
      "Sicredi Doca",
      "Sicredi Laurentino",
      "Aplicação Viacredi Laurentino",
      "ADIANTAMENTO CASA DA LAJE 150 MIL",
      "ADIANTAMENTO 250 MIL PRA CASA DA LAJE",
      "Aplicação CEF Laurentino",
      "Aplicação Doca Sicredi",
      "Aplicação Sicredi Laurentino",
      "ADIANTAMENTO CASA DA LAJE 175K",
      "ADIANTAMENTO CASA DA LAJE 200K",
      "ADIANTAMENTO CASA DA LAJE 250K",
      "ADIANTAMENTO 95K PORTARIS",
      "CAIXA Laurentino",
      "Conta Babi - Não Conciliar",
      "Conta Doca",
      "Conta Equilíbrio",
      "ESTOQUE",
      "PERMUTA INDAPAV BP1 LOTE 13 CASA 1",
      "Lançamentos Passados",
      "PERMUTA DISBRACON LOTE 10 E 14 BV3",
      "PERMUTA AB3 MR VIDROS",
      "PERMUTA AB 4 E 5 MONDINI",
      "PERMUTA DISBRACON ACACIA 26",
      "PERMUTA ACÁCIAS POMERODE LOTE 8 MINI CAR",
      "ADIANTAMENTO 115K VIDRAÇARIA ITOUP",
      "PERMUTA AMS LAJES GOLDEN PARK",
      "PERMUTA AMS LAJES LOTE 18 ACACIA POMEROD",
      "PERMUTA GAIO SERPA",
      "PERMUTA BERTELLI GREEN GARDEN LOTE 36",
      "PERMUTA BERTELLI KW 3 SALA 1",
      "PERMUTA BERTELLI KW 3 SALA 2",
      "PERMUTA BERTELLI SEVEN POINT SALA 9",
      "PERMUTA BONGO 2020 EBERT",
      "PERMUTA MONDINI LOTE 09 CASA 2 BP1",
      "PERMUTA BERTELLI VILA DAS NAÇÕES LT 15",
      "PERMUTA BERTELLI VILA DAS NAÇÕES LT 14",
      "PERMUTA PORTARIS CA 90 CS 1",
      "PERMUTA CALHAS INDAIAL LT NOVA INDAIAL 4",
      "PERMUTA CIVIC EBERT",
      "PERMUTA CASA MIA SEVEN POINT SALA 6",
      "PERMUTA CASA MIA VILA DAS NAÇÕES LT 18",
      "PERMUTA CONCRETOREI GREEN GARDEN LOTE 45",
      "PERMUTA DAINOX BP1 TERRENO MAT. 21884",
      "PERMUTA DAINOX SALA 7 SEVEN POINT",
      "PERMUTA DAINOX GOLDEN PARK LOT A DEF.",
      "PERMUTA DAMACENO",
      "PERMUTA DAMACENO LAND DO ALMIR",
      "PERMUTA DISBRACON ACÁCIA POMERODE LOTE16",
      "PERMUTA DISBRACON ACÁCIA POMERODE LOTE17",
      "PERMUTA DISBRACON SEVEN POINT SALA 4",
      "PERMUTA DISBRACON SEVEN POINT SALA 5",
      "PERMUTA EBERT ONIX",
      "PERMUTA ECO SPORT DAMACENO",
      "PERMUTA FCF RENEGADE (TROCOU DO PA 28)",
      "PERMUTA FIESTA (PEGEOUT) EBERT",
      "PERMUTA CASA MIA GARTEN DORF LOTE 1 E 19",
      "PERMUTA GG 21 DAINOX",
      "PERMUTA CASA DA LAJE GREEN GARDEN LT 30",
      "PERMUTA PORTARIS GG 31",
      "PERMUTA GREEN GARDEN GRAMEIRA FELIPPI 27",
      "PERMUTA GREEN GARDEN MULDE JANDT",
      "Permuta Gol Kosmos",
      "PERMUTA GRAMEIRA FELIPPI GOLDEN PARK",
      "PERMUTA GRAMEIRA FELIPPI GOLDEN PARK 2",
      "PERMUTA GRAMEIRA F. TOLEDO (TROCOU RY12)",
      "PERMUTA INDAPAV SALA 13 SEVEN POINT",
      "PERMUTA INDUMADER GREEN GARDEN 38",
      "PERMUTA INDUMADER GOLDEN PARK 2 LOTES",
      "PERMUTA INDUMADER GREEN VILLAGE 2LT 540M",
      "PERMUTA INDUMADER NOVA A DEFINIR",
      "PERMUTA INDUMADER GOLDEN PARK 3 LOTES",
      "PERMUTA INDUMADER ROYAL PARK 4/5/6",
      "PERMUTA INDUMADER ROYAL PARK 7/8/9",
      "PERMUTA INDAPAV WARNOW 12X80",
      "PERMUTA DAINOX GG LOTE 19",
      "PERMUTA INDAPAV VOLVO GARTEN DORF",
      "PERMUTA VIDRAÇARIA ITOUPAVAZINHA GP1 CS1",
      "PERMUTA JOURNEY EBERT",
      "PERMUTA KOSMOS GREEN GARDEN LOTE 23",
      "PERMUTAS KOSMOS SEVEN POINT SALA 10",
      "CASA DA LAJE ADIANTAMENTO 357K",
      "PERMUTA CASA DA LAJE KW 4 SALA 02",
      "PERMUTA CASA DA LAJE ACACIA POMEROD LT20",
      "PERMUTA CASA DA LAJE ROYAL PARK LT 13/14",
      "PERMUTA CASA DA LAJE SP SALA 8",
      "PERMUTA LF CALHAS GOLDEN PARK A DEF.",
      "PERMUTA LF CALHAS GREEN VILLAGE A DEF.",
      "PERMUTA LF CALHAS SEVEN POINT SALA 1",
      "PERMUTA LF CALHAS RUA TOLEDO",
      "PERMUTA CASA DA LAJE LOTE 03 BV3",
      "PERMUTA CASA DA LAJE LOTE 13 BV",
      "PERMUTA INDAPAV PC 35",
      "PERMUTA CASA DA LAJE LOTE 49 CA",
      "PERMUTA CALHAS INDAIAL LOTE 72 AA",
      "PERMUTA BERTELLI LOTE 10 BP1",
      "PERMUTA VIDRAÇARIA ITOUPAVAZINHA LT11BP1",
      "PERMUTA VIDR. ITOUPAZAVINHA 55 CA",
      "PERMUTA VIDR. ITOUPAVAZINHA 7/8 B. VISTA",
      "PERMUTA LOT. ROYAL PARK SCHROEDER E SCHM",
      "PERMUTA LIVINA EBERT",
      "PERMUTA MINI CARREGADEIRA VILA DAS NAÇÕE",
      "PERMUTA CASA MIA AA 62 CS2",
      "PERMUTA MICHELSON GOLDEN PARK",
      "PERMUTA MOTO ELÉTRICA GANHA SORTEIO DISB",
      "PERMUTA MR VIDROS GOLDEN PARK",
      "PERMUTA PARQUE DAS AREIAS LT 34 CASA MIA",
      "PERMUTA PARQUE DAS AREIAS LOTE 8 MÁRCIO",
      "PERMUTA PB SALA 5 DISBRACON",
      "PERMUTA PB SALA 6 WJ",
      "PERMUTA CONCRETOREI PB SALA A DEFINIR",
      "PERMUTA CASA MIA PB SALA 4",
      "PERMUTA PORTARIS KW4 SALA 1",
      "PERMUTA PORTARIS SEVEN POINT SALA 11",
      "PERMUTA RANGER OPA WJ",
      "PERMUTA GREEN GARDEN ROKA LOTE 29",
      "PERMUTA ROKA PARQUE DAS AREIAS LOTE 25",
      "PERMUTA S10 CHEVROLET LOT. ROYAL PARK",
      "PERMUTA RUI ELÉTRICA SEVEN POINT SALA 12",
      "PERMUTA RUI ELETRICISTA GG 33, 34 E 35",
      "PERMUTA TOMIO SEVEN POINT SALA 3",
      "PERMUTA TOLEDO 2 CONCRETO DJF CONCRETOS",
      "PERMUTA TOYOTA EBERT",
      "PERMUTA WJ RB CS3",
      "PERMUTA WJ SEVEN POINT SALA 2",
      "PERMUTA WARNOW PARK 1 INDAPAV",
      "PERMUTA WARNOW PARK 2 BERTELLI",
      "SALDO DEVOLULÇAO G LIGHT",
      "TROCA DO KW 5 CASA 3 PELO TERRENO SARTOR",
      "PERMUTA TERRENO CARIJÓS NO BP1 08",
      "PERMUTA TERRENO HENRIQUE KUNEN EST. AREA",
      "PERMUTA TOMIO GREEN GARDEN LOTE 46"
    ],
    flags: {
      imprimirParcelasReparceladas: false,
    },
    
    ordem: 'VL',
    
    processarLancamentos: 'CR',
  },
  {
    sheetName: 'Financiamento',
    pdfName: 'Financiamento',
    contasCorrente: [
      "Viacredi Laurentino",
      "CEF Laurentino",
      "Sicredi Doca",
      "Sicredi Laurentino",
      "Aplicação Viacredi Laurentino",
      "ADIANTAMENTO CASA DA LAJE 150 MIL",
      "ADIANTAMENTO 250 MIL PRA CASA DA LAJE",
      "Aplicação CEF Laurentino",
      "Aplicação Doca Sicredi",
      "Aplicação Sicredi Laurentino",
      "ADIANTAMENTO CASA DA LAJE 175K",
      "ADIANTAMENTO CASA DA LAJE 200K",
      "ADIANTAMENTO CASA DA LAJE 250K",
      "ADIANTAMENTO 95K PORTARIS",
      "CAIXA Laurentino",
      "Conta Babi - Não Conciliar",
      "Conta Doca",
      "Conta Equilíbrio",
      "ESTOQUE",
      "PERMUTA INDAPAV BP1 LOTE 13 CASA 1",
      "Lançamentos Passados",
      "PERMUTA DISBRACON LOTE 10 E 14 BV3",
      "PERMUTA AB3 MR VIDROS",
      "PERMUTA AB 4 E 5 MONDINI",
      "PERMUTA DISBRACON ACACIA 26",
      "PERMUTA ACÁCIAS POMERODE LOTE 8 MINI CAR",
      "ADIANTAMENTO 115K VIDRAÇARIA ITOUP",
      "PERMUTA AMS LAJES GOLDEN PARK",
      "PERMUTA AMS LAJES LOTE 18 ACACIA POMEROD",
      "PERMUTA GAIO SERPA",
      "PERMUTA BERTELLI GREEN GARDEN LOTE 36",
      "PERMUTA BERTELLI KW 3 SALA 1",
      "PERMUTA BERTELLI KW 3 SALA 2",
      "PERMUTA BERTELLI SEVEN POINT SALA 9",
      "PERMUTA BONGO 2020 EBERT",
      "PERMUTA MONDINI LOTE 09 CASA 2 BP1",
      "PERMUTA BERTELLI VILA DAS NAÇÕES LT 15",
      "PERMUTA BERTELLI VILA DAS NAÇÕES LT 14",
      "PERMUTA PORTARIS CA 90 CS 1",
      "PERMUTA CALHAS INDAIAL LT NOVA INDAIAL 4",
      "PERMUTA CIVIC EBERT",
      "PERMUTA CASA MIA SEVEN POINT SALA 6",
      "PERMUTA CASA MIA VILA DAS NAÇÕES LT 18",
      "PERMUTA CONCRETOREI GREEN GARDEN LOTE 45",
      "PERMUTA DAINOX BP1 TERRENO MAT. 21884",
      "PERMUTA DAINOX SALA 7 SEVEN POINT",
      "PERMUTA DAINOX GOLDEN PARK LOT A DEF.",
      "PERMUTA DAMACENO",
      "PERMUTA DAMACENO LAND DO ALMIR",
      "PERMUTA DISBRACON ACÁCIA POMERODE LOTE16",
      "PERMUTA DISBRACON ACÁCIA POMERODE LOTE17",
      "PERMUTA DISBRACON SEVEN POINT SALA 4",
      "PERMUTA DISBRACON SEVEN POINT SALA 5",
      "PERMUTA EBERT ONIX",
      "PERMUTA ECO SPORT DAMACENO",
      "PERMUTA FCF RENEGADE (TROCOU DO PA 28)",
      "PERMUTA FIESTA (PEGEOUT) EBERT",
      "PERMUTA CASA MIA GARTEN DORF LOTE 1 E 19",
      "PERMUTA GG 21 DAINOX",
      "PERMUTA CASA DA LAJE GREEN GARDEN LT 30",
      "PERMUTA PORTARIS GG 31",
      "PERMUTA GREEN GARDEN GRAMEIRA FELIPPI 27",
      "PERMUTA GREEN GARDEN MULDE JANDT",
      "Permuta Gol Kosmos",
      "PERMUTA GRAMEIRA FELIPPI GOLDEN PARK",
      "PERMUTA GRAMEIRA FELIPPI GOLDEN PARK 2",
      "PERMUTA GRAMEIRA F. TOLEDO (TROCOU RY12)",
      "PERMUTA INDAPAV SALA 13 SEVEN POINT",
      "PERMUTA INDUMADER GREEN GARDEN 38",
      "PERMUTA INDUMADER GOLDEN PARK 2 LOTES",
      "PERMUTA INDUMADER GREEN VILLAGE 2LT 540M",
      "PERMUTA INDUMADER NOVA A DEFINIR",
      "PERMUTA INDUMADER GOLDEN PARK 3 LOTES",
      "PERMUTA INDUMADER ROYAL PARK 4/5/6",
      "PERMUTA INDUMADER ROYAL PARK 7/8/9",
      "PERMUTA INDAPAV WARNOW 12X80",
      "PERMUTA DAINOX GG LOTE 19",
      "PERMUTA INDAPAV VOLVO GARTEN DORF",
      "PERMUTA VIDRAÇARIA ITOUPAVAZINHA GP1 CS1",
      "PERMUTA JOURNEY EBERT",
      "PERMUTA KOSMOS GREEN GARDEN LOTE 23",
      "PERMUTAS KOSMOS SEVEN POINT SALA 10",
      "CASA DA LAJE ADIANTAMENTO 357K",
      "PERMUTA CASA DA LAJE KW 4 SALA 02",
      "PERMUTA CASA DA LAJE ACACIA POMEROD LT20",
      "PERMUTA CASA DA LAJE ROYAL PARK LT 13/14",
      "PERMUTA CASA DA LAJE SP SALA 8",
      "PERMUTA LF CALHAS GOLDEN PARK A DEF.",
      "PERMUTA LF CALHAS GREEN VILLAGE A DEF.",
      "PERMUTA LF CALHAS SEVEN POINT SALA 1",
      "PERMUTA LF CALHAS RUA TOLEDO",
      "PERMUTA CASA DA LAJE LOTE 03 BV3",
      "PERMUTA CASA DA LAJE LOTE 13 BV",
      "PERMUTA INDAPAV PC 35",
      "PERMUTA CASA DA LAJE LOTE 49 CA",
      "PERMUTA CALHAS INDAIAL LOTE 72 AA",
      "PERMUTA BERTELLI LOTE 10 BP1",
      "PERMUTA VIDRAÇARIA ITOUPAVAZINHA LT11BP1",
      "PERMUTA VIDR. ITOUPAZAVINHA 55 CA",
      "PERMUTA VIDR. ITOUPAVAZINHA 7/8 B. VISTA",
      "PERMUTA LOT. ROYAL PARK SCHROEDER E SCHM",
      "PERMUTA LIVINA EBERT",
      "PERMUTA MINI CARREGADEIRA VILA DAS NAÇÕE",
      "PERMUTA CASA MIA AA 62 CS2",
      "PERMUTA MICHELSON GOLDEN PARK",
      "PERMUTA MOTO ELÉTRICA GANHA SORTEIO DISB",
      "PERMUTA MR VIDROS GOLDEN PARK",
      "PERMUTA PARQUE DAS AREIAS LT 34 CASA MIA",
      "PERMUTA PARQUE DAS AREIAS LOTE 8 MÁRCIO",
      "PERMUTA PB SALA 5 DISBRACON",
      "PERMUTA PB SALA 6 WJ",
      "PERMUTA CONCRETOREI PB SALA A DEFINIR",
      "PERMUTA CASA MIA PB SALA 4",
      "PERMUTA PORTARIS KW4 SALA 1",
      "PERMUTA PORTARIS SEVEN POINT SALA 11",
      "PERMUTA RANGER OPA WJ",
      "PERMUTA GREEN GARDEN ROKA LOTE 29",
      "PERMUTA ROKA PARQUE DAS AREIAS LOTE 25",
      "PERMUTA S10 CHEVROLET LOT. ROYAL PARK",
      "PERMUTA RUI ELÉTRICA SEVEN POINT SALA 12",
      "PERMUTA RUI ELETRICISTA GG 33, 34 E 35",
      "PERMUTA TOMIO SEVEN POINT SALA 3",
      "PERMUTA TOLEDO 2 CONCRETO DJF CONCRETOS",
      "PERMUTA TOYOTA EBERT",
      "PERMUTA WJ RB CS3",
      "PERMUTA WJ SEVEN POINT SALA 2",
      "PERMUTA WARNOW PARK 1 INDAPAV",
      "PERMUTA WARNOW PARK 2 BERTELLI",
      "SALDO DEVOLULÇAO G LIGHT",
      "TROCA DO KW 5 CASA 3 PELO TERRENO SARTOR",
      "PERMUTA TERRENO CARIJÓS NO BP1 08",
      "PERMUTA TERRENO HENRIQUE KUNEN EST. AREA",
      "PERMUTA TOMIO GREEN GARDEN LOTE 46"
    ],
    planoFinanceiro: [
    "Receita de Incorporação de Imóveis",
    "Receita de Bens/Imóveis Adquiridos de Terceiros",
    "Descontos Concedidos (Unidades Imobiliárias)",
    "xxxxxReceita de Aluguel (Locação)",
    "Permuta Externa",
    "Receita de Cliente Desconhecido",
    "Receita de Administração (Taxa Administrativa)",
    "Receita Empreitada Só Mão de Obra",
    "Receita de Projeto",
    "Receita de Locação (Aluguel)",
    "Entrada de Financiamento para Construção",
    "Entrada de Capital de Giro",
    "Receita de Aplicações Financeiras",
    "Receita de Empréstimos",
    "Aporte Capital de Sócio",
    "Descontos Obtidos",
    "Receita de Venda de Imobilizado",
    "Receita Aluguel PCS 35",
    "Estoque de Material",
    "Receita Energia Solar Locatários",
    "Venda de Passivos",
    "(-) Anulação de Custos",
    "(-) Anulação de Custos Mão de Obra",
    "(-) Retenção Caução/Sinal",
    "(-) Retenção de INSS",
    "(-) Retenção de ISS",
    "(-) Retenção de PIS",
    "(-) Retenção de CSLL",
    "(-) Retenção de COFINS",
    "(-) Retenção de IR",
    "(-) Retenção (PIS/COFINS/CSLL - 4,65%)",
    "(-) Retenção de Permuta",
    "(-) Transf. Materiais Estoque Laurentino P/ Finan",
    "(-) Transf. Materiais Estoque Laurentino P/ Bom J.",
    "(-) Transf. Mat. Estoque P/ Obras da Laurentino",
    "(-) Reembolso de Despesas Gerais (Obra)",
    "(-) Divisão de muro",
    "(-) Retenção Contribuição Sindical (Escritório)",
    "(-) Reembolso de Despesas Administrativas",
    "(-) Restituição IPTU Imóveis Alugados (Escritório)",
    "(-) Reembolso de Despesas Financeiras",
    "(-) Reembolso de Despesas Tributárias",
    "(-) Abatimento de Adiantamento",
    "(-) Restituição Condomínio",
    "(-) Restituição água"
],
documentos: ['CONTRATO'],
    condicoesPagamento: [
      'Novo Financiamento', 
      'Financiamento'
    ],
    flags: {
      imprimirParcelasReparceladas: false,
    },
    
    ordem: 'VL',
    
    processarLancamentos: 'CR',
  },
  {
    sheetName: 'Venda Lote',
    pdfName: 'Venda Lote',
    planoFinanceiro: ['Receita de Estoque de Terrenos'],
    
    documentos: [
      "Adiantamento",
      "APORTE DE CAPITAL",
      "AVISO DE LANÇAMENTO",
      "CAUÇÃO",
      "CONHECIMENTO DE FRETE",
      "CONTRATO",
      "CUPOM FISCAL",
      "Documento de Arrecadação de Receita Est",
      "Documento de arrecadação federal",
      "Documento de Arrecadação Municipal",
      "Doc. de Arrecadação do Simples Nacional",
      "Documento de arrecadação simples naciona",
      "Distrato",
      "FATURA",
      "FOLHA DE PAGAMENTO",
      "FATURAMENTO DIRETO",
      "Guia do FGTS Digital",
      "Guia de Previdência Social",
      "GUIA RECOLHIMENTO DE COMISSÃO SINDICAL",
      "Guia de recolhimento de FGTS",
      "GUIA PARA RECOLHIMENTO DE IMPOSTOS",
      "INVISÍVEL",
      "INICIALIZAÇÃO DE ESTOQUE",
      "DOC. CONTAB. SISTEMA CUSTO ORÇ. E INCOR",
      "Movimento de Frotas e Equipamento",
      "Movimento de Locação de Equipamentos",
      "Movimento de transf. de equipamento",
      "MOVIMENTO SEGURANÇA SAUDE NO TRABALHO",
      "NOTA FISCAL",
      "NOTA FISCAL ELETRÔNICA",
      "Nota Fiscal de Serviço",
      "PREVISÃO FINANCEIRA DE CONTR. DE MED.",
      "PERMUTA",
      "PREVISÃO FINANCEIRA DE PEDIDOS DE COMPRA",
      "PREVISÃO DE COMISSÃO",
      "PREVISÃO DE DISTRATO",
      "PREVISÃO DE PAGAMENTO/RECEBIMENTO",
      "REALOCAÇÃO DE MATERIAIS",
      "RECIBO",
      "REQUISIÇÃO DE MATERIAIS",
      "SN",
      "DOCUMENTO DE TRANSFERÊNCIA DE MATERIAIS",
      "TRANSF./RESERVA DE ESTOQUE P/SOLICITAÇÃO",
      "TRANSF. DE MATERIAIS POR SOLICITAÇÃO",
      "TAXAS"
    ],
    condicoesPagamento: [
      "Parcelas Mensais",
      "Parcelas Semestrais",
      "Parcelas Bimestrais",
      "Entrega das chaves",
      "Parcela na Escritura",
      "PERMUTA",
      "Parcela Única",
      "Parcela Anual",
      "Venda a Vista",
      "Novo Financiamento",
      "Resíduo",
      "Novo Parcelas Mensais",
      "Novo Parcelas Semestrais",
      "Novo Parcela Anual",
      "Novo Venda a Vista",
      "Novo Parcela na Escritura",
      "Novo Parcelas Bimestrais",
      "Novo Parcela Única",
      "Novo Parcelas Mensais Aluguel",
      "Parcela Locação Aluguel Mensal",
      "Ato",
      "Novo Ato",
      "Provisionado",
      "Financiamento"
    ],
    contasCorrente: [
      "Viacredi Laurentino",
      "CEF Laurentino",
      "Sicredi Doca",
      "Sicredi Laurentino",
      "Aplicação Viacredi Laurentino",
      "ADIANTAMENTO CASA DA LAJE 150 MIL",
      "ADIANTAMENTO 250 MIL PRA CASA DA LAJE",
      "Aplicação CEF Laurentino",
      "Aplicação Doca Sicredi",
      "Aplicação Sicredi Laurentino",
      "ADIANTAMENTO CASA DA LAJE 175K",
      "ADIANTAMENTO CASA DA LAJE 200K",
      "ADIANTAMENTO 95K PORTARIS",
      "CAIXA Laurentino",
      "Conta Babi - Não Conciliar",
      "Conta Doca",
      "Conta Equilíbrio",
      "ESTOQUE",
      "PERMUTA INDAPAV BP1 LOTE 13 CASA 1",
      "Lançamentos Passados",
      "PERMUTA DISBRACON LOTE 10 E 14 BV3",
      "PERMUTA AB3 MR VIDROS",
      "PERMUTA AB 4 E 5 MONDINI",
      "PERMUTA DISBRACON ACACIA 26",
      "PERMUTA ACÁCIAS POMERODE LOTE 8 MINI CAR",
      "ADIANTAMENTO 115K VIDRAÇARIA ITOUP",
      "PERMUTA AMS LAJES GOLDEN PARK",
      "PERMUTA AMS LAJES LOTE 18 ACACIA POMEROD",
      "PERMUTA GAIO SERPA",
      "PERMUTA BERTELLI GREEN GARDEN LOTE 36",
      "PERMUTA BERTELLI KW 3 SALA 1",
      "PERMUTA BERTELLI KW 3 SALA 2",
      "PERMUTA BERTELLI SEVEN POINT SALA 9",
      "PERMUTA BONGO 2020 EBERT",
      "PERMUTA MONDINI LOTE 09 CASA 2 BP1",
      "PERMUTA BERTELLI VILA DAS NAÇÕES LT 15",
      "PERMUTA BERTELLI VILA DAS NAÇÕES LT 14",
      "PERMUTA PORTARIS CA 90 CS 1",
      "PERMUTA CALHAS INDAIAL LT NOVA INDAIAL 4",
      "PERMUTA CIVIC EBERT",
      "PERMUTA CASA MIA SEVEN POINT SALA 6",
      "PERMUTA CASA MIA VILA DAS NAÇÕES LT 18",
      "PERMUTA CONCRETOREI GREEN GARDEN LOTE 45",
      "PERMUTA DAINOX BP1 TERRENO MAT. 21884",
      "PERMUTA DAINOX SALA 7 SEVEN POINT",
      "PERMUTA DAINOX GOLDEN PARK LOT A DEF.",
      "PERMUTA DAMACENO",
      "PERMUTA DAMACENO LAND DO ALMIR",
      "PERMUTA DISBRACON ACÁCIA POMERODE LOTE16",
      "PERMUTA DISBRACON ACÁCIA POMERODE LOTE17",
      "PERMUTA DISBRACON SEVEN POINT SALA 4",
      "PERMUTA DISBRACON SEVEN POINT SALA 5",
      "PERMUTA EBERT ONIX",
      "PERMUTA ECO SPORT DAMACENO",
      "PERMUTA FCF RENEGADE (TROCOU DO PA 28)",
      "PERMUTA FIESTA (PEGEOUT) EBERT",
      "PERMUTA CASA MIA GARTEN DORF LOTE 1 E 19",
      "PERMUTA GG 21 DAINOX",
      "PERMUTA CASA DA LAJE GREEN GARDEN LT 30",
      "PERMUTA PORTARIS GG 31",
      "PERMUTA GREEN GARDEN GRAMEIRA FELIPPI 27",
      "PERMUTA GREEN GARDEN MULDE JANDT",
      "Permuta Gol Kosmos",
      "PERMUTA GRAMEIRA FELIPPI GOLDEN PARK",
      "PERMUTA GRAMEIRA FELIPPI GOLDEN PARK 2",
      "PERMUTA GRAMEIRA F. TOLEDO (TROCOU RY12)",
      "PERMUTA INDAPAV SALA 13 SEVEN POINT",
      "PERMUTA INDUMADER GREEN GARDEN 38",
      "PERMUTA INDUMADER GOLDEN PARK 2 LOTES",
      "PERMUTA INDUMADER GREEN VILLAGE 2LT 540M",
      "PERMUTA INDUMADER NOVA A DEFINIR",
      "PERMUTA INDUMADER GOLDEN PARK 3 LOTES",
      "PERMUTA INDUMADER ROYAL PARK 4/5/6",
      "PERMUTA INDUMADER ROYAL PARK 7/8/9",
      "PERMUTA INDAPAV WARNOW 12X80",
      "PERMUTA DAINOX GG LOTE 19",
      "PERMUTA INDAPAV VOLVO GARTEN DORF",
      "PERMUTA VIDRAÇARIA ITOUPAVAZINHA GP1 CS1",
      "PERMUTA JOURNEY EBERT",
      "PERMUTA KOSMOS GREEN GARDEN LOTE 23",
      "PERMUTAS KOSMOS SEVEN POINT SALA 10",
      "CASA DA LAJE ADIANTAMENTO 357K",
      "PERMUTA CASA DA LAJE KW 4 SALA 02",
      "PERMUTA CASA DA LAJE ACACIA POMEROD LT20",
      "PERMUTA CASA DA LAJE ROYAL PARK LT 13/14",
      "PERMUTA CASA DA LAJE SP SALA 8",
      "PERMUTA LF CALHAS GOLDEN PARK A DEF.",
      "PERMUTA LF CALHAS GREEN VILLAGE A DEF.",
      "PERMUTA LF CALHAS SEVEN POINT SALA 1",
      "PERMUTA LF CALHAS RUA TOLEDO",
      "PERMUTA CASA DA LAJE LOTE 03 BV3",
      "PERMUTA CASA DA LAJE LOTE 13 BV",
      "PERMUTA INDAPAV PC 35",
      "PERMUTA CASA DA LAJE LOTE 49 CA",
      "PERMUTA CALHAS INDAIAL LOTE 72 AA",
      "PERMUTA BERTELLI LOTE 10 BP1",
      "PERMUTA VIDRAÇARIA ITOUPAVAZINHA LT11BP1",
      "PERMUTA VIDR. ITOUPAZAVINHA 55 CA",
      "PERMUTA VIDR. ITOUPAVAZINHA 7/8 B. VISTA",
      "PERMUTA LOT. ROYAL PARK SCHROEDER E SCHM",
      "PERMUTA LIVINA EBERT",
      "PERMUTA MINI CARREGADEIRA VILA DAS NAÇÕE",
      "PERMUTA CASA MIA AA 62 CS2",
      "PERMUTA MICHELSON GOLDEN PARK",
      "PERMUTA MOTO ELÉTRICA GANHA SORTEIO DISB",
      "PERMUTA MR VIDROS GOLDEN PARK",
      "PERMUTA PARQUE DAS AREIAS LT 34 CASA MIA",
      "PERMUTA PARQUE DAS AREIAS LOTE 8 MÁRCIO",
      "PERMUTA PB SALA 5 DISBRACON",
      "PERMUTA PB SALA 6 WJ",
      "PERMUTA CONCRETOREI PB SALA A DEFINIR",
      "PERMUTA CASA MIA PB SALA 4",
      "PERMUTA PORTARIS KW4 SALA 1",
      "PERMUTA PORTARIS SEVEN POINT SALA 11",
      "PERMUTA RANGER OPA WJ",
      "PERMUTA GREEN GARDEN ROKA LOTE 29",
      "PERMUTA ROKA PARQUE DAS AREIAS LOTE 25",
      "PERMUTA S10 CHEVROLET LOT. ROYAL PARK",
      "PERMUTA RUI ELÉTRICA SEVEN POINT SALA 12",
      "PERMUTA RUI ELETRICISTA GG 33, 34 E 35",
      "PERMUTA TOMIO SEVEN POINT SALA 3",
      "PERMUTA TOLEDO 2 CONCRETO DJF CONCRETOS",
      "PERMUTA TOYOTA EBERT",
      "PERMUTA WJ RB CS3",
      "PERMUTA WJ SEVEN POINT SALA 2",
      "PERMUTA WARNOW PARK 1 INDAPAV",
      "PERMUTA WARNOW PARK 2 BERTELLI",
      "SALDO DEVOLULÇAO G LIGHT",
      "TROCA DO KW 5 CASA 3 PELO TERRENO SARTOR",
      "PERMUTA TERRENO CARIJÓS NO BP1 08",
      "PERMUTA TERRENO HENRIQUE KUNEN EST. AREA",
      "PERMUTA TOMIO GREEN GARDEN LOTE 46"
    ],
    flags: {
      imprimirParcelasReparceladas: false,
    },
    
    ordem: 'VL',
    
    processarLancamentos: 'CR',
  },
  {
    sheetName: 'Empréstimo',
    pdfName: 'Empréstimo',
    planoFinanceiro: ['Receita de Empréstimos'],
    
    contasCorrente: [
      "Viacredi Laurentino",
      "CEF Laurentino",
      "Sicredi Doca",
      "Sicredi Laurentino",
      "Aplicação Viacredi Laurentino",
      "ADIANTAMENTO CASA DA LAJE 150 MIL",
      "ADIANTAMENTO 250 MIL PRA CASA DA LAJE",
      "Aplicação CEF Laurentino",
      "Aplicação Doca Sicredi",
      "Aplicação Sicredi Laurentino",
      "ADIANTAMENTO CASA DA LAJE 175K",
      "ADIANTAMENTO CASA DA LAJE 200K",
      "ADIANTAMENTO CASA DA LAJE 250K",
      "ADIANTAMENTO 95K PORTARIS",
      "CAIXA Laurentino",
      "Conta Babi - Não Conciliar",
      "Conta Doca",
      "Conta Equilíbrio",
      "ESTOQUE",
      "PERMUTA INDAPAV BP1 LOTE 13 CASA 1",
      "Lançamentos Passados",
      "PERMUTA DISBRACON LOTE 10 E 14 BV3",
      "PERMUTA AB3 MR VIDROS",
      "PERMUTA AB 4 E 5 MONDINI",
      "PERMUTA DISBRACON ACACIA 26",
      "PERMUTA ACÁCIAS POMERODE LOTE 8 MINI CAR",
      "ADIANTAMENTO 115K VIDRAÇARIA ITOUP",
      "PERMUTA AMS LAJES GOLDEN PARK",
      "PERMUTA AMS LAJES LOTE 18 ACACIA POMEROD",
      "PERMUTA GAIO SERPA",
      "PERMUTA BERTELLI GREEN GARDEN LOTE 36",
      "PERMUTA BERTELLI KW 3 SALA 1",
      "PERMUTA BERTELLI KW 3 SALA 2",
      "PERMUTA BERTELLI SEVEN POINT SALA 9",
      "PERMUTA BONGO 2020 EBERT",
      "PERMUTA MONDINI LOTE 09 CASA 2 BP1",
      "PERMUTA BERTELLI VILA DAS NAÇÕES LT 15",
      "PERMUTA BERTELLI VILA DAS NAÇÕES LT 14",
      "PERMUTA PORTARIS CA 90 CS 1",
      "PERMUTA CALHAS INDAIAL LT NOVA INDAIAL 4",
      "PERMUTA CIVIC EBERT",
      "PERMUTA CASA MIA SEVEN POINT SALA 6",
      "PERMUTA CASA MIA VILA DAS NAÇÕES LT 18",
      "PERMUTA CONCRETOREI GREEN GARDEN LOTE 45",
      "PERMUTA DAINOX BP1 TERRENO MAT. 21884",
      "PERMUTA DAINOX SALA 7 SEVEN POINT",
      "PERMUTA DAINOX GOLDEN PARK LOT A DEF.",
      "PERMUTA DAMACENO",
      "PERMUTA DAMACENO LAND DO ALMIR",
      "PERMUTA DISBRACON ACÁCIA POMERODE LOTE16",
      "PERMUTA DISBRACON ACÁCIA POMERODE LOTE17",
      "PERMUTA DISBRACON SEVEN POINT SALA 4",
      "PERMUTA DISBRACON SEVEN POINT SALA 5",
      "PERMUTA EBERT ONIX",
      "PERMUTA ECO SPORT DAMACENO",
      "PERMUTA FCF RENEGADE (TROCOU DO PA 28)",
      "PERMUTA FIESTA (PEGEOUT) EBERT",
      "PERMUTA CASA MIA GARTEN DORF LOTE 1 E 19",
      "PERMUTA GG 21 DAINOX",
      "PERMUTA CASA DA LAJE GREEN GARDEN LT 30",
      "PERMUTA PORTARIS GG 31",
      "PERMUTA GREEN GARDEN GRAMEIRA FELIPPI 27",
      "PERMUTA GREEN GARDEN MULDE JANDT",
      "Permuta Gol Kosmos",
      "PERMUTA GRAMEIRA FELIPPI GOLDEN PARK",
      "PERMUTA GRAMEIRA FELIPPI GOLDEN PARK 2",
      "PERMUTA GRAMEIRA F. TOLEDO (TROCOU RY12)",
      "PERMUTA INDAPAV SALA 13 SEVEN POINT",
      "PERMUTA INDUMADER GREEN GARDEN 38",
      "PERMUTA INDUMADER GOLDEN PARK 2 LOTES",
      "PERMUTA INDUMADER GREEN VILLAGE 2LT 540M",
      "PERMUTA INDUMADER NOVA A DEFINIR",
      "PERMUTA INDUMADER GOLDEN PARK 3 LOTES",
      "PERMUTA INDUMADER ROYAL PARK 4/5/6",
      "PERMUTA INDUMADER ROYAL PARK 7/8/9",
      "PERMUTA INDAPAV WARNOW 12X80",
      "PERMUTA DAINOX GG LOTE 19",
      "PERMUTA INDAPAV VOLVO GARTEN DORF",
      "PERMUTA VIDRAÇARIA ITOUPAVAZINHA GP1 CS1",
      "PERMUTA JOURNEY EBERT",
      "PERMUTA KOSMOS GREEN GARDEN LOTE 23",
      "PERMUTAS KOSMOS SEVEN POINT SALA 10",
      "CASA DA LAJE ADIANTAMENTO 357K",
      "PERMUTA CASA DA LAJE KW 4 SALA 02",
      "PERMUTA CASA DA LAJE ACACIA POMEROD LT20",
      "PERMUTA CASA DA LAJE ROYAL PARK LT 13/14",
      "PERMUTA CASA DA LAJE SP SALA 8",
      "PERMUTA LF CALHAS GOLDEN PARK A DEF.",
      "PERMUTA LF CALHAS GREEN VILLAGE A DEF.",
      "PERMUTA LF CALHAS SEVEN POINT SALA 1",
      "PERMUTA LF CALHAS RUA TOLEDO",
      "PERMUTA CASA DA LAJE LOTE 03 BV3",
      "PERMUTA CASA DA LAJE LOTE 13 BV",
      "PERMUTA INDAPAV PC 35",
      "PERMUTA CASA DA LAJE LOTE 49 CA",
      "PERMUTA CALHAS INDAIAL LOTE 72 AA",
      "PERMUTA BERTELLI LOTE 10 BP1",
      "PERMUTA VIDRAÇARIA ITOUPAVAZINHA LT11BP1",
      "PERMUTA VIDR. ITOUPAZAVINHA 55 CA",
      "PERMUTA VIDR. ITOUPAVAZINHA 7/8 B. VISTA",
      "PERMUTA LOT. ROYAL PARK SCHROEDER E SCHM",
      "PERMUTA LIVINA EBERT",
      "PERMUTA MINI CARREGADEIRA VILA DAS NAÇÕE",
      "PERMUTA CASA MIA AA 62 CS2",
      "PERMUTA MICHELSON GOLDEN PARK",
      "PERMUTA MOTO ELÉTRICA GANHA SORTEIO DISB",
      "PERMUTA MR VIDROS GOLDEN PARK",
      "PERMUTA PARQUE DAS AREIAS LT 34 CASA MIA",
      "PERMUTA PARQUE DAS AREIAS LOTE 8 MÁRCIO",
      "PERMUTA PB SALA 5 DISBRACON",
      "PERMUTA PB SALA 6 WJ",
      "PERMUTA CONCRETOREI PB SALA A DEFINIR",
      "PERMUTA CASA MIA PB SALA 4",
      "PERMUTA PORTARIS KW4 SALA 1",
      "PERMUTA PORTARIS SEVEN POINT SALA 11",
      "PERMUTA RANGER OPA WJ",
      "PERMUTA GREEN GARDEN ROKA LOTE 29",
      "PERMUTA ROKA PARQUE DAS AREIAS LOTE 25",
      "PERMUTA S10 CHEVROLET LOT. ROYAL PARK",
      "PERMUTA RUI ELÉTRICA SEVEN POINT SALA 12",
      "PERMUTA RUI ELETRICISTA GG 33, 34 E 35",
      "PERMUTA TOMIO SEVEN POINT SALA 3",
      "PERMUTA TOLEDO 2 CONCRETO DJF CONCRETOS",
      "PERMUTA TOYOTA EBERT",
      "PERMUTA WJ RB CS3",
      "PERMUTA WJ SEVEN POINT SALA 2",
      "PERMUTA WARNOW PARK 1 INDAPAV",
      "PERMUTA WARNOW PARK 2 BERTELLI",
      "SALDO DEVOLULÇAO G LIGHT",
      "TROCA DO KW 5 CASA 3 PELO TERRENO SARTOR",
      "PERMUTA TERRENO CARIJÓS NO BP1 08",
      "PERMUTA TERRENO HENRIQUE KUNEN EST. AREA",
      "PERMUTA TOMIO GREEN GARDEN LOTE 46"
    ],
    documentos: [
      "Adiantamento",
      "APORTE DE CAPITAL",
      "AVISO DE LANÇAMENTO",
      "CAUÇÃO",
      "CONHECIMENTO DE FRETE",
      "CONTRATO",
      "CUPOM FISCAL",
      "Documento de Arrecadação de Receita Est",
      "Documento de arrecadação federal",
      "Documento de Arrecadação Municipal",
      "Doc. de Arrecadação do Simples Nacional",
      "Documento de arrecadação simples naciona",
      "Distrato",
      "FATURA",
      "FOLHA DE PAGAMENTO",
      "FATURAMENTO DIRETO",
      "Guia do FGTS Digital",
      "Guia de Previdência Social",
      "GUIA RECOLHIMENTO DE COMISSÃO SINDICAL",
      "Guia de recolhimento de FGTS",
      "GUIA PARA RECOLHIMENTO DE IMPOSTOS",
      "INVISÍVEL",
      "INICIALIZAÇÃO DE ESTOQUE",
      "DOC. CONTAB. SISTEMA CUSTO ORÇ. E INCOR",
      "Movimento de Frotas e Equipamento",
      "Movimento de Locação de Equipamentos",
      "Movimento de transf. de equipamento",
      "MOVIMENTO SEGURANÇA SAUDE NO TRABALHO",
      "NOTA FISCAL",
      "NOTA FISCAL ELETRÔNICA",
      "Nota Fiscal de Serviço",
      "PREVISÃO FINANCEIRA DE CONTR. DE MED.",
      "PERMUTA",
      "PREVISÃO FINANCEIRA DE PEDIDOS DE COMPRA",
      "PREVISÃO DE COMISSÃO",
      "PREVISÃO DE DISTRATO",
      "PREVISÃO DE PAGAMENTO/RECEBIMENTO",
      "REALOCAÇÃO DE MATERIAIS",
      "RECIBO",
      "REQUISIÇÃO DE MATERIAIS",
      "SN",
      "DOCUMENTO DE TRANSFERÊNCIA DE MATERIAIS",
      "TRANSF./RESERVA DE ESTOQUE P/SOLICITAÇÃO",
      "TRANSF. DE MATERIAIS POR SOLICITAÇÃO",
      "TAXAS"
  ],
    condicoesPagamento: [
      "Parcelas Mensais",
      "Parcelas Semestrais",
      "Parcelas Bimestrais",
      "Entrega das chaves",
      "Parcela na Escritura",
      "PERMUTA",
      "Parcela Única",
      "Parcela Anual",
      "Venda a Vista",
      "Novo Financiamento",
      "Resíduo",
      "Novo Parcelas Mensais",
      "Novo Parcelas Semestrais",
      "Novo Parcela Anual",
      "Novo Venda a Vista",
      "Novo Parcela na Escritura",
      "Novo Parcelas Bimestrais",
      "Novo Parcela Única",
      "Novo Parcelas Mensais Aluguel",
      "Parcela Locação Aluguel Mensal",
      "Ato",
      "Novo Ato",
      "Provisionado",
      "Financiamento"
    ],
    flags: {
      imprimirParcelasReparceladas: false,
    },
    
    ordem: 'VL',
    
    processarLancamentos: 'CR',
  },
  {
    sheetName: 'Venda de Passivo',
    contasCorrente: [
      "Viacredi Laurentino",
      "CEF Laurentino",
      "Sicredi Doca",
      "Sicredi Laurentino",
      "Aplicação Viacredi Laurentino",
      "ADIANTAMENTO CASA DA LAJE 150 MIL",
      "ADIANTAMENTO 250 MIL PRA CASA DA LAJE",
      "Aplicação CEF Laurentino",
      "Aplicação Doca Sicredi",
      "Aplicação Sicredi Laurentino",
      "ADIANTAMENTO CASA DA LAJE 175K",
      "ADIANTAMENTO CASA DA LAJE 200K",
      "ADIANTAMENTO CASA DA LAJE 250K",
      "ADIANTAMENTO 95K PORTARIS",
      "CAIXA Laurentino",
      "Conta Babi - Não Conciliar",
      "Conta Doca",
      "Conta Equilíbrio",
      "ESTOQUE",
      "PERMUTA INDAPAV BP1 LOTE 13 CASA 1",
      "Lançamentos Passados",
      "PERMUTA DISBRACON LOTE 10 E 14 BV3",
      "PERMUTA AB3 MR VIDROS",
      "PERMUTA AB 4 E 5 MONDINI",
      "PERMUTA DISBRACON ACACIA 26",
      "PERMUTA ACÁCIAS POMERODE LOTE 8 MINI CAR",
      "ADIANTAMENTO 115K VIDRAÇARIA ITOUP",
      "PERMUTA AMS LAJES GOLDEN PARK",
      "PERMUTA AMS LAJES LOTE 18 ACACIA POMEROD",
      "PERMUTA GAIO SERPA",
      "PERMUTA BERTELLI GREEN GARDEN LOTE 36",
      "PERMUTA BERTELLI KW 3 SALA 1",
      "PERMUTA BERTELLI KW 3 SALA 2",
      "PERMUTA BERTELLI SEVEN POINT SALA 9",
      "PERMUTA BONGO 2020 EBERT",
      "PERMUTA MONDINI LOTE 09 CASA 2 BP1",
      "PERMUTA BERTELLI VILA DAS NAÇÕES LT 15",
      "PERMUTA BERTELLI VILA DAS NAÇÕES LT 14",
      "PERMUTA PORTARIS CA 90 CS 1",
      "PERMUTA CALHAS INDAIAL LT NOVA INDAIAL 4",
      "PERMUTA CIVIC EBERT",
      "PERMUTA CASA MIA SEVEN POINT SALA 6",
      "PERMUTA CASA MIA VILA DAS NAÇÕES LT 18",
      "PERMUTA CONCRETOREI GREEN GARDEN LOTE 45",
      "PERMUTA DAINOX BP1 TERRENO MAT. 21884",
      "PERMUTA DAINOX SALA 7 SEVEN POINT",
      "PERMUTA DAINOX GOLDEN PARK LOT A DEF.",
      "PERMUTA DAMACENO",
      "PERMUTA DAMACENO LAND DO ALMIR",
      "PERMUTA DISBRACON ACÁCIA POMERODE LOTE16",
      "PERMUTA DISBRACON ACÁCIA POMERODE LOTE17",
      "PERMUTA DISBRACON SEVEN POINT SALA 4",
      "PERMUTA DISBRACON SEVEN POINT SALA 5",
      "PERMUTA EBERT ONIX",
      "PERMUTA ECO SPORT DAMACENO",
      "PERMUTA FCF RENEGADE (TROCOU DO PA 28)",
      "PERMUTA FIESTA (PEGEOUT) EBERT",
      "PERMUTA CASA MIA GARTEN DORF LOTE 1 E 19",
      "PERMUTA GG 21 DAINOX",
      "PERMUTA CASA DA LAJE GREEN GARDEN LT 30",
      "PERMUTA PORTARIS GG 31",
      "PERMUTA GREEN GARDEN GRAMEIRA FELIPPI 27",
      "PERMUTA GREEN GARDEN MULDE JANDT",
      "Permuta Gol Kosmos",
      "PERMUTA GRAMEIRA FELIPPI GOLDEN PARK",
      "PERMUTA GRAMEIRA FELIPPI GOLDEN PARK 2",
      "PERMUTA GRAMEIRA F. TOLEDO (TROCOU RY12)",
      "PERMUTA INDAPAV SALA 13 SEVEN POINT",
      "PERMUTA INDUMADER GREEN GARDEN 38",
      "PERMUTA INDUMADER GOLDEN PARK 2 LOTES",
      "PERMUTA INDUMADER GREEN VILLAGE 2LT 540M",
      "PERMUTA INDUMADER NOVA A DEFINIR",
      "PERMUTA INDUMADER GOLDEN PARK 3 LOTES",
      "PERMUTA INDUMADER ROYAL PARK 4/5/6",
      "PERMUTA INDUMADER ROYAL PARK 7/8/9",
      "PERMUTA INDAPAV WARNOW 12X80",
      "PERMUTA DAINOX GG LOTE 19",
      "PERMUTA INDAPAV VOLVO GARTEN DORF",
      "PERMUTA VIDRAÇARIA ITOUPAVAZINHA GP1 CS1",
      "PERMUTA JOURNEY EBERT",
      "PERMUTA KOSMOS GREEN GARDEN LOTE 23",
      "PERMUTAS KOSMOS SEVEN POINT SALA 10",
      "CASA DA LAJE ADIANTAMENTO 357K",
      "PERMUTA CASA DA LAJE KW 4 SALA 02",
      "PERMUTA CASA DA LAJE ACACIA POMEROD LT20",
      "PERMUTA CASA DA LAJE ROYAL PARK LT 13/14",
      "PERMUTA CASA DA LAJE SP SALA 8",
      "PERMUTA LF CALHAS GOLDEN PARK A DEF.",
      "PERMUTA LF CALHAS GREEN VILLAGE A DEF.",
      "PERMUTA LF CALHAS SEVEN POINT SALA 1",
      "PERMUTA LF CALHAS RUA TOLEDO",
      "PERMUTA CASA DA LAJE LOTE 03 BV3",
      "PERMUTA CASA DA LAJE LOTE 13 BV",
      "PERMUTA INDAPAV PC 35",
      "PERMUTA CASA DA LAJE LOTE 49 CA",
      "PERMUTA CALHAS INDAIAL LOTE 72 AA",
      "PERMUTA BERTELLI LOTE 10 BP1",
      "PERMUTA VIDRAÇARIA ITOUPAVAZINHA LT11BP1",
      "PERMUTA VIDR. ITOUPAZAVINHA 55 CA",
      "PERMUTA VIDR. ITOUPAVAZINHA 7/8 B. VISTA",
      "PERMUTA LOT. ROYAL PARK SCHROEDER E SCHM",
      "PERMUTA LIVINA EBERT",
      "PERMUTA MINI CARREGADEIRA VILA DAS NAÇÕE",
      "PERMUTA CASA MIA AA 62 CS2",
      "PERMUTA MICHELSON GOLDEN PARK",
      "PERMUTA MOTO ELÉTRICA GANHA SORTEIO DISB",
      "PERMUTA MR VIDROS GOLDEN PARK",
      "PERMUTA PARQUE DAS AREIAS LT 34 CASA MIA",
      "PERMUTA PARQUE DAS AREIAS LOTE 8 MÁRCIO",
      "PERMUTA PB SALA 5 DISBRACON",
      "PERMUTA PB SALA 6 WJ",
      "PERMUTA CONCRETOREI PB SALA A DEFINIR",
      "PERMUTA CASA MIA PB SALA 4",
      "PERMUTA PORTARIS KW4 SALA 1",
      "PERMUTA PORTARIS SEVEN POINT SALA 11",
      "PERMUTA RANGER OPA WJ",
      "PERMUTA GREEN GARDEN ROKA LOTE 29",
      "PERMUTA ROKA PARQUE DAS AREIAS LOTE 25",
      "PERMUTA S10 CHEVROLET LOT. ROYAL PARK",
      "PERMUTA RUI ELÉTRICA SEVEN POINT SALA 12",
      "PERMUTA RUI ELETRICISTA GG 33, 34 E 35",
      "PERMUTA TOMIO SEVEN POINT SALA 3",
      "PERMUTA TOLEDO 2 CONCRETO DJF CONCRETOS",
      "PERMUTA TOYOTA EBERT",
      "PERMUTA WJ RB CS3",
      "PERMUTA WJ SEVEN POINT SALA 2",
      "PERMUTA WARNOW PARK 1 INDAPAV",
      "PERMUTA WARNOW PARK 2 BERTELLI",
      "SALDO DEVOLULÇAO G LIGHT",
      "TROCA DO KW 5 CASA 3 PELO TERRENO SARTOR",
      "PERMUTA TERRENO CARIJÓS NO BP1 08",
      "PERMUTA TERRENO HENRIQUE KUNEN EST. AREA",
      "PERMUTA TOMIO GREEN GARDEN LOTE 46"
    ],
    pdfName: 'Venda de Passivo',
    planoFinanceiro: ['Venda de Passivos'],
    
    documentos: [
      "Adiantamento",
      "APORTE DE CAPITAL",
      "AVISO DE LANÇAMENTO",
      "CAUÇÃO",
      "CONHECIMENTO DE FRETE",
      "CONTRATO",
      "CUPOM FISCAL",
      "Documento de Arrecadação de Receita Est",
      "Documento de arrecadação federal",
      "Documento de Arrecadação Municipal",
      "Doc. de Arrecadação do Simples Nacional",
      "Documento de arrecadação simples naciona",
      "Distrato",
      "FATURA",
      "FOLHA DE PAGAMENTO",
      "FATURAMENTO DIRETO",
      "Guia do FGTS Digital",
      "Guia de Previdência Social",
      "GUIA RECOLHIMENTO DE COMISSÃO SINDICAL",
      "Guia de recolhimento de FGTS",
      "GUIA PARA RECOLHIMENTO DE IMPOSTOS",
      "INVISÍVEL",
      "INICIALIZAÇÃO DE ESTOQUE",
      "DOC. CONTAB. SISTEMA CUSTO ORÇ. E INCOR",
      "Movimento de Frotas e Equipamento",
      "Movimento de Locação de Equipamentos",
      "Movimento de transf. de equipamento",
      "MOVIMENTO SEGURANÇA SAUDE NO TRABALHO",
      "NOTA FISCAL",
      "NOTA FISCAL ELETRÔNICA",
      "Nota Fiscal de Serviço",
      "PREVISÃO FINANCEIRA DE CONTR. DE MED.",
      "PERMUTA",
      "PREVISÃO FINANCEIRA DE PEDIDOS DE COMPRA",
      "PREVISÃO DE COMISSÃO",
      "PREVISÃO DE DISTRATO",
      "PREVISÃO DE PAGAMENTO/RECEBIMENTO",
      "REALOCAÇÃO DE MATERIAIS",
      "RECIBO",
      "REQUISIÇÃO DE MATERIAIS",
      "SN",
      "DOCUMENTO DE TRANSFERÊNCIA DE MATERIAIS",
      "TRANSF./RESERVA DE ESTOQUE P/SOLICITAÇÃO",
      "TRANSF. DE MATERIAIS POR SOLICITAÇÃO",
      "TAXAS"
    ],
    condicoesPagamento: [
      "Parcelas Mensais",
      "Parcelas Semestrais",
      "Parcelas Bimestrais",
      "Entrega das chaves",
      "Parcela na Escritura",
      "PERMUTA",
      "Parcela Única",
      "Parcela Anual",
      "Venda a Vista",
      "Novo Financiamento",
      "Resíduo",
      "Novo Parcelas Mensais",
      "Novo Parcelas Semestrais",
      "Novo Parcela Anual",
      "Novo Venda a Vista",
      "Novo Parcela na Escritura",
      "Novo Parcelas Bimestrais",
      "Novo Parcela Única",
      "Novo Parcelas Mensais Aluguel",
      "Parcela Locação Aluguel Mensal",
      "Ato",
      "Novo Ato",
      "Provisionado",
      "Financiamento"
    ],
    flags: {
      imprimirParcelasReparceladas: false,
    },
    
    ordem: 'VL',
    
    processarLancamentos: 'CR',
  },
  {
    sheetName: 'Receitas Diversas',
    pdfName: 'Receitas Diversas',
    contasCorrente: [
      "Viacredi Laurentino",
      "CEF Laurentino",
      "Sicredi Doca",
      "Sicredi Laurentino",
      "Aplicação Viacredi Laurentino",
      "ADIANTAMENTO CASA DA LAJE 150 MIL",
      "ADIANTAMENTO 250 MIL PRA CASA DA LAJE",
      "Aplicação CEF Laurentino",
      "Aplicação Doca Sicredi",
      "Aplicação Sicredi Laurentino",
      "ADIANTAMENTO CASA DA LAJE 175K",
      "ADIANTAMENTO CASA DA LAJE 200K",
      "ADIANTAMENTO CASA DA LAJE 250K",
      "ADIANTAMENTO 95K PORTARIS",
      "CAIXA Laurentino",
      "Conta Babi - Não Conciliar",
      "Conta Doca",
      "Conta Equilíbrio",
      "ESTOQUE",
      "PERMUTA INDAPAV BP1 LOTE 13 CASA 1",
      "Lançamentos Passados",
      "PERMUTA DISBRACON LOTE 10 E 14 BV3",
      "PERMUTA AB3 MR VIDROS",
      "PERMUTA AB 4 E 5 MONDINI",
      "PERMUTA DISBRACON ACACIA 26",
      "PERMUTA ACÁCIAS POMERODE LOTE 8 MINI CAR",
      "ADIANTAMENTO 115K VIDRAÇARIA ITOUP",
      "PERMUTA AMS LAJES GOLDEN PARK",
      "PERMUTA AMS LAJES LOTE 18 ACACIA POMEROD",
      "PERMUTA GAIO SERPA",
      "PERMUTA BERTELLI GREEN GARDEN LOTE 36",
      "PERMUTA BERTELLI KW 3 SALA 1",
      "PERMUTA BERTELLI KW 3 SALA 2",
      "PERMUTA BERTELLI SEVEN POINT SALA 9",
      "PERMUTA BONGO 2020 EBERT",
      "PERMUTA MONDINI LOTE 09 CASA 2 BP1",
      "PERMUTA BERTELLI VILA DAS NAÇÕES LT 15",
      "PERMUTA BERTELLI VILA DAS NAÇÕES LT 14",
      "PERMUTA PORTARIS CA 90 CS 1",
      "PERMUTA CALHAS INDAIAL LT NOVA INDAIAL 4",
      "PERMUTA CIVIC EBERT",
      "PERMUTA CASA MIA SEVEN POINT SALA 6",
      "PERMUTA CASA MIA VILA DAS NAÇÕES LT 18",
      "PERMUTA CONCRETOREI GREEN GARDEN LOTE 45",
      "PERMUTA DAINOX BP1 TERRENO MAT. 21884",
      "PERMUTA DAINOX SALA 7 SEVEN POINT",
      "PERMUTA DAINOX GOLDEN PARK LOT A DEF.",
      "PERMUTA DAMACENO",
      "PERMUTA DAMACENO LAND DO ALMIR",
      "PERMUTA DISBRACON ACÁCIA POMERODE LOTE16",
      "PERMUTA DISBRACON ACÁCIA POMERODE LOTE17",
      "PERMUTA DISBRACON SEVEN POINT SALA 4",
      "PERMUTA DISBRACON SEVEN POINT SALA 5",
      "PERMUTA EBERT ONIX",
      "PERMUTA ECO SPORT DAMACENO",
      "PERMUTA FCF RENEGADE (TROCOU DO PA 28)",
      "PERMUTA FIESTA (PEGEOUT) EBERT",
      "PERMUTA CASA MIA GARTEN DORF LOTE 1 E 19",
      "PERMUTA GG 21 DAINOX",
      "PERMUTA CASA DA LAJE GREEN GARDEN LT 30",
      "PERMUTA PORTARIS GG 31",
      "PERMUTA GREEN GARDEN GRAMEIRA FELIPPI 27",
      "PERMUTA GREEN GARDEN MULDE JANDT",
      "Permuta Gol Kosmos",
      "PERMUTA GRAMEIRA FELIPPI GOLDEN PARK",
      "PERMUTA GRAMEIRA FELIPPI GOLDEN PARK 2",
      "PERMUTA GRAMEIRA F. TOLEDO (TROCOU RY12)",
      "PERMUTA INDAPAV SALA 13 SEVEN POINT",
      "PERMUTA INDUMADER GREEN GARDEN 38",
      "PERMUTA INDUMADER GOLDEN PARK 2 LOTES",
      "PERMUTA INDUMADER GREEN VILLAGE 2LT 540M",
      "PERMUTA INDUMADER NOVA A DEFINIR",
      "PERMUTA INDUMADER GOLDEN PARK 3 LOTES",
      "PERMUTA INDUMADER ROYAL PARK 4/5/6",
      "PERMUTA INDUMADER ROYAL PARK 7/8/9",
      "PERMUTA INDAPAV WARNOW 12X80",
      "PERMUTA DAINOX GG LOTE 19",
      "PERMUTA INDAPAV VOLVO GARTEN DORF",
      "PERMUTA VIDRAÇARIA ITOUPAVAZINHA GP1 CS1",
      "PERMUTA JOURNEY EBERT",
      "PERMUTA KOSMOS GREEN GARDEN LOTE 23",
      "PERMUTAS KOSMOS SEVEN POINT SALA 10",
      "CASA DA LAJE ADIANTAMENTO 357K",
      "PERMUTA CASA DA LAJE KW 4 SALA 02",
      "PERMUTA CASA DA LAJE ACACIA POMEROD LT20",
      "PERMUTA CASA DA LAJE ROYAL PARK LT 13/14",
      "PERMUTA CASA DA LAJE SP SALA 8",
      "PERMUTA LF CALHAS GOLDEN PARK A DEF.",
      "PERMUTA LF CALHAS GREEN VILLAGE A DEF.",
      "PERMUTA LF CALHAS SEVEN POINT SALA 1",
      "PERMUTA LF CALHAS RUA TOLEDO",
      "PERMUTA CASA DA LAJE LOTE 03 BV3",
      "PERMUTA CASA DA LAJE LOTE 13 BV",
      "PERMUTA INDAPAV PC 35",
      "PERMUTA CASA DA LAJE LOTE 49 CA",
      "PERMUTA CALHAS INDAIAL LOTE 72 AA",
      "PERMUTA BERTELLI LOTE 10 BP1",
      "PERMUTA VIDRAÇARIA ITOUPAVAZINHA LT11BP1",
      "PERMUTA VIDR. ITOUPAZAVINHA 55 CA",
      "PERMUTA VIDR. ITOUPAVAZINHA 7/8 B. VISTA",
      "PERMUTA LOT. ROYAL PARK SCHROEDER E SCHM",
      "PERMUTA LIVINA EBERT",
      "PERMUTA MINI CARREGADEIRA VILA DAS NAÇÕE",
      "PERMUTA CASA MIA AA 62 CS2",
      "PERMUTA MICHELSON GOLDEN PARK",
      "PERMUTA MOTO ELÉTRICA GANHA SORTEIO DISB",
      "PERMUTA MR VIDROS GOLDEN PARK",
      "PERMUTA PARQUE DAS AREIAS LT 34 CASA MIA",
      "PERMUTA PARQUE DAS AREIAS LOTE 8 MÁRCIO",
      "PERMUTA PB SALA 5 DISBRACON",
      "PERMUTA PB SALA 6 WJ",
      "PERMUTA CONCRETOREI PB SALA A DEFINIR",
      "PERMUTA CASA MIA PB SALA 4",
      "PERMUTA PORTARIS KW4 SALA 1",
      "PERMUTA PORTARIS SEVEN POINT SALA 11",
      "PERMUTA RANGER OPA WJ",
      "PERMUTA GREEN GARDEN ROKA LOTE 29",
      "PERMUTA ROKA PARQUE DAS AREIAS LOTE 25",
      "PERMUTA S10 CHEVROLET LOT. ROYAL PARK",
      "PERMUTA RUI ELÉTRICA SEVEN POINT SALA 12",
      "PERMUTA RUI ELETRICISTA GG 33, 34 E 35",
      "PERMUTA TOMIO SEVEN POINT SALA 3",
      "PERMUTA TOLEDO 2 CONCRETO DJF CONCRETOS",
      "PERMUTA TOYOTA EBERT",
      "PERMUTA WJ RB CS3",
      "PERMUTA WJ SEVEN POINT SALA 2",
      "PERMUTA WARNOW PARK 1 INDAPAV",
      "PERMUTA WARNOW PARK 2 BERTELLI",
      "SALDO DEVOLULÇAO G LIGHT",
      "TROCA DO KW 5 CASA 3 PELO TERRENO SARTOR",
      "PERMUTA TERRENO CARIJÓS NO BP1 08",
      "PERMUTA TERRENO HENRIQUE KUNEN EST. AREA",
      "PERMUTA TOMIO GREEN GARDEN LOTE 46"
    ],
    planoFinanceiro: [
    "Receita de Incorporação de Imóveis",
    "Receita de Bens/Imóveis Adquiridos de Terceiros",
    "Descontos Concedidos (Unidades Imobiliárias)",
    "xxxxxReceita de Aluguel (Locação)",
    "Permuta Externa",
    "Receita de Cliente Desconhecido",
    "Receita de Administração (Taxa Administrativa)",
    "Receita Empreitada Só Mão de Obra",
    "Receita de Projeto",
    "Receita de Locação (Aluguel)",
    "Entrada de Financiamento para Construção",
    "Entrada de Capital de Giro",
    "Receita de Aplicações Financeiras",
    "Descontos Obtidos",
    "Receita de Venda de Imobilizado",
    "Receita Aluguel PCS 35",
    "Estoque de Material",
    "Receita Energia Solar Locatários",
    "(-) Anulação de Custos",
    "(-) Anulação de Custos Mão de Obra",
    "(-) Retenção Caução/Sinal",
    "(-) Retenção de INSS",
    "(-) Retenção de ISS",
    "(-) Retenção de PIS",
    "(-) Retenção de CSLL",
    "(-) Retenção de COFINS",
    "(-) Retenção de IR",
    "(-) Retenção (PIS/COFINS/CSLL - 4,65%)",
    "(-) Retenção de Permuta",
    "(-) Transf. Materiais Estoque Laurentino P/ Finan",
    "(-) Transf. Materiais Estoque Laurentino P/ Bom J.",
    "(-) Transf. Mat. Estoque P/ Obras da Laurentino",
    "(-) Reembolso de Despesas Gerais (Obra)",
    "(-) Divisão de muro",
    "(-) Retenção Contribuição Sindical (Escritório)",
    "(-) Reembolso de Despesas Administrativas",
    "(-) Restituição IPTU Imóveis Alugados (Escritório)",
    "(-) Reembolso de Despesas Financeiras",
    "(-) Reembolso de Despesas Tributárias",
    "(-) Abatimento de Adiantamento",
    "(-) Restituição Condomínio",
    "(-) Restituição água"
],
      documentos: [
      "Adiantamento",
      "APORTE DE CAPITAL",
      "AVISO DE LANÇAMENTO",
      "CAUÇÃO",
      "CONHECIMENTO DE FRETE",
      "CUPOM FISCAL",
      "Documento de Arrecadação de Receita Est",
      "Documento de arrecadação federal",
      "Documento de Arrecadação Municipal",
      "Doc. de Arrecadação do Simples Nacional",
      "Documento de arrecadação simples naciona",
      "Distrato",
      "FATURA",
      "FOLHA DE PAGAMENTO",
      "FATURAMENTO DIRETO",
      "Guia do FGTS Digital",
      "Guia de Previdência Social",
      "GUIA RECOLHIMENTO DE COMISSÃO SINDICAL",
      "Guia de recolhimento de FGTS",
      "GUIA PARA RECOLHIMENTO DE IMPOSTOS",
      "INVISÍVEL",
      "INICIALIZAÇÃO DE ESTOQUE",
      "DOC. CONTAB. SISTEMA CUSTO ORÇ. E INCOR",
      "Movimento de Frotas e Equipamento",
      "Movimento de Locação de Equipamentos",
      "Movimento de transf. de equipamento",
      "MOVIMENTO SEGURANÇA SAUDE NO TRABALHO",
      "NOTA FISCAL",
      "NOTA FISCAL ELETRÔNICA",
      "Nota Fiscal de Serviço",
      "PREVISÃO FINANCEIRA DE CONTR. DE MED.",
      "PERMUTA",
      "PREVISÃO FINANCEIRA DE PEDIDOS DE COMPRA",
      "PREVISÃO DE COMISSÃO",
      "PREVISÃO DE DISTRATO",
      "PREVISÃO DE PAGAMENTO/RECEBIMENTO",
      "REALOCAÇÃO DE MATERIAIS",
      "RECIBO",
      "REQUISIÇÃO DE MATERIAIS",
      "SN",
      "DOCUMENTO DE TRANSFERÊNCIA DE MATERIAIS",
      "TRANSF./RESERVA DE ESTOQUE P/SOLICITAÇÃO",
      "TRANSF. DE MATERIAIS POR SOLICITAÇÃO",
      "TAXAS"
    ],
    
    condicoesPagamento: [
      "Parcelas Mensais",
      "Parcelas Semestrais",
      "Parcelas Bimestrais",
      "Entrega das chaves",
      "Parcela na Escritura",
      "PERMUTA",
      "Parcela Única",
      "Parcela Anual",
      "Venda a Vista",
      "Novo Financiamento",
      "Resíduo",
      "Novo Parcelas Mensais",
      "Novo Parcelas Semestrais",
      "Novo Parcela Anual",
      "Novo Venda a Vista",
      "Novo Parcela na Escritura",
      "Novo Parcelas Bimestrais",
      "Novo Parcela Única",
      "Ato",
      "Novo Ato",
      "Provisionado",
      "Financiamento"
    ],
    flags: {
      imprimirParcelasReparceladas: false,
    },
    
    ordem: 'VL',
    
    processarLancamentos: 'CR',
  },
];

const MODAL_SELECTION_DEFINITIONS = [
  {
    key: 'planoFinanceiro',
    titleMatchers: ['plano financeiro'],
    frameSelectors: [
      'input[name="planoFinMultSelectedEntitiesList"]',
      'input#idObjRetorno[value="planoFinMult"]',
      'input#multSelecaoProperty[value="planosFin"]',
      'input[name="entity.nmConta"]',
      'input[name="entity.cdContaMask"]',
    ],
    searchInputSelectors: [
      'input[name="entity.nmConta"]',
      'input[name*="plano"]',
      'input[id*="plano"]',
      '#nmConta',
      'input[name*="nmConta"]',
    ],
    rowValueSelectors: [
      'td[reference="nmConta"]',
      'td[reference="cdContaMask"]',
    ],
    rowCheckboxSelector: 'input[type="checkbox"][name="rowSelect"]',
  },
  {
    key: 'contasCorrente',
    titleMatchers: ['contas correntes', 'conta corrente'],
    frameSelectors: [
      'input[name="contaCorrenteSelectedEntitiesList"]',
      'input#idObjRetorno[value="contaCorrente"]',
      'input#multSelecaoProperty[value="contasCorrentes"]',
      'input[name="entity.nmConta"]',
      'input[name="entity.contaCorrentePK.nuConta"]',
    ],
    searchInputSelectors: [
      'input[name="entity.contaCorrentePK.nuConta"]',
      '#entity.contaCorrentePK.nuConta',
      'input[name="entity.nmConta"]',
      '#entity.nmConta',
    ],
    rowValueSelectors: [
      'td[reference="nmContaCorrente"]',
      'td[reference="contaCorrentePK.nuConta"]',
    ],
    rowCheckboxSelector: 'input[type="checkbox"][name="rowSelect"]',
  },
  {
    key: 'documentos',
    titleMatchers: ['documentos', 'consulta de documentos'],
    inputName: 'docMultFilterContaRecebidas',
    entityName: 'docMultFilterContaRecebidas',
    frameSelectors: [
      'input[name="docMultFilterContaRecebidasSelectedEntitiesList"]',
      'input#idObjRetorno[value="docMultFilterContaRecebidas"]',
      'input#multSelecaoProperty[value="documentos"]',
      'input[name="entity.nmDocumento"]',
      'input[name="entity.documentoPK.cdDocumento"]',
      '#nmDocumento',
      'input[name*="documento"]',
      'input[id*="documento"]',
      'input[name*="nmDocumento"]',
      'input[id*="nmDocumento"]',
    ],
    searchInputSelectors: [
      'input[name="entity.nmDocumento"]',
      'input[name="nmDocumento"]',
      '#nmDocumento',
      'input[name*="documento"]',
      'input[id*="documento"]',
      'input[name*="nmDocumento"]',
      'input[id*="nmDocumento"]',
      'input[placeholder*="Documento"]',
      'input[placeholder*="documento"]',
      'input[aria-label*="Documento"]',
      'input[aria-label*="documento"]',
    ],
    rowValueSelectors: [
      'td[reference="nmDocumento"]',
      'td[reference="documentoPK.cdDocumento"]',
    ],
    rowCheckboxSelector: 'input[type="checkbox"][name="rowSelect"]',
  },
  {
    key: 'condicoesPagamento',
    titleMatchers: ['tipos de condicoes', 'condicoes pagamento', 'condição de pagamento'],
    frameSelectors: [
      'input[name="tipoCondicaoSelectedEntitiesList"]',
      'input#idObjRetorno[value="tipoCondicao"]',
      'input#multSelecaoProperty[value="tiposCondicao"]',
      'input[name="tipoCondicaoFilter.deTipoCondicao"]',
      'input[name="tipoCondicaoFilter.tipoCondicaoPK.cdTipoCondicao"]',
    ],
    searchInputSelectors: [
      'input[name="tipoCondicaoFilter.deTipoCondicao"]',
      'input[name="deTipoCondicao"]',
    ],
    rowValueSelectors: [
      'td[reference="deTipoCondicao"]',
      'td[reference="tipoCondicaoPK.cdTipoCondicao"]',
    ],
    rowCheckboxSelector: 'input[type="checkbox"][name="rowSelect"]',
  },
];

function normalizeSelectionText(value) {
  try {
    return String(value == null ? '' : value)
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^\w\s]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
  } catch (_) {
    return String(value == null ? '' : value)
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
  }
}

function getModalSelectionDefinition(modalTitle) {
  const normalizedTitle = normalizeSelectionText(modalTitle);
  
  for (const definition of MODAL_SELECTION_DEFINITIONS) {
    if (
      definition.titleMatchers.some(
        matcher => normalizedTitle.includes(normalizeSelectionText(matcher))
      )
    ) {
      return definition;
    }
  }
  
  return {
    key: 'generic',
    titleMatchers: [],
    frameSelectors: [],
    searchInputSelectors: [],
    rowValueSelectors: ['td'],
    rowCheckboxSelector: 'input[type="checkbox"][name="rowSelect"]',
  };
}

function getAllPlanoFinanceiroValues() {
  
  return uniqueNonEmpty(
    REPORT_DEFINITIONS.flatMap(
      r => r.planoFinanceiro || []
    )
  );
}

async function resetSomenteCamposEditaveis(surface) {
  
  logEvent({
    level: 'info',
    message: 'Resetando apenas campos editáveis.'
  });
  
  const selectors = [
    '#dtVenctoInicio',
    '#dtVenctoFim',
    '#dtRectoInicio',
    '#dtRectoFim',
    '#dtEmissaoInicio',
    '#dtEmissaoFim',
    '#nuTitulo',
    '#deTipoCondicao',
    '#nmDocumento',
    '#nmConta',
    '[id="cliente.nmCliente"]'
  ];
  
  for (const selector of selectors) {
    const field = surface.locator(selector).first();
    if (!(await field.count().catch(() => 0))) {
      continue;
    }
    
    
    
    await field.evaluate(el => {
      el.value = '';
      el.dispatchEvent(new Event('input', { bubbles: true }));
      el.dispatchEvent(new Event('change', { bubbles: true }));
      el.dispatchEvent(new Event('blur', { bubbles: true }));
    }).catch(() => {});
  }
  
  try {
    await surface.waitForSelector('body', { timeout: 1000 }).catch(() => {});
  } catch (_) {}
}


async function updateReportStatus(data) {
  // Persists a combined status file with overall status and per-report entries.
  // If `data` contains overall keys (running, total, progress, etc.) it will be
  // saved under `overall`. If it contains `perReportEntry: { report, props }`
  // it will merge into `reports[report]`.
  // Serializa acessos ao arquivo de status para evitar condições de corrida
  let cur = {};
  try {
    cur = JSON.parse(await fs.promises.readFile(STATUS_FILE, 'utf8')) || {};
  } catch {}
  
  const overallKeys = ['running', 'startedAt', 'finishedAt', 'total', 'current', 'progress', 'currentReport', 'completedReports', 'error'];
  const isOverall = overallKeys.some(k => Object.prototype.hasOwnProperty.call(data || {}, k));
  if (isOverall) {
    cur.overall = { ...(cur.overall || {}), ...(data || {}) };
    cur.overall.lastUpdated = nowIso();
    cur.overall.completedReports = cur.overall.completedReports || [];
  }
  
  if (data && data.perReportEntry) {
    cur.reports = cur.reports || {};
    // Normalize incoming report identifier to canonical form: report-<index>
    let name = data.perReportEntry.report;
    const incoming = (data.perReportEntry.props || {});
    
    if (typeof name === 'string' && !/^report-\d+$/.test(name)) {
      const foundIndex = REPORT_DEFINITIONS.findIndex(r => r.sheetName === name || r.pdfName === name);
      if (foundIndex >= 0) name = `report-${foundIndex + 1}`;
    }
    
    // Normalize existing cur.reports keys (merge any name-keyed entries into canonical keys)
    try {
      const normalized = {};
      for (const k of Object.keys(cur.reports || {})) {
        let canonical = k;
        if (!/^report-\d+$/.test(k)) {
          const idx = REPORT_DEFINITIONS.findIndex(r => r.sheetName === k || r.pdfName === k);
          if (idx >= 0) canonical = `report-${idx + 1}`;
        }
        normalized[canonical] = { ...(normalized[canonical] || {}), ...(cur.reports[k] || {}) };
        // ensure human-readable name property when possible
        const m = /^report-(\d+)$/.exec(canonical);
        if (m) {
          const i = Number(m[1]) - 1;
          if (REPORT_DEFINITIONS[i] && !normalized[canonical].name) normalized[canonical].name = REPORT_DEFINITIONS[i].sheetName;
        }
      }
      cur.reports = normalized;
    } catch (_) {}
    
    // Normalize and enrich per-report props with useful realtime fields
    const base = cur.reports[name] || {};
    // coerce progress to number when possible, allow explicit null for errors
    if (incoming && Object.prototype.hasOwnProperty.call(incoming, 'progress') && incoming.progress !== null && typeof incoming.progress === 'string') {
      const n = Number(incoming.progress);
      if (!Number.isNaN(n)) incoming.progress = n;
    }
    
    const merged = { ...base, ...incoming };
    
    // Ensure timestamp fields
    if (incoming.status === 'running' && !merged.startedAt) merged.startedAt = nowIso();
    if (incoming.status === 'completed' && !merged.finishedAt) merged.finishedAt = nowIso();
    
    // Standard realtime fields
    merged.lastUpdated = nowIso();
    merged.step = incoming.step || merged.step || null; // current sub-step
    merged.stepStartedAt = incoming.stepStartedAt || merged.stepStartedAt || null;
    // Preserve explicit null (for errors), otherwise ensure numeric progress defaulting to previous or 0
    if (Object.prototype.hasOwnProperty.call(incoming, 'progress')) {
      merged.progress = incoming.progress === null ? null : (typeof incoming.progress === 'number' ? incoming.progress : Number(incoming.progress) || 0);
    } else {
      merged.progress = typeof merged.progress === 'number' ? merged.progress : 0;
    }
    merged.totalItems = typeof incoming.totalItems !== 'undefined' ? incoming.totalItems : (merged.totalItems || null);
    merged.processedItems = typeof incoming.processedItems !== 'undefined' ? incoming.processedItems : (merged.processedItems || 0);
    merged.lastMessage = incoming.lastMessage || merged.lastMessage || null;
    merged.lastError = incoming.lastError || merged.lastError || null;
    
    // Preserve canonical index and name from existing base (do not allow callers to overwrite)
    if (base && typeof base.index !== 'undefined') {
      merged.index = base.index;
    } else if (typeof incoming.index !== 'undefined') {
      merged.index = incoming.index;
    }
    
    if (base && base.name) {
      merged.name = base.name;
    } else if (!merged.name) {
      // try to set human-readable name based on canonical key if possible
      const m = /^report-(\d+)$/.exec(name);
      if (m) {
        const i = Number(m[1]) - 1;
        if (REPORT_DEFINITIONS && REPORT_DEFINITIONS[i]) merged.name = REPORT_DEFINITIONS[i].sheetName;
      }
    }
    
    cur.reports[name] = merged;
  }
  
  // Recalcula resumo overall a partir das entradas por relatório quando disponíveis
  try {
    if (cur.reports) {
      const names = Object.keys(cur.reports);
      const reps = names.map(n => ({ name: n, ...(cur.reports[n] || {}) }));
      const progVals = reps.map(r => (typeof r.progress === 'number' ? r.progress : null)).filter(v => v !== null);
      if (!cur.overall) cur.overall = {};
      if (progVals.length) {
        cur.overall.progress = Math.round(progVals.reduce((a, b) => a + b, 0) / progVals.length);
      } else {
        // keep existing overall.progress or default to 0
        cur.overall.progress = typeof cur.overall.progress === 'number' ? cur.overall.progress : 0;
      }
      cur.overall.total = cur.overall.total || names.length;
      cur.overall.current = reps.filter(r => (r.status === 'completed' || r.status === 'done')).length;
      cur.overall.completedReports = reps.filter(r => (r.status === 'completed' || r.status === 'done')).map(r => ({ report: r.name, finishedAt: r.finishedAt || r.lastUpdated || null }));
      cur.overall.lastUpdated = nowIso();
    }
  } catch (_) {}
  
  // Ensure file contains instance identifier for easier aggregation
  try { cur.instanceId = cur.instanceId || INSTANCE_ID || String(process.pid); } catch (_) {}
  
  // Backward compatibility: if caller passed a plain object without wrapper,
  // and it wasn't recognized as overall, write it as-is.
  // Do not allow callers to replace the whole status file with a single-report
  // object. Always maintain the combined `overall` + `reports` structure so
  // the status file contains detailed progress for all reports.
  
  // Persist atomically: write to temp file then rename
  try {
    const tmpPath = STATUS_FILE + '.tmp';
    await fs.promises.writeFile(tmpPath, JSON.stringify(cur, null, 2), 'utf8');
    await fs.promises.rename(tmpPath, STATUS_FILE);
  } catch (err) {
    // fallback to direct write if atomic rename fails
    try { await fs.promises.writeFile(STATUS_FILE, JSON.stringify(cur, null, 2), 'utf8'); } catch (_) {}
  }
}

// Mutex simples para serializar atualizações do STATUS_FILE entre chamadas assíncronas
let STATUS_UPDATE_LOCK = Promise.resolve();

async function updateReportStatusSerialized(data) {
  const op = async () => updateReportStatus(data);
  // Encadeia a operação na fila
  STATUS_UPDATE_LOCK = STATUS_UPDATE_LOCK.then(op, op);
  return STATUS_UPDATE_LOCK;
}

// Helper para atualizações de progresso por relatório
async function setReportProgress(report, pct, extra) {
  if (!report) return;
  try {
    const props = Object.assign({ progress: pct, lastUpdated: nowIso() }, (extra || {}));
    await updateReportStatusSerialized({ perReportEntry: { report, props } });
  } catch (_) {}
}

function readLog() {
  if (!fs.existsSync(LOG_PATH)) return [];
  try {
    const arr = JSON.parse(fs.readFileSync(LOG_PATH, 'utf8'));
    return Array.isArray(arr) ? arr : [];
  } catch { return []; }
}
function logEvent(event) {
  const payload = { ts: nowIso(), ...event };
  const arr = readLog();
  arr.push(payload);
  try {
    fs.writeFile(LOG_PATH, JSON.stringify(arr, null, 2), 'utf8', () => {});
  } catch (_) {}
  const baseLine = `[${payload.ts}] [${String(payload.level || 'info').toUpperCase()}] ${payload.message || ''}`;
  
  const parts = [];
  if (payload.report) parts.push(`report=${payload.report}`);
  if (typeof payload.progress !== 'undefined') parts.push(`progress=${payload.progress}%`);
  if (payload.current) parts.push(`current=${payload.current}`);
  if (payload.total) parts.push(`total=${payload.total}`);
  if (payload.path) parts.push(`path=${payload.path}`);
  if (payload.url) parts.push(`url=${payload.url}`);
  if (payload.bytes) parts.push(`bytes=${payload.bytes}`);
  if (payload.shotName) parts.push(`shot=${payload.shotName}`);
  if (typeof payload.detail !== 'undefined') parts.push(`detail=${formatLogValue(payload.detail, 400)}`);
  
  if (parts.length) {
    console.log(baseLine + ' - ' + parts.join(' | '));
  } else {
    console.log(baseLine);
  }
}

function truncateForLog(value, maxLen = 500) {
  const s = String(value == null ? '' : value);
  return s.length > maxLen ? s.slice(0, maxLen) + '…' : s;
}

function formatLogValue(value, maxLen = 500) {
  try {
    if (typeof value === 'string') {
      return truncateForLog(value, maxLen);
    }
    if (typeof value === 'number' || typeof value === 'boolean' || value == null) {
      return truncateForLog(value, maxLen);
    }
    if (value instanceof Error) {
      return truncateForLog(JSON.stringify({
        name: value.name,
        message: value.message,
        stack: value.stack,
      }), maxLen);
    }
    return truncateForLog(JSON.stringify(value, null, 0), maxLen);
  } catch (_) {
    return truncateForLog(value, maxLen);
  }
}

// -------------------------
// Profiling / timing helpers
// -------------------------
const _measureTimers = new Map();
function startMeasure(name, meta = {}) {
  try {
    const id = `${name}-${Date.now()}-${Math.random().toString(36).slice(2,8)}`;
    _measureTimers.set(id, { name, meta, start: process.hrtime.bigint() });
    return id;
  } catch (err) {
    return null;
  }
}

async function stopMeasure(id, props = {}) {
  try {
    if (!id) return null;
    const rec = _measureTimers.get(id);
    if (!rec) return null;
    _measureTimers.delete(id);
    const diffNs = process.hrtime.bigint() - rec.start;
    const durationMs = Math.round(Number(diffNs) / 1e6);
    const payload = {
      level: 'info',
      message: `Measure ${rec.name}`,
      detail: JSON.stringify({ ...(rec.meta || {}), ...props, durationMs }),
    };
    try { logEvent(payload); } catch (_) {}
    try {
      // if we have report name in meta, persist a per-report duration
      const reportName = (rec.meta && rec.meta.report) || rec.name;
      await updateReportStatusSerialized({ perReportEntry: { report: reportName, props: { lastDurationMs: durationMs, lastRunAt: nowIso() } } });
    } catch (_) {}
    return durationMs;
  } catch (_) { return null; }
}

async function debugPageSnapshot(page, label) {
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
  if (!SAVE_SHOTS) return null;
  await ensureDir(SCREENSHOT_DIR);
  const file = path.join(SCREENSHOT_DIR, `${Date.now()}-${name}.png`);
  try {
    await page.screenshot({ path: file, fullPage: true });
    return file;
  } catch (e) { return null; }
}
async function saveHtml(page, name) {
  if (!DEBUG_HTML) return null;
  await ensureDir(SCREENSHOT_DIR);
  const file = path.join(SCREENSHOT_DIR, `${Date.now()}-${name}.html`);
  await fs.promises.writeFile(file, await page.content(), 'utf8');
  return file;
}
async function pageSummary(page) {
  let bodyText = '';
  try { bodyText = await page.locator('body').innerText({ timeout: 6000 }); } catch {}
  bodyText = String(bodyText || '').replace(/\s+/g, ' ').trim().slice(0, 3000);
  return { url: page.url(), title: await page.title().catch(() => ''), bodySnippet: bodyText };
}
async function saveStorageStateWithMeta(context) {
  try {
    await context.storageState({ path: STATE_PATH });
  } catch (err) {
    // bubble up
    throw err;
  }
  
  try {
    const meta = { instanceId: INSTANCE_ID, pid: process.pid, savedAt: nowIso(), statePath: STATE_PATH };
    await fs.promises.writeFile(STATE_PATH + '.meta.json', JSON.stringify(meta, null, 2), 'utf8').catch(() => {});
  } catch (e) {}
  
  return STATE_PATH;
}
async function logPageState(page, message, extra = {}) {
  const shot = await saveShot(page, (extra.shotName || 'state').replace(/[^a-z0-9_-]+/gi, '-'));
  const html = await saveHtml(page, (extra.shotName || 'state').replace(/[^a-z0-9_-]+/gi, '-'));
  const summary = await pageSummary(page);
  logEvent({
    level: extra.level || 'info',
    message,
    screenshot: shot,
    html,
    ...summary,
    ...Object.fromEntries(Object.entries(extra).filter(([k]) => !['shotName', 'level'].includes(k))),
  });
}
// Simple in-process modal lock to avoid concurrent modal actions
let _modalLock = false;
async function acquireModalLock(timeout = 15000) {
  const start = Date.now();
  while (_modalLock) {
    if (Date.now() - start > timeout) throw new Error('Timeout acquiring modal lock');
    await new Promise(r => setTimeout(r, 100));
  }
  _modalLock = true;
  return true;
}
function releaseModalLock() {
  _modalLock = false;
  return true;
}

async function saveModalCache() {
  try {
    await fs.promises.mkdir(path.dirname(MODAL_CACHE_PATH), { recursive: true }).catch(() => {});
    await fs.promises.writeFile(MODAL_CACHE_PATH, JSON.stringify(MODAL_CACHE || {}, null, 2), 'utf8').catch(() => {});
  } catch (e) {}
}

async function loadModalCache() {
  try {
    if (fs.existsSync(MODAL_CACHE_PATH)) {
      const txt = await fs.promises.readFile(MODAL_CACHE_PATH, 'utf8').catch(() => '{}');
      MODAL_CACHE = JSON.parse(txt || '{}') || {};
    }
  } catch (e) {
    MODAL_CACHE = {};
  }
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
  
  // espera JS legado: preferir waitForFunction ao invés de timeout fixo
  await page.waitForFunction(() => document.readyState === 'complete', { timeout: 5000 }).catch(() => {});
  
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

async function clickFirstVisible(label, locators) {
  for (const locator of locators) {
    try {
      const count = await locator.count().catch(() => 0);
      if (!count) continue;
      const first = locator.first();
      if (!(await first.isVisible({ timeout: 1500 }).catch(() => false))) continue;
      
      // try to ensure element is in view
      try { await first.scrollIntoViewIfNeeded().catch(() => {}); } catch (e) {}
      
      // Try normal click, then fallbacks: force click, then evaluate click on element handle
      try {
        await first.click({ timeout: 8000 });
        logEvent({ level: 'info', message: `${label}: clique realizado.` });
        return true;
      } catch (err) {
        logEvent({ level: 'debug', message: `${label}: clique padrão falhou, tentando alternativas.`, detail: String(err && err.message || err) });
        try {
          await first.click({ timeout: 3000, force: true });
          logEvent({ level: 'info', message: `${label}: clique realizado (force).` });
          return true;
        } catch (err2) {
          try {
            const handle = await first.elementHandle();
            if (handle) {
              await handle.evaluate((el) => { try { el.scrollIntoView({ block: 'center' }); el.click(); } catch (e) {} });
              logEvent({ level: 'info', message: `${label}: clique realizado via evaluate.` });
              return true;
            }
          } catch (err3) {
            logEvent({ level: 'debug', message: `${label}: evaluate click falhou.`, detail: String(err3 && err3.message || err3) });
          }
        }
      }
    } catch (err) {
      logEvent({ level: 'debug', message: `${label}: tentativa falhou.`, detail: String(err && err.message || err) });
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
async function configureOrdenacao(surface, ordem = 'CR') {
  
  const select = surface.locator('#flOrdem');
  
  await select.waitFor({
    state: 'visible',
    timeout: 15000,
  });
  
  await select.selectOption(ordem);
  
  await select.evaluate(el => {
    el.dispatchEvent(new Event('change', {
      bubbles: true
    }));
  });
  
  logEvent({
    level: 'info',
    message: 'Ordenação configurada.',
    ordem,
  });
}

// =========================================================
// PROCESSAR LANÇAMENTOS
// =========================================================
async function configureTipoLancamento(surface, tipo = 'CR') {
  
  const select = surface.locator('#flTipoSelecao');
  
  await select.waitFor({
    state: 'visible',
    timeout: 10000,
  });
  
  await select.selectOption(tipo);
  
  await select.evaluate(el => {
    el.dispatchEvent(new Event('change', {
      bubbles: true
    }));
  });
  
  logEvent({
    level: 'info',
    message: 'Tipo lançamento configurado.',
    tipo,
  });
}


// =========================================================
// FLAGS
// =========================================================
async function configureFlags(surface, flags = {}) {
  
  // =====================================================
  // IMPRIMIR PARCELAS REPARCELADAS
  // =====================================================
  if (
    typeof flags.imprimirParcelasReparceladas
    !== 'undefined'
  ) {
    
    const checkbox = surface.locator(
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

// =====================================================
// CONTAS CORRENTE
// =====================================================
async function selectContasCorrente({
  page,
  values = [],
  reportName,
}) {
  // defensive normalize: always pass a clean array to selectViaModal
  values = uniqueNonEmpty(Array.isArray(values) ? values.map(x => String(x || '').trim()) : [String(values || '').trim()]);

  await selectViaModal({
    
    page,
    
    triggerSelector:
    'img[onclick*="searchContaCorrenteCompleto.jsp"]',
    
    modalTitle:
    'Consulta de Contas Correntes',
    
    // campo "Nome"
    searchInputSelector: [
      'input[name="entity.nmConta"]',
      '#entity.nmConta',
      
    ],
    
    values,
    reportName,
    
  });
  try { await setReportProgress(reportName, 28, { lastMessage: 'plano_financeiro_selected', step: 'filter_plano' }); } catch (_) {}
}



// =========================================================
// CONFIGURADOR PRINCIPAL
// =========================================================
async function configureReportFilters({
  page,
  report,
}) {
  
  console.log(report);
  
  logEvent({
    level: 'info',
    message:
    `[${report.sheetName}] Iniciando configureReportFilters`,
  });
  
  try { await setReportProgress(report.sheetName, 4, { lastMessage: 'configure_filters_start', step: 'configure_start' }); } catch (_) {}
  
  const __configureMeasure = startMeasure('configureReportFilters', { report: report.sheetName });
  
  let __lastConfigureError = null;
  let attemptConfigure = 0;
  // retry once if frame was detached / execution context destroyed
  while (attemptConfigure < 2) {
    let reportFrame = null;
    try {
      reportFrame = await withFrameRetry(async () => await getReportFilterFrame(page), { name: 'getReportFilterFrame', retries: 2 });

      // =====================================================
      // RESET
      // =====================================================
      await withFrameRetry(async () => await resetSomenteCamposEditaveis(reportFrame), { name: 'resetSomenteCamposEditaveis', retries: 2 }).catch(() => {});
  
  // =====================================================
  // DATAS
  // =====================================================
  
  const now =
  new Date();
  
  const firstDayPrevMonth =
  new Date(
    now.getFullYear(),
    now.getMonth() - 1,
    1
  );
  
  const lastDayPrevMonth =
  new Date(
    now.getFullYear(),
    now.getMonth(),
    0
  );
  
  const formatDate = d =>
    
  String(d.getDate())
  .padStart(2, '0')
  
  + '/'
  
  + String(d.getMonth() + 1)
  .padStart(2, '0')
  
  + '/'
  
  + d.getFullYear();
  
  const dtInicio =
  formatDate(
    firstDayPrevMonth
  );
  
  const dtFim =
  formatDate(
    lastDayPrevMonth
  );
  
  // Allow overriding range via environment variables or constants
  const envStart = (process.env.REPORT_START_DATE || '').trim();
  const envEnd = (process.env.REPORT_END_DATE || '').trim();
  
  const tryParseToDate = s => {
    if (!s) return null;
    s = String(s).trim();
    // Already BR format dd/mm/yyyy
    if (/^\d{2}\/\d{2}\/\d{4}$/.test(s)) return s;
    // ISO yyyy-mm-dd or yyyy/mm/dd
    if (/^\d{4}[-\/]\d{2}[-\/]\d{2}$/.test(s)) {
      const parts = s.split(/[-\/]/);
      return `${parts[2].padStart(2,'0')}/${parts[1].padStart(2,'0')}/${parts[0]}`;
    }
    // dd-mm-yyyy or dd.mm.yyyy
    if (/^\d{2}[-\.]\d{2}[-\.]\d{4}$/.test(s)) {
      const parts = s.split(/[-\.]/);
      return `${parts[0].padStart(2,'0')}/${parts[1].padStart(2,'0')}/${parts[2]}`;
    }
    // Try Date constructor fallback
    const d = new Date(s);
    if (!isNaN(d.getTime())) return formatDate(d);
    return null;
  };
  
  const overriddenDtInicio = tryParseToDate(envStart) || (REPORT_PERIOD_START || dtInicio);
  const overriddenDtFim = tryParseToDate(envEnd) || (REPORT_PERIOD_END || dtFim);
  
  // Log raw + computed values to help debugging when spawned by other processes
  logEvent({ level: 'info', message: `[${report.sheetName}] Date override values`, envStart, envEnd, overriddenDtInicio, overriddenDtFim });
  try { console.log(`[${report.sheetName}] Date override: envStart=${envStart} envEnd=${envEnd} -> ${overriddenDtInicio} - ${overriddenDtFim}`); } catch {}
  
  // =====================================================
  // DATA INÍCIO
  // =====================================================
  
  const dtInicioInput =
  reportFrame.locator(
    '#dtRectoInicio'
  ).first();
  
  await dtInicioInput.waitFor({
    state: 'visible',
    timeout: 30000,
  });
  
  await dtInicioInput.fill('');
  
  await dtInicioInput.type(
    overriddenDtInicio,
    {
      delay: 40,
    }
  );
  
  await dtInicioInput.evaluate(el => {
    
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
  
  // =====================================================
  // DATA FIM
  // =====================================================
  
  const dtFimInput =
  reportFrame.locator(
    '#dtRectoFim'
  ).first();
  
  await dtFimInput.waitFor({
    state: 'visible',
    timeout: 30000,
  });
  
  await dtFimInput.fill('');
  
  await dtFimInput.type(
    overriddenDtFim,
    {
      delay: 40,
    }
  );
  
  await dtFimInput.evaluate(el => {
    
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
  
  const actualDtInicio = await dtInicioInput.inputValue().catch(() => overriddenDtInicio);
  const actualDtFim = await dtFimInput.inputValue().catch(() => overriddenDtFim);
  
  logEvent({
    level: 'info',
    message:
    `[${report.sheetName}] Datas configuradas`,
    overriddenDtInicio,
    overriddenDtFim,
    actualDtInicio,
    actualDtFim,
  });
  try { await setReportProgress(report.sheetName, 6, { lastMessage: 'dates_configured', step: 'dates' }); } catch (_) {}
  
  // =====================================================
  // CHECKBOX PARCELAS REPARCELADAS
  // =====================================================
  
  const reparceladasCheckbox = reportFrame.locator(
    '#flParcelasReparceladas'
  ).first();
  
  if (await reparceladasCheckbox.count()) {
    await reparceladasCheckbox.evaluate(el => {
      if (el.checked) {
        el.checked = false;
        el.dispatchEvent(
          new Event('change', {
            bubbles: true,
          })
        );
      }
    });
    
    logEvent({
      level: 'info',
      message:
      `[${report.sheetName}] Checkbox de parcelas reparceladas desmarcado`,
    });
  }
  
  // =====================================================
  // ORDEM
  // =====================================================
  
  await configureOrdenacao(
    reportFrame,
    report.ordem || 'VL'
  );
  
  try { await setReportProgress(report.sheetName, 7, { lastMessage: 'ordenacao_configured', step: 'ordenacao' }); } catch (_) {}
  
  // =====================================================
  // TIPO LANÇAMENTO
  // =====================================================
  
  await configureTipoLancamento(
    reportFrame,
    report.processarLancamentos || 'CR'
  );
  
  try { await setReportProgress(report.sheetName, 8, { lastMessage: 'tipo_lancamento_configured', step: 'tipo_lancamento' }); } catch (_) {}
  
  // =====================================================
  // FLAGS
  // =====================================================
  
  await configureFlags(
    reportFrame,
    report.flags || {}
  );
  
  try { await setReportProgress(report.sheetName, 9, { lastMessage: 'flags_configured', step: 'flags' }); } catch (_) {}
  
  // =====================================================
  // EXPANDE FILTROS
  // =====================================================
  
  await ensureExpandedFilters(
    reportFrame
  );
  
  await page.waitForTimeout(
    2000
  );
  
  // =====================================================
  // HELPERS
  // =====================================================
  
  const normalizeValues = values =>
    
  uniqueNonEmpty(
    (values || [])
    .map(x => String(x || '').trim())
  );
  
  const removeExcecoes = (
    values,
    excecoes
  ) => {
    
    const exc =
    normalizeValues(excecoes);
    
    return normalizeValues(values)
    .filter(x =>
      !exc.includes(x)
    );
  };
  
  // =====================================================
  // PLANO FINANCEIRO
  // =====================================================

  let planos = normalizeValues(report.planoFinanceiro);

  const selecionarTodosPlanos = !planos.length || planos.includes('*');

  if (!selecionarTodosPlanos) {
    if (planos.length) {
      logEvent({ level: 'info', message: `[${report.sheetName}] Selecionando plano financeiro`, planos });
      try { await setReportProgress(report.sheetName, 10, { lastMessage: 'selecting_plano_financeiro', step: 'plano_start' }); } catch (_) {}
      await selectPlanoFinanceiro({ page, values: planos, reportName: report.sheetName });
    }
  } else {
    logEvent({ level: 'info', message: `[${report.sheetName}] Plano financeiro = TODOS` });
  }
  
  // =====================================================
  // CONTAS CORRENTE
  // =====================================================

  let contas = normalizeValues(report.contasCorrente);

  const selecionarTodasContas = !contas.length || contas.includes('*');

  if (!selecionarTodasContas) {
    if (contas.length) {
      logEvent({ level: 'info', message: `[${report.sheetName}] Selecionando contas corrente`, contas });
      try { await setReportProgress(report.sheetName, 15, { lastMessage: 'selecting_contas', step: 'contas_start' }); } catch (_) {}
      await selectContasCorrente({ page, values: contas, reportName: report.sheetName });
    }
  } else {
    logEvent({ level: 'info', message: `[${report.sheetName}] Contas corrente = TODAS` });
  }
  
  // =====================================================
  // DOCUMENTOS
  // =====================================================

  let documentos = normalizeValues(report.documentos);

  const selecionarTodosDocs = !documentos.length || documentos.includes('*');

  if (!selecionarTodosDocs) {
    if (documentos.length) {
      logEvent({ level: 'info', message: `[${report.sheetName}] Selecionando documentos`, documentos });
      try { await setReportProgress(report.sheetName, 20, { lastMessage: 'selecting_documentos', step: 'documentos_start' }); } catch (_) {}
      await selectDocumentos({ page, values: documentos, enablePagination: true, reportName: report.sheetName });
    }
  } else {
    logEvent({ level: 'info', message: `[${report.sheetName}] Documentos = TODOS` });
  }
  
  // =====================================================
  // CONDIÇÕES PAGAMENTO
  // =====================================================

  let condicoes = normalizeValues(report.condicoesPagamento);

  const selecionarTodasCondicoes = !condicoes.length || condicoes.includes('*');

  if (!selecionarTodasCondicoes) {
    if (condicoes.length) {
      logEvent({ level: 'info', message: `[${report.sheetName}] Selecionando condições`, condicoes });
      try { await setReportProgress(report.sheetName, 25, { lastMessage: 'selecting_condicoes', step: 'condicoes_start' }); } catch (_) {}
      await selectCondicoesPagamento({ page, values: condicoes, reportName: report.sheetName });
    }
  } else {
    logEvent({ level: 'info', message: `[${report.sheetName}] Condições pagamento = TODAS` });
  }
  
  // =====================================================
  // FINAL
  // =====================================================
  
  await page.waitForTimeout(
    3000
  );
  logEvent({
    level: 'info',
    message:
    `[${report.sheetName}] Filtros finalizados`,
  });
  try { await setReportProgress(report.sheetName, 30, { lastMessage: 'filters_finished', step: 'filters_done' }); } catch (_) {}
  
      await stopMeasure(__configureMeasure, { phase: 'configureFilters' }).catch(() => {});
      // success
      __lastConfigureError = null;
      break;
    } catch (err) {
      __lastConfigureError = err;
      const msg = String(err && err.message || err);
      // retry only for frame-related transient errors
      if (attemptConfigure === 0 && (
        msg.includes('Frame was detached') ||
        msg.includes('Execution context was destroyed') ||
        msg.includes('Target closed')
      )) {
        logEvent({ level: 'warn', message: 'configureReportFilters: quadro foi recriado; refazendo tentativa', detail: msg });
        attemptConfigure++;
        await page.waitForTimeout(1500);
        continue;
      }
      // non-recoverable or second failure: rethrow
      throw err;
    }
  }
  if (__lastConfigureError) {
    // surface the last error if we exhausted retries
    throw __lastConfigureError;
  }
}

function uniqueNonEmpty(values = []) {
  
  return [...new Set(
    values.filter(v => v?.trim())
  )];
}

// =====================================================
// SELECIONA CONDIÇÕES PAGAMENTO
// =====================================================
async function selectCondicoesPagamento({
  page,
  values = [],
  reportName,
}) {
  // defensive normalize: always pass a clean array to selectViaModal
  values = uniqueNonEmpty(Array.isArray(values) ? values.map(x => String(x || '').trim()) : [String(values || '').trim()]);

  await selectViaModal({
    
    page,
    
    triggerSelector:
    'img[onclick*="consMultSelTipoCondicao"]',
    
    modalTitle:
    'Consulta de Tipos de Condições',
    
    searchInputSelector: [
      
      'input[name="tipoCondicaoFilter.deTipoCondicao"]',
      'input[name="deTipoCondicao"]',
      
    ],
    
    values,
    reportName,
    
  });
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

async function resolveDocumentosRenderingMode(modalFrame) {
  return await modalFrame.evaluate(() => {
    try {
      const currentDoc = document;
      const currentRows = currentDoc.querySelectorAll('tr[id^="linha_"]').length;
      const currentPager = !!currentDoc.querySelector('#infoTotalPaginas, .spwBotaoProximo, input[type="button"].spwBotaoProximo[title="Próxima página"]');

      const iframe = currentDoc.querySelector('#layerFormConsulta');
      const nestedDoc = iframe && (iframe.contentDocument || iframe.contentWindow?.document);
      const nestedRows = nestedDoc ? nestedDoc.querySelectorAll('tr[id^="linha_"]').length : 0;
      const nestedPager = !!(nestedDoc && nestedDoc.querySelector('#infoTotalPaginas, .spwBotaoProximo, input[type="button"].spwBotaoProximo[title="Próxima página"]'));

      return {
        useNested: !currentRows && (nestedRows > 0 || nestedPager),
        currentRows,
        nestedRows,
        hasCurrentPager: currentPager,
        hasNestedPager: nestedPager,
      };
    } catch (error) {
      return { useNested: false, currentRows: 0, nestedRows: 0, hasCurrentPager: false, hasNestedPager: false };
    }
  }).catch(() => ({ useNested: false, currentRows: 0, nestedRows: 0, hasCurrentPager: false, hasNestedPager: false }));
}

async function getDocumentosTotalPages(modalFrame, useNested = false) {
  const rawInfo = await modalFrame.evaluate((useNestedDoc) => {
    try {
      const currentDoc = document;
      const doc = useNestedDoc
        ? (currentDoc.querySelector('#layerFormConsulta')?.contentDocument || currentDoc.querySelector('#layerFormConsulta')?.contentWindow?.document)
        : currentDoc;

      return doc?.querySelector('#infoTotalPaginas')?.innerText || '';
    } catch (error) {
      return '';
    }
  }, useNested).catch(() => '');

  const match = String(rawInfo || '').match(/de\s+(\d+)/i);
  return match ? Number(match[1]) || 1 : 1;
}

async function waitForDocumentosDom(modalFrame, timeout = 45000) {
  const started = Date.now();

  while (Date.now() - started < timeout) {
    const info = await modalFrame.evaluate(() => {
      try {
        const iframe = document.querySelector('#layerFormConsulta');
        const doc = iframe && (iframe.contentDocument || iframe.contentWindow?.document) || document;
        const rowsCount = doc.querySelectorAll('tr[id^="linha_"]').length;
        const pagerText = String(doc.querySelector('#infoTotalPaginas')?.innerText || '').trim();
        const nextExists = !!doc.querySelector('input[type="button"].spwBotaoProximo[title="Próxima página"], .spwBotaoProximo');
        const inputExists = !!doc.querySelector('input[name="entity.nmDocumento"], input[id="nmDocumento"], input[name*="documento"], input[type="search"]');
        const totalMatch = pagerText.match(/de\s+(\d+)/i);
        return {
          ready: rowsCount > 0 || nextExists || inputExists || !!pagerText,
          useNested: !!iframe && doc !== document,
          rowsCount,
          totalPages: totalMatch ? Number(totalMatch[1]) || 1 : 1,
        };
      } catch (error) {
        return { ready: false, useNested: false, rowsCount: 0, totalPages: 1 };
      }
    }).catch(() => ({ ready: false, useNested: false, rowsCount: 0, totalPages: 1 }));

    if (info.ready) {
      return info;
    }

    await modalFrame.page().waitForTimeout(500).catch(() => {});
  }

  return { ready: false, useNested: false, rowsCount: 0, totalPages: 1 };
}

async function selectDocumentosAcrossPages({
  page,
  modalFrame,
  preferredFrame = null,
  modalTitle,
  values = [],
  rowCheckboxSelector = 'input[type="checkbox"][name="rowSelect"]',
}) {
  const desiredEntries = uniqueNonEmpty(
    Array.isArray(values)
      ? values.map(value => String(value || '').trim())
      : [String(values || '').trim()]
  );

  if (!desiredEntries.length) {
    return { found: 0, expected: 0, matched: [], documents: [], totalPages: 1 };
  }

  const desiredMap = new Map(
    desiredEntries.map(value => [normalizeSelectionText(value), value])
  );
  const pageRef = modalFrame?.page?.() || page || null;

  function normalize(value) {
    return normalizeSelectionText(value);
  }

  async function resolveDocumentosFrame(sourceFrame) {
    const pageRef = sourceFrame.page();
    const started = Date.now();
    let lastInfo = null;

    while (Date.now() - started < 15000) {
      lastInfo = await sourceFrame.evaluate(() => {
        try {
          const currentDoc = document;
          const directRows = currentDoc.querySelectorAll('tr[id^="linha_"]').length;
          const directPagerText = String(currentDoc.querySelector('#infoTotalPaginas')?.innerText || '').trim();
          const directNextExists = !!currentDoc.querySelector('input[type="button"].spwBotaoProximo[title="Próxima página"], .spwBotaoProximo');
          const iframe = currentDoc.querySelector('#layerFormConsulta');
          const nestedDoc = iframe && (iframe.contentDocument || iframe.contentWindow?.document);
          const nestedRows = nestedDoc ? nestedDoc.querySelectorAll('tr[id^="linha_"]').length : 0;
          const nestedPagerText = String(nestedDoc?.querySelector('#infoTotalPaginas')?.innerText || '').trim();
          const nestedNextExists = !!(nestedDoc && nestedDoc.querySelector('input[type="button"].spwBotaoProximo[title="Próxima página"], .spwBotaoProximo'));

          return {
            hasIframe: !!iframe,
            directRows,
            directPagerText,
            directNextExists,
            nestedRows,
            nestedPagerText,
            nestedNextExists,
          };
        } catch (error) {
          return {
            hasIframe: false,
            directRows: 0,
            directPagerText: '',
            directNextExists: false,
            nestedRows: 0,
            nestedPagerText: '',
            nestedNextExists: false,
          };
        }
      }).catch(() => null);

      const directReady = !!lastInfo && (
        lastInfo.directRows > 0 ||
        !!lastInfo.directPagerText ||
        lastInfo.directNextExists
      );
      const nestedReady = !!lastInfo && lastInfo.hasIframe && (
        lastInfo.nestedRows > 0 ||
        !!lastInfo.nestedPagerText ||
        lastInfo.nestedNextExists
      );

      if (nestedReady) {
        const iframeLocator = sourceFrame.locator('#layerFormConsulta').first();
        const iframeHandle = await iframeLocator.elementHandle().catch(() => null);
        const nestedFrame = iframeHandle ? await iframeHandle.contentFrame().catch(() => null) : null;
        if (nestedFrame) {
          return { frame: nestedFrame, mode: 'nested', info: lastInfo };
        }
      }

      if (directReady) {
        return { frame: sourceFrame, mode: 'direct', info: lastInfo };
      }

      await pageRef?.waitForTimeout(400).catch(() => {});
    }

    const iframeLocator = sourceFrame.locator('#layerFormConsulta').first();
    const iframeHandle = await iframeLocator.elementHandle().catch(() => null);
    const nestedFrame = iframeHandle ? await iframeHandle.contentFrame().catch(() => null) : null;
    return {
      frame: nestedFrame || sourceFrame,
      mode: nestedFrame ? 'nested' : 'direct',
      info: lastInfo,
    };
  }

  async function captureDocumentosPage(targetFrame) {
    return await targetFrame.evaluate(({ desiredValues, rowCheckboxSelector }) => {
      function normalizeLocal(value) {
        try {
          return String(value || '')
            .normalize('NFD')
            .replace(/[\u0300-\u036f]/g, '')
            .replace(/[^\w\s]/g, ' ')
            .replace(/\s+/g, ' ')
            .trim()
            .toLowerCase();
        } catch (error) {
          return String(value || '')
            .replace(/[^\w\s]/g, ' ')
            .replace(/\s+/g, ' ')
            .trim()
            .toLowerCase();
        }
      }

      function getDoc() {
        const currentDoc = document;
        const iframe = currentDoc.querySelector('#layerFormConsulta');
        const nestedDoc = iframe && (iframe.contentDocument || iframe.contentWindow?.document);
        const directRows = currentDoc.querySelectorAll('tr[id^="linha_"]').length;
        const nestedRows = nestedDoc ? nestedDoc.querySelectorAll('tr[id^="linha_"]').length : 0;
        const directPager = String(currentDoc.querySelector('#infoTotalPaginas')?.innerText || '').trim();
        const nestedPager = String(nestedDoc?.querySelector('#infoTotalPaginas')?.innerText || '').trim();
        const directNext = !!currentDoc.querySelector('input[type="button"].spwBotaoProximo[title="Próxima página"], .spwBotaoProximo');
        const nestedNext = !!(nestedDoc && nestedDoc.querySelector('input[type="button"].spwBotaoProximo[title="Próxima página"], .spwBotaoProximo'));

        if (nestedDoc && (nestedRows > 0 || nestedPager || nestedNext || (!directRows && !directPager && !directNext))) {
          return nestedDoc;
        }
        return currentDoc;
      }

      function getUrl(doc) {
        try {
          const iframe = document.querySelector('#layerFormConsulta');
          if (iframe?.contentWindow?.location?.href) {
            return iframe.contentWindow.location.href;
          }
        } catch (error) {}
        try {
          return doc?.defaultView?.location?.href || window.location.href || '';
        } catch (error) {
          return '';
        }
      }

      function getTotalPages(doc) {
        const raw = String(doc.querySelector('#infoTotalPaginas')?.innerText || '');
        const match = raw.match(/de\s+(\d+)/i);
        return match ? Number(match[1]) || 1 : 1;
      }

      const doc = getDoc();
      const rows = Array.from(doc.querySelectorAll('tr[id^="linha_"]'));
      const documentos = [];
      const matchedOnPage = [];
      const wanted = new Map(desiredValues.map(value => [normalizeLocal(value), value]));

      rows.forEach(tr => {
        const nome = tr.querySelectorAll('td')[2]?.innerText.trim();
        if (!nome) return;
        documentos.push(nome);

        const normalizedName = normalizeLocal(nome);
        if (!wanted.has(normalizedName)) return;

        const checkbox = tr.querySelector(rowCheckboxSelector || 'input[type="checkbox"]');
        if (checkbox && !checkbox.checked) {
          try {
            checkbox.click();
          } catch (error) {
            try {
              checkbox.checked = true;
              checkbox.dispatchEvent(new Event('change', { bubbles: true }));
            } catch (fallbackError) {}
          }
          try {
            checkbox.dispatchEvent(new Event('change', { bubbles: true }));
          } catch (error) {}
        }
        matchedOnPage.push(nome);
      });

      const pagerText = String(doc.querySelector('#infoTotalPaginas')?.innerText || '').trim();
      return {
        documentos,
        matchedOnPage,
        rowCount: rows.length,
        totalPaginas: getTotalPages(doc),
        currentUrl: getUrl(doc),
        pagerText,
      };
    }, {
      desiredValues: desiredEntries,
      rowCheckboxSelector,
    }).catch((error) => ({
      documentos: [],
      matchedOnPage: [],
      rowCount: 0,
      totalPaginas: 1,
      currentUrl: '',
      pagerText: '',
      error: String(error && (error.stack || error.message || error)),
    })).then(result => ({
      ...result,
      frameUrl: targetFrame.url() || '',
    }));
  }

  async function clickNextPage(targetFrame, pageIndex) {
    const selectors = [
      'input[type="button"].spwBotaoProximo[title="Próxima página"]',
      'input.spwBotaoProximo[title="Próxima página"]',
      '.spwBotaoProximo',
      'input[type="button"][onclick*="paginar(\'proximo\'"]',
      'input[type="button"][onclick*="paginar("]',
    ];

    logEvent({
      level: 'debug',
      message: 'Iniciando tentativa de avançar página de documentos',
      modalTitle,
      detail: {
        pageIndex,
        selectors: selectors.length,
      },
    });

    for (const selector of selectors) {
      logEvent({
        level: 'debug',
        message: 'Testando seletor de próxima página',
        modalTitle,
        detail: { selector, pageIndex },
      });

      const locator = targetFrame.locator(selector).first();
      const count = await locator.count().catch(() => 0);
      if (!count) {
        logEvent({
          level: 'debug',
          message: 'Seletor de próxima página inexistente ou desabilitado',
          modalTitle,
          detail: { selector, pageIndex, frameState: { exists: false } },
        });
        continue;
      }

      const state = await locator.evaluate((el) => {
        const attrs = {};
        Array.from(el.attributes || []).forEach(attr => {
          attrs[attr.name] = attr.value;
        });
        return {
          title: el.getAttribute('title') || '',
          tagName: el.tagName,
          className: el.className || '',
          outerHTML: el.outerHTML ? el.outerHTML.slice(0, 450) : '',
          hasPaginar: typeof window.paginar === 'function',
          hasParentPaginar: !!(window.parent && window.parent !== window && typeof window.parent.paginar === 'function'),
          hasTopPaginar: !!(window.top && window.top !== window && typeof window.top.paginar === 'function'),
          attributes: attrs,
        };
      }).catch(() => ({
        title: '',
        tagName: '',
        className: '',
        outerHTML: '',
        hasPaginar: false,
        hasParentPaginar: false,
        hasTopPaginar: false,
        attributes: {},
      }));

      logEvent({
        level: 'info',
        message: 'Tentando avançar página de documentos',
        modalTitle,
        detail: {
          selector,
          pageIndex,
          scope: targetFrame === modalFrame ? 'modalFrame' : 'nestedFrame',
          ...state,
          shotPath: await saveShot(page, `documentos-avancar-pagina-${pageIndex}`).catch(() => null),
        },
      });

      let clicked = false;
      try {
        await locator.scrollIntoViewIfNeeded().catch(() => {});
        await locator.click({ force: true, timeout: 8000 });
        clicked = true;
      } catch (clickError) {
        logEvent({
          level: 'debug',
          message: 'Clique padrão na próxima página falhou',
          modalTitle,
          detail: {
            selector,
            pageIndex,
            error: String(clickError && (clickError.message || clickError) || clickError),
          },
        });
        try {
          const handle = await locator.elementHandle().catch(() => null);
          if (handle) {
            await handle.evaluate(el => {
              try {
                el.scrollIntoView({ block: 'center', inline: 'center' });
                el.click();
              } catch (error) {}
            });
            clicked = true;
          }
        } catch (fallbackError) {
          logEvent({
            level: 'debug',
            message: 'Clique via evaluate na próxima página falhou',
            modalTitle,
            detail: {
              selector,
              pageIndex,
              error: String(fallbackError && (fallbackError.message || fallbackError) || fallbackError),
            },
          });
        }
      }

      logEvent({
        level: 'debug',
        message: 'Resultado da tentativa de clique em próxima página',
        modalTitle,
        detail: { selector, pageIndex, clicked },
      });

      if (clicked) {
        return { clicked: true, selector, state };
      }
    }

    logEvent({
      level: 'debug',
      message: 'Não foi possível avançar a página de documentos',
      modalTitle,
      detail: { pageIndex },
    });

    return { clicked: false };
  }

  const resolved = await withFrameRetry(async () => {
    const frameCandidate = await resolveDocumentosFrame(modalFrame);
    return frameCandidate;
  }, { name: 'selectDocumentos:resolve-frame', retries: 2 }).catch((error) => {
    logEvent({
      level: 'warn',
      message: 'Seleção paginada de documentos falhou ao resolver o frame do modal.',
      modalTitle,
      detail: {
        desiredCount: desiredEntries.length,
        error: String(error && (error.stack || error.message || error)),
      },
    });
    return null;
  });

  if (!resolved || !resolved.frame) {
    return {
      found: 0,
      expected: desiredEntries.length,
      matched: [],
      documents: [],
      totalPages: 1,
    };
  }

  const docsFrame = resolved.frame;
  const pageResults = [];
  const documents = new Set();
  const matched = new Set();
  const maxPages = 50;

  logEvent({
    level: 'debug',
    message: 'Iniciando seleção paginada de documentos',
    modalTitle,
    detail: {
      desiredCount: desiredEntries.length,
      resolvedMode: resolved.mode,
      resolvedInfo: resolved.info,
    },
  });

  let firstPage = await captureDocumentosPage(docsFrame);
  if ((!firstPage.rowCount || firstPage.rowCount === 0) && resolved.mode === 'nested') {
    await pageRef?.waitForTimeout(1000).catch(() => {});
    firstPage = await captureDocumentosPage(docsFrame);
  }

  pageResults.push({ pagina: 1, ...firstPage });
  firstPage.documentos.forEach(nome => documents.add(nome));
  firstPage.matchedOnPage.forEach(nome => matched.add(nome));

  let totalPages = Math.max(1, Number(firstPage.totalPaginas) || 1);

  logEvent({
    level: 'debug',
    message: 'Detalhe da página de documentos lida',
    modalTitle,
    detail: {
      pageIndex: 1,
      totalPages,
      documentsCount: firstPage.documentos.length,
      matchedOnPageCount: firstPage.matchedOnPage.length,
      capturedCount: documents.size,
      matchedCount: matched.size,
      sampleDocuments: firstPage.documentos.slice(0, 10),
      pageResult: { pagina: 1, ...firstPage },
    },
  });

  for (let pagina = 2; pagina <= totalPages && pagina <= maxPages; pagina++) {
    logEvent({
      level: 'debug',
      message: 'Solicitando avanço para próxima página de documentos',
      modalTitle,
      detail: {
        pageIndex: pagina - 1,
        totalPages,
        matchedCount: matched.size,
      },
    });

    const previousUrl = pageResults[pageResults.length - 1]?.frameUrl || pageResults[pageResults.length - 1]?.currentUrl || '';
    const advanceResult = await clickNextPage(docsFrame, pagina - 1);

    logEvent({
      level: 'debug',
      message: 'Resultado da solicitação de avanço',
      modalTitle,
      detail: {
        pageIndex: pagina - 1,
        totalPages,
        advanced: !!advanceResult.clicked,
      },
    });

    if (!advanceResult.clicked) {
      break;
    }

    let advanced = false;
    const started = Date.now();
    while (Date.now() - started < 20000) {
      const currentResolved = await withFrameRetry(async () => await resolveDocumentosFrame(modalFrame), {
        name: 'selectDocumentos:resolve-frame-after-advance',
        retries: 1,
      }).catch(() => null);
      const currentFrame = currentResolved?.frame || docsFrame;
      const currentPage = await captureDocumentosPage(currentFrame);
      const currentFrameUrl = currentPage.frameUrl || currentPage.currentUrl || currentFrame.url() || '';
      if (currentFrameUrl && currentFrameUrl !== previousUrl && currentPage.rowCount > 0) {
        advanced = true;
        pageResults.push({ pagina, ...currentPage, advanced });
        currentPage.documentos.forEach(nome => documents.add(nome));
        currentPage.matchedOnPage.forEach(nome => matched.add(nome));

        totalPages = Math.max(totalPages, Number(currentPage.totalPaginas) || totalPages || 1);
        logEvent({
          level: 'info',
          message: `Página ${pagina}/${totalPages} -> ${currentPage.rowCount} registros`,
          modalTitle,
          detail: {
            matched: matched.size,
            rowCount: currentPage.rowCount || 0,
          },
        });
        logEvent({
          level: 'debug',
          message: 'Detalhe da página de documentos lida',
          modalTitle,
          detail: {
            pageIndex: pagina,
            totalPages,
            documentsCount: currentPage.documentos.length,
            matchedOnPageCount: currentPage.matchedOnPage.length,
            capturedCount: documents.size,
            matchedCount: matched.size,
            sampleDocuments: currentPage.documentos.slice(0, 10),
            pageResult: { pagina, ...currentPage, advanced },
          },
        });
        break;
      }

      await pageRef?.waitForTimeout(500).catch(() => {});
    }

    if (!advanced) {
      const fallbackResolved = await withFrameRetry(async () => await resolveDocumentosFrame(modalFrame), {
        name: 'selectDocumentos:resolve-frame-fallback',
        retries: 1,
      }).catch(() => null);
      const fallbackFrame = fallbackResolved?.frame || docsFrame;
      const fallbackPage = await captureDocumentosPage(fallbackFrame);
      pageResults.push({ pagina, ...fallbackPage, advanced: false });
      fallbackPage.documentos.forEach(nome => documents.add(nome));
      fallbackPage.matchedOnPage.forEach(nome => matched.add(nome));

        logEvent({
          level: 'warn',
          message: 'Não foi possível confirmar a troca de página de documentos',
          modalTitle,
          detail: {
            pageIndex: pagina - 1,
            previousUrl,
            currentUrl: fallbackPage.frameUrl || fallbackPage.currentUrl || '',
            rowCount: fallbackPage.rowCount || 0,
            totalPages,
          },
        });
      break;
    }
  }

  logEvent({
    level: 'debug',
    message: 'Seleção paginada de documentos finalizada',
    modalTitle,
    detail: {
      found: matched.size,
      expected: desiredEntries.length,
      totalPages,
      capturedDocuments: documents.size,
      pageResults,
    },
  });

  return {
    found: matched.size,
    expected: desiredEntries.length,
    matched: Array.from(matched),
    documents: Array.from(documents),
    totalPages,
    pageResults,
  };
}

async function ensureSubmitEnabled(surface) {
  const state = {
    hasForm: (await surface.locator('form').count().catch(() => 0)) > 0,
    isEnableSubmit: await surface.locator('body').evaluate(() => {
      return typeof window.IS_enableSubmit === 'undefined'
      ? null
      : window.IS_enableSubmit;
    }).catch(() => null),
  };
  
  if (!state.hasForm) {
    throw new Error('Não encontrei o formulário do relatório para habilitar o submit.');
  }
  
  if (state.isEnableSubmit !== true) {
    logEvent({
      level: 'debug',
      message: 'IS_enableSubmit estava desativado; reabilitando antes do submit do relatório.',
      currentState: state.isEnableSubmit,
    });
    await surface.locator('body').evaluate(() => {
      window.IS_enableSubmit = true;
    }).catch(() => {});
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
  // Não limpar o campo para preservar valores já selecionados; apenas foca e digita o novo valor
  try { await locator.focus(); } catch (_) {}
  try {
    await locator.type(String(value), { delay: 15 });
  } catch (_) {
    try { await locator.fill(String(value)); } catch (_) {}
  }
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
  filePath,
  maxRetries = 4
) {
  
  const __saveMeasure = startMeasure('saveReportFromContext', { reportUrl, filePath });
  
  await ensureDir(path.dirname(filePath));
  
  const shouldRetryError = (error) => {
    if (!error) return false;
    const message = String(error.message || '').toLowerCase();
    return (
      message.includes('socket hang up') ||
      message.includes('connection reset') ||
      message.includes('econnreset') ||
      message.includes('etimedout') ||
      message.includes('timeout') ||
      message.includes('503') ||
      message.includes('502') ||
      message.includes('504')
    );
  };
  
  for (let attempt = 0; attempt < maxRetries; attempt++) {
    try {
      const response = await context.request.get(
        reportUrl,
        {
          timeout: 120000,
        }
      );
      
      if (!response.ok()) {
        const status = response.status();
        if (attempt < maxRetries - 1 && [500, 502, 503, 504].includes(status)) {
          logEvent({
            level: 'warning',
            message: `saveReportFromContext: HTTP ${status} temporário. Tentando novamente.`,
            attempt: attempt + 1,
            reportUrl,
          });
          await sleepMs(2000);
          continue;
        }
        
        throw new Error(
          `Falha ao baixar relatório: HTTP ${status}`
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
      
      await stopMeasure(__saveMeasure, { phase: 'download', attempt }).catch(() => {});
      return filePath;
    } catch (error) {
      if (attempt < maxRetries - 1 && shouldRetryError(error)) {
        logEvent({
          level: 'warning',
          message: `saveReportFromContext: tentativa ${attempt + 1}/${maxRetries} falhou por erro temporário. Retry.`,
          reportUrl,
          detail: String(error.message || error),
        });
        await sleepMs(2000);
        continue;
      }
      
      await stopMeasure(__saveMeasure, { phase: 'download-error', attempt, error: String(error && error.message || error) }).catch(() => {});
      
      throw error;
    }
  }
  
  throw new Error(`Falha ao baixar relatório após ${maxRetries} tentativas: ${reportUrl}`);
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

async function ensureExpandedFilters(surface) {
  
  if (!surface) {
    
    throw new Error(
      'ensureExpandedFilters: surface inválido.'
    );
  }
  
  const field = surface.locator(
    '#deTipoCondicao'
  );
  
  // já expandido
  if (await field.isVisible().catch(() => false)) {
    
    logEvent({
      level: 'info',
      message: 'Filtros já expandidos.',
    });
    
    return;
  }
  
  logEvent({
    level: 'info',
    message: 'Expandindo filtros.',
  });
  
  const toggle = surface.locator(
    'img[name="toggleFiltro"]'
  ).first();
  
  await toggle.waitFor({
    state: 'attached',
    timeout: 15000,
  });
  
  await toggle.scrollIntoViewIfNeeded();
  
  await toggle.click({
    force: true,
  });
  
  await new Promise(resolve => setTimeout(resolve, 1500));
  
  await field.waitFor({
    state: 'visible',
    timeout: 15000,
  });
  
  logEvent({
    level: 'info',
    message: 'Filtros expandidos.',
  });
}

async function getReportFilterFrame(page) {
  const iframeLocator =
  page.locator('iframe#iFramePage').first();
  
  await iframeLocator.waitFor({
    state: 'visible',
    timeout: 120000,
  });
  
  const frame =
  await iframeLocator.contentFrame();
  
  if (!frame) {
    throw new Error(
      'Não foi possível obter o frame iFramePage da tela de relatório.'
    );
  }
  
  return frame;
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
    message: 'Abrindo tela Relatório Contas Recebidas.',
  });
  
  //const __openReportsMeasure = startMeasure('openReportsPage');
  
  await fetchWithRetry(
    page,
    REPORT_FILTER_PAGE_URL,
    {
      waitUntil: 'domcontentloaded',
      timeout: 120000,
    }
  );
  
  // =====================================================
  // AGUARDA IFRAME PRINCIPAL
  // =====================================================
  
  const iframe =
  page.frameLocator(
    'iframe#iFramePage'
  );
  
  await iframe
  .locator('#btFiltrar')
  .waitFor({
    state: 'visible',
    timeout: 120000,
  });
  
  // =====================================================
  // ESPERA AJAX LEGADO
  // =====================================================
  
  await page.waitForTimeout(4000);
  
  logEvent({
    level: 'info',
    message: 'Tela de relatório carregada.',
    url: page.url(),
  });
  
  //await stopMeasure(__openReportsMeasure, { phase: 'openReportsPage' }).catch(() => {});
  
  return true;
}

async function goDirectToReportsPageAfterLogin(page) {
  logEvent({ level: 'info', message: 'Após login válido, tentando ir direto para a tela de relatório de contas recebidas.' });
  
  await page.waitForTimeout(2500);
  await dismissHomeOverlays(page);
  await page.waitForTimeout(1200);
  
  const attempts = [
    async () => {
      await fetchWithRetry(page, REPORT_FILTER_PAGE_URL, { waitUntil: 'domcontentloaded' });
      await waitForAppReady(page, 12000);
    },
    async () => {
      await fetchWithRetry(page, `${BASE_URL}/sienge/8/index.html`, { waitUntil: 'domcontentloaded' });
      await waitForAppReady(page, 12000);
      await fetchWithRetry(page, REPORT_FILTER_PAGE_URL, { waitUntil: 'domcontentloaded' });
      await waitForAppReady(page, 12000);
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

// =====================================================
// FRAME HELPERS
// =====================================================
async function findSelectionFrame(
  page,
  frameName,
  debugLabel,
  selectors = []
) {
  
  await page.waitForTimeout(4000);
  
  for (const frame of page.frames()) {
    
    try {
      
      const hasSelecionar = await frame.locator(
        'input[name="pbSelecionar"]'
      ).count();
      
      if (!hasSelecionar) {
        continue;
      }
      
      const url = frame.url();
      
      let matches =
      url.includes(frameName);
      
      if (!matches && selectors.length) {
        for (const selector of selectors) {
          const count = await frame.locator(selector).count().catch(() => 0);
          if (count > 0) {
            matches = true;
            break;
          }
        }
      }
      
      if (matches) {
        
        logEvent({
          level: 'info',
          message: `${debugLabel} encontrado.`,
          url,
        });
        
        return frame;
      }
      
    } catch (_) {}
  }
  
  return null;
}


// =====================================================
// FRAME DOCUMENTOS
// =====================================================
async function getDocumentosFrame(page) {
  
  return findSelectionFrame(
    page,
    'documento',
    'Frame de Documentos',
    getModalSelectionDefinition('Consulta de Documentos').frameSelectors
  );
}

// =====================================================
// FRAME CONDIÇÕES PAGAMENTO
// =====================================================
async function getCondicaoPagamentoFrame(page) {
  
  return findSelectionFrame(
    page,
    'condicao',
    'Frame de Condição de Pagamento',
    getModalSelectionDefinition('Consulta de Tipos de Condições').frameSelectors
  );
}

async function getPlanoFinanceiroFrame(page) {
  
  await page.waitForTimeout(4000);
  
  for (const frame of page.frames()) {
    
    try {
      
      const hasSelecionar = await frame.locator(
        'input[name="pbSelecionar"]'
      ).count();
      
      if (hasSelecionar > 0) {
        const selectors = getModalSelectionDefinition('Consulta de Plano Financeiro').frameSelectors;
        let matches = false;
        for (const selector of selectors) {
          if (await frame.locator(selector).count().catch(() => 0)) {
            matches = true;
            break;
          }
        }
        
        if (!matches) continue;
        
        logEvent({
          level: 'info',
          message: 'Frame do Plano Financeiro encontrado.',
          url: frame.url(),
        });
        
        return frame;
      }
      
    } catch (_) {}
  }
  
  return null;
}

async function selectPlanoFinanceiro({
  page,
  values = [],
  reportName,
}) {
  // defensive normalize: always pass a clean array to selectViaModal
  values = uniqueNonEmpty(Array.isArray(values) ? values.map(x => String(x || '').trim()) : [String(values || '').trim()]);

  await selectViaModal({
    
    page,
    
    triggerSelector:
    'img[onclick*="planoFin"]',
    
    modalTitle:
    'Consulta de Plano Financeiro',
    
    searchInputSelector: [
      // cobre variações: busca por campos relacionados a 'plano' ou 'nmConta'
      'input[name="entity.nmConta"]',
      'input[name*="plano"]',
      'input[id*="plano"]',
      '#nmConta',
      'input[name*="nmConta"]',
    ],
    values,
    reportName,
    
  });
}

// =====================================================
// SELECIONA DOCUMENTOS
// =====================================================
async function selectDocumentos({
  page,
  values = [],
  enablePagination,
  reportName,
}) {
  // defensive normalize: always pass a clean array to selectViaModal
  values = uniqueNonEmpty(Array.isArray(values) ? values.map(x => String(x || '').trim()) : [String(values || '').trim()]);

  await selectViaModal({
    
    page,
    
    triggerSelector:
    '#docMultFilterContaRecebidas img',
    
    modalTitle:
    'Consulta de Documentos',
    
    searchInputSelector: [
      
      'input[name="entity.nmDocumento"]',
      'input[name="nmDocumento"]',
      
    ],
    values,
    enablePagination: true,
    reportName,
    
  });
}

async function findModalFrame(
  page,
  marker,
  timeout = 30000
) {
  
  const started =
  Date.now();
  
  const normalizedMarker = String(marker).trim();
  const markerRegex = new RegExp(escRe(normalizedMarker), 'i');
  
  while (
    Date.now() - started <
    timeout
  ) {
    
    for (const frame of page.frames()) {
      
      try {
        
        const url =
        frame.url() || '';
        
        // IGNORA FRAME VAZIO
        if (
          url === 'about:blank'
        ) {
          continue;
        }
        
        // 1) Busca o texto do título no conteúdo bruto
        try {
          const content =
          await frame.content();
          
          if (
            content &&
            content.match(markerRegex)
          ) {
            return frame;
          }
        } catch {}
        
        // 2) Busca texto visível no frame
        try {
          const visibleTitle = await frame.locator(`text=${normalizedMarker}`).count();
          if (visibleTitle && visibleTitle > 0) {
            return frame;
          }
        } catch {}
        
        // 3) Busca seletor conhecido de modal de seleção
        try {
          const modalTrigger = await frame.locator(
            'input[name="pbSelecionar"], #pbSelecionar, input[type="button"][value*="TODOS" i], input[type="submit"][value*="TODOS" i], button:has-text("TODOS")'
          ).count();
          
          if (modalTrigger && modalTrigger > 0) {
            return frame;
          }
        } catch {}
        
        // 4) Busca cabeçalhos semelhantes ao título do modal
        try {
          const headerMatch = await frame.locator(
            `h1:has-text("${normalizedMarker}"), h2:has-text("${normalizedMarker}"), span:has-text("${normalizedMarker}")`
          ).count();
          
          if (headerMatch && headerMatch > 0) {
            return frame;
          }
        } catch {}
        
      } catch {}
    }
    
    await page.waitForTimeout(
      500
    );
  }
  
  return null;
}

async function safeFrameClick({
  
  page,
  
  frameMarker,
  
  selector,
  
  timeout = 45000,
  
}) {
  
  const started =
  Date.now();
  
  let lastError = null;
  
  while (
    Date.now() - started <
    timeout
  ) {
    
    try {
      
      // ============================================
      // RELOCALIZA FRAME
      // ============================================
      
      let targetFrame = null;
      
      for (const frame of page.frames()) {
        
        try {
          
          const url =
          frame.url() || '';
          
          if (
            url === 'about:blank'
          ) {
            continue;
          }
          
          const content =
          await frame.content();
          
          if (
            content.includes(
              frameMarker
            )
          ) {
            
            targetFrame = frame;
            
            break;
          }
          
        } catch {}
      }
      
      if (!targetFrame) {
        
        await page.waitForTimeout(
          1000
        );
        
        continue;
      }
      
      // ============================================
      // RELOCALIZA ELEMENTO
      // ============================================
      
      const btn =
      targetFrame.locator(
        selector
      ).first();
      
      await btn.waitFor({
        state: 'attached',
        timeout: 5000,
      });
      
      // ============================================
      // CLICK PLAYWRIGHT
      // MAIS RESILIENTE
      // ============================================
      
      await btn.click({
        force: true,
        timeout: 5000,
      });
      
      return true;
      
    } catch (err) {
      
      lastError = err;
      
      const msg =
      String(err);
      
      // ============================================
      // SIENGE RECRIOU FRAME
      // ============================================
      
      if (
        msg.includes(
          'Frame was detached'
        ) ||
        msg.includes(
          'Execution context was destroyed'
        ) ||
        msg.includes(
          'Target closed'
        )
      ) {
        
        await page.waitForTimeout(
          1500
        );
        
        continue;
      }
      
      // ============================================
      // ELEMENTO RECRIADO
      // ============================================
      
      if (
        msg.includes(
          'Element is not attached'
        )
      ) {
        
        await page.waitForTimeout(
          1000
        );
        
        continue;
      }
      
      await page.waitForTimeout(
        1000
      );
    }
  }
  
  throw new Error(
    `Falha click após múltiplas tentativas.\n${lastError}`
  );
}

async function findVisibleInput(surface, selectors) {
  
  const selectorList =
  Array.isArray(selectors)
  ? selectors
  : [selectors];
  
  const attachedCandidates = [];

  for (const selector of selectorList) {
    
    try {
      
      const locator =
      surface
      .locator(selector)
      .first();
      
      const count =
      await locator
      .count()
      .catch(() => 0);
      
      if (!count) {
        continue;
      }
      
      attachedCandidates.push(locator);
      
      await locator.waitFor({
        state: 'visible',
        timeout: 3000,
      }).catch(() => {});
      
      const visible =
      await locator
      .isVisible()
      .catch(() => false);
      
      if (visible) {
        
        return locator;
      }
      
    } catch (_) {}
  }

  const fallbackSelectors = [
    'input[name*="documento"]',
    'input[id*="documento"]',
    'input[placeholder*="Documento"]',
    'input[placeholder*="documento"]',
    'input[aria-label*="Documento"]',
    'input[aria-label*="documento"]',
    'input[name*="nmDocumento"]',
    'input[id*="nmDocumento"]',
    '#nmDocumento',
    'input[type="search"]',
  ];

  for (const selector of fallbackSelectors) {
    
    try {
      
      const locator =
      surface
      .locator(selector)
      .first();
      
      const count =
      await locator
      .count()
      .catch(() => 0);
      
      if (!count) {
        continue;
      }
      
      attachedCandidates.push(locator);
      
      await locator.waitFor({
        state: 'visible',
        timeout: 3000,
      }).catch(() => {});
      
      const visible =
      await locator
      .isVisible()
      .catch(() => false);
      
      if (visible) {
        
        return locator;
      }
      
    } catch (_) {}
  }

  try {
    const candidates =
    surface
    .locator('input[type="text"], input:not([type]), textarea');
    const count =
    await candidates
    .count()
    .catch(() => 0);

    for (let i = 0; i < count; i++) {
      const candidate = candidates.nth(i);
      const visible =
      await candidate
      .isVisible()
      .catch(() => false);
      
      if (visible) {
        return candidate;
      }

      attachedCandidates.push(candidate);
    }
  } catch (_) {}

  if (attachedCandidates.length) {
    return attachedCandidates[0];
  }

  return null;
}

async function resolveInput(surface, selectors, timeout = 30000) {
  const selectorList = uniqueNonEmpty(Array.isArray(selectors) ? selectors : [selectors]);
  const fallbackSelectors = [
    'input[name*="documento"]',
    'input[id*="documento"]',
    'input[placeholder*="Documento"]',
    'input[placeholder*="documento"]',
    'input[aria-label*="Documento"]',
    'input[aria-label*="documento"]',
    'input[name*="nmDocumento"]',
    'input[id*="nmDocumento"]',
    '#nmDocumento',
    'input[type="search"]',
    'input[type="text"]',
    'input:not([type])',
    'textarea',
  ];
  const candidates = uniqueNonEmpty([...selectorList, ...fallbackSelectors]);

  const started = Date.now();
  let attachedFirst = null;

  while (Date.now() - started < timeout) {
    for (const selector of candidates) {
      try {
        const locator = surface.locator(selector).first();
        const count = await locator.count().catch(() => 0);
        if (!count) continue;
        if (!attachedFirst) attachedFirst = locator;

        await locator.waitFor({ state: 'visible', timeout: 1000 }).catch(() => {});
        const visible = await locator.isVisible().catch(() => false);
        if (visible) return locator;
      } catch (_) {}
    }
    await new Promise(resolve => setTimeout(resolve, 250));
  }

  return attachedFirst;
}

// =====================================================
// HELPERS MODAL
// =====================================================

async function resolveModalLocator(
  page,
  selector,
  timeout = 30000
) {
  
  const started =
  Date.now();
  
  while (
    Date.now() - started <
    timeout
  ) {
    
    for (const frame of page.frames()) {
      
      try {
        
        const locator =
        frame
        .locator(selector)
        .first();
        
        const visible =
        await locator
        .isVisible()
        .catch(() => false);
        
        if (visible) {
          
          return {
            frame,
            locator,
          };
        }
        
      } catch (_) {}
    }
    
    await page.waitForTimeout(
      500
    );
  }
  
  throw new Error(
    `Elemento não encontrado no modal: ${selector}`
  );
  
}

async function findRowByText(
  page,
  text,
  timeout = 30000,
  preferredFrame = null
) {
  
  const started =
  Date.now();
  
  let frames = page.frames();
  
  // Se temos um frame preferido, o colocamos primeiro na lista
  if (preferredFrame) {
    frames = [
      preferredFrame,
      ...frames.filter(f => f !== preferredFrame)
    ];
  }
  
  function normalizeText(s) {
    try {
      return String(s || '')
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '')
      .replace(/[^\w\s]/g, ' ')
      .replace(/\s+/g, ' ')
      .trim()
      .toLowerCase();
    } catch (_) { return String(s || '').replace(/\s+/g, ' ').trim(); }
  }
  
  const normalizedText = normalizeText(text);
  
  while (
    Date.now() - started <
    timeout
  ) {
    
    for (const frame of frames) {
      try {
        // procura em uma gama maior de elementos (td, th, tr, span, div, a)
        const selectors = ['td', 'th', 'tr', 'span', 'div', 'a', 'label'];
        for (const sel of selectors) {
          const elems = frame.locator(sel);
          const count = await elems.count().catch(() => 0);
          for (let i = 0; i < count; i++) {
            const el = elems.nth(i);
            const content = (await el.innerText().catch(() => '')).trim();
            const normalizedContent = normalizeText(content);
            
            if (!normalizedContent) continue;
            
            // exact (case-sensitive after normalization)
            if (normalizedContent === normalizedText) {
              try { await el.waitFor({ state: 'visible', timeout: 3000 }).catch(() => {}); } catch (_) {}
              logEvent({ level: 'debug', message: `Linha encontrada (EXATO) em ${sel}: ${text}`, frameUrl: frame.url(), index: i });
              return { frame, cell: el };
            }
            
          }
        }
      } catch (_) {}
      
      // Fallback: busca por <tr> via script no contexto do frame, ignorando acentos
      try {
        const idx = await frame.evaluate((search) => {
          function normalize(s) {
            try {
              return String(s || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/[^\w\s]/g, ' ').replace(/\s+/g, ' ').trim().toLowerCase();
            } catch (_) { return String(s || '').replace(/\s+/g, ' ').trim().toLowerCase(); }
          }
          const rows = Array.from(document.querySelectorAll('tr'));
          for (let i = 0; i < rows.length; i++) {
            if (normalize(rows[i].innerText).trim() === search) return i + 1;
          }
          return null;
        }, normalizedText).catch(() => null);
        
        if (idx) {
          logEvent({ level: 'debug', message: `Linha encontrada via fallback <tr> index ${idx}: ${text}`, frameUrl: frame.url() });
          return { frame, cell: frame.locator(`xpath=(//tr)[${idx}]`) };
        }
      } catch (_) {}
    }
    
    await new Promise(resolve => setTimeout(resolve, 500));
  }
  
  throw new Error(
    `Linha não encontrada no modal: ${text}`
  );
}

async function findVisibleLocatorInFrames(
  page,
  selector,
  timeout = 30000
) {
  
  const started =
  Date.now();
  
  while (
    Date.now() - started <
    timeout
  ) {
    
    for (const frame of page.frames()) {
      
      try {
        
        const locator =
        frame
        .locator(selector)
        .first();
        
        const visible =
        await locator
        .isVisible()
        .catch(() => false);
        
        if (visible) {
          
          return {
            frame,
            locator,
          };
        }
        
      } catch (_) {}
    }
    
    await page.waitForTimeout(
      500
    );
  }
  
  throw new Error(
    `Elemento não encontrado: ${selector}`
  );
}

async function findSelectionModalFrame(
  page,
  modalTitle,
  timeout = 30000
) {
  
  const definition =
  getModalSelectionDefinition(modalTitle);
  
  const selectors =
  uniqueNonEmpty([
    ...definition.frameSelectors,
    ...definition.searchInputSelectors,
  ]);
  
  const started =
  Date.now();
  
  while (
    Date.now() - started <
    timeout
  ) {
    
    for (const frame of page.frames()) {
      
      try {
        
        const hasSelecionar =
        await frame.locator(
          'input[name="pbSelecionar"], button[name="pbSelecionar"], button#pbSelecionar, #pbSelecionar'
        ).count().catch(() => 0);
        
        if (!hasSelecionar) {
          continue;
        }
        
        let matches = false;
        
        for (const selector of selectors) {
          const count =
          await frame.locator(selector).count().catch(() => 0);
          
          if (count > 0) {
            matches = true;
            break;
          }
        }
        
        if (matches) {
          return frame;
        }
        
      } catch (_) {}
    }
    
    await page.waitForTimeout(500);
  }
  
  return null;
}

async function findSelectionRow(
  modalFrame,
  value,
  selectors = [],
  timeout = 15000
) {
  
  const started =
  Date.now();
  
  const normalizedValue =
  normalizeSelectionText(value);
  
  const rowSelectors = [
    '#tabelaResultado tr[indice]',
    '#tabelaResultado tr[linha="true"]',
    'tr[indice]',
    'tr[linha="true"]',
  ];
  
  while (
    Date.now() - started <
    timeout
  ) {
    
    let rows = null;
    
    for (const selector of rowSelectors) {
      try {
        const locator = modalFrame.locator(selector);
        const count = await locator.count().catch(() => 0);
        if (count > 0) {
          rows = locator;
          break;
        }
      } catch (_) {}
    }
    
    if (!rows) {
      await modalFrame.page().waitForTimeout(500);
      continue;
    }
    
    const count =
    await rows.count().catch(() => 0);
    
    for (let i = 0; i < count; i++) {
      const row = rows.nth(i);
      
      try {
        const rowText =
        normalizeSelectionText(
          await row.innerText().catch(() => '')
        );
        // 1) exact equality
        if (rowText === normalizedValue) return row;

        // 2) exact match on any configured cell selectors
        for (const selector of selectors) {
          const cell = row.locator(selector).first();
          const cellCount = await cell.count().catch(() => 0);
          if (!cellCount) continue;
          const cellText = normalizeSelectionText(await cell.innerText().catch(() => ''));
          if (cellText === normalizedValue) return row;
        }

        // only exact equality is accepted for this modal
      } catch (_) {}
    }
    
    await modalFrame.page().waitForTimeout(500);
  }
  try {
    const preview = await modalFrame.evaluate(() => {
      try {
        return Array.from(document.querySelectorAll('tr')).slice(0,20).map(r => r.innerText || '').slice(0,10);
      } catch (e) { return []; }
    }).catch(() => []);
    logEvent({ level: 'debug', message: `findSelectionRow final preview (first rows): ${preview.map(p => p.slice(0,200)).join(' || ')}`, frameUrl: modalFrame.url() });
  } catch (_) {}

  // Final fallback: try a DOM-level scan of all <tr> for exact normalized text.
  try {
    const normalized = normalizedValue;
    const idx = await modalFrame.evaluate((search) => {
      function normalize(s) {
        try { return String(s || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/[^\w\s]/g, ' ').replace(/\s+/g, ' ').trim().toLowerCase(); } catch (e) { return String(s || '').replace(/[^\w\s]/g, ' ').replace(/\s+/g, ' ').trim().toLowerCase(); }
      }
      const rows = Array.from(document.querySelectorAll('tr'));
      for (let i = 0; i < rows.length; i++) {
        try {
          const txt = normalize(rows[i].innerText || '');
          if (!txt) continue;
          if (txt === search) return i + 1;
        } catch (e) {}
      }
      return null;
    }, normalized).catch(() => null);

    if (idx) {
      logEvent({ level: 'debug', message: `findSelectionRow fallback DOM matched tr index=${idx}`, frameUrl: modalFrame.url() });
      return modalFrame.locator(`xpath=(//tr)[${idx}]`);
    }
  } catch (_) {}

  throw new Error(
    `Linha não encontrada no modal: ${value}`
  );
}

async function ensureCheckboxChecked(
  rowLocator,
  checkboxSelector = 'input[type="checkbox"][name="rowSelect"]'
) {
  const selector = checkboxSelector || 'input[type="checkbox"]';

  // 1) DOM-first: try to click matching checkboxes directly via evaluate inside the row
  try {
    const domResult = await rowLocator.evaluate((r, sel) => {
      try {
        const nodes = Array.from((sel ? r.querySelectorAll(sel) : r.querySelectorAll('input[type="checkbox"]')) || []);
        if (!nodes.length) return false;
        let anyChecked = false;
        for (const cb of nodes) {
          try {
            if (!cb.checked) {
              try { cb.click(); } catch (e) { cb.checked = true; }
              try { cb.dispatchEvent(new Event('change', { bubbles: true })); } catch (_) {}
            }
            if (cb.checked) anyChecked = true;
          } catch (e) {}
        }
        return anyChecked;
      } catch (e) { return false; }
    }, selector).catch(() => false);
    logEvent({ level: 'debug', message: `ensureCheckboxChecked DOM-first result=${domResult}`, selector });
    if (domResult) return true;
  } catch (e) {
    logEvent({ level: 'debug', message: `ensureCheckboxChecked DOM-first error: ${String(e && e.message || e)}` });
  }

  // 2) Fallback: use Playwright locators with a few attempts (preserve original robust behavior)
  const attempts = 4;
  for (let i = 0; i < attempts; i++) {
    try {
      const snippet = (await rowLocator.innerText().catch(() => '')).slice(0, 200);
      logEvent({ level: 'debug', message: `ensureCheckboxChecked attempt=${i + 1}/${attempts} row-snippet='${snippet}' selector='${selector}'` });

      const list = rowLocator.locator(selector);
      const count = await list.count().catch(() => 0);
      logEvent({ level: 'debug', message: `ensureCheckboxChecked: found ${count} checkbox(es) using selector='${selector}'` });

      if (count === 0) {
        const any = rowLocator.locator('input[type="checkbox"]');
        const anyCount = await any.count().catch(() => 0);
        logEvent({ level: 'debug', message: `ensureCheckboxChecked: any checkbox inside row count=${anyCount}` });
        for (let j = 0; j < anyCount; j++) {
          const cb = any.nth(j);
          try {
            const checked = await cb.isChecked().catch(() => false);
            if (!checked) {
              try { await cb.check(); } catch (err) { try { await cb.click({ force: true }); } catch (e) { await cb.evaluate(el => el.click()).catch(() => {}); } }
            }
          } catch (e) { logEvent({ level: 'debug', message: `ensureCheckboxChecked any-fallback error: ${String(e && e.message || e)}` }); }
        }
      } else {
        for (let idx = 0; idx < count; idx++) {
          const cb = list.nth(idx);
          try {
            const checked = await cb.isChecked().catch(() => false);
            if (!checked) {
              try { await cb.check(); } catch (err) {
                logEvent({ level: 'debug', message: `checkbox.check() failed: ${String(err && err.message || err)}` });
                try { await cb.click({ force: true }); } catch (e) {
                  logEvent({ level: 'debug', message: `checkbox.click() failed: ${String(e && e.message || e)}` });
                  try { await rowLocator.evaluate((r, params) => { const sel = params.sel; const index = params.index; const nodes = Array.from(r.querySelectorAll(sel)); if (nodes[index] && !nodes[index].checked) { nodes[index].click(); nodes[index].checked = true; nodes[index].dispatchEvent(new Event('change', { bubbles: true })); } }, { sel: selector, index: idx }).catch(() => {}); } catch (_) { logEvent({ level: 'debug', message: 'evaluate click fallback failed' }); }
                }
              }
            }
          } catch (e) { logEvent({ level: 'debug', message: `ensureCheckboxChecked per-cb error: ${String(e && e.message || e)}` }); }
        }
      }

      // verify at least one checkbox is checked now
      const finalList = rowLocator.locator(selector);
      const finalCount = await finalList.count().catch(() => 0);
      let anyChecked = false;
      for (let k = 0; k < finalCount; k++) {
        try { if (await finalList.nth(k).isChecked().catch(() => false)) { anyChecked = true; break; } } catch (_) {}
      }
      logEvent({ level: 'debug', message: `ensureCheckboxChecked: anyChecked=${anyChecked}` });
      if (anyChecked) return true;
    } catch (e) {
      logEvent({ level: 'debug', message: `ensureCheckboxChecked: attempt error: ${String(e && e.message || e)}` });
    }

    try { await rowLocator.page().waitForTimeout(300 + i * 200).catch(() => {}); } catch (_) {}
    try { await rowLocator.click({ force: true }).catch(() => {}); } catch (_) {}
  }

  // final attempt: DOM evaluate any checkbox inside the row
  try {
    await rowLocator.evaluate((r) => {
      try {
        const cb = r.querySelector('input[type="checkbox"]');
        if (cb && !cb.checked) { try { cb.click(); } catch (_) { cb.checked = true; } cb.dispatchEvent(new Event('change', { bubbles: true })); }
      } catch (e) {}
    }).catch(() => {});
  } catch (_) {}

  try {
    const final = rowLocator.locator(checkboxSelector).first();
    const nowChecked = await final.isChecked().catch(() => false);
    return !!nowChecked;
  } catch (_) {
    return false;
  }
}

async function selectViaModal({
  
  page,
  
  triggerSelector,
  
  modalTitle,
  
  searchInputSelector,
  
  value,
  values,

  enablePagination = false,
  
  // optional report name to update progress
  reportName,
  
}) {
  
  // =====================================================
  // NORMALIZA ENTRADAS
  // =====================================================
  
  const entries =
  uniqueNonEmpty(
    values || [value]
  );
  
  if (!entries.length) {
    return;
  }
  
  const modalDefinition =
  getModalSelectionDefinition(modalTitle);
  
  const searchSelectors =
  uniqueNonEmpty([
    ...(Array.isArray(searchInputSelector)
    ? searchInputSelector
    : [searchInputSelector]),
    ...modalDefinition.searchInputSelectors,
  ]);
  
  logEvent({
    level: 'info',
    message:
    `Selecionando valores: ${entries.join(', ')}`,
    modalTitle,
    count: entries.length,
  });
  
  // =====================================================
  // IFRAME PRINCIPAL
  // =====================================================
  
  const iframeLocator =
  page.locator(
    'iframe#iFramePage'
  );
  
  await iframeLocator.waitFor({
    state: 'attached',
    timeout: 30000,
  });
  
  const mainFrame =
  await iframeLocator.contentFrame();
  
  if (!mainFrame) {
    
    throw new Error(
      'Não foi possível obter contentFrame do iFramePage.'
    );
  }
  
  // Garante que os filtros avançados estejam expandidos no frame principal
  try {
    await ensureExpandedFilters(mainFrame);
  } catch (e) {
    // não é fatal — apenas log
    logEvent({ level: 'debug', message: `Falha ao expandir filtros no mainFrame: ${String(e && e.message || e)}` });
  }
  
  async function fillLegacyInput(
    locator,
    nextValue
  ) {
    
    // Tenta tornar visível e scrollar para a vista
    try {
      await locator.scrollIntoViewIfNeeded();
    } catch (_) {}
    
    await locator.evaluate(
      (el, value) => {
        
        // Tenta remover estilos que possam estar ocultando
        el.style.display = '';
        el.style.visibility = 'visible';
        el.style.opacity = '1';
        
        el.focus();
        
        el.value = '';
        
        el.dispatchEvent(
          new Event('input', {
            bubbles: true
          })
        );
        
        el.value = value;
        
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
          new KeyboardEvent(
            'keyup',
            {
              bubbles: true
            }
          )
        );
        
        el.dispatchEvent(
          new Event('blur', {
            bubbles: true
          })
        );
        
      },
      String(nextValue)
    );
  }
  
  const trigger =
  mainFrame.locator(
    `${triggerSelector}[src*="botProcurar.png"]`
  ).first();
  
  // =====================================================
  // MITIGAÇÃO DEFENSIVA: limpar estado de modal entre relatórios
  // - reseta cache em memória e em disco
  // - limpa listas e contadores temporários que podem vazar entre modais
  // =====================================================
  try {
    MODAL_CACHE = {};
    try { await saveModalCache(); } catch (_) {}
  } catch (_) {}

  try {
    await mainFrame.evaluate(() => {
      try {
        // limpa elementos que armazenam listas selecionadas
        Array.from(document.querySelectorAll('[id$="SelectedEntitiesList"], [class*="SelectedEntitiesList"]')).forEach(el => {
          if (el.tagName === 'INPUT' || el.tagName === 'TEXTAREA') el.value = '';
          else el.innerHTML = '';
        });

        // zera contadores auxiliares que o modal pode usar
        Array.from(document.querySelectorAll('[id*="contador"], [name*="contador"]')).forEach(el => {
          if (el.tagName === 'INPUT' || el.tagName === 'TEXTAREA') el.value = '0';
          else el.textContent = '0';
        });
      } catch (e) {
        // não fatal
      }
    }).catch(() => {});
  } catch (_) {}
  
  await trigger.click({
    force: true,
  });
  
  let modalFrame =
  await findSelectionModalFrame(
    page,
    modalTitle,
    30000
  );
  
  if (!modalFrame) {
    modalFrame =
    await findModalFrame(
      page,
      modalTitle,
      30000
    );
  }

  
  
  if (!modalFrame) {
    
    throw new Error(
      `Modal não encontrado: ${modalTitle}`
    );
  }
  
  logEvent({
    level: 'info',
    message: `Modal encontrado: ${modalTitle}`,
    frameUrl: modalFrame.url()
  });
  
  try { await setReportProgress(reportName, 8, { lastMessage: `modal_open:${modalTitle}`, step: 'modal_open' }); } catch (_) {}
  
  // helper: salva estado debug (HTML do modal, preview e screenshot)
  async function saveDebugState(tag) {
    if (!DEBUG_HTML) return null;
    try {
      await fs.promises.mkdir(SCREENSHOT_DIR, { recursive: true });
    } catch (_) {}
    const ts = nowIso().replace(/[:.]/g, '-');
    const base = sanitizeFileName(`${modalTitle}-${tag}-${ts}`);
    const htmlPath = path.join(SCREENSHOT_DIR, `${base}.html`);
    const previewPath = path.join(SCREENSHOT_DIR, `${base}-preview.json`);
    const shotPath = path.join(SCREENSHOT_DIR, `${base}.png`);
    try {
      const modalHtml = await modalFrame.content().catch(() => '<!-- erro ao capturar -->');
      await fs.promises.writeFile(htmlPath, modalHtml || '<!-- vazio -->').catch(() => {});
    } catch (e) {
      logEvent({ level: 'warn', message: `Erro ao salvar HTML debug: ${e.message}` });
    }
    try {
      const preview = await modalFrame.evaluate(() => Array.from(document.querySelectorAll('tr')).slice(0,30).map(r => r.innerText)).catch(() => []);
      await fs.promises.writeFile(previewPath, JSON.stringify(preview, null, 2)).catch(() => {});
    } catch (e) {
      logEvent({ level: 'warn', message: `Erro ao salvar preview debug: ${e.message}` });
    }
    try {
      await page.screenshot({ path: shotPath, fullPage: true }).catch(() => {});
    } catch (e) {
      logEvent({ level: 'warn', message: `Erro ao salvar screenshot debug: ${e.message}` });
    }
    logEvent({ level: 'info', message: `Debug salvo: html=${htmlPath}, preview=${previewPath}, shot=${shotPath}` });
    return { htmlPath, previewPath, shotPath };
  }
  
  // Aguarda o conteúdo do modal estar pronto
  for (const selector of searchSelectors) {
    try {
      await withFrameRetry(async () => {
        await modalFrame.locator(selector).first().waitFor({ state: 'attached', timeout: 5000 }).catch(() => {});
      }, { name: `waitFor:${selector}` }).catch(() => {});
    } catch (_) {}
  }
  
  // =====================================================
  // MARCA TODOS OS VALORES
  // =====================================================
  
  // prepara entradas normalizadas para matching DOM-first
  const normalizedEntries = entries.map(e => normalizeSelectionText(e)).filter(Boolean);

  function rowMatchesNormalizedNeedle(row, needle) {
    try {
      const parts = [];
      if (Array.isArray(modalDefinition.rowValueSelectors) && modalDefinition.rowValueSelectors.length) {
        for (const selector of modalDefinition.rowValueSelectors) {
          try {
            const cell = row.querySelector(selector);
            if (cell) parts.push(normalizeSelectionText(cell.innerText || ''));
          } catch (_) {}
        }
      }

      if (!parts.length) {
        const raw = String(row.innerText || '');
        parts.push(...raw.split(/\r?\n|\t/).map(text => normalizeSelectionText(text)).filter(Boolean));
      }

      return parts.some(part => part === needle);
    } catch (_) {
      return false;
    }
  }

  const modalInput =
  await resolveInput(
    modalFrame,
    searchSelectors,
    4
  ).catch(() => null);
  
  if (!modalInput) {
    await saveDebugState('input-not-found').catch(() => {});
    logEvent({ level: 'info', message: 'Input do modal nao encontrado — prosseguindo com seleção por DOM (sem buscas)', modalTitle });
    // nao interrompe o fluxo: tentaremos selecionar tudo diretamente no DOM
  }
  
  // TENTA SELEÇÃO EM LOTE (mais rápida) — seleciona checkboxes no DOM de uma só vez
  try {
    logEvent({ level: 'debug', message: 'Bulk selection: normalized entries', detail: JSON.stringify(normalizedEntries.slice(0,50)), modalTitle });
    if (normalizedEntries.length) {
      // Primeira tentativa: acionar checkbox de header (select-all) se existir
      try {
        await withFrameRetry(async () => {
          await modalFrame.evaluate((rowCheckboxSel) => {
            try {
              const header = document.querySelector('thead input[type="checkbox"], tr th input[type="checkbox"], .ui-table thead input[type="checkbox"], input[data-select-all], [data-select-all]');
              if (header && !header.checked) {
                try { header.click(); } catch (e) { header.checked = true; header.dispatchEvent(new Event('change', { bubbles: true })); }
              }
              return true;
            } catch (e) { return false; }
          }, modalDefinition.rowCheckboxSelector).catch(() => {});
        }, { name: 'dom-select-all-attempt' }).catch(() => {});
      } catch (_) {}

      // scroll modal to render virtualized rows before matching
      await withFrameRetry(async () => {
        await modalFrame.evaluate(() => {
          const scroller = document.querySelector('.ui-table') || document.querySelector('.modal-body') || document.scrollingElement || document.documentElement;
          if (!scroller) return;
          const height = scroller.scrollHeight || (document.body && document.body.scrollHeight) || 0;
          let y = 0;
          while (y < height) {
            scroller.scrollTop = y;
            y += 800;
          }
          scroller.scrollTop = height;
        }).catch(() => {});
      }, { name: 'bulk-scroll' }).catch(() => {});
      
        const isDocsModal = normalizeSelectionText(modalTitle || '').includes('consulta de documentos');
        let matched = null;
        if (isDocsModal) {
          try {
            const maxPages = 100;
            const accum = new Set();
            const accumDocumentos = new Set();
            let rowsCount = 0;
            let sampleRows = [];
            for (let pageIdx = 0; pageIdx < maxPages; pageIdx++) {
              const res = await withFrameRetry(async () => await modalFrame.evaluate(({ normalizedEntries, rowCheckboxSel, rowValueSelectors, nextSelectors }) => {
                function normalize(s) {
                  try { return (s || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/[^\w\s]/g, ' ').replace(/\s+/g, ' ').trim().toLowerCase(); } catch (e) { return (s || '').replace(/[^\w\s]/g, ' ').replace(/\s+/g, ' ').trim().toLowerCase(); }
                }
                const candidateRowSelectors = ['#tabelaResultado tr[indice]', '#tabelaResultado tr[linha="true"]', 'tr[indice]', 'tr[linha="true"]', 'tbody > tr', 'tr'];
                let rows = [];
                for (const sel of candidateRowSelectors) {
                  try { const r = Array.from(document.querySelectorAll(sel) || []); if (r && r.length) { rows = r; break; } } catch (e) {}
                }
                if (!rows || !rows.length) rows = Array.from(document.querySelectorAll('tr'));

                const matchedPage = new Set();
                const selectionFailures = [];
                const documentos = Array.from(document.querySelectorAll('tr[id^="linha_"]')).map(tr => tr.querySelectorAll('td')[2]?.innerText.trim()).filter(Boolean);
                for (const row of rows) {
                  try {
                    const parts = [];
                    if (Array.isArray(rowValueSelectors) && rowValueSelectors.length) {
                      for (const rsel of rowValueSelectors) {
                        try { const c = row.querySelector(rsel); if (c) parts.push(normalize(c.innerText || '')); } catch (e) {}
                      }
                    }
                    if (!parts.length) {
                      const raw = String(row.innerText || '');
                      parts.push(...raw.split(/\r?\n|\t/).map(s => normalize(s)).filter(Boolean));
                    }
                    for (let i = 0; i < normalizedEntries.length; i++) {
                      const ne = normalizedEntries[i];
                      if (!ne || matchedPage.has(ne)) continue;
                      const ok = parts.some(p => p === ne);
                      if (ok) {
                        const checkbox = row.querySelector(rowCheckboxSel || 'input[type="checkbox"]');
                        if (checkbox && !checkbox.checked) {
                          try { checkbox.click(); } catch (e) { try { checkbox.checked = true; checkbox.dispatchEvent(new Event('change', { bubbles: true })); } catch (_) {} }
                          try { checkbox.dispatchEvent(new Event('change', { bubbles: true })); } catch (_) {}
                        }
                        if (!(checkbox && checkbox.checked)) {
                          selectionFailures.push({
                            needle: ne,
                            rowIndex: rows.indexOf(row) + 1,
                            rowText: String(row.innerText || '').trim(),
                            checkboxSelector: rowCheckboxSel || 'input[type="checkbox"]',
                          });
                          continue;
                        }
                        matchedPage.add(ne);
                      }
                    }
                  } catch (e) { /* ignore row errors */ }
                }

                let nextClicked = false;
                for (const ns of nextSelectors) {
                  try {
                    const el = document.querySelector(ns);
                    if (el) {
                      const disabled = el.disabled || (el.classList && el.classList.contains && el.classList.contains('ui-state-disabled')) || el.getAttribute('aria-disabled') === 'true';
                      if (!disabled) {
                        try { el.click(); } catch (e) { }
                        nextClicked = true;
                        break;
                      }
                    }
                  } catch (e) {}
                }
                return { matchedPage: Array.from(matchedPage), rowsCount: rows.length, sampleRows: rows.slice(0,30).map(r => r.innerText), nextClicked, documentos, selectionFailures };
              }, { normalizedEntries, rowCheckboxSel: modalDefinition.rowCheckboxSelector, rowValueSelectors: (modalDefinition && modalDefinition.rowValueSelectors) || [], nextSelectors: DOCUMENTOS_NEXT_PAGE_SELECTORS }), { name: 'bulk-evaluate-page' });

              if (!res) break;
              rowsCount = rowsCount || res.rowsCount || 0;
              sampleRows = sampleRows.length ? sampleRows : (res.sampleRows || []);
              (res.matchedPage || []).forEach(m => accum.add(m));
              (res.documentos || []).forEach(d => accumDocumentos.add(d));
              if (Array.isArray(res.selectionFailures) && res.selectionFailures.length) {
                logEvent({ level: 'warn', message: 'Bulk selection page had checkbox failures', modalTitle, detail: JSON.stringify(res.selectionFailures.slice(0, 20)) });
              }
              if (accum.size >= normalizedEntries.length) {
                matched = { found: accum.size, expected: normalizedEntries.length, matched: Array.from(accum), rowsCount, sampleRows, documentos: Array.from(accumDocumentos) };
                break;
              }
              if (!res.nextClicked) break;
              try { await page.waitForTimeout(300); } catch (e) {}
            }
            if (!matched) matched = { found: accum.size, expected: normalizedEntries.length, matched: Array.from(accum), rowsCount, sampleRows, documentos: Array.from(accumDocumentos) };
          } catch (err) {
            matched = null;
          }
        } else {
          // Non-paginated: single-page evaluate (original behavior)
          try {
            matched = await withFrameRetry(async () => await modalFrame.evaluate(({ normalizedEntries, rowCheckboxSel, rowValueSelectors }) => {
              function normalize(s) {
                try { return (s || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/[^\w\s]/g, ' ').replace(/\s+/g, ' ').trim().toLowerCase(); } catch (e) { return (s || '').replace(/[^\w\s]/g, ' ').replace(/\s+/g, ' ').trim().toLowerCase(); }
              }
              const candidateRowSelectors = ['#tabelaResultado tr[indice]', '#tabelaResultado tr[linha="true"]', 'tr[indice]', 'tr[linha="true"]', 'tbody > tr', 'tr'];
              let rows = [];
              for (const sel of candidateRowSelectors) {
                try { const r = Array.from(document.querySelectorAll(sel) || []); if (r && r.length) { rows = r; break; } } catch (e) {}
              }
              if (!rows || !rows.length) rows = Array.from(document.querySelectorAll('tr'));
              const seen = new Set();
              const selectionFailures = [];
              for (const row of rows) {
                try {
                  const parts = [];
                  if (Array.isArray(rowValueSelectors) && rowValueSelectors.length) {
                    for (const rsel of rowValueSelectors) {
                      try { const c = row.querySelector(rsel); if (c) parts.push(normalize(c.innerText || '')); } catch (e) {}
                    }
                  }
                  if (!parts.length) {
                    const raw = String(row.innerText || '');
                    parts.push(...raw.split(/\r?\n|\t/).map(s => normalize(s)).filter(Boolean));
                  }
                  for (let i = 0; i < normalizedEntries.length; i++) {
                    const ne = normalizedEntries[i];
                    if (!ne || seen.has(ne)) continue;
                    const ok = parts.some(p => p === ne);
                    if (ok) {
                      const checkbox = row.querySelector(rowCheckboxSel || 'input[type="checkbox"]');
                      if (checkbox && !checkbox.checked) {
                        try { checkbox.click(); } catch (e) { try { checkbox.checked = true; checkbox.dispatchEvent(new Event('change', { bubbles: true })); } catch (_) {} }
                        try { checkbox.dispatchEvent(new Event('change', { bubbles: true })); } catch (_) {}
                      }
                      if (!(checkbox && checkbox.checked)) {
                        selectionFailures.push({
                          needle: ne,
                          rowIndex: rows.indexOf(row) + 1,
                          rowText: String(row.innerText || '').trim(),
                          checkboxSelector: rowCheckboxSel || 'input[type="checkbox"]',
                        });
                        continue;
                      }
                      seen.add(ne);
                    }
                  }
                } catch (e) { }
              }
              return { found: seen.size, expected: normalizedEntries.length, matched: Array.from(seen), rowsCount: rows.length, sampleRows: rows.slice(0,30).map(r => r.innerText), selectionFailures };
            }, { normalizedEntries, rowCheckboxSel: modalDefinition.rowCheckboxSelector, rowValueSelectors: (modalDefinition && modalDefinition.rowValueSelectors) || [] }), { name: 'bulk-evaluate' }).catch(() => null);
          } catch (err) {
            matched = null;
          }
        }
      
      if (matched) {
        logEvent({ level: 'debug', message: 'Bulk modal evaluate result', detail: JSON.stringify(matched), modalTitle });
        if (Array.isArray(matched.selectionFailures) && matched.selectionFailures.length) {
          logEvent({ level: 'warn', message: 'Bulk modal selection has unchecked matches', modalTitle, detail: JSON.stringify(matched.selectionFailures.slice(0, 20)) });
        }
      }

      if (matched && matched.found >= matched.expected) {
        logEvent({ level: 'info', message: `Bulk modal selection succeeded: ${matched.matched.join(', ')}`, modalTitle });
        
        try { await setReportProgress(reportName, 20, { lastMessage: 'bulk_selected', step: 'modal_bulk' }); } catch (_) {}
        
        // Antes de finalizar, captura lista de rows marcadas no modal para diagnóstico
        try {
          const modalChecked = await modalFrame.evaluate(() => Array.from(document.querySelectorAll('tr')).filter(r => { try { const cb = r.querySelector('input[type="checkbox"]'); return !!(cb && cb.checked); } catch(e){return false;} }).map(r => r.innerText.trim()).slice(0,200)).catch(() => []);
          logEvent({ level: 'info', message: `Modal checked rows (pre-select): ${modalChecked.length}`, detail: JSON.stringify(modalChecked.slice(0,50)), modalTitle });
          try {
            const normalize = value => normalizeSelectionText(value);
            const checkedNormalized = uniqueNonEmpty((Array.isArray(modalChecked) ? modalChecked : []).map(normalize));
            const missingChecked = normalizedEntries.filter(entry => !checkedNormalized.includes(entry));
            const extraChecked = checkedNormalized.filter(entry => !normalizedEntries.includes(entry));
            if (missingChecked.length || extraChecked.length) {
              logEvent({
                level: 'warn',
                message: 'Modal checked rows differ from expected bulk selection',
                modalTitle,
                detail: JSON.stringify({
                  expectedCount: normalizedEntries.length,
                  checkedCount: checkedNormalized.length,
                  missingChecked: missingChecked.slice(0, 20),
                  extraChecked: extraChecked.slice(0, 20),
                }),
              });
              matched = null;
            }
          } catch (diffError) {
            logEvent({ level: 'debug', message: 'Failed to diff modal checked rows against expected selection', modalTitle, detail: String(diffError && diffError.message || diffError) });
          }
        } catch (_) {}

          if (matched) {
            // Clica no botão selecionar e finaliza o modal
            const { locator: selecionarBtn } = await findVisibleLocatorInFrames(page, '#pbSelecionar');
            await selecionarBtn.click({ force: true });
            await page.waitForTimeout(300);

            // Após fechar modal, tenta ler o input próximo ao trigger no mainFrame e quaisquer listas SelectedEntitiesList
            try {
              const iframeLocator2 = page.locator('iframe#iFramePage');
              const mainFrame2 = await iframeLocator2.contentFrame().catch(() => null);
              // Re-obtem as linhas marcadas no modal e grava explicitamente no mainFrame para diagnóstico
              try {
                const modalCheckedPersist = await modalFrame.evaluate(() => Array.from(document.querySelectorAll('tr')).filter(r => { try { const cb = r.querySelector('input[type="checkbox"]'); return !!(cb && cb.checked); } catch(e){return false;} }).map(r => r.innerText.trim()).slice(0,200)).catch(() => []);
                if (mainFrame2 && modalCheckedPersist && modalCheckedPersist.length) {
                  await writeSelectedEntitiesIntoMainFrame(mainFrame2, modalCheckedPersist).catch(() => {});
                  logEvent({ level: 'info', message: 'Wrote modal selections into main frame (bulk)', detail: JSON.stringify(modalCheckedPersist.slice(0,10)), modalTitle });
                }
              } catch (e) {}
              if (mainFrame2) {
                // tenta extrair input visível associado ao triggerSelector
                try {
                  const nearValue = await mainFrame2.locator(`${triggerSelector}`).evaluate((el) => {
                    try {
                      // procura input text mais próximo subindo/descendo na arvore
                      function findNearInput(node) {
                        if (!node) return null;
                        if (node.tagName && node.tagName.toLowerCase() === 'input') return node;
                        for (const child of Array.from(node.childNodes || [])) {
                          if (child.tagName && child.tagName.toLowerCase() === 'input') return child;
                        }
                        return null;
                      }
                      let parent = el.parentElement;
                      while (parent) {
                        const inp = findNearInput(parent);
                        if (inp) return (inp.value || inp.innerText || '').toString();
                        parent = parent.parentElement;
                      }
                      return null;
                    } catch (e) { return null; }
                  }).catch(() => null);
                  logEvent({ level: 'info', message: `Main input near trigger after select: ${String(nearValue).slice(0,200)}`, modalTitle });
                } catch (e) {}

                try {
                  const selectedLists = await mainFrame2.evaluate(() => {
                    try {
                      const out = [];
                      Array.from(document.querySelectorAll('[id$="SelectedEntitiesList"], [class*="SelectedEntitiesList"]')).forEach(el => {
                        if (el.tagName === 'INPUT' || el.tagName === 'TEXTAREA') out.push(el.value || '');
                        else out.push(el.innerText || '');
                      });
                      return out;
                    } catch (e) { return []; }
                  }).catch(() => []);
                  logEvent({ level: 'info', message: `Main SelectedEntitiesList after select: ${selectedLists.length}`, detail: JSON.stringify(selectedLists.slice(0,10)), modalTitle });
                } catch (e) {}
              }
            } catch (_) {}

            logEvent({ level: 'info', message: `Modal finalizado (bulk): ${entries.join(', ')}`, modalTitle });
            try { await setReportProgress(reportName, 25, { lastMessage: 'modal_finished', step: 'modal_done' }); } catch (_) {}
            try { await saveShot(page, `modal-selected3-${sanitizeFileName(modalTitle)}`); } catch (_) {}
            return;
          }
      }
      // if matched is present but incomplete, log which entries were not matched
      if (matched && matched.found < matched.expected) {
        try {
          const unmatched = normalizedEntries.filter(ne => !(matched.matched || []).includes(ne));
          logEvent({ level: 'warn', message: `Bulk selection incomplete: found ${matched.found}/${matched.expected}`, detail: JSON.stringify({ unmatched: unmatched.slice(0,50) }), modalTitle });
          const sampleRows = matched.sampleRows || await modalFrame.evaluate(() => Array.from(document.querySelectorAll('tr')).slice(0,50).map(r => r.innerText)).catch(() => []);
          logEvent({ level: 'debug', message: 'Modal sample rows for diagnosis', detail: JSON.stringify(sampleRows.slice(0,50)), modalTitle });
        } catch (e) {}
      }
    }
  } catch (e) {
    // falhar na tentativa de bulk não deve interromper o fluxo — cai para seleção por item
    logEvent({ level: 'debug', message: `Bulk selection attempt failed: ${String(e && e.message || e)}`, modalTitle });
  }
  
  for (let vi = 0; vi < entries.length; vi++) {
    const currentValue = entries[vi];
    
    logEvent({
      level: 'info',
      message:
      `Selecionando no modal: ${currentValue}`,
      modalTitle,
    });
    
    // Wrap selection attempt in retries to recover from detached-frame errors
    let _withFrameResult = null;
    let _attemptSel = 0;
    let _selSuccess = false;
    while (_attemptSel < 3 && !_selSuccess) {
      try {
        // re-resolve modalFrame if it became detached
        try {
          await modalFrame.evaluate(() => true).catch(() => null);
        } catch (_) {
          // try to find modal frame again
          modalFrame = await findSelectionModalFrame(page, modalTitle, 5000).catch(() => null) || await findModalFrame(page, modalTitle, 5000).catch(() => null);
          if (!modalFrame) throw new Error('Modal frame desapareceu durante seleção');
        }

        // perform the selection using existing robust logic
        // Try quick DOM-only selection first to avoid triggering searches that may re-render
        // and clear previously checked items. If quick selection succeeds, skip typing.
        const normalizedToken = normalizeSelectionText(currentValue);
        const quickMarked = await modalFrame.evaluate(({ needle, rowValueSelectors }) => {
          try {
            const normalize = s => String(s || '')
              .normalize('NFD')
              .replace(/[\u0300-\u036f]/g, '')
              .replace(/[^\w\s]/g, ' ')
              .replace(/\s+/g, ' ')
              .trim()
              .toLowerCase();

            for (const row of document.querySelectorAll('tr')) {
              const parts = [];
              if (Array.isArray(rowValueSelectors) && rowValueSelectors.length) {
                for (const selector of rowValueSelectors) {
                  try {
                    const cell = row.querySelector(selector);
                    if (cell) parts.push(normalize(cell.innerText || ''));
                  } catch (e) {}
                }
              }
              if (!parts.length) {
                parts.push(...String(row.innerText || '').split(/\r?\n|\t/).map(text => normalize(text)).filter(Boolean));
              }

              if (!parts.some(part => part === needle)) continue;

              const cb = row.querySelector('input[type="checkbox"]');
              if (cb && !cb.checked) {
                try { cb.click(); } catch (e) { try { cb.checked = true; } catch(_){} }
                try { cb.dispatchEvent(new Event('change', { bubbles: true })); } catch(_){}
              }
              return !!(cb && cb.checked);
            }
          } catch (e) {}
          return false;
        }, { needle: normalizedToken, rowValueSelectors: modalDefinition.rowValueSelectors || [] }).catch(() => false);

        if (quickMarked) {
          logEvent({ level: 'info', message: `Checkbox marcado (quick DOM): ${currentValue}`, modalTitle });
          _selSuccess = true;
          break;
        }

        _withFrameResult = await withFrameRetry(async () => {
      // re-resolve modalInput fresh for this attempt
      const modalInputLocal = await resolveInput(modalFrame, searchSelectors, 4).catch(() => null);
      // LIMPA FILTRO ANTERIOR
      if (modalInputLocal) await modalInputLocal.fill('').catch(() => {});
      await page.waitForTimeout(500);
      const inputValue = modalInputLocal ? await modalInputLocal.inputValue().catch(() => '') : '';
      if (inputValue && inputValue.trim() && modalInputLocal) {
        await modalInputLocal.evaluate((el) => { el.value = ''; el.dispatchEvent(new Event('input', { bubbles: true })); });
        await page.waitForTimeout(300);
      }

      // DIGITA NOVO VALOR
      // Detect presence of any 'Limpar' control but DO NOT click it (clicking clears existing selections)
      const clickedClear = await modalFrame.evaluate(() => {
        try {
          const nodes = Array.from(document.querySelectorAll('input,button,a'));
          for (const n of nodes) {
            const txt = ((n.value || '') + ' ' + (n.innerText || '')).trim();
            if (/^\s*limpar\b/i.test(txt)) { return true; }
          }
        } catch (e) {}
        return false;
      }).catch(() => false);
      if (clickedClear) logEvent({ level: 'debug', message: 'Limpar encontrado (não clicado) antes de digitar novo valor', modalTitle });

      if (modalInputLocal) await fillLegacyInput(modalInputLocal, currentValue);
      if (modalInputLocal) await modalInputLocal.press('Enter').catch(() => {});

      const prevFirst = await modalFrame.evaluate(() => { try { const t = document.querySelector('tr'); return t ? t.innerText : ''; } catch (e) { return ''; } }).catch(() => '');

      const clicked = await modalFrame.evaluate(() => {
        try {
          const nodes = Array.from(document.querySelectorAll('input,button,a'));
          for (const n of nodes) {
            const txt = ((n.value || '') + ' ' + (n.innerText || '')).trim();
            if (/procurar/i.test(txt)) { try { n.click(); } catch(e){ try{ n.dispatchEvent(new MouseEvent('click',{bubbles:true})); }catch(_){} } return true; }
          }
        } catch (e) {}
        return false;
      }).catch(() => false);

      if (!clicked) {
        const procurarBtn = modalFrame.locator('#pbProcurar, input[name="pbProcurar"], input[id*="Procurar"], input[value*="Procurar"]');
        if (await procurarBtn.count().catch(() => 0)) {
          await procurarBtn.first().click({ force: true }).catch(() => {});
        }
      }

      const normalizedToken = normalizeSelectionText(currentValue);
      try { await modalFrame.waitForFunction((t) => { try { const norm = s => String(s || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/[^\w\s]/g, ' ').replace(/\s+/g, ' ').trim().toLowerCase(); return document.body && Array.from(String(document.body.innerText || '').split(/\r?\n/)).map(norm).some(line => line === t); } catch (e) { return false; } }, { timeout: 10000 }, normalizedToken).catch(() => {}); } catch (_) {}
      try { await modalFrame.locator('tr').first().waitFor({ timeout: 7000 }); } catch (_) {}

      // Primeiro: busca o índice da <tr> no DOM via evaluate (um roundtrip único, rápido)
    let cell = null;
    let cellFrame = null;
    let rowLocator = null;
    
    try {
          const idx = await modalFrame.evaluate(({ needle, rowValueSelectors }) => {
        try {
          function normalize(s) {
            try {
              return String(s || '')
                .normalize('NFD')
                .replace(/[\u0300-\u036f]/g, '')
                .replace(/[^\w\s]/g, ' ')
                .replace(/\s+/g, ' ')
                .trim()
                .toLowerCase();
            } catch (e) {
              return String(s || '').replace(/\s+/g, ' ').trim().toLowerCase();
            }
          }

          const rows = Array.from(document.querySelectorAll('tr'));
          for (let i = 0; i < rows.length; i++) {
                const parts = [];
                if (Array.isArray(rowValueSelectors) && rowValueSelectors.length) {
                  for (const selector of rowValueSelectors) {
                    try {
                      const cell = rows[i].querySelector(selector);
                      if (cell) parts.push(normalize(cell.innerText || ''));
                    } catch (e) {}
                  }
                }
                if (!parts.length) {
                  const text = normalize(rows[i].innerText || '');
                  if (text) parts.push(text);
                }
                if (parts.some(part => part === needle)) return i + 1;
          }
        } catch (e) {}
        return null;
          }, { needle: normalizedToken, rowValueSelectors: modalDefinition.rowValueSelectors || [] }).catch(() => null);
      logEvent({ level: 'debug', message: `Row index lookup for '${currentValue}': ${String(idx)}`, modalTitle });
      
      if (idx) {
        rowLocator = modalFrame.locator(`xpath=(//tr)[${idx}]`);
        cell = rowLocator;
        cellFrame = modalFrame;
      } else {
        // fallback para lógica mais ampla existente
        try {
          const res = await findRowByText(page, currentValue, 30000, modalFrame);
          cell = res.cell; cellFrame = res.frame;
        } catch (err) {
          // tenta fallback: remover prefixos comuns como 'Aplicacao' / 'Aplicação' e pesquisar o restante
          try {
            const alt = String(currentValue || '').replace(/^\s*(aplica[cç][ãa]o|aplicacao|conta)\s+/i, '').trim();
            if (alt && alt.length < String(currentValue || '').length) {
              logEvent({ level: 'warn', message: `Busca inicial falhou para '${currentValue}', tentando fallback com: ${alt}`, modalTitle });
              const res2 = await findRowByText(page, alt, 20000, modalFrame);
              cell = res2.cell; cellFrame = res2.frame;
            }
          } catch (_) {
            try {
              const tokens = String(currentValue || '').split(/\s+/).filter(Boolean);
              if (tokens.length > 1) {
                const lastTwo = tokens.slice(-2).join(' ');
                logEvent({ level: 'warn', message: `Tentando fallback tokens para '${currentValue}' -> '${lastTwo}'`, modalTitle });
                const res3 = await findRowByText(page, lastTwo, 20000, modalFrame);
                cell = res3.cell; cellFrame = res3.frame;
              }
            } catch (errFinal) {
              if (DEBUG_HTML) {
                try {
                  const modalHtml = await modalFrame.content().catch(() => '<!-- erro ao capturar -->');
                  const debugFile = path.join(SCREENSHOT_DIR, `debug-modal-${sanitizeFileName(modalTitle)}-${nowIso().replace(/[:.]/g, '-')}.html`);
                  await fs.promises.writeFile(debugFile, modalHtml || '<!-- vazio -->').catch(() => {});
                  const preview = await modalFrame.evaluate(() => Array.from(document.querySelectorAll('tr')).slice(0,20).map(r => r.innerText));
                  const previewFile = path.join(SCREENSHOT_DIR, `debug-modal-preview-${sanitizeFileName(modalTitle)}-${nowIso().replace(/[:.]/g, '-')}.json`);
                  await fs.promises.writeFile(previewFile, JSON.stringify(preview, null, 2)).catch(() => {});
                  logEvent({ level: 'info', message: `HTML do modal salvo em: ${debugFile}`, previewPath: previewFile });
                } catch (e) { logEvent({ level: 'warn', message: `Erro ao salvar HTML debug: ${e.message}` }); }
              }
              throw errFinal || err;
            }
          }
        }
      }
    } catch (e) {
      // segue para tentar localizar a row abaixo
    }
    try {
      const ancestor = cell.locator('xpath=ancestor::tr[1]').first();
      const ancCount = await ancestor.count().catch(() => 0);
      if (ancCount) {
        rowLocator = ancestor;
      } else {
        // Se o elemento encontrado já for uma <tr>, use-o
        const tagName = await cell.evaluate(el => el.tagName && el.tagName.toLowerCase()).catch(() => '');
        if (tagName === 'tr') rowLocator = cell;
      }
    } catch (_) {}
    
    // Se não obteve rowLocator, tenta encontrar uma <tr> que contenha o texto no mesmo frame
    if (!rowLocator) {
      try {
        const frameForSearch = cellFrame || modalFrame;
        const idx2 = await frameForSearch.evaluate((needle) => {
          try {
            function normalize(s) {
              try {
                return String(s || '')
                  .normalize('NFD')
                  .replace(/[\u0300-\u036f]/g, '')
                  .replace(/[^\w\s]/g, ' ')
                  .replace(/\s+/g, ' ')
                  .trim()
                  .toLowerCase();
              } catch (e) {
                return String(s || '').replace(/\s+/g, ' ').trim().toLowerCase();
              }
            }

            const rows = Array.from(document.querySelectorAll('tr'));
            for (let i = 0; i < rows.length; i++) {
              const text = normalize(rows[i].innerText || '');
              if (text && text === needle) return i + 1;
            }
          } catch (e) {}
          return null;
        }, normalizedToken).catch(() => null);
        if (idx2) rowLocator = frameForSearch.locator(`xpath=(//tr)[${idx2}]`);
      } catch (_) {}
    }
    
    // Finalmente, busca o checkbox dentro da row ou dentro do próprio elemento
    let checkbox = null;
    try {
      if (rowLocator) checkbox = rowLocator.locator('input[type="checkbox"]').first();
      else if (cell) checkbox = cell.locator('input[type="checkbox"]').first();
      
      if (checkbox && (await checkbox.count().catch(() => 0))) {
        const checked = await checkbox.isChecked().catch(() => false);
        if (!checked) {
          // attempt to check with frame-retry, re-resolving locator each attempt
          try {
            await withFrameRetry(async () => {
              const freshChk = (rowLocator || cell).locator('input[type="checkbox"]').first();
              if (await freshChk.count().catch(() => 0)) {
                await freshChk.check();
              } else {
                throw new Error('checkbox-not-found');
              }
            }, { name: `check-checkbox:${currentValue}`, retries: 2 });
          } catch (_) {
            // fallback: try click forced
            try {
              const freshChk2 = (rowLocator || cell).locator('input[type="checkbox"]').first();
              await freshChk2.click({ force: true }).catch(() => {});
            } catch (_) {
              // último recurso: marcar via evaluate no contexto do frame (único roundtrip)
              try {
                await (cellFrame || modalFrame).evaluate(({ needle, rowValueSelectors }) => {
                  const norm = s => String(s || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase();
                  for (const r of document.querySelectorAll('tr')) {
                    const parts = [];
                    if (Array.isArray(rowValueSelectors) && rowValueSelectors.length) {
                      for (const selector of rowValueSelectors) {
                        try {
                          const cell = r.querySelector(selector);
                          if (cell) parts.push(norm(cell.innerText || ''));
                        } catch (e) {}
                      }
                    }
                    if (!parts.length) {
                      const txt = norm(r.innerText || '');
                      if (txt) parts.push(txt);
                    }
                    if (parts.some(part => part === needle)) {
                      const cb = r.querySelector('input[type="checkbox"]');
                      if (cb && !cb.checked) { try { cb.click(); } catch(e) { cb.checked = true; } cb.dispatchEvent(new Event('change', { bubbles: true })); }
                      break;
                    }
                  }
                }, { needle: normalizedToken, rowValueSelectors: modalDefinition.rowValueSelectors || [] }).catch(() => {});
              } catch (_) {}
            }
          }
        }
        
        // aguarda confirmação de seleção por até 10s
        try {
          const startSel = Date.now();
          let confirmed = false;
          while (Date.now() - startSel < 10000) {
            confirmed = await (cellFrame || modalFrame).evaluate((needle) => {
              const norm = s => String(s || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase();
              for (const r of document.querySelectorAll('tr')) {
                if (norm(r.innerText || '') === needle) {
                  const cb = r.querySelector('input[type="checkbox"]');
                  return !!(cb && cb.checked);
                }
              }
              return false;
            }, normalizedToken).catch(() => false);
            if (confirmed) break;
            await page.waitForTimeout(250);
          }

          if (!confirmed) {
            await saveDebugState('checkbox-timeout').catch(() => {});
            logEvent({ level: 'warn', message: `Timeout ao selecionar valor no modal, pulando: ${currentValue}`, modalTitle });
            return { skipped: true };
          }

          logEvent({ level: 'info', message: `Checkbox marcado: ${currentValue}`, modalTitle });
          try {
            const pct = 8 + Math.round(((vi + 1) / entries.length) * 15);
            await setReportProgress(reportName, Math.min(25, pct), { lastMessage: `selected:${currentValue}`, step: 'modal_select' });
          } catch (_) {}
        } catch (e) {
          // em caso de erro durante a confirmação, registra e tenta continuar
          logEvent({ level: 'warn', message: `Erro ao confirmar seleção para ${currentValue}, pulando: ${String(e && e.message || e)}`, modalTitle });
          return { skipped: true };
        }
      } else {
        // se não encontrou checkbox via locators, tenta fallback evaluate para marcar
        await (cellFrame || modalFrame).evaluate((needle) => {
          const norm = s => s.normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase();
          for (const r of document.querySelectorAll('tr')) {
              if (norm(r.innerText) === needle) {
                const cb = r.querySelector('input[type="checkbox"]');
                if (cb && !cb.checked) cb.click();
                break;
              }
            }
        }, normalizedToken).catch(() => {});
        // confirma seleção por até 10s
        try {
          const startSelFb = Date.now();
          let confirmedFb = false;
          while (Date.now() - startSelFb < 10000) {
            confirmedFb = await (cellFrame || modalFrame).evaluate((needle) => {
              const norm = s => String(s || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase();
              for (const r of document.querySelectorAll('tr')) {
                if (norm(r.innerText || '') === needle) {
                  const cb = r.querySelector('input[type="checkbox"]');
                  return !!(cb && cb.checked);
                }
              }
              return false;
            }, normalizedToken).catch(() => false);
            if (confirmedFb) break;
            await page.waitForTimeout(250);
          }
          if (!confirmedFb) {
            await saveDebugState('checkbox-timeout').catch(() => {});
            logEvent({ level: 'warn', message: `Timeout ao selecionar valor no modal (fallback), pulando: ${currentValue}`, modalTitle });
            return { skipped: true };
          }
          logEvent({ level: 'info', message: `Checkbox marcado via fallback evaluate: ${currentValue}`, modalTitle });
        } catch (e) {
          logEvent({ level: 'warn', message: `Erro ao confirmar seleção via fallback para ${currentValue}, pulando: ${String(e && e.message || e)}`, modalTitle });
          return { skipped: true };
        }
      }
    
    } catch (err) {
      await saveDebugState('checkbox-error');
      logEvent({ level: 'error', message: `Erro ao manipular checkbox para ${currentValue}: ${err.message}` });
      throw err;
    }
        }, { name: `selectViaModal:${currentValue}`, retries: 2 });

        _selSuccess = true;
      } catch (err) {
        const msg = String(err && (err.message || err));
        logEvent({ level: 'warn', message: `Erro ao selecionar '${currentValue}', tentativa=${_attemptSel + 1}`, detail: msg, modalTitle });
        if (/Frame was detached|Execution context was destroyed|Target closed|Modal frame desapareceu/i.test(msg) && _attemptSel < 2) {
          _attemptSel++;
          await page.waitForTimeout(500 * _attemptSel);
          // try to re-find the modal frame before next attempt
          modalFrame = await findSelectionModalFrame(page, modalTitle, 5000).catch(() => null) || await findModalFrame(page, modalTitle, 5000).catch(() => null);
          continue;
        }
        // non-transient: rethrow
        throw err;
      }
    }

    if (_withFrameResult && _withFrameResult.skipped) continue;
    
    saveShot(page, `checkbox-${sanitizeFileName(currentValue)}`).catch(() => {});
    await page.waitForTimeout(200);
    
    // ================================================
    // LIMPA FILTRO PARA PRÓXIMA SELEÇÃO
    // ================================================
    // don't clear the modal input here; clearing can reset previous selections in some modals
    try { await page.waitForTimeout(300); } catch (_) {}
  }
  
  // =====================================================
  // BOTÃO SELECIONAR
  // =====================================================
  
  // Prefer clicking Selecionar inside the modal frame, fallback to cross-frame locator
  let clickedSelec = false;
  try {
    const selBtnFrame = modalFrame.locator('#pbSelecionar, input[id*="pbSelecionar"], input[name="pbSelecionar"], button#pbSelecionar').first();
    if (await selBtnFrame.count().catch(() => 0)) {
      await selBtnFrame.click({ force: true }).catch(() => {});
      clickedSelec = true;
    }
  } catch (_) {}
  let selecionarBtn = null;
  if (!clickedSelec) {
    const res = await findVisibleLocatorInFrames(page, '#pbSelecionar');
    selecionarBtn = res && res.locator ? res.locator : null;
  }
  // Antes de clicar em Selecionar, captura linhas marcadas no modal
  try {
    const modalChecked = await modalFrame.evaluate(() => Array.from(document.querySelectorAll('tr')).filter(r => { try { const cb = r.querySelector('input[type="checkbox"]'); return !!(cb && cb.checked); } catch(e){return false;} }).map(r => r.innerText.trim()).slice(0,200)).catch(() => []);
    logEvent({ level: 'info', message: `Modal checked rows (pre-select): ${modalChecked.length}`, detail: JSON.stringify(modalChecked.slice(0,50)), modalTitle });
    try {
      const iframeLocatorBefore = page.locator('iframe#iFramePage');
      const mainFrameBefore = await iframeLocatorBefore.contentFrame().catch(() => null);
      if (mainFrameBefore && modalChecked && modalChecked.length) {
        await writeSelectedEntitiesIntoMainFrame(mainFrameBefore, modalChecked).catch(() => {});
        logEvent({ level: 'info', message: 'Persisted modal selections into main frame (pre-click)', detail: JSON.stringify(modalChecked.slice(0,10)), modalTitle });
      }
    } catch (e) {}
  } catch (_) {}

  try {
    if (selecionarBtn) await selecionarBtn.click({ force: true });
    else if (!clickedSelec) {
      // final fallback: try to click any button in the modal via evaluate
      try { await modalFrame.evaluate(() => { const b = document.querySelector('#pbSelecionar,input[id*="pbSelecionar"],input[name*="pbSelecionar"],button#pbSelecionar'); if (b) try { b.click(); } catch(e) { try { b.dispatchEvent(new MouseEvent('click',{bubbles:true})); } catch(_){} } }).catch(() => {}); } catch(_) {}
    }
  } catch (_) {}

  await page.waitForTimeout(300);

  // Após fechar modal, tenta ler o input próximo ao trigger no mainFrame e quaisquer listas SelectedEntitiesList
  try {
    const iframeLocator2 = page.locator('iframe#iFramePage');
    const mainFrame2 = await iframeLocator2.contentFrame().catch(() => null);
    // Re-obtem as linhas marcadas no modal e grava explicitamente no mainFrame para diagnóstico
    try {
      const modalCheckedPersist = await modalFrame.evaluate(() => Array.from(document.querySelectorAll('tr')).filter(r => { try { const cb = r.querySelector('input[type="checkbox"]'); return !!(cb && cb.checked); } catch(e){return false;} }).map(r => r.innerText.trim()).slice(0,200)).catch(() => []);
      if (mainFrame2 && modalCheckedPersist && modalCheckedPersist.length) {
        await writeSelectedEntitiesIntoMainFrame(mainFrame2, modalCheckedPersist).catch(() => {});
        logEvent({ level: 'info', message: 'Wrote modal selections into main frame', detail: JSON.stringify(modalCheckedPersist.slice(0,10)), modalTitle });
      }
    } catch (e) {}
    if (mainFrame2) {
      try {
          const nearInfo = await mainFrame2.locator(`${triggerSelector}`).evaluate((el) => {
          try {
            function findNearInput(node) {
              if (!node) return null;
              if (node.tagName && node.tagName.toLowerCase() === 'input') return node;
              for (const child of Array.from(node.childNodes || [])) {
                if (child.tagName && child.tagName.toLowerCase() === 'input') return child;
              }
              return null;
            }
            let parent = el.parentElement;
            while (parent) {
              const inp = findNearInput(parent);
              if (inp) return (inp.value || inp.innerText || '').toString();
              parent = parent.parentElement;
            }
            return null;
          } catch (e) { return null; }
          }).catch(() => null);
          // tenta setar explicitamente o input visível próximo ao trigger com os nomes selecionados
          try {
            const combined = (modalCheckedPersist || []).join('\n');
            await mainFrame2.evaluate((sel, trigger) => {
              try {
                function findNearInput(node) {
                  if (!node) return null;
                  if (node.tagName && node.tagName.toLowerCase() === 'input') return node;
                  for (const child of Array.from(node.childNodes || [])) {
                    if (child.tagName && child.tagName.toLowerCase() === 'input') return child;
                  }
                  return null;
                }
                const el = document.querySelector(trigger);
                if (!el) return null;
                let parent = el.parentElement;
                while (parent) {
                  const inp = findNearInput(parent);
                  if (inp) {
                    try { inp.value = sel; inp.dispatchEvent(new Event('change', { bubbles: true })); } catch(e) {}
                    return true;
                  }
                  parent = parent.parentElement;
                }
                return false;
              } catch (e) { return false; }
            }, combined, triggerSelector).catch(() => {});
            logEvent({ level: 'info', message: `Wrote combined selection into near input (len=${String(modalCheckedPersist.length)})`, modalTitle });
          } catch (e) {}
      } catch (e) {}

      try {
        const selectedLists = await mainFrame2.evaluate(() => {
          try {
            const out = [];
            Array.from(document.querySelectorAll('[id$="SelectedEntitiesList"], [class*="SelectedEntitiesList"]')).forEach(el => {
              if (el.tagName === 'INPUT' || el.tagName === 'TEXTAREA') out.push(el.value || '');
              else out.push(el.innerText || '');
            });
            return out;
          } catch (e) { return []; }
        }).catch(() => []);
          try {
            logEvent({ level: 'info', message: `Main SelectedEntitiesList after select: ${selectedLists.length}`, detail: JSON.stringify(selectedLists.slice(0,10)), modalTitle });
            // If main frame lists don't include all selectedSoFar, re-write merged selections
            try {
              const flat = (Array.isArray(selectedLists) ? selectedLists.join('\n') : String(selectedLists || '')).split('\n').map(s=>s.trim()).filter(Boolean);
              const missing = (selectedSoFar || []).filter(s => !flat.includes(s.trim()));
              if (missing && missing.length) {
                await writeSelectedEntitiesIntoMainFrame(mainFrame2, selectedSoFar).catch(() => {});
                logEvent({ level: 'warn', message: 'Main frame missing some selections after click; re-wrote merged selections', detail: JSON.stringify({ missingCount: missing.length, missing: missing.slice(0,10) }), modalTitle });
              }
            } catch (e) {}
          } catch (e) { logEvent({ level: 'debug', message: 'Failed logging selectedLists', detail: String(e && e.message || e) }); }
      } catch (e) {}
    }
  } catch (_) {}
  
  logEvent({
    level: 'info',
    message:
    `Modal finalizado: ${entries.join(', ')}`,
    modalTitle,
  });
  try { await setReportProgress(reportName, 25, { lastMessage: 'modal_finished', step: 'modal_done' }); } catch (_) {}
  try { await saveShot(page, `modal-selected-${sanitizeFileName(modalTitle)}`); } catch (_) {}

  // =====================================================
  // MITIGAÇÃO DEFENSIVA: limpar possíveis estados remanescentes após fechar modal
  // ====================================================='
  try {
    // reseta cache em memória e grava
    MODAL_CACHE = {};
    try { await saveModalCache(); } catch (_) {}
  } catch (_) {}
  // Optional: post-modal cleanup of SelectedEntitiesList can interfere with app behavior
  // Enable explicitly via env var ENABLE_POST_MODAL_CLEANUP=1 when needed.
  try {
    if (process.env.ENABLE_POST_MODAL_CLEANUP === '1') {
      try {
        const iframeLocator2 = page.locator('iframe#iFramePage');
        const mainFrame2 = await iframeLocator2.contentFrame().catch(() => null);
        if (mainFrame2) {
          await mainFrame2.evaluate(() => {
            try {
              Array.from(document.querySelectorAll('[id$="SelectedEntitiesList"], [class*="SelectedEntitiesList"]')).forEach(el => {
                if (el.tagName === 'INPUT' || el.tagName === 'TEXTAREA') el.value = '';
                else el.innerHTML = '';
              });
              Array.from(document.querySelectorAll('[id*="contador"], [name*="contador"]')).forEach(el => {
                if (el.tagName === 'INPUT' || el.tagName === 'TEXTAREA') el.value = '0';
                else el.textContent = '0';
              });
            } catch (e) {}
          }).catch(() => {});
        }
      } catch (_) {}
    } else {
      logEvent({ level: 'debug', message: 'Skipping post-modal cleanup (ENABLE_POST_MODAL_CLEANUP!=1)', modalTitle });
    }
  } catch (_) {}
}

async function selectViaModalByGrid({
  page,
  triggerSelector,
  modalTitle,
  searchInputSelector,
  value,
  values,
  enablePagination = false,
}) {
  const entries = uniqueNonEmpty(values || [value]);
  if (!entries.length) return;
  const MODAL_SELECT_TIMEOUT_MS = Number(process.env.MODAL_SELECT_TIMEOUT_MS || 10000);
  
  const modalDefinition = getModalSelectionDefinition(modalTitle);
  const searchSelectors = uniqueNonEmpty([
    ...(Array.isArray(searchInputSelector) ? searchInputSelector : [searchInputSelector]),
    ...modalDefinition.searchInputSelectors,
  ]);
  
  logEvent({
    level: 'info',
    message: `Selecionando valores: ${entries.join(', ')}`,
    modalTitle,
    count: entries.length,
  });
  
  const iframeLocator = page.locator('iframe#iFramePage');
  await iframeLocator.waitFor({ state: 'attached', timeout: 30000 });
  const mainFrame = await iframeLocator.contentFrame();
  if (!mainFrame) throw new Error('Nao foi possivel obter contentFrame do iFramePage.');
  
  try {
    await ensureExpandedFilters(mainFrame);
  } catch (e) {
    logEvent({ level: 'debug', message: `Falha ao expandir filtros no mainFrame: ${String(e && e.message || e)}` });
  }
  
  async function fillLegacyInput(locator, nextValue) {
    try {
      await locator.scrollIntoViewIfNeeded();
    } catch (_) {}
    await locator.evaluate((el, value) => {
      el.style.display = '';
      el.style.visibility = 'visible';
      el.style.opacity = '1';
      el.focus();
      el.value = '';
      el.dispatchEvent(new Event('input', { bubbles: true }));
      el.value = value;
      el.dispatchEvent(new Event('input', { bubbles: true }));
      el.dispatchEvent(new Event('change', { bubbles: true }));
      el.dispatchEvent(new KeyboardEvent('keyup', { bubbles: true }));
      el.dispatchEvent(new Event('blur', { bubbles: true }));
    }, String(nextValue));
  }
  
  const trigger = mainFrame.locator(`${triggerSelector}[src*="botProcurar.png"]`).first();
  // MITIGAÇÃO DEFENSIVA: limpar estado modal antes de abrir (evita vazamento entre modais)
  try {
    MODAL_CACHE = {};
    try { await saveModalCache(); } catch (_) {}
  } catch (_) {}

  try {
    await mainFrame.evaluate(() => {
      try {
        Array.from(document.querySelectorAll('[id$="SelectedEntitiesList"], [class*="SelectedEntitiesList"]')).forEach(el => {
          if (el.tagName === 'INPUT' || el.tagName === 'TEXTAREA') el.value = '';
          else el.innerHTML = '';
        });
        Array.from(document.querySelectorAll('[id*="contador"], [name*="contador"]')).forEach(el => {
          if (el.tagName === 'INPUT' || el.tagName === 'TEXTAREA') el.value = '0';
          else el.textContent = '0';
        });
      } catch (e) {}
    }).catch(() => {});
  } catch (_) {}

  await acquireModalLock();
  let _modalLockHeld = true;
  try {
    await trigger.click({ force: true });

    let modalFrame = await findSelectionModalFrame(page, modalTitle, 30000);
  if (!modalFrame) {
    modalFrame = await findModalFrame(page, modalTitle, 30000);
  }
  if (!modalFrame) {
    throw new Error(`Modal nao encontrado: ${modalTitle}`);
  }
  
  logEvent({
    level: 'info',
    message: `Modal encontrado: ${modalTitle}`,
    frameUrl: modalFrame.url(),
  });
  
  async function saveDebugState(tag) {
    if (!DEBUG_HTML) return null;
    try {
      await fs.promises.mkdir(SCREENSHOT_DIR, { recursive: true });
    } catch (_) {}
    const ts = nowIso().replace(/[:.]/g, '-');
    const base = sanitizeFileName(`${modalTitle}-${tag}-${ts}`);
    const htmlPath = path.join(SCREENSHOT_DIR, `${base}.html`);
    const previewPath = path.join(SCREENSHOT_DIR, `${base}-preview.json`);
    const shotPath = path.join(SCREENSHOT_DIR, `${base}.png`);
    try {
      const modalHtml = await modalFrame.content().catch(() => '<!-- erro ao capturar -->');
      await fs.promises.writeFile(htmlPath, modalHtml || '<!-- vazio -->').catch(() => {});
    } catch (e) {
      logEvent({ level: 'warn', message: `Erro ao salvar HTML debug: ${e.message}` });
    }
    try {
      const preview = await modalFrame.evaluate(() => Array.from(document.querySelectorAll('tr')).slice(0, 30).map(r => r.innerText)).catch(() => []);
      await fs.promises.writeFile(previewPath, JSON.stringify(preview, null, 2)).catch(() => {});
    } catch (e) {
      logEvent({ level: 'warn', message: `Erro ao salvar preview debug: ${e.message}` });
    }
    try {
      await page.screenshot({ path: shotPath, fullPage: true }).catch(() => {});
    } catch (e) {
      logEvent({ level: 'warn', message: `Erro ao salvar screenshot debug: ${e.message}` });
    }
    logEvent({ level: 'info', message: `Debug salvo: html=${htmlPath}, preview=${previewPath}, shot=${shotPath}` });
    return { htmlPath, previewPath, shotPath };
  }
  
  for (const selector of searchSelectors) {
    try {
      await modalFrame.locator(selector).first().waitFor({ state: 'attached', timeout: 5000 }).catch(() => {});
    } catch (_) {}
  }

  const isDocsModal = normalizeSelectionText(modalTitle || '').includes('consulta de documentos');
  if (isDocsModal && enablePagination !== false) {
    const paginatedResult = await selectDocumentosAcrossPages({
      modalFrame,
      modalTitle,
      values: entries,
      rowCheckboxSelector: modalDefinition.rowCheckboxSelector,
    }).catch((error) => {
      logEvent({
        level: 'warn',
        message: 'Seleção paginada de documentos falhou ao executar o fluxo no modal.',
        modalTitle,
        detail: {
          desiredCount: entries.length,
          error: String(error && (error.stack || error.message || error)),
        },
      });
      return {
        found: 0,
        expected: entries.length,
        matched: [],
        documents: [],
        totalPages: 1,
        error: String(error && (error.stack || error.message || error)),
      };
    });

    if (paginatedResult && paginatedResult.found > 0) {
      logEvent({
        level: 'info',
        message: `Seleção paginada de documentos concluída: ${paginatedResult.matched.join(', ')}`,
        modalTitle,
        totalPages: paginatedResult.totalPages,
      });
    } else {
      logEvent({
        level: 'warn',
        message: `Seleção paginada de documentos incompleta: ${paginatedResult ? `${paginatedResult.found}/${paginatedResult.expected}` : 'falhou'}`,
        modalTitle,
        detail: paginatedResult && paginatedResult.error ? {
          error: paginatedResult.error,
          desiredCount: entries.length,
          found: paginatedResult.found,
          expected: paginatedResult.expected,
          totalPages: paginatedResult.totalPages,
        } : {
          desiredCount: entries.length,
          found: paginatedResult ? paginatedResult.found : 0,
          expected: paginatedResult ? paginatedResult.expected : entries.length,
          totalPages: paginatedResult ? paginatedResult.totalPages : 1,
        },
      });
    }

    try {
      const selBtn = modalFrame.locator('#pbSelecionar, input[id*="pbSelecionar"], input[name="pbSelecionar"], button#pbSelecionar').first();
      if (await selBtn.count().catch(() => 0)) {
        await selBtn.click({ force: true }).catch(() => {});
      } else {
        const { locator: selecionarBtn } = await findVisibleLocatorInFrames(page, '#pbSelecionar');
        if (selecionarBtn) await selecionarBtn.click({ force: true }).catch(() => {});
      }
    } catch (_) {}

    try { await saveShot(page, `modal-selected2-${sanitizeFileName(modalTitle)}`); } catch (_) {}
    return;
  }
  
  const modalInput = await findVisibleInput(modalFrame, searchSelectors).catch(() => null);
  if (!modalInput) {
    await saveDebugState('input-not-found').catch(() => {});
    logEvent({
      level: 'warn',
      message: `Nao encontrei input visivel no modal; prosseguindo com selecao via DOM: ${modalTitle}`,
      modalTitle,
      searchSelectors,
    });
  }

  // acumulador local de seleções já marcadas no modal (texto das linhas)
  let selectedSoFar = [];

  // TENTA SELEÇÃO EM LOTE ANTES DE PROCESSAR POR ITEM (DOM-FIRST)
  try {
    const normalizedEntries = entries.map(e => normalizeSelectionText(e)).filter(Boolean);
    logEvent({ level: 'debug', message: 'Bulk selection (grid): normalized entries', detail: JSON.stringify(normalizedEntries.slice(0,50)), modalTitle });
    if (normalizedEntries.length) {
      // tentativa rápida: acionar checkbox de header (select-all) se existir
      try {
        await withFrameRetry(async () => {
          await modalFrame.evaluate((rowCheckboxSel) => {
            try {
              const header = document.querySelector('thead input[type="checkbox"], tr th input[type="checkbox"], .ui-table thead input[type="checkbox"], input[data-select-all], [data-select-all]');
              if (header && !header.checked) {
                try { header.click(); } catch (e) { header.checked = true; header.dispatchEvent(new Event('change', { bubbles: true })); }
              }
              return true;
            } catch (e) { return false; }
          }, modalDefinition.rowCheckboxSelector).catch(() => {});
        }, { name: 'dom-select-all-attempt' }).catch(() => {});
      } catch (_) {}

      // scroll para renderizar linhas virtualizadas
      await withFrameRetry(async () => {
        await modalFrame.evaluate(() => {
          const scroller = document.querySelector('.ui-table') || document.querySelector('.modal-body') || document.scrollingElement || document.documentElement;
          if (!scroller) return;
          const height = scroller.scrollHeight || (document.body && document.body.scrollHeight) || 0;
          let y = 0;
          while (y < height) {
            scroller.scrollTop = y;
            y += 800;
          }
          scroller.scrollTop = height;
        }).catch(() => {});
      }, { name: 'bulk-scroll' }).catch(() => {});

      const useDocsFallback = false;
      let matched = null;
      if (useDocsFallback) {
        try {
          const maxPages = 100;
          const accum = new Set();
          const accumDocumentos = new Set();
          let rowsCount = 0;
          let sampleRows = [];
          for (let pageIdx = 0; pageIdx < maxPages; pageIdx++) {
            const res = await withFrameRetry(async () => await modalFrame.evaluate(({ normalizedEntries, rowCheckboxSel, rowValueSelectors, nextSelectors }) => {
              function normalize(s) {
                try { return (s || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/[^\w\s]/g, ' ').replace(/\s+/g, ' ').trim().toLowerCase(); } catch (e) { return (s || '').replace(/[^\w\s]/g, ' ').replace(/\s+/g, ' ').trim().toLowerCase(); }
              }
              const candidateRowSelectors = ['#tabelaResultado tr[indice]', '#tabelaResultado tr[linha="true"]', 'tr[indice]', 'tr[linha="true"]', 'tbody > tr', 'tr'];
              let rows = [];
              for (const sel of candidateRowSelectors) {
                try { const r = Array.from(document.querySelectorAll(sel) || []); if (r && r.length) { rows = r; break; } } catch (e) {}
              }
              if (!rows || !rows.length) rows = Array.from(document.querySelectorAll('tr'));

              const matchedPage = new Set();
              const documentos = Array.from(document.querySelectorAll('tr[id^="linha_"]')).map(tr => tr.querySelectorAll('td')[2]?.innerText.trim()).filter(Boolean);
              for (const row of rows) {
                try {
                  const parts = [];
                  if (Array.isArray(rowValueSelectors) && rowValueSelectors.length) {
                    for (const rsel of rowValueSelectors) {
                      try { const c = row.querySelector(rsel); if (c) parts.push(normalize(c.innerText || '')); } catch (e) {}
                    }
                  }
                  if (!parts.length) {
                    const raw = String(row.innerText || '');
                    parts.push(...raw.split(/\r?\n|\t/).map(s => normalize(s)).filter(Boolean));
                  }
                  for (let i = 0; i < normalizedEntries.length; i++) {
                    const ne = normalizedEntries[i];
                    if (!ne || matchedPage.has(ne)) continue;
                      const ok = parts.some(p => p === ne);
                    if (ok) {
                      const checkbox = row.querySelector(rowCheckboxSel || 'input[type="checkbox"]');
                      if (checkbox && !checkbox.checked) {
                        try { checkbox.click(); } catch (e) { try { checkbox.checked = true; checkbox.dispatchEvent(new Event('change', { bubbles: true })); } catch (_) {} }
                        try { checkbox.dispatchEvent(new Event('change', { bubbles: true })); } catch (_) {}
                      }
                      const selected = !!(checkbox && checkbox.checked);
                      if (!selected) {
                        continue;
                      }
                      matchedPage.add(ne);
                    }
                  }
                } catch (e) { }
              }

              let nextClicked = false;
              for (const ns of nextSelectors) {
                try {
                  const el = document.querySelector(ns);
                  if (el) {
                    const disabled = el.disabled || (el.classList && el.classList.contains && el.classList.contains('ui-state-disabled')) || el.getAttribute('aria-disabled') === 'true';
                    if (!disabled) {
                      try { el.click(); } catch (e) { }
                      nextClicked = true;
                      break;
                    }
                  }
                } catch (e) {}
              }
              return { matchedPage: Array.from(matchedPage), rowsCount: rows.length, sampleRows: rows.slice(0,30).map(r => r.innerText), nextClicked, documentos };
            }, { normalizedEntries, rowCheckboxSel: modalDefinition.rowCheckboxSelector, rowValueSelectors: (modalDefinition && modalDefinition.rowValueSelectors) || [], nextSelectors: DOCUMENTOS_NEXT_PAGE_SELECTORS }), { name: 'bulk-evaluate-page' });

            if (!res) break;
            rowsCount = rowsCount || res.rowsCount || 0;
            sampleRows = sampleRows.length ? sampleRows : (res.sampleRows || []);
            (res.matchedPage || []).forEach(m => accum.add(m));
            (res.documentos || []).forEach(d => accumDocumentos.add(d));
            if (accum.size >= normalizedEntries.length) {
              matched = { found: accum.size, expected: normalizedEntries.length, matched: Array.from(accum), rowsCount, sampleRows, documentos: Array.from(accumDocumentos) };
              break;
            }
            if (!res.nextClicked) break;
            try { await page.waitForTimeout(300); } catch (e) {}
          }
          if (!matched) matched = { found: accum.size, expected: normalizedEntries.length, matched: Array.from(accum), rowsCount, sampleRows, documentos: Array.from(accumDocumentos) };
        } catch (err) {
          matched = null;
          try {
            if (DEBUG_HTML) {
              try {
                const ts = nowIso().replace(/[:.]/g, '-');
                const dumpName = sanitizeFileName(`${modalTitle}-bulk-evaluate-error-${ts}.json`);
                const dumpPath = path.join(SCREENSHOT_DIR, dumpName);
                const payload = { error: String(err && (err.stack || err.message || err)), time: new Date().toISOString() };
                await fs.promises.writeFile(dumpPath, JSON.stringify(payload, null, 2)).catch(() => {});
                logEvent({ level: 'warn', message: `Saved bulk-evaluate error JSON`, detail: { dumpPath }, modalTitle });
              } catch (e) {}
            }
          } catch (e) {}
        }
      } else {
        // Non-paginated single-page evaluate (preserve original behavior)
        try {
          matched = await withFrameRetry(async () => await modalFrame.evaluate(({ normalizedEntries, rowCheckboxSel, rowValueSelectors }) => {
            function normalize(s) {
              try { return (s || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/[^\w\s]/g, ' ').replace(/\s+/g, ' ').trim().toLowerCase(); } catch (e) { return (s || '').replace(/[^\w\s]/g, ' ').replace(/\s+/g, ' ').trim().toLowerCase(); }
            }
            const candidateRowSelectors = ['#tabelaResultado tr[indice]', '#tabelaResultado tr[linha="true"]', 'tr[indice]', 'tr[linha="true"]', 'tbody > tr', 'tr'];
            let rows = [];
            for (const sel of candidateRowSelectors) {
              try { const r = Array.from(document.querySelectorAll(sel) || []); if (r && r.length) { rows = r; break; } } catch (e) {}
            }
            if (!rows || !rows.length) rows = Array.from(document.querySelectorAll('tr'));

            const seen = new Set();
            for (const row of rows) {
              try {
                const parts = [];
                if (Array.isArray(rowValueSelectors) && rowValueSelectors.length) {
                  for (const rsel of rowValueSelectors) {
                    try { const c = row.querySelector(rsel); if (c) parts.push(normalize(c.innerText || '')); } catch (e) {}
                  }
                }
                if (!parts.length) {
                  const raw = String(row.innerText || '');
                  parts.push(...raw.split(/\r?\n|\t/).map(s => normalize(s)).filter(Boolean));
                }
                for (let i = 0; i < normalizedEntries.length; i++) {
                  const ne = normalizedEntries[i];
                  if (!ne || seen.has(ne)) continue;
                  const ok = parts.some(p => p === ne);
                  if (ok) {
                    const checkbox = row.querySelector(rowCheckboxSel || 'input[type="checkbox"]');
                      if (checkbox && !checkbox.checked) {
                        try { checkbox.click(); } catch (e) { try { checkbox.checked = true; checkbox.dispatchEvent(new Event('change', { bubbles: true })); } catch (_) {} }
                        try { checkbox.dispatchEvent(new Event('change', { bubbles: true })); } catch (_) {}
                      }
                      const selected = !!(checkbox && checkbox.checked);
                      if (!selected) {
                        continue;
                    }
                    seen.add(ne);
                  }
                }
              } catch (e) { }
            }
            return { found: seen.size, expected: normalizedEntries.length, matched: Array.from(seen), rowsCount: rows.length, sampleRows: rows.slice(0,30).map(r => r.innerText) };
          }, { normalizedEntries, rowCheckboxSel: modalDefinition.rowCheckboxSelector, rowValueSelectors: (modalDefinition && modalDefinition.rowValueSelectors) || [] }), { name: 'bulk-evaluate' }).catch(() => null);
        } catch (err) {
          matched = null;
        }
      }

      if (matched) {
        logEvent({ level: 'debug', message: 'Bulk grid evaluate result', detail: JSON.stringify(matched), modalTitle });
        try {
          if (DEBUG_HTML) {
            try {
              const ts = nowIso().replace(/[:.]/g, '-');
              const dumpName = sanitizeFileName(`${modalTitle}-bulk-evaluate-${ts}.json`);
              const dumpPath = path.join(SCREENSHOT_DIR, dumpName);
              await fs.promises.writeFile(dumpPath, JSON.stringify(matched, null, 2)).catch(() => {});
              logEvent({ level: 'info', message: `Saved bulk-evaluate debug JSON`, detail: { dumpPath }, modalTitle });
            } catch (e) {}
          }
        } catch (e) {}
      }

      if (matched && matched.found >= matched.expected) {
        logEvent({ level: 'info', message: `Bulk modal (grid) selection succeeded: ${matched.matched.join(', ')}`, modalTitle });
        // clica em Selecionar dentro do modal
        try {
          const selBtn = modalFrame.locator('#pbSelecionar, input[id*="pbSelecionar"], input[name="pbSelecionar"], button#pbSelecionar').first();
          if (await selBtn.count().catch(() => 0)) {
            await selBtn.click({ force: true }).catch(() => {});
          } else {
            const { locator: selecionarBtn } = await findVisibleLocatorInFrames(page, '#pbSelecionar');
            if (selecionarBtn) await selecionarBtn.click({ force: true }).catch(() => {});
          }
        } catch (_) {}

        try { await saveShot(page, `modal-selected2-${sanitizeFileName(modalTitle)}`); } catch (_) {}
        return;
      }
      if (matched && matched.found < matched.expected) {
        const unmatched = normalizedEntries.filter(ne => !(matched.matched || []).includes(ne));
        logEvent({ level: 'warn', message: `Bulk grid selection incomplete: found ${matched.found}/${matched.expected}`, detail: JSON.stringify({ unmatched: unmatched.slice(0,50) }), modalTitle });
      }
    }
  } catch (e) {
    logEvent({ level: 'debug', message: `Bulk grid selection attempt failed: ${String(e && e.message || e)}`, modalTitle });
  }
  
  for (const currentValue of entries) {
    logEvent({ level: 'info', message: `Entry start: selecionando no modal: ${currentValue}`, modalTitle });
    logEvent({ level: 'debug', message: `Entry debug: selectedSoFar count=${selectedSoFar.length}`, detail: selectedSoFar.slice(0,50), modalTitle });

    let attempt = 0;
    let succeeded = false;

    while (attempt < 3 && !succeeded) {
      attempt++;
      try {
        // verifica se o modalFrame ainda está disponível, se não re-resolve
        try {
          await modalFrame.evaluate(() => true).catch(() => null);
        } catch (_) {
          modalFrame = await findSelectionModalFrame(page, modalTitle, 5000).catch(() => null) || await findModalFrame(page, modalTitle, 5000).catch(() => null);
          if (!modalFrame) throw new Error('Modal frame desapareceu durante seleção');
        }

        // re-resolve modal input fresh for this attempt
        const modalInputLocal = await findVisibleInput(modalFrame, searchSelectors).catch(() => null);
        if (!modalInputLocal) {
          await saveDebugState('input-not-found-attempt').catch(() => {});
          throw new Error('input-not-found');
        }

        // Reaplica seleções previamente marcadas (se houver) para evitar perda
        // quando uma nova busca reposicionar/limpar a lista.
        if (selectedSoFar && selectedSoFar.length) {
          try {
            await modalFrame.evaluate((prev) => {
              try {
                const norm = s => String(s || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase();
                const needles = prev.map(p => norm(p));
                for (const r of Array.from(document.querySelectorAll('tr'))) {
                  try {
                    const txt = norm(r.innerText || '');
                    if (!txt) continue;
                    if (needles.includes(txt)) {
                      const cb = r.querySelector('input[type="checkbox"]');
                      if (cb && !cb.checked) {
                        try { cb.click(); } catch (e) { try { cb.checked = true; } catch(_){} }
                        try { cb.dispatchEvent(new Event('change', { bubbles: true })); } catch(_){}
                      }
                    }
                  } catch (e) {}
                }
              } catch (e) {}
            }, selectedSoFar).catch(() => {});
            // after reapplying, log how many rows are checked now
            try {
              const rechecked = await modalFrame.evaluate(() => {
                try { return Array.from(document.querySelectorAll('tr')).filter(r => { try { const cb = r.querySelector('input[type="checkbox"]'); return !!(cb && cb.checked); } catch(e){return false;} }).map(r => r.innerText.trim()).slice(0,200); } catch (e) { return []; }
              }).catch(() => []);
              logEvent({ level: 'debug', message: 'After reapply selectedSoFar, modal checked rows count', detail: JSON.stringify({ count: (rechecked||[]).length, sample: (rechecked||[]).slice(0,20) }), modalTitle });
            } catch (_) {}
          } catch (_) {}
        }

        // Try quick DOM-only selection first to avoid triggering searches that may re-render
        // and clear previously checked items. If quick selection succeeds, skip typing.
        const qnorm = normalizeSelectionText(currentValue);
        // ensure virtualized rows are rendered by scrolling the modal briefly
        try {
          await modalFrame.evaluate(() => {
            const scroller = document.querySelector('.ui-table') || document.querySelector('.modal-body') || document.scrollingElement || document.documentElement;
            if (!scroller) return;
            const height = scroller.scrollHeight || (document.body && document.body.scrollHeight) || 0;
            let y = 0;
            while (y < height) {
              scroller.scrollTop = y;
              y += 800;
            }
            scroller.scrollTop = height;
          }).catch(() => {});
        } catch (_) {}

        const quickMarked = await modalFrame.evaluate(({ needle, rowCheckboxSel }) => {
          try {
            const normalize = s => {
              try { return String(s || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/[^\w\s]/g, ' ').replace(/\s+/g, ' ').trim().toLowerCase(); } catch (e) { return String(s || '').replace(/[^\w\s]/g, ' ').replace(/\s+/g, ' ').trim().toLowerCase(); }
            };
            for (const r of document.querySelectorAll('tr')) {
              const rowText = normalize(r.innerText || '');
              const cells = Array.from(r.querySelectorAll('td,th')).map(c => normalize(c.innerText || ''));
              const ok = rowText === needle || cells.some(ct => ct === needle);
              if (ok) {
                const cb = r.querySelector(rowCheckboxSel || 'input[type="checkbox"]');
                if (cb && !cb.checked) {
                  try { cb.click(); } catch (e) { try { cb.checked = true; } catch(_){} }
                  try { cb.dispatchEvent(new Event('change', { bubbles: true })); } catch(_){ }
                }
                return !!(cb && cb.checked);
              }
            }
          } catch (e) {}
          return false;
        }, { needle: qnorm, rowCheckboxSel: modalDefinition.rowCheckboxSelector }).catch(() => false);

        if (quickMarked) {
          logEvent({ level: 'info', message: `Entry quick-mark succeeded: ${currentValue}`, modalTitle });
          succeeded = true;
          break;
        }

        // limpa e digita
          try {
          const beforeInputVal = await modalInputLocal.inputValue().catch(() => '');
          logEvent({ level: 'debug', message: 'Modal input before clear', detail: JSON.stringify({ beforeInputVal }), modalTitle });
        } catch (_) {}
        await modalInputLocal.fill('').catch(() => {});
        await page.waitForTimeout(300);
        const inputValue = await modalInputLocal.inputValue().catch(() => '');
        if (inputValue && inputValue.trim()) {
          await modalInputLocal.evaluate((el) => { el.value = ''; el.dispatchEvent(new Event('input', { bubbles: true })); }).catch(() => {});
          await page.waitForTimeout(200);
        }

        await fillLegacyInput(modalInputLocal, currentValue);
        logEvent({ level: 'debug', message: 'Typed into modal input', detail: JSON.stringify({ value: String(currentValue).slice(0,200) }), modalTitle });
        try { await modalInputLocal.focus(); } catch (_) {}
        await modalInputLocal.press('Enter').catch(() => {});

        // aciona procurar
        const clicked = await modalFrame.evaluate(() => {
          try {
            const nodes = Array.from(document.querySelectorAll('input,button,a'));
            for (const n of nodes) {
              const txt = ((n.value || '') + ' ' + (n.innerText || '')).trim();
              if (/procurar/i.test(txt)) { try { n.click(); } catch(e){ try{ n.dispatchEvent(new MouseEvent('click',{bubbles:true})); }catch(_){} } return true; }
            }
          } catch (e) {}
          return false;
        }).catch(() => false);

        if (!clicked) {
          const procurarBtn = modalFrame.locator('#pbProcurar, input[name="pbProcurar"], input[id*="Procurar"], input[value*="Procurar"]');
          if (await procurarBtn.count().catch(() => 0)) {
            await procurarBtn.first().click({ force: true }).catch(() => {});
          }
        }

        try {
          let normalizedToken;
          try {
            normalizedToken = normalizeSelectionText(currentValue);
          } catch (e) {
            normalizedToken = normalizeSelectionText(currentValue);
          }

          await modalFrame.waitForFunction((needle) => {
            try {
              function norm(s) {
                return String(s || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/[^\w\s]/g, ' ').replace(/\s+/g, ' ').trim().toLowerCase();
              }
              return document.body && Array.from(String(document.body.innerText || '').split(/\r?\n/)).map(norm).some(line => line === needle);
            } catch (e) {
              return false;
            }
          }, { timeout: 10000 }, normalizedToken).catch(() => {});
        } catch (_) {}

        // Re-apply previously selected rows AFTER search results load to avoid the search
        // clearing previously-checked items (important for multi-select scenarios).
        try {
          if (selectedSoFar && selectedSoFar.length) {
            await modalFrame.evaluate((prev) => {
              try {
                const norm = s => String(s || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/[^\w\s]/g, ' ').replace(/\s+/g, ' ').trim().toLowerCase();
                const needles = prev.map(p => norm(p));
                for (const r of Array.from(document.querySelectorAll('tr'))) {
                  try {
                    const txt = norm(r.innerText || '');
                    if (!txt) continue;
                    for (const needle of needles) {
                      if (!needle) continue;
                      if (txt === needle) {
                        const cb = r.querySelector('input[type="checkbox"]');
                        if (cb && !cb.checked) {
                          try { cb.click(); } catch (e) { try { cb.checked = true; } catch(_){} }
                          try { cb.dispatchEvent(new Event('change', { bubbles: true })); } catch(_){ }
                        }
                        break;
                      }
                    }
                  } catch (e) {}
                }
              } catch (e) {}
            }, selectedSoFar).catch(() => {});
            try {
              const rechecked = await modalFrame.evaluate(() => {
                try { return Array.from(document.querySelectorAll('tr')).filter(r => { try { const cb = r.querySelector('input[type="checkbox"]'); return !!(cb && cb.checked); } catch(e){return false;} }).map(r => r.innerText.trim()).slice(0,200); } catch (e) { return []; }
              }).catch(() => []);
              logEvent({ level: 'debug', message: 'After reapply-post-search selectedSoFar, modal checked rows count', detail: JSON.stringify({ count: (rechecked||[]).length, sample: (rechecked||[]).slice(0,20) }), modalTitle });
            } catch (_) {}
          }
        } catch (_) {}

        // tenta encontrar a linha e marcar o checkbox com frame-retry
        let rowLocator = null;
        try {
          rowLocator = await withFrameRetry(async () => await findSelectionRow(modalFrame, currentValue, modalDefinition.rowValueSelectors, 25000), { name: 'findSelectionRow', retries: 2 });
        } catch (err) {
          await saveDebugState('row-not-found').catch(() => {});
          logEvent({ level: 'warn', message: `Linha não encontrada no modal, tentativa=${attempt}, pulando: ${currentValue}`, detail: String(err && err.message || err), modalTitle });
          // se foi a última tentativa, efetivamente pula
          if (attempt >= 3) break;
          await page.waitForTimeout(300 * attempt);
          continue;
        }

        const selPromise = async () => await ensureCheckboxChecked(rowLocator, modalDefinition.rowCheckboxSelector);
        const selOk = await Promise.race([withFrameRetry(selPromise, { name: `ensureCheckboxChecked:${currentValue}`, retries: 2 }), new Promise(res => setTimeout(() => res(false), MODAL_SELECT_TIMEOUT_MS))]);
        if (!selOk) {
          await saveDebugState('checkbox-timeout').catch(() => {});
          logEvent({ level: 'warn', message: `Timeout ao selecionar valor no modal, tentativa=${attempt}: ${currentValue}`, modalTitle });
          // tenta reabrir modal e refazer se houver tentativas restantes
          try {
            logEvent({ level: 'info', message: 'Attempting to close and reopen modal for retry', detail: { attempt, currentValue }, modalTitle });
            try { await modalFrame.locator('#pbFechar, input[id*="pbFechar"]').first().click({ force: true }); } catch (_) {}
            await page.waitForTimeout(400);
            await trigger.click({ force: true }).catch(() => {});
            modalFrame = await findSelectionModalFrame(page, modalTitle, 15000).catch(() => null) || await findModalFrame(page, modalTitle, 15000).catch(() => null);
            logEvent({ level: 'info', message: 'Modal re-open attempt complete', detail: { reopened: !!modalFrame }, modalTitle });
            if (!modalFrame) {
              logEvent({ level: 'warn', message: `Não foi possível reabrir modal para retry, tentativa=${attempt}, pulando: ${currentValue}`, modalTitle });
              if (attempt >= 3) break;
              await page.waitForTimeout(300 * attempt);
              continue;
            }
            // se reabriu, faz mais uma iteração do loop
            await page.waitForTimeout(300);
            continue;
          } catch (e) {
            logEvent({ level: 'warn', message: `Erro durante retry para ${currentValue}: ${String(e && e.message || e)}`, modalTitle });
            if (attempt >= 3) break;
            await page.waitForTimeout(300 * attempt);
            continue;
          }
        }

        logEvent({ level: 'info', message: `Checkbox marcado: ${currentValue}`, modalTitle });
        try { logEvent({ level: 'debug', message: 'selectedSoFar before update', detail: JSON.stringify(selectedSoFar.slice(0,50)), modalTitle }); } catch (_) {}
        // Persist selection immediately into main frame to avoid losing it when modal changes
        try {
          const modalCheckedNow = await modalFrame.evaluate(() => Array.from(document.querySelectorAll('tr')).filter(r => { try { const cb = r.querySelector('input[type="checkbox"]'); return !!(cb && cb.checked); } catch(e){return false;} }).map(r => r.innerText.trim()).slice(0,200)).catch(() => []);
          try {
            const iframeLocator2 = page.locator('iframe#iFramePage');
            const mainFrame2 = await iframeLocator2.contentFrame().catch(() => null);
            if (mainFrame2 && modalCheckedNow && modalCheckedNow.length) {
              await writeSelectedEntitiesIntoMainFrame(mainFrame2, modalCheckedNow).catch(() => {});
              logEvent({ level: 'debug', message: 'Persisted modal selections into main frame (per-item)', detail: JSON.stringify(modalCheckedNow.slice(0,10)), modalTitle });
            }
          } catch (e) {}
            // atualiza acumulador local de seleções sem duplicatas
            try {
              for (const m of (modalCheckedNow || [])) {
                if (!selectedSoFar.includes(m)) selectedSoFar.push(m);
              }
              logEvent({ level: 'debug', message: 'selectedSoFar updated', detail: JSON.stringify({ count: selectedSoFar.length, sample: selectedSoFar.slice(0,50) }), modalTitle });
            } catch (_) {}
        } catch (e) {}
        succeeded = true;
      } catch (err) {
        const msg = String(err && (err.message || err));
        logEvent({ level: 'warn', message: `Erro ao selecionar '${currentValue}', tentativa=${attempt}`, detail: msg, modalTitle });
        if (/Frame was detached|Execution context was destroyed|Cannot find context|Cannot find object|Target closed|Modal frame desapareceu/i.test(msg) && attempt < 3) {
          // tenta re-resolver modal frame e repetir
          modalFrame = await findSelectionModalFrame(page, modalTitle, 5000).catch(() => null) || await findModalFrame(page, modalTitle, 5000).catch(() => null);
          await page.waitForTimeout(500 * attempt);
          continue;
        }
        // se for erro irreversível, rethrow
        if (attempt >= 3) throw err;
      }
    }

    if (!succeeded) {
      logEvent({ level: 'warn', message: `Falha ao selecionar após tentativas: ${currentValue}`, modalTitle });
      continue;
    }
    saveShot(page, `checkbox-${sanitizeFileName(currentValue)}`).catch(() => {});
    await page.waitForTimeout(700);
    // don't clear the modal input here; clearing can reset previous selections in some modals
    // keep input as-is and move to next iteration which will set new value explicitly
    await page.waitForTimeout(200);

    // Log main frame selected lists after this iteration to help understand persistence
    try {
      const iframeLocator2 = page.locator('iframe#iFramePage');
      const mainFrame2 = await iframeLocator2.contentFrame().catch(() => null);
      if (mainFrame2) {
        const selLists = await mainFrame2.evaluate(() => {
          try {
            return Array.from(document.querySelectorAll('[id$="SelectedEntitiesList"], [class*="SelectedEntitiesList"]')).map(el => (el.value || el.innerText || '').toString()).slice(0,20);
          } catch (e) { return []; }
        }).catch(() => []);
        logEvent({ level: 'debug', message: 'Main frame SelectedEntitiesList after iteration', detail: { count: selLists.length, sample: selLists.slice(0,10) }, modalTitle });
      }
    } catch (_) {}
  }
  
  // Antes de confirmar, re-aplica todas as seleções acumuladas no modal
  try {
    let finalModal = await findSelectionModalFrame(page, modalTitle, 2000).catch(() => null) || await findModalFrame(page, modalTitle, 2000).catch(() => null);
    if (finalModal && selectedSoFar && selectedSoFar.length) {
      try {
        await finalModal.evaluate((prev) => {
          try {
            const norm = s => String(s || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/[^\w\s]/g, ' ').replace(/\s+/g, ' ').trim().toLowerCase();
            const needles = Array.from(new Set(prev.map(p => norm(p))));
            for (const r of Array.from(document.querySelectorAll('tr'))) {
              try {
                const txt = norm(r.innerText || '');
                if (!txt) continue;
                for (const needle of needles) {
                  if (!needle) continue;
                  if (txt === needle) {
                    const cb = r.querySelector('input[type="checkbox"]');
                    if (cb && !cb.checked) {
                      try { cb.click(); } catch (e) { try { cb.checked = true; } catch(_){} }
                      try { cb.dispatchEvent(new Event('change', { bubbles: true })); } catch(_){ }
                    }
                    break;
                  }
                }
              } catch (e) {}
            }
          } catch (e) {}
        }, selectedSoFar).catch(() => {});
        await page.waitForTimeout(300);
      } catch (_) {}
    }
  } catch (_) {}

    // --- DEBUG: dump modal state (HTML + diagnostic JSON) before clicking Selecionar
    try {
      if (DEBUG_HTML) {
        try {
          await fs.promises.mkdir(SCREENSHOT_DIR, { recursive: true }).catch(() => {});
          const ts = nowIso().replace(/[:.]/g, '-');
          const base = sanitizeFileName(`${modalTitle}-preselect-${ts}`);
          const htmlPath = path.join(SCREENSHOT_DIR, `${base}.html`);
          const diagPath = path.join(SCREENSHOT_DIR, `${base}-diag.json`);
          try {
            const modalHtml = await finalModal.content().catch(() => '<!-- erro ao capturar -->');
            await fs.promises.writeFile(htmlPath, modalHtml || '<!-- vazio -->').catch(() => {});
          } catch (e) {}
          try {
            const diag = await finalModal.evaluate(() => {
              try {
                const norm = s => String(s || '').trim();
                const checkedRows = Array.from(document.querySelectorAll('tr')).filter(r => { try { const cb = r.querySelector('input[type="checkbox"]'); return !!(cb && cb.checked); } catch(e){return false;} }).map(r => ({ text: norm(r.innerText), html: r.innerHTML.slice(0,1000) }));
                const inputs = Array.from(document.querySelectorAll('input,textarea,select')).map(i => ({ tag: i.tagName, type: i.type, id: i.id, name: i.name, value: i.value, display: (i.style && i.style.display) || getComputedStyle(i).display }));
                const hidden = inputs.filter(i => i.type === 'hidden' || i.display === 'none');
                const selectedEntities = Array.from(document.querySelectorAll('[id$="SelectedEntitiesList"],[class*="SelectedEntitiesList"]')).map(el => ({ id: el.id || null, class: el.className || null, tag: el.tagName, value: el.value || el.innerText || '' }));
                const dataAttrs = Array.from(document.querySelectorAll('*')).filter(el => el && el.getAttribute && Array.from(el.attributes).some(a => a.name.indexOf('data-')===0)).slice(0,200).map(el => ({ tag: el.tagName, id: el.id||null, attrs: Array.from(el.attributes).filter(a => a.name.indexOf('data-')===0).map(a=>({name:a.name,value:a.value})) }));
                return { checkedRows, inputs, hidden, selectedEntities, dataAttrs };
              } catch (e) { return { error: String(e && e.message || e) }; }
            }).catch(() => ({ error: 'evaluate-failed' }));
            await fs.promises.writeFile(diagPath, JSON.stringify(diag, null, 2)).catch(() => {});
            logEvent({ level: 'info', message: `Saved modal debug HTML+diag`, detail: { htmlPath, diagPath }, modalTitle });
          } catch (e) {}
        } catch (e) {}
      }
    } catch (e) {}

    // Clique no botão Selecionar preferencialmente dentro do frame do modal
    let clickedSelec = false;
    try {
      const finalModalFrame = finalModal || await findSelectionModalFrame(page, modalTitle, 2000).catch(() => null) || await findModalFrame(page, modalTitle, 2000).catch(() => null);
      if (finalModalFrame) {
        const selBtn = finalModalFrame.locator('#pbSelecionar, input[id*="pbSelecionar"], input[name="pbSelecionar"], button#pbSelecionar').first();
        if (await selBtn.count().catch(() => 0)) {
          await selBtn.click({ force: true }).catch(() => {});
          clickedSelec = true;
        }
      }
    } catch (_) {}

    if (!clickedSelec) {
      const { locator: selecionarBtn } = await findVisibleLocatorInFrames(page, '#pbSelecionar');
      await selecionarBtn.click({ force: true });
    }
  
  await page.waitForTimeout(2000);
  await page.waitForTimeout(2000);
  
  logEvent({
    level: 'info',
    message: `Modal finalizado: ${entries.join(', ')}`,
    modalTitle,
  });
  try { await saveShot(page, `modal-selected2-${sanitizeFileName(modalTitle)}`); } catch (_) {}
  } finally {
    if (typeof _modalLockHeld !== 'undefined' && _modalLockHeld) {
      try { releaseModalLock(); } catch (_) {}
      _modalLockHeld = false;
    }
  }
}

selectViaModal = selectViaModalByGrid;

async function writeSelectedEntitiesIntoMainFrame(frame, values, config = {}) {
  const {
    inputName,
    entityName,
    hiddenMappings = {},
  } = config;

  if (!inputName || !entityName) {
    throw new Error('writeSelectedEntitiesIntoMainFrame requires inputName and entityName');
  }

  return setInputSelectState(frame, {
    inputName,
    entityName,
    values,
    hiddenMappings,
  });
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
      let el = document.getElementById(id);
      if (!el) {
        el = document.querySelector(`[name="${id}"]`);
      }
      
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

async function downloadPdf(
  context,
  reportUrl,
  outputPath
) {
  
  if (!reportUrl) {
    
    throw new Error(
      'downloadPdf: reportUrl inválida.'
    );
  }
  
  logEvent({
    level: 'info',
    message: 'Iniciando download PDF.',
    reportUrl,
    outputPath,
  });
  
  // =====================================================
  // REQUEST CONTEXT
  // =====================================================
  const response =
  await context.request.get(
    reportUrl,
    {
      timeout: 120000,
    }
  );
  
  if (!response.ok()) {
    
    throw new Error(
      `Falha download PDF. Status: ${response.status()}`
    );
  }
  
  const body =
  await response.body();
  
  // =====================================================
  // GARANTE DIRETÓRIO
  // =====================================================
  await fs.promises.mkdir(
    path.dirname(outputPath),
    {
      recursive: true,
    }
  );
  
  // =====================================================
  // SALVA
  // =====================================================
  await fs.promises.writeFile(
    outputPath,
    body
  );
  
  // =====================================================
  // VALIDA
  // =====================================================
  const stat =
  await fs.promises.stat(outputPath);
  
  if (!stat.size) {
    
    throw new Error(
      'Arquivo PDF salvo vazio.'
    );
  }
  
  logEvent({
    level: 'info',
    message: 'PDF salvo.',
    outputPath,
    size: stat.size,
  });
  
  return outputPath;
}

function extractFinalReportUrl(reportUrl) {
  
  if (!reportUrl) {
    return null;
  }
  
  try {
    
    let current = reportUrl;
    
    for (let i = 0; i < 5; i++) {
      
      const parsed =
      new URL(current);
      
      // =====================================================
      // URL ENCAPSULADA
      // =====================================================
      const inner =
      parsed.searchParams.get('url');
      
      if (inner) {
        
        current = new URL(
          decodeURIComponent(inner),
          parsed.origin
        ).toString();
        
        continue;
      }
      
      // =====================================================
      // URL RELATÓRIO
      // =====================================================
      if (
        /viewReportSPW|gerarRelatorio|viewer|\.pdf/i
        .test(current)
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
      'Falha ao extrair URL final.',
      reportUrl,
      error: String(err),
    });
    
    return reportUrl;
  }
}

// =====================================================
// RANGE MÊS ANTERIOR
// =====================================================
function getPreviousMonthRange() {
  
  const now = new Date();
  
  // primeiro dia do mês atual
  const firstDayCurrentMonth =
  new Date(
    now.getFullYear(),
    now.getMonth(),
    1
  );
  
  // último dia do mês anterior
  const lastDayPreviousMonth =
  new Date(
    firstDayCurrentMonth.getTime() - 1
  );
  
  // primeiro dia do mês anterior
  const firstDayPreviousMonth =
  new Date(
    lastDayPreviousMonth.getFullYear(),
    lastDayPreviousMonth.getMonth(),
    1
  );
  
  const format = date => {
    
    const day =
    String(date.getDate())
    .padStart(2, '0');
    
    const month =
    String(date.getMonth() + 1)
    .padStart(2, '0');
    
    const year =
    date.getFullYear();
    
    return `${day}/${month}/${year}`;
  };
  
  return {
    inicio: format(firstDayPreviousMonth),
    fim: format(lastDayPreviousMonth),
  };
}

async function generateSingleReport(
  context,
  page,
  report
) {
  
  logEvent({
    level: 'info',
    message: 'Entrou em generateSingleReport',
    report: report.sheetName,
  });
  
  // PERFIL: marca início da geração do relatório para o profiler externo
  try { console.log(`STEP START generateSingleReport ${report.sheetName}`); } catch (e) {}
  
  const __genMeasure = startMeasure('generateSingleReport', { report: report.sheetName });
  
  try {
    // Isola cache de modais e estado por relatório para evitar cross-talk
    try { MODAL_CACHE = {}; } catch (_) {}
    
    await openReportsPage(page);
    await saveShot(page, `reports-opened-${sanitizeFileName(report.sheetName)}`).catch(() => {});
    
    // =====================================================
    // CONFIGURA FILTROS
    // =====================================================
    await configureReportFilters({
      page,
      report
    });
    await saveShot(page, `filters-configured-${sanitizeFileName(report.sheetName)}`).catch(() => {});
    
    const reportFrame =
    await getReportFilterFrame(page);
    
    await ensureSubmitEnabled(reportFrame);
    
    await closeLegacyPopups(page)
    .catch(() => {});
    
    await page.waitForTimeout(2000);
    
    const finalPdfPath = path.join(
      REPORT_OUTPUT_DIR,
      `${sanitizeFileName(report.pdfName)}.pdf`
    );
    
    // =====================================================
    // PROMISE NOVA JANELA
    // =====================================================
    // =====================================================
    // REGEX RELATÓRIO
    // =====================================================
    
    const REPORT_REGEX =
    /viewReportSPW|please_wait_frame|ReportServlet|jasper|viewer|\.pdf/i;
    
    let reportUrl = null;
    
    // =====================================================
    // MONITOR REQUEST/RESPONSE
    // =====================================================
    
    let detectedReportUrl = null;
    
    const requestListener = request => {
      
      try {
        
        const url = request.url();
        
        if (
          REPORT_REGEX.test(url)
        ) {
          
          detectedReportUrl = url;
          
          logEvent({
            level: 'info',
            message: 'Request relatório detectada.',
            url,
          });
        }
        
      } catch {}
    };
    
    const responseListener = response => {
      
      try {
        
        const url =
        response.url();
        
        const contentType =
        response.headers()['content-type'] || '';
        
        if (
          REPORT_REGEX.test(url) ||
          contentType.includes('pdf')
        ) {
          
          detectedReportUrl = url;
          
          logEvent({
            level: 'info',
            message: 'Response relatório detectada.',
            url,
            contentType,
          });
        }
        
      } catch {}
    };
    
    context.on(
      'request',
      requestListener
    );
    
    context.on(
      'response',
      responseListener
    );
    
    // =====================================================
    // POPUP REAL
    // =====================================================
    
    const popupPromise =
    page.waitForEvent(
      'popup',
      {
        timeout: 60000,
      }
    ).catch(() => null);
    
    // =====================================================
    // CLICK VISUALIZAR
    // =====================================================
    
    logEvent({ level: 'info', message: `Relatório "${report.sheetName}": clicando Visualizar.`, });
    console.log(`Relatório "${report.sheetName}": acionando Visualizar (aguardando popup)...`);
    
    await saveShot(page, `visualizar_click-${sanitizeFileName(report.sheetName)}`).catch(() => {});
    const btnFiltrar =
    reportFrame.locator(
      '#btFiltrar, input[name="btFiltrar"]'
    ).first();
    
    await btnFiltrar.waitFor({
      state: 'visible',
      timeout: 30000,
    });
    
    await btnFiltrar.scrollIntoViewIfNeeded()
    .catch(() => {});
    
    await btnFiltrar.click({
      force: true,
      timeout: 30000,
    });
    
    try { await setReportProgress(report.sheetName, 5, { lastMessage: 'view_clicked', step: 'view', stepStartedAt: nowIso() }); } catch (_) {}
    
    // =====================================================
    // AGUARDA POPUP
    // =====================================================
    
    logEvent({
      level: 'info',
      message:
      `Relatório "${report.sheetName}": aguardando popup.`,
    });
    
    const popup =
    await popupPromise;
    
    if (!popup) {
      
      logEvent({
        level: 'info',
        message:
        `Relatório "${report.sheetName}": não retornou nenhuma informação.`,
      });
      return ;
    }
    
    // report: popup detected -> update progress
    try {
      await updateReportStatusSerialized({ perReportEntry: { report: report.sheetName, props: { progress: 10, lastMessage: 'popup_detected', step: 'popup', stepStartedAt: nowIso() } } });
    } catch (_) {}
    
    
    // =====================================================
    // ESPERA CARREGAMENTO REAL
    // =====================================================
    
    await popup.waitForLoadState(
      'domcontentloaded',
      {
        timeout: 120000,
      }
    );
    
    await popup.waitForTimeout(
      12000
    );
    // debug: captura do popup assim que carregado
    try { await saveShot(popup, `popup-loaded-${sanitizeFileName(report.sheetName)}`); } catch (_) {}
    
    try { await setReportProgress(report.sheetName, 30, { lastMessage: 'popup_loaded', step: 'popup_loaded' }); } catch (_) {}
    
    // =====================================================
    // LOOP URL FINAL
    // =====================================================
    // report: about to look for final URL
    try {
      await setReportProgress(report.sheetName, 35, { lastMessage: 'waiting_for_url', step: 'url_wait', stepStartedAt: nowIso() });
    } catch (_) {}
    
    const start =
    Date.now();
    
    while (
      Date.now() - start < 120000
    ) {
      
      const urls = [];
      
      try {
        
        urls.push(
          popup.url()
        );
        
      } catch {}
      
      for (const frame of popup.frames()) {
        
        try {
          
          urls.push(
            frame.url()
          );
          
        } catch {}
      }
      
      // remove vazias
      const validUrls =
      urls.filter(Boolean);
      
      // LOG DEBUG
      logEvent({
        level: 'info',
        message:
        'URLs popup monitoradas.',
        urls: validUrls,
      });
      
      // procura url final
      const found =
      validUrls.find(url =>
        
        /ReportServlet|viewReportSPW|viewer|jasper|\.pdf/i
        .test(url)
        
      );
      
      if (found) {
        
        try { await setReportProgress(report.sheetName, 50, { lastMessage: 'found_final_url', step: 'url_found' }); } catch (_) {}
        
        reportUrl = found;
        
        break;
      }
      
      // ===================================================
      // TENTA EXTRAIR DO PLEASE WAIT
      // ===================================================
      
      const pleaseWait =
      validUrls.find(url =>
        url.includes(
          'please_wait_frame.jsp'
        )
      );
      
      if (pleaseWait) {
        
        try {
          
          const parsed =
          new URL(pleaseWait);
          
          const internalUrl =
          parsed.searchParams.get(
            'url'
          );
          
          if (internalUrl) {
            
            try { await setReportProgress(report.sheetName, 50, { lastMessage: 'found_internal_url', step: 'url_found' }); } catch (_) {}
            
            reportUrl =
            internalUrl.startsWith('http')
            ? internalUrl
            : `${BASE_URL}${internalUrl}`;
            
            break;
          }
          
        } catch {}
      }
      
      await popup.waitForTimeout(
        1500
      );
    }
    
    // =====================================================
    // FALLBACK REQUESTS
    // =====================================================
    
    if (
      !reportUrl &&
      detectedReportUrl
    ) {
      
      reportUrl =
      detectedReportUrl;
    }
    
    reportUrl =
    extractFinalReportUrl(reportUrl);
    
    logEvent({
      level: 'info',
      message: 'URL final do relatório resolvida.',
      reportUrl,
    });
    try { await saveShot(popup, `reporturl-resolved-${sanitizeFileName(report.sheetName)}`); } catch (_) {}
    
    // report: URL resolved
    try {
      await updateReportStatusSerialized({ perReportEntry: { report: report.sheetName, props: { progress: 60, lastMessage: 'report_url_resolved', step: 'url', stepStartedAt: nowIso() } } });
    } catch (_) {}
    
    // =====================================================
    // LIMPEZA
    // =====================================================
    
    context.off(
      'request',
      requestListener
    );
    
    context.off(
      'response',
      responseListener
    );
    
    // =====================================================
    // DEBUG FINAL
    // =====================================================
    
    logEvent({
      level: 'info',
      message:
      'Resultado captura relatório.',
      reportUrl,
      popupDetected: !!popup,
    });
    
    // =====================================================
    // ERRO
    // =====================================================
    
    if (!reportUrl) {
      
      const html =
      await popup.content()
      .catch(() => '');
      
      try { await saveShot(popup, `error-no-url-${sanitizeFileName(report.sheetName)}`); } catch (_) {}
      console.log(
        html.substring(0, 5000)
      );
      
      throw new Error(
        `Não foi possível localizar URL final do relatório "${report.sheetName}".`
      );
    }
    
    // =====================================================
    // DOWNLOAD PDF
    // =====================================================
    
    logEvent({
      level: 'info',
      message:
      `Baixando PDF "${report.sheetName}".`,
      reportUrl,
    });
    console.log(`Relatório "${report.sheetName}": iniciando download do PDF...`);
    
    await saveShot(popup, `before-download-${sanitizeFileName(report.sheetName)}`).catch(() => {});
    try {
      await updateReportStatusSerialized({ perReportEntry: { report: report.sheetName, props: { progress: 70, lastMessage: 'downloading', step: 'download', stepStartedAt: nowIso() } } });
    } catch (_) {}
    
    const response =
    await context.request.get(
      reportUrl,
      {
        timeout: 120000,
      }
    );
    
    try { await setReportProgress(report.sheetName, 75, { lastMessage: 'download_request_sent' }); } catch (_) {}
    
    const contentType =
    String(
      response.headers()['content-type'] || ''
    ).toLowerCase();
    
    const buffer =
    await response.body();
    
    try { await setReportProgress(report.sheetName, 85, { lastMessage: 'downloaded' }); } catch (_) {}
    
    const head =
    buffer.subarray(0, 8).toString('utf8');
    
    if (
      !contentType.includes('application/pdf') &&
      !head.startsWith('%PDF')
    ) {
      
      const snippet =
      buffer.toString('utf8', 0, Math.min(buffer.length, 500));
      
      throw new Error(
        `Resposta não é PDF. content-type=${contentType || 'n/a'}; head=${JSON.stringify(head)}; snippet=${JSON.stringify(snippet)}`
      );
    }
    
    if (fs.existsSync(finalPdfPath)) {
      try {
        fs.unlinkSync(finalPdfPath);
      } catch {};
    }
    
    await fs.promises.writeFile(finalPdfPath, buffer).catch((e) => { throw e; });
    
    try { await setReportProgress(report.sheetName, 95, { lastMessage: 'saved_tmp', step: 'saving' }); } catch (_) {}
    // captura final após salvar PDF
    try { await saveShot(page, `pdf-saved-${sanitizeFileName(report.sheetName)}`); } catch (_) {}
    
    const stats =
    fs.statSync(finalPdfPath);
    logEvent({ level: 'info', message: 'PDF salvo.', bytes: stats.size, path: finalPdfPath });
    console.log(`Relatório "${report.sheetName}" salvo em: ${finalPdfPath} (${stats.size} bytes)`);
    
    // report: saved
    try {
      await updateReportStatusSerialized({ perReportEntry: { report: report.sheetName, props: { progress: 100, lastMessage: 'saved', step: 'finished', finishedAt: nowIso(), processedItems: 1, totalItems: 1 } } });
    } catch (_) {}
    
    if (stats.size < 5000) {
      try { await saveShot(page, `pdf-invalid-${sanitizeFileName(report.sheetName)}-${stats.size}`); } catch (_) {}
      throw new Error(`PDF inválido (${stats.size} bytes).`);
    }
    
    await popup.close().catch(() => {});
    
    logEvent({ level: 'info', message: `Relatório "${report.sheetName}" salvo com sucesso.`, path: finalPdfPath });
    try { console.log(`STEP END generateSingleReport ${report.sheetName}`); } catch (e) {}
    
    return { path: finalPdfPath, bytes: stats.size };
  } finally {
    await stopMeasure(__genMeasure, { phase: 'generateSingleReport' }).catch(() => {});
  }
}

async function closeLegacyPopups(page) {
  
  try {
    
    // modal jquery ui
    const closeButtons = page.locator(`
      .ui-dialog-titlebar-close,
      .popup-close,
      .modal-close,
      img[title="Fechar"],
      input[value="FECHAR"],
      input[value="Fechar"]
    `);
      
      const count =
      await closeButtons.count()
      .catch(() => 0);
      
      for (let i = 0; i < count; i++) {
        
        await closeButtons
        .nth(i)
        .click({
          force: true,
        })
        .catch(() => {});
      }
      
      // overlays
      await page.evaluate(() => {
        
        document
        .querySelectorAll(`
          .ui-widget-overlay,
          .modal-backdrop
        `)
          .forEach(el => el.remove());
          
        });
        
      } catch {}
    }
    
    async function runReports(context, page) {
      
      await ensureDir(REPORT_OUTPUT_DIR);
      
      // PERFIL: marca início da rotina de geração de relatórios
      try { console.log('STEP START runReports'); } catch (e) {}
      
      page.setDefaultTimeout(15000);
      page.setDefaultNavigationTimeout(30000);
      
      let reports = REPORT_DEFINITIONS.map(report => ({ ...report }));
      
      // Suporta modo de execução de um único relatório via env `SINGLE_REPORT` ou CLI `--single <index|name>`
      if (SINGLE_REPORT_ARG) {
        const arg = String(SINGLE_REPORT_ARG).trim();
        let selected = null;
        if (/^\d+$/.test(arg)) {
          const idx = Math.max(0, Number(arg) - 1);
          if (idx >= 0 && idx < reports.length) selected = reports[idx];
        } else {
          selected = reports.find(r => r.sheetName.toLowerCase() === arg.toLowerCase());
        }
        
        if (selected) {
          logEvent({ level: 'info', message: `Modo single report ativado: ${arg}` });
          reports = [selected];
        } else {
          logEvent({ level: 'warn', message: `Single report não encontrado: ${arg}. Executando todos.` });
        }
      }
      
      const startedAt = new Date().toISOString();
      
      const status = {
        running: true,
        startedAt,
        finishedAt: null,
        total: reports.length,
        current: 0,
        progress: 0,
        currentReport: null,
        completedReports: [],
        error: null,
      };
      
      // Inicializa entradas por relatório como 'pending' e grava status inicial
      for (const r of reports) {
        try {
          await updateReportStatusSerialized({ perReportEntry: { report: r.sheetName, props: { status: 'pending', createdAt: nowIso(), lastUpdated: nowIso(), progress: 0, processedItems: 0, totalItems: null } } });
        } catch {}
      }
      
      // STATUS INICIAL
      await updateReportStatusSerialized(status);
      
      logEvent({
        level: 'info',
        message: 'Iniciando geração de relatórios.',
        reports: reports.map(report => report.sheetName),
        outputDir: REPORT_OUTPUT_DIR,
      });
      
      // Abra todas as abas primeiro, depois gere os relatórios em paralelo
      const entries = [];
      for (let index = 0; index < reports.length; index++) {
        const report = reports[index];
        console.log(`Criando aba para relatório ${index + 1}/${reports.length}: ${report.sheetName}`);
        const reportPage = await context.newPage();
        reportPage.setDefaultTimeout(15000);
        reportPage.setDefaultNavigationTimeout(30000);
        await attachPageDebug(reportPage);
        console.log(`Aba criada: ${report.sheetName}`);
        entries.push({ report, page: reportPage, index });
      }
      
      const tasks = entries.map(entry => (async () => {
        const { report, page, index } = entry;
        const reportName = report.sheetName;
        const startedAtReport = new Date().toISOString();
        try {
          // marca início — coloque o relatório no passo 'initializing' com 2% de progresso
          await updateReportStatusSerialized({ perReportEntry: { report: reportName, props: { status: 'running', startedAt: startedAtReport, index: index + 1, step: 'initializing', stepStartedAt: nowIso(), progress: 2, lastMessage: 'initializing' } } });
          logEvent({ level: 'info', message: 'Relatório iniciado.', report: reportName, current: index + 1, total: reports.length });
          
          const result = await generateSingleReport(context, page, report);
          
          const finishedAt = new Date().toISOString();
          
          await updateReportStatusSerialized({ perReportEntry: { report: reportName, props: { status: 'completed', finishedAt, path: result && result.path ? result.path : (typeof result === 'string' ? result : null), bytes: result && result.bytes ? result.bytes : null, progress: 100, lastMessage: 'completed' } } });
          logEvent({ level: 'info', message: 'Relatório concluído.', report: reportName, path: result && result.path ? result.path : null, bytes: result && result.bytes ? result.bytes : null });
          
          return { report: reportName, result };
        } catch (err) {
          // Captura estado da página para diagnóstico (inclui screenshot e html)
          try {
            await logPageState(page, `Erro no relatório ${reportName}: ${String(err && err.message || err)}`, {
              level: 'error',
              shotName: `report-error-${sanitizeFileName(reportName)}-${nowIso().replace(/[:.]/g, '-')}`,
            });
          } catch (_) {}
          
          const finishedAt = new Date().toISOString();
          await updateReportStatusSerialized({ perReportEntry: { report: reportName, props: { status: 'error', finishedAt, lastError: String(err && err.message || err), lastMessage: 'error', progress: null } } });
          logEvent({ level: 'error', message: 'Falha no relatório.', report: reportName, detail: String(err && err.message || err) });
          // rethrow para Promise.allSettled manejar
          throw err;
        } finally {
          try { await page.close(); } catch {}
        }
      })());
      
      const settled = await Promise.allSettled(tasks);
      
      // atualiza relatório geral com resultados
      const completed = settled.filter(s => s.status === 'fulfilled').length;
      status.completedReports = [];
      for (const s of settled) {
        if (s.status === 'fulfilled') {
          status.completedReports.push({ report: s.value.report, finishedAt: new Date().toISOString() });
        }
      }
      
      
      status.current = completed;
      status.progress = Math.round((completed / reports.length) * 100);
      status.currentReport = null;
      status.error = settled.some(s => s.status === 'rejected') ? 'some_reports_failed' : null;
      
      await updateReportStatusSerialized(status);
      
      status.running = false;
      status.finishedAt = new Date().toISOString();
      status.current = reports.length;
      status.progress = 100;
      status.currentReport = null;
      status.error = null;
      
      await updateReportStatus(status);
      
      logEvent({
        level: 'info',
        message: 'Geração de relatórios concluída.'
      });
      
      try { console.log('STEP END runReports'); } catch (e) {}
      
      return {
        success: true,
        completedReports: status.completedReports,
      };
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
    
    async function promptUserForCode() {
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
      // Try multiple times with several selector strategies, then fallback to evaluate click
      const candidatesFactory = () => [
        page.locator('button').filter({ hasText: new RegExp(escRe(USERNAME), 'i') }),
        page.locator('[role="button"]').filter({ hasText: new RegExp(escRe(USERNAME), 'i') }),
        page.locator('a').filter({ hasText: new RegExp(escRe(USERNAME), 'i') }),
        page.locator('button').filter({ hasText: /conectado/i }),
        page.getByText(new RegExp(escRe(USERNAME), 'i')).locator('xpath=ancestor::button[1]'),
        page.getByText(new RegExp(escRe(USERNAME), 'i')).locator('xpath=ancestor::a[1]'),
      ];
      
      let lastErr = null;
      for (let attempt = 1; attempt <= 3; attempt++) {
        try {
          const clicked = await clickFirstVisible(`Conta salva Sienge ID (attempt ${attempt})`, candidatesFactory());
          if (clicked) {
            await waitForAppReady(page, 20000);
            await logPageState(page, 'Conta salva clicada com sucesso.', { shotName: `after-account-click-attempt-${attempt}` });
            return true;
          }
          
          // fallback: try to click via evaluate searching for text content
          const userText = String(USERNAME || '').trim();
          if (userText) {
            const evalClicked = await page.evaluate((u) => {
              try {
                const norm = s => (s || '').trim();
                const nodes = Array.from(document.querySelectorAll('button, a, div, span'));
                for (const n of nodes) {
                  if (norm(n.innerText).toLowerCase().includes(u.toLowerCase())) {
                    try { n.click(); return true; } catch(e) {}
                  }
                }
              } catch (e) {}
              return false;
            }, userText).catch(() => false);
            
            if (evalClicked) {
              await waitForAppReady(page, 20000);
              await logPageState(page, 'Conta salva clicada via evaluate.', { shotName: `after-account-click-eval-${attempt}` });
              return true;
            }
          }
        } catch (err) {
          lastErr = err;
          logEvent({ level: 'debug', message: `tryAccountChooser attempt ${attempt} failed`, detail: String(err && err.message || err) });
          await page.waitForTimeout(800 + attempt * 300);
        }
      }
      
      // final diagnostic: save page state and throw
      await logPageState(page, 'Falha ao clicar na conta salva do Sienge ID.', { level: 'error', shotName: 'account-click-failed' });
      throw new Error('Não consegui clicar na conta salva do Sienge ID.' + (lastErr ? ` Caused by: ${String(lastErr && lastErr.message || lastErr)}` : ''));
    }
    async function detectMfaPinInputs(page) {
      const locators = [
        page.locator('input.grua-pin-input-textfield'),
        page.locator('input[maxlength="1"]'),
        page.locator('.sc-fqkvVR input'),
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
          await inputs[i].click({ timeout: 3000 }).catch(() => {});
          await inputs[i].fill('').catch(() => {});
          await inputs[i].type(digits[i], { delay: 20 }).catch(async () => { await inputs[i].fill(digits[i]); });
        }
        return true;
      }
      return false;
    }
    async function handleMfa(page) {
      let summary = await pageSummary(page);
      if (!isMfaMethodStep(summary) && !isMfaCodeStep(summary)) return false;
      
      if (isMfaMethodStep(summary)) {
        logEvent({ level: 'info', message: 'Tela de escolha do método MFA detectada. Selecionando "Autenticação por e-mail".' });
        const clicked = await clickFirstVisible('Autenticação por e-mail', [
          page.locator('[data-testid="mfa-choose-method-email"]'),
          page.getByRole('button', { name: /autenticação por e-mail|autenticacao por e-mail/i }),
          page.getByText(/autenticação por e-mail|autenticacao por e-mail/i).locator('xpath=ancestor::button[1]'),
          page.locator('button').filter({ hasText: /autenticação por e-mail|autenticacao por e-mail/i }),
        ]);
        if (!clicked) throw new Error('Não consegui selecionar "Autenticação por e-mail".');
        await waitForAppReady(page, 20000);
        await logPageState(page, 'Método MFA por e-mail selecionado.', { shotName: 'after-mfa-email-method' });
        summary = await pageSummary(page);
      }
      
      if (!isMfaCodeStep(summary)) {
        await page.waitForTimeout(1500);
        summary = await pageSummary(page);
      }
      if (!isMfaCodeStep(summary)) throw new Error('A tela para inserir o código MFA não apareceu.');
      
      logEvent({ level: 'info', message: 'Tela de código MFA detectada. Aguardando o código digitado no terminal.' });
      const code = await promptUserForCode();
      const ok = await fillMfaCode(page, code);
      if (!ok) throw new Error('Não encontrei o input para inserir o código MFA.');
      
      await page.waitForTimeout(200);
      const verifyClicked = await clickFirstVisible('Botão Verificar', [
        page.getByRole('button', { name: /verificar/i }),
        page.getByText(/^VERIFICAR$/i),
        page.locator('button[type="submit"]'),
      ]);
      if (!verifyClicked) {
        const pins = await detectMfaPinInputs(page);
        if (pins.length) await pins[pins.length - 1].press('Enter').catch(() => {});
      }
      await waitForAppReady(page, 25000);
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
      await fetchWithRetry(page, `${BASE_URL}/sienge/`, { waitUntil: 'domcontentloaded' });
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
      await fetchWithRetry(page, `${BASE_URL}/sienge/8/index.html`, { waitUntil: 'domcontentloaded' });
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
      
      logEvent({ level: 'info', message: 'Sessão inválida ou ausente. Aguardando revalidação do arquivo de estado externo.' });

      const start = Date.now();
      const timeout = STATE_REVALIDATE_TIMEOUT_MS;
      const poll = STATE_REVALIDATE_POLL_MS;

      async function applyStorageStateToContext(statePath) {
        try {
          const raw = await fs.promises.readFile(statePath, 'utf8').catch(() => null);
          if (!raw) return false;
          const parsed = JSON.parse(raw);
          const cookies = Array.isArray(parsed.cookies) ? parsed.cookies : [];
          const origins = Array.isArray(parsed.origins) ? parsed.origins : [];

          if (cookies.length) {
            try { await context.addCookies(cookies); } catch (e) { logEvent({ level: 'warning', message: 'Falha ao aplicar cookies do storageState', detail: String(e && e.message || e) }); }
          }

          for (const o of origins) {
            try {
              const originUrl = (o && o.origin) ? String(o.origin) : null;
              if (!originUrl) continue;
              const temp = await context.newPage();
              try {
                await temp.goto(originUrl, { waitUntil: 'domcontentloaded', timeout: 10000 }).catch(() => {});
                if (Array.isArray(o.localStorage)) {
                  await temp.evaluate((items) => {
                    try {
                      for (const it of items) {
                        try { localStorage.setItem(it.name, it.value); } catch (e) {}
                      }
                    } catch (e) {}
                  }, o.localStorage).catch(() => {});
                }
              } catch (e) {}
              try { await temp.close(); } catch (e) {}
            } catch (e) { logEvent({ level: 'debug', message: 'Falha ao aplicar localStorage para origin', origin: o && o.origin, detail: String(e && e.message || e) }); }
          }
          return true;
        } catch (e) {
          logEvent({ level: 'warning', message: 'Erro ao aplicar storageState', detail: String(e && e.message || e) });
          return false;
        }
      }

      let applied = false;
      while (Date.now() - start < timeout) {
        if (fs.existsSync(STATE_PATH)) {
          logEvent({ level: 'info', message: 'Arquivo de estado detectado; aplicando ao contexto.' });
          applied = await applyStorageStateToContext(STATE_PATH).catch(() => false);
          if (applied) {
            try {
              await page.reload({ waitUntil: 'domcontentloaded' }).catch(() => {});
              await waitForAppReady(page, 8000);
              summary = await pageSummary(page);
              if (isLoggedArea(summary)) {
                logEvent({ level: 'info', message: 'Sessão válida após aplicação do estado externo.' });
                if (TASK_MODE === 'reports') {
                  await goDirectToReportsPageAfterLogin(page);
                } else {
                  await goDirectToAuthorizationPageAfterLogin(page);
                }
                return;
              }
            } catch (e) { logEvent({ level: 'debug', message: 'Revalidação após aplicação falhou', detail: String(e && e.message || e) }); }
          }
        }
        await new Promise(r => setTimeout(r, poll));
      }

      // timed out without valid session
      throw new Error('Sessão inválida ou ausente e revalidação do arquivo de estado expirou (' + STATE_REVALIDATE_TIMEOUT_MS + 'ms). Aguarde o processo de autenticação externo atualizar ' + STATE_PATH);
    }
    
    
    
    async function goDirectToAuthorizationPageAfterLogin(page) {
      logEvent({ level: 'info', message: 'Após login válido, tentando ir direto para a URL da tela 1777.' });
      
      await page.waitForTimeout(2500);
      await dismissHomeOverlays(page);
      await page.waitForTimeout(1200);
      
      const attempts = [
        async () => {
          await fetchWithRetry(page, TARGET_PAGE_URL, { waitUntil: 'domcontentloaded' });
          await waitForAppReady(page, 12000);
          return await waitForRealAuthorizationPage(page, 10000);
        },
        async () => {
          await fetchWithRetry(page, `${BASE_URL}/sienge/8/index.html`, { waitUntil: 'domcontentloaded' });
          await waitForAppReady(page, 12000);
          await page.evaluate(() => {
            try { window.location.hash = '#/common/page/1777'; } catch {}
          });
          await page.waitForTimeout(3000);
          return await waitForRealAuthorizationPage(page, 10000);
        },
        async () => {
          await fetchWithRetry(page, `${BASE_URL}/sienge/8/index.html`, { waitUntil: 'domcontentloaded' });
          await waitForAppReady(page, 12000);
          await clickRecentAuthorizationLink(page);
          return await waitForRealAuthorizationPage(page, 10000);
        },
        async () => {
          await fetchWithRetry(page, `${BASE_URL}/sienge/8/index.html`, { waitUntil: 'domcontentloaded' });
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
        await fetchWithRetry(page, url, { waitUntil: 'domcontentloaded' });
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
    
    // lightweight dedupe map for console messages per page
    const _consoleSeenMap = new WeakMap();
    async function attachPageDebug(page) {
      function _shouldIgnoreConsoleText(text) {
        if (!text) return false;
        const t = String(text);
        // ignore noisy network/resource failures and common irrelevant messages
        const ignoreRe = /Failed to load resource|net::ERR_FAILED|ERR_FAILED|the server responded with a status of|favicon\.ico|favicon|\/sockjs\.|WebSocket connection to|Source map error/i;
        return ignoreRe.test(t);
      }

      function _hashConsole(type, text) {
        const t = String(text || '').trim().slice(0, 240);
        return `${type}|${t}`;
      }

      const seen = new Map();
      _consoleSeenMap.set(page, seen);

      // Log only page runtime errors (important) and non-noisy console errors once.
      page.on('pageerror', (err) => {
        try {
          logEvent({ level: 'error', message: 'Erro JS na página.', detail: String(err && (err.stack || err.message || err)) });
        } catch (e) {}
      });

      page.on('console', (msg) => {
        try {
          const type = msg.type();
          if (!['error', 'warning'].includes(type)) return;
          const text = msg.text();
          if (_shouldIgnoreConsoleText(text)) return;
          const key = _hashConsole(type, text);
          const now = Date.now();
          const last = seen.get(key) || 0;
          // dedupe identical messages for 5 minutes
          if (now - last < 5 * 60 * 1000) return;
          seen.set(key, now);
          // promote console error severity to 'error' level, warnings stay 'warning'
          const level = type === 'error' ? 'error' : 'warning';
          logEvent({ level, message: `Console ${type} da página.`, detail: text });
        } catch (e) {}
      });

      // Log failed requests only for important resource types to avoid noise
      page.on('requestfailed', (req) => {
        try {
          const f = req.failure();
          if (!f) return;
          const rt = req.resourceType();
          // only care about document/xhr/fetch/script/api resources
          const important = ['document', 'xhr', 'fetch', 'script', 'stylesheet'];
          if (!important.includes(rt)) return;
          const url = req.url();
          // ignore known third-party analytic hosts
          if (/google-analytics|googlesyndication|doubleclick|gstatic|hotjar|facebook|ads|adservice/i.test(url)) return;
          const text = `${req.method()} ${url} - ${f.errorText || f}`;
          const key = `requestfailed|${rt}|${url}`;
          const now = Date.now();
          const last = seen.get(key) || 0;
          if (now - last < 5 * 60 * 1000) return;
          seen.set(key, now);
          logEvent({ level: 'warning', message: 'Requisição falhada na página.', detail: text, url });
        } catch (e) {}
      });
      try {
        await page.route('**/*', (route) => {
          const req = route.request();
          const url = req.url();
          const rt = req.resourceType();
          // block heavy/static assets that don't affect selection
          if (['image', 'media', 'font'].includes(rt)) return route.abort();
          // block common analytics/ads third-party hosts
          const blockedHosts = ['google-analytics', 'googlesyndication', 'doubleclick', 'gstatic', 'analytics', 'hotjar', 'facebook', 'ads', 'adservice'];
          for (const h of blockedHosts) if (url.includes(h)) return route.abort();
          return route.continue();
        });
      } catch (e) {}
    }
    
    async function fetchWithRetry(page, url, options = {}, maxRetries = 3) {
      const settings = {
        timeout: 120000,
        ...options,
      };
      
      for (let i = 0; i < maxRetries; i++) {
        try {
          const response = await page.goto(url, settings);
          if (response && response.ok()) {
            return response;
          }
          if ((!response || !response.ok()) && i < maxRetries - 1) {
            logEvent({ level: 'warning', message: `fetchWithRetry: response not OK, retrying`, attempt: i + 1, url, status: response ? response.status() : 'no-response' });
            await sleepMs(2000);
            continue;
          }
          return response;
        } catch (error) {
          const message = String(error.message || '');
          if (message.includes('socket hang up') && i < maxRetries - 1) {
            logEvent({ level: 'warning', message: `fetchWithRetry retry ${i + 1}/${maxRetries} after socket hang up`, url });
            await sleepMs(2000);
            continue;
          }
          throw error;
        }
      }
    }
    
    async function runAuthorization(context, page) {
      const authSurface = await openAuthorizationPage(page);
      const filterResult = await configureFilters(page, authSurface);
      
      if (filterResult && filterResult.hasResults === false) {
        const idleShot = await saveShot(page, 'final-no-results');
        const idleSummary = await pageSummary(page);
        logEvent({
          level: 'info',
          message: 'Execução concluída sem parcelas pendentes para autorizar.',
          screenshot: idleShot,
          ...idleSummary
        });
      } else {
        await markAllAndSave(page, authSurface);
        const finalShot = await saveShot(page, 'final-success');
        const summary = await pageSummary(page);
        logEvent({ level: 'info', message: 'Processo de autorização concluído com sucesso.', screenshot: finalShot, ...summary });
      }
      
      await saveStorageStateWithMeta(context);
      logEvent({ level: 'info', message: 'Estado da sessão salvo com sucesso.', statePath: STATE_PATH });
    }
    
    async function run() {
      await ensureDir(SCREENSHOT_DIR);
      await ensureDir(REPORT_OUTPUT_DIR);
      // Prepare launch options; non-headless in containers often needs extra flags
      const launchOpts = { headless: HEADLESS };
      try {
        if (!HEADLESS) {
          launchOpts.args = [
            '--no-sandbox',
            '--disable-setuid-sandbox',
            '--disable-dev-shm-usage',
            '--disable-gpu',
            '--window-size=1280,1024'
          ];
        }
      } catch (e) {}

      let browser;
      try {
        browser = await chromium.launch(launchOpts);
      } catch (err) {
        // provide richer logging to help diagnose headful failures
        try {
          logEvent({ level: 'error', message: 'Falha ao lançar browser Playwright.', detail: String(err && (err.stack || err.message || err)) });
        } catch (e) {}
        throw err;
      }
      const context = await browser.newContext(fs.existsSync(STATE_PATH) ? { storageState: STATE_PATH } : {});
      if (BLOCK_RESOURCES) {
        try {
          // Abort images and fonts to reduce chance of resource failures and frame reloads
          await context.route('**/*', (route) => {
            try {
              const req = route.request();
              const type = req.resourceType();
              if (type === 'image' || type === 'font') {
                return route.abort();
              }
            } catch (e) {}
            try { return route.continue(); } catch (e) { try { route.abort(); } catch (_) {} }
          });
        } catch (e) {
          logEvent({ level: 'warn', message: `Não foi possível ativar bloqueio de recursos: ${e.message}` });
        }
      }
      const page = await context.newPage();
      page.setDefaultTimeout(120000);
      page.setDefaultNavigationTimeout(120000);
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
          reportPage.setDefaultTimeout(120000);
          reportPage.setDefaultNavigationTimeout(120000);
          await attachPageDebug(reportPage);
          await runReports(context, reportPage);
        } else {
          await runAuthorization(context, page);
        }
        
        await saveStorageStateWithMeta(context);
        logEvent({ level: 'info', message: 'Estado da sessão salvo com sucesso.', statePath: STATE_PATH });
      } catch (err) {
        const errorShot = await saveShot(page, 'error');
        const html = await saveHtml(page, 'error');
        const summary = await pageSummary(page);
        logEvent({ level: 'error', message: err.message, screenshot: errorShot, html, stack: String(err.stack || ''), ...summary });
        console.error(err);
        process.exitCode = 1;
      } finally {
        await browser.close();
        logEvent({ level: 'info', message: 'Execução finalizada.' });
      }
    }
    
    
    
    const PM2_LOOP = String(process.env.PM2_LOOP || 'false').toLowerCase() === 'true';
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
    });
    
    process.on('SIGTERM', async () => {
      logEvent({ level: 'info', message: 'SIGTERM recebido. Encerrando loop PM2.' });
      __stopping = true;
    });
    
    startPm2Loop().catch(err => {
      logEvent({
        level: 'error',
        message: 'Falha fatal ao inicializar o modo PM2.',
        detail: String(err && err.message || err),
        stack: String(err && err.stack || '')
      });
      process.exit(1);
    });
    
