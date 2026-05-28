#!/usr/bin/env node
const fs = require('fs');
const path = require('path');
const readline = require('readline/promises');
const { stdin: input, stdout: output } = require('process');
const { chromium } = require('playwright');
require('dotenv').config();

const STATUS_FILE = "/tmp/report-status.json";
const BASE_URL = process.env.SIENGE_BASE_URL;
const USERNAME = process.env.SIENGE_USERNAME;
const PASSWORD = process.env.SIENGE_PASSWORD;
const HEADLESS = (process.env.HEADLESS ?? 'true').toLowerCase() !== 'false';
const STATE_PATH = process.env.STATE_PATH || path.resolve(process.cwd(), 'sienge-storage-state.json');
const SCREENSHOT_DIR = process.env.SCREENSHOT_DIR || path.resolve(process.cwd(), 'screenshots');
const LOG_PATH = process.env.LOG_PATH || path.resolve(process.cwd(), 'sienge-authorize-log.json');
const DEBUG_HTML = (process.env.DEBUG_HTML ?? 'false').toLowerCase() === 'true';
const TASK_MODE = (process.env.TASK_MODE || 'authorize').toLowerCase();
function getCliArg(name) {
  const idx = process.argv.indexOf(name);
  if (idx >= 0 && process.argv.length > idx + 1) return process.argv[idx + 1];
  return null;
}

const SINGLE_REPORT_ARG = (process.env.SINGLE_REPORT || getCliArg('--single') || getCliArg('-s')) || null;
const REPORT_OUTPUT_DIR = process.env.REPORT_OUTPUT_DIR || path.resolve(process.cwd(), 'reports');
const TARGET_PAGE_URL = `${BASE_URL}/sienge/8/index.html#/common/page/1777`;
const REPORT_FILTER_PAGE_URL = `${BASE_URL}/sienge/8/index.html#/common/page/4929`;
const TARGET_END_DATE = '31/12/2040';
const REPORT_PERIOD_START = '01/04/2026';
const REPORT_PERIOD_END = '30/04/2026';

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
    planoFinanceiro: ['Receita de Incorporação de Imóveis', 'Receita de Bens/Imóveis Adquiridos de Terceiros', 'Descontos Concedidos (Unidades Imobiliárias)', 'xxxxxReceita de Aluguel (Locação)', 'Permuta Externa', 'Receita de Cliente Desconhecido', 'Receita de Administração (Taxa Administrativa)', 'Receita Empreitada Só Mão de Obra', 'Receita de Projeto', 'Receita de Locação (Aluguel)', 'Entrada de Financiamento para Construção', 'Entrada de Capital de Giro', 'Receita de Aplicações Financeiras', 'Receita de Empréstimos', 'Aporte Capital de Sócio', 'Descontos Obtidos', 'Receita de Venda de Imobilizado', 'Receita Aluguel PCS 35', 'Estoque de Material', 'Receita Energia Solar Locatários', 'Venda de Passivos', '(-) Anulação de Custos', '(-) Anulação de Custos Mão de Obra', '(-) Retenção Caução/Sinal', '(-) Retenção de INSS', '(-) Retenção de ISS', '(-) Retenção de PIS', '(-) Retenção de CSLL', '(-) Retenção de COFINS', '(-) Retenção de IR', '(-) Retenção (PIS/COFINS/CSLL -  4,65%)', '(-) Retenção de Permuta', '(-) Transf. Materiais Estoque Laurentino P/ Finan', '(-) Transf. Materiais Estoque Laurentino P/ Bom J.', '(-) Transf. Mat. Estoque P/ Obras da Laurentino', '(-) Reembolso de Despesas Gerais (Obra)', '(-) Divisão de muro', '(-) Retenção Contribuição Sindical (Escritório)', '(-) Reembolso de Despesas Administrativas', '(-) Restituição IPTU Imóveis Alugados (Escritório)', '(-) Reembolso de Despesas Financeiras', '(-) Reembolso de Despesas Tributárias', '(-) Abatimento de Adiantamento', '(-) Restituição Condomínio', '(-) Restituição água'],
    planoFinanceiroExcecao: [''],
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
        "Emissao de Cheques",
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
        "PERMUTA CONCRETOREI GREEN GARDEN LOTE 45",
        "PERMUTA DAINOX BP1 TERRENO MAT. 21884",
        "PERMUTA DAINOX SALA 7 SEVEN POINT",
        "PERMUTA DAINOX GOLDEN PARK LOT A DEF.",
        "PERMUTA DAMACENO",
        "PERMUTA DAMACENO ASFALTO GREEN GARDEN",
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
    documentosExcecao: [''],
    condicoesPagamento: ['Parcelas Mensais', 'Novo Parcelas Mensais'],
    condicoesPagamentoExcecao: [''],
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
        "Emissao de Cheques",
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
        "PERMUTA CONCRETOREI GREEN GARDEN LOTE 45",
        "PERMUTA DAINOX BP1 TERRENO MAT. 21884",
        "PERMUTA DAINOX SALA 7 SEVEN POINT",
        "PERMUTA DAINOX GOLDEN PARK LOT A DEF.",
        "PERMUTA DAMACENO",
        "PERMUTA DAMACENO ASFALTO GREEN GARDEN",
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
        "PERMUTA WJ SEVEN POINT SALA 2",
        "PERMUTA WARNOW PARK 1 INDAPAV",
        "PERMUTA WARNOW PARK 2 BERTELLI",
        "SALDO DEVOLULÇAO G LIGHT",
        "TROCA DO KW 5 CASA 3 PELO TERRENO SARTOR",
        "PERMUTA TERRENO CARIJÓS NO BP1 08",
        "PERMUTA TERRENO HENRIQUE KUNEN EST. AREA",
        "PERMUTA TOMIO GREEN GARDEN LOTE 46"
    ],
    planoFinanceiro: ['Receita de Incorporação de Imóveis', 'Receita de Bens/Imóveis Adquiridos de Terceiros', 'Descontos Concedidos (Unidades Imobiliárias)', 'xxxxxReceita de Aluguel (Locação)', 'Permuta Externa', 'Receita de Cliente Desconhecido', 'Receita de Administração (Taxa Administrativa)', 'Receita Empreitada Só Mão de Obra', 'Receita de Projeto', 'Receita de Locação (Aluguel)', 'Entrada de Financiamento para Construção', 'Entrada de Capital de Giro', 'Receita de Aplicações Financeiras', 'Receita de Empréstimos', 'Aporte Capital de Sócio', 'Descontos Obtidos', 'Receita de Venda de Imobilizado', 'Receita Aluguel PCS 35', 'Estoque de Material', 'Receita Energia Solar Locatários', 'Venda de Passivos', '(-) Anulação de Custos', '(-) Anulação de Custos Mão de Obra', '(-) Retenção Caução/Sinal', '(-) Retenção de INSS', '(-) Retenção de ISS', '(-) Retenção de PIS', '(-) Retenção de CSLL', '(-) Retenção de COFINS', '(-) Retenção de IR', '(-) Retenção (PIS/COFINS/CSLL -  4,65%)', '(-) Retenção de Permuta', '(-) Transf. Materiais Estoque Laurentino P/ Finan', '(-) Transf. Materiais Estoque Laurentino P/ Bom J.', '(-) Transf. Mat. Estoque P/ Obras da Laurentino', '(-) Reembolso de Despesas Gerais (Obra)', '(-) Divisão de muro', '(-) Retenção Contribuição Sindical (Escritório)', '(-) Reembolso de Despesas Administrativas', '(-) Restituição IPTU Imóveis Alugados (Escritório)', '(-) Reembolso de Despesas Financeiras', '(-) Reembolso de Despesas Tributárias', '(-) Abatimento de Adiantamento', '(-) Restituição Condomínio', '(-) Restituição água'],
    planoFinanceiroExcecao: [''],
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
    planoFinanceiro: ['Receita de Incorporação de Imóveis', 'Receita de Bens/Imóveis Adquiridos de Terceiros', 'Descontos Concedidos (Unidades Imobiliárias)', 'xxxxxReceita de Aluguel (Locação)', 'Permuta Externa', 'Receita de Cliente Desconhecido', 'Receita de Administração (Taxa Administrativa)', 'Receita Empreitada Só Mão de Obra', 'Receita de Projeto', 'Receita de Locação (Aluguel)', 'Entrada de Financiamento para Construção', 'Entrada de Capital de Giro', 'Receita de Aplicações Financeiras', 'Receita de Empréstimos', 'Aporte Capital de Sócio', 'Descontos Obtidos', 'Receita de Venda de Imobilizado', 'Receita Aluguel PCS 35', 'Estoque de Material', 'Receita Energia Solar Locatários', 'Venda de Passivos', '(-) Anulação de Custos', '(-) Anulação de Custos Mão de Obra', '(-) Retenção Caução/Sinal', '(-) Retenção de INSS', '(-) Retenção de ISS', '(-) Retenção de PIS', '(-) Retenção de CSLL', '(-) Retenção de COFINS', '(-) Retenção de IR', '(-) Retenção (PIS/COFINS/CSLL -  4,65%)', '(-) Retenção de Permuta', '(-) Transf. Materiais Estoque Laurentino P/ Finan', '(-) Transf. Materiais Estoque Laurentino P/ Bom J.', '(-) Transf. Mat. Estoque P/ Obras da Laurentino', '(-) Reembolso de Despesas Gerais (Obra)', '(-) Divisão de muro', '(-) Retenção Contribuição Sindical (Escritório)', '(-) Reembolso de Despesas Administrativas', '(-) Restituição IPTU Imóveis Alugados (Escritório)', '(-) Reembolso de Despesas Financeiras', '(-) Reembolso de Despesas Tributárias', '(-) Abatimento de Adiantamento', '(-) Restituição Condomínio', '(-) Restituição água'],
    planoFinanceiroExcecao: [],
    documentos: [
        "ADTO",
        "Adiantamento",
        "APORTE DE CAPITAL",
        "AVISO DE LANÇAMENTO",
        "CAUÇÃO",
        "CONHECIMENTO DE FRETE",
        "CONTRATO",
        "CUPOM FISCAL",
        "Documento de Arrecadação de Receita Est"
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
        "Emissao de Cheques",
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
        "PERMUTA CONCRETOREI GREEN GARDEN LOTE 45",
        "PERMUTA DAINOX BP1 TERRENO MAT. 21884",
        "PERMUTA DAINOX SALA 7 SEVEN POINT",
        "PERMUTA DAINOX GOLDEN PARK LOT A DEF.",
        "PERMUTA DAMACENO",
        "PERMUTA DAMACENO ASFALTO GREEN GARDEN",
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
        "Emissao de Cheques",
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
        "PERMUTA CONCRETOREI GREEN GARDEN LOTE 45",
        "PERMUTA DAINOX BP1 TERRENO MAT. 21884",
        "PERMUTA DAINOX SALA 7 SEVEN POINT",
        "PERMUTA DAINOX GOLDEN PARK LOT A DEF.",
        "PERMUTA DAMACENO",
        "PERMUTA DAMACENO ASFALTO GREEN GARDEN",
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
        "PERMUTA WJ SEVEN POINT SALA 2",
        "PERMUTA WARNOW PARK 1 INDAPAV",
        "PERMUTA WARNOW PARK 2 BERTELLI",
        "SALDO DEVOLULÇAO G LIGHT",
        "TROCA DO KW 5 CASA 3 PELO TERRENO SARTOR",
        "PERMUTA TERRENO CARIJÓS NO BP1 08",
        "PERMUTA TERRENO HENRIQUE KUNEN EST. AREA",
        "PERMUTA TOMIO GREEN GARDEN LOTE 46"
    ],
    planoFinanceiro: ['Receita de Incorporação de Imóveis', 'Receita de Bens/Imóveis Adquiridos de Terceiros', 'Descontos Concedidos (Unidades Imobiliárias)', 'xxxxxReceita de Aluguel (Locação)', 'Permuta Externa', 'Receita de Cliente Desconhecido', 'Receita de Administração (Taxa Administrativa)', 'Receita Empreitada Só Mão de Obra', 'Receita de Projeto', 'Receita de Locação (Aluguel)', 'Entrada de Financiamento para Construção', 'Entrada de Capital de Giro', 'Receita de Aplicações Financeiras', 'Receita de Empréstimos', 'Aporte Capital de Sócio', 'Descontos Obtidos', 'Receita de Venda de Imobilizado', 'Receita Aluguel PCS 35', 'Estoque de Material', 'Receita Energia Solar Locatários', 'Venda de Passivos', '(-) Anulação de Custos', '(-) Anulação de Custos Mão de Obra', '(-) Retenção Caução/Sinal', '(-) Retenção de INSS', '(-) Retenção de ISS', '(-) Retenção de PIS', '(-) Retenção de CSLL', '(-) Retenção de COFINS', '(-) Retenção de IR', '(-) Retenção (PIS/COFINS/CSLL -  4,65%)', '(-) Retenção de Permuta', '(-) Transf. Materiais Estoque Laurentino P/ Finan', '(-) Transf. Materiais Estoque Laurentino P/ Bom J.', '(-) Transf. Mat. Estoque P/ Obras da Laurentino', '(-) Reembolso de Despesas Gerais (Obra)', '(-) Divisão de muro', '(-) Retenção Contribuição Sindical (Escritório)', '(-) Reembolso de Despesas Administrativas', '(-) Restituição IPTU Imóveis Alugados (Escritório)', '(-) Reembolso de Despesas Financeiras', '(-) Reembolso de Despesas Tributárias', '(-) Abatimento de Adiantamento', '(-) Restituição Condomínio', '(-) Restituição água'],
    planoFinanceiroExcecao: [''],
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
    planoFinanceiro: ['Receita de Incorporação de Imóveis', 'Receita de Bens/Imóveis Adquiridos de Terceiros', 'Descontos Concedidos (Unidades Imobiliárias)', 'xxxxxReceita de Aluguel (Locação)', 'Permuta Externa', 'Receita de Cliente Desconhecido', 'Receita de Administração (Taxa Administrativa)', 'Receita Empreitada Só Mão de Obra', 'Receita de Projeto', 'Receita de Locação (Aluguel)', 'Entrada de Financiamento para Construção', 'Entrada de Capital de Giro', 'Receita de Aplicações Financeiras', 'Receita de Empréstimos', 'Aporte Capital de Sócio', 'Descontos Obtidos', 'Receita de Venda de Imobilizado', 'Receita Aluguel PCS 35', 'Estoque de Material', 'Receita Energia Solar Locatários', 'Venda de Passivos', '(-) Anulação de Custos', '(-) Anulação de Custos Mão de Obra', '(-) Retenção Caução/Sinal', '(-) Retenção de INSS', '(-) Retenção de ISS', '(-) Retenção de PIS', '(-) Retenção de CSLL', '(-) Retenção de COFINS', '(-) Retenção de IR', '(-) Retenção (PIS/COFINS/CSLL -  4,65%)', '(-) Retenção de Permuta', '(-) Transf. Materiais Estoque Laurentino P/ Finan', '(-) Transf. Materiais Estoque Laurentino P/ Bom J.', '(-) Transf. Mat. Estoque P/ Obras da Laurentino', '(-) Reembolso de Despesas Gerais (Obra)', '(-) Divisão de muro', '(-) Retenção Contribuição Sindical (Escritório)', '(-) Reembolso de Despesas Administrativas', '(-) Restituição IPTU Imóveis Alugados (Escritório)', '(-) Reembolso de Despesas Financeiras', '(-) Reembolso de Despesas Tributárias', '(-) Abatimento de Adiantamento', '(-) Restituição Condomínio', '(-) Restituição água'],
    planoFinanceiroExcecao: [],
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
        "Emissao de Cheques",
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
        "PERMUTA CONCRETOREI GREEN GARDEN LOTE 45",
        "PERMUTA DAINOX BP1 TERRENO MAT. 21884",
        "PERMUTA DAINOX SALA 7 SEVEN POINT",
        "PERMUTA DAINOX GOLDEN PARK LOT A DEF.",
        "PERMUTA DAMACENO",
        "PERMUTA DAMACENO ASFALTO GREEN GARDEN",
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
        "Emissao de Cheques",
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
        "PERMUTA CONCRETOREI GREEN GARDEN LOTE 45",
        "PERMUTA DAINOX BP1 TERRENO MAT. 21884",
        "PERMUTA DAINOX SALA 7 SEVEN POINT",
        "PERMUTA DAINOX GOLDEN PARK LOT A DEF.",
        "PERMUTA DAMACENO",
        "PERMUTA DAMACENO ASFALTO GREEN GARDEN",
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
        "PERMUTA WJ SEVEN POINT SALA 2",
        "PERMUTA WARNOW PARK 1 INDAPAV",
        "PERMUTA WARNOW PARK 2 BERTELLI",
        "SALDO DEVOLULÇAO G LIGHT",
        "TROCA DO KW 5 CASA 3 PELO TERRENO SARTOR",
        "PERMUTA TERRENO CARIJÓS NO BP1 08",
        "PERMUTA TERRENO HENRIQUE KUNEN EST. AREA",
        "PERMUTA TOMIO GREEN GARDEN LOTE 46"
    ],
    planoFinanceiro: ['Receita de Incorporação de Imóveis', 'Receita de Bens/Imóveis Adquiridos de Terceiros', 'Descontos Concedidos (Unidades Imobiliárias)', 'xxxxxReceita de Aluguel (Locação)', 'Permuta Externa', 'Receita de Cliente Desconhecido', 'Receita de Administração (Taxa Administrativa)', 'Receita Empreitada Só Mão de Obra', 'Receita de Projeto', 'Receita de Locação (Aluguel)', 'Entrada de Financiamento para Construção', 'Entrada de Capital de Giro', 'Receita de Aplicações Financeiras', 'Receita de Empréstimos', 'Aporte Capital de Sócio', 'Descontos Obtidos', 'Receita de Venda de Imobilizado', 'Receita Aluguel PCS 35', 'Estoque de Material', 'Receita Energia Solar Locatários', 'Venda de Passivos', '(-) Anulação de Custos', '(-) Anulação de Custos Mão de Obra', '(-) Retenção Caução/Sinal', '(-) Retenção de INSS', '(-) Retenção de ISS', '(-) Retenção de PIS', '(-) Retenção de CSLL', '(-) Retenção de COFINS', '(-) Retenção de IR', '(-) Retenção (PIS/COFINS/CSLL -  4,65%)', '(-) Retenção de Permuta', '(-) Transf. Materiais Estoque Laurentino P/ Finan', '(-) Transf. Materiais Estoque Laurentino P/ Bom J.', '(-) Transf. Mat. Estoque P/ Obras da Laurentino', '(-) Reembolso de Despesas Gerais (Obra)', '(-) Divisão de muro', '(-) Retenção Contribuição Sindical (Escritório)', '(-) Reembolso de Despesas Administrativas', '(-) Restituição IPTU Imóveis Alugados (Escritório)', '(-) Reembolso de Despesas Financeiras', '(-) Reembolso de Despesas Tributárias', '(-) Abatimento de Adiantamento', '(-) Restituição Condomínio', '(-) Restituição água'],
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
    planoFinanceiroExcecao: [],
    documentos: [
        "ADTO",
        "Adiantamento",
        "APORTE DE CAPITAL",
        "AVISO DE LANÇAMENTO",
        "CAUÇÃO",
        "CONHECIMENTO DE FRETE",
        "CONTRATO",
        "CUPOM FISCAL",
        "Documento de Arrecadação de Receita Est"
    ],
    condicoesPagamento: [
        "PM",
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
        "ADIANTAMENTO CASA DA LAJE 250K",
        "ADIANTAMENTO 95K PORTARIS",
        "CAIXA Laurentino",
        "Conta Babi - Não Conciliar",
        "Conta Doca",
        "Conta Equilíbrio",
        "Emissao de Cheques",
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
        "PERMUTA CONCRETOREI GREEN GARDEN LOTE 45",
        "PERMUTA DAINOX BP1 TERRENO MAT. 21884",
        "PERMUTA DAINOX SALA 7 SEVEN POINT",
        "PERMUTA DAINOX GOLDEN PARK LOT A DEF.",
        "PERMUTA DAMACENO",
        "PERMUTA DAMACENO ASFALTO GREEN GARDEN",
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
    planoFinanceiroExcecao: [''],
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
        "Emissao de Cheques",
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
        "PERMUTA CONCRETOREI GREEN GARDEN LOTE 45",
        "PERMUTA DAINOX BP1 TERRENO MAT. 21884",
        "PERMUTA DAINOX SALA 7 SEVEN POINT",
        "PERMUTA DAINOX GOLDEN PARK LOT A DEF.",
        "PERMUTA DAMACENO",
        "PERMUTA DAMACENO ASFALTO GREEN GARDEN",
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
        "ADTO",
        "Adiantamento",
        "APORTE DE CAPITAL",
        "AVISO DE LANÇAMENTO",
        "CAUÇÃO",
        "CONHECIMENTO DE FRETE",
        "CONTRATO",
        "CUPOM FISCAL",
        "Documento de Arrecadação de Receita Est"
    ],
    condicoesPagamento: [
        "PM",
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
        "Emissao de Cheques",
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
        "PERMUTA CONCRETOREI GREEN GARDEN LOTE 45",
        "PERMUTA DAINOX BP1 TERRENO MAT. 21884",
        "PERMUTA DAINOX SALA 7 SEVEN POINT",
        "PERMUTA DAINOX GOLDEN PARK LOT A DEF.",
        "PERMUTA DAMACENO",
        "PERMUTA DAMACENO ASFALTO GREEN GARDEN",
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
    planoFinanceiroExcecao: [],
    documentos: [
        "ADTO",
        "Adiantamento",
        "APORTE DE CAPITAL",
        "AVISO DE LANÇAMENTO",
        "CAUÇÃO",
        "CONHECIMENTO DE FRETE",
        "CONTRATO",
        "CUPOM FISCAL",
        "Documento de Arrecadação de Receita Est"
    ],
    condicoesPagamento: [
        "PM",
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
        "Emissao de Cheques",
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
        "PERMUTA CONCRETOREI GREEN GARDEN LOTE 45",
        "PERMUTA DAINOX BP1 TERRENO MAT. 21884",
        "PERMUTA DAINOX SALA 7 SEVEN POINT",
        "PERMUTA DAINOX GOLDEN PARK LOT A DEF.",
        "PERMUTA DAMACENO",
        "PERMUTA DAMACENO ASFALTO GREEN GARDEN",
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
        "PERMUTA WJ SEVEN POINT SALA 2",
        "PERMUTA WARNOW PARK 1 INDAPAV",
        "PERMUTA WARNOW PARK 2 BERTELLI",
        "SALDO DEVOLULÇAO G LIGHT",
        "TROCA DO KW 5 CASA 3 PELO TERRENO SARTOR",
        "PERMUTA TERRENO CARIJÓS NO BP1 08",
        "PERMUTA TERRENO HENRIQUE KUNEN EST. AREA",
        "PERMUTA TOMIO GREEN GARDEN LOTE 46"
    ],
    planoFinanceiro: ['Receita de Incorporação de Imóveis', 'Receita de Bens/Imóveis Adquiridos de Terceiros', 'Descontos Concedidos (Unidades Imobiliárias)', 'xxxxxReceita de Aluguel (Locação)', 'Permuta Externa', 'Receita de Cliente Desconhecido', 'Receita de Administração (Taxa Administrativa)', 'Receita Empreitada Só Mão de Obra', 'Receita de Projeto', 'Receita de Locação (Aluguel)', 'Entrada de Financiamento para Construção', 'Entrada de Capital de Giro', 'Receita de Aplicações Financeiras', 'Descontos Obtidos', 'Receita de Venda de Imobilizado', 'Receita Aluguel PCS 35', 'Estoque de Material', 'Receita Energia Solar Locatários', '(-) Anulação de Custos', '(-) Anulação de Custos Mão de Obra', '(-) Retenção Caução/Sinal', '(-) Retenção de INSS', '(-) Retenção de ISS', '(-) Retenção de PIS', '(-) Retenção de CSLL', '(-) Retenção de COFINS', '(-) Retenção de IR', '(-) Retenção (PIS/COFINS/CSLL -  4,65%)', '(-) Retenção de Permuta', '(-) Transf. Materiais Estoque Laurentino P/ Finan', '(-) Transf. Materiais Estoque Laurentino P/ Bom J.', '(-) Transf. Mat. Estoque P/ Obras da Laurentino', '(-) Reembolso de Despesas Gerais (Obra)', '(-) Divisão de muro', '(-) Retenção Contribuição Sindical (Escritório)', '(-) Reembolso de Despesas Administrativas', '(-) Restituição IPTU Imóveis Alugados (Escritório)', '(-) Reembolso de Despesas Financeiras', '(-) Reembolso de Despesas Tributárias', '(-) Abatimento de Adiantamento', '(-) Restituição Condomínio', '(-) Restituição água'],
    planoFinanceiroExcecao: [],
    documentos: [
        "ADTO",
        "Adiantamento",
        "APORTE DE CAPITAL",
        "AVISO DE LANÇAMENTO",
        "CAUÇÃO",
        "CONHECIMENTO DE FRETE",
        "CUPOM FISCAL",
        "Documento de Arrecadação de Receita Est"
    ],
    documentosExcecao: [],
    condicoesPagamento: [
        "PM",
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
    condicoesPagamentoExcecao: [],
    flags: {
      imprimirParcelasReparceladas: false,
    },

    ordem: 'VL',

    processarLancamentos: 'CR',
  },
];

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

  await new Promise(resolve => setTimeout(resolve, 800));
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
  }

  if (data && data.perReportEntry) {
    cur.reports = cur.reports || {};
    const name = data.perReportEntry.report;
    cur.reports[name] = { ...(cur.reports[name] || {}), ...(data.perReportEntry.props || {}) };
  }

  // Backward compatibility: if caller passed a plain object without wrapper,
  // and it wasn't recognized as overall, write it as-is.
  if (!isOverall && !(data && data.perReportEntry)) {
    cur = data;
  }

  await fs.promises.writeFile(STATUS_FILE, JSON.stringify(cur, null, 2), 'utf8');
}

// Mutex simples para serializar atualizações do STATUS_FILE entre chamadas assíncronas
let STATUS_UPDATE_LOCK = Promise.resolve();

async function updateReportStatusSerialized(data) {
  const op = async () => updateReportStatus(data);
  // Encadeia a operação na fila
  STATUS_UPDATE_LOCK = STATUS_UPDATE_LOCK.then(op, op);
  return STATUS_UPDATE_LOCK;
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
  fs.writeFileSync(LOG_PATH, JSON.stringify(arr, null, 2), 'utf8');
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
  if (payload.detail) parts.push(`detail=${truncateForLog(payload.detail, 200)}`);

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
  await ensureDir(SCREENSHOT_DIR);
  const file = path.join(SCREENSHOT_DIR, `${Date.now()}-${name}.png`);
  await page.screenshot({ path: file, fullPage: true });
  return file;
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
}) {

  await selectViaModal({

    page,

    triggerSelector:
      'img[onclick*="searchContaCorrenteCompleto.jsp"]',

    modalTitle:
      'Consulta de Contas Correntes',

    // campo "Nome"
    searchInputSelector: [
      // usa os seletores reais do modal de contas correntes
      'input[name="entity.contaCorrentePK.nuConta"]',
      '#entity.contaCorrentePK.nuConta',
      'input[name="entity.nmConta"]',
      '#entity.nmConta',

    ],

    values,

  });
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

  await selectPlanoFinanceiro({
    page,
    values,
  });

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

  const reportFrame = await getReportFilterFrame(page);

  // =====================================================
  // RESET
  // =====================================================

  await resetSomenteCamposEditaveis(
    reportFrame
  );

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

  const empresa = reportFrame.locator('#cdEmpresaView').first();

  await empresa.waitFor({
    state: 'visible',
    timeout: 30000,
  });
  
  const dtInicioInput =
    reportFrame.locator(
      '#dtRectoInicio'
    ).first();

  await dtInicioInput.waitFor({
    state: 'visible',
    timeout: 30000,
  });

  await empresa.fill('1');

  await empresa.evaluate(el => {
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

  saveShot(page, `empresa-filled-${report.sheetName}`);
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

  // =====================================================
  // TIPO LANÇAMENTO
  // =====================================================

  await configureTipoLancamento(
    reportFrame,
    report.processarLancamentos || 'CR'
  );

  // =====================================================
  // FLAGS
  // =====================================================

  await configureFlags(
    reportFrame,
    report.flags || {}
  );

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

  let planos =
    normalizeValues(
      report.planoFinanceiro
    );

  const planosExcecao =
    normalizeValues(
      report.planoFinanceiroExcecao
    );

  const selecionarTodosPlanos =

    !planos.length
    || planos.includes('*');

  if (
    !selecionarTodosPlanos
  ) {

    planos =
      removeExcecoes(
        planos,
        planosExcecao
      );

    if (planos.length) {

      logEvent({
        level: 'info',
        message:
          `[${report.sheetName}] Selecionando plano financeiro`,
        planos,
      });

      await selectPlanoFinanceiro({

        page,

        values: planos,

      });
    }

  } else {

    logEvent({
      level: 'info',
      message:
        `[${report.sheetName}] Plano financeiro = TODOS`,
      excecoes: planosExcecao,
    });
  }

  // =====================================================
  // CONTAS CORRENTE
  // =====================================================

  let contas =
    normalizeValues(
      report.contasCorrente
    );

  const contasExcecao =
    normalizeValues(
      report.contasCorrenteExcecao
    );

  const selecionarTodasContas =

    !contas.length
    || contas.includes('*');

  if (
    !selecionarTodasContas
  ) {

    contas =
      removeExcecoes(
        contas,
        contasExcecao
      );

    if (contas.length) {

      logEvent({
        level: 'info',
        message:
          `[${report.sheetName}] Selecionando contas corrente`,
        contas,
      });

      await selectContasCorrente({

        page,

        values: contas,

      });
    }

  } else {

    logEvent({
      level: 'info',
      message:
        `[${report.sheetName}] Contas corrente = TODAS`,
      excecoes: contasExcecao,
    });
  }

  // =====================================================
  // DOCUMENTOS
  // =====================================================

  let documentos =
    normalizeValues(
      report.documentos
    );

  const documentosExcecao =
    normalizeValues(
      report.documentosExcecao
    );

  const selecionarTodosDocs =

    !documentos.length
    || documentos.includes('*');

  if (
    !selecionarTodosDocs
  ) {

    documentos =
      removeExcecoes(
        documentos,
        documentosExcecao
      );

    if (documentos.length) {

      logEvent({
        level: 'info',
        message:
          `[${report.sheetName}] Selecionando documentos`,
        documentos,
      });

      await selectDocumentos({

        page,

        values: documentos,

      });
    }

  } else {

    logEvent({
      level: 'info',
      message:
        `[${report.sheetName}] Documentos = TODOS`,
      excecoes: documentosExcecao,
    });
  }

  // =====================================================
  // CONDIÇÕES PAGAMENTO
  // =====================================================

  let condicoes =
    normalizeValues(
      report.condicoesPagamento
    );

  const condicoesExcecao =
    normalizeValues(
      report.condicoesPagamentoExcecao
    );

  const selecionarTodasCondicoes =

    !condicoes.length
    || condicoes.includes('*');

  if (
    !selecionarTodasCondicoes
  ) {

    condicoes =
      removeExcecoes(
        condicoes,
        condicoesExcecao
      );

    if (condicoes.length) {

      logEvent({
        level: 'info',
        message:
          `[${report.sheetName}] Selecionando condições`,
        condicoes,
      });

      await selectCondicoesPagamento({

        page,

        values: condicoes,

      });
    }

  } else {

    logEvent({
      level: 'info',
      message:
        `[${report.sheetName}] Condições pagamento = TODAS`,
      excecoes: condicoesExcecao,
    });
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
}) {

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
  filePath,
  maxRetries = 4
) {

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
  debugLabel
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

      const matches =
        url.includes(frameName);

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
    'Frame de Documentos'
  );
}

// =====================================================
// FRAME CONDIÇÕES PAGAMENTO
// =====================================================
async function getCondicaoPagamentoFrame(page) {

  return findSelectionFrame(
    page,
    'condicao',
    'Frame de Condição de Pagamento'
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
}) {
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

  });
}

// =====================================================
// SELECIONA DOCUMENTOS
// =====================================================
async function selectDocumentos({
  page,
  values = [],
}) {
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

  return null;
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

            // partial (substring)
            if (normalizedContent.includes(normalizedText)) {
              logEvent({ level: 'debug', message: `Linha encontrada (PARTIAL) em ${sel}: ${text}`, frameUrl: frame.url(), index: i });
              return { frame, cell: el };
            }

            // token match: split search into tokens and ensure each token exists in content
            const tokens = normalizedText.split(' ').filter(Boolean);
            if (tokens.length > 1) {
              let all = true;
              for (const t of tokens) {
                if (!normalizedContent.includes(t)) { all = false; break; }
              }
              if (all) {
                logEvent({ level: 'debug', message: `Linha encontrada (TOKENS) em ${sel}: ${text}`, frameUrl: frame.url(), index: i });
                return { frame, cell: el };
              }
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
            if (normalize(rows[i].innerText).includes(search.toLowerCase())) return i + 1;
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

async function selectViaModal({

  page,

  triggerSelector,

  modalTitle,

  searchInputSelector,

  value,
  values,

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

  // =====================================================
  // HELPERS
  // =====================================================

  async function resolveInput(
    container,
    selectors = [],
    retries = 3
  ) {

    for (let attempt = 0; attempt < retries; attempt++) {

      for (const selector of selectors) {

        const locators =
          container.locator(selector);

        try {

          const count =
            await locators.count();

          if (count === 0 && attempt === 0) {
            continue;
          }

          for (let i = 0; i < count; i++) {

            const input =
              locators.nth(i);

            try {

              // Aguarda o elemento ficar visível com timeout
              await input.waitFor({
                state: 'visible',
                timeout: 3000
              }).catch(() => {});

              const visible =
                await input.isVisible();

              if (!visible) {
                continue;
              }

              const box =
                await input.boundingBox();

              if (!box) {
                continue;
              }

              logEvent({
                level: 'debug',
                message: `Input encontrado no modal. Seletor: ${selector}, Tentativa: ${attempt + 1}/${retries}`
              });

              return input;

            } catch (_) {}
          }

        } catch (_) {}
      }

      // Se não encontrou, aguarda um pouco antes de tentar novamente
      if (attempt < retries - 1) {
        logEvent({
          level: 'debug',
          message: `Input não encontrado na tentativa ${attempt + 1}. Aguardando 800ms...`
        });
        // Aguarda 800ms usando promise
        await new Promise(resolve => setTimeout(resolve, 800));
      }
    }

    logEvent({
      level: 'error',
      message: `Falha ao encontrar input após ${retries} tentativas`,
      selectors: selectors
    });

    // Última tentativa: encontra qualquer input, mesmo que não visível
    try {
      const anyInput = await container.locator('input').first();
      if (anyInput) {
        logEvent({
          level: 'warn',
          message: `Usando fallback: primeiro input encontrado (pode não estar visível)`
        });
        return anyInput;
      }
    } catch (_) {}

    // Debug: Tenta salvar o HTML do modal para diagnóstico
    if (DEBUG_HTML) {
      try {
        const modalHtml = await container.content().catch(() => '<!-- erro ao capturar -->');;
        const debugFile = path.join(
          SCREENSHOT_DIR,
          `debug-modal-${nowIso().replace(/[:.]/g, '-')}.html`
        );
        fs.writeFileSync(debugFile, modalHtml || '<!-- vazio -->');
        logEvent({
          level: 'info',
          message: `HTML do modal salvo em: ${debugFile}`
        });
      } catch (e) {
        logEvent({
          level: 'warn',
          message: `Erro ao salvar HTML debug: ${e.message}`
        });
      }
    }

    throw new Error(
      `Não encontrei input visível. Selectors: ${selectors.join(', ')}`
    );
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

  // =====================================================
  // VALOR ÚNICO
  // MAIN FRAME DIRETO
  // =====================================================

  if (entries.length === 1) {

    const singleValue =
      entries[0];

    // Para o caso único, tentamos uma lista mais ampla de seletores
    const mainSelectors = Array.isArray(searchInputSelector) ? [...searchInputSelector] : [searchInputSelector];
    mainSelectors.push(
      'input[name*="plano"]',
      'input[id*="plano"]',
      'input[class*="plano"]',
      'input[placeholder*="Plano"]',
      'input[placeholder*="plano"]',
      'input'
    );

    const input = await resolveInput(mainFrame, mainSelectors, 4).catch(() => null);

    if (!input) {
      logEvent({ level: 'warn', message: `Não encontrei input do tipo 'plano' no frame principal. Tentando localizar qualquer input visível.` });
      // última tentativa: pega primeiro input visível no mainFrame
      try {
        const any = mainFrame.locator('input').first();
        if (await any.count().catch(() => 0)) {
          await fillLegacyInput(any, singleValue);
        }
      } catch (e) {
        logEvent({ level: 'error', message: `Falha ao preencher input único do mainFrame: ${e.message}` });
      }
    } else {
      await fillLegacyInput(input, singleValue);
      // força perda de foco/validação
      try { await input.press('Tab').catch(() => {}); } catch (_) {}
    }

    await page.waitForTimeout(1200);

    logEvent({
      level: 'info',
      message:
        `Valor único aplicado direto no frame principal: ${singleValue}`,
      modalTitle,
    });

    return;
  }

  const trigger =
    mainFrame.locator(
      `${triggerSelector}[src*="botProcurar.png"]`
    ).first();

  await trigger.click({
    force: true,
  });

  let modalFrame =
    await findModalFrame(
      page,
      modalTitle,
      30000
    );

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

  // Verifica se o frame encontrado contém o input esperado; caso contrário, tenta localizar
  // um frame que tenha os seletores de busca (evita abrir modal errado, ex: Consulta de Empresas)
  try {
    const searchSelectors = Array.isArray(searchInputSelector) ? searchInputSelector : [searchInputSelector];

    async function frameHasAnyInput(f) {
      try {
        for (const sel of searchSelectors) {
          const c = await (f.locator ? f.locator(sel).count().catch(() => 0) : 0);
          if (c && c > 0) return true;
        }
      } catch (_) {}
      return false;
    }

    let hasInput = await frameHasAnyInput(modalFrame);

    if (!hasInput) {
      logEvent({ level: 'warn', message: `Frame modal não contém input esperado. Tentando fechar modais errados e reabrir (3 tentativas)...`, modalTitle });

      // tenta múltiplas vezes: fecha modal de empresas se aberto, re-clica trigger e procura frame que contenha os seletores
      for (let attempt = 1; attempt <= 3 && !hasInput; attempt++) {
        logEvent({ level: 'info', message: `Tentativa ${attempt} para obter modal correto: ${modalTitle}` });

        // detecta e fecha modal de Empresas se presente
        try {
          const emp = page.locator('div.dojoDialog:has-text("Consulta de Empresas"), .spwDialog:has-text("Consulta de Empresas"), table[id="empresa"]');
          if (await emp.count().catch(() => 0)) {
            logEvent({ level: 'info', message: 'Modal de Empresas detectado; fechando antes de reabrir o modal correto.' });
            const closeBtn = page.locator('.spwAlertaFechar, img[title="Fechar"], .ui-dialog-titlebar-close, input[value="FECHAR"], button:has-text("FECHAR")').first();
            if (await closeBtn.count().catch(() => 0)) {
              await closeBtn.click({ force: true }).catch(() => {});
              await page.waitForTimeout(600);
            }
          }
        } catch (e) {
          logEvent({ level: 'debug', message: 'Erro tentando fechar modal de Empresas', detail: String(e && e.message || e) });
        }

        // tenta reabrir o modal clicando novamente no trigger (relocaliza trigger para evitar stale)
        try {
          const retrigger = mainFrame.locator(`${triggerSelector}[src*="botProcurar.png"]`).first();
          if (await retrigger.count().catch(() => 0)) {
            await retrigger.click({ force: true }).catch(() => {});
            await page.waitForTimeout(600 + attempt * 300);
          }
        } catch (e) {
          logEvent({ level: 'debug', message: 'Falha ao reclicar no trigger do modal', detail: String(e && e.message || e) });
        }

        // procura em frames por qualquer seletor esperado
        for (const f of page.frames()) {
          if (await frameHasAnyInput(f)) {
            modalFrame = f;
            hasInput = true;
            logEvent({ level: 'info', message: `Trocando para frame que contém seletor (após tentativa ${attempt})`, frameUrl: f.url(), modalTitle });
            break;
          }
        }

        // verifica overlay no próprio page
        if (!hasInput) {
          for (const sel of searchSelectors) {
            try {
              const cc = await page.locator(sel).count().catch(() => 0);
              if (cc && cc > 0) {
                modalFrame = page;
                hasInput = true;
                logEvent({ level: 'info', message: `Trocando para page (overlay) que contém seletor ${sel} (após tentativa ${attempt})`, frameUrl: page.url(), modalTitle });
                break;
              }
            } catch (_) {}
          }
        }
      }

      if (!hasInput) {
        logEvent({ level: 'warn', message: `Ainda não encontrei input esperado no modal para: ${modalTitle}. Salvando debug.`, modalTitle });
        await saveDebugState('modal-wrong-frame').catch(() => {});
        throw new Error(`Modal encontrado mas não contém os inputs esperados: ${modalTitle}`);
      }
    }
  } catch (e) {
    // se alguma verificação falhar, rethrow para caller
    throw e;
  }

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
      fs.writeFileSync(htmlPath, modalHtml || '<!-- vazio -->');
    } catch (e) {
      logEvent({ level: 'warn', message: `Erro ao salvar HTML debug: ${e.message}` });
    }
    try {
      const preview = await modalFrame.evaluate(() => Array.from(document.querySelectorAll('tr')).slice(0,30).map(r => r.innerText)).catch(() => []);
      fs.writeFileSync(previewPath, JSON.stringify(preview, null, 2));
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
  for (const selector of searchInputSelector) {
    try {
      await modalFrame.locator(selector).first().waitFor({
        state: 'attached',
        timeout: 5000
      }).catch(() => {});
    } catch (_) {}
  }

  // =====================================================
  // MARCA TODOS OS VALORES
  // =====================================================

  // =====================================================
// MARCA TODOS OS VALORES
// =====================================================

  // Tenta com seletores padrão + alternativos
  const selectorsToTry = [
    ...searchInputSelector,
    // Fallbacks: input de qualquer tipo no modal
    'input[type="text"]:not([readonly])',
    'input:not([readonly])',
    'input[formatType="TEXT"]',
    'td.spwCelulaGrid input[type="text"]',
    'input[class*="Grid"]',
    'input'
  ];

  const modalInput =
    await resolveInput(
      modalFrame,
      selectorsToTry
    );

  for (const currentValue of entries) {

    logEvent({
      level: 'info',
      message:
        `Selecionando no modal: ${currentValue}`,
      modalTitle,
    });

    // ================================================
    // LIMPA FILTRO ANTERIOR
    // ================================================

    await modalInput.fill('')
      .catch(() => {});

    // Aguarda e garante que o input foi realmente limpo
    await page.waitForTimeout(500);

    // Verifica se foi realmente limpo
    const inputValue = await modalInput.inputValue().catch(() => '');
    if (inputValue && inputValue.trim()) {
      // Se ainda tem valor, tenta novamente
      await modalInput.evaluate((el) => {
        el.value = '';
        el.dispatchEvent(new Event('input', { bubbles: true }));
      });
      await page.waitForTimeout(300);
    }

    // ================================================
    // DIGITA NOVO VALOR
    // ================================================

    await fillLegacyInput(
      modalInput,
      currentValue
    );

    await modalInput.press('Enter')
      .catch(() => {});

    // Aguarda resultado da busca ser carregado (um pouco mais tolerante)
    await page.waitForTimeout(2300);

    // Aguarda a tabela ser atualizada
    try {
      await modalFrame.locator('td').first().waitFor({ timeout: 7000 });
    } catch (_) {}

    // ================================================
    // LOCALIZA LINHA (com fallback inteligente)
    // ================================================

    let cell = null;
    let cellFrame = null;

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
        // outro fallback: busca apenas os últimos dois tokens
        try {
          const tokens = String(currentValue || '').split(/\s+/).filter(Boolean);
          if (tokens.length > 1) {
            const lastTwo = tokens.slice(-2).join(' ');
            logEvent({ level: 'warn', message: `Tentando fallback tokens para '${currentValue}' -> '${lastTwo}'`, modalTitle });
            const res3 = await findRowByText(page, lastTwo, 20000, modalFrame);
            cell = res3.cell; cellFrame = res3.frame;
          }
        } catch (errFinal) {
          // grava HTML do modal e preview das primeiras linhas quando DEBUG_HTML ativo
          if (DEBUG_HTML) {
            try {
              const modalHtml = await modalFrame.content().catch(() => '<!-- erro ao capturar -->');
              const debugFile = path.join(SCREENSHOT_DIR, `debug-modal-${sanitizeFileName(modalTitle)}-${nowIso().replace(/[:.]/g, '-')}.html`);
              fs.writeFileSync(debugFile, modalHtml || '<!-- vazio -->');

              // também grava preview das primeiras 20 linhas normalizadas
              const preview = await modalFrame.evaluate(() => Array.from(document.querySelectorAll('tr')).slice(0,20).map(r => r.innerText));
              const previewFile = path.join(SCREENSHOT_DIR, `debug-modal-preview-${sanitizeFileName(modalTitle)}-${nowIso().replace(/[:.]/g, '-')}.json`);
              fs.writeFileSync(previewFile, JSON.stringify(preview, null, 2));
              logEvent({ level: 'info', message: `HTML do modal salvo em: ${debugFile}`, previewPath: previewFile });
            } catch (e) {
              logEvent({ level: 'warn', message: `Erro ao salvar HTML debug: ${e.message}` });
            }
          }

          // rethrow para o caller lidar com o erro
          throw errFinal || err;
        }
      }
    }

    // Tenta localizar a linha (tr) mais próxima e então o checkbox dentro dela
    let rowLocator = null;
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
        const xpath = `//tr[.//text()[contains(normalize-space(.), "${currentValue}")]]`;
        const frameForSearch = cellFrame || modalFrame;
        const maybeRow = frameForSearch.locator(`xpath=${xpath}`).first();
        if (await maybeRow.count().catch(() => 0)) rowLocator = maybeRow;
      } catch (_) {}
    }

    // Finalmente, busca o checkbox dentro da row ou dentro do próprio elemento
    let checkbox = null;
    if (rowLocator) {
      checkbox = rowLocator.locator('input[type="checkbox"]').first();
    } else {
      checkbox = cell.locator('input[type="checkbox"]').first();
    }

    try {
      const checked = await checkbox.isChecked().catch(() => false);
      if (!checked) {
        try {
          await checkbox.check();
        } catch (_) {
          try { await checkbox.click({ force: true }); } catch (_) {
            // último recurso: click via evaluate
            try {
              await checkbox.evaluate((el) => el.click());
            } catch (_) {}
          }
        }
      }

      logEvent({
        level: 'info',
        message:
          `Checkbox marcado: ${currentValue}`,
        modalTitle,
      });
    } catch (err) {
      await saveDebugState('checkbox-error');
      logEvent({ level: 'error', message: `Erro ao manipular checkbox para ${currentValue}: ${err.message}` });
      throw err;
    }

    await page.waitForTimeout(800);

    // ================================================
    // LIMPA FILTRO PARA PRÓXIMA SELEÇÃO
    // ================================================
    try {
      await modalInput.fill('');
      await page.waitForTimeout(300);
    } catch (_) {}
  }

  // =====================================================
  // BOTÃO SELECIONAR
  // =====================================================

  const { locator: selecionarBtn } = await findVisibleLocatorInFrames( page, '#pbSelecionar' ); 
  await selecionarBtn.click({ force: true, });

  await page.waitForTimeout(
    2000
  );

  await page.waitForTimeout(
    2000
  );

  logEvent({
    level: 'info',
    message:
      `Modal finalizado: ${entries.join(', ')}`,
    modalTitle,
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

  await openReportsPage(page);

  // =====================================================
  // CONFIGURA FILTROS
  // =====================================================
  await configureReportFilters({
    page,
    report
  });

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

  logEvent({
    level: 'info',
    message:
      `Relatório "${report.sheetName}": clicando Visualizar.`,
  });
  console.log(`Relatório "${report.sheetName}": acionando Visualizar (aguardando popup)...`);

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

  // =====================================================
  // LOOP URL FINAL
  // =====================================================

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

  const response =
    await context.request.get(
      reportUrl,
      {
        timeout: 120000,
      }
    );

  const contentType =
    String(
      response.headers()['content-type'] || ''
    ).toLowerCase();

  const buffer =
    await response.body();

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

  fs.writeFileSync(
    finalPdfPath,
    buffer
  );

  const stats =
    fs.statSync(finalPdfPath);
  logEvent({ level: 'info', message: 'PDF salvo.', bytes: stats.size, path: finalPdfPath });
  console.log(`Relatório "${report.sheetName}" salvo em: ${finalPdfPath} (${stats.size} bytes)`);

  if (stats.size < 5000) {
    throw new Error(`PDF inválido (${stats.size} bytes).`);
  }

  await popup.close().catch(() => {});

  logEvent({ level: 'info', message: `Relatório "${report.sheetName}" salvo com sucesso.`, path: finalPdfPath });

  return { path: finalPdfPath, bytes: stats.size };
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
      await updateReportStatusSerialized({ perReportEntry: { report: r.sheetName, props: { status: 'pending' } } });
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
      // marca início
      await updateReportStatusSerialized({ perReportEntry: { report: reportName, props: { status: 'running', startedAt: startedAtReport, index: index + 1 } } });
      logEvent({ level: 'info', message: 'Relatório iniciado.', report: reportName, current: index + 1, total: reports.length });

      const result = await generateSingleReport(context, page, report);

      const finishedAt = new Date().toISOString();

      await updateReportStatusSerialized({ perReportEntry: { report: reportName, props: { status: 'done', finishedAt, path: result && result.path ? result.path : (typeof result === 'string' ? result : null), bytes: result && result.bytes ? result.bytes : null } } });
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
      await updateReportStatusSerialized({ perReportEntry: { report: reportName, props: { status: 'error', finishedAt, error: String(err && err.message || err) } } });
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

  await page.waitForTimeout(800);
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

async function attachPageDebug(page) {
  page.on('pageerror', (err) => logEvent({ level: 'debug', message: 'Erro JS na página.', detail: String(err.message || err) }));
  page.on('console', (msg) => {
    if (['error', 'warning'].includes(msg.type())) {
      logEvent({ level: 'debug', message: `Console ${msg.type()} da página.`, detail: msg.text() });
    }
  });
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

  await context.storageState({ path: STATE_PATH });
  logEvent({ level: 'info', message: 'Estado da sessão salvo com sucesso.', statePath: STATE_PATH });
}

async function run() {
  await ensureDir(SCREENSHOT_DIR);
  await ensureDir(REPORT_OUTPUT_DIR);
  const browser = await chromium.launch({ headless: HEADLESS });
  const context = await browser.newContext(fs.existsSync(STATE_PATH) ? { storageState: STATE_PATH } : {});
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

    await context.storageState({ path: STATE_PATH });
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
