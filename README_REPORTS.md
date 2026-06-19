# Relatórios — `script.js`

Este README documenta o script de geração de relatórios `script.js`, responsável por automatizar a navegação no Sienge, filtrar períodos e coletar/autorizar itens de relatórios.

**Visão geral**
- **Arquivo principal:** [script.js](script.js)
- **Objetivo:** autenticar no Sienge, navegar até páginas de relatório, aplicar filtros de período, processar itens (coleta/validação/autorizações) e persistir estado e logs para retomada e auditoria.

**Variáveis de ambiente principais**
- `SIENGE_BASE_URL`: URL base do Sienge (obrigatório).
- `SIENGE_USERNAME` / `SIENGE_PASSWORD`: credenciais de login (obrigatório).
- `INSTANCE_ID`: identificador da instância para distinguir arquivos gerados (recomendado).
- `HEADLESS`: `true`/`false` — controla se o navegador roda em headless.
- `STATE_PATH`: caminho do arquivo JSON de estado (padrão: `sienge-storage-state-<INSTANCE_ID>.json`).
- `LOG_PATH`: caminho do log de execução (padrão: `sienge-authorize-log-<INSTANCE_ID>.json`).
- `STATUS_FILE`: arquivo de status global para várias instâncias (padrão: `/tmp/report-status.json`).
- `MODAL_CACHE_PATH`: caminho do cache de modais (padrão: `.modal-cache-<INSTANCE_ID>.json`).
- `SINGLE_REPORT` (ou `--single` / `-s`): processa somente um relatório específico.
- `TASK_MODE`: modo de operação (ex.: `authorize`, `watch`).
- `SAVE_SHOTS` / `DEBUG_HTML`: salvamento de screenshots/HTML para depuração.

**Comportamento principal**
1. Valida e carrega configurações/variáveis de ambiente.
2. Inicializa o navegador via Playwright (`chromium`).
3. Restaura estado anterior a partir de `STATE_PATH` e carrega `MODAL_CACHE_PATH` para evitar reprocessamento de modais.
4. Navega até `TARGET_PAGE_URL` ou `REPORT_FILTER_PAGE_URL` e aplica filtros (`REPORT_PERIOD_START`, `REPORT_PERIOD_END`).
5. Itera sobre os relatórios/itens encontrados, processando conforme `TASK_MODE` (ex.: autorizar parcelas quando `authorize`).
6. Persiste progresso incrementalmente em `STATE_PATH` e grava eventos em `LOG_PATH`.
7. Em caso de `SAVE_SHOTS` ou `DEBUG_HTML`, salva screenshots e HTML em `SCREENSHOT_DIR`.

**Recomendações de execução**
- Execução simples:

```bash
SIENGE_BASE_URL=https://sienge.example.com \
SIENGE_USERNAME=usuario SIENGE_PASSWORD=senha node script.js
```

- Executar processando somente um relatório:

```bash
SINGLE_REPORT=123 node script.js
```

- Executar com identificação de instância e captura de telas:

```bash
INSTANCE_ID=2 SAVE_SHOTS=true SCREENSHOT_DIR=./screenshots-2 \
SIENGE_BASE_URL=https://sienge.example.com \
SIENGE_USERNAME=usuario SIENGE_PASSWORD=senha node script.js
```

**Arquivos gerados**
- `STATE_PATH`: arquivo JSON com progresso e metadados para retomar execução.
- `LOG_PATH`: log detalhado em JSON com eventos, erros e ações.
- `SCREENSHOT_DIR`: screenshots/HTML de páginas quando ativado.
- `MODAL_CACHE_PATH`: cache local de conteúdo de modais.

**Concorrência e locking**
- O script possui um semáforo em-processo para serializar manipulação de modais (`_modalSemaphore`), evitando colisões quando há múltiplas tarefas no mesmo processo.
- Para paralelismo entre processos, defina `INSTANCE_ID` distinto para cada processo.

**Tratamento de erros e retomada**
- Antes de abortar, o script geralmente persiste o estado atual para `STATE_PATH`, permitindo retomadas.
- Logs no `LOG_PATH` e o `STATUS_FILE` ajudam na identificação de pontos de falha e estratégias de retry.

Para detalhes de implementação e lógica de retry, consulte o código em [script.js](script.js).
