# Geração de Relatórios (script.js)

Este documento descreve, de forma técnica, o propósito e o comportamento do script de geração de relatórios presente no repositório (`script.js`). O objetivo principal do script é navegar pela interface web do Sienge, filtrar e gerar relatórios periódicos, salvar estados locais, capturar screenshots e registrar logs de execução.

**Visão Geral**
- **Arquivo principal**: `script.js` - script Node.js que usa Playwright para automação do navegador.
- **Função**: autenticar-se no Sienge, navegar até páginas específicas de relatórios, aplicar filtros de data/período, executar ações de autorização/checagem e gravar resultados em arquivos de estado e logs.

**Fluxo de execução**
1. Carrega variáveis de ambiente e configurações (URL base, credenciais, paths).
2. Inicializa Playwright (Chromium) com opções `HEADLESS` configuráveis.
3. Carrega o arquivo de estado local (`STATE_PATH`) e o cache de modais (`MODAL_CACHE_PATH`).
4. Abre a página alvo (`TARGET_PAGE_URL` ou `REPORT_FILTER_PAGE_URL`) e aplica filtros de período (`REPORT_PERIOD_START`, `REPORT_PERIOD_END`, etc.).
5. Para cada relatório encontrado (ou um único quando `SINGLE_REPORT` usado), executa o fluxo de coleta/validação/autorizações conforme `TASK_MODE`.
6. Persiste resultados incrementais em `STATE_PATH` e registra eventos/erros em `LOG_PATH` e `STATUS_FILE`.
7. Opcionalmente salva screenshots e HTML do estado da página quando `SAVE_SHOTS` ou `DEBUG_HTML` estiverem ativados.

**Principais variáveis de ambiente**
- `SIENGE_BASE_URL`: URL base do Sienge.
- `SIENGE_USERNAME` / `SIENGE_PASSWORD`: credenciais de autenticação.
- `INSTANCE_ID`: identificador da instância para isolar arquivos (`STATE_PATH`, `SCREENSHOT_DIR`, `LOG_PATH`).
- `HEADLESS`: `true`/`false` para rodar o navegador em modo headless.
- `STATE_PATH`: caminho para arquivo JSON de estado (padrão: `sienge-storage-state-<INSTANCE_ID>.json`).
- `LOG_PATH`: caminho para o log de execução (padrão: `sienge-authorize-log-<INSTANCE_ID>.json`).
- `STATUS_FILE`: arquivo de status global (padrão: `/tmp/report-status.json`).
- `MODAL_CACHE_PATH`: caminho para cache de modais (padrão: `.modal-cache-<INSTANCE_ID>.json`).
- `SINGLE_REPORT` ou `--single` / `-s`: processa apenas um relatório específico.
- `TASK_MODE`: modo de operação (ex.: `authorize`, `watch`, etc.).
- `SAVE_SHOTS`, `DEBUG_HTML`: controlam captura de telas e salvamento de HTML para debug.

Use backticks para passar variáveis de ambiente no shell, por exemplo:

```bash
INSTANCE_ID=1 SIENGE_BASE_URL=https://sienge.example.com \
SIENGE_USERNAME=usuario SIENGE_PASSWORD=senha node script.js
```

**Arquivos de estado e logs**
- `STATE_PATH`: contém o progresso, itens processados e metadados para retomar execuções.
- `LOG_PATH`: registro estruturado (JSON) de eventos, erros e ações executadas.
- `STATUS_FILE`: arquivo compartilhado para agregar progresso entre várias instâncias do script.
- `MODAL_CACHE_PATH`: cache local de conteúdo de modais para evitar reaberturas desnecessárias e acelerar execuções.

**Concorrência e semáforos**
O script implementa um semáforo em processo para serializar interações com modais (`_modalSemaphore`) — isso evita que múltiplas interações de modais conflitem quando o mesmo processo roda várias tarefas em paralelo. Para rodar execuções paralelas isoladas, utilize `INSTANCE_ID` distinto por processo.

**Modos de operação relevantes**
- `TASK_MODE=authorize` (padrão): realiza ações para autorizar parcelas/itens nos relatórios.
- `TASK_MODE=watch`: observa e registra mudanças de status sem tomar ações de autorização.

**Capturas e debugging**
- Quando `SAVE_SHOTS=true` ou `DEBUG_HTML=true`, o script salva screenshots e/ou HTML na pasta `SCREENSHOT_DIR` (padrão `screenshots-<INSTANCE_ID>`).
- Logs e arquivos salvos facilitam reprodutibilidade e análise pós-execução.

**Erros e retry**
O script registra erros no `LOG_PATH` e faz persistência de estado antes de abortar para permitir retomar processamento posterior. O comportamento de retry depende de pontos específicos no fluxo (por exemplo, reautenticação ou recarregamento de páginas). Consulte o código em `script.js` para estratégias detalhadas de retry.

**Execução e exemplos**
- Execução simples (headless):

```bash
SIENGE_BASE_URL=https://sienge.example.com \
SIENGE_USERNAME=usuario SIENGE_PASSWORD=senha node script.js
```

- Execução com captura de telas e instância identificada:

```bash
INSTANCE_ID=2 SAVE_SHOTS=true SCREENSHOT_DIR=./screenshots-2 \
SIENGE_BASE_URL=https://sienge.example.com \
SIENGE_USERNAME=usuario SIENGE_PASSWORD=senha node script.js
```

**Pontos importantes para manutenção**
- Atualize `TARGET_PAGE_URL` e `REPORT_FILTER_PAGE_URL` caso as rotas internas do Sienge mudem.
- Garanta que as credenciais e permissões sejam adequadas para as ações de autorização automatizadas.
- Verifique o tamanho dos arquivos `STATE_PATH`/`LOG_PATH` em execuções longas e faça rotação/backup quando necessário.

---

**Documentação específica**
- Processos PM2 (modo `authorize`): consulte [README_AUTHORIZE.md](README_AUTHORIZE.md).
- Script de relatórios (`script.js`): consulte [README_REPORTS.md](README_REPORTS.md).

Para detalhes de implementação (funções, retry logic, tratamento de modais), consulte [script.js](script.js).
