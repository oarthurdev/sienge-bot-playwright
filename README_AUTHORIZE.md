# Task Mode: authorize — Processo PM2

Este README descreve o processo que roda o modo `authorize` através do gerenciador PM2. O processo é definido em [ecosystem.config.cjs](ecosystem.config.cjs) e executa o arquivo [authorize.js](authorize.js).

**Descrição do processo**
- **Nome do app:** `sienge-autorizar-parcelas` (definido em `ecosystem.config.cjs`).
- **Script executado:** `authorize.js` — script Node.js que contém o laço/loop de monitoramento e execução das ações de autorização quando `task_mode` está configurado para `authorize`.
- **Interprete:** `node` — o processo é executado com o Node.js.

**Variáveis de ambiente configuradas no PM2**
- `NODE_ENV`: ambiente de execução (ex.: `production`).
- `PM2_LOOP`: indica que o script deve executar em loop (quando utilizado pelo próprio `authorize.js`).
- `PM2_INTERVAL_MS`: intervalo em milissegundos entre iterações do loop (ex.: `60000`).
- `PM2_STOP_ON_FATAL`: se `true`, indica comportamento de parada ante falhas fatais.
- `HEADLESS`: controla se o navegador (Playwright/Chromium) roda em modo headless (`true`/`false`).
- `DEBUG_HTML`: `true` para salvar HTML de debug — afeta salvamento de artefatos em disco.

Estas variáveis podem ser sobrescritas no ambiente do sistema ou no próprio comando `pm2 start`.

**Comandos PM2 úteis**
- Iniciar a aplicação definida no `ecosystem.config.cjs`:

```bash
pm2 start ecosystem.config.cjs
```

- Reiniciar (aplica alterações no código/env):

```bash
pm2 restart sienge-autorizar-parcelas
```

- Parar e remover do PM2:

```bash
pm2 stop sienge-autorizar-parcelas
pm2 delete sienge-autorizar-parcelas
```

- Ver logs em tempo real:

```bash
pm2 logs sienge-autorizar-parcelas --lines 200
```

**Observações operacionais**
- O processo é pensado para rodar continuamente em produção; `PM2_INTERVAL_MS` controla a frequência de checagem/execução.
- Garanta que as credenciais (variáveis `SIENGE_USERNAME`, `SIENGE_PASSWORD`) estejam corretas e seguras no ambiente do PM2 (use `pm2 ecosystem` ou variáveis de ambiente do sistema).
- Use `INSTANCE_ID` distinto para isolar arquivos de estado, logs e screenshots quando rodar múltiplas instâncias paralelas no mesmo host.
- Em caso de falhas repetidas, analise os logs apontados por `LOG_PATH` e o arquivo `STATUS_FILE` (padrão `/tmp/report-status.json`) para diagnóstico.

Para detalhes da implementação e variáveis adicionais usadas pelo fluxo de autorização, consulte `authorize.js` e [ecosystem.config.cjs](ecosystem.config.cjs).
