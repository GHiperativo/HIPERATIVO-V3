# Strava — runbook de recuperação segura

**Status:** trilha VERDE. Este documento não autoriza `clasp push`, deploy, criação/remoção de trigger, refresh real de token ou OAuth de atleta.

## Evidência observada em 10/10/2026

- Supabase contém 7.412 atividades de 24 atletas.
- Última atividade/importação observada: 23/09/2026.
- Existem 29 registros em `tokens_strava`; todos estavam expirados na leitura de 10/10/2026.
- Última atualização de token observada: 23/09/2026.
- Existem 6 eventos em `strava_eventos_webhook`; último recebido/processado em 29/07/2026.
- A Edge Function `strava-webhook` está ativa e a configuração registra subscription/callback/Apps Script configurados.

## O que o código versionado atual já faz

No `main` atual:

- `apps-script/live/SupaSync.js` contém `renovacaoProativaTokens()` e `instalarAcionadorRenovacao()`.
- `renovacaoProativaTokens()` persiste credenciais pelo fluxo central quando disponível e mantém cópia no Supabase.
- `apps-script/live/WebApp.js` contém `configurarTriggers(silencioso)` com configuração seletiva de automações essenciais.
- `configurarTriggers()` inclui importação a cada 3h, renovação a cada 4h, monitor diário, formulário SHE e limpeza semanal.
- `diagnosticarTriggersEssenciais()` é somente leitura e não cria/remove triggers.

Portanto, **não reescrever OAuth/refresh antes de verificar a versão efetivamente publicada**.

## Hipótese operacional prioritária

O código versionado e o estado vivo podem estar desalinhados. Possibilidades:

1. versão publicada do Apps Script anterior ao `main` atual;
2. triggers essenciais não instalados para a conta ativa;
3. trigger instalado por outra conta/owner e não visível ao operador atual;
4. execução falhando por credencial/configuração apesar de o código estar correto;
5. webhook Strava e importação periódica com falhas independentes.

Nenhuma dessas hipóteses deve ser convertida em correção sem evidência.

## Diagnóstico permitido antes de produção

### A. Read-only

1. Confirmar autenticação `clasp`.
2. Executar `clasp deployments` e registrar deployment/version sem alterar nada.
3. Comparar arquivos puxados do projeto vivo com `apps-script/live/` do GitHub.
4. Se a função existir no projeto vivo, executar **somente** `diagnosticarTriggersEssenciais()` via mecanismo de execução autorizado.
5. Ler logs de execução sem alterar propriedades.

### B. Não executar como diagnóstico

- `configurarTriggers()` porque recria automações.
- `instalarAcionadorRenovacao()` porque altera triggers.
- `desativarTriggers()`.
- `runAuditarTriggers()` de snapshots antigos: apesar do nome, versões antigas podem excluir triggers.
- `renovacaoProativaTokens()` porque chama a API Strava e altera tokens reais.
- qualquer função que faça OAuth, importação ou escrita em Supabase/Sheets.

## Gate para reparo live

Antes de qualquer ação live, registrar:

- versão publicada;
- lista de triggers essenciais presentes/ausentes;
- owner/conta quando identificável;
- erro de execução mais recente;
- diferença entre projeto vivo e `main`;
- ação mínima proposta;
- rollback.

Somente então pedir um GO específico, por exemplo:

`GO STRAVA — RECONFIGURAR TRIGGERS`

ou

`GO STRAVA — PUBLICAR VERSÃO VALIDADA`

## Critério de saída

Considerar o P0 Strava encerrado apenas quando houver uma destas evidências:

1. atividade nova chegando automaticamente e tokens com expiração futura; ou
2. causa exata documentada e correção preparada com gate de produção definido.

Webhook e importação periódica devem ser validados separadamente. Um não prova o funcionamento do outro.
