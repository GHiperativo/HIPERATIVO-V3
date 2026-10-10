# Supabase Migration Ledger

Auditoria de consistência entre o histórico de migrations em produção e os arquivos versionados da Plataforma Hiperativo.

## Decisão de baseline

O histórico remoto anterior e igual a `20261010234625` é tratado como legado congelado. Ele é registrado em `supabase/migration-baseline.json` apenas por metadados seguros: versão, nome, MD5 do SQL e quantidade de caracteres.

O SQL histórico completo não é copiado automaticamente para o repositório público. Isso evita fabricar migrations antigas e reduz o risco de versionar literais sensíveis que possam ter existido em migrations operacionais.

Dois arquivos locais anteriores ao corte são dívida técnica conhecida e aparecem como `warning`:

- `20260722155300_ensure_atleta_before_strava_token.sql`
- `20260722204500_strava_webhook_event_queue.sql`

## Regra após o corte

Para qualquer migration com versão posterior a `20261010234625`:

- migration existente em produção sem arquivo local correspondente é erro;
- arquivo e produção com hashes diferentes é erro;
- arquivo local ainda não aplicado em produção é warning;
- alteração de uma migration congelada que também exista localmente é erro.

## Acesso remoto

O RPC público `public.platform_migration_ledger()` é `SECURITY INVOKER` e executável apenas por `service_role`. A leitura privilegiada da tabela interna de migrations fica em `private.platform_migration_ledger_internal()`, com `SECURITY DEFINER`, `search_path = ''` e ACL restrita.

A resposta contém somente:

- `version`
- `name`
- `sql_md5`
- `sql_chars`

Nenhum statement SQL, token, segredo ou dado de negócio é retornado.

## Execução

```bash
npm run audit:supabase-migrations
```

Sem `SUPABASE_URL` e `SUPABASE_SERVICE_ROLE_KEY`, a checagem remota fica `skipped` e o contrato local continua sendo validado. Com essas variáveis fornecidas por ambiente seguro, o comando também compara o ledger remoto.

Nunca grave a service role no repositório ou em logs.
