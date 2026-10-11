# Platform Health

Auditoria agregada e somente leitura da fundação técnica da Plataforma Hiperativo.

## Verificações

- runtime Node.js esperado;
- presença dos arquivos mínimos de tooling;
- contagem de migrations SQL versionadas;
- alcance do Supabase por meio do RPC dedicado `public.platform_health_ping()`;
- saúde dos metadados de credenciais Strava quando `STRAVA_ATH_ID` estiver configurado.

## Saúde x cobertura

`overall` mede o resultado dos checks executados:

- `healthy`: nenhum check executado falhou;
- `degraded`: pelo menos um check executado falhou.

`verification` mede quanto da Plataforma foi efetivamente verificado:

- `full`: Supabase e Strava foram executados;
- `partial`: apenas parte dos checks remotos pôde ser executada;
- `local_only`: somente a fundação local foi verificada.

O bloco `coverage` detalha `local`, `supabase` e `strava` com valores `complete`, `incomplete` ou `skipped`. Assim, `overall: healthy` com `verification: partial` significa que nenhuma falha foi encontrada nos checks executados, mas a auditoria não cobriu toda a Plataforma.

Um erro retornado por um serviço após o check ter sido realmente executado mantém a cobertura daquele serviço como `complete`; nesse caso `overall` fica `degraded`. Falta de configuração necessária para executar o check conta como cobertura `incomplete`.

## Supabase reachability

O check genérico de Supabase não consulta mais `tokens_strava` nem outra tabela de domínio. Ele chama `POST /rest/v1/rpc/platform_health_ping`, cuja função retorna somente `ok` e não lê dados da aplicação.

A função é `SECURITY INVOKER`, `STABLE` e possui `search_path = ''`. Dessa forma, o health check do Supabase fica desacoplado do módulo Strava.

## Segurança

A ferramenta não imprime `SUPABASE_SERVICE_ROLE_KEY`, tokens Strava, `ath_id` ou UUIDs internos do Vault.

Checks remotos são opcionais. Ausência deliberada de configuração pode produzir `skipped` sem degradar `overall`, mas reduz `verification`. Configuração parcial incapaz de executar um check produz erro de saúde e cobertura `incomplete`.

## Execução

```bash
npm run audit:platform-health
```

Para incluir os checks remotos, forneça por ambiente seguro:

- `SUPABASE_URL`
- `SUPABASE_SERVICE_ROLE_KEY`
- `STRAVA_ATH_ID` para incluir a auditoria Strava

Nunca grave esses valores no repositório ou em logs.
