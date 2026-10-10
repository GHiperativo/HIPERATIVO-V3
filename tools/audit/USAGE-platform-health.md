# Platform Health

Auditoria agregada e somente leitura da fundação técnica da Plataforma Hiperativo.

## Verificações

- runtime Node.js esperado;
- presença dos arquivos mínimos de tooling;
- contagem de migrations SQL versionadas;
- alcance do Supabase quando credenciais seguras estiverem configuradas;
- saúde dos metadados de credenciais Strava quando `STRAVA_ATH_ID` estiver configurado.

## Segurança

A ferramenta não imprime `SUPABASE_SERVICE_ROLE_KEY`, tokens Strava, `ath_id` ou UUIDs internos do Vault.

Checks remotos são opcionais: em ambientes sem credenciais, aparecem como `skipped`. Se uma credencial remota estiver parcialmente configurada ou um check configurado falhar, o resultado global vira `degraded` e o processo termina com código 1.

## Execução

```bash
npm run audit:platform-health
```

Para incluir os checks remotos, forneça por ambiente seguro:

- `SUPABASE_URL`
- `SUPABASE_SERVICE_ROLE_KEY`
- `STRAVA_ATH_ID` para incluir a auditoria Strava

Nunca grave esses valores no repositório ou em logs.
