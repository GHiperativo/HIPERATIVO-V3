# Auditoria de credenciais Strava

Executa uma verificação somente leitura sobre os metadados do registro correspondente em `public.tokens_strava`.

## Saída permitida

A ferramenta imprime apenas:

- `status`
- `vault_only`
- `token_version`
- `vault_reference`: `complete`, `partial` ou `missing`

Os UUIDs internos do Vault são usados somente para determinar se as duas referências existem. Eles não são exibidos. Tokens de acesso e refresh não são consultados.

## Variáveis necessárias

- `SUPABASE_URL`
- `SUPABASE_SERVICE_ROLE_KEY`
- `STRAVA_ATH_ID`

Esses valores devem ser fornecidos por ambiente seguro. Nunca devem ser gravados no repositório ou impressos em logs.

## Execução

```bash
npm run audit:strava-credentials
```

A ferramenta termina com código `1` quando `vault_only` não está ativo ou quando as referências do Vault não estão completas.

O workflow `Tooling Check` valida automaticamente os testes sempre que arquivos em `tools/` forem alterados.
