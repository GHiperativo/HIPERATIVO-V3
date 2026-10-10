# Audit

Ferramentas de auditoria e verificação da Plataforma Hiperativo.

Por padrão, devem operar em modo somente leitura e retornar evidências objetivas: status, referências, contagens, versões, integridade e diferenças esperadas. Não devem revelar segredos para comprovar funcionamento.

## Ferramentas atuais

- `strava-credentials-status.mjs`: verifica metadados da credencial Strava sem consultar ou imprimir tokens. A saída pública é limitada a `status`, `vault_only`, `token_version` e `vault_reference`.
