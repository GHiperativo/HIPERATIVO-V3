# Live test

A validação em produção deve confirmar apenas o estado seguro da credencial, sem exibir segredos ou identificadores internos.

Resultado esperado para uma credencial completamente migrada:

```json
{
  "status": "active",
  "vault_only": true,
  "token_version": 1,
  "vault_reference": "complete"
}
```
