# Platform Health snapshots

Histórico versionado e sanitizado do estado técnico da PLATAFORMA HIPERATIVO.

## Objetivo

Permitir comparação objetiva entre execuções do `Platform Health` sem versionar credenciais, tokens, IDs internos do Vault ou outros segredos.

## Captura

```bash
npm run snapshot:platform-health
```

O comando grava um JSON em `snapshots/platform-health/` usando timestamp UTC no nome do arquivo.

## Comparação

```bash
npm run snapshot:compare -- snapshots/platform-health/<anterior>.json snapshots/platform-health/<atual>.json
```

A comparação ignora apenas `captured_at` e informa as demais mudanças por caminho.

## Detecção de regressão

Comparação explícita:

```bash
npm run snapshot:regression -- snapshots/platform-health/<anterior>.json snapshots/platform-health/<atual>.json
```

Comparação automática dos dois snapshots mais recentes:

```bash
npm run snapshot:regression -- --auto
```

O `Tooling Check` executa o modo automático sempre que `tools/`, `package.json`, o workflow ou `snapshots/platform-health/` mudam.

Falham o CI, entre outros:

- `overall` deixar de ser `healthy`;
- runtime Node.js incompatível;
- arquivos obrigatórios desaparecerem;
- quantidade de migrations versionadas diminuir;
- Supabase ou credenciais Strava passarem de `ok` para `error`;
- `vault_only` deixar de ser verdadeiro;
- referências do Vault deixarem de estar completas;
- `token_version` diminuir;
- baseline agregado perder cobertura completa de Vault.

Checks remotos que passam de `ok` para `skipped` por ausência de credenciais no ambiente geram `warning`, não falha. Assim o CI público não precisa armazenar segredos para avaliar o histórico.

## Segurança

O snapshot usa lista explícita de campos permitidos. Mesmo que uma auditoria interna passe a carregar novos campos sensíveis no futuro, eles não entram automaticamente no histórico.

Snapshots com `overall = degraded` ainda podem ser gravados como evidência, mas o comando termina com código de erro para impedir que degradação passe despercebida.
