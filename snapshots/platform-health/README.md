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

## Segurança

O snapshot usa lista explícita de campos permitidos. Mesmo que uma auditoria interna passe a carregar novos campos sensíveis no futuro, eles não entram automaticamente no histórico.

Snapshots com `overall = degraded` ainda podem ser gravados como evidência, mas o comando termina com código de erro para impedir que degradação passe despercebida.
