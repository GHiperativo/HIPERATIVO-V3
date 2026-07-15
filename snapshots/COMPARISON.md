# Snapshot comparison

## Scope

This report compares:

- GitHub `main` before this snapshot;
- Apps Script editor state (`HEAD`);
- immutable Apps Script version 10;
- immutable Apps Script version 38.

It is descriptive only. No production source or deployment was changed.

## Inventory

| State | Files | Functions | Notes |
|---|---:|---:|---|
| GitHub `apps-script/src` | 14 | 143 | Previous repository state |
| Apps Script version 10 | 11 | 160 | Deployment referenced by the athlete OAuth links inspected in the spreadsheet |
| Apps Script version 38 | 20 | 203 | Current maintenance deployment family |
| Apps Script editor `HEAD` | 20 | 287 | Current saved editor state |

## Router entry points

### Version 10

- `doGet`: `WebApp.js`
- `doPost`: `Webhooks.js`
- OAuth callback: `_processarCallbackOAuth` in `WebApp.js`
- automatic import: `triggerImportacaoAutomatica` in `Strava.js`

### Version 38

- `doGet`: `WebApp.js`
- OAuth callback: `_processarCallbackOAuth` in `WebApp.js`
- automatic import: `triggerImportacaoAutomatica` in `Strava.js`
- queue processing: `processarFilaStrava` in `Queue.js`
- Supabase token fallback: `supaGetRefreshToken` in `Supabase.js`

### Editor HEAD

- `doGet`: `WebApp.js`
- OAuth callback: `_processarCallbackOAuth` in `WebApp.js`
- queue processing: `processarFilaStrava` in `Queue.js`
- two definitions of `triggerImportacaoAutomatica`:
  - `Strava.js`
  - `ImportacaoRawConvertida.js`

The duplicate trigger definition is preserved exactly as found. Its behavior
must be resolved in a later change; this snapshot does not choose one.

## Duplicate global functions

- GitHub main: duplicate `doGet` in `Strava.gs` and `WebApp.gs`.
- Version 10: duplicate `_linhaAtividadeStrava_` in `Strava.js` and
  `Webhooks.js`.
- Version 38: duplicate `diagnosticoRapido` and `fixWebAppUrl`.
- Editor HEAD: duplicate `triggerImportacaoAutomatica` and
  `enviarLinkStravaDesconectados`.

These are findings, not changes.

## Version 38 versus editor HEAD

Unchanged files:

- `atleta.html`
- `Dashboard.js`
- `Metricas.js`
- `MetricasFallbackSetup.js`
- `SISRUN.js`

Changed common files:

- `Analise.js`
- `appsscript.json`
- `cadastro.html`
- `Cadastro.js`
- `Config.js`
- `Queue.js`
- `Strava.js`
- `WebApp.js`

Present only in version 38:

- `05_Setup.js`
- `06_Ranking.js`
- `07_InputManual.js`
- `08_Desafios.js`
- `Ajustes.js`
- `Restaurar.js`
- `Supabase.js`

Present only in editor HEAD:

- `Correcoes.js`
- `ImportacaoRawConvertida.js`
- `Normalizar.js`
- `Sem título.js`
- `Setup.js`
- `StravaPipeline.js`
- `SupaSync.js`

## Credential scan

The extracted source contains no literal JWT, Google API key, GitHub token,
private key, refresh token, Supabase service-role key, or literal Strava client
secret matching the pre-commit scan.

Credentials are read from `PropertiesService`, which is not exported by
`clasp`. Therefore the OAuth credentials and working refresh tokens were not
copied into this repository.

Security note for later remediation: diagnostic functions in the extracted
code log variables named `secret`. No secret value is embedded in the
snapshot, but executing those diagnostics may expose a configured value in
Apps Script execution logs.

## Validation

- All extracted JavaScript files pass `node --check`.
- Every source file is listed in `SHA256SUMS.txt`.
- Generated `.clasp.json` files were excluded to prevent accidental pushes.
- No `clasp push`, deployment, trigger execution, spreadsheet mutation, or
  Supabase mutation occurred.

