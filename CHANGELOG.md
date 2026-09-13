# CHANGELOG

## [0.2.1] — 2026-09-13
### Added
- Rotina segura `configurarSpreadsheetId()` para registrar o ID da planilha ativa em `PropertiesService` sem hardcode.
- Diagnóstico `diagnosticarPropriedadesEssenciais()` para informar apenas se `ADMIN_EMAIL` e `SPREADSHEET_ID` estão configurados, sem exibir seus valores.
- Implementação espelhada em `apps-script/live/SetupProperties.js` e `apps-script/src/SetupProperties.gs`.

### Notes
- Alteração preparada no repositório. O projeto Apps Script vivo ainda exige sincronização autenticada via `clasp push` ou atualização pelo editor do Apps Script.

## [0.2.0] — 2026-06-08
### Added
- Campos de fallback manual na aba `📈 MÉTRICAS`: perfil, volume, intensidade, origem, confiança e observações.
- Cálculo de métricas com estimativa guiada por múltipla escolha quando não há corridas recentes válidas.
- Constantes `H.MET` para reduzir risco de erro por coluna hardcoded.
- Script complementar `MetricasFallbackSetup.gs` para configurar os campos manuais de métricas sem alterar o fluxo Strava.

### Changed
- Documentação atualizada com resumo do projeto, estrutura de planilhas e melhorias sugeridas.

### Notes
- Integração Strava não foi alterada nesta versão.

## [0.1.0] — 2026-06-05
### Added
- Estrutura inicial do repositório
- README, docs, templates
