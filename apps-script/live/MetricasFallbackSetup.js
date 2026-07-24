/**
 * HIPERATIVO V3 — MetricasFallbackSetup.gs
 * Ajuste seguro da aba MÉTRICAS para fallback manual por múltipla escolha.
 * Não altera integração, tokens ou callbacks Strava.
 */

function configurarFallbackManualMetricas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ws = ss.getSheetByName(H.SHEETS.METRICAS);
  if (!ws) throw new Error('Aba MÉTRICAS não encontrada.');

  _garantirEstruturaMetricas_(ws);

  _dropdown(ws, H.MET.PERFIL_MAN, 3, 1000, ['Iniciante','Intermediário','Avançado','Retorno/lesão']);
  _dropdown(ws, H.MET.INTENS_MAN, 3, 1000, ['Leve','Moderado','Forte','Competitivo']);

  _log('SYSTEM', 'INFO', 'configurarFallbackManualMetricas', 'Campos manuais preservados sem criar estimativas fisiológicas', '');
  try {
    SpreadsheetApp.getUi().alert('✅ Campos manuais configurados sem gerar métricas artificiais.');
  } catch (_) {}
}
