/**
 * HIPERATIVO V3 — MetricasFallbackSetup.gs
 * Ajuste seguro da aba MÉTRICAS para fallback manual por múltipla escolha.
 * Não altera integração, tokens ou callbacks Strava.
 */

function configurarFallbackManualMetricas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ws = ss.getSheetByName(H.SHEETS.METRICAS);
  if (!ws) throw new Error('Aba MÉTRICAS não encontrada.');

  const headers = [
    'Perfil Manual', 'Volume Manual', 'Intensidade Manual',
    'Origem dos Dados', 'Confiança', 'Observações'
  ];
  ws.getRange(2, H.MET.PERFIL_MAN, 1, headers.length).setValues([headers]);

  _dropdown(ws, H.MET.PERFIL_MAN, 3, 1000, ['Iniciante','Intermediário','Avançado','Retorno/lesão']);
  _dropdown(ws, H.MET.VOLUME_MAN, 3, 1000, ['Baixo','Moderado','Alto','Muito alto']);
  _dropdown(ws, H.MET.INTENS_MAN, 3, 1000, ['Leve','Moderado','Forte','Competitivo']);
  _dropdown(ws, H.MET.CONFIANCA, 3, 1000, ['Alta','Média','Baixa']);

  for (let r = 3; r <= 202; r++) {
    ws.getRange(r, H.MET.Z1_LENTO).setFormula(`=IFERROR(TEXT(INT(E${r}*1.20/60),"0")&":"&TEXT(MOD(ROUND(E${r}*1.20,0),60),"00"),"")`);
    ws.getRange(r, H.MET.Z1_RAPIDO).setFormula(`=IFERROR(TEXT(INT(E${r}*1.08/60),"0")&":"&TEXT(MOD(ROUND(E${r}*1.08,0),60),"00"),"")`);
    ws.getRange(r, H.MET.Z2_LENTO).setFormula(`=IFERROR(TEXT(INT(E${r}*1.04/60),"0")&":"&TEXT(MOD(ROUND(E${r}*1.04,0),60),"00"),"")`);
    ws.getRange(r, H.MET.Z2_RAPIDO).setFormula(`=IFERROR(TEXT(INT(E${r}*0.96/60),"0")&":"&TEXT(MOD(ROUND(E${r}*0.96,0),60),"00"),"")`);
    ws.getRange(r, H.MET.Z3_LENTO).setFormula(`=IFERROR(TEXT(INT(E${r}*0.94/60),"0")&":"&TEXT(MOD(ROUND(E${r}*0.94,0),60),"00"),"")`);
    ws.getRange(r, H.MET.Z3_RAPIDO).setFormula(`=IFERROR(TEXT(INT(E${r}*0.87/60),"0")&":"&TEXT(MOD(ROUND(E${r}*0.87,0),60),"00"),"")`);
    ws.getRange(r, H.MET.Z4_LENTO).setFormula(`=IFERROR(TEXT(INT(E${r}*0.84/60),"0")&":"&TEXT(MOD(ROUND(E${r}*0.84,0),60),"00"),"")`);
    ws.getRange(r, H.MET.Z5_MIN).setFormula(`=IFERROR(TEXT(INT(F${r}/60),"0")&":"&TEXT(MOD(F${r},60),"00"),"")`);
  }

  ws.getRange(2, H.MET.PERFIL_MAN, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#1D9E75')
    .setFontColor('#FFFFFF')
    .setHorizontalAlignment('center')
    .setWrap(true);

  _log('SYSTEM', 'INFO', 'configurarFallbackManualMetricas', 'Fallback manual de métricas configurado', '');
  try {
    SpreadsheetApp.getUi().alert('✅ Fallback manual de métricas configurado.');
  } catch (_) {}
}
