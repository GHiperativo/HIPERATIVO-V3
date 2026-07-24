/**
 * Mantém à vista apenas as abas de trabalho do treinador.
 * Nada é excluído: bases, tokens, logs, backups e protótipos permanecem
 * acessíveis ao sistema e podem ser reexibidos pelo editor da planilha.
 */
function organizarAbasOperacionais() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const visiveis = [
    '📊 PAINEL',
    '👤 CADASTRO',
    '📲 WHATSAPP STRAVA',
    '🏃 ATIVIDADES',
    '📈 MÉTRICAS',
    '📅 PLANO SEMANAL',
    '💬 FEEDBACK',
    '🔬 ANÁLISE',
    '🏆 RANKING',
    '📝 INPUT MANUAL'
  ];
  const ocultas = [
    '🏃 ATIVIDADES_CONVERTIDAS',
    '🏃 STRAVA_RAW',
    '📈 MÉTRICAS_BETA',
    '📋 AUDITORIA',
    '🔴 ERROS',
    '📉 GRÁFICOS',
    '🔐 TOKENS',
    '⚙️ CONFIG',
    '🏆 RANKING COMPLETO',
    'STRAVA_WEBHOOK_QUEUE',
    '📡 STRAVA STATUS',
    '📨 STRAVA ENVIOS',
    '🎯 DESAFIOS',
    '📊 DESAFIOS_PROGRESSO',
    '👤 CADASTRO_BACKUP_PRE_NORMALIZACAO',
    '🔐 TOKENS_BACKUP_PRE_NORMALIZACAO',
    '🏃 ATIVIDADES_BACKUP_PRE_NORMALIZACAO',
    '📈 MÉTRICAS_BACKUP_PRE_NORMALIZACAO',
    '📡 STRAVA STATUS_BACKUP_PRE_NORMALIZACAO'
  ];

  visiveis.forEach(nome => {
    const sh = ss.getSheetByName(nome);
    if (sh) sh.showSheet();
  });
  ocultas.forEach(nome => {
    const sh = ss.getSheetByName(nome);
    if (sh && !visiveis.includes(nome)) sh.hideSheet();
  });

  const ordem = visiveis.map(nome => ss.getSheetByName(nome)).filter(Boolean);
  ordem.forEach((sh, i) => {
    if (sh.getIndex() === i + 1) return;
    try {
      ss.setActiveSheet(sh);
      ss.moveActiveSheet(i + 1);
    } catch (e) {
      // A visibilidade é a parte crítica. Uma seleção sobre célula mesclada
      // pode impedir a reordenação no Google Sheets, sem afetar os dados.
      _log('SISTEMA', 'AVISO', 'organizarAbasOperacionais',
        'Aba ' + sh.getName() + ' mantida na posição atual: ' + e.message, '');
    }
  });
  _log('SISTEMA', 'INFO', 'organizarAbasOperacionais', 'Abas técnicas ocultadas sem exclusão.', '');
  return { visiveis: ordem.length, ocultas: ocultas.length };
}
