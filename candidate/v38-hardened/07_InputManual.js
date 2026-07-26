/**
 * ═══════════════════════════════════════════════════════════════════════
 * HIPERATIVO V3 — 07_InputManual.gs
 * Aba de input manual de atividades (atletas sem GPS / sem Strava)
 * ═══════════════════════════════════════════════════════════════════════
 */

// ── CRIAR ABA INPUT MANUAL ────────────────────────────────────────────────────
function criarAbaInputManual() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  let ws = ss.getSheetByName(H.SHEETS.INPUT_MANUAL);
  if (ws) ss.deleteSheet(ws);
  ws = ss.insertSheet(H.SHEETS.INPUT_MANUAL);

  // Título
  ws.getRange(1, 1, 1, 15).mergeAcross()
    .setValue('📝 INPUT MANUAL DE ATIVIDADES — HIPERATIVO V3')
    .setFontFamily('Arial').setFontSize(13).setFontWeight('bold')
    .setFontColor('#FFFFFF').setBackground('#001F3F')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  ws.setRowHeight(1, 36);

  // Cabeçalhos
  const cabecalhos = [
    'Data', 'ID Atleta', 'Tipo de Atividade', 'Duração (min)',
    'Distância (km)', 'Pace médio (MM:SS)', 'FC Média',
    'FC Máxima', 'Zona sentida', 'RPE (esforço)',
    'Qualidade do treino', 'Terreno', 'Período', 'Observações', 'Importado?'
  ];
  ws.getRange(2, 1, 1, cabecalhos.length).setValues([cabecalhos])
    .setFontFamily('Arial').setFontSize(10).setFontWeight('bold')
    .setBackground('#003D7A').setFontColor('#FFFFFF')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  ws.setRowHeight(2, 26);

  // Zebra nas linhas de dados
  for (let r = 3; r <= 202; r++) {
    ws.getRange(r, 1, 1, 15)
      .setBackground(r % 2 === 0 ? '#FFFFFF' : '#F0F8FF')
      .setFontFamily('Arial').setFontSize(10);
    ws.setRowHeight(r, 22);
  }

  const S = 3, E = 202;

  // Dropdown Atleta_ID — lista dinâmica do CADASTRO
  const wsCad = ss.getSheetByName(H.SHEETS.CADASTRO);
  if (wsCad) {
    const idsAtletas = wsCad.getDataRange().getValues().slice(2)
      .filter(function(r) { return r[H.CAD.ID - 1]; })
      .map(function(r) { return String(r[H.CAD.ID - 1]); });
    if (idsAtletas.length > 0) {
      _dropdown(ws, 2, S, E, idsAtletas);
    }
  }

  _dropdown(ws, 3, S, E, ['Corrida','Caminhada','Ciclismo','Natação','Funcional','Musculação','Outro']);

  // Duração: inteiro 1-600 min
  ws.getRange(S, 4, E - S + 1, 1)
    .setDataValidation(SpreadsheetApp.newDataValidation()
      .requireNumberBetween(1, 600).setAllowInvalid(false).build());

  // Distância: 0-200 km (opcional)
  ws.getRange(S, 5, E - S + 1, 1)
    .setDataValidation(SpreadsheetApp.newDataValidation()
      .requireNumberBetween(0, 200).setAllowInvalid(true).build());

  // FC Média e FC Máxima: 40-220 bpm
  ws.getRange(S, 7, E - S + 1, 1)
    .setDataValidation(SpreadsheetApp.newDataValidation()
      .requireNumberBetween(40, 220).setAllowInvalid(true).build());
  ws.getRange(S, 8, E - S + 1, 1)
    .setDataValidation(SpreadsheetApp.newDataValidation()
      .requireNumberBetween(40, 220).setAllowInvalid(true).build());

  _dropdown(ws, 9,  S, E, ['Z1 - Muito leve','Z2 - Leve','Z3 - Moderado','Z4 - Difícil','Z5 - Máximo']);
  _dropdown(ws, 10, S, E, ['1','2','3','4','5','6','7','8','9','10']);
  _dropdown(ws, 11, S, E, ['Ruim','Regular','Boa','Excelente']);
  _dropdown(ws, 12, S, E, ['Asfalto','Terra/Trail','Pista','Esteira','Misto']);
  _dropdown(ws, 13, S, E, ['Manhã','Tarde','Noite']);

  // Col 15: Checkbox "Importado?"
  ws.getRange(S, 15, E - S + 1, 1)
    .setDataValidation(SpreadsheetApp.newDataValidation()
      .requireCheckbox().setAllowInvalid(false).build())
    .setValue(false);

  // Formato de data na col 1
  ws.getRange(S, 1, E - S + 1, 1).setNumberFormat('DD/MM/YYYY');

  // Larguras das colunas
  const larguras = [12, 14, 18, 12, 13, 16, 9, 9, 18, 12, 18, 12, 10, 30, 11];
  larguras.forEach(function(w, i) { ws.setColumnWidth(i + 1, w * 7); });

  ws.setFrozenRows(2);

  _log('SYSTEM', 'INFO', 'criarAbaInputManual', 'Aba "' + H.SHEETS.INPUT_MANUAL + '" criada com validações', '');
  ui.alert('✅ Aba criada!',
    '"' + H.SHEETS.INPUT_MANUAL + '" configurada com dropdowns e validações.\n\n' +
    'Preencha as atividades e use:\nMenu → 📝 Input Manual → Importar atividades manuais.',
    ui.ButtonSet.OK);
}

// ── IMPORTAR ATIVIDADES MANUAIS → ATIVIDADES ─────────────────────────────────
function importarAtividadesManuaisUI() {
  const n = importarAtividadesManuais();
  if (n > 0) {
    SpreadsheetApp.getUi().alert('✅ ' + n + ' atividade(s) importada(s) com sucesso para ATIVIDADES.');
  } else {
    SpreadsheetApp.getUi().alert('ℹ️ Nenhuma atividade nova para importar.\n(Linhas já importadas ou incompletas são ignoradas.)');
  }
}

function importarAtividadesManuais() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetInput     = ss.getSheetByName(H.SHEETS.INPUT_MANUAL);
  const sheetAtividades = ss.getSheetByName(H.SHEETS.ATIVIDADES);

  if (!sheetInput) {
    _log('SYSTEM', 'ERRO', 'importarAtividadesManuais', 'Aba INPUT MANUAL não encontrada', '');
    return 0;
  }
  if (!sheetAtividades) {
    _log('SYSTEM', 'ERRO', 'importarAtividadesManuais', 'Aba ATIVIDADES não encontrada', '');
    return 0;
  }

  const dados = sheetInput.getDataRange().getValues();
  if (dados.length <= 2) return 0;

  let importadas = 0;

  for (let i = 2; i < dados.length; i++) {
    const row = dados[i];
    if (row[14] === true) continue;    // col 15 = Importado?
    if (!row[0] || !row[1]) continue;  // data ou atleta_id obrigatórios

    const durMin = parseFloat(row[3]) || 0;
    const distKm = parseFloat(row[4]) || 0;
    const distM  = Math.round(distKm * 1000 * 10) / 10;
    const paceS  = _parsePaceParaSegundos(String(row[5] || ''));
    const paceMs = paceS > 0
      ? Math.floor(paceS / 60) + ':' + String(paceS % 60).padStart(2, '0')
      : '';
    const fcMed  = parseInt(row[6]) || '';
    const fcMax  = parseInt(row[7]) || '';
    const durHMS = durMin > 0 ? _minsParaHMS(durMin) : '';
    const velKmh = (durMin > 0 && distKm > 0)
      ? Math.round(distKm / (durMin / 60) * 10) / 10
      : (paceS > 0 ? Math.round(3600 / paceS * 10) / 10 : '');

    const nomeAtleta = _getNomeAtleta(String(row[1]));

    const obsPartes = [
      row[8]  ? 'Zona: '      + row[8]  : '',
      row[10] ? 'Qualidade: ' + row[10] : '',
      row[11] ? 'Terreno: '   + row[11] : '',
      row[12] ? 'Período: '   + row[12] : '',
      row[13] ? String(row[13]) : '',
    ].filter(Boolean).join(' | ');

    const novaLinha = [
      'MANUAL_' + Utilities.getUuid().replace(/-/g,'').slice(0,8).toUpperCase(),
      String(row[1]),                       // ATH_ID
      nomeAtleta,                            // NOME
      row[0],                                // DATA
      String(row[2] || 'Outro'),             // TIPO
      'Manual',                              // FONTE
      '',                                    // STRAVA_ID
      'Atividade manual',                    // NOME_ATIV
      durHMS,                                // MOV_HMS
      durHMS,                                // TOTAL_HMS
      distM,                                 // DIST_M
      distKm,                                // DIST_KM
      velKmh,                                // VEL_KMH
      paceS || '',                           // PACE_S
      paceMs,                                // PACE_MS
      '',                                    // PACE_RAPIDO
      fcMed,                                 // FC_MED
      fcMax,                                 // FC_MAX
      '',                                    // ELEV
      '',                                    // CAL
      row[9] || '',                          // RPE
      obsPartes,                             // OBS
      String(row[2] || 'Outro'),             // SPORT_TYPE
      '',                                    // CADENCIA
      '',                                    // ESFORCO_REL
      new Date(),                            // DATA_IMPORT
    ];

    sheetAtividades.appendRow(novaLinha);
    sheetInput.getRange(i + 1, 15).setValue(true);
    importadas++;
  }

  _log('SYSTEM', 'INFO', 'importarAtividadesManuais', importadas + ' atividades manuais importadas', '');
  return importadas;
}

// ── HELPERS INTERNOS ──────────────────────────────────────────────────────────
function _parsePaceParaSegundos(paceStr) {
  if (!paceStr) return 0;
  const parts = paceStr.trim().split(':');
  if (parts.length !== 2) return 0;
  const mins = parseInt(parts[0]);
  const secs = parseInt(parts[1]);
  if (isNaN(mins) || isNaN(secs)) return 0;
  return mins * 60 + secs;
}

function _minsParaHMS(mins) {
  const totalSecs = Math.round(mins * 60);
  const h = Math.floor(totalSecs / 3600);
  const m = Math.floor((totalSecs % 3600) / 60);
  const s = totalSecs % 60;
  return h + ':' + String(m).padStart(2, '0') + ':' + String(s).padStart(2, '0');
}
