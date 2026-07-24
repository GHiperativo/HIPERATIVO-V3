/**
 * HIPERATIVO V3 — pipeline único de atividades Strava.
 *
 * Fonte bruta:       🏃 STRAVA_RAW (oculta)
 * Fonte operacional: 🏃 ATIVIDADES (compatível com H.ATIV)
 *
 * A migração é idempotente: preserva PSE, consolida duplicações por Strava ID,
 * incorpora registros do fluxo CONVERTIDAS e nunca apaga a aba legada.
 */

var STRAVA_RAW_HEADERS_CANONICOS = [
  'ATH_ID', 'Atleta', 'Activity ID', 'Name', 'Sport Type', 'Type',
  'Start Date Local', 'Start Date UTC', 'Timezone', 'Distance m',
  'Moving Time s', 'Elapsed Time s', 'Total Elevation Gain m',
  'Average Speed m/s', 'Max Speed m/s', 'Average Heartrate',
  'Max Heartrate', 'Average Cadence', 'Average Watts', 'Kilojoules',
  'Calories', 'Achievement Count', 'Kudos Count', 'Comment Count',
  'Athlete Count', 'PR Count', 'Manual', 'Gear ID', 'Private', 'Trainer',
  'Commute', 'Has Heartrate', 'Suffer Score', 'Workout Type', 'Map ID',
  'Summary Polyline', 'Location City', 'Location State', 'Location Country',
  'UTC Offset', 'RAW JSON', 'Fonte', 'Importado Em', 'Status'
];

// As 25 primeiras posições permanecem exatamente compatíveis com H.ATIV.
var ATIVIDADES_OPERACIONAIS_HEADERS = [
  'ID Interno', 'ATH_ID', 'Atleta', 'Data/Hora', 'Tipo', 'Fonte',
  'Strava ID', 'Nome da Atividade', 'Tempo Movimento s', 'Tempo Total s',
  'Distância m', 'Distância km', 'Velocidade m/s', 'Velocidade km/min',
  'Pace s/km', 'Pace', 'FC Média', 'FC Máxima', 'Elevação m', 'Calorias',
  'Cadência', 'Potência W', 'Rota', 'Importado em', 'PSE',
  'Tempo Movimento', 'Tempo Total', 'Velocidade km/h', 'Velocidade',
  'Distância', 'Tipo Original', 'Data', 'Hora', 'Status'
];

var _atividadesAlteradasExecucao_ = false;
var _rawIdsExecucao_ = null;
var _rawLegadoRowsExecucao_ = null;
var _rawRowsExecucao_ = null;
var _rawAlteradoExecucao_ = false;
var _estruturaAtividadesGarantidaExecucao_ = false;

/** Função pública para execução manual/API e item do menu. */
function garantirEstruturaAtividadesUnificada() {
  return _garantirEstruturaAtividadesUnificada_(true);
}

function organizarExtracoesStrava() {
  var resumo = _garantirEstruturaAtividadesUnificada_(true);
  try {
    SpreadsheetApp.getUi().alert(
      '✅ Extrações Strava organizadas',
      'RAW: ' + resumo.raw + ' registros (aba oculta)\n' +
      'ATIVIDADES: ' + resumo.atividades + ' registros convertidos\n' +
      'Duplicações removidas: ' + resumo.duplicadas + '\n' +
      'Registros incorporados do fluxo legado: ' + resumo.migradas,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (_) {}
  return resumo;
}

/**
 * Garante cabeçalhos e, quando solicitado, consolida todo o histórico.
 * @param {boolean} migrarLegado
 */
function _garantirEstruturaAtividadesUnificada_(migrarLegado) {
  if (!migrarLegado && _estruturaAtividadesGarantidaExecucao_) {
    var ssPronto = SpreadsheetApp.getActiveSpreadsheet();
    var rawPronto = ssPronto.getSheetByName('🏃 STRAVA_RAW');
    var ativPronto = ssPronto.getSheetByName(H.SHEETS.ATIVIDADES);
    return {
      raw: rawPronto ? Math.max(0, rawPronto.getLastRow() - 1) : 0,
      atividades: ativPronto ? Math.max(0, ativPronto.getLastRow() - 2) : 0,
      duplicadas: 0,
      migradas: 0
    };
  }
  var lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var shRaw = ss.getSheetByName('🏃 STRAVA_RAW') || ss.insertSheet('🏃 STRAVA_RAW');
    var shAtiv = ss.getSheetByName(H.SHEETS.ATIVIDADES) || ss.insertSheet(H.SHEETS.ATIVIDADES);
    var shLegado = ss.getSheetByName('🏃 ATIVIDADES_CONVERTIDAS');

    _garantirCabecalhoRaw_(shRaw);
    _garantirCabecalhoAtividades_(shAtiv);
    _estruturaAtividadesGarantidaExecucao_ = true;
    try { shRaw.hideSheet(); } catch (_) {}
    if (shLegado) try { shLegado.hideSheet(); } catch (_) {}

    var resumo = { raw: Math.max(0, shRaw.getLastRow() - 1), atividades: Math.max(0, shAtiv.getLastRow() - 2), duplicadas: 0, migradas: 0 };
    if (!migrarLegado) return resumo;

    var rawPorId = _lerRawPorId_(shRaw);
    var linhasAtiv = shAtiv.getLastRow() >= 3
      ? shAtiv.getRange(3, 1, shAtiv.getLastRow() - 2, ATIVIDADES_OPERACIONAIS_HEADERS.length).getValues()
      : [];
    var porId = {};
    var semStrava = [];

    linhasAtiv.forEach(function(row) {
      var normalizada = _normalizarLinhaAtividadeExistente_(row, rawPorId);
      var sid = String(normalizada[H.ATIV.STRAVA_ID - 1] || '').trim();
      if (!sid) {
        if (normalizada.some(function(v) { return v !== '' && v !== null; })) semStrava.push(normalizada);
        return;
      }
      if (porId[sid]) {
        porId[sid] = _fundirLinhasAtividade_(porId[sid], normalizada);
        resumo.duplicadas++;
      } else {
        porId[sid] = normalizada;
      }
    });

    if (shLegado && shLegado.getLastRow() >= 2) {
      var legadoHeaders = shLegado.getRange(1, 1, 1, shLegado.getLastColumn()).getValues()[0];
      var legadoIdx = _indiceHeaders_(legadoHeaders);
      var legadoRows = shLegado.getRange(2, 1, shLegado.getLastRow() - 1, shLegado.getLastColumn()).getValues();
      legadoRows.forEach(function(row) {
        var convertida = _linhaLegadaConvertidaParaOperacional_(row, legadoIdx);
        var sid = String(convertida[H.ATIV.STRAVA_ID - 1] || '').trim();
        if (!sid) return;
        convertida = _normalizarLinhaAtividadeExistente_(convertida, rawPorId);
        if (porId[sid]) {
          var antes = JSON.stringify(porId[sid]);
          porId[sid] = _fundirLinhasAtividade_(porId[sid], convertida);
          if (JSON.stringify(porId[sid]) !== antes) resumo.migradas++;
        } else {
          porId[sid] = convertida;
          resumo.migradas++;
        }
      });
    }

    var finais = Object.keys(porId).map(function(id) { return porId[id]; }).concat(semStrava);
    finais.sort(function(a, b) { return _dataMs_(b[H.ATIV.DATA - 1]) - _dataMs_(a[H.ATIV.DATA - 1]); });
    _reescreverAtividades_(shAtiv, finais);

    // O RAW passa a representar todo o histórico. Dados antigos sem JSON são
    // marcados como migrados; novas extrações guardam a resposta completa.
    finais.forEach(function(row) {
      var sid = String(row[H.ATIV.STRAVA_ID - 1] || '').trim();
      if (sid && !rawPorId[sid]) rawPorId[sid] = _linhaRawLegadaDaAtividade_(row);
    });
    var rawFinais = Object.keys(rawPorId).map(function(id) { return rawPorId[id]; });
    rawFinais.sort(function(a, b) { return _dataMs_(b[6]) - _dataMs_(a[6]); });
    _reescreverRaw_(shRaw, rawFinais);

    resumo.raw = rawFinais.length;
    resumo.atividades = finais.length;
    PropertiesService.getScriptProperties().setProperty('ATIVIDADES_PIPELINE_V2_MIGRADO_EM', new Date().toISOString());
    _rawIdsExecucao_ = null;
    _rawLegadoRowsExecucao_ = null;
    _rawAlteradoExecucao_ = false;
    SpreadsheetApp.flush();
    return resumo;
  } finally {
    lock.releaseLock();
  }
}

function _garantirCabecalhoRaw_(sh) {
  if (sh.getMaxColumns() < STRAVA_RAW_HEADERS_CANONICOS.length) {
    sh.insertColumnsAfter(sh.getMaxColumns(), STRAVA_RAW_HEADERS_CANONICOS.length - sh.getMaxColumns());
  }
  var atual = sh.getRange(1, 1, 1, STRAVA_RAW_HEADERS_CANONICOS.length).getValues()[0];
  if (String(atual[0] || '').trim() && String(atual[0] || '').trim() !== 'ATH_ID') sh.insertRowBefore(1);
  sh.getRange(1, 1, 1, STRAVA_RAW_HEADERS_CANONICOS.length)
    .setValues([STRAVA_RAW_HEADERS_CANONICOS])
    .setFontWeight('bold').setFontColor('#FFFFFF').setBackground('#455A64')
    .setHorizontalAlignment('center').setWrap(true);
  sh.setFrozenRows(1);
}

function _garantirCabecalhoAtividades_(sh) {
  var n = ATIVIDADES_OPERACIONAIS_HEADERS.length;
  if (sh.getMaxColumns() < n) sh.insertColumnsAfter(sh.getMaxColumns(), n - sh.getMaxColumns());

  var row1 = sh.getRange(1, 1, 1, Math.min(n, sh.getMaxColumns())).getValues()[0];
  var row2 = sh.getRange(2, 1, 1, Math.min(n, sh.getMaxColumns())).getValues()[0];
  if (String(row1[0] || '').trim() === ATIVIDADES_OPERACIONAIS_HEADERS[0]) {
    sh.insertRowBefore(1);
    row2 = sh.getRange(2, 1, 1, n).getValues()[0];
  }
  var temHeader = String(row2[0] || '').trim() === ATIVIDADES_OPERACIONAIS_HEADERS[0] ||
    (String(row2[1] || '').trim().toUpperCase() === 'ATH_ID' && row2.map(String).indexOf('Strava ID') >= 0);
  var row2TemConteudo = row2.some(function(v) { return v !== '' && v !== null; });
  if (!temHeader && row2TemConteudo) sh.insertRowsAfter(1, 1);

  try { sh.getRange(1, 1, 1, n).breakApart(); } catch (_) {}
  sh.getRange(1, 1, 1, n).merge()
    .setValue('🏃 ATIVIDADES — DADOS CONVERTIDOS DO STRAVA')
    .setHorizontalAlignment('center').setFontWeight('bold')
    .setFontColor('#FFFFFF').setBackground('#C33500');
  sh.getRange(2, 1, 1, n).setValues([ATIVIDADES_OPERACIONAIS_HEADERS])
    .setFontWeight('bold').setFontColor('#FFFFFF').setBackground('#003366')
    .setHorizontalAlignment('center').setWrap(true);
  sh.setFrozenRows(2);
  _aplicarLayoutAtividades_(sh);
  _formatarAbaAtividades_(sh);
}

function _aplicarLayoutAtividades_(sh) {
  var props = PropertiesService.getScriptProperties();
  var chave = 'ATIVIDADES_LAYOUT_V3_APLICADO';
  if (props.getProperty(chave) === '1') return;

  var grupos = [
    [1, 1, 120], [2, 1, 110], [3, 1, 190], [4, 1, 155], [5, 1, 115],
    [6, 1, 90], [7, 1, 125], [8, 1, 220], [9, 2, 110], [11, 2, 100],
    [13, 3, 105], [16, 1, 120], [17, 6, 90], [23, 1, 100], [24, 1, 150],
    [25, 1, 60], [26, 2, 115], [28, 1, 110], [29, 1, 115], [30, 1, 100],
    [31, 1, 115], [32, 1, 110], [33, 1, 90], [34, 1, 100]
  ];
  grupos.forEach(function(g) { sh.setColumnWidths(g[0], g[1], g[2]); });
  sh.setRowHeight(1, 30);
  sh.setRowHeight(2, 54);
  sh.setHiddenGridlines(true);
  props.setProperty(chave, '1');
}

function _formatarAbaAtividades_(sh) {
  var linhas = Math.max(1, sh.getMaxRows() - 2);
  sh.getRange(3, 4, linhas, 1).setNumberFormat('dd/MM/yyyy HH:mm:ss');
  sh.getRange(3, 9, linhas, 2).setNumberFormat('0');
  sh.getRange(3, 11, linhas, 1).setNumberFormat('0');
  sh.getRange(3, 12, linhas, 1).setNumberFormat('0.0');
  sh.getRange(3, 13, linhas, 3).setNumberFormat('0.00');
  sh.getRange(3, 28, linhas, 1).setNumberFormat('0.00');
  sh.getRange(3, 32, linhas, 1).setNumberFormat('dd/MM/yyyy');
  sh.getRange(3, 33, linhas, 1).setNumberFormat('HH:mm:ss');
}

/** Registra a resposta original antes da conversão operacional. */
function registrarAtividadesBrutasStrava_(athId, nomeAtleta, atividades, opcoes) {
  opcoes = opcoes || {};
  if (!atividades || !atividades.length) return 0;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('🏃 STRAVA_RAW') || ss.insertSheet('🏃 STRAVA_RAW');
  _garantirCabecalhoRaw_(sh);
  try { sh.hideSheet(); } catch (_) {}

  if (!_rawIdsExecucao_) {
    _rawIdsExecucao_ = new Set();
    _rawLegadoRowsExecucao_ = {};
    _rawRowsExecucao_ = {};
    if (sh.getLastRow() >= 2) {
      sh.getRange(2, 1, sh.getLastRow() - 1, STRAVA_RAW_HEADERS_CANONICOS.length).getValues().forEach(function(r, i) {
        var sidExistente = String(r[2] || '').trim();
        if (!sidExistente) return;
        _rawIdsExecucao_.add(sidExistente);
        if (!_rawRowsExecucao_[sidExistente]) _rawRowsExecucao_[sidExistente] = i + 2;
        var fonte = String(r[41] || '').trim();
        var json = String(r[40] || '').trim();
        if (fonte !== 'STRAVA_API' || !json) _rawLegadoRowsExecucao_[sidExistente] = i + 2;
      });
    }
  }

  var novas = [];
  var reparos = [];
  atividades.forEach(function(act) {
    var sid = String(act && act.id || '').trim();
    if (!sid) return;
    var linhaReal = _atividadeParaLinhaRaw_(act, athId, nomeAtleta);
    if (_rawIdsExecucao_.has(sid)) {
      var linhaExistente = opcoes.forcarAtualizacao
        ? _rawRowsExecucao_[sid]
        : _rawLegadoRowsExecucao_[sid];
      if (linhaExistente) {
        reparos.push({ row: linhaExistente, values: linhaReal });
        delete _rawLegadoRowsExecucao_[sid];
      }
      return;
    }
    novas.push(linhaReal);
    _rawIdsExecucao_.add(sid);
    _rawRowsExecucao_[sid] = sh.getLastRow() + novas.length;
  });
  if (reparos.length) _gravarReparosRaw_(sh, reparos);
  if (novas.length) {
    sh.getRange(sh.getLastRow() + 1, 1, novas.length, STRAVA_RAW_HEADERS_CANONICOS.length).setValues(novas);
  }
  if (reparos.length || novas.length) _rawAlteradoExecucao_ = true;
  return novas.length + reparos.length;
}

function _gravarReparosRaw_(sh, reparos) {
  reparos.sort(function(a, b) { return a.row - b.row; });
  var inicio = 0;
  while (inicio < reparos.length) {
    var fim = inicio + 1;
    while (fim < reparos.length && reparos[fim].row === reparos[fim - 1].row + 1) fim++;
    var bloco = reparos.slice(inicio, fim);
    sh.getRange(bloco[0].row, 1, bloco.length, STRAVA_RAW_HEADERS_CANONICOS.length)
      .setValues(bloco.map(function(item) { return item.values; }));
    inicio = fim;
  }
}

function _atividadeParaLinhaRaw_(act, athId, nomeAtleta) {
  var mapa = act.map || {};
  var json = '';
  try { json = JSON.stringify(act).substring(0, 49000); } catch (_) {}
  return [
    athId || '', nomeAtleta || '', String(act.id || ''), act.name || '',
    act.sport_type || act.type || '', act.type || '', act.start_date_local || '',
    act.start_date || '', act.timezone || '', Number(act.distance) || 0,
    Number(act.moving_time) || 0, Number(act.elapsed_time) || 0,
    Number(act.total_elevation_gain) || 0, Number(act.average_speed) || 0,
    Number(act.max_speed) || 0, act.average_heartrate || '', act.max_heartrate || '',
    act.average_cadence || '', act.average_watts || '', act.kilojoules || '',
    act.calories || '', act.achievement_count || 0, act.kudos_count || 0,
    act.comment_count || 0, act.athlete_count || 1, act.pr_count || 0,
    !!act.manual, act.gear_id || '', !!act.private, !!act.trainer, !!act.commute,
    !!act.has_heartrate, act.suffer_score || '', act.workout_type || '',
    mapa.id || '', mapa.summary_polyline || '', act.location_city || '',
    act.location_state || '', act.location_country || '', act.utc_offset || '',
    json, 'STRAVA_API', new Date(), 'ATIVO'
  ];
}

function ordenarAtividadesMaisRecentes_() {
  if (!_atividadesAlteradasExecucao_) return false;
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(H.SHEETS.ATIVIDADES);
  if (!sh || sh.getLastRow() < 4) return false;
  sh.getRange(3, 1, sh.getLastRow() - 2, ATIVIDADES_OPERACIONAIS_HEADERS.length)
    .sort({ column: H.ATIV.DATA, ascending: false });
  _formatarAbaAtividades_(sh);
  _atividadesAlteradasExecucao_ = false;
  return true;
}

function ordenarStravaRawMaisRecentes_() {
  if (!_rawAlteradoExecucao_) return false;
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('🏃 STRAVA_RAW');
  if (!sh || sh.getLastRow() < 3) return false;
  sh.getRange(2, 1, sh.getLastRow() - 1, STRAVA_RAW_HEADERS_CANONICOS.length)
    .sort({ column: 7, ascending: false });
  sh.getRange(2, 7, sh.getLastRow() - 1, 2).setNumberFormat('yyyy-MM-dd HH:mm:ss');
  sh.getRange(2, 43, sh.getLastRow() - 1, 1).setNumberFormat('dd/MM/yyyy HH:mm:ss');
  _rawAlteradoExecucao_ = false;
  try { sh.hideSheet(); } catch (_) {}
  return true;
}

function _normalizarLinhaAtividadeExistente_(entrada, rawPorId) {
  var row = entrada.slice(0, ATIVIDADES_OPERACIONAIS_HEADERS.length);
  while (row.length < ATIVIDADES_OPERACIONAIS_HEADERS.length) row.push('');
  var sid = String(row[H.ATIV.STRAVA_ID - 1] || '').trim();
  var raw = sid && rawPorId[sid] ? rawPorId[sid] : null;
  if (sid) row[H.ATIV.FONTE - 1] = 'Strava';
  if (typeof traduzirNomeAtividadeStrava_ === 'function') {
    row[7] = traduzirNomeAtividadeStrava_(row[7]);
  }
  var jaTemFormatado = !!(row[25] || row[26]);
  row[8] = _segundosOperacionais_(row[8], jaTemFormatado);
  row[9] = _segundosOperacionais_(row[9], jaTemFormatado);
  row[10] = Math.round(_numeroAtividade_(row[10]));
  row[11] = _arredondar_(_numeroAtividade_(row[11]) || row[10] / 1000, 3);
  row[12] = _numeroAtividade_(row[12]);
  row[13] = _numeroAtividade_(row[13]) || (row[12] > 0 ? _arredondar_(row[12] * 0.06, 3) : 0);
  row[14] = Math.round(_numeroAtividade_(row[14]) || (row[12] > 0 ? 1000 / row[12] : 0));
  row[15] = row[15] || _paceOperacional_(row[14], row[4]);
  row[25] = _tempoOperacional_(row[8]);
  row[26] = _tempoOperacional_(row[9]);
  row[27] = _arredondar_(_numeroAtividade_(row[27]) || row[12] * 3.6, 2);
  row[28] = row[27] > 0 ? String(row[27].toFixed(2)).replace('.', ',') + ' km/h' : '';
  row[29] = row[11] > 0 ? String(row[11].toFixed(1)).replace('.', ',') + ' km' : '';
  row[30] = row[30] || (raw ? raw[4] || raw[5] : '');
  var data = raw && raw[6] ? _dataAtividade_(raw[6]) : _dataAtividade_(row[3]);
  row[3] = data || row[3] || '';
  row[31] = data || '';
  row[32] = data ? Utilities.formatDate(data, Session.getScriptTimeZone(), 'HH:mm:ss') : '';
  row[33] = row[33] || 'Importado';
  return row;
}

function _linhaLegadaConvertidaParaOperacional_(row, idx) {
  function g(nome, fallback) {
    var i = idx[nome];
    return i === undefined ? (fallback === undefined ? '' : fallback) : row[i];
  }
  var data = _dataAtividade_(g('Data/Hora')) || _dataAtividade_(g('Data'));
  var distKm = _numeroAtividade_(g('Distância km'));
  var velKmh = _numeroAtividade_(g('Velocidade km/h'));
  var velMps = velKmh > 0 ? velKmh / 3.6 : 0;
  var movS = _numeroAtividade_(g('Tempo Movimento s'));
  var totalS = _numeroAtividade_(g('Tempo Total s'));
  var paceS = _numeroAtividade_(g('Pace s/km'));
  var id = String(g('Strava ID') || '').trim();
  return [
    g('ID Interno') || ('ATIV_' + id), g('ATH_ID'), g('Atleta'), data || '',
    g('Tipo'), 'Strava', id, g('Nome da Atividade'), movS, totalS,
    Math.round(distKm * 1000), distKm, velMps, velMps * 0.06, paceS,
    g('Pace'), _numeroAtividade_(g('FC Média')), _numeroAtividade_(g('FC Máx.')),
    _numeroAtividade_(g('Elevação m')), _numeroAtividade_(g('Calorias')),
    _numeroAtividade_(g('Cadência')), _numeroAtividade_(g('Potência W')), '',
    _dataAtividade_(g('Importado em')) || new Date(), '', g('Tempo Movimento'),
    g('Tempo Total'), velKmh, g('Velocidade'), g('Distância'), g('Tipo Original'),
    data || '', g('Hora'), g('Status') || 'Importado'
  ];
}

function _fundirLinhasAtividade_(principal, candidata) {
  var a = principal.slice();
  var b = candidata.slice();
  for (var i = 0; i < a.length; i++) {
    if ((a[i] === '' || a[i] === null || a[i] === 0) && b[i] !== '' && b[i] !== null) a[i] = b[i];
  }
  // Dados manuais e identificadores já usados por outras abas sempre vencem.
  if (principal[0]) a[0] = principal[0];
  if (principal[24] !== '' && principal[24] !== null) a[24] = principal[24];
  if (principal[23]) a[23] = principal[23];
  return a;
}

function _lerRawPorId_(sh) {
  var mapa = {};
  if (sh.getLastRow() < 2) return mapa;
  var rows = sh.getRange(2, 1, sh.getLastRow() - 1, STRAVA_RAW_HEADERS_CANONICOS.length).getValues();
  rows.forEach(function(row) {
    var id = String(row[2] || '').trim();
    if (!id) return;
    if (!mapa[id] || String(row[40] || '').length > String(mapa[id][40] || '').length) mapa[id] = row;
  });
  return mapa;
}

function _linhaRawLegadaDaAtividade_(row) {
  var data = _dataAtividade_(row[3]);
  var resumo = { migrated_from: 'ATIVIDADES', strava_id: String(row[6] || '') };
  return [
    row[1] || '', row[2] || '', String(row[6] || ''), row[7] || '',
    row[30] || '', row[30] || '', data || '', data || '', '',
    _numeroAtividade_(row[10]), _numeroAtividade_(row[8]), _numeroAtividade_(row[9]),
    _numeroAtividade_(row[18]), _numeroAtividade_(row[12]), 0,
    _numeroAtividade_(row[16]) || '', _numeroAtividade_(row[17]) || '',
    _numeroAtividade_(row[20]) || '', _numeroAtividade_(row[21]) || '', '',
    _numeroAtividade_(row[19]) || '', 0, 0, 0, 1, 0, false, '', false, false,
    false, !!row[16], '', '', '', '', '', '', '', '', JSON.stringify(resumo),
    'LEGADO_ATIVIDADES', row[23] || new Date(), 'MIGRADO'
  ];
}

function _reescreverAtividades_(sh, rows) {
  var n = ATIVIDADES_OPERACIONAIS_HEADERS.length;
  var necessario = rows.length + 2;
  if (sh.getMaxRows() < necessario) sh.insertRowsAfter(sh.getMaxRows(), necessario - sh.getMaxRows());
  if (sh.getLastRow() >= 3) sh.getRange(3, 1, sh.getLastRow() - 2, n).clearContent();
  if (rows.length) sh.getRange(3, 1, rows.length, n).setValues(rows);
  _formatarAbaAtividades_(sh);
  var filtro = sh.getFilter();
  if (filtro) filtro.remove();
  sh.getRange(2, 1, Math.max(1, rows.length + 1), n).createFilter();
}

function _reescreverRaw_(sh, rows) {
  var n = STRAVA_RAW_HEADERS_CANONICOS.length;
  var necessario = rows.length + 1;
  if (sh.getMaxRows() < necessario) sh.insertRowsAfter(sh.getMaxRows(), necessario - sh.getMaxRows());
  if (sh.getLastRow() >= 2) sh.getRange(2, 1, sh.getLastRow() - 1, n).clearContent();
  if (rows.length) sh.getRange(2, 1, rows.length, n).setValues(rows);
  sh.getRange(2, 7, Math.max(1, rows.length), 2).setNumberFormat('yyyy-MM-dd HH:mm:ss');
  sh.getRange(2, 43, Math.max(1, rows.length), 1).setNumberFormat('dd/MM/yyyy HH:mm:ss');
  try { sh.hideSheet(); } catch (_) {}
}

function _indiceHeaders_(headers) {
  var idx = {};
  headers.forEach(function(h, i) { if (h !== '' && h !== null) idx[String(h).trim()] = i; });
  return idx;
}

function _numeroAtividade_(v) {
  if (typeof v === 'number') return isFinite(v) ? v : 0;
  var s = String(v === null || v === undefined ? '' : v).trim();
  if (!s) return 0;
  var n = Number(s.replace(/\s/g, '').replace(',', '.').replace(/[^0-9.\-]/g, ''));
  return isFinite(n) ? n : 0;
}

function _segundosOperacionais_(v, jaTemFormatado) {
  var n = _numeroAtividade_(v);
  if (!jaTemFormatado && n > 0 && n < 1) return Math.round(n * 86400);
  return Math.round(n);
}

function _tempoOperacional_(segundos) {
  var s = Math.max(0, Math.round(_numeroAtividade_(segundos)));
  var h = Math.floor(s / 3600);
  var m = Math.floor((s % 3600) / 60);
  var r = s % 60;
  return String(h).padStart(2, '0') + ':' + String(m).padStart(2, '0') + ':' + String(r).padStart(2, '0');
}

function _paceOperacional_(paceS, tipo) {
  var s = Math.round(_numeroAtividade_(paceS));
  if (!s || s > 3600) return '';
  var t = String(tipo || '').toLowerCase();
  var sufixo = t.indexOf('nata') >= 0 ? ' /100m' : ' /km';
  return Math.floor(s / 60) + ':' + String(s % 60).padStart(2, '0') + sufixo;
}

function _dataAtividade_(v) {
  if (v instanceof Date && !isNaN(v.getTime())) return v;
  if (!v) return null;
  var s = String(v).trim();
  var isoLocal = s.match(/^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2}):(\d{2})/);
  if (isoLocal) return new Date(Number(isoLocal[1]), Number(isoLocal[2]) - 1, Number(isoLocal[3]), Number(isoLocal[4]), Number(isoLocal[5]), Number(isoLocal[6]));
  var br = s.match(/^(\d{2})\/(\d{2})\/(\d{4})(?:[ T](\d{2}):(\d{2})(?::(\d{2}))?)?/);
  if (br) return new Date(Number(br[3]), Number(br[2]) - 1, Number(br[1]), Number(br[4] || 0), Number(br[5] || 0), Number(br[6] || 0));
  var d = new Date(s);
  return isNaN(d.getTime()) ? null : d;
}

function _dataMs_(v) {
  var d = _dataAtividade_(v);
  return d ? d.getTime() : 0;
}

function _arredondar_(n, casas) {
  var fator = Math.pow(10, casas || 0);
  return Math.round((Number(n) || 0) * fator) / fator;
}
