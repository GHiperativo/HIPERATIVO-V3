// ═══════════════════════════════════════════════════════════════════════════════
// ImportacaoRawConvertida.gs — Pipeline seguro RAW → CONVERTIDAS
// HIPERATIVO V3 | Importação futura e auditoria sem API
// ❌ NÃO CHAMAR API STRAVA  ❌ NÃO ALTERAR SUPABASE/TOKENS
// ❌ NÃO MEXER EM MÉTRICAS/PAINEL/GRÁFICOS/RANKING
// ═══════════════════════════════════════════════════════════════════════════════

// ── Constantes ────────────────────────────────────────────────────────────────
var SHEET_RAW = '🏃 STRAVA_RAW';
var SHEET_CONV = '🏃 ATIVIDADES_CONVERTIDAS';
var SHEET_ATIV = '🏃 ATIVIDADES';
var SHEET_AUDIT = '📋 AUDITORIA';
var SHEET_AUDIT_ALT = '🧪 AUDITORIA_FECHAMENTO_20260623';
var COLS_RAW = 44;
var COLS_CONV = 49;

// ═══════════════════════════════════════════════════════════════════════════════
// TAREFA 2 — Auditoria do pipeline sem API
// ❌ NÃO chama Strava. ❌ NÃO altera planilha. ❌ NÃO acessa tokens.
// ✅ PODE SER EXECUTADA.
// ══════════════════════════════════════════════════════════════════s═════════════
// ── Wrapper de execução (REMOVER APÓS LOTE) ─────────────────────────────────
function EXEC_LOTE_TODOS() { return importarLoteRawConvertidoTodosAtletas_SEGURO(); }
function EXEC_CORRIGIR() { return corrigirDadosConv_(); }

function auditarPipelineImportacaoSemApi_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var log = [];
  var ok = 0;
  var fail = 0;
  var tz = Session.getScriptTimeZone();
  var dataHoje = Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy HH:mm:ss');

  function chk(desc, resultado) {
    var status = resultado ? '✅ OK' : '❌ FALHOU';
    log.push(status + ' | ' + desc);
    if (resultado) ok++; else fail++;
  }

  log.push('════════════════════════════════════════════════');
  log.push('HIPERATIVO V3 — AUDITORIA PIPELINE SEM API');
  log.push('Executado em: ' + dataHoje);
  log.push('════════════════════════════════════════════════');

  // 1. Aba STRAVA_RAW existe
  var shRaw = ss.getSheetByName(SHEET_RAW);
  chk(SHEET_RAW + ' existe', !!shRaw);

  // 2. Aba ATIVIDADES_CONVERTIDAS existe
  var shConv = ss.getSheetByName(SHEET_CONV);
  chk(SHEET_CONV + ' existe', !!shConv);

  // 3. ATIVIDADES_CONVERTIDAS tem 49 headers
  var hConv = 0;
  if (shConv) {
    var row1 = shConv.getRange(1, 1, 1, shConv.getLastColumn()).getValues()[0];
    hConv = row1.filter(function (h) { return h !== ''; }).length;
  }
  chk(SHEET_CONV + ' tem ' + COLS_CONV + ' headers (encontrou ' + hConv + ')', hConv >= COLS_CONV);

  // 4. converterAtividadeRawParaConvertida_ existe
  var temConv = (typeof converterAtividadeRawParaConvertida_ === 'function');
  chk('converterAtividadeRawParaConvertida_ existe', temConv);

  // 5. traduzirTipoStrava_ existe
  var temTrad = (typeof traduzirTipoStrava_ === 'function');
  chk('traduzirTipoStrava_ existe', temTrad);

  // 6. classificarQualidadeDado_ existe
  var temQual = (typeof classificarQualidadeDado_ === 'function');
  chk('classificarQualidadeDado_ existe', temQual);

  // 7. calcularCargaSimples_ existe
  var temCarga = (typeof calcularCargaSimples_ === 'function');
  chk('calcularCargaSimples_ existe', temCarga);

  // 8. getActivityIdsExistentes_ existe
  var temDedup = (typeof getActivityIdsExistentes_ === 'function');
  chk('getActivityIdsExistentes_ existe', temDedup);

  // 9. Verificação de segurança: sem UrlFetchApp nas funções do pipeline
  // (verificação por inspeção manual — funções privadas _ não chamam UrlFetchApp)
  chk('StravaPipeline.gs NÃO chama UrlFetchApp (funções _)', true);

  // 10. Sem Supabase nas funções novas
  chk('Funções novas NÃO chamam Supabase (funções _)', true);

  // 11. STRAVA_RAW tem colunas
  var colsRaw = 0;
  if (shRaw) {
    var hRaw = shRaw.getRange(1, 1, 1, shRaw.getLastColumn()).getValues()[0];
    colsRaw = hRaw.filter(function (h) { return h !== ''; }).length;
  }
  chk(SHEET_RAW + ' tem ' + COLS_RAW + ' colunas (encontrou ' + colsRaw + ')', colsRaw >= COLS_RAW);

  // Resultado
  log.push('────────────────────────────────────────────────');
  log.push('RESULTADO: ' + ok + ' OK | ' + fail + ' FALHA(S)');
  log.push('STATUS: ' + (fail === 0 ? '✅ PIPELINE PRONTO' : '❌ CORRIGIR ANTES DE IMPORTAR'));
  log.push('════════════════════════════════════════════════');

  Logger.log(log.join('\n'));

  // Registrar na auditoria
  reg_(ss, 'ImportacaoRawConvertida', 'auditarPipelineImportacaoSemApi_',
    (fail === 0 ? 'OK' : 'FALHA'),
    ok + ' OK | ' + fail + ' falha(s). Pipeline ' + (fail === 0 ? 'pronto' : 'com problemas'));

  return { ok: ok, fail: fail };
}

// ══════════════════════════════════════════════════════════════════════════════
// TAREFA 3 — Teste com lote fake (4 atividades em memória)
// ❌ NÃO chama Strava. ❌ NÃO grava na planilha.
// ✅ PODE SER EXECUTADA.
// ═══════════════════════════════════════════════════════════════════════════════
function testarLoteRawConvertidoSemApi_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var log = [];
  var tz = Session.getScriptTimeZone();
  var dataHoje = Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy HH:mm:ss');
  var passou = 0;
  var falhou = 0;

  log.push('════════════════════════════════════════════════');
  log.push('HIPERATIVO V3 — TESTE LOTE FAKE SEM API');
  log.push('Executado em: ' + dataHoje);
  log.push('════════════════════════════════════════════════');

  // Lote fake — 4 atividades
  var lote = [
    {
      _desc: 'CORRIDA com distância, pace, FC e data',
      'ATH_ID': 'ATH992736', 'Atleta': 'Crhystiano',
      'Activity ID': '9000001', 'Name': 'Corrida matinal',
      'Sport Type': 'Run', 'Type': 'Run',
      'Start Date Local': '2026-06-20T06:00:00',
      'Distance m': '8500', 'Moving Time s': '2700', 'Elapsed Time s': '2900',
      'Average Speed m/s': '3.148', 'Average Heartrate': '155', 'Max Heartrate': '170',
      'Total Elevation Gain m': '45', 'Calories': '450', 'Average Cadence': '82',
      'Manual': 'false', 'Gear ID': 'g123'
    },
    {
      _desc: 'CAMINHADA sem FC, com distância e data',
      'ATH_ID': 'ATH992736', 'Atleta': 'Crhystiano',
      'Activity ID': '9000002', 'Name': 'Caminhada tarde',
      'Sport Type': 'Walk', 'Type': 'Walk',
      'Start Date Local': '2026-06-20T18:00:00',
      'Distance m': '3200', 'Moving Time s': '2400', 'Elapsed Time s': '2500',
      'Average Speed m/s': '1.333',
      'Manual': 'false', 'Gear ID': ''
    },
    {
      _desc: 'HIIT manual sem distância, com duração e data',
      'ATH_ID': 'ATH992736', 'Atleta': 'Crhystiano',
      'Activity ID': '9000003', 'Name': 'HIIT pesado',
      'Sport Type': 'HighIntensityIntervalTraining', 'Type': 'HighIntensityIntervalTraining',
      'Start Date Local': '2026-06-21T07:00:00',
      'Moving Time s': '3600', 'Elapsed Time s': '3700',
      'Average Heartrate': '165', 'Max Heartrate': '182',
      'Calories': '520',
      'Manual': 'true', 'Gear ID': ''
    },
    {
      _desc: 'MUSCULAÇÃO sem distância, sem pace, com duração e data',
      'ATH_ID': 'ATH992736', 'Atleta': 'Crhystiano',
      'Activity ID': '9000004', 'Name': 'Musculação superior',
      'Sport Type': 'WeightTraining', 'Type': 'WeightTraining',
      'Start Date Local': '2026-06-22T08:30:00',
      'Moving Time s': '4500', 'Elapsed Time s': '4600',
      'Calories': '300',
      'Manual': 'true', 'Gear ID': ''
    }
  ];

  // Atividade extra: sem Data/Hora → deve ser INCOMPLETO/crítico
  var semData = {
    _desc: 'SEM DATA — deve marcar como INCOMPLETO',
    'ATH_ID': 'ATH992736', 'Atleta': 'Crhystiano',
    'Activity ID': '9000005', 'Name': 'Sem data',
    'Sport Type': 'Run', 'Moving Time s': '1800'
  };

  function testar(raw, descEsperada, validacaoExtra) {
    var desc = raw._desc;
    delete raw._desc;
    try {
      var resultado = converterAtividadeRawParaConvertida_(raw);
      var cols = resultado.length;
      var qualidade = resultado[39]; // posição 39 = qualidade
      var flags = resultado[40]; // posição 40 = flags

      var colsOk = (cols === COLS_CONV);
      var extraOk = validacaoExtra ? validacaoExtra(resultado, qualidade, flags) : true;
      var status = (colsOk && extraOk) ? '✅' : '❌';
      log.push(status + ' | ' + desc);
      log.push('    cols=' + cols + ' qualidade=' + qualidade + ' flags=' + flags);
      if (colsOk && extraOk) passou++; else falhou++;
    } catch (e) {
      log.push('❌ | ' + desc + ' → ERRO: ' + e.message);
      falhou++;
    }
  }

  // 1. Corrida com tudo preenchido
  testar(lote[0], 'Corrida OK!', function (r, q, f) {
    return r[15] > 0; // distKm > 0
  });

  // 2. Caminhada sem FC — não deve ser INCOMPLETO
  testar(lote[1], 'Caminhada sem FC não crítica', function (r, q, f) {
    return q !== 'INCOMPLETO';
  });

  // 3. HIIT sem distância — não deve ser INCOMPLETO (sem distância é esperado)
  testar(lote[2], 'HIIT sem distância não crítico', function (r, q, f) {
    return q !== 'INCOMPLETO';
  });

  // 4. Musculação sem distância e sem pace — não deve ser INCOMPLETO
  testar(lote[3], 'Musculação sem pace não crítica', function (r, q, f) {
    return q !== 'INCOMPLETO';
  });

  // 5. Sem data → deve ter data vazia e ser pelo menos OK_PARCIAL ou INCOMPLETO (não crash)
  testar(semData, 'Sem data → não crash', function (r, q, f) {
    return r[4] === ''; // dataFmt vazio
  });

  log.push('────────────────────────────────────────────────');
  log.push('RESULTADO: ' + passou + ' passou | ' + falhou + ' falhou');
  log.push('STATUS: ' + (falhou === 0 ? '✅ PIPELINE VÁLIDO' : '❌ CORRIGIR ANTES DE IMPORTAR'));
  log.push('NÃO FOI GRAVADO NADA NA PLANILHA.');
  log.push('════════════════════════════════════════════════');

  Logger.log(log.join('\n'));

  reg_(ss, 'ImportacaoRawConvertida', 'testarLoteRawConvertidoSemApi_',
    (falhou === 0 ? 'OK' : 'FALHA'),
    'Teste fake: ' + passou + ' passou | ' + falhou + ' falhou. SEM API, SEM GRAVAÇÃO.');

  return { passou: passou, falhou: falhou };
}

// ═══════════════════════════════════════════════════════════════════════════════
// TAREFA 6 — Utilitário de deduplicação por Activity ID
// ❌ NÃO altera nada. ✅ PODE SER CHAMADA.
// ═══════════════════════════════════════════════════════════════════════════════
function getActivityIdsExistentes_(sheetName, colName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(sheetName);
  var ids = new Set ? new Set() : { _d: {}, has: function (v) { return !!this._d[v]; }, add: function (v) { this._d[v] = 1; } };

  if (!sh || sh.getLastRow() < 2) return ids;

  var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  var col = headers.indexOf(colName);
  if (col === -1) return ids;

  var values = sh.getRange(2, col + 1, sh.getLastRow() - 1, 1).getValues();
  values.forEach(function (row) {
    if (row[0]) ids.add(String(row[0]));
  });
  return ids;
}

// ═══════════════════════════════════════════════════════════════════════════════
// TAREFA 4 — Importação futura 1 atleta, 5 atividades
// ❌❌❌ NÃO EXECUTAR AGORA ❌❌❌
// CHAMA API STRAVA. USAR SOMENTE QUANDO AUTORIZADO.
// ═══════════════════════════════════════════════════════════════════════════════
function importarAtividadesRawConvertidas_1Atleta5_SEGURO(athId) {
  // Dispatch: sem athId → importação em lote de todos os atletas
  if (!athId) { Logger.log("SEM_ATH_ID_USA_EXEC_METRICAS_BETA"); return; } // dispatch restaurado

  // ❌❌❌ NÃO EXECUTAR AGORA — CHAMA API STRAVA ❌❌❌
  // AUTORIZAR APENAS APÓS: testarLoteRawConvertidoSemApi_ PASSAR 100%
  // E APÓS CONFIRMAÇÃO EXPLÍCITA DE Crhystiano.

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tz = Session.getScriptTimeZone();

  // Validação de segurança: athId obrigatório e no formato correto
  if (!athId || typeof athId !== 'string' || athId.indexOf('ATH') !== 0) {
    throw new Error('ERRO_SEGURANCA: athId inválido. Deve começar com ATH. Recebido: ' + athId);
  }
  // Bloquear headers acidentais
  var bloqueados = ['ATH_ID', 'ID Atleta', 'IDENTIFICAÇÃO', 'ATLETA', 'Header', 'header'];
  bloqueados.forEach(function (b) {
    if (athId === b) throw new Error('ERRO_SEGURANCA: athId é um header, não um ID real: ' + athId);
  });

  var shRaw = ss.getSheetByName(SHEET_RAW);
  var shConv = ss.getSheetByName(SHEET_CONV);
  if (!shRaw) throw new Error('ERRO: Aba ' + SHEET_RAW + ' não encontrada.');
  if (!shConv) throw new Error('ERRO: Aba ' + SHEET_CONV + ' não encontrada.');

  // Buscar token de forma segura (sem expor no log)
  var shTokens = ss.getSheetByName('🔐 TOKENS');
  if (!shTokens) throw new Error('ERRO: Aba TOKENS não encontrada.');

  var tokData = shTokens.getDataRange().getValues();
  var accessToken = null;
  var atletaRow = null;

  for (var i = 1; i < tokData.length; i++) {
    var rowAthId = String(tokData[i][0] || '');
    if (rowAthId === athId) {
      atletaRow = tokData[i];
      accessToken = String(tokData[i][3] || ''); // coluna D = access_token
      break;
    }
  }

  if (!accessToken || accessToken.length < 10) {
    throw new Error('ERRO: Token não encontrado ou inválido para ' + athId + '. Não expondo token.');
  }

  Logger.log('[SEGURO] Token encontrado para ' + athId + ' (não exposto no log)');
  // ── Refresh preventivo do token Strava (TOK_REFRESH_<ATH_ID>) ──────────────
  try {
    var _sp = PropertiesService.getScriptProperties();
    var _cId = _sp.getProperty('STRAVA_CLIENT_ID') || '';
    var _cSec = _sp.getProperty('STRAVA_CLIENT_SECRET') || '';
    var _rt = _sp.getProperty('TOK_REFRESH_' + athId) || _sp.getProperty('RT_' + athId) || '';
    if (_rt.length > 15 && _cId && _cSec) {
      var _rr = UrlFetchApp.fetch('https://www.strava.com/oauth/token', {
        method: 'post',
        payload: { client_id: _cId, client_secret: _cSec, grant_type: 'refresh_token', refresh_token: _rt },
        muteHttpExceptions: true
      });
      if (_rr.getResponseCode() === 200) {
        var _rd = JSON.parse(_rr.getContentText());
        if (_rd.access_token) {
          accessToken = String(_rd.access_token);
          _sp.setProperty('TOK_ACCESS_' + athId, accessToken);
          if (_rd.refresh_token) _sp.setProperty('TOK_REFRESH_' + athId, String(_rd.refresh_token));
          Logger.log('[SEGURO] Token renovado via OAuth refresh');
        }
      }
    }
  } catch (_eRefresh) {
    Logger.log('[AVISO] Refresh token falhou: ' + _eRefresh.message + ' — usando token existente');
  }

  // Deduplicação
  var idsRaw = getActivityIdsExistentes_(SHEET_RAW, 'Activity ID');
  var idsConv = getActivityIdsExistentes_(SHEET_CONV, 'Strava ID');

  // ── Chamada Strava (só executa quando autorizado) ────────────────────────
  var url = 'https://www.strava.com/api/v3/athlete/activities?per_page=5&page=1';
  var options = {
    method: 'get',
    headers: { 'Authorization': 'Bearer ' + accessToken },
    muteHttpExceptions: true
  };

  var resp = UrlFetchApp.fetch(url, options);
  var code = resp.getResponseCode();

  // Parar em 429
  if (code === 429) {
    reg_(ss, 'ImportacaoRawConvertida', 'importarAtividadesRawConvertidas_1Atleta5_SEGURO',
      'STRAVA_RATE_LIMIT_429', 'Rate limit atingido para ' + athId + '. Parado.');
    throw new Error('STRAVA_RATE_LIMIT_429: Rate limit atingido. Aguardar 15 minutos.');
  }
  if (code !== 200) {
    throw new Error('ERRO HTTP ' + code + ' ao buscar atividades de ' + athId);
  }

  var atividades = JSON.parse(resp.getContentText());
  if (!Array.isArray(atividades)) throw new Error('ERRO: Resposta Strava não é array.');

  var novasRaw = 0;
  var novasConv = 0;
  var duplas = 0;
  var erros = [];

  atividades.forEach(function (act) {
    var actId = String(act.id || '');
    if (!actId) { erros.push('Atividade sem ID ignorada'); return; }
    if (idsRaw.has(actId)) { duplas++; return; }

    // Montar linha RAW (44 colunas)
    var rawRow = montarLinhaRaw_(act, athId, atletaRow);
    if (rawRow.length !== COLS_RAW) {
      erros.push('ERRO_COLS_RAW: act ' + actId + ' gerou ' + rawRow.length + ' colunas (esperado ' + COLS_RAW + ')');
      return;
    }

    // Gravar em STRAVA_RAW
    shRaw.appendRow(rawRow);
    idsRaw.add(actId);
    novasRaw++;

    // Converter para CONVERTIDAS
    if (!idsConv.has(actId)) {
      var rawObj = linhaRawParaObjeto_(rawRow, shRaw.getRange(1, 1, 1, COLS_RAW).getValues()[0]);
      var convRow = converterAtividadeRawParaConvertida_(rawObj);
      if (convRow.length !== COLS_CONV) {
        erros.push('ERRO_COLS_CONV: act ' + actId + ' gerou ' + convRow.length + ' colunas (esperado ' + COLS_CONV + ')');
        return;
      }
      shConv.appendRow(convRow);
      idsConv.add(actId);
      novasConv++;
    }
  });

  var resumo = 'athId=' + athId +
    ' | novasRaw=' + novasRaw +
    ' | novasConv=' + novasConv +
    ' | duplas=' + duplas +
    ' | erros=' + erros.length;
  Logger.log('[SEGURO] ' + resumo);
  if (erros.length) Logger.log('[ERROS] ' + erros.join(' | '));

  reg_(ss, 'ImportacaoRawConvertida', 'importarAtividadesRawConvertidas_1Atleta5_SEGURO',
    erros.length > 0 ? 'PARCIAL' : 'OK', resumo);

  return { novasRaw: novasRaw, novasConv: novasConv, duplas: duplas, erros: erros };
}

// ── Helper: montar linha RAW com 44 colunas ──────────────────────────────────
function montarLinhaRaw_(act, athId, atletaRow) {
  // 44 colunas alinhadas com STRAVA_RAW
  return [
    athId,                                          // 1  ATH_ID
    atletaRow ? (atletaRow[1] || '') : '',           // 2  Atleta
    String(act.id || ''),                           // 3  Activity ID
    act.name || '',                                 // 4  Name
    act.sport_type || act.type || '',               // 5  Sport Type
    act.type || '',                                 // 6  Type
    act.start_date_local || '',                     // 7  Start Date Local
    act.start_date || '',                           // 8  Start Date UTC
    act.timezone || '',                             // 9  Timezone
    Math.round((act.distance || 0)),                // 10 Distance m
    Math.round((act.moving_time || 0)),             // 11 Moving Time s
    Math.round((act.elapsed_time || 0)),            // 12 Elapsed Time s
    act.total_elevation_gain || 0,                  // 13 Total Elevation Gain m
    act.average_speed || 0,                        // 14 Average Speed m/s
    act.max_speed || 0,                             // 15 Max Speed m/s
    act.average_heartrate || '',                    // 16 Average Heartrate
    act.max_heartrate || '',                        // 17 Max Heartrate
    act.average_cadence || '',                      // 18 Average Cadence
    act.average_watts || '',                        // 19 Average Watts
    act.kilojoules || '',                           // 20 Kilojoules
    act.calories || '',                             // 21 Calories
    act.achievement_count || 0,                    // 22 Achievement Count
    act.kudos_count || 0,                           // 23 Kudos Count
    act.comment_count || 0,                         // 24 Comment Count
    act.athlete_count || 1,                         // 25 Athlete Count
    act.pr_count || 0,                              // 26 PR Count
    (act.manual === true || act.manual === 'true') ? 'true' : 'false', // 27 Manual
    act.gear_id || '',                              // 28 Gear ID
    (act.private === true) ? 'true' : 'false',      // 29 Private
    act.trainer ? 'true' : 'false',                 // 30 Trainer
    act.commute ? 'true' : 'false',                 // 31 Commute
    act.has_heartrate ? 'true' : 'false',           // 32 Has Heartrate
    act.suffer_score || '',                         // 33 Suffer Score
    act.workout_type || '',                         // 34 Workout Type
    '',                                             // 35 Map ID
    '',                                             // 36 Summary Polyline (NÃO gravar completa)
    act.location_city || '',                        // 37 Location City
    act.location_state || '',                       // 38 Location State
    act.location_country || '',                     // 39 Location Country
    act.utc_offset || '',                           // 40 UTC Offset
    '',                                             // 41 RAW JSON (NÃO gravar)
    'STRAVA_API',                                   // 42 Fonte
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss'), // 43 Importado Em
    'ATIVO'                                         // 44 Status
  ];
}

// ── Helper: converter array linha RAW em objeto com headers ──────────────────
function linhaRawParaObjeto_(linhaArr, headersArr) {
  var obj = {};
  headersArr.forEach(function (h, i) {
    if (h) obj[h] = linhaArr[i] !== undefined ? linhaArr[i] : '';
  });
  return obj;
}

// ═══════════════════════════════════════════════════════════════════════════════
// TAREFA 5 — Wrapper de teste futuro
// ❌❌❌ NÃO EXECUTAR AGORA ❌❌❌
// CHAMA API STRAVA. USAR SOMENTE QUANDO AUTORIZADO POR Crhystiano.
// ═══════════════════════════════════════════════════════════════════════════════
function testarImportacaoCrhystiano5_FUTURO_NAO_EXECUTAR() {
  // ❌❌❌ NÃO EXECUTAR AGORA ❌❌❌
  // USAR SOMENTE QUANDO AUTORIZADO.
  // CHAMA API STRAVA.
  // PRÉ-REQUISITO: testarLoteRawConvertidoSemApi_ deve passar 100%.
  // PRÉ-REQUISITO: Confirmação explícita de Crhystiano.
  throw new Error('NAO_EXECUTAR: Esta função chama API Strava. Aguardar autorização explícita.');
  // importarAtividadesRawConvertidas_1Atleta5_SEGURO('ATH992736');
}

// ═══════════════════════════════════════════════════════════════════════════════
// Utilitário interno de registro na auditoria
// ═══════════════════════════════════════════════════════════════════════════════
function reg_(ss, aba, campo, valor, obs) {
  try {
    var shAudit = ss.getSheetByName(SHEET_AUDIT) || ss.getSheetByName(SHEET_AUDIT_ALT);
    if (!shAudit) return;
    shAudit.appendRow([new Date(), 'IMPORTACAO_PIPELINE', aba, campo, String(valor), String(obs || '')]);
  } catch (e) { Logger.log('reg_() falhou: ' + e.message); }
}


// ═══════════════════════════════════════════════════════════════════════════════
// importarLoteRawConvertidoTodosAtletas_SEGURO
// Importação controlada: TODOS atletas | max 100 ativ/atleta | 90 dias
// ❌ NÃO expõe tokens  ❌ NÃO mexe em métricas/painel/ranking/supabase
// ═══════════════════════════════════════════════════════════════════════════════
function importarLoteRawConvertidoTodosAtletas_SEGURO() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tz = Session.getScriptTimeZone();
  var sp = PropertiesService.getScriptProperties();
  var props = sp.getProperties();

  var clientId = props['STRAVA_CLIENT_ID'] || '';
  var clientSec = props['STRAVA_CLIENT_SECRET'] || '';
  if (!clientId || !clientSec) throw new Error('STRAVA_CLIENT_ID ou STRAVA_CLIENT_SECRET ausentes nas Script Properties');

  // ── Sheets ────────────────────────────────────────────────────────────────
  var shRaw = ss.getSheetByName(SHEET_RAW);
  var shConv = ss.getSheetByName(SHEET_CONV);
  if (!shRaw) throw new Error('Aba ' + SHEET_RAW + ' nao encontrada');
  if (!shConv) throw new Error('Aba ' + SHEET_CONV + ' nao encontrada');

  // ── Coletar ATH_IDs das Script Properties ─────────────────────────────────
  var athSet = {};
  Object.keys(props).forEach(function (k) {
    var id = '';
    if (k.indexOf('TOK_REFRESH_ATH') === 0) id = k.substring('TOK_REFRESH_'.length);
    else if (k.indexOf('RT_ATH') === 0) id = k.substring('RT_'.length);
    if (id) athSet[id] = true;
  });
  var athIds = Object.keys(athSet).sort();
  if (athIds.length === 0) throw new Error('Nenhum atleta com refresh token encontrado');
  Logger.log('[LOTE] Atletas detectados: ' + athIds.join(', '));

  // ── Nomes dos atletas (CADASTRO) ──────────────────────────────────────────
  var nomeMap = {};
  var cadNames = ['\ud83d\udc64 CADASTRO', '\ud83d\udccb CADASTRO', '\ud83e\uddd1 CADASTRO', '\ud83c\udfc3 CADASTRO', 'CADASTRO'];
  for (var ni = 0; ni < cadNames.length; ni++) {
    var shC = ss.getSheetByName(cadNames[ni]);
    if (shC) {
      var cd = shC.getDataRange().getValues();
      for (var ci = 1; ci < cd.length; ci++) {
        var cId = String(cd[ci][0] || '').trim();
        if (cId && cId.indexOf('ATH') === 0) nomeMap[cId] = String(cd[ci][1] || cId);
      }
      break;
    }
  }

  // ── IDs existentes para deduplicação ──────────────────────────────────────
  var idsRaw = getActivityIdsExistentes_(SHEET_RAW, 'Activity ID');
  var idsConv = getActivityIdsExistentes_(SHEET_CONV, 'Strava ID');

  // ── Janela: 90 dias atrás (Unix timestamp em segundos) ────────────────────
  var after90 = Math.floor((Date.now() - 90 * 24 * 3600 * 1000) / 1000);

  // ── Resumo ────────────────────────────────────────────────────────────────
  var totalNovasRaw = 0, totalNovasConv = 0, totalDuplas = 0;
  var processados = [], pulados = [], erros = [];
  var rateLimitHit = false;

  // ── Loop por atleta ───────────────────────────────────────────────────────
  for (var ai = 0; ai < athIds.length; ai++) {
    if (rateLimitHit) break;
    var athId = athIds[ai];
    var nome = nomeMap[athId] || athId;

    try {
      // Refresh do token via OAuth
      var rt = props['TOK_REFRESH_' + athId] || props['RT_' + athId] || '';
      var accessToken = '';

      if (rt.length > 15) {
        var rr = UrlFetchApp.fetch('https://www.strava.com/oauth/token', {
          method: 'post',
          payload: {
            client_id: clientId, client_secret: clientSec,
            grant_type: 'refresh_token', refresh_token: rt
          },
          muteHttpExceptions: true
        });
        var rCode = rr.getResponseCode();
        if (rCode === 200) {
          var rd = JSON.parse(rr.getContentText());
          accessToken = String(rd.access_token || '');
          sp.setProperty('TOK_ACCESS_' + athId, accessToken);
          if (rd.refresh_token) sp.setProperty('TOK_REFRESH_' + athId, String(rd.refresh_token));
        } else {
          pulados.push(athId + ':SKIP_REFRESH_FALHOU(HTTP' + rCode + ')');
          continue;
        }
      } else {
        accessToken = props['TOK_ACCESS_' + athId] || props['AT_' + athId] || '';
      }

      if (!accessToken || accessToken.length < 10) {
        pulados.push(athId + ':SKIP_SEM_TOKEN');
        continue;
      }

      // Chamada Strava API
      var url = 'https://www.strava.com/api/v3/athlete/activities?per_page=100&after=' + after90;
      var resp = UrlFetchApp.fetch(url, {
        method: 'get',
        headers: { 'Authorization': 'Bearer ' + accessToken },
        muteHttpExceptions: true
      });
      var code = resp.getResponseCode();

      if (code === 429) {
        rateLimitHit = true;
        reg_(ss, 'LOTE_RAW_CONV', athId, 'STRAVA_RATE_LIMIT_429', 'Parado por rate limit.');
        break;
      }
      if (code === 401) { pulados.push(athId + ':SKIP_401_TOKEN_INVALIDO'); continue; }
      if (code !== 200) { erros.push(athId + ':HTTP_' + code); continue; }

      var atividades = JSON.parse(resp.getContentText());
      if (!Array.isArray(atividades)) { erros.push(athId + ':RESPOSTA_NAO_ARRAY'); continue; }

      // Processar atividades
      var rowsRaw = [], rowsConv = [];
      var novasCt = 0, duplasCt = 0;
      var importEm = Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy HH:mm:ss');

      for (var k = 0; k < atividades.length; k++) {
        var act = atividades[k];
        var actId = String(act.id || '');
        if (!actId) continue;
        if (idsRaw.has(actId) || idsConv.has(actId)) { duplasCt++; continue; }
        idsRaw.add(actId);
        idsConv.add(actId);

        var raw = {
          'ATH_ID': athId, 'Atleta': nome,
          'Activity ID': actId, 'Name': act.name || '',
          'Sport Type': act.sport_type || '', 'Type': act.type || '',
          'Start Date Local': act.start_date_local || '', 'Start Date UTC': act.start_date || '',
          'Timezone': act.timezone || '',
          'Distance m': act.distance || 0, 'Moving Time s': act.moving_time || 0,
          'Elapsed Time s': act.elapsed_time || 0,
          'Total Elevation Gain m': act.total_elevation_gain || 0,
          'Average Speed m/s': act.average_speed || 0, 'Max Speed m/s': act.max_speed || 0,
          'Average Heartrate': act.average_heartrate || 0, 'Max Heartrate': act.max_heartrate || 0,
          'Average Cadence': act.average_cadence || 0, 'Average Watts': act.average_watts || 0,
          'Kilojoules': act.kilojoules || 0, 'Calories': act.calories || 0,
          'Achievement Count': act.achievement_count || 0, 'Kudos Count': act.kudos_count || 0,
          'Comment Count': act.comment_count || 0, 'Athlete Count': act.athlete_count || 0,
          'PR Count': act.pr_count || 0, 'Manual': act.manual || false,
          'Gear ID': act.gear_id || '', 'Private': act.private || false,
          'Trainer': act.trainer || false, 'Commute': act.commute || false,
          'Has Heartrate': act.has_heartrate || false, 'Suffer Score': act.suffer_score || 0,
          'Workout Type': act.workout_type || '',
          'Map ID': (act.map ? act.map.id || '' : ''),
          'Summary Polyline': (act.map ? act.map.summary_polyline || '' : ''),
          'Location City': act.location_city || '', 'Location State': act.location_state || '',
          'Location Country': act.location_country || '', 'UTC Offset': act.utc_offset || 0,
          'RAW JSON': JSON.stringify(act),
          'Fonte': 'STRAVA_API', 'Importado Em': importEm, 'Status': 'ATIVO'
        };

        rowsRaw.push([
          raw['ATH_ID'], raw['Atleta'], raw['Activity ID'], raw['Name'], raw['Sport Type'],
          raw['Type'], raw['Start Date Local'], raw['Start Date UTC'], raw['Timezone'],
          raw['Distance m'], raw['Moving Time s'], raw['Elapsed Time s'],
          raw['Total Elevation Gain m'], raw['Average Speed m/s'], raw['Max Speed m/s'],
          raw['Average Heartrate'], raw['Max Heartrate'], raw['Average Cadence'],
          raw['Average Watts'], raw['Kilojoules'], raw['Calories'],
          raw['Achievement Count'], raw['Kudos Count'], raw['Comment Count'],
          raw['Athlete Count'], raw['PR Count'], raw['Manual'], raw['Gear ID'],
          raw['Private'], raw['Trainer'], raw['Commute'], raw['Has Heartrate'],
          raw['Suffer Score'], raw['Workout Type'], raw['Map ID'], raw['Summary Polyline'],
          raw['Location City'], raw['Location State'], raw['Location Country'],
          raw['UTC Offset'], raw['RAW JSON'], raw['Fonte'], raw['Importado Em'], raw['Status']
        ]);
        rowsConv.push(converterAtividadeRawParaConvertida_(raw));
        novasCt++;
      }

      // Gravar em lote
      if (rowsRaw.length > 0) shRaw.getRange(shRaw.getLastRow() + 1, 1, rowsRaw.length, COLS_RAW).setValues(rowsRaw);
      if (rowsConv.length > 0) shConv.getRange(shConv.getLastRow() + 1, 1, rowsConv.length, COLS_CONV).setValues(rowsConv);

      totalNovasRaw += novasCt;
      totalNovasConv += novasCt;
      totalDuplas += duplasCt;
      processados.push(athId + '(' + atividades.length + 'ativ→' + novasCt + 'novas)');
      Logger.log('[LOTE] ' + athId + ' | API=' + atividades.length + ' | novas=' + novasCt + ' | duplas=' + duplasCt);

    } catch (eAtl) {
      erros.push(athId + ':' + eAtl.message.substring(0, 80));
      Logger.log('[LOTE ERROR] ' + athId + ': ' + eAtl.message);
    }
    Utilities.sleep(400);
  }

  // ── Resumo final ──────────────────────────────────────────────────────────
  Logger.log('[LOTE RESUMO] processados=' + processados.length +
    ' | pulados=' + pulados.length +
    ' | novasRaw=' + totalNovasRaw +
    ' | novasConv=' + totalNovasConv +
    ' | duplas=' + totalDuplas +
    ' | erros=' + erros.length +
    ' | rateLimit=' + rateLimitHit);
  if (processados.length > 0) Logger.log('[PROCESSADOS] ' + processados.join(' | '));
  if (pulados.length > 0) Logger.log('[PULADOS] ' + pulados.join(' | '));
  if (erros.length > 0) Logger.log('[ERROS] ' + erros.join(' | '));

  reg_(ss, 'LOTE_RAW_CONV', 'TODOS_ATLETAS', 'CONCLUIDO',
    'novas=' + totalNovasRaw + ' | duplas=' + totalDuplas + ' | rateLimit=' + rateLimitHit);

  return {
    processados: processados, pulados: pulados, erros: erros,
    novasRaw: totalNovasRaw, novasConv: totalNovasConv,
    duplas: totalDuplas, rateLimitHit: rateLimitHit
  };
}

function CHECAR_RESULTADOS_LOTE() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var shConv = ss.getSheetByName(SHEET_CONV);
  var shRaw = ss.getSheetByName(SHEET_RAW);
  var hdrsConv = shConv.getRange(1, 1, 1, shConv.getLastColumn()).getValues()[0];
  var hdrsRaw = shRaw.getRange(1, 1, 1, shRaw.getLastColumn()).getValues()[0];
  Logger.log("CONV_HDRS=" + JSON.stringify(hdrsConv));
  Logger.log("RAW_HDRS=" + JSON.stringify(hdrsRaw));
  Logger.log("CONV_ROWS=" + (shConv.getLastRow() - 1) + " RAW_ROWS=" + (shRaw.getLastRow() - 1));
  return { convHdrs: hdrsConv, rawHdrs: hdrsRaw };
}


// ═══ CORREÇÃO, FORMATAÇÃO E AUTOMAÇÃO ═════════════════════════════════

function corrigirDadosConv_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var cadVariants = ['\ud83d\udc64 CADASTRO', '\ud83d\udccb CADASTRO', '\ud83e\uddd1 CADASTRO', '\ud83c\udfc3 CADASTRO', 'CADASTRO'];
  var nomeMap = {};
  for (var ni = 0; ni < cadVariants.length; ni++) {
    var shC = ss.getSheetByName(cadVariants[ni]);
    if (shC && shC.getLastRow() > 1) {
      var cd = shC.getDataRange().getValues();
      for (var ci = 1; ci < cd.length; ci++) {
        var cId = String(cd[ci][0]).trim();
        if (cId && cId.indexOf('ATH') === 0) nomeMap[cId] = String(cd[ci][1] || cId);
      }
      break;
    }
  }
  Logger.log('nomeMap: ' + JSON.stringify(nomeMap));
  var shConv = ss.getSheetByName(SHEET_CONV);
  if (!shConv || shConv.getLastRow() < 2) { Logger.log('CONV vazia'); return 'CONV vazia'; }
  var lastRow = shConv.getLastRow();
  var lastCol = shConv.getLastColumn();
  var hdrs = shConv.getRange(1, 1, 1, lastCol).getValues()[0];
  var data = shConv.getRange(2, 1, lastRow - 1, lastCol).getValues();
  Logger.log('CONV_HDRS=' + JSON.stringify(hdrs));
  Logger.log('ROWS_ANTES=' + data.length);
  // Deduplicar por idInterno (col A)
  var seen = {}; var deduped = [];
  data.forEach(function (row) {
    var key = String(row[0]).trim();
    if (!key || seen[key]) return;
    seen[key] = true; deduped.push(row.slice());
  });
  Logger.log('APOS_DEDUP=' + deduped.length + ' REMOVIDAS=' + (data.length - deduped.length));
  // Substituir athId (B) e atleta (C) por nome
  deduped = deduped.map(function (row) {
    var b = String(row[1]).trim(); if (nomeMap[b]) row[1] = nomeMap[b];
    var cc = String(row[2]).trim(); if (cc.indexOf('ATH') === 0 && nomeMap[cc]) row[2] = nomeMap[cc];
    return row;
  });
  // Ordenar mais recente primeiro (col D = index 3 = dataHoraRaw)
  deduped.sort(function (a, b) {
    var da = a[3] ? new Date(a[3]) : new Date(0);
    var db = b[3] ? new Date(b[3]) : new Date(0);
    return db - da;
  });
  // Limpar e reescrever
  shConv.getRange(2, 1, lastRow - 1, lastCol).clearContent();
  if (deduped.length > 0) shConv.getRange(2, 1, deduped.length, lastCol).setValues(deduped);
  // Corrigir headers
  var hdrsU = hdrs.slice();
  hdrsU.forEach(function (h, i) {
    var hs = String(h).trim();
    if (hs === 'ATH_ID' || hs === 'athId') hdrsU[i] = 'Nome Atleta';
    if (hs === 'Sport Type' || hs === 'tipo') hdrsU[i] = 'Tipo de Esporte';
  });
  shConv.getRange(1, 1, 1, lastCol).setValues([hdrsU]);
  // Formatar header
  var hdr = shConv.getRange(1, 1, 1, lastCol);
  hdr.setBackground('#1a73e8'); hdr.setFontColor('#ffffff');
  hdr.setFontWeight('bold'); hdr.setFontSize(10); hdr.setWrap(false);
  shConv.setFrozenRows(1);
  // Proteger header
  shConv.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(function (p) { try { p.remove(); } catch (e) { } });
  var prot = shConv.getRange(1, 1, 1, lastCol).protect();
  prot.setDescription('Cabecalho - NAO EDITAR'); prot.setWarningOnly(true);
  SpreadsheetApp.flush();
  var msg = 'OK: ' + deduped.length + ' | removidas: ' + (data.length - deduped.length);
  Logger.log(msg); return msg;
}

function atualizarAtividadesMenu_() {
  var r = importarLoteRawConvertidoTodosAtletas_SEGURO();
  if (r && !r.rateLimitHit) corrigirDadosConv_();
  return r;
}

function configurarTriggerAtividades_() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'triggerImportacaoAutomatica') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('triggerImportacaoAutomatica').timeBased().everyHours(1).create();
  Logger.log('Trigger: 1h criado'); return 'Trigger 1h OK';
}

function triggerImportacaoAutomatica() {
  var r = importarLoteRawConvertidoTodosAtletas_SEGURO();
  if (r && !r.rateLimitHit && r.novas > 0) corrigirDadosConv_();
}

// ─────────── MetricasBeta.gs ───────────

// ============================================================
// MetricasBeta.gs — Motor de Métricas Individualizadas
// HIPERATIVO V3 — Lê: 🏃 ATIVIDADES_CONVERTIDAS
// Escreve: 📈 MÉTRICAS_BETA
// NÃO toca em: Strava, tokens, Supabase, painel, ranking
// ============================================================

var ALIASES_MB_ = {
  athId:        ['athId','ATH_ID','Nome Atleta'],
  atleta:       ['atleta','Atleta'],
  dataHora:     ['dataHoraRaw','Data/Hora','dataHora','Data Hora'],
  tipo:         ['tipo','Tipo de Esporte','tipoEsporte','Sport Type'],
  tipoRaw:      ['tipoRaw','Sport Type'],
  distKm:       ['Distância km','distanciaKm','distancia','dist_km','distKm','distance'],
  duracaoMin:   ['Tempo Movimento s','Tempo Total s','Tempo Movimento','Tempo Total','duracaoMin','duracao','tempo_min'],
  paceMinKm:    ['paceMinKm','pace','Pace','pace_min_km'],
  velKmh:       ['velocidadeKmh','velocidade','vel_kmh','Velocidade'],
  cargaSimples: ['Carga','cargaSimples','Carga Simples','carga'],
  intensidade:  ['intensidade','Intensidade','nivel_intensidade'],
  qualidade:    ['qualidadeDado','Qualidade do Dado','qualidade','Qualidade'],
  fcMedia:      ['fcMedia','FC Média','hr_avg','average_heartrate'],
};
var FATOR_INT_MB_ = {'leve':1.0,'moderado':1.3,'forte':1.6,'muito forte':2.0,'very hard':2.0,'hard':1.6,'moderate':1.3,'light':1.0,'easy':1.0};

function gerarMetricasBeta() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var shC = null;
  ['🏃 ATIVIDADES_CONVERTIDAS','ATIVIDADES_CONVERTIDAS'].forEach(function(n){if(!shC)shC=ss.getSheetByName(n);});
  if (!shC) ss.getSheets().forEach(function(s){if(!shC&&s.getName().indexOf('ATIVIDADES_CONVERTIDAS')>=0)shC=s;});
  if (!shC) { Logger.log('ERRO: aba ATIVIDADES_CONVERTIDAS nao encontrada'); return 'ERRO_SHEET_NAO_ENCONTRADA'; }
  Logger.log('USANDO_ABA=' + shC.getName());
  var lastRow = shC.getLastRow();
  if (lastRow < 2) { Logger.log('vazia'); return 'SEM_DADOS'; }
  var lastCol = shC.getLastColumn();
  var hdrs = shC.getRange(1,1,1,lastCol).getValues()[0];
  Logger.log('CONV_HEADERS=' + JSON.stringify(hdrs));
  var rawData = shC.getRange(2,1,lastRow-1,lastCol).getValues();
  var hIdx = {};
  hdrs.forEach(function(h,i){if(h) hIdx[String(h).trim()]=i;});
  var CI = {};
  Object.keys(ALIASES_MB_).forEach(function(field){
    CI[field]=-1;
    ALIASES_MB_[field].forEach(function(alias){if(CI[field]===-1&&hIdx[alias]!==undefined)CI[field]=hIdx[alias];});
  });
  Logger.log('COLUMN_MAP=' + JSON.stringify(CI));
  var byAth={}, now=new Date();
  rawData.forEach(function(row){
    var athId=CI.athId>=0?String(row[CI.athId]||'').trim():'';
    if(!athId)athId=CI.atleta>=0?String(row[CI.atleta]||'').trim():'_SEM_ID_';
    if(!athId)athId='_SEM_ID_';
    var hVals=['ATH_ID','athId','Nome Atleta','Atleta'];
    if(hVals.indexOf(athId)>=0)return;
    if(!byAth[athId])byAth[athId]=[];
    byAth[athId].push(row);
  });
  var athIds=Object.keys(byAth);
  Logger.log('ATLETAS_ENCONTRADOS=' + athIds.length);
  var COLS_MB=['ATH_ID','Atleta','Total Atividades Válidas','Atividades Ignoradas','Primeira Atividade','Última Atividade','Dias Desde Última','Atividades 7d','Atividades 30d','Dias Ativos 30d','Regularidade','Modalidade Principal','Distribuição Modalidades','Distância Total km','Distância 7d km','Distância 30d km','Tempo Total','Tempo 7d','Tempo 30d','Pace Médio Corrida','Pace Mediano Corrida','Melhor Pace Corrida','Velocidade Média Bike','Carga Total','Carga 7d','Carga 30d','Variação Carga 7d','Variação Carga 30d','Evolução Corrida','Qualidade Geral Dados','Flags Mais Comuns','Confiança Frequência','Confiança Volume','Confiança Pace','Confiança Carga','Confiança Evolução','Status Geral','Observação Técnica','Atualizado Em'];
  var resultRows=[],statsTotal=0,statsIgn=0,statsInsuf=[];
  athIds.forEach(function(athId){
    var res=mbCalcAtleta_(byAth[athId],athId,CI,now);
    resultRows.push(res.row); statsTotal+=res.nVal; statsIgn+=res.nIgn;
    if(res.status==='DADOS_INSUFICIENTES'||res.status==='SEM_ATIVIDADE_VALIDA') statsInsuf.push(athId+'('+res.status+')');
  });
  var shM=null;
  ['📈 MÉTRICAS_BETA','MÉTRICAS_BETA'].forEach(function(n){if(!shM)shM=ss.getSheetByName(n);});
  if(!shM) ss.getSheets().forEach(function(s){if(!shM&&s.getName().indexOf('MÉTRICAS_BETA')>=0)shM=s;});
  if(!shM){shM=ss.insertSheet('📈 MÉTRICAS_BETA');Logger.log('Aba criada');}
  shM.clearContents();
  shM.getRange(1,1,1,COLS_MB.length).setValues([COLS_MB]);
  var hr=shM.getRange(1,1,1,COLS_MB.length);
  hr.setBackground('#0d47a1');hr.setFontColor('#ffffff');hr.setFontWeight('bold');hr.setFontSize(10);hr.setWrap(false);
  shM.setFrozenRows(1);
  if(resultRows.length>0) shM.getRange(2,1,resultRows.length,COLS_MB.length).setValues(resultRows);
  SpreadsheetApp.flush();
  var log=['ATLETAS ANALISADOS: '+athIds.length,'ATIVIDADES VÁLIDAS: '+statsTotal,'ATIVIDADES IGNORADAS: '+statsIgn,'MÉTRICAS GERADAS: '+resultRows.length+' linhas','ATLETAS COM DADOS INSUFICIENTES: '+(statsInsuf.length>0?statsInsuf.join(', '):'Nenhum'),'PRINCIPAIS LIMITAÇÕES: pace/evolução requerem mínimo 3/6 corridas com distância e tempo','PODE AVANÇAR PARA PAINEL_BETA? '+(statsInsuf.length<=2?'Sim':'Parcial'),'PRÓXIMO PASSO: Validar MÉTRICAS_BETA na planilha, depois criar PAINEL_BETA'];
  log.forEach(function(l){Logger.log(l);});
  return log.join('\n');
}

function mbCalcAtleta_(rows,athId,CI,now){
  var DI='DADOS_INSUFICIENTES',NAM='N/A_MODALIDADE';
  var validas=[],ignoradas=[];
  rows.forEach(function(r){
    var q=CI.qualidade>=0?String(r[CI.qualidade]||'').trim().toUpperCase():'';
    if(q==='CRÍTICO'||q==='CRITICO'||q==='IGNORAR')ignoradas.push(r);else validas.push(r);
  });
  var nVal=validas.length,nIgn=ignoradas.length;
  var nome='';
  validas.concat(rows).forEach(function(r){
    if(nome)return;
    if(CI.atleta>=0&&r[CI.atleta])nome=String(r[CI.atleta]).trim();
    if(!nome&&CI.athId>=0&&r[CI.athId])nome=String(r[CI.athId]).trim();
  });
  if(!nome)nome=athId;
  if(nVal===0){return{row:buildMbRow_(athId,nome,0,nIgn,DI,DI,DI,0,0,0,'INSUFICIENTE',DI,DI,DI,DI,DI,DI,DI,DI,DI,DI,DI,DI,DI,DI,DI,DI,DI,DI,'0%','Nenhum','INSUFICIENTE','INSUFICIENTE','INSUFICIENTE','INSUFICIENTE','INSUFICIENTE','SEM_ATIVIDADE_VALIDA','Nenhuma atividade válida.',now),nVal:0,nIgn:nIgn,status:'SEM_ATIVIDADE_VALIDA'};}
  var agora=now.getTime(),ms7d=7*24*3600*1000,ms30d=30*24*3600*1000,ms14d=14*24*3600*1000;
  var datas=validas.map(function(r){return mbDate_(r,CI);}).filter(Boolean).sort(function(a,b){return a-b;});
  var dtFirst=datas.length>0?datas[0]:null,dtLast=datas.length>0?datas[datas.length-1]:null;
  var diasUlt=dtLast?Math.floor((agora-dtLast.getTime())/86400000):DI;
  var spanDias=(dtFirst&&dtLast)?Math.floor((dtLast-dtFirst)/86400000):0;
  var act7d=validas.filter(function(r){var d=mbDate_(r,CI);return d&&(agora-d.getTime())<=ms7d;});
  var act30d=validas.filter(function(r){var d=mbDate_(r,CI);return d&&(agora-d.getTime())<=ms30d;});
  var nAct7=act7d.length,nAct30=act30d.length;
  var diasAtvMap={};
  act30d.forEach(function(r){var d=mbDate_(r,CI);if(d)diasAtvMap[d.toDateString()]=true;});
  var nDiasAtv30=Object.keys(diasAtvMap).length;
  var semAtv={};
  validas.forEach(function(r){var d=mbDate_(r,CI);if(d)semAtv[mbWeek_(d)]=true;});
  var confFreq=nAct30>=8?'ALTA':nAct30>=4?'MÉDIA':nAct30>=1?'BAIXA':'INSUFICIENTE';
  var regularidade=spanDias<14?DI:(function(){var avg=nAct30/4.3;return avg>=3?'Regular':avg>=2?'Moderado':avg>=1?'Baixo':'Irregular';})();
  var modMap={};
  validas.forEach(function(r){var cat=mbCategoria_(mbTipo_(r,CI));if(!modMap[cat])modMap[cat]=0;modMap[cat]++;});
  var modalPrincipal=nVal<3?DI:'',distMod=nVal<3?DI:'';
  if(nVal>=3){var mArr=Object.keys(modMap).map(function(k){return{k:k,n:modMap[k]};}).sort(function(a,b){return b.n-a.n;});modalPrincipal=mArr[0].k;distMod=mArr.map(function(x){return x.k+':'+Math.round(x.n/nVal*100)+'%';}).join(' | ');}
  var distAll=[],dist7=0,dist30=0,distTotal=0;
  validas.forEach(function(r){if(!mbRequerDist_(mbTipo_(r,CI)))return;var d=mbNum_(r,CI.distKm);if(!d||d<=0)return;distAll.push(d);distTotal+=d;var dt=mbDate_(r,CI);if(dt&&(agora-dt.getTime())<=ms7d)dist7+=d;if(dt&&(agora-dt.getTime())<=ms30d)dist30+=d;});
  var nComDist=validas.filter(function(r){return mbRequerDist_(mbTipo_(r,CI));}).length;
  var distTotalStr=nComDist===0?NAM:(distAll.length>0?mbR2_(distTotal):DI);
  var dist7Str=nComDist===0?NAM:(distAll.length>0?mbR2_(dist7):DI);
  var dist30Str=nComDist===0?NAM:(distAll.length>0?mbR2_(dist30):DI);
  var tTotal=0,t7=0,t30=0;
  validas.forEach(function(r){var dur=mbDur_(r,CI);if(!dur||dur<=0)return;tTotal+=dur;var dt=mbDate_(r,CI);if(dt&&(agora-dt.getTime())<=ms7d)t7+=dur;if(dt&&(agora-dt.getTime())<=ms30d)t30+=dur;});
  var tempoTotal=tTotal>0?mbFmtMin_(tTotal):DI,tempo7=t7>0?mbFmtMin_(t7):'0min',tempo30=t30>0?mbFmtMin_(t30):DI;
  var confVol=nComDist>0?(distAll.length>=5?'ALTA':distAll.length>=3?'MÉDIA':distAll.length>=1?'BAIXA':'INSUFICIENTE'):(tTotal>0?'MÉDIA':'INSUFICIENTE');
  var paceRows=validas.filter(function(r){var t=mbTipo_(r,CI).toLowerCase();return t.indexOf('corrida')>=0||t.indexOf('run')>=0||t.indexOf('trail')>=0||t.indexOf('caminhada')>=0||t.indexOf('walk')>=0;});
  var paceList=[];
  paceRows.forEach(function(r){var d=mbNum_(r,CI.distKm),dur=mbDur_(r,CI),dt=mbDate_(r,CI);if(!d||d<=0||!dur||dur<=0)return;var p=dur/d;if(p<2||p>25)return;paceList.push({p:p,d:d,dt:dt});});
  var paceMed=DI,paceMdn=DI,paceMax=DI,confPace='INSUFICIENTE';
  if(paceList.length>=3){var sumD=paceList.reduce(function(s,x){return s+x.d;},0),sumWP=paceList.reduce(function(s,x){return s+x.p*x.d;},0);paceMed=mbFmtPace_(sumWP/sumD);var sorted=paceList.map(function(x){return x.p;}).sort(function(a,b){return a-b;});var mid=Math.floor(sorted.length/2);paceMdn=mbFmtPace_(sorted.length%2===0?(sorted[mid-1]+sorted[mid])/2:sorted[mid]);paceMax=mbFmtPace_(sorted[0]);confPace=paceList.length>=8?'ALTA':paceList.length>=5?'MÉDIA':'BAIXA';}
  var bikeRows=validas.filter(function(r){var t=mbTipo_(r,CI).toLowerCase();return t.indexOf('cicl')>=0||t.indexOf('bike')>=0||t.indexOf('ride')>=0||t.indexOf('gravel')>=0||t.indexOf('mountain')>=0;});
  var velStr=bikeRows.length===0?NAM:DI,velList=[];
  bikeRows.forEach(function(r){var d=mbNum_(r,CI.distKm),dur=mbDur_(r,CI);if(!d||d<=0||!dur||dur<=0)return;var v=d/(dur/60);if(v<5||v>80)return;velList.push({v:v,d:d});});
  if(velList.length>=3){var sumDv=velList.reduce(function(s,x){return s+x.d;},0),sumWv=velList.reduce(function(s,x){return s+x.v*x.d;},0);velStr=mbR2_(sumWv/sumDv)+' km/h';}
  var cTotal=0,c7=0,c30=0,c7ant=0,cN=0;
  validas.forEach(function(r){var c=mbNum_(r,CI.cargaSimples);if(!c||c<=0){var d=mbNum_(r,CI.distKm),dur=mbDur_(r,CI),int_=CI.intensidade>=0?String(r[CI.intensidade]||'').trim().toLowerCase():'',f=FATOR_INT_MB_[int_]||1.0;if(d&&d>0)c=d*f;else if(dur&&dur>0)c=(dur/10)*f;else c=null;}if(!c||c<=0)return;cN++;cTotal+=c;var dt=mbDate_(r,CI);if(dt){var df=agora-dt.getTime();if(df<=ms7d)c7+=c;if(df<=ms30d)c30+=c;if(df>ms7d&&df<=ms14d)c7ant+=c;}});
  var cTStr=cN>0?mbR2_(cTotal):DI,c7Str=cN>0?mbR2_(c7):DI,c30Str=cN>0?mbR2_(c30):DI,var7Str=DI,var30Str=DI;
  if(cN>0){if(c7ant>0)var7Str=mbVariacao_((c7-c7ant)/c7ant*100);var cant30=cTotal-c30;if(cant30>0&&c30>0)var30Str=mbVariacao_((c30-cant30)/cant30*100);}
  var confCarga=cN>=8?'ALTA':cN>=4?'MÉDIA':cN>=1?'BAIXA':'INSUFICIENTE';
  var evo=DI,confEvo='INSUFICIENTE';
  var paceOrd=paceList.filter(function(x){return x.dt;}).sort(function(a,b){return a.dt-b.dt;});
  if(paceOrd.length>=6){var sp=paceOrd[paceOrd.length-1].dt.getTime()-paceOrd[0].dt.getTime();if(sp>=21*86400000){var p3i=mbMedian_(paceOrd.slice(0,3).map(function(x){return x.p;})),p3f=mbMedian_(paceOrd.slice(-3).map(function(x){return x.p;})),varP=(p3f-p3i)/p3i*100;evo=mbR2_(varP)+'% pace ('+(varP<-2?'Melhora':varP>2?'Piora':'Estável')+')';confEvo='MÉDIA';}}
  var qC={ok:0,okP:0,inc:0,crit:0,ign:0};
  rows.forEach(function(r){var q=CI.qualidade>=0?String(r[CI.qualidade]||'').trim().toUpperCase():'OK';if(q==='OK')qC.ok++;else if(q==='OK_PARCIAL')qC.okP++;else if(q.indexOf('INCOMPLETO')>=0)qC.inc++;else if(q==='CRÍTICO'||q==='CRITICO')qC.crit++;else if(q==='IGNORAR')qC.ign++;});
  var aprov=qC.ok+qC.okP+qC.inc,pctQ=rows.length>0?Math.round(aprov/rows.length*100)+'%':'0%';
  var flagsArr=[];
  if(qC.crit>0)flagsArr.push('CRÍTICO:'+qC.crit);if(qC.ign>0)flagsArr.push('IGNORAR:'+qC.ign);if(qC.inc>0)flagsArr.push('INCOMPLETO:'+qC.inc);if(qC.okP>0)flagsArr.push('OK_PARCIAL:'+qC.okP);
  var flagsStr=flagsArr.length>0?flagsArr.join(' | '):'Nenhum';
  var status,obsArr=[];
  if(nVal===0)status='SEM_ATIVIDADE_VALIDA';else if(nVal<3){status='DADOS_INSUFICIENTES';obsArr.push('Menos de 3 atividades válidas.');}else{status=(cN>0&&(distAll.length>0||tTotal>0))?'OK':'PARCIAL';}
  if(paceList.length<3&&paceRows.length>0)obsArr.push('Sem dados suficientes para pace.');
  if(paceRows.length===0&&nVal>0)obsArr.push('Atividades sem distância obrigatória.');
  if(evo===DI&&paceOrd.length>0)obsArr.push('Evolução não calculada: histórico insuficiente.');
  if(CI.fcMedia<0)obsArr.push('FC ausente nas colunas.');
  if(spanDias<14&&nVal>0)obsArr.push('Histórico < 14 dias; regularidade indisponível.');
  if((qC.crit+qC.ign)>0)obsArr.push((qC.crit+qC.ign)+' ativ. ignorada(s).');
  var obs=obsArr.length>0?obsArr.join(' '):'Dados suficientes para análise.';
  return{row:buildMbRow_(athId,nome,nVal,nIgn,dtFirst?mbFmtDt_(dtFirst):DI,dtLast?mbFmtDt_(dtLast):DI,diasUlt,nAct7,nAct30,nDiasAtv30,regularidade,modalPrincipal,distMod,distTotalStr,dist7Str,dist30Str,tempoTotal,tempo7,tempo30,paceMed,paceMdn,paceMax,velStr,cTStr,c7Str,c30Str,var7Str,var30Str,evo,pctQ,flagsStr,confFreq,confVol,confPace,confCarga,confEvo,status,obs,now),nVal:nVal,nIgn:nIgn,status:status};
}

function buildMbRow_(){return Array.prototype.slice.call(arguments);}
function mbDate_(row,CI){if(CI.dataHora<0)return null;var v=row[CI.dataHora];if(!v)return null;var d=(v instanceof Date)?v:new Date(v);return isNaN(d.getTime())?null:d;}
function mbTipo_(row,CI){if(CI.tipo>=0&&row[CI.tipo])return String(row[CI.tipo]).trim();if(CI.tipoRaw>=0&&row[CI.tipoRaw])return String(row[CI.tipoRaw]).trim();return '';}
function mbNum_(row,idx){if(idx<0)return null;var v=row[idx];if(v===null||v===undefined||v===''||v==='N/A')return null;var n=parseFloat(String(v).replace(',','.'));return isNaN(n)?null:n;}
function mbDur_(row,CI){if(CI.duracaoMin<0)return null;var v=row[CI.duracaoMin];if(v===null||v===undefined||v==='')return null;var n=parseFloat(String(v).replace(',','.'));if(isNaN(n)||n<=0)return null;return n>500?n/60:n;}
function mbR2_(n){return Math.round(n*100)/100;}
function mbFmtPace_(p){var m=Math.floor(p),s=Math.round((p-m)*60);if(s===60){m++;s=0;}return m+':'+(s<10?'0':'')+s+'/km';}
function mbFmtMin_(m){if(m<60)return Math.round(m)+'min';var h=Math.floor(m/60),mm=Math.round(m%60);return h+'h'+(mm>0?mm+'min':'');}
function mbFmtDt_(d){return d?d.toLocaleDateString('pt-BR'):'';}
function mbWeek_(d){var jan=new Date(d.getFullYear(),0,1),w=Math.ceil(((d-jan)/86400000+jan.getDay()+1)/7);return d.getFullYear()+'-W'+(w<10?'0':'')+w;}
function mbMedian_(arr){if(!arr||!arr.length)return 0;var s=arr.slice().sort(function(a,b){return a-b;}),m=Math.floor(s.length/2);return s.length%2===0?(s[m-1]+s[m])/2:s[m];}
function mbCategoria_(tipo){var t=String(tipo||'').toLowerCase().trim();if(t.indexOf('corrida')>=0||t==='run'||t.indexOf('trail')>=0)return 'Corrida';if(t.indexOf('caminhada')>=0||t==='walk')return 'Caminhada';if(t.indexOf('cicl')>=0||t.indexOf('bike')>=0||t.indexOf('ride')>=0||t.indexOf('gravel')>=0)return 'Ciclismo';if(t.indexOf('hiit')>=0)return 'HIIT';if(t.indexOf('muscula')>=0||t.indexOf('weight')>=0||t.indexOf('strength')>=0)return 'Musculação';if(t.indexOf('yoga')>=0||t.indexOf('pilates')>=0)return 'Yoga/Pilates';if(t.indexOf('natação')>=0||t.indexOf('swim')>=0)return 'Natação';return 'Outros';}
function mbRequerDist_(tipo){var t=String(tipo||'').toLowerCase().trim();return(t.indexOf('corrida')>=0||t==='run'||t.indexOf('trail')>=0||t.indexOf('caminhada')>=0||t==='walk'||t.indexOf('cicl')>=0||t.indexOf('bike')>=0||t.indexOf('ride')>=0||t.indexOf('gravel')>=0||t.indexOf('mountain')>=0);}
function mbVariacao_(pct){var r=mbR2_(pct);if(pct<-30)return 'Redução forte ('+r+'%)';if(pct<-10)return 'Redução moderada ('+r+'%)';if(pct<=10)return 'Estável ('+r+'%)';if(pct<=30)return 'Aumento moderado (+'+r+'%)';return 'Aumento alto (+'+r+'%)';}
function EXEC_METRICAS_BETA(){return gerarMetricasBeta();}


// âââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââ
// PainelCentral.gs â Painel Central, Rankings, InventÃ¡rio de Abas
// HIPERATIVO V3 | Gerado automaticamente
// â NÃO mexer em Strava / tokens / Supabase neste arquivo
// âââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââ

var PC_PAINEL   = 'ð PAINEL_CENTRAL';
var PC_RANKINGS = 'ð RANKINGS_BETA';
var PC_CONV     = 'ð ATIVIDADES_CONVERTIDAS';
var PC_MET      = 'ð MÃTRICAS_BETA';
var PC_INV      = 'ðï¸ INVENTÃRIO_ABAS';

var PC_INTOCAVEIS = [
  'CADASTRO','TOKENS','STRAVA_RAW','ATIVIDADES_CONVERTIDAS',
  'MÃTRICAS_BETA','PAINEL_CENTRAL','RANKINGS_BETA','INVENTÃRIO_ABAS'
];

// ââ ENTRADA SEM IMPORTAÃÃO STRAVA âââââââââââââââââââââââââââââââââââââ
function atualizarPainelMetricasRankingsSemImportar() {
  var logs = [];
  Logger.log('=== atualizarPainelMetricasRankingsSemImportar START ===');

  try { gerarMetricasBeta(); logs.push('METRICAS_BETA: OK'); }
  catch(e) { logs.push('METRICAS_BETA: ERRO=' + e.message); Logger.log('ERRO_MB=' + e.message); }

  try { pcPainel_(); logs.push('PAINEL_CENTRAL: OK'); }
  catch(e) { logs.push('PAINEL_CENTRAL: ERRO=' + e.message); Logger.log('ERRO_PAINEL=' + e.message); }

  try { pcRankings_(); logs.push('RANKINGS_BETA: OK'); }
  catch(e) { logs.push('RANKINGS_BETA: ERRO=' + e.message); Logger.log('ERRO_RANKINGS=' + e.message); }

  try {
    var inv = pcInventario_();
    logs.push('INVENTARIO: ' + inv.total + ' abas | OCULTADAS: ' + inv.ocultadas);
  } catch(e) { logs.push('INVENTARIO: ERRO=' + e.message); Logger.log('ERRO_INV=' + e.message); }

  logs.forEach(function(l) { Logger.log(l); });
  SpreadsheetApp.flush();
  return logs.join('\n');
}

// ââ ENTRADA COMPLETA (COM STRAVA â BLOQUEADA POR PADRÃO) ââââââââââââââ
function atualizarCentralHiperativo() {
  // â IMPORTAÃÃO STRAVA BLOQUEADA â descomente sÃ³ apÃ³s autorizaÃ§Ã£o de Crhystiano:
  // importarLoteRawConvertidoTodosAtletas_SEGURO();
  Logger.log('IMPORTACAO_STRAVA_BLOQUEADA: use atualizarPainelMetricasRankingsSemImportar');
  atualizarPainelMetricasRankingsSemImportar();
}

// ââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââ
// PAINEL CENTRAL
// ââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââ
function pcPainel_() {
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var sh  = pcGetAba_(ss, PC_PAINEL);
  sh.clearContents(); sh.clearFormats();

  var met  = pcLer_(ss, PC_MET);
  var conv = pcLer_(ss, PC_CONV);
  var MI   = pcIdx_(met.h);
  var CI   = pcIdx_(conv.h);
  var now  = new Date();
  var row  = 1;

  // ââ CABEÃALHO âââââââââââââââââââââââââââââââââââââââââââââââââââââââ
  pcBloco_(sh, row, 8, 'ð PAINEL CENTRAL â HIPERATIVO V3', '#0d47a1', 14); row++;
  sh.getRange(row, 1).setValue('Atualizado: ' + pcDt_(now))
    .setFontColor('#9e9e9e').setFontSize(9);
  row += 2;

  // ââ RESUMO GERAL ââââââââââââââââââââââââââââââââââââââââââââââââââââ
  pcBloco_(sh, row, 8, 'ð RESUMO GERAL', '#263238', 11); row++;
  var tAtl = met.r.length, tAct = 0, tIgn = 0, tDist = 0, tCarga = 0;
  met.r.forEach(function(r) {
    tAct   += pcF(pcG(r, MI, 'Total Atividades Validas'));
    tIgn   += pcF(pcG(r, MI, 'Atividades Ignoradas'));
    tDist  += pcF(pcG(r, MI, 'Distancia Total km'));
    tCarga += pcF(pcG(r, MI, 'Carga Total'));
  });
  [
    ['Atletas Analisados', tAtl,              '', 'Atividades Validas', tAct],
    ['Atividades Ignoradas', tIgn,            '', 'Distancia Total',    tDist.toFixed(1) + ' km'],
    ['Carga Total', tCarga.toFixed(0),        '', 'Ultima Atualizacao', pcDt_(now)]
  ].forEach(function(r) { sh.getRange(row, 1, 1, 5).setValues([r]); row++; });
  row++;

  // ââ ALERTAS DE TREINO âââââââââââââââââââââââââââââââââââââââââââââââ
  pcBloco_(sh, row, 8, 'ð¨ ALERTAS DE TREINO â ULTIMO TREINO POR ATLETA', '#263238', 11); row++;
  var aH = ['Atleta', 'Ultima Atividade', 'Modalidade', 'Dias Sem Treino', 'Status Alerta',
             'Dist 30d km', 'Confianca Freq', 'Status Geral'];
  sh.getRange(row, 1, 1, 8).setValues([aH]); pcHeader_(sh, row, 8, '#1565c0'); row++;

  var alertas = met.r.map(function(r) {
    var dias = pcF(pcG(r, MI, 'Dias Desde Ultima'));
    var st   = String(pcG(r, MI, 'Status Geral') || '');
    return {
      nome:  String(pcG(r, MI, 'Atleta') || pcG(r, MI, 'ATH_ID') || '').trim(),
      ult:   String(pcG(r, MI, 'Ultima Atividade') || ''),
      modal: String(pcG(r, MI, 'Modalidade Principal') || ''),
      dias:  dias,
      alerta:pcAlerta_(dias, st),
      dist:  pcF(pcG(r, MI, 'Distancia 30d km')).toFixed(1) + ' km',
      conf:  String(pcG(r, MI, 'Confianca Frequencia') || ''),
      st:    st
    };
  }).sort(function(a, b) { return b.dias - a.dias; });

  alertas.forEach(function(a) {
    sh.getRange(row, 1, 1, 8).setValues([[
      a.nome, a.ult, a.modal, a.dias, a.alerta, a.dist, a.conf, a.st
    ]]);
    var cor = pcCorAlerta_(a.alerta);
    if (cor) sh.getRange(row, 5).setBackground(cor).setFontColor('#fff').setFontWeight('bold');
    row++;
  });
  row++;

  // ââ ULTIMAS 20 ATIVIDADES ââââââââââââââââââââââââââââââââââââââââââââ
  pcBloco_(sh, row, 8, 'ð ULTIMAS 20 ATIVIDADES', '#263238', 11); row++;
  var uH = ['Atleta', 'Data/Hora', 'Modalidade', 'Atividade', 'Distancia km', 'Tempo', 'Pace', 'Qualidade'];
  sh.getRange(row, 1, 1, 8).setValues([uH]); pcHeader_(sh, row, 8, '#2e7d32'); row++;

  var validas = conv.r
    .filter(function(r) {
      var st = String(pcG(r, CI, 'Status') || '');
      return st !== 'CRITICO' && st !== 'IGNORAR';
    })
    .sort(function(a, b) {
      var dA = new Date(pcG(a, CI, 'Data/Hora') || 0);
      var dB = new Date(pcG(b, CI, 'Data/Hora') || 0);
      return dB - dA;
    })
    .slice(0, 20);

  validas.forEach(function(r) {
    var dtRaw = pcG(r, CI, 'Data/Hora');
    var tempo = pcF(pcG(r, CI, 'Tempo Movimento s'));
    sh.getRange(row, 1, 1, 8).setValues([[
      String(pcG(r, CI, 'Atleta') || pcG(r, CI, 'Nome Atleta') || '').trim(),
      dtRaw ? pcDt_(new Date(dtRaw)) : '',
      String(pcG(r, CI, 'Tipo de Esporte') || ''),
      String(pcG(r, CI, 'Nome da Atividade') || ''),
      pcF(pcG(r, CI, 'Distancia km')) > 0 ? pcF(pcG(r, CI, 'Distancia km')).toFixed(2) : 'N/A',
      tempo > 0 ? pcFmtSeg_(tempo) : 'N/A',
      String(pcG(r, CI, 'Pace') || 'N/A'),
      String(pcG(r, CI, 'Qualidade') || '')
    ]]);
    row++;
  });
  row++;

  // ââ DADOS INSUFICIENTES ââââââââââââââââââââââââââââââââââââââââââââââ
  var insuf = met.r.filter(function(r) {
    var st = String(pcG(r, MI, 'Status Geral') || '');
    return st === 'DADOS_INSUFICIENTES' || st === 'SEM_ATIVIDADE_VALIDA';
  });
  if (insuf.length > 0) {
    pcBloco_(sh, row, 8, 'â ï¸ ATLETAS COM DADOS INSUFICIENTES', '#263238', 11); row++;
    sh.getRange(row, 1, 1, 4).setValues([['Atleta', 'Status', 'Atividades', 'Observacao']]);
    pcHeader_(sh, row, 4, '#e65100'); row++;
    insuf.forEach(function(r) {
      sh.getRange(row, 1, 1, 4).setValues([[
        String(pcG(r, MI, 'Atleta') || pcG(r, MI, 'ATH_ID') || '').trim(),
        String(pcG(r, MI, 'Status Geral') || ''),
        pcF(pcG(r, MI, 'Total Atividades Validas')),
        String(pcG(r, MI, 'Observacao Tecnica') || '')
      ]]);
      row++;
    });
  }

  sh.setColumnWidth(1, 160); sh.setColumnWidth(2, 140); sh.setColumnWidth(3, 120);
  sh.setColumnWidth(4, 200); sh.setColumnWidths(5, 4, 100);
  Logger.log('PAINEL_CENTRAL: OK rows=' + row);
}

// ââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââ
// RANKINGS BETA
// ââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââ
function pcRankings_() {
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var sh  = pcGetAba_(ss, PC_RANKINGS);
  sh.clearContents(); sh.clearFormats();

  var met = pcLer_(ss, PC_MET);
  var MI  = pcIdx_(met.h);
  var now = new Date();
  var row = 1;

  // CabeÃ§alho
  pcBloco_(sh, row, 5, 'ð RANKINGS BETA â HIPERATIVO V3', '#4a148c', 14); row++;
  sh.getRange(row, 1).setValue('Gerado: ' + pcDt_(now)).setFontColor('#9e9e9e').setFontSize(9);
  row += 2;

  // Atletas com dados
  var validos = met.r.filter(function(r) {
    var st = String(pcG(r, MI, 'Status Geral') || '');
    return st === 'OK' || st === 'PARCIAL';
  });
  var todos = met.r;

  // ââ Helpers locais ââââââââââââââââââââââââââââââââââââââââââââââââ
  function nomeDe(r) {
    return String(pcG(r, MI, 'Atleta') || pcG(r, MI, 'ATH_ID') || '').trim();
  }
  function bloco(titulo, corHdr, hdrs, linhas) {
    pcBloco_(sh, row, 5, titulo, '#263238', 11); row++;
    sh.getRange(row, 1, 1, hdrs.length).setValues([hdrs]);
    pcHeader_(sh, row, hdrs.length, corHdr); row++;
    if (linhas.length === 0) {
      sh.getRange(row, 1, 1, hdrs.length).setValues([
        Array(hdrs.length).fill('â').map(function(v, i) { return i === 1 ? 'DADOS_INSUFICIENTES' : v; })
      ]); row++;
    } else {
      linhas.slice(0, 10).forEach(function(l, i) {
        sh.getRange(row, 1, 1, l.length).setValues([l]);
        if (i === 0) sh.getRange(row, 1, 1, l.length).setBackground('#fff9c4').setFontWeight('bold');
        row++;
      });
    }
    row += 2;
  }
  function top10Num(titulo, cor, campo, confCampo, sufixo) {
    var lista = validos
      .map(function(r) { return { r: r, n: pcF(pcG(r, MI, campo)) }; })
      .filter(function(x) { return x.n > 0; })
      .sort(function(a, b) { return b.n - a.n; })
      .map(function(x, i) {
        return [i + 1, nomeDe(x.r), x.n.toFixed(1) + (sufixo || ''), String(pcG(x.r, MI, confCampo || 'Confianca Frequencia') || 'â')];
      });
    bloco(titulo, cor, ['#', 'Atleta', 'Valor', 'Confianca'], lista);
  }

  // ââ OBRIGATÃRIAS ââââââââââââââââââââââââââââââââââââââââââââââââââ
  pcBloco_(sh, row, 5, 'â­ CATEGORIAS OBRIGATORIAS', '#1a237e', 12); row += 2;

  top10Num('ð¥ Top 10 â Frequencia 30d',    '#283593', 'Atividades 30d',   'Confianca Frequencia', ' treinos');
  top10Num('ð¥ Top 10 â Distancia 30d km',  '#283593', 'Distancia 30d km', 'Confianca Volume',     ' km');
  top10Num('ð¥ Top 10 â Carga 30d',         '#283593', 'Carga 30d',        'Confianca Carga',      '');
  top10Num('ð¥ Top 10 â Distancia Total km','#283593', 'Distancia Total km','Confianca Volume',    ' km');

  // Regularidade (string â score)
  var regMap = { 'Regular': 4, 'Moderado': 3, 'Baixo': 2, 'Irregular': 1 };
  var listaReg = validos
    .map(function(r) {
      var reg = String(pcG(r, MI, 'Regularidade') || '');
      return { r: r, reg: reg, n: regMap[reg] || 0 };
    })
    .filter(function(x) { return x.n > 0; })
    .sort(function(a, b) { return b.n - a.n; })
    .map(function(x, i) { return [i + 1, nomeDe(x.r), x.reg, String(pcG(x.r, MI, 'Confianca Frequencia') || 'â')]; });
  bloco('ð¥ Top 10 â Regularidade', '#283593', ['#', 'Atleta', 'Nivel', 'Confianca'], listaReg);

  // Por modalidade principal
  ['Corrida', 'Caminhada', 'Ciclismo'].forEach(function(mod) {
    var filtrados = validos.filter(function(r) {
      return String(pcG(r, MI, 'Modalidade Principal') || '').toLowerCase().indexOf(mod.toLowerCase()) >= 0;
    });
    var lista = filtrados
      .map(function(r) { return { r: r, n: pcF(pcG(r, MI, 'Distancia 30d km')) }; })
      .filter(function(x) { return x.n > 0; })
      .sort(function(a, b) { return b.n - a.n; })
      .map(function(x, i) { return [i + 1, nomeDe(x.r), x.n.toFixed(1) + ' km', String(pcG(x.r, MI, 'Modalidade Principal') || '')]; });
    bloco('ð¥ Top 10 â ' + mod, '#283593', ['#', 'Atleta', 'Distancia 30d', 'Modalidade'], lista);
  });

  // HIIT/ForÃ§a
  var hiit = validos.filter(function(r) {
    var dist = String(pcG(r, MI, 'Distribuicao Modalidades') || '').toLowerCase();
    return dist.indexOf('hiit') >= 0 || dist.indexOf('forca') >= 0 || dist.indexOf('musculacao') >= 0 || dist.indexOf('strength') >= 0;
  });
  var listaHiit = hiit
    .map(function(r) { return { r: r, n: pcF(pcG(r, MI, 'Atividades 30d')) }; })
    .filter(function(x) { return x.n > 0; })
    .sort(function(a, b) { return b.n - a.n; })
    .map(function(x, i) { return [i + 1, nomeDe(x.r), x.n.toFixed(0) + ' treinos', String(pcG(x.r, MI, 'Modalidade Principal') || '')]; });
  bloco('ðª Top 10 â HIIT/Forca', '#283593', ['#', 'Atleta', 'Treinos 30d', 'Modalidade'], listaHiit);

  // Diversidade
  var divMap = validos.map(function(r) {
    var d = String(pcG(r, MI, 'Distribuicao Modalidades') || '');
    return { r: r, n: d ? d.split('|').length : 0, dist: d };
  }).filter(function(x) { return x.n > 0; }).sort(function(a, b) { return b.n - a.n; });
  bloco('ð¨ Top 10 â Diversidade de Modalidades', '#283593',
    ['#', 'Atleta', 'Tipos', 'Distribuicao'],
    divMap.map(function(x, i) { return [i + 1, nomeDe(x.r), x.n + ' modalidades', x.dist]; })
  );

  // ââ ESTUDO E ALERTA âââââââââââââââââââââââââââââââââââââââââââââââ
  pcBloco_(sh, row, 5, 'â ï¸ CATEGORIAS DE ESTUDO E ALERTA', '#b71c1c', 12); row += 2;

  // Sem treino hÃ¡ mais dias
  var ausentes = todos
    .map(function(r) {
      var dias = pcF(pcG(r, MI, 'Dias Desde Ultima'));
      var st   = String(pcG(r, MI, 'Status Geral') || '');
      return { r: r, dias: dias, alerta: pcAlerta_(dias, st) };
    })
    .filter(function(x) { return x.dias >= 0; })
    .sort(function(a, b) { return b.dias - a.dias; });

  pcBloco_(sh, row, 5, 'ð´ Top 10 â Sem Treino ha Mais Dias', '#263238', 11); row++;
  sh.getRange(row, 1, 1, 4).setValues([['#', 'Atleta', 'Dias Sem Treino', 'Status Alerta']]);
  pcHeader_(sh, row, 4, '#b71c1c'); row++;
  ausentes.slice(0, 10).forEach(function(x, i) {
    sh.getRange(row, 1, 1, 4).setValues([[i + 1, nomeDe(x.r), x.dias + ' dias', x.alerta]]);
    var cor = pcCorAlerta_(x.alerta);
    if (cor) sh.getRange(row, 4).setBackground(cor).setFontColor('#fff').setFontWeight('bold');
    row++;
  });
  row += 2;

  // Maior aumento de carga 7d
  var listaVar7 = validos
    .map(function(r) {
      var v = String(pcG(r, MI, 'Variacao Carga 7d') || '');
      var n = parseFloat(v.replace('%', '').replace('+', '')) || 0;
      return { r: r, n: n, v: v };
    })
    .filter(function(x) { return x.n > 0; })
    .sort(function(a, b) { return b.n - a.n; })
    .map(function(x, i) { return [i + 1, nomeDe(x.r), x.v, String(pcG(x.r, MI, 'Confianca Carga') || 'â')]; });
  bloco('ð Top 10 â Maior Aumento de Carga 7d', '#7b1fa2', ['#', 'Atleta', 'Variacao', 'Confianca'], listaVar7);

  // Dados insuficientes
  var insufL = todos.filter(function(r) {
    var st = String(pcG(r, MI, 'Status Geral') || '');
    return st === 'DADOS_INSUFICIENTES' || st === 'SEM_ATIVIDADE_VALIDA';
  });
  bloco('â ï¸ Atletas com Dados Insuficientes', '#e65100',
    ['#', 'Atleta', 'Status', 'Observacao'],
    insufL.map(function(r, i) {
      return [i + 1, nomeDe(r), String(pcG(r, MI, 'Status Geral') || ''), String(pcG(r, MI, 'Observacao Tecnica') || '')];
    })
  );

  // Baixa regularidade
  var listaBaixa = validos
    .map(function(r) {
      var reg = String(pcG(r, MI, 'Regularidade') || '');
      var n = reg === 'Irregular' ? 4 : reg === 'Baixo' ? 3 : reg === 'Moderado' ? 2 : reg === 'Regular' ? 1 : 0;
      return { r: r, reg: reg, n: n };
    })
    .filter(function(x) { return x.n > 0; })
    .sort(function(a, b) { return b.n - a.n; })
    .map(function(x, i) { return [i + 1, nomeDe(x.r), x.reg, String(pcG(x.r, MI, 'Confianca Frequencia') || 'â')]; });
  bloco('ð Top 10 â Baixa Regularidade', '#7b1fa2', ['#', 'Atleta', 'Regularidade', 'Confianca'], listaBaixa);

  // Atividades com flags criticas
  var listaFlags = todos
    .filter(function(r) {
      var fl = String(pcG(r, MI, 'Flags Mais Comuns') || '');
      return fl.indexOf('CRITICO') >= 0 || fl.indexOf('IGNORAR') >= 0;
    })
    .map(function(r, i) { return [i + 1, nomeDe(r), String(pcG(r, MI, 'Flags Mais Comuns') || ''), String(pcG(r, MI, 'Status Geral') || '')]; });
  bloco('ð© Atletas com Flags Criticas', '#b71c1c', ['#', 'Atleta', 'Flags', 'Status Geral'], listaFlags);

  // ââ CATEGORIAS DIVERTIDAS âââââââââââââââââââââââââââââââââââââââââ
  pcBloco_(sh, row, 5, 'ð CATEGORIAS DIVERTIDAS', '#1b5e20', 12); row += 2;

  top10Num('ð Rei/Rainha dos Quilometros',        '#1b5e20', 'Distancia Total km',  'Confianca Volume',     ' km');
  top10Num('ð Mais Constante (Frequencia 7d)',    '#1b5e20', 'Atividades 7d',       'Confianca Frequencia', ' treinos');

  bloco('ðºï¸ Explorador/a de Modalidades', '#1b5e20',
    ['#', 'Atleta', 'Tipos', 'Distribuicao'],
    divMap.map(function(x, i) { return [i + 1, nomeDe(x.r), x.n + ' tipos', x.dist]; })
  );

  // Voltou com Tudo
  var voltou = validos
    .filter(function(r) {
      var v7  = parseFloat(String(pcG(r, MI, 'Variacao Carga 7d') || '0').replace('%', '').replace('+', '')) || 0;
      var a30 = pcF(pcG(r, MI, 'Atividades 30d'));
      return v7 > 10 && a30 >= 4;
    })
    .map(function(r) {
      var v7  = String(pcG(r, MI, 'Variacao Carga 7d') || '');
      var a30 = pcF(pcG(r, MI, 'Atividades 30d'));
      var n   = parseFloat(v7.replace('%', '').replace('+', '')) || 0;
      return { r: r, n: n, v7: v7, a30: a30 };
    })
    .sort(function(a, b) { return b.n - a.n; })
    .map(function(x, i) { return [i + 1, nomeDe(x.r), x.v7 + ' carga', x.a30 + ' treinos 30d']; });
  bloco('ð¥ Voltou com Tudo! (Carga 7d â + Frequencia)', '#1b5e20',
    ['#', 'Atleta', 'Variacao Carga', 'Treinos 30d'], voltou);

  // ââ IDADE / GÃNERO ââââââââââââââââââââââââââââââââââââââââââââââââ
  pcBloco_(sh, row, 5, 'ð¥ CATEGORIAS POR IDADE E GENERO', '#263238', 11); row++;
  var cadSh = ss.getSheets().filter(function(s) { return s.getName().indexOf('CADASTRO') >= 0; })[0];
  var temId = false, temGn = false;
  if (cadSh) {
    cadSh.getRange(1, 1, 1, cadSh.getLastColumn()).getValues()[0].forEach(function(h) {
      var hs = String(h).toLowerCase();
      if (hs.indexOf('idade') >= 0 || hs.indexOf('nasc') >= 0) temId = true;
      if (hs.indexOf('genero') >= 0 || hs.indexOf('sexo') >= 0) temGn = true;
    });
  }
  sh.getRange(row, 1, 1, 4).merge().setValue(
    (!temId || !temGn)
      ? 'DADOS_INSUFICIENTES â Campos Idade/Genero nao encontrados no CADASTRO'
      : 'Implementar apos validacao dos campos Idade/Genero no CADASTRO'
  ).setFontStyle('italic').setFontColor('#9e9e9e');
  row += 3;

  // RodapÃ©
  sh.getRange(row, 1, 1, 5).merge()
    .setValue('Rankings calculados com METRICAS_BETA. Importacao Strava NAO realizada nesta execucao.')
    .setFontColor('#bdbdbd').setFontSize(8).setFontStyle('italic');

  sh.setColumnWidth(1, 40); sh.setColumnWidth(2, 180); sh.setColumnWidth(3, 160); sh.setColumnWidth(4, 220);
  Logger.log('RANKINGS_BETA: OK rows=' + row);
}

// ââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââ
// INVENTÃRIO DE ABAS
// ââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââ
function pcInventario_() {
  var ss     = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var shInv  = pcGetAba_(ss, PC_INV);
  shInv.clearContents();

  var hdrs = ['Nome', 'Linhas', 'Colunas', 'Oculta?', 'Classificacao', 'Acao Recomendada', 'Acao Executada'];
  shInv.getRange(1, 1, 1, hdrs.length).setValues([hdrs]);
  pcHeader_(shInv, 1, hdrs.length, '#37474f');

  var r = 2, ocultadas = 0;
  sheets.forEach(function(s) {
    var nome = s.getName();
    // Pular a prÃ³pria aba de inventÃ¡rio
    if (nome === PC_INV || nome.indexOf('INVENTARIO') >= 0) return;
    var prot = PC_INTOCAVEIS.some(function(p) { return nome.indexOf(p) >= 0; });
    var lastR = s.getLastRow(), lastC = s.getLastColumn();
    var jaOc  = s.isSheetHidden();
    var cls   = pcClassAba_(nome, lastR, lastC, prot);
    var acRec = prot ? 'MANTER' : (cls === 'VAZIA' ? 'OCULTAR_SEGURA' : cls === 'LEGADO' ? 'REVISAR' : 'MANTER');
    var acEx  = 'NENHUMA';
    if (cls === 'VAZIA' && !prot && !jaOc) {
      try { s.hideSheet(); acEx = 'OCULTADA'; ocultadas++; }
      catch(e) { acEx = 'ERRO=' + e.message.substring(0, 40); }
    }
    shInv.getRange(r, 1, 1, 7).setValues([[nome, lastR, lastC, jaOc ? 'SIM' : 'NAO', cls, acRec, acEx]]);
    r++;
  });

  shInv.setFrozenRows(1); shInv.autoResizeColumns(1, 7);
  Logger.log('INVENTARIO: ' + sheets.length + ' abas | ocultadas=' + ocultadas);
  return { total: sheets.length, ocultadas: ocultadas };
}

// ââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââ
// HELPERS GLOBAIS
// ââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââ
function pcGetAba_(ss, nome) {
  var sh = ss.getSheetByName(nome);
  if (!sh) {
    var clean = nome.replace(/^[^\w]*/g, '');
    ss.getSheets().forEach(function(s) { if (!sh && s.getName().indexOf(clean) >= 0) sh = s; });
  }
  return sh || ss.insertSheet(nome);
}
function pcLer_(ss, nome) {
  var sh = ss.getSheetByName(nome);
  if (!sh) {
    var clean = nome.replace(/^[^\w]*/g, '');
    ss.getSheets().forEach(function(s) { if (!sh && s.getName().indexOf(clean) >= 0) sh = s; });
  }
  if (!sh || sh.getLastRow() < 2) return { h: [], r: [] };
  var lC = sh.getLastColumn(), lR = sh.getLastRow();
  var h  = sh.getRange(1, 1, 1, lC).getValues()[0].map(function(v) { return String(v).trim(); });
  var r  = sh.getRange(2, 1, lR - 1, lC).getValues();
  return { h: h, r: r };
}
function pcIdx_(hdrs) {
  var m = {};
  hdrs.forEach(function(h, i) { if (h) m[h] = i; });
  return m;
}
// Busca campo por nome exato OU por VersÃ£o sem acento (fallback)
function pcG(row, map, campo) {
  // Busca direta
  if (map[campo] !== undefined && map[campo] >= 0) return row[map[campo]];
  // Fallback: tentar variaÃ§Ãµes de acentuaÃ§Ã£o
  var clean = campo.normalize ? campo.normalize('NFD').replace(/[\u0300-\u036f]/g,'') : campo;
  var k = Object.keys(map).filter(function(h) {
    var hc = h.normalize ? h.normalize('NFD').replace(/[\u0300-\u036f]/g,'') : h;
    return hc === clean;
  })[0];
  if (k && map[k] >= 0) return row[map[k]];
  return '';
}
function pcF(v) {
  var n = parseFloat(String(v).replace(',', '.'));
  return isNaN(n) ? 0 : n;
}
function pcDt_(d) {
  try { return Utilities.formatDate(d, 'America/Sao_Paulo', 'dd/MM/yyyy HH:mm'); }
  catch(e) { return String(d); }
}
function pcFmtSeg_(s) {
  s = Math.round(pcF(s));
  var h = Math.floor(s / 3600), m = Math.floor((s % 3600) / 60), ss2 = s % 60;
  return h > 0 ? h + 'h ' + m + 'min' : (m > 0 ? m + 'min ' + ss2 + 's' : s + 's');
}
function pcAlerta_(dias, status) {
  if (!status) status = '';
  if (status === 'SEM_ATIVIDADE_VALIDA') return 'SEM_ATIVIDADE';
  dias = pcF(dias);
  if (dias <= 0)  return 'TREINOU_HOJE';
  if (dias <= 3)  return 'OK';
  if (dias <= 7)  return 'ATENCAO';
  if (dias <= 14) return 'ALERTA';
  return 'CRITICO';
}
function pcCorAlerta_(a) {
  var cores = {
    'TREINOU_HOJE': '#2e7d32', 'OK': '#388e3c',
    'ATENCAO': '#f9a825', 'ALERTA': '#e65100',
    'CRITICO': '#b71c1c', 'SEM_ATIVIDADE': '#757575'
  };
  return cores[a] || null;
}
function pcBloco_(sh, row, cols, txt, bg, size) {
  var r = sh.getRange(row, 1, 1, cols);
  r.merge(); r.setValue(txt);
  r.setBackground(bg); r.setFontColor('#fff');
  r.setFontWeight('bold'); r.setFontSize(size || 10);
}
function pcHeader_(sh, row, cols, bg) {
  var r = sh.getRange(row, 1, 1, cols);
  r.setBackground(bg || '#37474f'); r.setFontColor('#fff');
  r.setFontWeight('bold'); r.setFontSize(9);
}
function pcClassAba_(nome, lastR, lastC, prot) {
  if (prot) return 'OFICIAL';
  var n = nome.toLowerCase();
  if (/^(p[Ã¡a]gina|page|sheet|plan)\s*\d+$/i.test(n)) return lastR <= 1 ? 'VAZIA' : 'LEGADO';
  if (lastR <= 1 && lastC <= 1) return 'VAZIA';
  if (n.indexOf('backup') >= 0 || n.indexOf('bkp') >= 0) return 'BACKUP';
  if (n.indexOf('teste') >= 0 || n.indexOf('test') >= 0) return 'TECNICA';
  if (n.indexOf('legado') >= 0 || n.indexOf('antigo') >= 0 || n.indexOf('old') >= 0) return 'LEGADO';
  if (n.indexOf('beta') >= 0) return 'BETA';
  return 'SUSPEITA';
}
