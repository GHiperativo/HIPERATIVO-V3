/**
 * ═══════════════════════════════════════════════════════════════════════
 * HIPERATIVO V3 — 09_Intervals.gs
 * Integração com Intervals.icu API
 * CTL/ATL/TSB + diagnóstico algorítmico (módulos A/B/C/D)
 * ═══════════════════════════════════════════════════════════════════════
 */

var INTERVALS_API_BASE = 'https://intervals.icu/api/v1';

// ── CONFIGURAÇÃO ──────────────────────────────────────────────────────────────

/**
 * Salva as credenciais do admin no PropertiesService.
 * Execute uma única vez pelo menu após o deploy.
 */
function configurarIntervalsAdmin() {
  var props = PropertiesService.getScriptProperties();
  props.setProperties({
    'INTERVALS_ADMIN_ID':      'i610008',
    'INTERVALS_ADMIN_API_KEY': '6tbjqfhsy8a8i9tyhiqc8bpyt'
  });
  Logger.log('Credenciais Intervals.icu salvas no PropertiesService');
  try {
    SpreadsheetApp.getUi().alert(
      '✅ Intervals.icu configurado',
      'ID: i610008\nUse "Atualizar CTL/ATL/TSB" para buscar os dados.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch(e) {
    // Sem contexto de UI (ex: rodando pelo editor) — credenciais já foram salvas
    Logger.log('✅ Configurado com sucesso. ID: i610008');
  }
}

// ── CLIENTE API ───────────────────────────────────────────────────────────────

/**
 * Busca dados de fitness (CTL/ATL/TSB/VO2max) de um atleta no Intervals.icu
 * @param {string} intervalId  ex: "i610008"
 * @param {string} apiKey
 * @returns {Object|null}  { ctl, atl, tsb, rampRate, vo2max }
 */
function buscarDadosIntervalsAtleta(intervalId, apiKey) {
  var hoje = Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'yyyy-MM-dd');
  var url = INTERVALS_API_BASE + '/athlete/' + intervalId +
            '/fitness?oldest=' + hoje + '&newest=' + hoje;

  var credenciais = Utilities.base64Encode('API_KEY:' + apiKey);
  var options = {
    method: 'GET',
    headers: {
      'Authorization': 'Basic ' + credenciais,
      'Accept': 'application/json'
    },
    muteHttpExceptions: true
  };

  try {
    var resp = UrlFetchApp.fetch(url, options);
    if (resp.getResponseCode() !== 200) {
      Logger.log('Intervals fitness HTTP ' + resp.getResponseCode() + ' para ' + intervalId);
      return null;
    }
    var dados = JSON.parse(resp.getContentText());
    if (!dados || dados.length === 0) return null;

    var ultimo = dados[dados.length - 1];
    return {
      ctl:      ultimo.ctl      || 0,
      atl:      ultimo.atl      || 0,
      tsb:      ultimo.tsb      || 0,
      rampRate: ultimo.rampRate || 0,
      vo2max:   ultimo.vo2max   || null
    };
  } catch (e) {
    Logger.log('buscarDadosIntervalsAtleta: ' + e.message);
    return null;
  }
}

/**
 * Busca distribuição de zonas de FC dos últimos 30 dias.
 * @returns {Object|null}  { z1_pct…z5_pct, totalMinutos }
 */
function buscarZonasIntervalsAtleta(intervalId, apiKey) {
  var hoje    = new Date();
  var ha30d   = new Date(hoje.getTime() - 30 * 24 * 60 * 60 * 1000);
  var oldest  = Utilities.formatDate(ha30d, 'America/Sao_Paulo', 'yyyy-MM-dd');
  var newest  = Utilities.formatDate(hoje,  'America/Sao_Paulo', 'yyyy-MM-dd');

  var url = INTERVALS_API_BASE + '/athlete/' + intervalId +
            '/activities?oldest=' + oldest + '&newest=' + newest;

  var credenciais = Utilities.base64Encode('API_KEY:' + apiKey);
  var options = {
    method: 'GET',
    headers: { 'Authorization': 'Basic ' + credenciais },
    muteHttpExceptions: true
  };

  try {
    var resp = UrlFetchApp.fetch(url, options);
    if (resp.getResponseCode() !== 200) return null;

    var atividades = JSON.parse(resp.getContentText());
    if (!atividades || atividades.length === 0) return null;

    var zonas = { z1: 0, z2: 0, z3: 0, z4: 0, z5: 0, total: 0 };
    atividades.forEach(function(a) {
      if (a.heartRateZones) {
        a.heartRateZones.forEach(function(z, i) {
          var key = 'z' + (i + 1);
          if (zonas[key] !== undefined) zonas[key] += (z.secs || 0);
          zonas.total += (z.secs || 0);
        });
      }
    });

    if (zonas.total === 0) return null;
    return {
      z1_pct: Math.round(zonas.z1 / zonas.total * 100),
      z2_pct: Math.round(zonas.z2 / zonas.total * 100),
      z3_pct: Math.round(zonas.z3 / zonas.total * 100),
      z4_pct: Math.round(zonas.z4 / zonas.total * 100),
      z5_pct: Math.round(zonas.z5 / zonas.total * 100),
      totalMinutos: Math.round(zonas.total / 60)
    };
  } catch (e) {
    Logger.log('buscarZonasIntervalsAtleta: ' + e.message);
    return null;
  }
}

// ── MÓDULOS DE DIAGNÓSTICO ────────────────────────────────────────────────────

/** Módulo A — Carga de treino (TSB + Ramp Rate) */
function _diagnosticoCarga(fitness) {
  if (!fitness) return '⚠️ Dados de carga não disponíveis';

  var tsb = fitness.tsb;
  var rr  = fitness.rampRate;
  var linhas = [];

  if (tsb < -30) {
    linhas.push('🔴 SOBRECARGA ALTA — TSB ' + tsb.toFixed(1));
    linhas.push('Risco elevado de overtraining ou lesão. Reduzir volume/intensidade.');
  } else if (tsb < -10) {
    linhas.push('🟡 EM CARGA — TSB ' + tsb.toFixed(1));
    linhas.push('Período produtivo de adaptação. Manter consistência.');
  } else if (tsb <= 10) {
    linhas.push('🟢 NEUTRO — TSB ' + tsb.toFixed(1));
    linhas.push('Equilíbrio fadiga/fitness. Bom para treinos de qualidade.');
  } else {
    linhas.push('✅ FRESCO — TSB ' + tsb.toFixed(1));
    linhas.push('Pronto para competição ou carga mais intensa.');
  }

  if (rr > 10) {
    linhas.push('⚠️ Ramp Rate ' + rr.toFixed(1) + '/sem — excede limite de 10. Risco de lesão por sobrecarga aguda.');
  }

  return linhas.join('\n');
}

/** Módulo B — Distribuição de esforço (zonas de FC) */
function _diagnosticoZonas(zonas) {
  if (!zonas) return '⚠️ Dados de zonas não disponíveis (atividades sem FC nos últimos 30d)';

  var z12  = zonas.z1_pct + zonas.z2_pct;
  var z45  = zonas.z4_pct + zonas.z5_pct;
  var linhas = [
    'Z1+Z2=' + z12 + '% | Z3=' + zonas.z3_pct + '% | Z4+Z5=' + z45 + '%',
    '(Total com FC: ' + zonas.totalMinutos + ' min)'
  ];

  if (z12 < 70) {
    linhas.push('⚠️ Base aeróbica insuficiente (Z1+Z2=' + z12 + '%, ideal 70-85%). Incluir treinos longos em baixa intensidade.');
  } else if (z12 > 85) {
    linhas.push('🟡 Treino excessivamente conservador (Z1+Z2=' + z12 + '%). Considerar intervalados.');
  } else {
    linhas.push('✅ Distribuição aeróbica adequada (Z1+Z2=' + z12 + '%)');
  }

  if (z45 < 10) {
    linhas.push('🟡 Pouco estímulo de alta intensidade (Z4+Z5=' + z45 + '%, ideal 10-20%)');
  } else if (z45 > 25) {
    linhas.push('⚠️ Excesso de alta intensidade (Z4+Z5=' + z45 + '%). Risco de overtraining.');
  }

  return linhas.join('\n');
}

/** Módulo C — Consistência (últimos 30 dias, via API) */
function _diagnosticoConsistencia(intervalId, apiKey) {
  var hoje   = new Date();
  var ha30d  = new Date(hoje.getTime() - 30 * 24 * 60 * 60 * 1000);
  var oldest = Utilities.formatDate(ha30d, 'America/Sao_Paulo', 'yyyy-MM-dd');
  var newest = Utilities.formatDate(hoje,  'America/Sao_Paulo', 'yyyy-MM-dd');

  var url = INTERVALS_API_BASE + '/athlete/' + intervalId +
            '/activities?oldest=' + oldest + '&newest=' + newest;
  var credenciais = Utilities.base64Encode('API_KEY:' + apiKey);

  try {
    var resp = UrlFetchApp.fetch(url, {
      method: 'GET',
      headers: { 'Authorization': 'Basic ' + credenciais },
      muteHttpExceptions: true
    });
    if (resp.getResponseCode() !== 200) return '⚠️ Não foi possível calcular consistência';

    var atividades = JSON.parse(resp.getContentText()) || [];
    var total = atividades.length;

    var semanasAtivas = {};
    atividades.forEach(function(a) {
      var d = new Date(a.startDateLocal || a.start_date_local);
      var semana = Math.floor((hoje - d) / (7 * 24 * 60 * 60 * 1000));
      semanasAtivas[semana] = true;
    });
    var nSemanas = Object.keys(semanasAtivas).length;
    var mediaSem = (total / 4.3).toFixed(1);

    var linhas = [
      'Atividades 30d: ' + total + ' (média ' + mediaSem + '/sem)',
      'Semanas ativas: ' + nSemanas + '/4'
    ];

    if (nSemanas >= 4) {
      linhas.push('✅ Consistência excelente');
    } else if (nSemanas >= 3) {
      linhas.push('🟡 Consistência boa (' + nSemanas + '/4 semanas)');
    } else {
      linhas.push('🔴 Baixa consistência (' + nSemanas + '/4 semanas) — priorizar regularidade antes de aumentar intensidade');
    }
    return linhas.join('\n');
  } catch (e) {
    return '⚠️ Erro ao calcular consistência: ' + e.message;
  }
}

/** Módulo D — Flags de saúde (lidas do CADASTRO) */
function _diagnosticoSaude(atletaData) {
  var flags = [];
  if (atletaData.dm1)          flags.push('🩺 DM1: monitorar glicemia antes/durante/após treinos; gel de emergência.');
  if (atletaData.asma)         flags.push('🫁 Asma: verificar condições climáticas; broncodilatador acessível no treino.');
  if (atletaData.lombar)       flags.push('🦴 Lombar: aquecimento obrigatório; fortalecer core; evitar impacto em dias de dor.');
  if (atletaData.fibromialgia) flags.push('⚡ Fibromialgia: respeitar janelas de recuperação; intensidade moderada; evitar picos.');

  return flags.length === 0
    ? '✅ Sem flags de saúde ativas'
    : '⚠️ FLAGS DE SAÚDE:\n' + flags.join('\n');
}

// ── DIAGNÓSTICO COMPLETO ──────────────────────────────────────────────────────

/**
 * Gera relatório diagnóstico completo para um atleta (A+B+C+D).
 * @param {string} atletaId  ID interno do atleta (col ID_ATLETA do CADASTRO)
 * @returns {string}  Texto do diagnóstico
 */
function gerarDiagnosticoAtleta(atletaId) {
  var props    = PropertiesService.getScriptProperties();
  var adminKey = props.getProperty('INTERVALS_ADMIN_API_KEY');
  if (!adminKey) return 'ERRO: configure INTERVALS_ADMIN_API_KEY via menu Intervals.icu → Configurar credenciais';

  var ss       = SpreadsheetApp.getActive();
  var sheetCad = ss.getSheetByName(H.SHEETS.CADASTRO);
  var dados    = sheetCad.getDataRange().getValues();
  var hdrs     = dados[0];

  var colId    = hdrs.indexOf('ID_ATLETA');
  var colIid   = hdrs.indexOf('INTERVALS_ID');
  var colSaude = hdrs.indexOf('SAUDE');  // coluna de texto livre de saúde

  var row = null;
  for (var i = 1; i < dados.length; i++) {
    if (String(dados[i][colId]) === String(atletaId)) { row = dados[i]; break; }
  }
  if (!row) return 'Atleta ' + atletaId + ' não encontrado no CADASTRO';

  var intervalId = colIid >= 0 ? row[colIid] : null;
  if (!intervalId) return 'Atleta ' + atletaId + ' não tem INTERVALS_ID configurado no CADASTRO';

  var fitness = buscarDadosIntervalsAtleta(intervalId, adminKey);
  var zonas   = buscarZonasIntervalsAtleta(intervalId, adminKey);

  // Flags de saúde extraídas do campo SAUDE (texto livre) ou de colunas dedicadas
  var saudeTexto = colSaude >= 0 ? String(row[colSaude] || '').toLowerCase() : '';
  var atletaData = {
    dm1:          saudeTexto.indexOf('dm1')          >= 0 || saudeTexto.indexOf('diabetes') >= 0,
    asma:         saudeTexto.indexOf('asma')         >= 0,
    lombar:       saudeTexto.indexOf('lombar')       >= 0,
    fibromialgia: saudeTexto.indexOf('fibromialgia') >= 0
  };

  var data = Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'dd/MM/yyyy HH:mm');

  return [
    '══════════ DIAGNÓSTICO ' + data + ' ══════════',
    '',
    '📊 A. CARGA DE TREINO',
    _diagnosticoCarga(fitness),
    '',
    '🎯 B. DISTRIBUIÇÃO DE ESFORÇO (30d)',
    _diagnosticoZonas(zonas),
    '',
    '📅 C. CONSISTÊNCIA',
    _diagnosticoConsistencia(intervalId, adminKey),
    '',
    '🏥 D. SAÚDE',
    _diagnosticoSaude(atletaData)
  ].join('\n');
}

// ── ATUALIZAÇÃO EM LOTE ───────────────────────────────────────────────────────

/**
 * Atualiza CTL/ATL/TSB/Ramp_Rate/Fonte_VO2max em MÉTRICAS
 * para todos os atletas com INTERVALS_ID no CADASTRO.
 */
function atualizarMetricasIntervals() {
  var props    = PropertiesService.getScriptProperties();
  var adminKey = props.getProperty('INTERVALS_ADMIN_API_KEY');

  if (!adminKey) {
    SpreadsheetApp.getUi().alert(
      '❌ Não configurado',
      'Execute primeiro: Menu → Intervals.icu → Configurar credenciais',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }

  var ss       = SpreadsheetApp.getActive();
  var sheetCad = ss.getSheetByName(H.SHEETS.CADASTRO);
  var sheetMet = ss.getSheetByName(H.SHEETS.METRICAS);

  if (!sheetCad || !sheetMet) {
    Logger.log('CADASTRO ou MÉTRICAS não encontradas');
    return;
  }

  // Garantir colunas CTL–Fonte_VO2max no cabeçalho de MÉTRICAS
  _garantirColunasMetricas(sheetMet);

  var dadosCad = sheetCad.getDataRange().getValues();
  var hdrsCad  = dadosCad[0];
  var colIdCad = hdrsCad.indexOf('ID_ATLETA');
  var colIid   = hdrsCad.indexOf('INTERVALS_ID');

  if (colIid < 0) {
    SpreadsheetApp.getUi().alert(
      '❌ Coluna INTERVALS_ID ausente',
      'Adicione a coluna INTERVALS_ID no CADASTRO e preencha os IDs dos atletas.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }

  var dadosMet = sheetMet.getDataRange().getValues();
  var hdrsMet  = dadosMet[0];
  var colIdMet = hdrsMet.indexOf('ID_ATLETA');
  var colCTL   = hdrsMet.indexOf('CTL');
  var colATL   = hdrsMet.indexOf('ATL');
  var colTSB   = hdrsMet.indexOf('TSB');
  var colRamp  = hdrsMet.indexOf('Ramp_Rate');
  var colFonte = hdrsMet.indexOf('Fonte_VO2max');

  var atualizados = 0;
  var erros = 0;

  for (var i = 1; i < dadosCad.length; i++) {
    var atletaId  = dadosCad[i][colIdCad];
    var intervalId = dadosCad[i][colIid];

    if (!atletaId || !intervalId) continue;

    var fitness = buscarDadosIntervalsAtleta(String(intervalId), adminKey);
    if (!fitness) { erros++; continue; }

    // Localizar atleta em MÉTRICAS
    for (var j = 1; j < dadosMet.length; j++) {
      if (String(dadosMet[j][colIdMet]) === String(atletaId)) {
        var linha = j + 1;
        if (colCTL  >= 0) sheetMet.getRange(linha, colCTL  + 1).setValue(fitness.ctl);
        if (colATL  >= 0) sheetMet.getRange(linha, colATL  + 1).setValue(fitness.atl);
        if (colTSB  >= 0) sheetMet.getRange(linha, colTSB  + 1).setValue(fitness.tsb);
        if (colRamp >= 0) sheetMet.getRange(linha, colRamp + 1).setValue(fitness.rampRate);
        if (colFonte >= 0) sheetMet.getRange(linha, colFonte + 1).setValue('Intervals.icu');
        atualizados++;
        break;
      }
    }

    Utilities.sleep(300); // respeitar rate limit
  }

  var msg = 'Intervals.icu: ' + atualizados + ' atleta(s) atualizados com CTL/ATL/TSB';
  if (erros > 0) msg += '\n(' + erros + ' sem dados disponíveis)';

  Logger.log(msg);
  SpreadsheetApp.getUi().alert('✅ Atualização concluída', msg, SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Garante que as 5 colunas de métricas Intervals existam no cabeçalho de MÉTRICAS.
 */
function _garantirColunasMetricas(sheet) {
  var hdr      = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var novas    = ['CTL', 'ATL', 'TSB', 'Ramp_Rate', 'Fonte_VO2max'];
  var ultimaCol = hdr.length;

  novas.forEach(function(col) {
    if (hdr.indexOf(col) < 0) {
      ultimaCol++;
      sheet.getRange(1, ultimaCol).setValue(col);
      hdr.push(col);
    }
  });
}
