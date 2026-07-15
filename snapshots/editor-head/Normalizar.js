// ============================================================
// Normalizar.gs — HIPERATIVO V3
// Padrão oficial de conversão e normalização de dados Strava
// Criado: 21/06/2026
//
// REGRAS:
// - Distância: metros → km (2 casas)
// - Tempo: segundos → HH:MM:SS
// - Velocidade: m/s → km/h (2 casas) ou pace MM:SS min/km
// - FC / Potência / Elevação: inteiros
// - Tipo: traduzido para português (sport_type como primário)
// - Rota/polyline: NÃO incluída no objeto normalizado
// - Ordenação: mais recente → mais antiga (responsabilidade do caller)
// ============================================================

// — CABEÇALHO OFICIAL — 34 COLUNAS ——————————————————————

const ATIVIDADES_NORMALIZADAS_HEADERS = [
  'ID Interno',
  'ATH_ID',
  'Atleta',
  'Data/Hora',
  'Mês',
  'Tipo',
  'Tipo Original',
  'Strava ID',
  'Nome da Atividade',
  'Distância km',
  'Distância',
  'Tempo Movimento s',
  'Tempo Movimento',
  'Tempo Total s',
  'Tempo Total',
  'Pace s/km',
  'Pace',
  'Velocidade km/h',
  'Velocidade',
  'FC Média',
  'FC Média fmt',
  'FC Máx.',
  'FC Máx. fmt',
  'Elevação m',
  'Elevação',
  'Calorias',
  'Calorias fmt',
  'Cadência',
  'Cadência fmt',
  'Potência W',
  'Potência',
  'Fonte',
  'Importado em',
  'Status'
];

// — DICIONÁRIO DE TIPOS ——————————————————————————————

const SPORT_TYPE_DICT = {
  Run: 'Corrida',
  TrailRun: 'Corrida em trilha',
  Walk: 'Caminhada',
  Ride: 'Ciclismo',
  VirtualRide: 'Ciclismo virtual',
  MountainBikeRide: 'Mountain bike',
  GravelRide: 'Gravel',
  Swim: 'Natação',
  Workout: 'Treino',
  WeightTraining: 'Musculação',
  HighIntensityIntervalTraining: 'HIIT',
  Yoga: 'Yoga',
  Pilates: 'Pilates',
  Hike: 'Trilha',
  Elliptical: 'Elíptico',
  StairStepper: 'Escada',
  Rowing: 'Remo',
  VirtualRow: 'Remo virtual',
  Soccer: 'Futebol',
  Tennis: 'Tênis',
  Padel: 'Padel',
  Pickleball: 'Pickleball',
  Badminton: 'Badminton',
  Dance: 'Dança',
  PhysicalTherapy: 'Fisioterapia',
  Crossfit: 'Crossfit',
  InlineSkate: 'Patins',
  Wheelchair: 'Cadeira de rodas'
};

const TIPOS_CORRIDA = ['Run', 'TrailRun', 'Walk', 'Hike'];
const TIPOS_CICLO   = ['Ride', 'VirtualRide', 'MountainBikeRide', 'GravelRide'];

// — CONVERSORES BÁSICOS ——————————————————————————————

function toKm(metros) {
    if (!metros && metros !== 0) return null;
      return Math.round(Number(metros)) / 1000;
}

function toTempo(segundos) {
  if (segundos === null || segundos === undefined || segundos === '') return null;
  var s = Math.round(Number(segundos));
  if (s < 0) return null;
  var h   = Math.floor(s / 3600);
  var m   = Math.floor((s % 3600) / 60);
  var sec = s % 60;
  var mm  = String(m).padStart(2, '0');
  var ss  = String(sec).padStart(2, '0');
  return String(h).padStart(2, '0') + ':' + mm + ':' + ss;
}

function toPaceS(avgSpeedMs) {
  if (!avgSpeedMs || Number(avgSpeedMs) <= 0) return null;
  return Math.round(1000 / Number(avgSpeedMs));
}

function toPaceMinKm(avgSpeedMs) {
  if (!avgSpeedMs || Number(avgSpeedMs) <= 0) return null;
  var paceDecimal = 1000 / (Number(avgSpeedMs) * 60);
  var min = Math.floor(paceDecimal);
  var sec = Math.round((paceDecimal - min) * 60);
  if (sec === 60) { min += 1; sec = 0; }
  return min + ':' + String(sec).padStart(2, '0');
}

function toVelKmh(avgSpeedMs) {
  if (avgSpeedMs === null || avgSpeedMs === undefined || Number(avgSpeedMs) <= 0) return null;
  return Math.round(Number(avgSpeedMs) * 3.6 * 100) / 100;
}

function traduzirTipo(sportType, typeFallback) {
  var chave = sportType || typeFallback || '';
  return SPORT_TYPE_DICT[chave] || 'Outro';
}

// — NORMALIZAÇÃO PRINCIPAL ———————————————————————————

function normalizarAtividadeStrava(activity, atleta) {
  var ath_id   = atleta.ath_id || atleta.ATH_ID || '';
  var nome_atl = atleta.nome   || atleta.Nome   || '';

  var tipo_original = activity.sport_type || activity.type || '';
  var tipo          = traduzirTipo(activity.sport_type, activity.type);

  var ehCorrida = TIPOS_CORRIDA.includes(tipo_original);
  var ehCiclo   = TIPOS_CICLO.includes(tipo_original);

  var dist_km  = toKm(activity.distance) ?? 0;
  var dist_fmt = dist_km !== null
    ? String(dist_km.toFixed(2)).replace('.', ',') + ' km'
    : '';

  var tempo_mov_s     = activity.moving_time  != null ? Math.round(Number(activity.moving_time))  : null;
  var tempo_mov_fmt   = toTempo(tempo_mov_s)   || '';
  var tempo_total_s   = activity.elapsed_time != null ? Math.round(Number(activity.elapsed_time)) : null;
  var tempo_total_fmt = toTempo(tempo_total_s) || '';

  var pace_s_km = null;
  var pace_fmt  = '';
  if (ehCorrida && activity.average_speed && Number(activity.average_speed) > 0) {
    pace_s_km = toPaceS(activity.average_speed);
    pace_fmt  = toPaceMinKm(activity.average_speed) + ' min/km';
  }

  var vel_kmh     = null;
  var vel_kmh_fmt = '';
  if (activity.average_speed && Number(activity.average_speed) > 0) {
    vel_kmh     = toVelKmh(activity.average_speed);
    vel_kmh_fmt = String(vel_kmh.toFixed(2)).replace('.', ',') + ' km/h';
  }

  var fc_media     = activity.average_heartrate != null ? Math.round(Number(activity.average_heartrate)) : null;
  var fc_media_fmt = fc_media !== null ? fc_media + ' bpm' : '';
  var fc_max       = activity.max_heartrate     != null ? Math.round(Number(activity.max_heartrate))     : null;
  var fc_max_fmt   = fc_max   !== null ? fc_max   + ' bpm' : '';

  var elev_m   = activity.total_elevation_gain != null ? Math.round(Number(activity.total_elevation_gain)) : null;
  var elev_fmt = elev_m !== null ? elev_m + ' m' : '';

  var calorias     = activity.calories != null ? Math.round(Number(activity.calories)) : null;
  var calorias_fmt = calorias !== null ? calorias + ' kcal' : '';

  var cadencia     = null;
  var cadencia_fmt = '';
  if (activity.average_cadence != null) {
    cadencia     = Math.round(Number(activity.average_cadence));
    cadencia_fmt = cadencia + (ehCiclo ? ' rpm' : ' spm');
  }

  var potencia_w   = activity.average_watts != null ? Math.round(Number(activity.average_watts)) : null;
  var potencia_fmt = potencia_w !== null ? potencia_w + ' W' : '';

  var data_hora = null;
  if (activity.start_date_local) {
    data_hora = new Date(activity.start_date_local);
  } else if (activity.start_date) {
    data_hora = new Date(activity.start_date);
  }

  var mes_ref = '';
  if (data_hora) {
    var mo = data_hora.getMonth() + 1;
    mes_ref = data_hora.getFullYear() + '-' + String(mo).padStart(2, '0');
  }

  var strava_id  = String(activity.id || '');
  var id_interno = ath_id + '_' + strava_id;

  return {
    id_interno,
    ath_id,
    atleta:         nome_atl,
    data_hora,
    mes_ref,
    tipo,
    tipo_original,
    strava_id,
    nome_atividade: activity.name || '',
    dist_km,
    dist_fmt,
    tempo_mov_s,
    tempo_mov_fmt,
    tempo_total_s,
    tempo_total_fmt,
    pace_s_km,
    pace_fmt,
    vel_kmh,
    vel_kmh_fmt,
    fc_media,
    fc_media_fmt,
    fc_max,
    fc_max_fmt,
    elev_m,
    elev_fmt,
    calorias,
    calorias_fmt,
    cadencia,
    cadencia_fmt,
    potencia_w,
    potencia_fmt,
    fonte:        'Strava',
    importado_em: new Date(),
    status:       'Importado'
  };
}

// — CONVERSÃO PARA LINHA (34 COLUNAS) ————————————————

function normalizadoParaLinha(norm) {
  return [
    norm.id_interno      || '',
    norm.ath_id          || '',
    norm.atleta          || '',
    norm.data_hora       || '',
    norm.mes_ref         || '',
    norm.tipo            || '',
    norm.tipo_original   || '',
    norm.strava_id       || '',
    norm.nome_atividade  || '',
    norm.dist_km         ?? '',
    norm.dist_fmt        || '',
    norm.tempo_mov_s     ?? '',
    norm.tempo_mov_fmt   || '',
    norm.tempo_total_s   ?? '',
    norm.tempo_total_fmt || '',
    norm.pace_s_km       ?? '',
    norm.pace_fmt        || '',
    norm.vel_kmh         ?? '',
    norm.vel_kmh_fmt     || '',
    norm.fc_media        ?? '',
    norm.fc_media_fmt    || '',
    norm.fc_max          ?? '',
    norm.fc_max_fmt      || '',
    norm.elev_m          ?? '',
    norm.elev_fmt        || '',
    norm.calorias        ?? '',
    norm.calorias_fmt    || '',
    norm.cadencia        ?? '',
    norm.cadencia_fmt    || '',
    norm.potencia_w      ?? '',
    norm.potencia_fmt    || '',
    norm.fonte           || 'Strava',
    norm.importado_em    || new Date(),
    norm.status          || 'Importado'
  ];
}

// — PREPARAR ABA ATIVIDADES ——————————————————————————

function prepararAbaAtividadesNormalizada() {
  const ss = SpreadsheetApp.openById('1bI5pnt-HOAD5p8M2hqjEsU9P816hc94wy4mqx0J_xOM');
  const sh = ss.getSheetByName('🏃 ATIVIDADES');

  if (!sh) throw new Error('Aba 🏃 ATIVIDADES não encontrada.');

  sh.clear();

  sh.getRange(1, 1, 1, ATIVIDADES_NORMALIZADAS_HEADERS.length)
    .merge()
    .setValue('🏃 ATIVIDADES — DADOS IMPORTADOS DO STRAVA')
    .setHorizontalAlignment('center')
    .setFontWeight('bold')
    .setFontColor('#FFFFFF')
    .setBackground('#C33500');

  sh.getRange(2, 1, 1, ATIVIDADES_NORMALIZADAS_HEADERS.length)
    .setValues([ATIVIDADES_NORMALIZADAS_HEADERS])
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setFontColor('#FFFFFF')
    .setBackground('#003366');

  sh.setFrozenRows(2);
  sh.autoResizeColumns(1, ATIVIDADES_NORMALIZADAS_HEADERS.length);

  return {
    ok: true,
    aba: '🏃 ATIVIDADES',
    colunas: ATIVIDADES_NORMALIZADAS_HEADERS.length
  };
}

// — TESTES ——————————————————————————————————————————

function testarNormalizacaoStrava() {
  var resultados = [];
  var todoOk = true;

  // TESTE 1: Corrida
  var n1 = normalizarAtividadeStrava({
    id: '9999001', name: 'Corrida matinal',
    sport_type: 'Run', type: 'Run',
    distance: 5680, moving_time: 2061, elapsed_time: 2100,
    average_speed: 2.757, average_heartrate: 145, max_heartrate: 168,
    total_elevation_gain: 32, calories: 320, average_cadence: 172,
    start_date_local: '2026-06-20T06:30:00Z'
  }, { ath_id: 'ATH992736', nome: 'Crhystiano' });
  var l1 = normalizadoParaLinha(n1);
  var ok1 = n1.tipo === 'Corrida'
    && Math.abs(n1.dist_km - 5.68) < 0.01
    && n1.dist_fmt === '5,68 km'
    && n1.tempo_mov_s === 2061
    && n1.tempo_mov_fmt === '00:34:21'
    && n1.fc_media === 145
    && n1.fc_media_fmt === '145 bpm'
    && l1.length === 34
    && !JSON.stringify(n1).includes('polyline')
    && !JSON.stringify(n1).includes('summary_polyline');
  if (!ok1) todoOk = false;
  resultados.push('[' + (ok1 ? 'OK' : 'FALHOU') + '] CORRIDA'
    + ' | tipo=' + n1.tipo
    + ' | dist_fmt=' + n1.dist_fmt
    + ' | tempo_mov_fmt=' + n1.tempo_mov_fmt
    + ' | pace_fmt=' + n1.pace_fmt
    + ' | fc_media_fmt=' + n1.fc_media_fmt
    + ' | cols=' + l1.length);

  // TESTE 2: HIIT
  var n2 = normalizarAtividadeStrava({
    id: '9999002', name: 'HIIT matinal',
    sport_type: 'HighIntensityIntervalTraining', type: 'Workout',
    distance: 0, moving_time: 1800, elapsed_time: 1800,
    average_speed: 0, average_heartrate: 152, max_heartrate: 178,
    total_elevation_gain: 0, calories: 280,
    start_date_local: '2026-06-21T07:00:00Z'
  }, { ath_id: 'ATH992736', nome: 'Crhystiano' });
  var l2 = normalizadoParaLinha(n2);
  var ok2 = n2.tipo === 'HIIT'
    && n2.dist_km === 0
    && !n2.pace_s_km
    && n2.pace_fmt === ''
    && n2.tempo_mov_fmt === '00:30:00'
    && l2.length === 34;
  if (!ok2) todoOk = false;
  resultados.push('[' + (ok2 ? 'OK' : 'FALHOU') + '] HIIT'
    + ' | tipo=' + n2.tipo
    + ' | dist_km=' + n2.dist_km
    + ' | pace_fmt="' + n2.pace_fmt + '"'
    + ' | tempo_mov_fmt=' + n2.tempo_mov_fmt
    + ' | cols=' + l2.length);

  // TESTE 3: Ciclismo
  var n3 = normalizarAtividadeStrava({
    id: '9999003', name: 'Pedalada tarde',
    sport_type: 'Ride', type: 'Ride',
    distance: 42000, moving_time: 7200, elapsed_time: 7500,
    average_speed: 5.833, average_heartrate: 138, max_heartrate: 165,
    total_elevation_gain: 210, calories: 850,
    average_cadence: 82, average_watts: 180,
    start_date_local: '2026-06-19T16:00:00Z'
  }, { ath_id: 'ATH992736', nome: 'Crhystiano' });
  var l3 = normalizadoParaLinha(n3);
  var ok3 = n3.tipo === 'Ciclismo'
    && Math.abs(n3.vel_kmh - 21.0) < 0.5
    && n3.cadencia_fmt === '82 rpm'
    && n3.potencia_fmt === '180 W'
    && l3.length === 34;
  if (!ok3) todoOk = false;
  resultados.push('[' + (ok3 ? 'OK' : 'FALHOU') + '] CICLISMO'
    + ' | tipo=' + n3.tipo
    + ' | vel_kmh=' + n3.vel_kmh
    + ' | vel_kmh_fmt=' + n3.vel_kmh_fmt
    + ' | cadencia_fmt=' + n3.cadencia_fmt
    + ' | potencia_fmt=' + n3.potencia_fmt
    + ' | cols=' + l3.length);

  // HEADERS
  var hOk = ATIVIDADES_NORMALIZADAS_HEADERS.length === 34;
  if (!hOk) todoOk = false;
  resultados.push('[' + (hOk ? 'OK' : 'FALHOU') + '] HEADERS: ' + ATIVIDADES_NORMALIZADAS_HEADERS.length + ' colunas');

  var msg = resultados.join('\n');
  Logger.log(msg);
  SpreadsheetApp.getUi().alert('TESTES NORMALIZAÇÃO\n\n' + msg + '\n\nSTATUS: ' + (todoOk ? 'TODOS OK' : 'HA FALHAS'));
  return { ok: todoOk, resultados };
}

// ═══════════════════════════════════════════════════════════════════════
// IMPORTAÇÃO SEGURA — v1.0 — 22/06/2026
// Bypass das funções antigas quebradas do menu
// NÃO toca: Supabase, Painel, Métricas, Ranking
// SÓ escreve em: 🏃 ATIVIDADES
// ═══════════════════════════════════════════════════════════════════════

/**
 * Importa as últimas 20 atividades Strava por atleta conectado.
 * Seguro: sem Supabase, sem Painel, sem Métricas, sem Ranking.
 * Chame pelo menu ou diretamente pelo editor.
 */
function importarUltimas20AtividadesPorAtleta() {
  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const wsAtiv = ss.getSheetByName(H.SHEETS.ATIVIDADES);
  const wsTok  = ss.getSheetByName(H.SHEETS.TOKENS);

  if (!wsAtiv) { Logger.log('[ERRO] Aba ATIVIDADES não encontrada'); return; }
  if (!wsTok)  { Logger.log('[ERRO] Aba TOKENS não encontrada'); return; }

  const props        = PropertiesService.getScriptProperties();
  const clientId     = props.getProperty('STRAVA_CLIENT_ID');
  const clientSecret = props.getProperty('STRAVA_CLIENT_SECRET');

  if (!clientId || !clientSecret) {
    Logger.log('[ERRO] Credenciais Strava não configuradas. Use Menu > Configurações > Configurar credenciais Strava.');
    return;
  }

  // Deduplicação: coletar Strava IDs já gravados (coluna 8 = índice 7)
  const ativVals = wsAtiv.getDataRange().getValues();
  const existingIds = new Set();
  for (let i = 2; i < ativVals.length; i++) {
    const sid = String(ativVals[i][7] || '').trim();
    if (sid && sid !== 'undefined') existingIds.add(sid);
  }

  const tokVals = wsTok.getDataRange().getValues();
  const agora   = Math.floor(Date.now() / 1000);
  const log     = [];
  let totalImportadas = 0;
  let totalErros      = 0;
  const novasLinhas   = [];

  for (let i = 1; i < tokVals.length; i++) {
    const row         = tokVals[i];
    const athId       = String(row[H.TOK.ATH_ID   - 1] || '').trim();
    const nomeAtleta  = String(row[H.TOK.NOME      - 1] || athId).trim();
    let   accessToken = String(row[H.TOK.ACCESS    - 1] || '').trim();
    const refreshTok  = String(row[H.TOK.REFRESH   - 1] || '').trim();
    const expiresRaw  = row[H.TOK.EXPIRES - 1];

    if (!athId || !accessToken) {
      log.push('[SKIP] ' + (athId || 'linha ' + (i + 1)) + ' — sem token');
      continue;
    }

    // Calcular expiração
    let expiresTs = 0;
    if (expiresRaw instanceof Date) expiresTs = Math.floor(expiresRaw.getTime() / 1000);
    else if (typeof expiresRaw === 'number' && expiresRaw > 1000000000) expiresTs = expiresRaw;

    // Renovar se expirado ou < 5 min para expirar
    if (expiresTs > 0 && agora > expiresTs - 300) {
      if (!refreshTok) {
        log.push('[SKIP] ' + athId + ' — token expirado sem refresh_token');
        continue;
      }
      try {
        const rResp = UrlFetchApp.fetch(STRAVA_TOKEN_URL, {
          method: 'post',
          payload: {
            client_id:     clientId,
            client_secret: clientSecret,
            grant_type:    'refresh_token',
            refresh_token: refreshTok
          },
          muteHttpExceptions: true
        });
        if (rResp.getResponseCode() !== 200) throw new Error('HTTP ' + rResp.getResponseCode());
        const rd = JSON.parse(rResp.getContentText());
        accessToken = rd.access_token;
        wsTok.getRange(i + 1, H.TOK.ACCESS,  1, 1).setValue(rd.access_token);
        wsTok.getRange(i + 1, H.TOK.REFRESH, 1, 1).setValue(rd.refresh_token || refreshTok);
        wsTok.getRange(i + 1, H.TOK.EXPIRES, 1, 1).setValue(new Date(rd.expires_at * 1000));
        log.push('[REFRESH OK] ' + athId);
      } catch (e) {
        log.push('[ERRO REFRESH] ' + athId + ' — ' + e.message);
        totalErros++;
        continue;
      }
    }

    // Chamar Strava API
    try {
      const apiResp = UrlFetchApp.fetch(
        STRAVA_API_BASE + '/athlete/activities?per_page=20&page=1',
        { headers: { Authorization: 'Bearer ' + accessToken }, muteHttpExceptions: true }
      );
      if (apiResp.getResponseCode() !== 200) {
        log.push('[ERRO API] ' + athId + ' — HTTP ' + apiResp.getResponseCode());
        totalErros++;
        continue;
      }
      const atividades = JSON.parse(apiResp.getContentText());
      if (!Array.isArray(atividades) || atividades.length === 0) {
        log.push('[SKIP] ' + athId + ' — sem atividades');
        continue;
      }
      let novas = 0;
      for (const raw of atividades) {
        const sid = String(raw.id || '');
        if (!sid || existingIds.has(sid)) continue;
        const norm  = normalizarAtividadeStrava(raw, { ath_id: athId, ATH_ID: athId, nome: nomeAtleta, Atleta: nomeAtleta });
        const linha = normalizadoParaLinha(norm);
        novasLinhas.push(linha);
        existingIds.add(sid);
        novas++;
      }
      log.push('[OK] ' + athId + ' (' + nomeAtleta + ') — ' + novas + ' novas / ' + atividades.length + ' recebidas');
      totalImportadas += novas;
    } catch (e) {
      log.push('[ERRO] ' + athId + ' — ' + e.message);
      totalErros++;
    }

    Utilities.sleep(300);
  }

  // Gravar em batch
  if (novasLinhas.length > 0) {
    const startRow = Math.max(wsAtiv.getLastRow() + 1, 3);
    wsAtiv.getRange(startRow, 1, novasLinhas.length, 34).setValues(novasLinhas);
  }

  const resumo = [
    'IMPORTAÇÃO SEGURA CONCLUÍDA',
    'Novas: ' + totalImportadas + '  |  Erros: ' + totalErros,
    '',
    log.join('\n')
  ].join('\n');

  Logger.log(resumo);
}  

// ── DIAGNÓSTICO TEMPORÁRIO ──────────────────────────────────────
function _diagImport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) { console.log('[DIAG] ss=NULL'); return; }
  console.log('[DIAG] ss=' + ss.getName());

  const wsAtiv = ss.getSheetByName(H.SHEETS.ATIVIDADES);
  const wsTok  = ss.getSheetByName(H.SHEETS.TOKENS);
  console.log('[DIAG] wsAtiv=' + (wsAtiv ? 'OK' : 'NULL') + ' | wsTok=' + (wsTok ? 'OK' : 'NULL'));

  const props = PropertiesService.getScriptProperties();
  const cid = props.getProperty('STRAVA_CLIENT_ID');
  console.log('[DIAG] STRAVA_CLIENT_ID=' + (cid ? 'SET(' + cid.substring(0,4) + '...)' : 'NULL'));

  if (wsTok) {
    const rows = wsTok.getDataRange().getValues();
    console.log('[DIAG] TOKENS rows=' + rows.length + ' (inc. header)');
    if (rows.length > 1) {
      const r = rows[1];
      console.log('[DIAG] row1[0-5]=' + r.slice(0,6).join(' | '));
    }
  }
}
