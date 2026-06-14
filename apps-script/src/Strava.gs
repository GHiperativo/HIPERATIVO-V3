/**
 * ═══════════════════════════════════════════════════════════════════════
 * HIPERATIVO V3 — 02_Strava.gs
 * OAuth2 + Importação de atividades + Refresh automático
 * ═══════════════════════════════════════════════════════════════════════
 */

// ── GERAR LINK OAUTH PARA O ATLETA ───────────────────────────────────────────
function gerarLinkStrava() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt('Strava OAuth', 'Digite o ID do atleta (ex: ATH_001):', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;

  const athId = r.getResponseText().trim();
  if (!athId) { ui.alert('ID inválido'); return; }

  try {
    const url = _gerarUrlOAuth(athId);
    ui.alert('✅ Link gerado — envie para o atleta',
      `${url}\n\nQuando autorizar, o sistema salva o token automaticamente.`,
      ui.ButtonSet.OK);
    _log(athId, 'INFO', 'gerarLinkStrava', 'Link OAuth gerado', '');
  } catch (e) {
    ui.alert('❌ Erro', e.message, ui.ButtonSet.OK);
    _log(athId, 'ERRO', 'gerarLinkStrava', e.message, '');
  }
}

function _gerarUrlOAuth(athId) {
  const props = PropertiesService.getScriptProperties();
  const clientId  = props.getProperty('STRAVA_CLIENT_ID');
  const webAppUrl = props.getProperty('WEBAPP_URL');

  if (!clientId || !webAppUrl) {
    throw new Error('Configure STRAVA_CLIENT_ID e WEBAPP_URL primeiro (Menu → Setup → Configurar propriedades)');
  }

  return 'https://www.strava.com/oauth/authorize'
    + '?client_id='     + clientId
    + '&response_type=code'
    + '&redirect_uri='  + encodeURIComponent(webAppUrl + '?action=strava_callback')
    + '&approval_prompt=auto'
    + '&scope='         + encodeURIComponent('read,activity:read_all')
    + '&state='         + encodeURIComponent(athId);
}

// ── WEB APP — ROTEADOR PRINCIPAL ─────────────────────────────────────────────
function doGet(e) {
  const p = (e && e.parameter) ? e.parameter : {};

  // Auto-atualizar WEBAPP_URL com a URL real do deployment atual (auto-healing)
  try {
    const _url = ScriptApp.getService().getUrl();
    if (_url) PropertiesService.getScriptProperties().setProperty('WEBAPP_URL', _url);
  } catch (_e) {}

  // Verificação de webhook Strava (hub.challenge) — deve vir ANTES do restante
  if (p['hub.mode'] === 'subscribe') {
    const verifyToken = PropertiesService.getScriptProperties().getProperty('STRAVA_VERIFY_TOKEN');
    if (p['hub.verify_token'] === verifyToken) {
      return ContentService
        .createTextOutput(JSON.stringify({'hub.challenge': p['hub.challenge']}))
        .setMimeType(ContentService.MimeType.JSON);
    }
    _log('WEBHOOK', 'AVISO', 'doGet', 'hub.verify_token inválido recebido', '');
    return ContentService
      .createTextOutput(JSON.stringify({'error': 'invalid verify token'}))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // Callback OAuth do Strava
  if (p.action === 'strava_callback') return _processarCallback(p);

  // Ações JSON — devem vir ANTES do fallback de página
  if (p.action === 'salvar_cadastro') return _salvarCadastro(p);
  if (p.action === 'listar_atletas')  return _listarAtletas();
  if (p.action === 'get_metricas')    return _getMetricas(p.athId);

  // Painel do treinador (com PIN)
  if (p.page === 'treinador') return _paginaTreinador(p);

  // Página de cadastro (padrão — sem page ou page=cadastro)
  return _paginaCadastro(p);
}

// ── CALLBACK OAUTH ────────────────────────────────────────────────────────────
function _processarCallback(p) {
  const code  = p.code;
  const athId = decodeURIComponent(p.state || '');

  if (!code || !athId) {
    return HtmlService.createHtmlOutput('<h2>❌ Parâmetros inválidos no callback.</h2>');
  }

  try {
    const props = PropertiesService.getScriptProperties();
    const payload = {
      client_id:     props.getProperty('STRAVA_CLIENT_ID'),
      client_secret: props.getProperty('STRAVA_CLIENT_SECRET'),
      code:          code,
      grant_type:    'authorization_code',
    };

    const resp = UrlFetchApp.fetch('https://www.strava.com/oauth/token', {
      method: 'post',
      payload: payload,
      muteHttpExceptions: true,
    });

    const data = JSON.parse(resp.getContentText());
    if (!data.access_token) throw new Error('Token não retornado: ' + resp.getContentText());

    // Salvar na aba TOKENS
    _salvarToken(athId, data);

    // Atualizar CADASTRO
    _atualizarCampoAtleta(athId, H.CAD.STRAVA_OK, 'Sim');
    _atualizarCampoAtleta(athId, H.CAD.STRAVA_ID, data.athlete ? data.athlete.id : '');

    const nomeAtleta = _getNomeAtleta(athId);
    _log(athId, 'INFO', '_processarCallback', 'Strava conectado com sucesso', '');

    return HtmlService.createHtmlOutput(
      `<html><head><meta charset="UTF-8"></head><body style="font-family:Arial;text-align:center;padding:40px">
      <h2 style="color:#1D9E75">✅ Strava conectado!</h2>
      <p>Olá, <strong>${nomeAtleta}</strong>! Seus treinos serão importados automaticamente.</p>
      <p style="color:#888;font-size:13px">Pode fechar esta janela.</p>
      </body></html>`
    );

  } catch (e) {
    _log(athId, 'ERRO', '_processarCallback', e.message, '');
    return HtmlService.createHtmlOutput(`<h2>❌ Erro: ${e.message}</h2>`);
  }
}

// ── SALVAR TOKEN ──────────────────────────────────────────────────────────────
function _salvarToken(athId, data) {
  const ws = SpreadsheetApp.getActive().getSheetByName(H.SHEETS.TOKENS);
  if (!ws) {
    // Lança erro para que _processarCallback registre no log de ERROS
    throw new Error('Aba "' + H.SHEETS.TOKENS + '" não encontrada. Verifique o nome da aba na planilha.');
  }

  if (!data.refresh_token) {
    throw new Error('Strava não retornou refresh_token. Dados recebidos: ' + JSON.stringify(data).slice(0, 200));
  }

  const rows  = ws.getDataRange().getValues();
  const agora = new Date();
  const linha = [
    athId,
    data.athlete ? data.athlete.id : '',
    data.access_token,
    data.refresh_token,
    data.expires_at,
    agora,
    agora,
    'Válido',
  ];

  // Procura linha existente do atleta e atualiza
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === athId) {
      ws.getRange(i + 1, 1, 1, linha.length).setValues([linha]);
      _log(athId, 'INFO', '_salvarToken', 'Token salvo/atualizado na aba TOKENS (linha ' + (i+1) + ')', '');
      return;
    }
  }
  // Nova linha
  ws.appendRow(linha);
  _log(athId, 'INFO', '_salvarToken', 'Token salvo como nova linha na aba TOKENS', '');
}

// ── REFRESH DO TOKEN ──────────────────────────────────────────────────────────
function _getAccessToken(athId) {
  const ws   = SpreadsheetApp.getActive().getSheetByName(H.SHEETS.TOKENS);
  if (!ws) throw new Error('Aba TOKENS não encontrada');

  const rows = ws.getDataRange().getValues();
  let tokenRow = null, rowIdx = -1;

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === athId) {
      tokenRow = rows[i];
      rowIdx   = i + 1;
      break;
    }
  }

  if (!tokenRow) throw new Error(`Atleta ${athId} sem token. Gere o link OAuth primeiro.`);

  const expiresAt  = Number(tokenRow[H.TOKENS.EXPIRES - 1]);
  const agora      = Math.floor(Date.now() / 1000);

  // Token válido: retorna direto
  if (expiresAt && agora < expiresAt - 60) return String(tokenRow[H.TOKENS.ACCESS - 1]);

  // Precisa refresh
  const props = PropertiesService.getScriptProperties();
  const resp = UrlFetchApp.fetch('https://www.strava.com/oauth/token', {
    method: 'post',
    payload: {
      client_id:     props.getProperty('STRAVA_CLIENT_ID'),
      client_secret: props.getProperty('STRAVA_CLIENT_SECRET'),
      grant_type:    'refresh_token',
      refresh_token: String(tokenRow[H.TOKENS.REFRESH - 1]),
    },
    muteHttpExceptions: true,
  });

  const httpCode = resp.getResponseCode();
  const data     = JSON.parse(resp.getContentText());

  if (!data.access_token) {
    // Token inválido/revogado — marcar para que o atleta precise reconectar
    const motivo = httpCode === 400 ? 'Revogado' : 'Expirado';
    ws.getRange(rowIdx, H.TOKENS.STATUS, 1, 1).setValue(motivo);
    ws.getRange(rowIdx, H.TOKENS.ULT_ATU, 1, 1).setValue(new Date());
    // Atualizar CADASTRO: Strava = "Não" para refletir no Painel
    _atualizarCampoAtleta(athId, H.CAD.STRAVA_OK, 'Não');
    const msg = 'Token ' + motivo.toLowerCase() + ' (HTTP ' + httpCode + '). '
      + 'Atleta precisa reconectar o Strava.';
    _log(athId, 'ERRO', '_getAccessToken', msg, resp.getContentText().slice(0, 200));
    throw new Error(msg);
  }

  // Sucesso — salvar novos tokens (Strava rotaciona o refresh_token a cada uso)
  ws.getRange(rowIdx, H.TOKENS.ACCESS,   1, 1).setValue(data.access_token);
  ws.getRange(rowIdx, H.TOKENS.REFRESH,  1, 1).setValue(data.refresh_token);
  ws.getRange(rowIdx, H.TOKENS.EXPIRES,  1, 1).setValue(data.expires_at);
  ws.getRange(rowIdx, H.TOKENS.ULT_ATU,  1, 1).setValue(new Date());
  ws.getRange(rowIdx, H.TOKENS.STATUS,   1, 1).setValue('Válido');
  _atualizarCampoAtleta(athId, H.CAD.STRAVA_OK, 'Sim');

  _log(athId, 'INFO', '_getAccessToken', 'Token renovado com sucesso', '');
  return data.access_token;
}

// ── IMPORTAR ATIVIDADES — TODOS ───────────────────────────────────────────────
function importarStravaAll() {
  const ws     = SpreadsheetApp.getActive().getSheetByName(H.SHEETS.CADASTRO);
  if (!ws) { _log('SYSTEM', 'ERRO', 'importarStravaAll', 'Aba CADASTRO não encontrada', ''); return; }

  const rows = ws.getDataRange().getValues().slice(2); // pula header
  let importados = 0, erros = 0;

  rows.forEach(row => {
    const athId    = String(row[H.CAD.ID - 1] || '').trim();
    const stravaOk = String(row[H.CAD.STRAVA_OK - 1] || '').trim();
    const status   = String(row[H.CAD.STATUS - 1] || '').trim();

    if (!athId || stravaOk !== 'Sim' || status === 'Inativo') return;

    try {
      const n = _importarAtleta(athId);
      importados += n;
    } catch (e) {
      _log(athId, 'ERRO', 'importarStravaAll', e.message, '');
      erros++;
    }
  });

  _log('SYSTEM', 'INFO', 'importarStravaAll',
    `Importação concluída: ${importados} atividades | ${erros} erros`, '');
}

// ── HELPER: segundos → "H:MM:SS" ─────────────────────────────────────────────
function _segsParaHMS(s) {
  s = Math.round(Number(s) || 0);
  if (s <= 0) return '0:00:00';
  const h   = Math.floor(s / 3600);
  const m   = Math.floor((s % 3600) / 60);
  const seg = s % 60;
  return h + ':' + String(m).padStart(2, '0') + ':' + String(seg).padStart(2, '0');
}

// ── IMPORTAR ATIVIDADES — UM ATLETA ──────────────────────────────────────────
function importarStravaUm() {
  const ui = SpreadsheetApp.getUi();
  const r  = ui.prompt('Importar Strava', 'ID do atleta:', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;

  const athId = r.getResponseText().trim();
  try {
    const n = _importarAtleta(athId);
    ui.alert('✅ Importado', `${n} atividades importadas para ${athId} (mês vigente)`, ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('❌ Erro', e.message, ui.ButtonSet.OK);
    _log(athId, 'ERRO', 'importarStravaUm', e.message, '');
  }
}

// ── IMPORTAR STRAVA — PERÍODO SELECIONÁVEL ────────────────────────────────────
function importarStravaPeriodo() {
  const ui = SpreadsheetApp.getUi();

  const r1 = ui.prompt('📅 Importar Período — Passo 1/2',
    'Data de início (DD/MM/AAAA):', ui.ButtonSet.OK_CANCEL);
  if (r1.getSelectedButton() !== ui.Button.OK) return;

  const r2 = ui.prompt('📅 Importar Período — Passo 2/2',
    'Data de fim (DD/MM/AAAA):', ui.ButtonSet.OK_CANCEL);
  if (r2.getSelectedButton() !== ui.Button.OK) return;

  function parseData(str) {
    const p = str.trim().split('/');
    if (p.length !== 3) return null;
    const d = new Date(Number(p[2]), Number(p[1]) - 1, Number(p[0]));
    return isNaN(d.getTime()) ? null : d;
  }

  const inicio = parseData(r1.getResponseText());
  const fim    = parseData(r2.getResponseText());

  if (!inicio || !fim) {
    ui.alert('❌ Datas inválidas', 'Use o formato DD/MM/AAAA (ex: 01/05/2025)', ui.ButtonSet.OK);
    return;
  }
  if (inicio > fim) {
    ui.alert('❌ Erro', 'A data de início deve ser anterior à data de fim.', ui.ButtonSet.OK);
    return;
  }

  // Fim do dia final
  fim.setHours(23, 59, 59, 999);

  const afterTs  = Math.floor(inicio.getTime() / 1000);
  const beforeTs = Math.floor(fim.getTime() / 1000);

  const ws = SpreadsheetApp.getActive().getSheetByName(H.SHEETS.CADASTRO);
  if (!ws) { ui.alert('❌ Aba CADASTRO não encontrada'); return; }

  const rows = ws.getDataRange().getValues().slice(2);
  let total = 0, erros = 0;

  rows.forEach(row => {
    const athId    = String(row[H.CAD.ID - 1] || '').trim();
    const stravaOk = String(row[H.CAD.STRAVA_OK - 1] || '').trim();
    const status   = String(row[H.CAD.STATUS - 1] || '').trim();
    if (!athId || stravaOk !== 'Sim' || status === 'Inativo') return;
    try {
      const n = _importarAtleta(athId, afterTs, beforeTs);
      total += n;
    } catch (e) {
      _log(athId, 'ERRO', 'importarStravaPeriodo', e.message, '');
      erros++;
    }
  });

  _log('SYSTEM', 'INFO', 'importarStravaPeriodo',
    `Período ${r1.getResponseText()} a ${r2.getResponseText()}: ${total} atividades | ${erros} erros`, '');
  ui.alert('✅ Importação concluída',
    `Período: ${r1.getResponseText()} → ${r2.getResponseText()}\n` +
    `✅ ${total} atividades importadas\n❌ ${erros} atletas com erro`,
    ui.ButtonSet.OK);
}

// ── RECONCILIAR ATIVIDADES DIÁRIAS ────────────────────────────────────────────
// Importa todos os atletas ativos e loga o total real de novas atividades.
function reconciliarAtividadesDiarias() {
  const wsCad = SpreadsheetApp.getActive().getSheetByName(H.SHEETS.CADASTRO);
  if (!wsCad) {
    _log('SYSTEM', 'ERRO', 'reconciliarAtividadesDiarias', 'Aba CADASTRO não encontrada', '');
    return 0;
  }

  const rows = wsCad.getDataRange().getValues().slice(2);
  let totalNovas = 0;
  let erros = 0;

  rows.forEach(row => {
    const athId    = String(row[H.CAD.ID - 1] || '').trim();
    const stravaOk = String(row[H.CAD.STRAVA_OK - 1] || '').trim();
    const status   = String(row[H.CAD.STATUS - 1] || '').trim();
    if (!athId || stravaOk !== 'Sim' || status === 'Inativo') return;
    try {
      const novas = _importarAtleta(athId);
      totalNovas += (novas || 0);
    } catch (e) {
      _log(athId, 'ERRO', 'reconciliarAtividadesDiarias', e.message, '');
      erros++;
    }
  });

  _log('SYSTEM', 'INFO', 'reconciliarAtividadesDiarias',
    totalNovas + ' nova(s) atividade(s) importada(s)' + (erros > 0 ? ' | ' + erros + ' erros' : ''), '');
  return totalNovas;
}

// ── ALIAS PARA TRIGGER AUTOMÁTICO ────────────────────────────────────────────
// O trigger externo (criado por outro usuário) dispara a cada 5 min.
// Guard interno impede execuções mais frequentes que 22 h para preservar a cota Strava.
function triggerImportacaoAutomatica() {
  const props = PropertiesService.getScriptProperties();
  const CHAVE = 'ULTIMA_IMPORTACAO_AUTO';
  const agora = Date.now();
  const ultima = Number(props.getProperty(CHAVE) || 0);
  if (agora - ultima < 22 * 3600 * 1000) {
    _log('SISTEMA', 'INFO', 'triggerImportacaoAutomatica',
      'Ignorado — última execução foi há ' + Math.round((agora - ultima) / 3600000) + 'h (mínimo: 22h)', '');
    return;
  }
  props.setProperty(CHAVE, String(agora));

  const ws = SpreadsheetApp.getActive().getSheetByName(H.SHEETS.TOKENS);
  if (!ws || ws.getLastRow() < 2) {
    _log('SISTEMA', 'INFO', 'triggerImportacaoAutomatica',
      '0 nova(s) atividade(s) importada(s) — sem tokens', '');
    return;
  }

  const tokens = ws.getDataRange().getValues().slice(1)
    .filter(r => r[H.TOKENS.REFRESH - 1]);

  if (!tokens.length) {
    _log('SISTEMA', 'INFO', 'triggerImportacaoAutomatica',
      '0 nova(s) atividade(s) importada(s) — sem refresh_token ativo', '');
    return;
  }

  _log('SISTEMA', 'INFO', 'triggerImportacaoAutomatica', 'Iniciando reconciliação diária...', '');
  reconciliarAtividadesDiarias();
}

/**
 * _importarAtleta(athId, afterTs?, beforeTs?)
 *
 * afterTs  — Unix timestamp (segundos) do início do período (opcional)
 * beforeTs — Unix timestamp (segundos) do fim do período (opcional)
 *
 * Sem parâmetros: importa o mês vigente (do dia 1º até agora).
 */
function _importarAtleta(athId, afterTs, beforeTs) {
  const token = _getAccessToken(athId);

  // Período: mês vigente por padrão
  let after, before;
  if (afterTs !== undefined && afterTs !== null) {
    after  = afterTs;
    before = beforeTs || Math.floor(Date.now() / 1000);
  } else {
    const agora     = new Date();
    const inicioMes = new Date(agora.getFullYear(), agora.getMonth(), 1);
    after  = Math.floor(inicioMes.getTime() / 1000);
    before = Math.floor(agora.getTime() / 1000);
  }

  const url = `https://www.strava.com/api/v3/athlete/activities?per_page=100&after=${after}&before=${before}`;
  const resp = UrlFetchApp.fetch(url,
    { headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true }
  );

  if (resp.getResponseCode() === 429) {
    Utilities.sleep(60000);
    throw new Error('Rate limit Strava — aguarde 1 min e tente de novo');
  }

  const atividades = JSON.parse(resp.getContentText());
  if (!Array.isArray(atividades)) throw new Error('Resposta inválida do Strava');

  const ws         = SpreadsheetApp.getActive().getSheetByName(H.SHEETS.ATIVIDADES);
  const existentes = _getStravaIdsExistentes(ws);
  const nomeAtleta = _getNomeAtleta(athId);
  const TIPOS      = ['Run','TrailRun','Walk','Hike','Ride','VirtualRide',
                       'Swim','WeightTraining','Rowing','Yoga','Workout','EBikeRide'];
  const novas      = [];

  atividades.forEach(a => {
    if (!TIPOS.includes(a.type) && !TIPOS.includes(a.sport_type)) return;
    if (existentes.has(String(a.id))) return;

    const velMps   = a.average_speed || 0;
    const velKmh   = velMps > 0 ? Math.round(velMps * 36) / 10 : 0;       // km/h, 1 decimal
    const paceS    = velMps > 0 ? Math.round(1000 / velMps) : 0;           // s/km
    const paceMs   = paceS  > 0
      ? Math.floor(paceS / 60) + ':' + String(paceS % 60).padStart(2, '0')
      : '';
    const paceRap  = paceS  > 0
      ? Math.floor(paceS * 0.82 / 60) + ':' + String(Math.round(paceS * 0.82) % 60).padStart(2, '0')
      : '';
    const distM    = Math.round((a.distance || 0) * 10) / 10;             // metros, 1 decimal
    const distKm   = Math.round((a.distance || 0) / 100) / 10;            // km, 1 decimal

    const tipoPT = {
      Run: 'Corrida', WeightTraining: 'Academia',
      Ride: 'Ciclismo', Walk: 'Caminhada', VirtualRide: 'Ergométrica',
      TrailRun: 'Corrida (Trail)', Swim: 'Natação', Hike: 'Trilha',
      Rowing: 'Remo', Yoga: 'Yoga', Workout: 'Treino',
    }[a.sport_type || a.type] || (a.sport_type || a.type);

    novas.push([
      // ── Identificação (cols 1-8) ──────────────────────────────────────────
      'ATIV_' + Math.random().toString(36).slice(2, 10).toUpperCase(),
      athId, nomeAtleta,
      new Date(a.start_date_local || a.start_date),
      tipoPT, 'Strava', String(a.id),
      a.name || '',
      // ── Tempo (cols 9-10): hh:mm:ss ──────────────────────────────────────
      _segsParaHMS(a.moving_time  || 0),
      _segsParaHMS(a.elapsed_time || 0),
      // ── Distância (cols 11-12) ────────────────────────────────────────────
      distM,    // metros (1 decimal)
      distKm,   // km (1 decimal)
      // ── Velocidade e Pace (cols 13-16) ───────────────────────────────────
      velKmh,   // km/h (1 decimal)
      paceS,    // s/km (para fórmulas do painel)
      paceMs,   // min:ss/km (texto formatado)
      paceRap,  // pace mais rápido estimado
      // ── Frequência cardíaca (cols 17-18) ─────────────────────────────────
      a.average_heartrate ? Math.round(a.average_heartrate * 10) / 10 : '',
      a.max_heartrate     || '',
      // ── Esforço físico (cols 19-20) ───────────────────────────────────────
      Math.round((a.total_elevation_gain || 0) * 10) / 10,
      a.calories          || '',
      // ── Campos manuais (cols 21-22) ───────────────────────────────────────
      '', // RPE (preenchido manualmente)
      '', // Observações
      // ── Dados extras Strava (cols 23-26) ──────────────────────────────────
      a.sport_type        || a.type || '',             // Tipo Esporte (detalhado)
      a.average_cadence   ? Math.round(a.average_cadence) : '',  // Cadência (ppm)
      a.suffer_score      || '',                       // Esforço Relativo (0-100)
      new Date(),                                      // Data Importação
    ]);
  });

  if (novas.length > 0) {
    const linha = _primeiraLinhaVazia(ws, H.ATIV.EXEC_ID);
    ws.getRange(linha, 1, novas.length, 26).setValues(novas);
    SpreadsheetApp.flush();
    // Recalcula métricas imediatamente após gravar novas atividades
    try { _calcularMetricasAtleta(athId); } catch(e) {
      _log(athId, 'AVISO', '_importarAtleta', 'Métricas: ' + e.message, '');
    }
  }

  const periodoDesc = afterTs !== undefined
    ? `período personalizado (after=${afterTs})`
    : 'mês vigente';
  _log(athId, 'INFO', '_importarAtleta', novas.length + ' atividades importadas — ' + periodoDesc, '');
  return novas.length;
}

function _getStravaIdsExistentes(ws) {
  const lastRow = ws.getLastRow();
  if (lastRow < 3) return new Set();
  const ids = ws.getRange(3, H.ATIV.STRAVA_ID, lastRow - 2, 1)
    .getValues().flat().filter(Boolean).map(String);
  return new Set(ids);
}

// Retorna a primeira linha onde a coluna colIdx está vazia (a partir da linha 3)
function _primeiraLinhaVazia(ws, colIdx) {
  const lastRow = ws.getLastRow();
  if (lastRow < 3) return 3;
  const vals = ws.getRange(3, colIdx, lastRow - 2, 1).getValues().flat();
  for (let i = 0; i < vals.length; i++) {
    if (!vals[i] && vals[i] !== 0) return i + 3;
  }
  return lastRow + 1;
}

function _getDataCadastroAtleta(athId) {
  const ws   = SpreadsheetApp.getActive().getSheetByName(H.SHEETS.CADASTRO);
  if (!ws) return new Date();
  const rows = ws.getDataRange().getValues().slice(2);
  const r    = rows.find(row => String(row[H.CAD.ID - 1]) === athId);
  if (!r) return new Date();
  const d = r[H.CAD.DATA_CAD - 1];
  return d instanceof Date ? d : new Date();
}

// ── HELPERS ───────────────────────────────────────────────────────────────────
function _getNomeAtleta(athId) {
  const ws   = SpreadsheetApp.getActive().getSheetByName(H.SHEETS.CADASTRO);
  if (!ws) return athId;
  const rows = ws.getDataRange().getValues().slice(2);
  const r    = rows.find(row => String(row[H.CAD.ID - 1]) === athId);
  return r ? String(r[H.CAD.NOME - 1]) : athId;
}

function _atualizarCampoAtleta(athId, colIdx, valor) {
  const ws   = SpreadsheetApp.getActive().getSheetByName(H.SHEETS.CADASTRO);
  if (!ws) return;
  const rows = ws.getDataRange().getValues();
  for (let i = 2; i < rows.length; i++) {
    if (String(rows[i][H.CAD.ID - 1]) === athId) {
      ws.getRange(i + 1, colIdx).setValue(valor);
      return;
    }
  }
}

// ── CONFIGURAR TRIGGER AUTOMÁTICO (mês corrente, diário) ─────────────────────
function criarTriggerImportacaoDiaria() {
  // Remove triggers antigos do mesmo tipo para evitar duplicação
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'triggerImportacaoAutomatica') {
      ScriptApp.deleteTrigger(t);
    }
  });

  // Cria trigger diário às 06:00 (horário do servidor — Brasília é UTC-3)
  ScriptApp.newTrigger('triggerImportacaoAutomatica')
    .timeBased()
    .everyDays(1)
    .atHour(6)
    .create();

  try {
    SpreadsheetApp.getUi().alert(
      '✅ Trigger criado',
      'Importação automática agendada para todo dia às 06h.\n\n' +
      'Importa apenas atividades do MÊS CORRENTE.\n' +
      'Para período personalizado use: Menu → Strava → Importar período.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch(e) {
    Logger.log('✅ Trigger diário criado: triggerImportacaoAutomatica às 06h');
  }
}

function removerTriggersDiarios() {
  let count = 0;
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'triggerImportacaoAutomatica') {
      ScriptApp.deleteTrigger(t);
      count++;
    }
  });
  try {
    SpreadsheetApp.getUi().alert('🗑️ ' + count + ' trigger(s) removido(s).');
  } catch(e) {
    Logger.log('🗑️ ' + count + ' trigger(s) removido(s).');
  }
}

// ── PÁGINAS HTML MÍNIMAS (doGet) ──────────────────────────────────────────────
function _paginaCadastro(p) {
  const conectado = p.strava === 'conectado';
  const nome      = p.nome || '';
  const html = `<!DOCTYPE html><html lang="pt-BR"><head><meta charset="UTF-8">
  <meta name="viewport" content="width=device-width,initial-scale=1">
  <title>Grupo Hiperativo — Cadastro</title>
  <style>
    body{font-family:Arial,sans-serif;max-width:520px;margin:40px auto;padding:0 20px;color:#111}
    h2{color:#001F3F}input,select{width:100%;padding:9px;margin:6px 0 14px;border:1px solid #ccc;border-radius:6px;font-size:14px}
    button{width:100%;padding:12px;background:#00B4FF;color:#fff;border:none;border-radius:8px;font-size:15px;font-weight:bold;cursor:pointer}
    .ok{background:#D6F5EC;border:1px solid #1D9E75;padding:14px;border-radius:8px;color:#0A5C3A;margin-bottom:20px}
    label{font-size:13px;font-weight:bold;color:#444}
  </style></head><body>
  ${conectado ? `<div class="ok">✅ Strava conectado com sucesso${nome ? ', ' + nome : ''}! Seus treinos serão importados automaticamente.</div>` : ''}
  <h2>⚡ Cadastro — Grupo Hiperativo</h2>
  <form id="f">
    <label>Nome completo</label><input name="nome" required>
    <label>E-mail</label><input type="email" name="email" required>
    <label>WhatsApp (com DDD)</label><input name="whats" required>
    <label>Data de nascimento</label><input type="date" name="nasc" required>
    <label>Sexo biológico</label>
    <select name="sexo"><option>Masculino</option><option>Feminino</option><option>Outro</option></select>
    <label>Peso (kg)</label><input type="number" name="peso" min="30" max="200">
    <label>Modalidade de interesse</label>
    <select name="mod"><option>Corrida</option><option>Academia</option><option>Ergométrica</option><option>Híbrido</option></select>
    <label>Nível atual</label>
    <select name="nivel"><option>Iniciante</option><option>Intermediário</option><option>Avançado</option></select>
    <label>Objetivo principal</label>
    <select name="obj"><option>Emagrecimento</option><option>Condicionamento</option><option>Performance</option><option>Saúde Mental</option><option>Força</option></select>
    <label>Dias disponíveis por semana</label>
    <select name="freq"><option>2x</option><option>3x</option><option>4x</option><option>5x</option></select>
    <label>Plano de interesse</label>
    <select name="plano"><option>Experimental (avulso)</option><option>Online R$100</option><option>Presencial R$120</option><option>Híbrido R$160</option></select>
    <label>Como nos encontrou?</label>
    <select name="origem"><option>Instagram</option><option>Indicação</option><option>Strava</option><option>Google</option><option>Outro</option></select>
    <button type="submit">Enviar cadastro</button>
  </form>
  <script>
    document.getElementById('f').onsubmit = function(e){
      e.preventDefault();
      const fd  = new FormData(e.target);
      const dados = Object.fromEntries(fd);
      const btn = e.target.querySelector('button[type=submit]');
      btn.disabled = true;
      btn.innerText = 'Enviando...';
      google.script.run
        .withSuccessHandler(function(res){
          btn.disabled = false;
          btn.innerText = 'Enviar cadastro';
          if(res.ok){ alert('✅ Cadastro salvo! ID: ' + res.athId); e.target.reset(); }
          else { alert('❌ Erro: ' + res.erro); }
        })
        .withFailureHandler(function(err){
          btn.disabled = false;
          btn.innerText = 'Enviar cadastro';
          alert('❌ Erro: ' + err.message);
        })
        .salvarCadastroForm(dados);
    };
  </script></body></html>`;
  return HtmlService.createHtmlOutput(html).setTitle('Cadastro Hiperativo');
}

function _paginaTreinador(p) {
  const pinConfig = getCfg('PIN_TREINADOR') || '1234';
  const pinEnv    = p.pin || '';
  if (pinEnv !== pinConfig) {
    return HtmlService.createHtmlOutput(
      `<html><head><meta charset="UTF-8"></head><body style="font-family:Arial;max-width:360px;margin:60px auto">
      <h3>🔐 Painel do Treinador</h3>
      <input id="pin" type="password" placeholder="PIN" style="width:100%;padding:10px;font-size:16px;margin:10px 0;border:1px solid #ccc;border-radius:6px">
      <button onclick="location.href=location.href+'&pin='+document.getElementById('pin').value"
        style="width:100%;padding:12px;background:#001F3F;color:#fff;border:none;border-radius:8px;cursor:pointer;font-size:15px">
        Entrar</button></body></html>`
    );
  }
  return HtmlService.createHtmlOutput(
    '<html><head><meta charset="UTF-8"></head><body style="font-family:Arial;padding:40px">' +
    '<h2>⚡ Painel do Treinador</h2>' +
    '<p>Use o artefato do prescritor (Claude Designer) para gerar e enviar prescrições.</p>' +
    '<p>Este Web App serve como backend para o artefato.</p>' +
    '</body></html>'
  );
}

// ── CADASTRO VIA google.script.run (chamado pelo formulário HTML) ─────────────
function salvarCadastroForm(dados) {
  try {
    const ss = SpreadsheetApp.getActive();
    const ws = ss.getSheetByName(H.SHEETS.CADASTRO);
    if (!ws) throw new Error(
      'Aba "' + H.SHEETS.CADASTRO + '" não encontrada. ' +
      'Execute o Setup da planilha primeiro (Menu ⚡ HIPERATIVO → 🛠️ Setup).'
    );

    const athId = 'ATH_' + Date.now();
    const agora = new Date();

    ws.appendRow([
      athId, dados.nome, dados.email, dados.whats, dados.nasc,
      dados.sexo, dados.peso || '', dados.mod, dados.nivel, dados.obj,
      dados.freq, '', '', '', dados.plano, dados.origem,
      agora, 'Não', '', 'Trial', '',
    ]);
    SpreadsheetApp.flush();

    // Auto-inicializa linha de métricas para o novo atleta
    try { _calcularMetricasAtleta(athId); } catch(e) {
      _log(athId, 'AVISO', 'salvarCadastroForm', 'Métricas não iniciadas: ' + e.message, '');
    }

    _log(athId, 'INFO', 'salvarCadastroForm', 'Novo atleta: ' + dados.nome, '');
    return { ok: true, athId: athId };
  } catch (e) {
    _log('FORM', 'ERRO', 'salvarCadastroForm', e.message,
         JSON.stringify(dados || {}).slice(0, 300));
    return { ok: false, erro: e.message };
  }
}

// ── AÇÕES JSON (chamadas via doGet pelo artefato externo) ─────────────────────
function _salvarCadastro(p) {
  try {
    const res = salvarCadastroForm(p);
    return ContentService.createTextOutput(JSON.stringify(res))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (e) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, erro: e.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function _listarAtletas() {
  try {
    const ws   = SpreadsheetApp.getActive().getSheetByName(H.SHEETS.CADASTRO);
    const rows = ws.getDataRange().getValues().slice(2);
    const lista = rows
      .filter(r => r[H.CAD.STATUS - 1] !== 'Inativo' && r[H.CAD.ID - 1])
      .map(r => ({ id: r[H.CAD.ID-1], nome: r[H.CAD.NOME-1], strava: r[H.CAD.STRAVA_OK-1] }));
    return ContentService.createTextOutput(JSON.stringify({ ok: true, lista }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (e) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, erro: e.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function _getMetricas(athId) {
  try {
    const ws   = SpreadsheetApp.getActive().getSheetByName(H.SHEETS.METRICAS);
    const rows = ws.getDataRange().getValues().slice(2);
    const r    = rows.find(row => String(row[H.MET.ATH_ID - 1]) === athId);
    if (!r) return ContentService.createTextOutput(JSON.stringify({ ok: false, erro: 'Sem métricas' }))
      .setMimeType(ContentService.MimeType.JSON);
    const m = {
      vo2: r[H.MET.VO2-1], paceMed: r[H.MET.PACE_MED-1],
      paceRap: r[H.MET.PACE_RAP-1], paceLento: r[H.MET.PACE_LEN-1],
      fcMax: r[H.MET.FC_MAX-1], fcMed: r[H.MET.FC_MED-1],
      volSem: r[H.MET.VOL_SEM-1],
    };
    return ContentService.createTextOutput(JSON.stringify({ ok: true, metricas: m }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (e) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, erro: e.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ── WEBHOOK STRAVA (POST) ─────────────────────────────────────────────────────
function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return ContentService.createTextOutput('OK');
    }
    const payload = JSON.parse(e.postData.contents);

    // Validar que é uma notificação legítima do Strava
    if (!payload.object_type || !payload.object_id) {
      _log('WEBHOOK', 'AVISO', 'doPost', 'Payload inválido rejeitado', JSON.stringify(payload).slice(0, 200));
      return ContentService.createTextOutput('OK');
    }

    _log('WEBHOOK', 'INFO', 'doPost',
      'Evento: ' + payload.object_type + '/' + payload.aspect_type,
      JSON.stringify(payload).slice(0, 500));

    // Nova atividade criada → importar imediatamente
    if (payload.object_type === 'activity' && payload.aspect_type === 'create') {
      _processarWebhookAtividade(String(payload.owner_id || ''), payload.object_id);
    }

    return ContentService.createTextOutput('EVENT_RECEIVED');
  } catch (err) {
    _log('WEBHOOK', 'ERRO', 'doPost', err.message, '');
    return ContentService.createTextOutput('OK');
  }
}

function _processarWebhookAtividade(stravaAthId, stravaActId) {
  if (!stravaAthId) return;
  const wsTok = SpreadsheetApp.getActive().getSheetByName(H.SHEETS.TOKENS);
  if (!wsTok) return;

  const rows = wsTok.getDataRange().getValues().slice(1);
  const tokenRow = rows.find(r => String(r[H.TOKENS.STRAVA_ATH_ID - 1]) === stravaAthId);
  if (!tokenRow) return;

  const athId = String(tokenRow[H.TOKENS.ATH_ID - 1]);
  _log(athId, 'INFO', '_processarWebhookAtividade', 'Nova atividade via webhook, Strava ID: ' + stravaActId, '');

  const agora = Math.floor(Date.now() / 1000);
  try {
    _importarAtleta(athId, agora - 86400, agora); // janela de 24h
    // métricas já recalculadas dentro de _importarAtleta quando há novas atividades
  } catch (e) {
    _log(athId, 'ERRO', '_processarWebhookAtividade', e.message, '');
  }
}

// ── SEGURANÇA: VERIFICAR ACESSO DE ADMIN ─────────────────────────────────────
function verificarAcessoAdmin() {
  const emailUsuario = Session.getActiveUser().getEmail();
  const prop = PropertiesService.getScriptProperties().getProperty('ADMIN_EMAILS');
  const admins = prop
    ? prop.split(',').map(function(s) { return s.trim(); })
    : ['crhystianoh@gmail.com'];
  if (!admins.includes(emailUsuario)) {
    throw new Error('Acesso não autorizado para esta operação: ' + emailUsuario);
  }
}

// ── VERIFICAR E LIMPAR TOKENS DUPLICADOS / REVOGADOS ─────────────────────────
function limparTokensInvalidos() {
  let ui = null;
  try { ui = SpreadsheetApp.getUi(); } catch(e) {}
  const ss    = SpreadsheetApp.getActive();
  const wsTok = ss.getSheetByName(H.SHEETS.TOKENS);
  if (!wsTok) { if (ui) ui.alert('Aba TOKENS não encontrada'); Logger.log('Aba TOKENS não encontrada'); return; }

  const rows = wsTok.getDataRange().getValues();
  const vistos   = {};
  const remover  = []; // linhas a deletar (índice base 1, do maior para o menor)
  let msg = '';

  for (let i = 1; i < rows.length; i++) {
    const athId  = String(rows[i][H.TOKENS.ATH_ID - 1] || '').trim();
    const status = String(rows[i][H.TOKENS.STATUS - 1]  || '').trim();
    if (!athId) { remover.push(i + 1); continue; }

    if (vistos[athId]) {
      // Duplicado — remover a linha mais antiga (a anterior)
      remover.push(i + 1);
      msg += '• Duplicado removido: ' + athId + ' (linha ' + (i + 1) + ')\n';
    } else {
      vistos[athId] = i + 1;
    }
  }

  // Deletar de baixo para cima para não deslocar índices
  remover.sort((a, b) => b - a).forEach(r => wsTok.deleteRow(r));

  // Relatar tokens revogados que precisam de reautorização
  const rowsAtual = wsTok.getDataRange().getValues().slice(1);
  const revogados = rowsAtual.filter(r =>
    String(r[H.TOKENS.STATUS - 1]) === 'Revogado' ||
    !String(r[H.TOKENS.REFRESH - 1]).trim()
  );

  if (revogados.length > 0) {
    msg += '\nATLETAS QUE PRECISAM RECONECTAR O STRAVA:\n';
    revogados.forEach(r => {
      msg += '• ' + String(r[H.TOKENS.ATH_ID - 1]) + '\n';
    });
    msg += '\nUse: Menu → Atletas → Gerar link Strava para cada um.';
  }

  if (!msg) msg = '✅ Nenhum token duplicado ou inválido encontrado.';
  _log('SYSTEM', 'INFO', 'limparTokensInvalidos', msg.replace(/\n/g,' | '), '');
  Logger.log('limparTokensInvalidos: ' + msg);
  if (ui) ui.alert('🔧 Tokens', msg, ui.ButtonSet.OK);
}

// ── SINCRONIZAR STATUS STRAVA (TOKENS → CADASTRO) ────────────────────────────
// Roda no onOpen e no atualizarPainel. Não faz chamadas externas.
function sincronizarStatusStrava() {
  const ss    = SpreadsheetApp.getActive();
  const wsCad = ss.getSheetByName(H.SHEETS.CADASTRO);
  const wsTok = ss.getSheetByName(H.SHEETS.TOKENS);
  if (!wsCad || !wsTok) return;

  // Montar mapa athId → tem refresh_token
  const tokenMap = {};
  const tokData  = wsTok.getDataRange().getValues().slice(1);
  tokData.forEach((r, idx) => {
    const athId   = String(r[H.TOKENS.ATH_ID   - 1] || '').trim();
    const refresh = String(r[H.TOKENS.REFRESH   - 1] || '').trim();
    if (!athId) return;

    tokenMap[athId] = !!refresh;

    // Garantir que STATUS = 'Válido' quando refresh existe (sem chamar API)
    if (refresh && String(r[H.TOKENS.STATUS - 1]) !== 'Válido') {
      wsTok.getRange(idx + 2, H.TOKENS.STATUS).setValue('Válido');
    }
  });

  // Atualizar coluna STRAVA_OK no CADASTRO
  const cadRows = wsCad.getDataRange().getValues();
  let atualizados = 0;
  for (let i = 2; i < cadRows.length; i++) {
    const athId = String(cadRows[i][H.CAD.ID - 1] || '').trim();
    if (!athId) continue;

    const esperado = tokenMap[athId] ? 'Sim' : 'Não';
    if (String(cadRows[i][H.CAD.STRAVA_OK - 1]).trim() !== esperado) {
      wsCad.getRange(i + 1, H.CAD.STRAVA_OK).setValue(esperado);
      atualizados++;
    }
  }

  if (atualizados > 0) {
    _log('SYSTEM', 'INFO', 'sincronizarStatusStrava',
      atualizados + ' status Strava atualizados no CADASTRO', '');
  }
}

// ── DIAGNÓSTICO DE CONEXÃO STRAVA ─────────────────────────────────────────────
function diagnosticarConexaoStrava() {
  const ui     = SpreadsheetApp.getUi();
  const wsCad  = SpreadsheetApp.getActive().getSheetByName(H.SHEETS.CADASTRO);
  const wsTok  = SpreadsheetApp.getActive().getSheetByName(H.SHEETS.TOKENS);
  const wsAtiv = SpreadsheetApp.getActive().getSheetByName(H.SHEETS.ATIVIDADES);

  if (!wsCad) { ui.alert('Aba CADASTRO não encontrada'); return; }

  const atletas = wsCad.getDataRange().getValues().slice(2).filter(r => r[H.CAD.ID - 1]);
  if (!atletas.length) { ui.alert('Nenhum atleta cadastrado ainda.'); return; }

  const tokens   = wsTok  ? wsTok.getDataRange().getValues().slice(2) : [];
  const ativRows = wsAtiv ? wsAtiv.getDataRange().getValues().slice(2).filter(r => r[H.ATIV.EXEC_ID - 1]) : [];

  let msg = 'ATLETAS CADASTRADOS:\n' + '─'.repeat(40) + '\n';

  atletas.forEach(r => {
    const athId       = String(r[H.CAD.ID - 1]);
    const nome        = String(r[H.CAD.NOME - 1] || '—');
    const strOk       = String(r[H.CAD.STRAVA_OK - 1] || 'Não');
    const status      = String(r[H.CAD.STATUS - 1] || '—');
    const tokenRow    = tokens.find(t => String(t[H.TOKENS.ATH_ID - 1]) === athId);
    const tokenStatus = tokenRow ? String(tokenRow[H.TOKENS.STATUS - 1] || '—') : 'Sem token';
    const qtdAtiv     = ativRows.filter(a => String(a[H.ATIV.ATH_ID - 1]) === athId).length;
    const icon        = strOk === 'Sim' ? '✅' : '❌';

    msg += '\n' + icon + ' ' + nome + ' (' + athId + ')\n';
    msg += '   Status: ' + status + ' | Strava: ' + strOk + ' | Token: ' + tokenStatus + '\n';
    msg += '   Atividades importadas: ' + qtdAtiv + '\n';
  });

  msg += '\n' + '─'.repeat(40) + '\n';
  msg += 'Strava ❌ = link OAuth não foi autorizado.\nUse Menu → Atletas → Gerar link Strava.';

  ui.alert('🔍 Diagnóstico Strava', msg, ui.ButtonSet.OK);
}

// ── GERAR LINKS DE RECONEXÃO PARA TODOS OS TOKENS INVÁLIDOS ──────────────────
function gerarLinksReconexaoTodos() {
  const ss  = SpreadsheetApp.getActive();
  const ui  = SpreadsheetApp.getUi();
  const wsTok = ss.getSheetByName(H.SHEETS.TOKENS);
  if (!wsTok) { ui.alert('Aba TOKENS não encontrada.'); return; }

  const props    = PropertiesService.getScriptProperties();
  const clientId = props.getProperty('STRAVA_CLIENT_ID');
  const webAppUrl = props.getProperty('WEBAPP_URL');
  if (!clientId || !webAppUrl) {
    ui.alert('Configure STRAVA_CLIENT_ID e WEBAPP_URL primeiro (Menu → Setup).');
    return;
  }

  const rows = wsTok.getDataRange().getValues().slice(1);
  const invalidos = rows.filter(r =>
    String(r[H.TOKENS.STATUS - 1]) === 'Revogado' ||
    !String(r[H.TOKENS.REFRESH - 1]).trim()
  );

  if (invalidos.length === 0) {
    ui.alert('✅ Todos os tokens estão válidos. Nenhum link necessário.');
    return;
  }

  const wsCad = ss.getSheetByName(H.SHEETS.CADASTRO);
  const cadRows = wsCad ? wsCad.getDataRange().getValues().slice(2) : [];
  function getNome(id) {
    const r = cadRows.find(row => String(row[H.CAD.ID - 1]) === id);
    return r ? String(r[H.CAD.NOME - 1]) : '';
  }

  let html = '<style>'
    + 'body{font-family:Arial,sans-serif;font-size:13px;padding:8px}'
    + 'h3{color:#FC4C02;margin:0 0 8px}'
    + '.atleta{margin-bottom:14px}'
    + '.nome{font-weight:bold;margin-bottom:3px}'
    + 'input{width:100%;padding:6px;font-size:11px;border:1px solid #ccc;border-radius:4px;box-sizing:border-box;cursor:pointer}'
    + 'input:focus{outline:2px solid #FC4C02}'
    + '</style>'
    + '<h3>🔗 Links de Reconexão Strava</h3>'
    + '<p style="margin:0 0 12px;color:#555">Clique no link para selecionar, depois Ctrl+C para copiar.</p>';

  invalidos.forEach(r => {
    const athId = String(r[H.TOKENS.ATH_ID - 1]);
    const nome  = getNome(athId);
    const url   = 'https://www.strava.com/oauth/authorize'
      + '?client_id=' + clientId
      + '&response_type=code'
      + '&redirect_uri=' + encodeURIComponent(webAppUrl + '?action=strava_callback')
      + '&approval_prompt=force'
      + '&scope=read,activity:read_all'
      + '&state=' + athId;
    html += '<div class="atleta">'
      + '<div class="nome">' + athId + (nome ? ' — ' + nome : '') + '</div>'
      + '<input type="text" value="' + url + '" onclick="this.select()" readonly>'
      + '</div>';
  });

  ui.showModalDialog(
    HtmlService.createHtmlOutput(html).setWidth(640).setHeight(Math.min(120 + invalidos.length * 80, 500)),
    '🔗 Reconexão Strava (' + invalidos.length + ' atleta' + (invalidos.length > 1 ? 's' : '') + ')'
  );
}
