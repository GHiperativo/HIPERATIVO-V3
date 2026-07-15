/**
 * HIPERATIVO V3 - Webhooks Strava
 * Recebe eventos rapidamente, coloca em fila e processa de forma assíncrona.
 */

const STRAVA_SUBSCRIPTIONS_URL = 'https://www.strava.com/api/v3/push_subscriptions';
const STRAVA_WEBHOOK_QUEUE = 'STRAVA_WEBHOOK_QUEUE';
const STRAVA_WEBHOOK_EVENT_PREFIX = 'STRAVA_WEBHOOK_EVENT_';

function _respostaJson_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj || {}))
    .setMimeType(ContentService.MimeType.JSON);
}

function _validarWebhookStrava_(p) {
  const props = PropertiesService.getScriptProperties();
  const esperado = props.getProperty('STRAVA_WEBHOOK_VERIFY_TOKEN') || '';
  const recebido = String(p['hub.verify_token'] || '');
  const challenge = String(p['hub.challenge'] || '');
  if (!esperado || recebido !== esperado || !challenge) {
    return _respostaJson_({ error: 'verification_failed' });
  }
  return _respostaJson_({ 'hub.challenge': challenge });
}

function doPost(e) {
  try {
    const raw = e && e.postData ? e.postData.contents : '';
    const evento = JSON.parse(raw || '{}');
    _enfileirarEventoWebhook_(evento);
  } catch (err) {
    _logErroSistema('doPost', 'Webhook não enfileirado: ' + err.message);
  }
  // O Strava exige resposta 200 em até dois segundos.
  return _respostaJson_({ ok: true });
}

function _obterFilaWebhook_() {
  const ss = SpreadsheetApp.openById(_getSsId());
  let sh = ss.getSheetByName(STRAVA_WEBHOOK_QUEUE);
  if (!sh) {
    sh = ss.insertSheet(STRAVA_WEBHOOK_QUEUE);
    sh.getRange(1, 1, 1, 12).setValues([[
      'Recebido em', 'Status', 'Tentativas', 'Tipo Objeto', 'Tipo Evento',
      'ID Objeto', 'ID Strava Atleta', 'ID Inscrição', 'Data Evento',
      'Atualizações JSON', 'Erro', 'Processado em'
    ]]);
    sh.setFrozenRows(1);
    sh.hideSheet();
  }
  return sh;
}

function _enfileirarEventoWebhook_(evento) {
  const tipos = ['activity', 'athlete'];
  const aspectos = ['create', 'update', 'delete'];
  if (tipos.indexOf(evento.object_type) < 0 || aspectos.indexOf(evento.aspect_type) < 0) {
    throw new Error('Evento Strava inválido.');
  }

  const subscriptionEsperada = PropertiesService.getScriptProperties()
    .getProperty('STRAVA_WEBHOOK_SUBSCRIPTION_ID');
  if (subscriptionEsperada && String(evento.subscription_id || '') !== subscriptionEsperada) {
    throw new Error('Evento recebido de inscrição Strava desconhecida.');
  }

  // Script Properties é mais rápido que abrir a planilha e ajuda o doPost
  // a responder dentro dos dois segundos exigidos pelo Strava.
  const chave = STRAVA_WEBHOOK_EVENT_PREFIX + [
    evento.subscription_id || '',
    evento.owner_id || '',
    evento.object_type || '',
    evento.object_id || '',
    evento.aspect_type || '',
    evento.event_time || Date.now()
  ].join('_');
  PropertiesService.getScriptProperties().setProperty(chave, JSON.stringify({
    recebido_em: new Date().toISOString(),
    tentativas: 0,
    evento: evento
  }));
}

function processarFilaWebhookStrava() {
  const props = PropertiesService.getScriptProperties();
  const todos = props.getProperties();
  const chaves = Object.keys(todos).filter(k => k.indexOf(STRAVA_WEBHOOK_EVENT_PREFIX) === 0).slice(0, 25);
  if (!chaves.length) return 'Fila vazia.';

  const sh = _obterFilaWebhook_();
  let processados = 0;
  chaves.forEach(chave => {
    let envelope;
    try {
      envelope = JSON.parse(todos[chave] || '{}');
      _processarEventoWebhook_(envelope.evento || {});
      _registrarEventoWebhook_(sh, envelope, 'Concluído', '');
      props.deleteProperty(chave);
    } catch (err) {
      envelope = envelope || { recebido_em: new Date().toISOString(), tentativas: 0, evento: {} };
      envelope.tentativas = Number(envelope.tentativas || 0) + 1;
      if (envelope.tentativas >= 5) {
        _registrarEventoWebhook_(sh, envelope, 'Erro definitivo', err.message);
        props.deleteProperty(chave);
      } else {
        props.setProperty(chave, JSON.stringify(envelope));
      }
      _logErroSistema('processarFilaWebhookStrava', err.message);
    }
    processados++;
  });
  return processados + ' evento(s) processado(s).';
}

function _registrarEventoWebhook_(sh, envelope, status, erro) {
  const evento = envelope.evento || {};
  sh.appendRow([
    envelope.recebido_em ? new Date(envelope.recebido_em) : new Date(),
    status, Number(envelope.tentativas || 0) + 1, evento.object_type || '', evento.aspect_type || '',
    String(evento.object_id || ''), String(evento.owner_id || ''), String(evento.subscription_id || ''),
    evento.event_time || '', JSON.stringify(evento.updates || {}), erro || '', new Date()
  ]);
}

function processarFilaWebhookStravaAgora() {
  const ui = SpreadsheetApp.getUi();
  try {
    ui.alert('Fila Strava', processarFilaWebhookStrava(), ui.ButtonSet.OK);
  } catch (err) {
    ui.alert('Erro ao processar fila Strava', err.message, ui.ButtonSet.OK);
  }
}

function _processarEventoWebhook_(evento) {
  if (evento.object_type === 'athlete' &&
      evento.aspect_type === 'update' &&
      String(evento.updates.authorized) === 'false') {
    _marcarConexaoRevogada_(evento.owner_id);
    return;
  }

  if (evento.object_type !== 'activity') return;
  const athId = _getAthIdPorStravaId_(evento.owner_id);
  if (!athId) throw new Error('Atleta Strava ' + evento.owner_id + ' não associado ao cadastro.');

  if (evento.aspect_type === 'delete') {
    _excluirAtividadeStrava_(evento.object_id, athId);
    return;
  }

  const atividade = _buscarAtividadeStrava_(athId, evento.object_id);
  _upsertAtividadeStrava_(athId, _getNomeAtleta(athId), atividade);
}

function _getAthIdPorStravaId_(stravaId) {
  const sh = SpreadsheetApp.openById(_getSsId()).getSheetByName(H.SHEETS.TOKENS);
  if (!sh) return '';
  const dados = sh.getDataRange().getValues();
  for (let i = 2; i < dados.length; i++) {
    if (String(dados[i][H.TOK.STRAVA_ID - 1] || '') === String(stravaId || '')) {
      return String(dados[i][H.TOK.ATH_ID - 1] || '').trim().toUpperCase();
    }
  }
  return '';
}

function _buscarAtividadeStrava_(athId, activityId) {
  const token = _getValidAccessToken(athId);
  const resp = UrlFetchApp.fetch(STRAVA_API_BASE + '/activities/' + encodeURIComponent(activityId), {
    headers: { Authorization: 'Bearer ' + token },
    muteHttpExceptions: true
  });
  _registrarChamadaStrava();
  if (resp.getResponseCode() !== 200) {
    throw new Error('Falha ao buscar atividade ' + activityId + ': HTTP ' + resp.getResponseCode());
  }
  return JSON.parse(resp.getContentText());
}

function _linhaAtividadeStrava_(athId, nomeAtleta, a, execId) {
  const distKm = (a.distance || 0) / 1000;
  const tempoS = a.moving_time || 0;
  const velMps = tempoS > 0 ? (a.distance || 0) / tempoS : 0;
  return [
    execId || ('ATIV_' + Utilities.getUuid().substring(0, 8).toUpperCase()),
    athId, nomeAtleta,
    a.start_date_local ? new Date(a.start_date_local) : '',
    _traduzirTipo(a.sport_type || a.type || ''), 'Strava', String(a.id || ''), a.name || '',
    tempoS, a.elapsed_time || 0, a.distance || 0, distKm, velMps, velMps * 3.6,
    _calcPace(tempoS, distKm), _calcPaceDecimal(tempoS, distKm),
    a.average_heartrate || '', a.max_heartrate || '', a.total_elevation_gain || '',
    a.calories || '', a.average_cadence || '', a.average_watts || '',
    a.map ? (a.map.summary_polyline || '') : '', new Date()
  ];
}

function _upsertAtividadeStrava_(athId, nomeAtleta, atividade) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const sh = SpreadsheetApp.openById(_getSsId()).getSheetByName(H.SHEETS.ATIVIDADES);
    if (!sh) throw new Error('Aba ATIVIDADES não encontrada.');
    const sid = String(atividade.id || '');
    const dados = sh.getDataRange().getValues();
    for (let i = 2; i < dados.length; i++) {
      if (String(dados[i][H.ATIV.STRAVA_ID - 1] || '') === sid) {
        sh.getRange(i + 1, 1, 1, 24).setValues([_linhaAtividadeStrava_(athId, nomeAtleta, atividade, dados[i][0])]);
        return 'Atualizada';
      }
    }
    sh.appendRow(_linhaAtividadeStrava_(athId, nomeAtleta, atividade, ''));
    return 'Criada';
  } finally {
    lock.releaseLock();
  }
}

function _excluirAtividadeStrava_(activityId, athId) {
  const sh = SpreadsheetApp.openById(_getSsId()).getSheetByName(H.SHEETS.ATIVIDADES);
  if (!sh) return;
  const dados = sh.getDataRange().getValues();
  for (let i = dados.length - 1; i >= 2; i--) {
    if (String(dados[i][H.ATIV.STRAVA_ID - 1] || '') === String(activityId)) {
      sh.deleteRow(i + 1);
      _log(athId, 'INFO', '_excluirAtividadeStrava_', 'Atividade Strava excluída: ' + activityId, '');
      return;
    }
  }
}

function _marcarConexaoRevogada_(stravaId) {
  const ss = SpreadsheetApp.openById(_getSsId());
  const tok = ss.getSheetByName(H.SHEETS.TOKENS);
  if (!tok) return;
  const dados = tok.getDataRange().getValues();
  for (let i = 2; i < dados.length; i++) {
    if (String(dados[i][H.TOK.STRAVA_ID - 1] || '') !== String(stravaId)) continue;
    const athId = String(dados[i][H.TOK.ATH_ID - 1] || '');
    tok.getRange(i + 1, H.TOK.ACCESS, 1, 3).clearContent();
    tok.getRange(i + 1, H.TOK.STATUS).setValue('Revogado');
    tok.getRange(i + 1, H.TOK.ULT_ATU).setValue(new Date());
    _atualizarStatusCadastro(athId, false, '');
    _log(athId, 'INFO', '_marcarConexaoRevogada_', 'Atleta revogou o acesso Strava.', '');
    return;
  }
}

function configurarWebhookStrava() {
  const ui = SpreadsheetApp.getUi();
  try {
    const props = PropertiesService.getScriptProperties();
    const clientId = props.getProperty('STRAVA_CLIENT_ID') || '';
    const secret = props.getProperty('STRAVA_CLIENT_SECRET') || '';
    const callbackUrl = props.getProperty('WEBAPP_URL') || '';
    if (!clientId || !secret || !callbackUrl) throw new Error('Configure Client ID, Client Secret e WEBAPP_URL primeiro.');
    _validarUrlWebApp(callbackUrl);

    const existentes = _consultarWebhookStrava_();
    if (existentes.length) {
      props.setProperty('STRAVA_WEBHOOK_SUBSCRIPTION_ID', String(existentes[0].id));
      const mesmoCallback = String(existentes[0].callback_url || '') === callbackUrl;
      ui.alert('Webhook já configurado',
        'Inscrição ativa: ' + existentes[0].id +
          '\nCallback: ' + existentes[0].callback_url +
          (mesmoCallback ? '' : '\n\n⚠️ O callback ativo é diferente da WEBAPP_URL atual. Remova e recrie a inscrição.'),
        ui.ButtonSet.OK);
      return;
    }

    const verifyToken = Utilities.getUuid().replace(/-/g, '');
    props.setProperty('STRAVA_WEBHOOK_VERIFY_TOKEN', verifyToken);
    const resp = UrlFetchApp.fetch(STRAVA_SUBSCRIPTIONS_URL, {
      method: 'post',
      payload: {
        client_id: clientId,
        client_secret: secret,
        callback_url: callbackUrl,
        verify_token: verifyToken
      },
      muteHttpExceptions: true
    });
    if (resp.getResponseCode() < 200 || resp.getResponseCode() >= 300) {
      throw new Error('Strava recusou o webhook: HTTP ' + resp.getResponseCode() + ' - ' + resp.getContentText());
    }
    const result = JSON.parse(resp.getContentText() || '{}');
    props.setProperty('STRAVA_WEBHOOK_SUBSCRIPTION_ID', String(result.id || ''));
    _obterFilaWebhook_();
    ui.alert('✅ Webhook Strava configurado', 'Inscrição criada: ' + result.id, ui.ButtonSet.OK);
  } catch (err) {
    _logErroSistema('configurarWebhookStrava', err.message);
    ui.alert('❌ Erro ao configurar webhook', err.message, ui.ButtonSet.OK);
  }
}

function _consultarWebhookStrava_() {
  const props = PropertiesService.getScriptProperties();
  const url = STRAVA_SUBSCRIPTIONS_URL +
    '?client_id=' + encodeURIComponent(props.getProperty('STRAVA_CLIENT_ID') || '') +
    '&client_secret=' + encodeURIComponent(props.getProperty('STRAVA_CLIENT_SECRET') || '');
  const resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  if (resp.getResponseCode() !== 200) throw new Error('Falha ao consultar webhook: HTTP ' + resp.getResponseCode());
  return JSON.parse(resp.getContentText() || '[]');
}

function consultarWebhookStrava() {
  const ui = SpreadsheetApp.getUi();
  try {
    const itens = _consultarWebhookStrava_();
    ui.alert('Status do webhook Strava',
      itens.length ? 'Ativo. Inscrição: ' + itens[0].id + '\nCallback: ' + itens[0].callback_url : 'Nenhuma inscrição ativa.',
      ui.ButtonSet.OK);
  } catch (err) {
    ui.alert('Erro', err.message, ui.ButtonSet.OK);
  }
}

function removerWebhookStrava() {
  const ui = SpreadsheetApp.getUi();
  try {
    const itens = _consultarWebhookStrava_();
    if (!itens.length) {
      ui.alert('Nenhuma inscrição webhook ativa.');
      return;
    }
    const confirma = ui.alert('Remover webhook Strava',
      'Remover a inscrição ' + itens[0].id + '? Os eventos deixarão de chegar até uma nova configuração.',
      ui.ButtonSet.YES_NO);
    if (confirma !== ui.Button.YES) return;

    const props = PropertiesService.getScriptProperties();
    const url = STRAVA_SUBSCRIPTIONS_URL + '/' + encodeURIComponent(itens[0].id) +
      '?client_id=' + encodeURIComponent(props.getProperty('STRAVA_CLIENT_ID') || '') +
      '&client_secret=' + encodeURIComponent(props.getProperty('STRAVA_CLIENT_SECRET') || '');
    const resp = UrlFetchApp.fetch(url, { method: 'delete', muteHttpExceptions: true });
    if (resp.getResponseCode() !== 204) {
      throw new Error('Falha ao remover webhook: HTTP ' + resp.getResponseCode() + ' - ' + resp.getContentText());
    }
    props.deleteProperty('STRAVA_WEBHOOK_SUBSCRIPTION_ID');
    ui.alert('✅ Webhook Strava removido.');
  } catch (err) {
    _logErroSistema('removerWebhookStrava', err.message);
    ui.alert('Erro', err.message, ui.ButtonSet.OK);
  }
}

function configurarAutomacaoStrava() {
  const handlers = ['processarFilaWebhookStrava', 'reconciliarAtividadesDiarias', 'limparLogsAntigos', 'triggerImportacaoAutomatica'];
  ScriptApp.getProjectTriggers().forEach(t => {
    if (handlers.indexOf(t.getHandlerFunction()) >= 0) ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('processarFilaWebhookStrava').timeBased().everyMinutes(5).create();
  ScriptApp.newTrigger('reconciliarAtividadesDiarias').timeBased().everyDays(1).atHour(3).create();
  ScriptApp.newTrigger('limparLogsAntigos').timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(6).create();
  _logInfoSistema('configurarAutomacaoStrava', 'Fila a cada 5 min + reconciliação diária às 03h.');
  try {
    SpreadsheetApp.getUi().alert('✅ Automação Strava configurada',
      'Webhooks serão processados a cada 5 minutos.\nReconciliação diária será executada por volta das 03h.',
      SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (_) {}
}

function reconciliarAtividadesDiarias() {
  const resultado = _importarTodosAtletas(2);
  _limparFilaWebhookAntiga_();
  _log('SISTEMA', 'INFO', 'reconciliarAtividadesDiarias', resultado, '');
  return resultado;
}

function _limparFilaWebhookAntiga_() {
  const sh = _obterFilaWebhook_();
  const dados = sh.getDataRange().getValues();
  const limite = new Date(Date.now() - 30 * 24 * 60 * 60 * 1000);
  for (let i = dados.length - 1; i >= 1; i--) {
    if (String(dados[i][1] || '') === 'Concluído' && dados[i][11] instanceof Date && dados[i][11] < limite) {
      sh.deleteRow(i + 1);
    }
  }
}
