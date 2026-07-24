/**
 * HIPERATIVO V3 — atualização Strava orientada a eventos.
 *
 * Strava -> Supabase Edge Function (validação/fila) -> doPost assinado ->
 * busca oficial da atividade -> RAW + ATIVIDADES + MÉTRICAS do atleta.
 *
 * Nenhuma rotina deste arquivo apaga ou substitui refresh_token. Em caso de
 * falha, o evento permanece no Supabase para reconciliação posterior.
 */

var RT_STRAVA_CALLBACK_URL_ =
  'https://korlpbclqgmqvpbrungc.supabase.co/functions/v1/strava-webhook';
var RT_WEBAPP_URL_ =
  'https://script.google.com/macros/s/AKfycbyNrmCjUxRYjUVjKMBPK7n_qFMknas2yfEXVciAMAIcOO1dr-9zH5haSuiGxMlIIO4Fqg/exec';
var RT_MAX_EVENT_AGE_S_ = 5 * 60;

/** Endpoint privado chamado pela Edge Function. */
function doPost(e) {
  var envelope = null;
  try {
    envelope = _rtValidarEnvelope_(e);
    var lock = LockService.getUserLock();
    lock.waitLock(25000);
    try {
      var resultado = _processarEventoStravaTempoReal_(envelope.event);
      return _rtJson_({ ok: true, resultado: resultado });
    } finally {
      try { lock.releaseLock(); } catch (_) {}
    }
  } catch (erro) {
    var mensagem = String(erro && erro.message || erro || 'Falha desconhecida').substring(0, 500);
    if (envelope && envelope.event) {
      try {
        var athFalha = _rtEncontrarAthId_(envelope.event.owner_id);
        if (athFalha && typeof registrarAlertaStravaOperacional_ === 'function') {
          registrarAlertaStravaOperacional_(athFalha, 'ERRO_TEMPORARIO', mensagem, true);
          if (typeof sincronizarVisoesAposInput_ === 'function') {
            sincronizarVisoesAposInput_('erro_webhook', athFalha, { permitirGlobal:false });
          }
        }
      } catch (_) {}
    }
    try { _log('SISTEMA', 'ERRO', 'doPost.webhookStrava', mensagem, ''); } catch (_) {}
    return _rtJson_({ ok: false, erro: mensagem });
  }
}

function _rtJson_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function _rtValidarEnvelope_(e) {
  var texto = e && e.postData && e.postData.contents ? String(e.postData.contents) : '';
  if (!texto || texto.length > 20000) throw new Error('Envelope ausente ou grande demais.');
  var envelope;
  try { envelope = JSON.parse(texto); } catch (_) { throw new Error('JSON inválido.'); }
  if (!envelope || envelope.source !== 'supabase-strava-webhook') {
    throw new Error('Origem não autorizada.');
  }

  var timestamp = Number(envelope.timestamp);
  var agora = Math.floor(Date.now() / 1000);
  if (!Number.isFinite(timestamp) || Math.abs(agora - timestamp) > RT_MAX_EVENT_AGE_S_) {
    throw new Error('Assinatura expirada.');
  }
  var nonce = String(envelope.nonce || '');
  if (!/^[0-9a-f-]{20,80}$/i.test(nonce)) throw new Error('Nonce inválido.');

  var segredo = PropertiesService.getScriptProperties()
    .getProperty('STRAVA_WEBHOOK_APPS_SCRIPT_SECRET');
  if (!segredo) throw new Error('Webhook ainda não configurado.');
  var evento = _rtValidarEvento_(envelope.event);
  var canonico = timestamp + '.' + nonce + '.' + JSON.stringify(evento);
  var esperado = _rtHmacHex_(segredo, canonico);
  if (!_rtIgualConstante_(esperado, String(envelope.signature || '').toLowerCase())) {
    throw new Error('Assinatura inválida.');
  }

  var cache = CacheService.getScriptCache();
  var chaveNonce = 'RT_NONCE_' + nonce;
  if (cache.get(chaveNonce)) throw new Error('Evento repetido.');
  cache.put(chaveNonce, '1', 600);
  return { event: evento };
}

function _rtValidarEvento_(evento) {
  if (!evento || typeof evento !== 'object') throw new Error('Evento ausente.');
  var tipo = String(evento.object_type || '');
  var aspecto = String(evento.aspect_type || '');
  var objetoId = Number(evento.object_id);
  var ownerId = Number(evento.owner_id);
  var subscriptionId = Number(evento.subscription_id);
  var eventTime = Number(evento.event_time);
  if (!/^(activity|athlete)$/.test(tipo) || !/^(create|update|delete)$/.test(aspecto) ||
      !Number.isSafeInteger(objetoId) || objetoId <= 0 ||
      !Number.isSafeInteger(ownerId) || ownerId <= 0 ||
      !Number.isSafeInteger(subscriptionId) || subscriptionId < 0 ||
      !Number.isSafeInteger(eventTime) || eventTime <= 0) {
    throw new Error('Estrutura do evento inválida.');
  }
  return {
    queue_id: Number(evento.queue_id) || 0,
    object_type: tipo,
    object_id: objetoId,
    aspect_type: aspecto,
    owner_id: ownerId,
    subscription_id: subscriptionId,
    event_time: eventTime,
    updates: evento.updates && typeof evento.updates === 'object' ? evento.updates : {}
  };
}

function _rtHmacHex_(segredo, mensagem) {
  return Utilities.computeHmacSha256Signature(mensagem, segredo).map(function(byte) {
    var n = byte < 0 ? byte + 256 : byte;
    return ('0' + n.toString(16)).slice(-2);
  }).join('');
}

function _rtIgualConstante_(a, b) {
  a = String(a || '');
  b = String(b || '');
  var diferenca = a.length ^ b.length;
  var tamanho = Math.max(a.length, b.length);
  for (var i = 0; i < tamanho; i++) {
    diferenca |= (a.charCodeAt(i % Math.max(1, a.length)) || 0) ^
      (b.charCodeAt(i % Math.max(1, b.length)) || 0);
  }
  return diferenca === 0;
}

function _processarEventoStravaTempoReal_(evento) {
  evento = _rtValidarEvento_(evento);
  var athId = _rtEncontrarAthId_(evento.owner_id);
  if (!athId) throw new Error('Conta Strava não vinculada a um ATH_ID.');

  if (evento.object_type === 'athlete') {
    var autorizado = evento.updates && evento.updates.authorized;
    if (evento.aspect_type === 'update' && (autorizado === 'false' || autorizado === false)) {
      _rtMarcarRevogacao_(athId, evento.owner_id);
      if (typeof registrarAlertaStravaOperacional_ === 'function') {
        registrarAlertaStravaOperacional_(athId, 'REVOGADO_CONFIRMADO',
          'Evento oficial de desautorização recebido da Strava.', true);
      }
      if (typeof sincronizarVisoesAposInput_ === 'function') {
        sincronizarVisoesAposInput_('revogacao_strava', athId, { permitirGlobal:false, metricasJaAtualizadas:true });
      }
      return { acao: 'revogacao_confirmada', ath_id: athId };
    }
    return { acao: 'ignorado', motivo: 'evento de atleta sem revogação' };
  }

  if (evento.aspect_type === 'delete') {
    var removidas = _rtExcluirAtividadeExata_(athId, evento.object_id);
    if (typeof sincronizarVisoesAposInput_ === 'function') {
      sincronizarVisoesAposInput_('exclusao_strava', athId, { metricasJaAtualizadas:true });
    }
    return { acao: 'exclusao', ath_id: athId, removidas: removidas };
  }

  var atividade = _rtBuscarAtividadeOficial_(athId, evento.owner_id, evento.object_id);
  if (!atividade) {
    var apagadas = _rtExcluirAtividadeExata_(athId, evento.object_id);
    if (typeof sincronizarVisoesAposInput_ === 'function') {
      sincronizarVisoesAposInput_('exclusao_404_strava', athId, { metricasJaAtualizadas:true });
    }
    return { acao: 'exclusao_confirmada_404', ath_id: athId, removidas: apagadas };
  }

  var nome = typeof _getNomeAtleta === 'function' ? _getNomeAtleta(athId) : athId;
  var novas = _gravarAtividades(athId, nome, [atividade], {
    forcarAtualizacao: evento.aspect_type === 'update'
  });
  if (typeof ordenarAtividadesMaisRecentes_ === 'function') ordenarAtividadesMaisRecentes_();
  if (typeof ordenarStravaRawMaisRecentes_ === 'function') ordenarStravaRawMaisRecentes_();
  _rtMarcarConectado_(athId, evento.owner_id);
  if (typeof resolverAlertasStravaOperacionais_ === 'function') resolverAlertasStravaOperacionais_(athId);
  if (typeof sincronizarVisoesAposInput_ === 'function') {
    sincronizarVisoesAposInput_('atividade_strava', athId, { metricasJaAtualizadas:true });
  }
  PropertiesService.getScriptProperties().setProperty('RT_GLOBAL_DIRTY', '1');
  return {
    acao: evento.aspect_type === 'update' ? 'atualizacao' : 'criacao',
    ath_id: athId,
    strava_id: String(evento.object_id),
    novas: novas
  };
}

function _rtEncontrarAthId_(ownerId) {
  var local = typeof _encontrarDonoStravaId_ === 'function'
    ? _encontrarDonoStravaId_(String(ownerId), '') : '';
  if (local) return String(local).trim().toUpperCase();
  if (!_supaConfigurado_()) return '';
  var resposta = _rtSupaFetch_('/rest/v1/tokens_strava?strava_id=eq.' +
    encodeURIComponent(String(ownerId)) + '&select=ath_id&limit=2', { method: 'get' });
  if (resposta.code !== 200) throw new Error('Falha ao consultar vínculo de atleta.');
  var linhas = JSON.parse(resposta.text || '[]');
  if (linhas.length > 1) throw new Error('Conta Strava vinculada a mais de um atleta.');
  return linhas.length ? String(linhas[0].ath_id || '').trim().toUpperCase() : '';
}

function _rtBuscarAtividadeOficial_(athId, ownerId, activityId) {
  var access = _getValidAccessToken(athId);
  if (!access) throw new Error('Access token indisponível; refresh token foi preservado.');
  var url = STRAVA_API_BASE + '/activities/' + encodeURIComponent(String(activityId));
  var resposta = _rtFetchStrava_(url, access);
  if (resposta.code === 401) {
    var renovado = _forcarRefreshAccessToken_(athId, access);
    if (!renovado) throw new Error('HTTP 401; refresh token preservado para nova tentativa.');
    resposta = _rtFetchStrava_(url, renovado);
  }
  if (resposta.code === 404) return null;
  if (resposta.code === 429) throw new Error('Limite Strava atingido; evento guardado para nova tentativa.');
  if (resposta.code !== 200) throw new Error('Strava respondeu HTTP ' + resposta.code + '.');
  var atividade = JSON.parse(resposta.text || '{}');
  var donoResposta = Number(atividade && atividade.athlete && atividade.athlete.id);
  if (!donoResposta || donoResposta !== Number(ownerId)) {
    throw new Error('Atividade não pertence à conta Strava esperada.');
  }
  return atividade;
}

function _rtFetchStrava_(url, accessToken) {
  var resp = UrlFetchApp.fetch(url, {
    method: 'get',
    headers: { Authorization: 'Bearer ' + accessToken },
    muteHttpExceptions: true
  });
  if (typeof _qRegistrarRate_ === 'function') _qRegistrarRate_(resp);
  return { code: resp.getResponseCode(), text: resp.getContentText() };
}

function _rtExcluirAtividadeExata_(athId, activityId) {
  var sid = String(activityId);
  var ss = SpreadsheetApp.openById(_getSsId());
  var removidas = 0;
  var shAtiv = ss.getSheetByName(H.SHEETS.ATIVIDADES);
  if (shAtiv && shAtiv.getLastRow() >= 3) {
    var dados = shAtiv.getRange(3, 1, shAtiv.getLastRow() - 2,
      Math.max(H.ATIV.STRAVA_ID, H.ATIV.ATH_ID)).getValues();
    for (var i = dados.length - 1; i >= 0; i--) {
      if (String(dados[i][H.ATIV.STRAVA_ID - 1] || '').trim() === sid &&
          String(dados[i][H.ATIV.ATH_ID - 1] || '').trim().toUpperCase() === athId) {
        shAtiv.deleteRow(i + 3);
        removidas++;
      }
    }
  }
  var shRaw = ss.getSheetByName(H.SHEETS.STRAVA_RAW);
  if (shRaw && shRaw.getLastRow() >= 2) {
    var raw = shRaw.getRange(2, 1, shRaw.getLastRow() - 1, 3).getValues();
    for (var r = raw.length - 1; r >= 0; r--) {
      if (String(raw[r][2] || '').trim() === sid &&
          String(raw[r][0] || '').trim().toUpperCase() === athId) {
        shRaw.deleteRow(r + 2);
      }
    }
    try { shRaw.hideSheet(); } catch (_) {}
  }
  if (removidas && typeof recalcularMetricasAposAtividade === 'function') {
    recalcularMetricasAposAtividade(athId);
  }
  PropertiesService.getScriptProperties().setProperty('RT_GLOBAL_DIRTY', '1');
  _log(athId, 'INFO', '_rtExcluirAtividadeExata_',
    'Exclusão Strava aplicada ao ID ' + sid + '; linhas=' + removidas, '');
  return removidas;
}

function _rtMarcarRevogacao_(athId, ownerId) {
  var props = PropertiesService.getScriptProperties();
  props.setProperty('STRAVA_REVOGADO_' + athId, String(Date.now()));
  _atualizarStatusCadastro(athId, false, String(ownerId));
  if (typeof supaAtualizarStravaOk === 'function') supaAtualizarStravaOk(athId, 'Reconectar');

  var ss = SpreadsheetApp.openById(_getSsId());
  var tokens = ss.getSheetByName(H.SHEETS.TOKENS);
  if (tokens && tokens.getLastRow() >= 2) {
    var dados = tokens.getRange(2, 1, tokens.getLastRow() - 1, H.TOK.STATUS).getValues();
    for (var i = 0; i < dados.length; i++) {
      if (String(dados[i][H.TOK.ATH_ID - 1] || '').trim().toUpperCase() === athId) {
        tokens.getRange(i + 2, H.TOK.STATUS).setValue('Revogado pela Strava — token preservado');
        break;
      }
    }
  }
  if (_supaConfigurado_()) {
    _rtSupaFetch_('/rest/v1/tokens_strava?ath_id=eq.' + encodeURIComponent(athId), {
      method: 'patch',
      headers: { Prefer: 'return=minimal' },
      payload: JSON.stringify({ status: 'Revogado pela Strava — token preservado', ult_atu: new Date().toISOString() })
    });
  }
  _log(athId, 'AVISO', '_rtMarcarRevogacao_',
    'A Strava confirmou a revogação; credenciais preservadas para auditoria.', '');
}

function _rtMarcarConectado_(athId, ownerId) {
  PropertiesService.getScriptProperties().deleteProperty('STRAVA_REVOGADO_' + athId);
  if (typeof _atualizarStatusCadastro === 'function') {
    _atualizarStatusCadastro(athId, true, String(ownerId));
  }
  if (typeof supaAtualizarStravaOk === 'function') supaAtualizarStravaOk(athId, 'Conectado');
}

/** Cria segredos internamente e configura o Supabase sem exibi-los. */
function configurarWebhookTempoRealSeguro() {
  if (!_supaConfigurado_()) throw new Error('Supabase não configurado no Apps Script.');
  var props = PropertiesService.getScriptProperties();
  var verify = props.getProperty('STRAVA_WEBHOOK_VERIFY_TOKEN');
  var segredo = props.getProperty('STRAVA_WEBHOOK_APPS_SCRIPT_SECRET');
  if (!verify) verify = _rtGerarSegredo_();
  if (!segredo) segredo = _rtGerarSegredo_() + _rtGerarSegredo_();
  props.setProperties({
    STRAVA_WEBHOOK_VERIFY_TOKEN: verify,
    STRAVA_WEBHOOK_APPS_SCRIPT_SECRET: segredo,
    STRAVA_WEBHOOK_CALLBACK_URL: RT_STRAVA_CALLBACK_URL_
  }, false);

  var resposta = _rtSupaFetch_('/rest/v1/strava_webhook_config?on_conflict=id', {
    method: 'post',
    headers: { Prefer: 'resolution=merge-duplicates,return=minimal' },
    payload: JSON.stringify({
      id: 1,
      verify_token: verify,
      apps_script_secret: segredo,
      apps_script_url: RT_WEBAPP_URL_,
      callback_url: RT_STRAVA_CALLBACK_URL_,
      modo: 'shadow',
      atualizado_at: new Date().toISOString()
    })
  });
  if (resposta.code >= 300) throw new Error('Supabase recusou a configuração: HTTP ' + resposta.code);
  _log('SISTEMA', 'INFO', 'configurarWebhookTempoRealSeguro',
    'Configuração privada criada em modo espelho; segredos não foram exibidos.', '');
  return { ok: true, modo: 'shadow', callback_url: RT_STRAVA_CALLBACK_URL_ };
}

function _rtGerarSegredo_() {
  return (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, '');
}

/** Registra a única assinatura permitida para o aplicativo Strava. */
function instalarWebhookStravaTempoReal() {
  var config = _rtBuscarConfig_();
  if (!config) {
    configurarWebhookTempoRealSeguro();
    config = _rtBuscarConfig_();
  }
  var cred = _getStravaAppCredentials_();
  if (!cred.clientId || !cred.clientSecret) throw new Error('Credenciais do aplicativo Strava ausentes.');
  var urlLista = STRAVA_API_BASE + '/push_subscriptions?client_id=' +
    encodeURIComponent(cred.clientId) + '&client_secret=' + encodeURIComponent(cred.clientSecret);
  var listaResp = UrlFetchApp.fetch(urlLista, { method: 'get', muteHttpExceptions: true });
  if (listaResp.getResponseCode() !== 200) {
    throw new Error('Não foi possível consultar assinaturas Strava: HTTP ' + listaResp.getResponseCode());
  }
  var lista = JSON.parse(listaResp.getContentText() || '[]');
  var atual = lista.length ? lista[0] : null;
  if (atual && String(atual.callback_url || '') !== RT_STRAVA_CALLBACK_URL_) {
    throw new Error('Já existe uma assinatura Strava (ID ' + Number(atual.id || 0) +
      ') em ' + String(atual.callback_url || 'endereço não informado') +
      '. Nada foi removido.');
  }
  var subscriptionId = atual ? Number(atual.id) : 0;
  if (!subscriptionId) {
    var criar = UrlFetchApp.fetch(STRAVA_API_BASE + '/push_subscriptions', {
      method: 'post',
      payload: {
        client_id: cred.clientId,
        client_secret: cred.clientSecret,
        callback_url: RT_STRAVA_CALLBACK_URL_,
        verify_token: PropertiesService.getScriptProperties().getProperty('STRAVA_WEBHOOK_VERIFY_TOKEN')
      },
      muteHttpExceptions: true
    });
    if (criar.getResponseCode() >= 300) {
      throw new Error('Strava recusou a assinatura: HTTP ' + criar.getResponseCode() + ' — ' +
        String(criar.getContentText() || '').substring(0, 250));
    }
    subscriptionId = Number(JSON.parse(criar.getContentText() || '{}').id);
  }
  if (!subscriptionId) throw new Error('A Strava não retornou o ID da assinatura.');
  _rtAtualizarConfig_({ subscription_id: subscriptionId, modo: 'shadow' });
  return { ok: true, subscription_id: subscriptionId, modo: 'shadow' };
}

/**
 * Substitui somente o proxy legado conhecido. Não aceita apagar callbacks
 * desconhecidos e tenta restaurar o legado se a criação nova falhar.
 */
function migrarWebhookLegadoParaSupabaseSeguro() {
  var LEGADO = 'https://webhook.ghiperativo.com.br';
  var config = _rtBuscarConfig_();
  if (!config) throw new Error('Configure o webhook seguro antes da migração.');
  var cred = _getStravaAppCredentials_();
  if (!cred.clientId || !cred.clientSecret) throw new Error('Credenciais Strava ausentes.');

  var listaUrl = STRAVA_API_BASE + '/push_subscriptions?client_id=' +
    encodeURIComponent(cred.clientId) + '&client_secret=' + encodeURIComponent(cred.clientSecret);
  var consulta = UrlFetchApp.fetch(listaUrl, { method: 'get', muteHttpExceptions: true });
  if (consulta.getResponseCode() !== 200) {
    throw new Error('Não foi possível consultar a assinatura existente.');
  }
  var itens = JSON.parse(consulta.getContentText() || '[]');
  if (!itens.length) return instalarWebhookStravaTempoReal();
  var atual = itens[0];
  var callbackAtual = String(atual.callback_url || '').replace(/\/$/, '');
  if (callbackAtual === RT_STRAVA_CALLBACK_URL_.replace(/\/$/, '')) {
    _rtAtualizarConfig_({ subscription_id: Number(atual.id), modo: 'shadow' });
    return { ok: true, migrado: false, subscription_id: Number(atual.id), modo: 'shadow' };
  }
  if (callbackAtual !== LEGADO) {
    throw new Error('Callback diferente do legado conhecido; nada foi removido: ' + callbackAtual);
  }

  var excluirUrl = STRAVA_API_BASE + '/push_subscriptions/' + encodeURIComponent(String(atual.id)) +
    '?client_id=' + encodeURIComponent(cred.clientId) +
    '&client_secret=' + encodeURIComponent(cred.clientSecret);
  var exclusao = UrlFetchApp.fetch(excluirUrl, { method: 'delete', muteHttpExceptions: true });
  if (exclusao.getResponseCode() !== 204) {
    throw new Error('A assinatura legada não foi removida: HTTP ' + exclusao.getResponseCode());
  }

  var verify = PropertiesService.getScriptProperties().getProperty('STRAVA_WEBHOOK_VERIFY_TOKEN');
  var criar = UrlFetchApp.fetch(STRAVA_API_BASE + '/push_subscriptions', {
    method: 'post',
    payload: {
      client_id: cred.clientId,
      client_secret: cred.clientSecret,
      callback_url: RT_STRAVA_CALLBACK_URL_,
      verify_token: verify
    },
    muteHttpExceptions: true
  });
  if (criar.getResponseCode() >= 300) {
    var restauracao = UrlFetchApp.fetch(STRAVA_API_BASE + '/push_subscriptions', {
      method: 'post',
      payload: {
        client_id: cred.clientId,
        client_secret: cred.clientSecret,
        callback_url: LEGADO,
        verify_token: verify
      },
      muteHttpExceptions: true
    });
    throw new Error('Novo callback recusado (HTTP ' + criar.getResponseCode() +
      '). Restauração do legado: HTTP ' + restauracao.getResponseCode() + '.');
  }

  var novoId = Number(JSON.parse(criar.getContentText() || '{}').id);
  if (!novoId) throw new Error('Nova assinatura criada sem ID retornado.');
  _rtAtualizarConfig_({ subscription_id: novoId, modo: 'shadow' });
  PropertiesService.getScriptProperties().setProperty(
    'STRAVA_WEBHOOK_SUBSCRIPTION_ID', String(novoId)
  );
  _log('SISTEMA', 'INFO', 'migrarWebhookLegadoParaSupabaseSeguro',
    'Assinatura legada ' + atual.id + ' substituída pela assinatura segura ' + novoId + '.', '');
  return {
    ok: true,
    migrado: true,
    anterior: Number(atual.id),
    subscription_id: novoId,
    modo: 'shadow'
  };
}

function ativarWebhookStravaTempoReal() {
  var config = _rtBuscarConfig_();
  if (!config || !Number(config.subscription_id)) {
    throw new Error('Instale e valide a assinatura antes de ativar.');
  }
  _rtAtualizarConfig_({ modo: 'ativo' });
  return { ok: true, modo: 'ativo', subscription_id: Number(config.subscription_id) };
}

function pausarWebhookStravaTempoReal() {
  _rtAtualizarConfig_({ modo: 'shadow' });
  return { ok: true, modo: 'shadow', observacao: 'Eventos continuam guardados sem alterar a planilha.' };
}

/** Teste completo e idempotente usando uma atividade já existente. */
function testarWebhookTempoRealSeguro() {
  var config = _rtBuscarConfig_();
  if (!config || !Number(config.subscription_id)) throw new Error('Assinatura Strava não instalada.');
  var amostra = _rtAtividadeAmostra_();
  var eventTime = Math.floor(Date.now() / 1000);
  var recebido = false;
  _rtAtualizarConfig_({ modo: 'ativo' });
  try {
    var evento = {
      object_type: 'activity',
      object_id: Number(amostra.activity.id),
      aspect_type: 'update',
      owner_id: Number(amostra.activity.athlete.id),
      subscription_id: Number(config.subscription_id),
      event_time: eventTime,
      updates: { title: 'teste_idempotente_hiperativo' }
    };
    var envio = UrlFetchApp.fetch(RT_STRAVA_CALLBACK_URL_, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(evento),
      muteHttpExceptions: true
    });
    if (envio.getResponseCode() >= 300) {
      throw new Error('Recepção do webhook falhou: HTTP ' + envio.getResponseCode());
    }
    recebido = true;
    for (var i = 0; i < 30; i++) {
      Utilities.sleep(1000);
      var fila = _rtConsultarEvento_(evento);
      if (fila && fila.status === 'processado') {
        return {
          ok: true,
          modo: 'ativo',
          ath_id: amostra.athId,
          strava_id: String(evento.object_id),
          fila: 'processado'
        };
      }
      if (fila && fila.status === 'falha') throw new Error('Fila registrou falha: ' + (fila.ultimo_erro || 'sem detalhe'));
    }
    throw new Error('O evento foi recebido, mas não concluiu dentro da janela de teste.');
  } catch (erro) {
    // Só pausa quando nem a recepção inicial foi aceita. Um evento já guardado
    // continua protegido pela fila e não deve derrubar o modo de produção por
    // simples atraso do teste.
    if (!recebido) _rtAtualizarConfig_({ modo: 'shadow' });
    throw erro;
  }
}

function _rtAtividadeAmostra_() {
  var ss = SpreadsheetApp.openById(_getSsId());
  var sh = ss.getSheetByName(H.SHEETS.TOKENS);
  if (!sh || sh.getLastRow() < 2) throw new Error('Nenhum token disponível para o teste.');
  var dados = sh.getRange(2, 1, sh.getLastRow() - 1, H.TOK.STATUS).getValues();
  for (var i = 0; i < dados.length; i++) {
    var athId = String(dados[i][H.TOK.ATH_ID - 1] || '').trim().toUpperCase();
    if (!_isAthIdValido_(athId)) continue;
    var access = _getValidAccessToken(athId);
    if (!access) continue;
    var resposta = _rtFetchStrava_(STRAVA_API_BASE + '/athlete/activities?per_page=1&page=1', access);
    if (resposta.code !== 200) continue;
    var atividades = JSON.parse(resposta.text || '[]');
    if (atividades.length && atividades[0].athlete && atividades[0].athlete.id) {
      return { athId: athId, activity: atividades[0] };
    }
  }
  throw new Error('Nenhuma atividade acessível para executar o teste seguro.');
}

/**
 * Repara somente linhas estruturais herdadas sem data válida. Alertas de pace,
 * FC ou duração potencialmente legítimos não são alterados automaticamente.
 */
function repararAtividadesEstruturaisViaStravaSeguro() {
  var ss = SpreadsheetApp.openById(_getSsId());
  var sh = ss.getSheetByName(H.SHEETS.ATIVIDADES);
  if (!sh || sh.getLastRow() < 3) return { reparadas: 0, falhas: 0 };
  var dados = sh.getRange(3, 1, sh.getLastRow() - 2, 34).getValues();
  var alvos = [];
  dados.forEach(function(row) {
    var status = String(row[H.ATIV.STATUS - 1] || '');
    var sid = String(row[H.ATIV.STRAVA_ID - 1] || '').trim();
    var athId = String(row[H.ATIV.ATH_ID - 1] || '').trim().toUpperCase();
    if (status.indexOf('data ausente/inválida') >= 0 && /^\d+$/.test(sid) && _isAthIdValido_(athId)) {
      alvos.push({ athId: athId, sid: sid });
    }
  });
  if (alvos.length > 10) throw new Error('Mais de 10 reparos estruturais pendentes; revisão manual exigida.');

  var reparadas = 0;
  var removidas = 0;
  var falhas = 0;
  alvos.forEach(function(alvo) {
    try {
      var token = _getTokenRow_(alvo.athId);
      var ownerId = Number(token && token.stravaId);
      if (!ownerId) throw new Error('ID Strava do atleta não encontrado.');
      var atividade = _rtBuscarAtividadeOficial_(alvo.athId, ownerId, Number(alvo.sid));
      if (!atividade) {
        removidas += _rtExcluirAtividadeExata_(alvo.athId, alvo.sid);
        return;
      }
      _gravarAtividades(alvo.athId, _getNomeAtleta(alvo.athId), [atividade], {
        forcarAtualizacao: true
      });
      reparadas++;
    } catch (erro) {
      falhas++;
      Logger.log('REPARO_ESTRUTURAL_FALHA ' + alvo.athId + ' ' + alvo.sid + ' ' +
        String(erro.message || erro).substring(0, 300));
      _log(alvo.athId, 'AVISO', 'repararAtividadesEstruturaisViaStravaSeguro',
        'Atividade ' + alvo.sid + ' não reparada: ' + erro.message, '');
    }
  });
  if (typeof ordenarAtividadesMaisRecentes_ === 'function') ordenarAtividadesMaisRecentes_();
  if (typeof ordenarStravaRawMaisRecentes_ === 'function') ordenarStravaRawMaisRecentes_();
  var resultado = {
    encontradas: alvos.length,
    reparadas: reparadas,
    removidas: removidas,
    falhas: falhas
  };
  Logger.log('REPARO_ESTRUTURAL ' + JSON.stringify(resultado));
  return resultado;
}

function _rtConsultarEvento_(evento) {
  var query = '/rest/v1/strava_eventos_webhook?subscription_id=eq.' + evento.subscription_id +
    '&object_id=eq.' + evento.object_id + '&event_time=eq.' + evento.event_time +
    '&select=id,status,ultimo_erro,tentativas&limit=1';
  var resposta = _rtSupaFetch_(query, { method: 'get' });
  if (resposta.code !== 200) return null;
  var linhas = JSON.parse(resposta.text || '[]');
  return linhas[0] || null;
}

/** Reprocessa somente falhas guardadas; chamada pela reconciliação diária. */
function processarFilaWebhookSupabase() {
  if (!_supaConfigurado_()) return { processados: 0, falhas: 0 };
  var resposta = _rtSupaFetch_('/rest/v1/strava_eventos_webhook?status=eq.falha' +
    '&tentativas=lt.5&select=id,subscription_id,object_type,object_id,aspect_type,owner_id,event_time,updates,tentativas' +
    '&order=recebido_at.asc&limit=10', { method: 'get' });
  if (resposta.code !== 200) throw new Error('Não foi possível ler a fila de segurança.');
  var eventos = JSON.parse(resposta.text || '[]');
  var processados = 0;
  var falhas = 0;
  eventos.forEach(function(evento) {
    var tentativas = Number(evento.tentativas || 0) + 1;
    try {
      _rtPatchEvento_(evento.id, { status: 'processando', tentativas: tentativas, ultimo_erro: null });
      _processarEventoStravaTempoReal_(evento);
      _rtPatchEvento_(evento.id, {
        status: 'processado', processado_at: new Date().toISOString(),
        ultimo_erro: null, proxima_tentativa: null
      });
      processados++;
    } catch (erro) {
      _rtPatchEvento_(evento.id, {
        status: 'falha', tentativas: tentativas,
        ultimo_erro: String(erro.message || erro).substring(0, 1000),
        proxima_tentativa: new Date(Date.now() + 15 * 60 * 1000).toISOString()
      });
      falhas++;
    }
  });
  return { processados: processados, falhas: falhas };
}

function _rtPatchEvento_(id, valores) {
  var resposta = _rtSupaFetch_('/rest/v1/strava_eventos_webhook?id=eq.' + Number(id), {
    method: 'patch', headers: { Prefer: 'return=minimal' }, payload: JSON.stringify(valores)
  });
  if (resposta.code >= 300) throw new Error('Falha ao atualizar fila: HTTP ' + resposta.code);
}

function diagnosticarWebhookTempoReal() {
  var config = _rtBuscarConfig_();
  var resposta = _rtSupaFetch_('/rest/v1/strava_eventos_webhook?select=status', { method: 'get' });
  var contagem = {};
  if (resposta.code === 200) {
    JSON.parse(resposta.text || '[]').forEach(function(r) {
      contagem[r.status] = (contagem[r.status] || 0) + 1;
    });
  }
  var resultado = {
    configurado: !!config,
    modo: config ? config.modo : 'ausente',
    subscription_id: config ? config.subscription_id : null,
    callback_url: config ? config.callback_url : RT_STRAVA_CALLBACK_URL_,
    eventos: contagem
  };
  Logger.log(JSON.stringify(resultado));
  return resultado;
}

/**
 * Evita que acionadores instalados por contas diferentes repitam o ciclo
 * pesado. Em caso de indisponibilidade do Supabase, falha aberto para manter a
 * reconciliação tradicional funcionando.
 */
function _rtDeveExecutarReconciliacaoCompleta_() {
  var config;
  try {
    config = _rtBuscarConfig_();
  } catch (_) {
    return true;
  }
  if (!config || config.modo !== 'ativo') return true;

  var lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return false;
  try {
    var props = PropertiesService.getScriptProperties();
    var chave = 'RT_ULTIMA_RECONCILIACAO_COMPLETA';
    var agora = Date.now();
    var ultima = Number(props.getProperty(chave) || 0);
    if (ultima && agora - ultima < 20 * 60 * 60 * 1000) return false;
    props.setProperty(chave, String(agora));
    return true;
  } finally {
    lock.releaseLock();
  }
}

/** Executar somente depois de testarWebhookTempoRealSeguro() retornar ok. */
function migrarTriggersParaTempoReal() {
  var config = _rtBuscarConfig_();
  if (!config || config.modo !== 'ativo') throw new Error('Webhook não está ativo.');
  var checagem = _rtSupaFetch_('/rest/v1/strava_eventos_webhook?status=eq.processado&select=id&limit=1', { method: 'get' });
  var processados = checagem.code === 200 ? JSON.parse(checagem.text || '[]') : [];
  if (!processados.length) throw new Error('Ainda não existe evento processado; gatilhos foram preservados.');

  var remover = {
    triggerImportacaoAutomatica: true,
    renovacaoProativaTokens: true,
    atualizarStatusStravaEmCadastro: true,
    sincronizarFilaWhatsAppCadastros: true
  };
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (remover[t.getHandlerFunction()]) ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('triggerImportacaoAutomatica')
    .timeBased().everyDays(1).atHour(3).create();
  return {
    ok: true,
    preservados: ['monitorarStravaOk', 'limparLogsAntigos'],
    reconciliacao: 'triggerImportacaoAutomatica diária às 03h'
  };
}

function _rtBuscarConfig_() {
  if (!_supaConfigurado_()) return null;
  var resposta = _rtSupaFetch_('/rest/v1/strava_webhook_config?id=eq.1' +
    '&select=id,apps_script_url,callback_url,subscription_id,modo&limit=1', { method: 'get' });
  if (resposta.code !== 200) throw new Error('Não foi possível ler a configuração do webhook.');
  var linhas = JSON.parse(resposta.text || '[]');
  return linhas[0] || null;
}

function _rtAtualizarConfig_(valores) {
  valores = Object.assign({}, valores, { atualizado_at: new Date().toISOString() });
  var resposta = _rtSupaFetch_('/rest/v1/strava_webhook_config?id=eq.1', {
    method: 'patch', headers: { Prefer: 'return=minimal' }, payload: JSON.stringify(valores)
  });
  if (resposta.code >= 300) throw new Error('Falha ao atualizar configuração: HTTP ' + resposta.code);
}

function _rtSupaFetch_(path, opcoes) {
  opcoes = opcoes || {};
  var headers = Object.assign({}, _supaHeaders_(), opcoes.headers || {});
  var requisicao = {
    method: opcoes.method || 'get',
    headers: headers,
    muteHttpExceptions: true
  };
  if (opcoes.payload !== undefined) requisicao.payload = opcoes.payload;
  var resp = UrlFetchApp.fetch(_supaUrl_() + path, requisicao);
  return { code: resp.getResponseCode(), text: resp.getContentText() };
}
