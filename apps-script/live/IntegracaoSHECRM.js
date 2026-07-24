/**
 * HIPERATIVO V3 — integração complementar e idempotente com o CRM SHE.
 *
 * Fluxo:
 * Formulário SHE -> CRM_ALUNOS -> HIPERATIVO V3 -> situação Strava ->
 * mensagem pronta no WhatsApp.
 *
 * Regras de segurança:
 * - nunca lê, apaga ou substitui refresh_token;
 * - cadastro existente vence por ID SHE, e-mail exato ou WhatsApp;
 * - dados esportivos já preenchidos só são completados quando estão vazios;
 * - o estado real do token vence a declaração feita no formulário;
 * - um acionador por envio substitui a varredura periódica de WhatsApp.
 */

var SHECRM_CFG_ = {
  spreadsheetId: '1Sn58dggyaalkNyVWWCAqx8U-xXyx8uDtaq3phLR7_GE',
  respostas: 'Respostas ao formulário 1',
  crm: 'CRM_ALUNOS',
  integracao: 'INTEGRACAO_SHE_V3',
  url: 'https://docs.google.com/spreadsheets/d/1Sn58dggyaalkNyVWWCAqx8U-xXyx8uDtaq3phLR7_GE/edit',
  timezone: 'America/Sao_Paulo'
};

var SHECRM_LOG_HEADERS_ = [
  'DATA_SYNC', 'ID_SHE', 'ID_ATLETA_V3', 'NOME', 'EMAIL_NORMALIZADO',
  'WHATSAPP_SHE', 'STATUS_VINCULO', 'ORIGEM_VINCULO', 'STRAVA_V3',
  'LINK_CRM', 'LINK_V3', 'OBSERVACAO'
];

function instalarIntegracaoSHECRMTempoReal() {
  var removidos = 0;
  ScriptApp.getProjectTriggers().forEach(function(t) {
    var fn = t.getHandlerFunction();
    if (fn === 'processarCadastroSHEEmTempoReal' ||
        fn === 'sincronizarFilaWhatsAppCadastros') {
      ScriptApp.deleteTrigger(t);
      removidos++;
    }
  });

  ScriptApp.newTrigger('processarCadastroSHEEmTempoReal')
    .forSpreadsheet(SHECRM_CFG_.spreadsheetId)
    .onFormSubmit()
    .create();

  var inicial = sincronizarCadastrosSHEPendentes();
  _log('SISTEMA', 'INFO', 'instalarIntegracaoSHECRMTempoReal',
    'Acionador SHE instalado; ' + removidos + ' acionador(es) antigo(s) removido(s).', '');
  return {
    ok: true,
    modo: 'onFormSubmit',
    varreduraWhatsApp15Min: false,
    sincronizacaoInicial: inicial
  };
}

/**
 * Executado somente quando o formulário SHE recebe uma resposta.
 * A pequena espera permite que o gatilho principal do CRM crie CRM_ALUNOS
 * antes da reconciliação. Se ainda não estiver pronto, novas tentativas
 * evitam o falso CRM_NAO_LOCALIZADO observado em produção.
 */
function processarCadastroSHEEmTempoReal(e) {
  var email = _sheEmailDoEvento_(e);
  if (!email) {
    _log('SISTEMA', 'AVISO', 'processarCadastroSHEEmTempoReal',
      'Envio SHE sem e-mail reconhecido; será reconciliado no ciclo de segurança.', '');
    return { ok: false, adiado: true, motivo: 'EMAIL_NAO_LOCALIZADO' };
  }

  Utilities.sleep(5000);
  for (var tentativa = 1; tentativa <= 3; tentativa++) {
    var resultado = _sheSincronizarPorEmail_(email);
    if (resultado && resultado.ok) {
      _sheAtualizarCamadasDerivadas_(resultado.athId, 'formulario_she');
      return resultado;
    }
    if (tentativa < 3) Utilities.sleep(5000);
  }

  _sheRegistrarFalhaTransitória_(email,
    'CRM ainda não criou a linha após três tentativas; ciclo diário fará a reconciliação.');
  return { ok: false, adiado: true, email: email, motivo: 'CRM_EM_PROCESSAMENTO' };
}

function sincronizarCadastrosSHEPendentes() {
  // Usa document lock para não disputar com o lock específico da fila
  // WhatsApp durante a mesma reconciliação.
  var lock = LockService.getDocumentLock();
  if (!lock.tryLock(20000)) return { ok: false, adiado: true };
  try {
    var origem = SpreadsheetApp.openById(SHECRM_CFG_.spreadsheetId);
    var shCrm = origem.getSheetByName(SHECRM_CFG_.crm);
    if (!shCrm || shCrm.getLastRow() < 2) {
      return { ok: true, encontrados: 0, sincronizados: 0 };
    }

    var dados = shCrm.getDataRange().getValues();
    var cab = _sheMapaCabecalho_(dados[0]);
    var declaracoes = _sheMapaDeclaracoesStrava_(origem);
    var contexto = _sheCriarContextoV3_();
    var sincronizados = 0;
    var erros = [];

    for (var i = 1; i < dados.length; i++) {
      var idShe = _sheValor_(dados[i], cab, 'ID_ALUNO');
      var email = _sheNormalizarEmail_(
        _sheValor_(dados[i], cab, 'EMAIL_NORMALIZADO') ||
        _sheValor_(dados[i], cab, 'EMAIL')
      );
      if (!idShe || !email) continue;
      try {
        var r = _sheSincronizarLinhaCRM_(origem, shCrm, i + 1, dados[i], cab,
          declaracoes[email] || {}, contexto);
        if (r.ok) sincronizados++;
      } catch (erro) {
        erros.push(idShe + ': ' + erro.message);
        _log('SISTEMA', 'AVISO', 'sincronizarCadastrosSHEPendentes',
          idShe + ': ' + erro.message, '');
      }
    }

    try { sincronizarFilaWhatsAppCadastros(); } catch (eFila) {
      erros.push('Fila WhatsApp: ' + eFila.message);
    }
    try {
      sincronizarVisoesAposInput_('crm_she_reconciliacao', '', {
        forcarGlobal: true,
        metricasJaAtualizadas: true
      });
    } catch (eVisual) {
      erros.push('Painel: ' + eVisual.message);
    }
    return {
      ok: erros.length === 0,
      encontrados: dados.length - 1,
      sincronizados: sincronizados,
      erros: erros
    };
  } finally {
    lock.releaseLock();
  }
}

function _sheSincronizarPorEmail_(email) {
  var origem = SpreadsheetApp.openById(SHECRM_CFG_.spreadsheetId);
  var shCrm = origem.getSheetByName(SHECRM_CFG_.crm);
  if (!shCrm || shCrm.getLastRow() < 2) return null;
  var dados = shCrm.getDataRange().getValues();
  var cab = _sheMapaCabecalho_(dados[0]);
  var alvo = _sheNormalizarEmail_(email);
  var declaracoes = _sheMapaDeclaracoesStrava_(origem);
  var contexto = _sheCriarContextoV3_();
  for (var i = dados.length - 1; i >= 1; i--) {
    var atual = _sheNormalizarEmail_(
      _sheValor_(dados[i], cab, 'EMAIL_NORMALIZADO') ||
      _sheValor_(dados[i], cab, 'EMAIL')
    );
    if (atual === alvo) {
      return _sheSincronizarLinhaCRM_(origem, shCrm, i + 1, dados[i], cab,
        declaracoes[alvo] || {}, contexto);
    }
  }
  return null;
}

function _sheCriarContextoV3_() {
  var ss = SpreadsheetApp.openById(_getSsId());
  var cad = ss.getSheetByName(H.SHEETS.CADASTRO);
  if (!cad) throw new Error('Aba CADASTRO não encontrada no V3.');
  var largura = Math.max(H.CAD.LINK_CRM, cad.getLastColumn());
  var valores = cad.getRange(1, 1, cad.getLastRow(), largura).getValues();
  var porIdShe = {}, porEmail = {}, porWhats = {};
  for (var i = 3; i < valores.length; i++) {
    var athId = String(valores[i][H.CAD.ID - 1] || '').trim().toUpperCase();
    if (!_isAthIdValido_(athId)) continue;
    var idShe = String(valores[i][H.CAD.ID_SHE - 1] || '').trim().toUpperCase();
    var email = _sheNormalizarEmail_(
      valores[i][H.CAD.EMAIL_NORM - 1] || valores[i][H.CAD.EMAIL - 1]);
    var whats = _sheChaveWhats_(
      valores[i][H.CAD.WHATS_NORM - 1] || valores[i][H.CAD.WHATS - 1]);
    if (idShe) porIdShe[idShe] = i + 1;
    if (email && !porEmail[email]) porEmail[email] = i + 1;
    if (whats && !porWhats[whats]) porWhats[whats] = i + 1;
  }
  return {
    ss: ss,
    cad: cad,
    valores: valores,
    largura: largura,
    porIdShe: porIdShe,
    porEmail: porEmail,
    porWhats: porWhats,
    ultimasAtividades: _sheMapaUltimasAtividades_(ss)
  };
}

function _sheSincronizarLinhaCRM_(origem, shCrm, linhaCrm, row, cab, declaracao, ctx) {
  var idShe = String(_sheValor_(row, cab, 'ID_ALUNO') || '').trim().toUpperCase();
  var nome = String(_sheValor_(row, cab, 'NOME_COMPLETO') || '').trim();
  var email = _sheNormalizarEmail_(
    _sheValor_(row, cab, 'EMAIL_NORMALIZADO') || _sheValor_(row, cab, 'EMAIL'));
  var whatsLocal = String(
    _sheValor_(row, cab, 'WHATSAPP_NORMALIZADO') || _sheValor_(row, cab, 'WHATSAPP') || ''
  ).replace(/\D/g, '');
  var whatsKey = _sheChaveWhats_(whatsLocal);

  var linhaV3 = ctx.porIdShe[idShe] || 0;
  var origemVinculo = linhaV3 ? 'ID_SHE' : '';
  if (!linhaV3 && email && ctx.porEmail[email]) {
    linhaV3 = ctx.porEmail[email];
    origemVinculo = 'EMAIL_EXATO';
  }
  if (!linhaV3 && whatsKey && ctx.porWhats[whatsKey]) {
    linhaV3 = ctx.porWhats[whatsKey];
    origemVinculo = 'WHATSAPP_EXATO';
  }

  var criado = false;
  if (!linhaV3) {
    criado = true;
    linhaV3 = ctx.cad.getLastRow() + 1;
    var novo = new Array(H.CAD.LINK_CRM).fill('');
    novo[H.CAD.ID - 1] = _sheGerarAthId_(idShe, ctx);
    ctx.cad.getRange(linhaV3, 1, 1, novo.length).setValues([novo]);
    ctx.valores[linhaV3 - 1] = novo;
    origemVinculo = 'CRIADO_PELO_SHE';
  }

  var atual = ctx.cad.getRange(linhaV3, 1, 1, H.CAD.LINK_CRM).getValues()[0];
  var athId = String(atual[H.CAD.ID - 1] || '').trim().toUpperCase();
  var emailAnterior = _sheNormalizarEmail_(atual[H.CAD.EMAIL - 1]);
  var whatsAnterior = _sheChaveWhats_(atual[H.CAD.WHATS - 1]);
  var conflitos = [];
  if (emailAnterior && email && emailAnterior !== email) conflitos.push('e-mail diverge');
  if (whatsAnterior && whatsKey && whatsAnterior !== whatsKey) conflitos.push('WhatsApp diverge');

  _shePreencherSeVazio_(atual, H.CAD.NOME, nome);
  _shePreencherSeVazio_(atual, H.CAD.EMAIL, email);
  _shePreencherSeVazio_(atual, H.CAD.WHATS, _sheWhatsComPais_(whatsLocal));
  _shePreencherSeVazio_(atual, H.CAD.NASC, _sheValor_(row, cab, 'DATA_NASCIMENTO'));
  _shePreencherSeVazio_(atual, H.CAD.NIVEL, _sheValor_(row, cab, 'NIVEL_CORRIDA'));
  _shePreencherSeVazio_(atual, H.CAD.OBJ, _sheValor_(row, cab, 'OBJETIVO_PRINCIPAL'));
  _shePreencherSeVazio_(atual, H.CAD.FREQ, _sheFreqSemanal_(_sheValor_(row, cab, 'TREINOS_SEMANA')));
  _shePreencherSeVazio_(atual, H.CAD.HORARIO, _sheValor_(row, cab, 'TEMPO_SESSAO'));
  _shePreencherSeVazio_(atual, H.CAD.SAUDE, _sheValor_(row, cab, 'ALERTA_SAUDE'));
  _shePreencherSeVazio_(atual, H.CAD.PLANO, _sheValor_(row, cab, 'PLANO'));
  _shePreencherSeVazio_(atual, H.CAD.CIDADE, _sheValor_(row, cab, 'CIDADE'));
  _shePreencherSeVazio_(atual, H.CAD.ESTADO, _sheValor_(row, cab, 'UF'));
  _shePreencherSeVazio_(atual, H.CAD.ORIGEM, _sheValor_(row, cab, 'ORIGEM_LEAD') || 'Formulário SHE');
  _shePreencherSeVazio_(atual, H.CAD.DATA_CAD, _sheValor_(row, cab, 'DATA_CADASTRO') || new Date());
  _shePreencherSeVazio_(atual, H.CAD.STATUS, _sheValor_(row, cab, 'STATUS') || 'Ativo');
  var enriquecimentoSaude = _sheEnriquecerCadastroSaudeSHE_(atual, declaracao);

  var tokenValido = _temRefreshTokenValido_(athId);
  var revogado = !!PropertiesService.getScriptProperties().getProperty('STRAVA_REVOGADO_' + athId);
  var declarado = _sheClassificarDeclaracaoStrava_(declaracao, _sheValor_(row, cab, 'STRAVA'));
  if (tokenValido && !revogado) atual[H.CAD.STRAVA_OK - 1] = 'Sim';
  else if (revogado) atual[H.CAD.STRAVA_OK - 1] = 'Pendente';
  else if (declarado === 'nao') atual[H.CAD.STRAVA_OK - 1] = 'Não';
  else if (!String(atual[H.CAD.STRAVA_OK - 1] || '').trim() ||
           _textoNormalizado_(atual[H.CAD.STRAVA_OK - 1]) === 'sim') {
    atual[H.CAD.STRAVA_OK - 1] = 'Pendente';
  }

  var statusStrava = tokenValido && !revogado ? 'Conectado' :
    (revogado ? 'Desconectado' : declarado === 'nao' ? 'Nao utiliza' : 'Pendente OAuth');
  var statusVinculo = criado ? 'CRIADO_E_VINCULADO' :
    (conflitos.length ? 'VINCULADO_COM_ATENCAO' : 'VINCULADO');
  var agora = new Date();
  var linkCrm = SHECRM_CFG_.url + '#gid=' + shCrm.getSheetId() + '&range=A' + linhaCrm;
  var linkV3 = 'https://docs.google.com/spreadsheets/d/' + _getSsId() +
    '/edit#gid=' + ctx.cad.getSheetId() + '&range=A' + linhaV3;

  atual[H.CAD.ID_SHE - 1] = idShe;
  atual[H.CAD.EMAIL_NORM - 1] = email;
  atual[H.CAD.WHATS_NORM - 1] = whatsLocal;
  atual[H.CAD.SHE_PREENCHIDO - 1] = 'Sim';
  atual[H.CAD.CLASS_SHE - 1] = _sheValor_(row, cab, 'CLASSIFICACAO_INTERNA');
  atual[H.CAD.DISP_RC - 1] = _sheValor_(row, cab, 'DISPONIBILIDADE_RC');
  atual[H.CAD.ULT_SYNC_SHE - 1] = agora;
  atual[H.CAD.STATUS_SYNC_SHE - 1] = statusVinculo;
  atual[H.CAD.LINK_CRM - 1] = linkCrm;
  ctx.cad.getRange(linhaV3, 1, 1, atual.length).setValues([atual]);
  ctx.cad.getRange(linhaV3, H.CAD.ULT_SYNC_SHE).setNumberFormat('dd/MM/yyyy HH:mm');
  _sheRegistrarAuditoriaSaudeV3_(
    ctx.ss,
    athId,
    enriquecimentoSaude.campos,
    enriquecimentoSaude.preservados
  );

  _sheAtualizarIndicesContexto_(ctx, linhaV3, idShe, email, whatsKey);
  _sheAtualizarRetornoCRM_(shCrm, linhaCrm, cab, {
    ID_ATLETA_V3: athId,
    STATUS_V3: String(atual[H.CAD.STATUS - 1] || 'Ativo'),
    STRAVA_STATUS_V3: statusStrava,
    ULTIMA_ATIVIDADE_V3: ctx.ultimasAtividades[athId] || '',
    ULTIMA_SYNC_V3: agora,
    STATUS_INTEGRACAO_V3: statusVinculo
  });

  var obs = (criado ? 'Novo atleta criado no V3 sem alterar tokens.' :
    'Vínculo confirmado; dados esportivos existentes foram preservados.') +
    (conflitos.length ? ' Atenção: ' + conflitos.join(' e ') + '; o V3 foi preservado.' : '') +
    (declarado === 'nao' && tokenValido ? ' Token válido encontrado; estado real conectado prevaleceu.' : '');
  _sheAtualizarLog_(origem, {
    data: agora, idShe: idShe, athId: athId, nome: nome, email: email,
    whats: whatsLocal, status: statusVinculo, origem: origemVinculo,
    strava: statusStrava, linkCrm: linkCrm, linkV3: linkV3, obs: obs
  });

  try { if (typeof supaGarantirAtleta === 'function') supaGarantirAtleta(athId); } catch (_) {}
  try {
    if (typeof atualizarStatusConexaoStravaAtleta === 'function') {
      atualizarStatusConexaoStravaAtleta(athId, {
        atualizarFila: false,
        origem: 'integracao_she'
      });
    }
  } catch (_) {}
  try {
    registrarFilaWhatsAppCadastro_({
      athId: athId,
      nome: atual[H.CAD.NOME - 1],
      email: atual[H.CAD.EMAIL - 1],
      whats: atual[H.CAD.WHATS - 1],
      dataCadastro: atual[H.CAD.DATA_CAD - 1],
      stravaOk: atual[H.CAD.STRAVA_OK - 1],
      stravaId: atual[H.CAD.STRAVA_ID - 1]
    });
  } catch (_) {}

  return {
    ok: true,
    athId: athId,
    idShe: idShe,
    criado: criado,
    statusVinculo: statusVinculo,
    strava: statusStrava
  };
}

function _sheAtualizarCamadasDerivadas_(athId, origem) {
  try { sincronizarFilaWhatsAppCadastros(); } catch (_) {}
  try {
    sincronizarVisoesAposInput_(origem || 'crm_she', athId, {
      permitirGlobal: false,
      metricasJaAtualizadas: true
    });
  } catch (_) {}
}

function _sheMapaDeclaracoesStrava_(ss) {
  var sh = ss.getSheetByName(SHECRM_CFG_.respostas);
  var mapa = {};
  if (!sh || sh.getLastRow() < 2) return mapa;
  var dados = sh.getDataRange().getValues();
  var cab = _sheMapaCabecalho_(dados[0]);
  for (var i = 1; i < dados.length; i++) {
    var email = _sheNormalizarEmail_(_sheValor_(dados[i], cab, 'E_MAIL'));
    if (!email) continue;
    mapa[email] = {
      usa: _sheValor_(dados[i], cab, 'VOCE_UTILIZA_O_STRAVA'),
      perfil: _sheValor_(dados[i], cab, 'LINK_PUBLICO_DO_SEU_PERFIL_NO_STRAVA'),
      cabecalho: cab,
      resposta: dados[i].slice()
    };
  }
  return mapa;
}

function _sheClassificarDeclaracaoStrava_(declaracao, perfilCrm) {
  var uso = _textoNormalizado_(declaracao && declaracao.usa || '');
  var perfil = String(declaracao && declaracao.perfil || perfilCrm || '').trim();
  if (uso.indexOf('ainda nao') >= 0 || uso === 'nao' || uso.indexOf('nao utiliz') >= 0) return 'nao';
  if (uso.indexOf('sim') >= 0 || /strava\.(com|app)|strava\.app\.link/i.test(perfil)) return 'sim';
  return 'desconhecido';
}

function _sheMapaUltimasAtividades_(ss) {
  var mapa = {};
  var sh = ss.getSheetByName(H.SHEETS.ATIVIDADES);
  if (!sh || sh.getLastRow() < 3) return mapa;
  var dados = sh.getRange(3, 1, sh.getLastRow() - 2,
    Math.max(H.ATIV.ATH_ID, H.ATIV.DATA)).getValues();
  dados.forEach(function(r) {
    var id = String(r[H.ATIV.ATH_ID - 1] || '').trim().toUpperCase();
    var data = r[H.ATIV.DATA - 1];
    if (!_isAthIdValido_(id) || !(data instanceof Date) || isNaN(data.getTime())) return;
    if (!mapa[id] || data > mapa[id]) mapa[id] = data;
  });
  return mapa;
}

function _sheAtualizarRetornoCRM_(sh, linha, cab, valores) {
  Object.keys(valores).forEach(function(nome) {
    var idx = cab[_sheNormHeader_(nome)];
    if (idx === undefined) return;
    sh.getRange(linha, idx + 1).setValue(valores[nome]);
    if (nome.indexOf('ULTIMA_') === 0) sh.getRange(linha, idx + 1).setNumberFormat('dd/MM/yyyy HH:mm');
  });
}

function _sheAtualizarLog_(ss, item) {
  var sh = ss.getSheetByName(SHECRM_CFG_.integracao);
  if (!sh) sh = ss.insertSheet(SHECRM_CFG_.integracao);
  var cabAtual = sh.getRange(1, 1, 1, SHECRM_LOG_HEADERS_.length).getValues()[0];
  if (cabAtual.join('|') !== SHECRM_LOG_HEADERS_.join('|')) {
    sh.getRange(1, 1, 1, SHECRM_LOG_HEADERS_.length).setValues([SHECRM_LOG_HEADERS_]);
  }
  var ultima = sh.getLastRow();
  var dados = ultima > 1 ? sh.getRange(2, 1, ultima - 1, 12).getValues() : [];
  var linha = 0;
  for (var i = dados.length - 1; i >= 0; i--) {
    var mesmoId = item.idShe && String(dados[i][1] || '').trim().toUpperCase() === item.idShe;
    var mesmoErroEmail = item.email &&
      _sheNormalizarEmail_(dados[i][4]) === item.email &&
      String(dados[i][6] || '').trim().toUpperCase() === 'ERRO';
    if (mesmoId || mesmoErroEmail) {
      linha = i + 2;
      break;
    }
  }
  var valores = [[
    item.data, item.idShe, item.athId, item.nome, item.email, item.whats,
    item.status, item.origem, item.strava, item.linkCrm, item.linkV3, item.obs
  ]];
  if (!linha) linha = sh.getLastRow() + 1;
  sh.getRange(linha, 1, 1, 12).setValues(valores);
  sh.getRange(linha, 1).setNumberFormat('dd/MM/yyyy HH:mm');
  sh.setFrozenRows(1);
}

function _sheRegistrarFalhaTransitória_(email, obs) {
  try {
    var ss = SpreadsheetApp.openById(SHECRM_CFG_.spreadsheetId);
    _sheAtualizarLog_(ss, {
      data: new Date(),
      idShe: '',
      athId: '',
      nome: '',
      email: _sheNormalizarEmail_(email),
      whats: '',
      status: 'AGUARDANDO_CRM',
      origem: 'FORM_SUBMIT',
      strava: '',
      linkCrm: SHECRM_CFG_.url,
      linkV3: 'https://docs.google.com/spreadsheets/d/' + _getSsId() + '/edit',
      obs: obs
    });
  } catch (_) {}
}

function diagnosticarIntegracaoSHECRM() {
  var ss = SpreadsheetApp.openById(SHECRM_CFG_.spreadsheetId);
  var crm = ss.getSheetByName(SHECRM_CFG_.crm);
  var v3 = SpreadsheetApp.openById(_getSsId()).getSheetByName(H.SHEETS.CADASTRO);
  var triggerCount = ScriptApp.getProjectTriggers().filter(function(t) {
    return t.getHandlerFunction() === 'processarCadastroSHEEmTempoReal';
  }).length;
  var crmRows = crm ? Math.max(0, crm.getLastRow() - 1) : 0;
  var v3Rows = v3 ? Math.max(0, v3.getLastRow() - 3) : 0;
  var ids = {};
  var duplicados = [];
  if (v3 && v3.getLastRow() >= 4) {
    v3.getRange(4, H.CAD.ID_SHE, v3.getLastRow() - 3, 1).getValues().forEach(function(r) {
      var id = String(r[0] || '').trim().toUpperCase();
      if (!id) return;
      if (ids[id]) duplicados.push(id);
      ids[id] = true;
    });
  }
  return {
    ok: !!crm && !!v3 && triggerCount === 1 && duplicados.length === 0,
    crmRows: crmRows,
    v3Rows: v3Rows,
    vinculadosSHE: Object.keys(ids).length,
    duplicadosIdSHE: duplicados,
    triggerTempoReal: triggerCount,
    tokensAlterados: false
  };
}

function _sheEmailDoEvento_(e) {
  if (!e) return '';
  var named = e.namedValues || {};
  var chaves = Object.keys(named);
  for (var i = 0; i < chaves.length; i++) {
    var h = _sheNormHeader_(chaves[i]);
    if (h === 'E_MAIL' || h === 'EMAIL') {
      var valor = named[chaves[i]];
      return _sheNormalizarEmail_(Array.isArray(valor) ? valor[0] : valor);
    }
  }
  if (e.values && e.range) {
    var sh = e.range.getSheet();
    var cab = _sheMapaCabecalho_(sh.getRange(1, 1, 1, e.values.length).getValues()[0]);
    return _sheNormalizarEmail_(_sheValor_(e.values, cab, 'E_MAIL'));
  }
  return '';
}

function _sheMapaCabecalho_(headers) {
  var mapa = {};
  (headers || []).forEach(function(h, i) {
    var chave = _sheNormHeader_(h);
    if (chave && mapa[chave] === undefined) mapa[chave] = i;
  });
  return mapa;
}

function _sheValor_(row, cab, nome) {
  var idx = cab[_sheNormHeader_(nome)];
  return idx === undefined ? '' : row[idx];
}

function _sheNormHeader_(valor) {
  return String(valor || '').trim().toUpperCase()
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .replace(/[^A-Z0-9]+/g, '_').replace(/^_+|_+$/g, '');
}

function _sheNormalizarEmail_(valor) {
  return String(valor || '').trim().toLowerCase();
}

function _sheChaveWhats_(valor) {
  var n = String(valor || '').replace(/\D/g, '');
  if (n.length > 11) n = n.slice(-11);
  return n;
}

function _sheWhatsComPais_(valor) {
  var n = String(valor || '').replace(/\D/g, '');
  if (!n) return '';
  if (n.length === 10 || n.length === 11) return '55' + n;
  return n;
}

function _sheFreqSemanal_(valor) {
  var n = String(valor || '').trim();
  return n && /^\d+$/.test(n) ? n + 'x por semana' : n;
}

function _shePreencherSeVazio_(row, col, valor) {
  if (!String(row[col - 1] || '').trim() && valor !== '' && valor !== null && valor !== undefined) {
    row[col - 1] = valor;
  }
}

function _sheGerarAthId_(idShe, ctx) {
  var numero = String(idShe || '').replace(/\D/g, '');
  var base = numero ? 'ATHSHE' + ('0000' + numero).slice(-4) :
    'ATHSHE' + Utilities.getUuid().replace(/-/g, '').slice(0, 8).toUpperCase();
  var candidato = base;
  var seq = 2;
  var usados = {};
  (ctx.valores || []).forEach(function(r) {
    var id = String(r && r[H.CAD.ID - 1] || '').trim().toUpperCase();
    if (id) usados[id] = true;
  });
  while (usados[candidato]) candidato = base + '_' + seq++;
  return candidato;
}

function _sheAtualizarIndicesContexto_(ctx, linha, idShe, email, whats) {
  if (idShe) ctx.porIdShe[idShe] = linha;
  if (email) ctx.porEmail[email] = linha;
  if (whats) ctx.porWhats[whats] = linha;
}
