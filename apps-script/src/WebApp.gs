/**
 * ==========================================================
 * HIPERATIVO V3 — WebApp.gs  (v4.1 — 04/06/2026)
 * Roteador principal doGet | Triggers | Rate Limit Strava
 * ==========================================================
 */

// ── doGet: roteador principal ──────────────────────────────
function doGet(e) {
  const p = (e && e.parameter) ? e.parameter : {};

  // Callback OAuth do Strava
  if (p.code && p.state) {
    return _processarCallbackOAuth(e);
  }

  // Processar formulário de cadastro (POST via form)
  if (p.salvar === 'true') {
    return _processarFormCadastro(p);
  }

  // Exibir formulário de cadastro (cadastro.html externo)
  try {
    const tmpl = HtmlService.createTemplateFromFile('cadastro');
    tmpl.athId = p.athId || '';
    return tmpl.evaluate()
      .setTitle('Cadastro — Grupo Hiperativo')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  } catch(err) {
    // Fallback: usar página HTML embutida do Strava.gs
    try {
      return HtmlService.createHtmlOutput(_paginaCadastro(p.athId || ''))
        .setTitle('Cadastro — Grupo Hiperativo')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    } catch(err2) {
      return _paginaMensagemWA('Erro', 'Erro ao carregar formulário: ' + err2.message, '#c00');
    }
  }
}

// ── Processar callback OAuth do Strava ─────────────────────
function _processarCallbackOAuth(e) {
  const p = (e && e.parameter) ? e.parameter : {};
  const code  = p.code  || '';
  const state = p.state || '';
  const error = p.error || '';

  if (error) {
    return _paginaMensagemWA('Strava não autorizado',
      'Você cancelou a autorização do Strava. Nenhum dado foi salvo.', '#e65100');
  }

  if (!code || !state) {
    return _paginaMensagemWA('Parâmetros inválidos',
      'Código ou estado ausente na resposta do Strava.', '#c00');
  }

  // state = athId + '__' + timestamp
  const athId = state.split('__')[0];
  if (!athId) {
    return _paginaMensagemWA('ID inválido', 'Identificador do atleta ausente.', '#c00');
  }

  try {
    // Trocar código por tokens (função do Strava.gs)
    const tokens = _trocarCodigoPorToken(athId, code);
    if (!tokens || tokens.error) {
      throw new Error(tokens ? (tokens.error_description || tokens.error) : 'Resposta vazia');
    }

    // Buscar perfil e atualizar cadastro
    try { importarPerfilAtleta(athId); } catch(pe) { /* não bloquear */ }

    // Email de confirmação
    try { _enviarEmailConexaoStrava(athId); } catch(ee) { /* não bloquear */ }

    return _paginaMensagemWA('✅ Strava conectado!',
      'Sua conta Strava foi vinculada com sucesso! A partir de agora suas atividades serão importadas automaticamente.',
      '#00897b');
  } catch(err) {
    _logErro('_processarCallbackOAuth', athId, err.message);
    return _paginaMensagemWA('Erro ao conectar Strava',
      'Não foi possível conectar: ' + err.message, '#c00');
  }
}

// ── Buscar atividades com rate limit (trigger automático) ──
function buscarAtividadesTodosAtletas() {
  const props   = PropertiesService.getScriptProperties();
  const limite15 = 160; // 80% de 200 req/15min
  const limiteDay = 1600; // 80% de 2000 req/dia

  // Contadores de rate limit
  const agora   = new Date();
  const hojeStr = Utilities.formatDate(agora, 'America/Sao_Paulo', 'yyyy-MM-dd');
  const janela  = Math.floor(agora.getTime() / (15 * 60 * 1000)); // janela de 15min

  let contDia   = parseInt(props.getProperty('RATE_DIA_' + hojeStr)   || '0');
  let contJanela= parseInt(props.getProperty('RATE_15M_' + janela)    || '0');

  // Verificar limites antes de iniciar
  if (contDia >= limiteDay) {
    _logErroSistema('buscarAtividadesTodosAtletas', 'Limite diário atingido: ' + contDia + '/' + limiteDay);
    return;
  }
  if (contJanela >= limite15) {
    _logErroSistema('buscarAtividadesTodosAtletas', 'Limite 15min atingido: ' + contJanela + '/' + limite15);
    return;
  }

  // Delegar para o Strava.gs que tem a lógica completa
  try {
    _importarTodosAtletas();
    _logInfoSistema('buscarAtividadesTodosAtletas', 'Ciclo de importação concluído');
  } catch(err) {
    _logErroSistema('buscarAtividadesTodosAtletas', 'Erro no ciclo: ' + err.message);
  }
}

// ── Configurar triggers automáticos ────────────────────────
function configurarTriggers() {
  // Remover todos os triggers existentes primeiro
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));

  // Trigger de importação: a cada 3 horas
  ScriptApp.newTrigger('triggerImportacaoAutomatica')
    .timeBased()
    .everyHours(3)
    .create();

  // Trigger de limpeza de logs: toda segunda-feira às 6h
  ScriptApp.newTrigger('limparLogsAntigos')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(6)
    .create();

  // Notificar no log
  _logInfoSistema('configurarTriggers', 'Triggers configurados: importação 3h + limpeza semanal');

  // Feedback na UI se chamado manualmente
  try {
    SpreadsheetApp.getUi().alert(
      '✅ Triggers Configurados',
      'Importação automática a cada 3 horas.\nLimpeza de logs toda segunda às 6h.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch(e) { /* sem UI em trigger */ }
}

// ── Desativar todos os triggers ─────────────────────────────
function desativarTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => ScriptApp.deleteTrigger(t));
  _logInfoSistema('desativarTriggers', triggers.length + ' triggers removidos');
  try {
    SpreadsheetApp.getUi().alert('✅ ' + triggers.length + ' trigger(s) desativado(s).');
  } catch(e) { /* sem UI */ }
}

// ── Limpar logs antigos (> 90 dias) ─────────────────────────
function limparLogsAntigos() {
  try {
    const ss     = SpreadsheetApp.openById(_getSsId());
    const shErro = ss.getSheetByName(H.SHEETS.ERROS);
    if (!shErro) return;

    const dados   = shErro.getDataRange().getValues();
    const limite  = new Date();
    limite.setDate(limite.getDate() - 90);

    // Encontrar linhas antigas (a partir da linha 3, pois linha 1=título, 2=cabeçalho)
    let linhasRemover = [];
    for (let i = dados.length - 1; i >= 2; i--) {
      const dataCell = dados[i][0];
      if (dataCell instanceof Date && dataCell < limite) {
        linhasRemover.push(i + 1); // 1-indexed
      }
    }

    linhasRemover.forEach(row => shErro.deleteRow(row));
    _logInfoSistema('limparLogsAntigos', linhasRemover.length + ' registros antigos removidos');
  } catch(err) {
    _logErroSistema('limparLogsAntigos', err.message);
  }
}

// ── Enviar email de confirmação de conexão Strava ──────────
function _enviarEmailConexaoStrava(athId) {
  try {
    const ss      = SpreadsheetApp.openById(_getSsId());
    const shCad   = ss.getSheetByName(H.SHEETS.CADASTRO);
    const dados   = shCad.getDataRange().getValues();

    let email = '', nome = '';
    for (let i = 1; i < dados.length; i++) {
      if ((dados[i][H.CAD.ID - 1] || '').toString().trim().toUpperCase() === athId.toUpperCase()) {
        email = dados[i][H.CAD.EMAIL - 1] || '';
        nome  = dados[i][H.CAD.NOME  - 1] || 'Atleta';
        break;
      }
    }

    if (!email) return;

    const assunto = '✅ Strava conectado ao HIPERATIVO!';
    const corpo   = '<h2 style="color:#1a3a8a">Parabéns, ' + nome + '!</h2>' +
      '<p>Sua conta Strava foi conectada com sucesso ao sistema HIPERATIVO.</p>' +
      '<p>A partir de agora, suas atividades serão importadas automaticamente ' +
      'e você poderá acompanhar seu progresso no painel.</p>' +
      '<hr><p style="font-size:12px;color:#666">Grupo Hiperativo — CABEÇA • CORAÇÃO • CORPO</p>';

    GmailApp.sendEmail(email, assunto, '', { htmlBody: corpo, name: 'Grupo Hiperativo' });
  } catch(err) {
    Logger.log('_enviarEmailConexaoStrava erro: ' + err.message);
  }
}

// ── Helpers internos ────────────────────────────────────────
function _getSsId() {
  const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (id) return id;
  return SpreadsheetApp.getActiveSpreadsheet().getId();
}

function _logInfoSistema(funcao, msg) {
  _logSistemaWA('INFO', funcao, msg);
}

function _logErroSistema(funcao, msg) {
  _logSistemaWA('ERRO', funcao, msg);
}

function _logSistemaWA(nivel, funcao, msg) {
  try {
    const ss    = SpreadsheetApp.openById(_getSsId());
    const shLog = ss.getSheetByName(H.SHEETS.ERROS);
    if (!shLog) return;
    shLog.appendRow([
      new Date(),
      nivel,
      'SYSTEM',
      funcao,
      msg,
      '',
      'Não',
      ''
    ]);
  } catch(e) {
    Logger.log('[' + nivel + '] ' + funcao + ': ' + msg);
  }
}

function _logErro(funcao, athId, msg) {
  try {
    const ss    = SpreadsheetApp.openById(_getSsId());
    const shLog = ss.getSheetByName(H.SHEETS.ERROS);
    if (!shLog) return;
    shLog.appendRow([new Date(), 'ERRO', athId || 'SYSTEM', funcao, msg, '', 'Não', '']);
  } catch(e) {
    Logger.log('ERRO ' + funcao + ': ' + msg);
  }
}

function _paginaMensagemWA(titulo, msg, cor) {
  const html = '<!DOCTYPE html><html><head><meta charset="UTF-8">' +
    '<meta name="viewport" content="width=device-width,initial-scale=1">' +
    '<style>body{font-family:Arial,sans-serif;display:flex;align-items:center;' +
    'justify-content:center;min-height:100vh;margin:0;background:#f5f5f5;}' +
    '.card{background:#fff;padding:40px;border-radius:12px;box-shadow:0 4px 20px rgba(0,0,0,.1);' +
    'max-width:480px;text-align:center;}' +
    'h2{color:' + (cor||'#1a3a8a') + ';margin-bottom:16px;}' +
    'p{color:#444;line-height:1.6;}' +
    '.logo{font-size:48px;margin-bottom:16px;}' +
    '</style></head><body><div class="card">' +
    '<div class="logo">⚡</div>' +
    '<h2>' + titulo + '</h2>' +
    '<p>' + msg + '</p>' +
    '<p style="margin-top:24px;font-size:13px;color:#888">GRUPO HIPERATIVO<br>CABEÇA • CORAÇÃO • CORPO</p>' +
    '</div></body></html>';
  return HtmlService.createHtmlOutput(html)
    .setTitle(titulo)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
