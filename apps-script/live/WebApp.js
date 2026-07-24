/**
 * ==========================================================
 * HIPERATIVO V3 — WebApp.gs  (v4.1 — 04/06/2026)
 * Roteador principal doGet | Triggers | Rate Limit Strava
 * ==========================================================
 */

const HIPERATIVO_WEBAPP_URL_OFICIAL_ =
  'https://script.google.com/macros/s/AKfycbyNrmCjUxRYjUVjKMBPK7n_qFMknas2yfEXVciAMAIcOO1dr-9zH5haSuiGxMlIIO4Fqg/exec';

function corrigirUrlWebAppOficial() {
  PropertiesService.getScriptProperties().setProperty(
    'WEBAPP_URL', HIPERATIVO_WEBAPP_URL_OFICIAL_
  );
  _log('SISTEMA', 'INFO', 'corrigirUrlWebAppOficial',
    'WEBAPP_URL alinhada à implantação pública oficial', '');
  return HIPERATIVO_WEBAPP_URL_OFICIAL_;
}

// ── doGet: roteador principal ──────────────────────────────
function doGet(e) {
  const p = (e && e.parameter) ? e.parameter : {};

  // Callback OAuth do Strava
  if (p.code && p.state) {
    return _processarCallbackOAuth(e);
  }

  // Link geral para atletas já cadastrados conectarem o Strava.
  // Exige ATH_ID + email e nunca sobrescreve um refresh token existente.
  if (p.conectar === 'true') {
    return _paginaConexaoStravaGeral(p);
  }

  // Página individual do atleta (?atleta=ATH_001)
  if (p.atleta) {
    return _paginaAtleta(p.atleta.toUpperCase().trim());
  }

  // Ranking divertido público (?ranking=true) — só atletas com AZ != "Não"
  if (p.ranking === 'true') {
    return _paginaRankingPublico();
  }

  // Processar formulário de cadastro (POST via form)
  if (p.salvar === 'true') {
    return _processarFormCadastro(p);
  }

  // Exibir formulário de cadastro (cadastro.html externo)
  try {
    const tmpl = HtmlService.createTemplateFromFile('cadastro');
    tmpl.athId = p.athId || '';
    tmpl.utmSource = p.utm_source || '';
    tmpl.utmMedium = p.utm_medium || '';
    tmpl.utmCampaign = p.utm_campaign || '';
    tmpl.utmContent = p.utm_content || '';
    tmpl.refOrigem = p.ref || '';
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

// ── Conexão Strava para atleta já cadastrado ────────────────────────────────
function _paginaConexaoStravaGeral(p) {
  p = p || {};
  const athInicial = String(p.athId || '').trim().toUpperCase().replace(/[^A-Z0-9_-]/g, '');
  const emailInicial = String(p.email || '').trim().replace(/["'<>]/g, '');
  const html = `<!DOCTYPE html><html lang="pt-BR"><head>
  <meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
  <base target="_top"><title>Conectar Strava — Grupo Hiperativo</title>
  <style>
    *{box-sizing:border-box}body{margin:0;min-height:100vh;display:flex;align-items:center;justify-content:center;
    padding:24px;font-family:Arial,sans-serif;background:#f4f6f9;color:#1a2340}
    .card{width:100%;max-width:520px;background:#fff;border-radius:16px;padding:30px;box-shadow:0 8px 30px rgba(0,0,0,.12)}
    .logo{text-align:center;font-size:34px;font-weight:900;color:#1a3a8a}.logo span{color:#00a846}
    h1{text-align:center;font-size:24px;margin:18px 0 8px}p{line-height:1.5;color:#566078;text-align:center}
    label{display:block;margin:18px 0 6px;font-weight:700;color:#1a3a8a}
    input{width:100%;padding:13px;border:1px solid #cbd3e1;border-radius:9px;font-size:16px;text-transform:none}
    #athId{text-transform:uppercase}button{width:100%;margin-top:24px;padding:14px;border:0;border-radius:9px;
    background:#fc4c02;color:#fff;font-size:16px;font-weight:800;cursor:pointer}button:disabled{opacity:.55}
    #continuar{display:none;width:100%;margin-top:14px;padding:14px;border-radius:9px;text-align:center;
    background:#fc4c02;color:#fff;text-decoration:none;font-size:16px;font-weight:800;box-shadow:0 4px 12px rgba(252,76,2,.24)}
    #msg{display:none;margin-top:16px;padding:12px;border-radius:8px;line-height:1.45;text-align:center}
    .nota{font-size:12px;color:#778197;margin-top:18px}
  </style></head><body><main class="card">
    <div class="logo">⚡ HIPER<span>ATIVO</span></div>
    <h1>Conectar minha conta Strava</h1>
    <p>Use os mesmos dados do seu cadastro. A validação evita vincular a conta à pessoa errada.</p>
    <form id="form">
      <label for="athId">Código do atleta</label>
      <input id="athId" name="athId" value="${athInicial}" placeholder="Ex.: ATH123456" required autocomplete="off">
      <label for="email">E-mail cadastrado</label>
      <input id="email" name="email" type="email" value="${emailInicial}" placeholder="seu@email.com" required autocomplete="email">
      <button id="btn" type="submit">CONECTAR COM STRAVA</button>
    </form>
    <div id="msg"></div>
    <a id="continuar" href="#" target="_top" rel="noopener">ABRIR O STRAVA AGORA</a>
    <p class="nota">Se sua conta já estiver conectada, o token atual será preservado e nenhuma nova autorização será solicitada.</p>
  </main><script>
    const form=document.getElementById('form'),btn=document.getElementById('btn'),
      msg=document.getElementById('msg'),continuar=document.getElementById('continuar');
    let espera;
    function mostrar(texto,ok){msg.textContent=texto;msg.style.display='block';msg.style.color=ok?'#176b36':'#a61b1b';
      msg.style.background=ok?'#e9f7ef':'#fdecec';}
    function liberarFormulario(texto){clearTimeout(espera);btn.disabled=false;btn.textContent=texto||'CONECTAR COM STRAVA';}
    form.addEventListener('submit',function(ev){ev.preventDefault();continuar.style.display='none';
      btn.disabled=true;btn.textContent='VALIDANDO...';
      espera=setTimeout(function(){liberarFormulario();mostrar('A validação demorou mais que o esperado. Tente novamente.',false);},25000);
      google.script.run.withSuccessHandler(function(resp){
        clearTimeout(espera);
        if(resp&&resp.ok&&resp.url){
          liberarFormulario('VALIDADO ✓');
          mostrar('Cadastro validado. Toque no botão abaixo para abrir o Strava.',true);
          continuar.href=resp.url;
          continuar.style.display='block';
          document.getElementById('athId').readOnly=true;
          document.getElementById('email').readOnly=true;
          return;
        }
        liberarFormulario();mostrar((resp&&resp.erro)||'Não foi possível validar os dados.',false);
      }).withFailureHandler(function(){liberarFormulario();
        mostrar('Falha temporária. Tente novamente em instantes.',false);
      }).iniciarConexaoStrava({athId:document.getElementById('athId').value,email:document.getElementById('email').value});
    });
  </script></body></html>`;
  return HtmlService.createHtmlOutput(html)
    .setTitle('Conectar Strava — Grupo Hiperativo')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function iniciarConexaoStrava(p) {
  const athId = String((p && p.athId) || '').trim().toUpperCase();
  const email = String((p && p.email) || '').trim().toLowerCase();
  const idValido = typeof _isAthIdValido_ === 'function'
    ? _isAthIdValido_(athId)
    : /^ATH[A-Z0-9_-]{3,30}$/.test(athId);
  if (!idValido || !email || email.indexOf('@') < 1) {
    return { ok: false, erro: 'Confira o código do atleta e o e-mail cadastrado.' };
  }

  const ss = SpreadsheetApp.openById(_getSsId());
  const shCad = ss.getSheetByName(H.SHEETS.CADASTRO);
  if (!shCad) return { ok: false, erro: 'Cadastro temporariamente indisponível.' };

  const cad = shCad.getDataRange().getValues();
  let encontrado = false;
  for (let i = 0; i < cad.length; i++) {
    const idLinha = String(cad[i][H.CAD.ID - 1] || '').trim().toUpperCase();
    const emailLinha = String(cad[i][H.CAD.EMAIL - 1] || '').trim().toLowerCase();
    const status = String(cad[i][H.CAD.STATUS - 1] || '').trim().toLowerCase();
    if (idLinha === athId && emailLinha === email && status !== 'inativo') {
      encontrado = true;
      break;
    }
  }
  if (!encontrado) {
    return { ok: false, erro: 'Os dados não conferem com um cadastro ativo. Fale com o treinador.' };
  }

  const shTok = ss.getSheetByName(H.SHEETS.TOKENS);
  if (shTok) {
    const tok = shTok.getDataRange().getValues();
    // Algumas versões antigas da aba TOKENS gravaram o primeiro atleta na
    // linha 2. Percorrer todas as linhas e validar o ATH_ID preserva também
    // esse refresh token sem confundir títulos/cabeçalhos com atletas.
    for (let i = 0; i < tok.length; i++) {
      const idTok = String(tok[i][H.TOK.ATH_ID - 1] || '').trim().toUpperCase();
      const refresh = String(tok[i][H.TOK.REFRESH - 1] || '').trim();
      const idTokValido = typeof _isAthIdValido_ === 'function'
        ? _isAthIdValido_(idTok)
        : /^ATH[A-Z0-9_-]{3,30}$/.test(idTok);
      if (idTokValido && idTok === athId && refresh) {
        _log(athId, 'INFO', 'iniciarConexaoStrava', 'Conexão já existente; refresh token preservado.', '');
        return { ok: false, erro: 'Seu Strava já está conectado. Nenhuma nova autorização é necessária.' };
      }
    }
  }

  _log(athId, 'INFO', 'iniciarConexaoStrava', 'Cadastro validado para iniciar OAuth.', '');
  return { ok: true, url: _gerarUrlOAuth(athId) };
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

    // Agendar importação histórica para a próxima rodada da fila
    try { registrarAtletaParaHistorico(athId); } catch(qe) { /* não bloquear */ }

    // Email de confirmação
    try { _enviarEmailConexaoStrava(athId); } catch(ee) { /* não bloquear */ }

    try {
      if (typeof resolverAlertasStravaOperacionais_ === 'function') resolverAlertasStravaOperacionais_(athId);
      if (typeof sincronizarVisoesAposInput_ === 'function') {
        sincronizarVisoesAposInput_('oauth_strava', athId, { forcarGlobal:true });
      }
    } catch (ve) { _logErro('_processarCallbackOAuth.visual', athId, ve.message); }

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
function configurarTriggers(silencioso) {
  // Preservar automações alheias ao fluxo. Remover somente duplicatas das
  // rotinas essenciais do HIPERATIVO e recriá-las de forma determinística.
  const gerenciados = [
    'triggerImportacaoAutomatica',
    'renovacaoProativaTokens',
    'atualizarStatusStravaEmCadastro',
    'monitorarStravaOk',
    'processarCadastroSHEEmTempoReal',
    'sincronizarFilaWhatsAppCadastros',
    'limparLogsAntigos'
  ];
  ScriptApp.getProjectTriggers().forEach(t => {
    if (gerenciados.indexOf(t.getHandlerFunction()) >= 0) ScriptApp.deleteTrigger(t);
  });

  // Trigger de importação: a cada 3 horas
  ScriptApp.newTrigger('triggerImportacaoAutomatica')
    .timeBased()
    .everyHours(3)
    .create();

  // Renovação preventiva: mantém refresh/access tokens sincronizados.
  ScriptApp.newTrigger('renovacaoProativaTokens')
    .timeBased()
    .everyHours(4)
    .create();

  // Monitor diário com alerta consolidado e sem repetição excessiva.
  ScriptApp.newTrigger('monitorarStravaOk')
    .timeBased()
    .everyDays(1)
    .atHour(8)
    .create();

  // Cadastro SHE em tempo real; a própria chegada do formulário atualiza
  // cadastro, painel e mensagem. Não há varredura periódica de WhatsApp.
  ScriptApp.newTrigger('processarCadastroSHEEmTempoReal')
    .forSpreadsheet(SHECRM_CFG_.spreadsheetId)
    .onFormSubmit()
    .create();

  // Trigger de limpeza de logs: toda segunda-feira às 6h
  ScriptApp.newTrigger('limparLogsAntigos')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(6)
    .create();

  // Notificar no log
  _logInfoSistema('configurarTriggers',
    'Automações essenciais configuradas: SHE em tempo real, importação 3h, renovação 4h, monitor diário e limpeza semanal');

  // Feedback na UI se chamado manualmente
  try {
    if (silencioso === true) return diagnosticarTriggersEssenciais();
    SpreadsheetApp.getUi().alert(
      '✅ Triggers Configurados',
      'Importação: 3h\nRenovação de tokens: 4h\n' +
      'Monitor Strava: diário às 8h\nCadastro SHE: tempo real\nLimpeza de logs: segunda às 6h.\n\n' +
      'A rotina antiga de status permanece desativada.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch(e) { /* sem UI em trigger */ }
  return diagnosticarTriggersEssenciais();
}

/** Diagnóstico somente leitura das automações essenciais. */
function diagnosticarTriggersEssenciais() {
  const esperados = [
    'triggerImportacaoAutomatica',
    'renovacaoProativaTokens',
    'monitorarStravaOk',
    'processarCadastroSHEEmTempoReal',
    'limparLogsAntigos'
  ];
  const instalados = ScriptApp.getProjectTriggers().map(function(t) {
    return {
      funcao: t.getHandlerFunction(),
      origem: String(t.getTriggerSource()),
      evento: String(t.getEventType())
    };
  });
  const contagem = {};
  instalados.forEach(function(t) { contagem[t.funcao] = (contagem[t.funcao] || 0) + 1; });
  const status = esperados.map(function(fn) {
    return { funcao: fn, quantidade: contagem[fn] || 0, ok: contagem[fn] === 1 };
  });
  return { ok: status.every(function(s) { return s.ok; }), status: status, instalados: instalados };
}

/**
 * Wrapper com alerta visível — diagnosticarTriggersEssenciais() só
 * retorna um objeto e nunca aparecia na tela quando chamada pelo menu.
 * Não diz a cadência (3h vs diária) — a API de triggers do Apps Script
 * não expõe isso depois de criado — só confirma se há exatamente 1 de
 * cada trigger essencial ou se há zero/duplicado.
 */
function mostrarDiagnosticoTriggers() {
  const r = diagnosticarTriggersEssenciais();
  const linhas = r.status.map(function(s) {
    const icone = s.quantidade === 1 ? '✅' : s.quantidade === 0 ? '❌' : '⚠️';
    return icone + ' ' + s.funcao + ' — ' + s.quantidade + ' instalado(s)';
  });
  const gerenciados = r.status.map(function(s) { return s.funcao; });
  const extras = r.instalados.filter(function(t) { return gerenciados.indexOf(t.funcao) < 0; });
  let msg = linhas.join('\n');
  if (extras.length) {
    msg += '\n\nOutros triggers instalados:\n' + extras.map(function(t) { return '• ' + t.funcao; }).join('\n');
  }
  msg += '\n\n' + (r.ok
    ? 'Contagem OK (1 de cada). Isso NÃO confirma o intervalo — o Apps Script não permite consultar depois de criado.'
    : 'Atenção: algum item com 0 ou mais de 1 instalado.');
  try {
    SpreadsheetApp.getUi().alert('🩺 Diagnóstico de Triggers', msg, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch(_) { Logger.log(msg); }
  return r;
}

// ── Desativar todos os triggers ─────────────────────────────
function desativarTriggers(confirmacao) {
  if (confirmacao !== 'REMOVER_TODOS_OS_TRIGGERS') {
    throw new Error('Operação protegida. Use configurarTriggers() para reparar as automações essenciais.');
  }
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

// ── Página individual do atleta ─────────────────────────────
/**
 * Página pública do Ranking Divertido. Lê a mesma base de cálculo da
 * aba interna (_construirBaseRankings_, Dashboard.js) — nunca duplica
 * a lógica — e só exibe quem tem 👤 CADASTRO.AZ diferente de "Não".
 */
function _paginaRankingPublico() {
  try {
    const ss  = SpreadsheetApp.openById(_getSsId());
    const base = _construirBaseRankings_(ss);
    const ats = base.ats.filter(function(a) { return a.participar; });
    const medal = ['🥇', '🥈', '🥉'];
    const pos = function(i) { return medal[i] || (i + 1) + 'º'; };
    const cats = [];

    function addCat(id, titulo, emoji, itens, colDefs, topN) {
      topN = topN || 10;
      cats.push({
        id: id,
        titulo: titulo,
        emoji: emoji,
        cabecalho: colDefs.map(function(c) { return c.titulo; }),
        linhas: itens.slice(0, topN).map(function(it, i) {
          return colDefs.map(function(c) { return String(c.valor(it, i)); });
        })
      });
    }

    const colAtleta = { titulo: 'Atleta', valor: function(it) { return it.nome || it.athId; } };
    const colPos    = { titulo: '#', valor: function(it, i) { return pos(i); } };

    addCat('corrida', 'Volume de Corrida (30d)', '🏃',
      ats.filter(function(a){ return a.kmCorrida30 > 0; }).sort(function(a,b){ return b.kmCorrida30 - a.kmCorrida30; }),
      [ colPos, colAtleta, { titulo:'km', valor:function(it){ return it.kmCorrida30; } } ], 20);

    addCat('caminhada', 'Volume de Caminhada (30d)', '🚶',
      ats.filter(function(a){ return a.kmCaminhada30 > 0; }).sort(function(a,b){ return b.kmCaminhada30 - a.kmCaminhada30; }),
      [ colPos, colAtleta, { titulo:'km', valor:function(it){ return it.kmCaminhada30; } } ]);

    addCat('ciclismo', 'Volume de Ciclismo (30d)', '🚴',
      ats.filter(function(a){ return a.kmCiclismo30 > 0; }).sort(function(a,b){ return b.kmCiclismo30 - a.kmCiclismo30; }),
      [ colPos, colAtleta, { titulo:'km', valor:function(it){ return it.kmCiclismo30; } } ]);

    addCat('forca', 'Força & Funcional (30d)', '🏋️',
      ats.filter(function(a){ return a.forcaFuncional30 > 0; }).sort(function(a,b){ return b.forcaFuncional30 - a.forcaFuncional30; }),
      [ colPos, colAtleta, { titulo:'Sessões', valor:function(it){ return it.forcaFuncional30; } } ]);

    addCat('atividades', 'Número de Atividades (30d)', '🔢',
      ats.filter(function(a){ return a.treinos30 > 0; }).sort(function(a,b){ return b.treinos30 - a.treinos30; }),
      [ colPos, colAtleta, { titulo:'Atividades', valor:function(it){ return it.treinos30; } } ]);

    addCat('streak', 'Sequência Ativa', '🔥',
      ats.filter(function(a){ return a.streak > 0; }).sort(function(a,b){ return b.streak - a.streak; }),
      [ colPos, colAtleta, { titulo:'Dias seguidos', valor:function(it){ return it.streak; } } ]);

    addCat('madrugador', 'Madrugador(a)', '🌅',
      ats.filter(function(a){ return a.madrugada30 > 0; }).sort(function(a,b){ return b.madrugada30 - a.madrugada30; }),
      [ colPos, colAtleta, { titulo:'Treinos', valor:function(it){ return it.madrugada30; } } ]);

    addCat('coruja', 'Coruja', '🌙',
      ats.filter(function(a){ return a.coruja30 > 0; }).sort(function(a,b){ return b.coruja30 - a.coruja30; }),
      [ colPos, colAtleta, { titulo:'Treinos', valor:function(it){ return it.coruja30; } } ]);

    addCat('finde', 'Guerreiro(a) de Fim de Semana', '📅',
      ats.filter(function(a){ return a.fimDeSemana30 > 0; }).sort(function(a,b){ return b.fimDeSemana30 - a.fimDeSemana30; }),
      [ colPos, colAtleta, { titulo:'Treinos', valor:function(it){ return it.fimDeSemana30; } } ]);

    addCat('montanha', 'Rei/Rainha da Montanha', '🏔️',
      ats.filter(function(a){ return a.kmTotal >= 5 && a.elevPorKm > 0; }).sort(function(a,b){ return b.elevPorKm - a.elevPorKm; }),
      [ colPos, colAtleta, { titulo:'m/km', valor:function(it){ return it.elevPorKm; } } ]);

    addCat('comeback', 'Comeback do Mês', '🎽',
      ats.filter(function(a){ return a.crescimento30 !== null && a.crescimento30 >= 20 && a.km30 > 0; }).sort(function(a,b){ return b.crescimento30 - a.crescimento30; }),
      [ colPos, colAtleta, { titulo:'Crescimento', valor:function(it){ return it.crescimento30 + '%'; } } ]);

    addCat('leve', 'Regenerativo Consciente', '🐢',
      ats.filter(function(a){ return a.pseLeve30 > 0; }).sort(function(a,b){ return b.pseLeve30 - a.pseLeve30; }),
      [ colPos, colAtleta, { titulo:'Treinos leves', valor:function(it){ return it.pseLeve30; } } ]);

    const tmpl = HtmlService.createTemplateFromFile('ranking');
    tmpl.dataJson = JSON.stringify(cats);
    tmpl.totalParticipantes = ats.length;
    return tmpl.evaluate()
      .setTitle('Ranking Divertido — Hiperativo')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  } catch (err) {
    return _paginaMensagemWA('Ranking indisponível', 'Não foi possível carregar o ranking agora: ' + err.message, '#c00');
  }
}

function _paginaAtleta(athId) {
  try {
    const ss     = SpreadsheetApp.openById(_getSsId());
    const wsAtiv = ss.getSheetByName(H.SHEETS.ATIVIDADES);
    const wsMet  = ss.getSheetByName(H.SHEETS.METRICAS);
    const wsCad  = ss.getSheetByName(H.SHEETS.CADASTRO);

    // Dados do cadastro
    let nome = athId;
    if (wsCad) {
      const cadRows = wsCad.getDataRange().getValues();
      for (let i = 1; i < cadRows.length; i++) {
        if (String(cadRows[i][H.CAD.ID - 1] || '').trim().toUpperCase() === athId) {
          nome = String(cadRows[i][H.CAD.NOME - 1] || athId).trim();
          break;
        }
      }
    }

    // Métricas
    let vo2Str = '—', paceMedStr = '—', volSemStr = '—';
    let z1='—', z2='—', z3='—', z4='—', z5='—';
    if (wsMet) {
      const metRows = wsMet.getDataRange().getValues();
      for (let i = 2; i < metRows.length; i++) {
        if (String(metRows[i][H.MET.ATH_ID - 1] || '') === athId) {
          const pm  = Number(metRows[i][H.MET.PACE_MED  - 1]) || 0;
          const vo2 = Number(metRows[i][H.MET.VO2       - 1]) || 0;
          const vs  = Number(metRows[i][H.MET.VOL_SEM   - 1]) || 0;
          if (vo2 > 0) vo2Str     = String(vo2);
          if (vs  > 0) volSemStr  = String(vs);
          if (pm  > 0) {
            paceMedStr = _fmtPaceWA(pm);
            z1 = _fmtPaceWA(Math.round(pm * 1.28)) + '–' + _fmtPaceWA(Math.round(pm * 1.18));
            z2 = _fmtPaceWA(Math.round(pm * 1.18)) + '–' + _fmtPaceWA(Math.round(pm * 1.05));
            z3 = _fmtPaceWA(Math.round(pm * 0.98));
            z4 = _fmtPaceWA(Math.round(pm * 0.93));
            z5 = '< ' + _fmtPaceWA(Math.round(pm * 0.87));
          }
          break;
        }
      }
    }

    // Análise científica
    let tsbStr = '—', tsbCor = '', acwrStr = '—';
    let pctZ1z2 = 0, pctZ3 = 0, pctZ4z5 = 0, modeloZonas = '';
    let prescricao = [];
    let treinos28 = 0;
    try {
      const analise = calcularAnaliseAtleta(athId);
      if (analise) {
        tsbStr   = analise.tsb > 0 ? '+' + analise.tsb : String(analise.tsb);
        tsbCor   = analise.tsb < -20 ? 'vermelho' : analise.tsb < -10 ? 'laranja' : analise.tsb > 10 ? 'verde' : '';
        acwrStr  = analise.acwr.acwr ? String(analise.acwr.acwr) : '—';
        pctZ1z2  = analise.distribuicao.z1z2;
        pctZ3    = analise.distribuicao.z3;
        pctZ4z5  = analise.distribuicao.z4z5;
        modeloZonas = analise.distribuicao.modelo || '';
        prescricao  = analise.prescricao || [];
      }
    } catch(ae) { /* análise opcional */ }

    // Atividades recentes (últimas 20)
    const atividades = [];
    if (wsAtiv) {
      const atvRows = wsAtiv.getDataRange().getValues().slice(2)
        .filter(r => String(r[H.ATIV.ATH_ID - 1] || '') === athId && r[H.ATIV.DATA - 1] instanceof Date)
        .sort((a, b) => b[H.ATIV.DATA - 1] - a[H.ATIV.DATA - 1]);

      const d28 = new Date(Date.now() - 28 * 86400000);
      treinos28 = atvRows.filter(r => r[H.ATIV.DATA - 1] >= d28).length;

      const tagMap = { 'Corrida':'tag-corrida','Trail Run':'tag-corrida','Ciclismo':'tag-bike',
        'Natação':'tag-swim','Caminhada':'tag-cam' };

      atvRows.slice(0, 20).forEach(r => {
        const tipo  = String(r[H.ATIV.TIPO    - 1] || '');
        const frac  = Number(r[H.ATIV.MOV_S   - 1]) || 0;
        const mins  = frac > 0 && frac < 1 ? frac * 1440 : frac / 60;
        const h     = Math.floor(mins / 60);
        const m     = Math.round(mins % 60);
        const tempo = mins > 0 ? (h > 0 ? h + 'h' + String(m).padStart(2,'0') + 'm' : m + 'min') : '—';
        const dist  = Number(r[H.ATIV.DIST_KM - 1]) || 0;
        atividades.push({
          data:   Utilities.formatDate(r[H.ATIV.DATA - 1], Session.getScriptTimeZone(), 'dd/MM/yy'),
          tipo,
          tagCls: tagMap[tipo] || 'tag-outro',
          nome:   String(r[H.ATIV.NOME_ATIV  - 1] || '—'),
          dist:   dist > 0 ? dist + 'km' : '—',
          tempo,
          pace:   String(r[H.ATIV.PACE_FMT   - 1] || '—'),
          fc:     r[H.ATIV.FC_MED - 1] || '—',
          elev:   r[H.ATIV.ELEV   - 1] ? r[H.ATIV.ELEV - 1] + 'm' : '—',
          pse:    r[H.ATIV.PSE    - 1] || '—',
        });
      });
    }

    const ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');

    const tmpl = HtmlService.createTemplateFromFile('atleta');
    tmpl.athId       = athId;
    tmpl.nome        = nome;
    tmpl.vo2         = vo2Str;
    tmpl.paceMed     = paceMedStr;
    tmpl.volSem      = volSemStr;
    tmpl.tsbStr      = tsbStr;
    tmpl.tsbCor      = tsbCor;
    tmpl.acwrStr     = acwrStr;
    tmpl.treinos28   = treinos28;
    tmpl.z1          = z1;
    tmpl.z2          = z2;
    tmpl.z3          = z3;
    tmpl.z4          = z4;
    tmpl.z5          = z5;
    tmpl.pctZ1z2     = pctZ1z2;
    tmpl.pctZ3       = pctZ3;
    tmpl.pctZ4z5     = pctZ4z5;
    tmpl.modeloZonas = modeloZonas;
    tmpl.prescricao  = prescricao;
    tmpl.atividades  = atividades;
    tmpl.ts          = ts;

    return tmpl.evaluate()
      .setTitle(nome + ' — Hiperativo')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');

  } catch(err) {
    return _paginaMensagemWA('Atleta não encontrado',
      'Não foi possível carregar o painel de ' + athId + ': ' + err.message, '#c00');
  }
}

// ── Helper: formatar pace (seg → "M:SS") ───────────────────
function _fmtPaceWA(seg) {
  if (!seg || seg <= 0) return '—';
  const m = Math.floor(seg / 60);
  const s = seg % 60;
  return m + ':' + String(s).padStart(2, '0');
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
