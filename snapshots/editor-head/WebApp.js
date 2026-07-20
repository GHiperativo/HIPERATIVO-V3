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

  // Link geral para atletas já cadastrados conectarem o Strava.
  // Exige ATH_ID + email e nunca sobrescreve um refresh token existente.
  if (p.conectar === 'true') {
    return _paginaConexaoStravaGeral();
  }

  // Página individual do atleta (?atleta=ATH_001)
  if (p.atleta) {
    return _paginaAtleta(p.atleta.toUpperCase().trim());
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

// ── Conexão Strava para atleta já cadastrado ────────────────────────────────
function _paginaConexaoStravaGeral() {
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
    #msg{display:none;margin-top:16px;padding:12px;border-radius:8px;line-height:1.45;text-align:center}
    .nota{font-size:12px;color:#778197;margin-top:18px}
  </style></head><body><main class="card">
    <div class="logo">⚡ HIPER<span>ATIVO</span></div>
    <h1>Conectar minha conta Strava</h1>
    <p>Use os mesmos dados do seu cadastro. A validação evita vincular a conta à pessoa errada.</p>
    <form id="form">
      <label for="athId">Código do atleta</label>
      <input id="athId" name="athId" placeholder="Ex.: ATH123456" required autocomplete="off">
      <label for="email">E-mail cadastrado</label>
      <input id="email" name="email" type="email" placeholder="seu@email.com" required autocomplete="email">
      <button id="btn" type="submit">CONECTAR COM STRAVA</button>
    </form>
    <div id="msg"></div>
    <p class="nota">Se sua conta já estiver conectada, o token atual será preservado e nenhuma nova autorização será solicitada.</p>
  </main><script>
    const form=document.getElementById('form'),btn=document.getElementById('btn'),msg=document.getElementById('msg');
    function mostrar(texto,ok){msg.textContent=texto;msg.style.display='block';msg.style.color=ok?'#176b36':'#a61b1b';
      msg.style.background=ok?'#e9f7ef':'#fdecec';}
    form.addEventListener('submit',function(ev){ev.preventDefault();btn.disabled=true;btn.textContent='VALIDANDO...';
      google.script.run.withSuccessHandler(function(resp){
        if(resp&&resp.ok&&resp.url){mostrar('Cadastro validado. Abrindo o Strava...',true);window.top.location.href=resp.url;return;}
        btn.disabled=false;btn.textContent='CONECTAR COM STRAVA';mostrar((resp&&resp.erro)||'Não foi possível validar os dados.',false);
      }).withFailureHandler(function(){btn.disabled=false;btn.textContent='CONECTAR COM STRAVA';
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
  for (let i = 2; i < cad.length; i++) {
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
    for (let i = 2; i < tok.length; i++) {
      const idTok = String(tok[i][H.TOK.ATH_ID - 1] || '').trim().toUpperCase();
      const refresh = String(tok[i][H.TOK.REFRESH - 1] || '').trim();
      if (idTok === athId && refresh) {
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
  // Preservar automações alheias ao fluxo. Remover apenas duplicatas das duas
  // rotinas gerenciadas aqui e recriá-las de forma determinística.
  const gerenciados = ['triggerImportacaoAutomatica', 'limparLogsAntigos'];
  ScriptApp.getProjectTriggers().forEach(t => {
    if (gerenciados.indexOf(t.getHandlerFunction()) >= 0) ScriptApp.deleteTrigger(t);
  });

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

// ── Página individual do atleta ─────────────────────────────
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
