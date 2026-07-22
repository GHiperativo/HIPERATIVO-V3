/**
 * ══════════════════════════════════════════════════════════════════════
 * HIPERATIVO V3 — Strava.gs  (v3.2 — fluxo único cadastro+OAuth 04/06/2026)
 * OAuth2 Strava: link único cadastro → formulário → autorização → extração
 * ═══════════════════════════════════════════════════════════════════════
 */

// ── CONSTANTES STRAVA ───────────────────────────────────────────────────────
const STRAVA_AUTH_URL = 'https://www.strava.com/oauth/authorize';
const STRAVA_TOKEN_URL = 'https://www.strava.com/oauth/token';
const STRAVA_API_BASE = 'https://www.strava.com/api/v3';
const STRAVA_SCOPE = 'read,activity:read_all,profile:read_all';

// — 1. GERAR LINK SEGURO PARA ATLETAS JÁ CADASTRADOS ————————————————————————
function gerarLinkStrava() {
  const ui = SpreadsheetApp.getUi();
  try {
    const url = _gerarUrlConexaoStrava();
    ui.alert(
      '✅ Link geral de conexão Strava',
      'Envie este link somente para atletas já cadastrados:\n\n' + url +
      '\n\nO atleta confirma código e e-mail. Se já existir refresh token, o sistema preserva a conexão e não inicia nova autorização.',
      ui.ButtonSet.OK
    );
    _log('SISTEMA', 'INFO', 'gerarLinkStrava', 'Link geral seguro de conexão exibido', '');
  } catch (e) {
    ui.alert('❌ Erro ao gerar link', e.message, ui.ButtonSet.OK);
    _log('SISTEMA', 'ERRO', 'gerarLinkStrava', e.message, e.stack || '');
  }
}

function _gerarUrlCadastro(athId) {
  const props = PropertiesService.getScriptProperties();
  const webAppUrl = props.getProperty('WEBAPP_URL') || '';
  if (!webAppUrl) throw new Error('WEBAPP_URL não configurado. Implante o WebApp primeiro.');
  return webAppUrl + '?cadastro=true&athId=' + encodeURIComponent(athId);
}

function _gerarUrlConexaoStrava() {
  const webAppUrl = PropertiesService.getScriptProperties().getProperty('WEBAPP_URL') || '';
  if (!webAppUrl) throw new Error('WEBAPP_URL não configurado. Implante o WebApp primeiro.');
  return webAppUrl + '?conectar=true';
}

function _gerarUrlOAuth(athId) {
  const props = PropertiesService.getScriptProperties();
  const clientId = props.getProperty('STRAVA_CLIENT_ID') || '';
  const webAppUrl = props.getProperty('WEBAPP_URL') || '';
  if (!clientId) throw new Error('STRAVA_CLIENT_ID não configurado.');
  if (!webAppUrl) throw new Error('WEBAPP_URL não configurado. Implante o WebApp primeiro.');
  const state = athId + '__' + Date.now();
  const params = [
    'client_id=' + encodeURIComponent(clientId),
    'redirect_uri=' + encodeURIComponent(webAppUrl),
    'response_type=code',
    'approval_prompt=auto',
    'scope=' + encodeURIComponent(STRAVA_SCOPE),
    'state=' + encodeURIComponent(state)
  ].join('&');
  return STRAVA_AUTH_URL + '?' + params;
}

function _processarFormCadastro(p) {
  const athId = (p.athId || '').toUpperCase().trim();
  if (!athId) return HtmlService.createHtmlOutput(_paginaErro('ID inválido', 'Identificador do atleta ausente.'));

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(H.SHEETS.CADASTRO);
    if (!sheet) throw new Error('Aba CADASTRO não encontrada.');

    const data = sheet.getDataRange().getValues();
    let rowIdx = -1;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][H.CAD.ID - 1] || '').trim().toUpperCase() === athId) { rowIdx = i + 1; break; }
    }

    const agora = new Date();
    const agoraStr = agora.toLocaleString('pt-BR', { timeZone: 'America/Sao_Paulo' });
    const waLink = p.waLink || (p.whats ? 'https://wa.me/' + String(p.whats).replace(/\D/g, '') : '');
    const obsBase = p.obs || '';
    const obsText = waLink ? (waLink + (obsBase ? ' | ' + obsBase : '')) : obsBase;
    const usaStrava = String(p.usa_strava || 'Sim').trim().toLowerCase()
      .normalize('NFD').replace(/[\u0300-\u036f]/g, '') !== 'nao';
    const atribuicao = [
      p.origem || '',
      p.utm_source ? 'utm_source=' + p.utm_source : '',
      p.utm_medium ? 'utm_medium=' + p.utm_medium : '',
      p.utm_campaign ? 'utm_campaign=' + p.utm_campaign : '',
      p.utm_content ? 'utm_content=' + p.utm_content : '',
      p.ref ? 'ref=' + p.ref : ''
    ].filter(String).join(' | ');
    const origemFinal = atribuicao || 'Formulário Web';

    // MAPEAMENTO CORRETO: 26 valores → cols 2–27 (H.CAD)
    const vals = [
      p.nome || '',               // col  2  NOME
      p.email || '',               // col  3  EMAIL
      p.whats || '',               // col  4  WHATS
      p.nasc || '',               // col  5  NASC
      p.sexo || '',               // col  6  SEXO
      p.peso || '',               // col  7  PESO
      p.altura || '',               // col  8  ALTURA
      p.mod || p.modAgg || '',   // col  9  MOD
      p.nivel || '',               // col 10  NIVEL
      p.obj || p.objAgg || '',   // col 11  OBJ
      p.freq || '',               // col 12  FREQ
      p.horario || '',               // col 13  HORARIO
      p.saude || p.saudeAgg || '', // col 14  SAUDE
      p.lesao || '',               // col 15  LESAO
      p.med || '',               // col 16  MED
      p.prova || '',               // col 17  PROVA
      p.plano || '',               // col 18  PLANO
      p.cidade || '',               // col 19  CIDADE
      p.estado || '',               // col 20  ESTADO
      p.cpf || '',               // col 21  CPF
      p.origem || 'Link de Cadastro Web', // col 22 ORIGEM
      agoraStr,                      // col 23  DATA_CAD
      'Pendente',                    // col 24  STRAVA_OK
      '',                            // col 25  STRAVA_ID
      'Ativo',                       // col 26  STATUS
      obsText,                       // col 27  OBS
    ];

    if (rowIdx === -1) {
      sheet.appendRow([athId, ...vals]);
    } else {
      sheet.getRange(rowIdx, 2, 1, vals.length).setValues([vals]);
    }

    _log(athId, 'INFO', '_processarFormCadastro', 'Cadastro salvo: ' + (p.nome || 'N/A'), '');

    const oauthUrl = _gerarUrlOAuth(athId);
    return HtmlService.createHtmlOutput(_paginaRedirecionando(oauthUrl, p.nome || athId))
      .setTitle('Conectando ao Strava...');

  } catch (err) {
    _log(athId, 'ERRO', '_processarFormCadastro', err.message, err.stack || '');
    return HtmlService.createHtmlOutput(_paginaErro('Erro ao salvar cadastro', err.message));
  }
}

// ── 3b. SALVAR CADASTRO VIA AJAX (google.script.run) ─────────────────────────
// Retorna objeto simples. Gera ATH_ID automatico se nao vier no link.
// Envia email de confirmacao com link Strava apos salvar.
function salvarCadastroAjax(p) {
  const props = PropertiesService.getScriptProperties();
  const webAppUrl = props.getProperty('WEBAPP_URL') || '';
  const adminEmail = props.getProperty('ADMIN_EMAIL') || 'contato@ghiperativo.com.br';

  // Gerar ATH_ID se nao veio no link
  let athId = (p.athId || '').trim().toUpperCase();
  if (!athId) {
    athId = 'ATH' + String(Date.now()).slice(-6);
  }

  try {
    const ss = SpreadsheetApp.openById(_getSsId());
    const sheet = ss.getSheetByName(H.SHEETS.CADASTRO);
    if (!sheet) return { ok: false, erro: 'Aba CADASTRO nao encontrada. Execute o Setup primeiro.' };

    const data = sheet.getDataRange().getValues();
    let rowIdx = -1;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][H.CAD.ID - 1] || '').trim().toUpperCase() === athId) {
        rowIdx = i + 1; break;
      }
    }

    const agora = new Date().toLocaleString('pt-BR', { timeZone: 'America/Sao_Paulo' });
    const usaStrava = String(p.usa_strava || 'Sim').trim().toLowerCase()
      .normalize('NFD').replace(/[\u0300-\u036f]/g, '') !== 'nao';
    const atribuicao = [
      p.origem || '',
      p.utm_source ? 'utm_source=' + p.utm_source : '',
      p.utm_medium ? 'utm_medium=' + p.utm_medium : '',
      p.utm_campaign ? 'utm_campaign=' + p.utm_campaign : '',
      p.utm_content ? 'utm_content=' + p.utm_content : '',
      p.ref ? 'ref=' + p.ref : ''
    ].filter(String).join(' | ');
    const origemFinal = atribuicao || 'Formulário Web';
    // wa.me link vem normalizado do formulário; fallback: derivar do número
    const waLink = p.waLink || (p.whats ? 'https://wa.me/' + String(p.whats).replace(/\D/g, '') : '');
    const obsBase = p.obs || '';
    const obsText = waLink ? (waLink + (obsBase ? ' | ' + obsBase : '')) : obsBase;

    // MAPEAMENTO: 33 valores → cols 2–34 (H.CAD)
    const vals = [
      p.nome || '',               // col  2  NOME
      p.email || '',               // col  3  EMAIL
      p.whats || '',               // col  4  WHATS (+5561999999999)
      p.nasc || '',               // col  5  NASC
      p.sexo || '',               // col  6  SEXO
      p.peso || '',               // col  7  PESO
      p.altura || '',               // col  8  ALTURA
      p.mod || '',               // col  9  MOD
      p.nivel || '',               // col 10  NIVEL
      p.obj || '',               // col 11  OBJ
      p.freq || '',               // col 12  FREQ
      p.horario || '',               // col 13  HORARIO
      p.saude || '',               // col 14  SAUDE
      p.lesao || '',               // col 15  LESAO
      p.med || '',               // col 16  MED
      p.prova || '',               // col 17  PROVA
      p.plano || '',               // col 18  PLANO
      p.cidade || '',               // col 19  CIDADE
      p.estado || '',               // col 20  ESTADO
      p.cpf || '',               // col 21  CPF
      origemFinal,                         // col 22  ORIGEM + atribuição UTM
      agora,                              // col 23  DATA_CAD
      usaStrava ? 'Pendente' : 'Não',     // col 24  STRAVA_OK
      '',                                 // col 25  STRAVA_ID
      'Ativo',                            // col 26  STATUS
      obsText,                            // col 27  OBS (wa.me link + obs)
      p.emerg_nome || '',               // col 28  EMERG_NOME
      p.emerg_tel || '',               // col 29  EMERG_TEL
      p.emerg_rel || '',               // col 30  EMERG_REL
      p.parq || '',               // col 31  PAR_Q (7 respostas ex: "N/N/N/N/N/N/N")
      p.pr_tempos || '',               // col 32  PR_TEMPOS (ex: "5k:25:00|10k:52:00|21k:—|42k:—")
      p.assinatura || '',               // col 33  ASSINATURA (nome digitado)
      agora,                              // col 34  DATA_ASSINATURA
      // ── CRM Avançado — inicialmente vazios, preenchidos manualmente ──────
      p.instagram || '',               // col 35  INSTAGRAM
      p.prox_prova || '',               // col 36  PROX_PROVA
      p.data_prova || '',               // col 37  DATA_PROVA
      '',                                 // col 38  PLANO_PAG (manual)
      '',                                 // col 39  DATA_INICIO (manual)
      '',                                 // col 40  ULTIMA_AVAL (manual)
    ];

    if (rowIdx === -1) {
      sheet.appendRow([athId, ...vals]);
    } else {
      sheet.getRange(rowIdx, 2, 1, vals.length).setValues([vals]);
    }

    _log(athId, 'INFO', 'salvarCadastroAjax', 'Cadastro salvo: ' + (p.nome || 'N/A'), '');

    // Garante a linha-pai antes de qualquer token chegar ao Supabase.
    // Se o Supabase estiver indisponível, o cadastro e os tokens locais seguem preservados.
    try {
      if (typeof supaGarantirAtleta === 'function') supaGarantirAtleta(athId);
    } catch (eSupaCadastro) {
      _log(athId, 'AVISO', 'salvarCadastroAjax',
        'Cadastro salvo; sincronização imediata com Supabase ficou pendente', eSupaCadastro.message);
    }

    // Cria/atualiza a mensagem operacional para copiar e colar no WhatsApp.
    try {
      if (typeof registrarFilaWhatsAppCadastro_ === 'function') {
        registrarFilaWhatsAppCadastro_({
          athId: athId,
          nome: p.nome || '',
          email: p.email || '',
          whats: p.whats || '',
          dataCadastro: agora,
          stravaOk: usaStrava ? 'Pendente' : 'Não',
          stravaId: ''
        });
      }
    } catch (eFila) {
      _log(athId, 'AVISO', 'salvarCadastroAjax',
        'Cadastro salvo; fila WhatsApp será reparada pelo acionador automático', eFila.message);
    }

    // Gerar URL OAuth para retornar ao frontend (conexão Strava inline, sem depender de email)
    let oauthUrl = '';
    if (usaStrava) {
      try { oauthUrl = _gerarUrlOAuth(athId); } catch (oe) { /* credenciais não configuradas ainda */ }
    }

    // ── Email de boas-vindas ──────────────────────────────────────────────────
    const emailAtleta = p.email || '';
    const nomeAtleta = (p.nome || '').split(' ')[0] || 'Atleta';
    if (emailAtleta) {
      try {
        const assunto = 'Bem-vindo(a) ao Hiperativo! 🏃';
        const corpo = _htmlEmailConfirmacaoCadastro(nomeAtleta, athId, oauthUrl);
        MailApp.sendEmail({ to: emailAtleta, subject: assunto, htmlBody: corpo, replyTo: adminEmail });
        _log(athId, 'INFO', 'salvarCadastroAjax', 'Email boas-vindas enviado para ' + emailAtleta, '');
      } catch (eEmail) {
        _log(athId, 'AVISO', 'salvarCadastroAjax', 'Falha ao enviar email: ' + eEmail.message, '');
      }
    }

    // ── WhatsApp de boas-vindas (se API configurada) ──────────────────────
    const whatsAtleta = p.whats || '';
    if (whatsAtleta && oauthUrl) {
      try {
        const numero = whatsAtleta.replace(/\D/g, '');
        const msgWA =
          'Olá, ' + nomeAtleta + '! 👋\n\n' +
          'Cadastro recebido no *Grupo Hiperativo*! ✅\n\n' +
          'Seu código de atleta: *' + athId + '*\n\n' +
          '🔗 Conecte seu Strava agora para seu treinador acompanhar seus treinos automaticamente:\n' +
          oauthUrl + '\n\n' +
          'Qualquer dúvida é só chamar. Vamos juntos! 🏃⚡';
        _enviarWhatsApp(numero, msgWA, athId);
      } catch (eWA) {
        _log(athId, 'AVISO', 'salvarCadastroAjax', 'Falha ao enviar WhatsApp: ' + eWA.message, '');
      }
    }

    return { ok: true, athId: athId, nome: p.nome || '', oauthUrl: oauthUrl, usaStrava: usaStrava };

  } catch (err) {
    _log(athId, 'ERRO', 'salvarCadastroAjax', err.message, err.stack || '');
    return { ok: false, erro: 'Erro ao salvar: ' + err.message };
  }
}

// ── Template email confirmacao de cadastro ─────────────────────────────────────
function _htmlEmailConfirmacaoCadastro(nome, athId, linkStrava) {
  const blocoStrava = linkStrava ? `
  <hr style="border:none;border-top:1px solid #eee;margin:24px 0">
  <h3 style="color:#1a3a8a;margin:0 0 8px">🏃 Conecte seu Strava</h3>
  <p style="color:#444;line-height:1.6">Conectando sua conta, seu treinador poderá acompanhar seus treinos automaticamente.</p>
  <div style="text-align:center;margin:24px 0">
    <a href="${linkStrava}" style="background:#fc4c02;color:#fff;padding:14px 32px;border-radius:8px;text-decoration:none;font-weight:bold;font-size:16px;display:inline-block">⚡ Conectar meu Strava agora</a>
  </div>
  <p style="color:#888;font-size:13px;text-align:center">Você também pode conectar mais tarde pelo mesmo link.</p>` : `
  <hr style="border:none;border-top:1px solid #eee;margin:24px 0">
  <p style="color:#444;line-height:1.6;text-align:center">Seu cadastro foi concluído sem Strava. Você não receberá pedidos de conexão.</p>`;
  return `<!DOCTYPE html>
<html><head><meta charset="UTF-8"></head>
<body style="margin:0;padding:0;background:#f4f6f9;font-family:Arial,sans-serif">
<table width="100%" cellpadding="0" cellspacing="0">
<tr><td align="center" style="padding:32px 16px">
<table width="580" cellpadding="0" cellspacing="0" style="background:#fff;border-radius:12px;overflow:hidden;box-shadow:0 4px 20px rgba(0,0,0,.12)">
<tr><td style="background:linear-gradient(135deg,#1a3a8a,#0d2560);padding:32px;text-align:center">
  <div style="font-size:48px">⚡</div>
  <h1 style="color:#fff;margin:12px 0 4px;font-size:26px;letter-spacing:2px">HIPERATIVO</h1>
  <p style="color:#00c853;margin:0;font-size:12px;letter-spacing:3px">CABEÇA • CORAÇÃO • CORPO</p>
</td></tr>
<tr><td style="padding:32px">
  <h2 style="color:#1a3a8a;margin:0 0 12px">Olá, ${nome}! 🎉</h2>
  <p style="color:#444;line-height:1.6">Seu cadastro foi recebido com sucesso. Estamos felizes em ter você no <strong>Grupo Hiperativo</strong>!</p>
  <p style="color:#444;line-height:1.6">Seu código de atleta é:</p>
  <div style="background:#f4f6f9;border-radius:8px;padding:16px;text-align:center;margin:16px 0">
    <span style="font-size:22px;font-weight:bold;color:#1a3a8a;letter-spacing:3px">${athId}</span>
  </div>
  ${blocoStrava}
  <p style="color:#888;font-size:13px;text-align:center">Guarde seu código de atleta: <strong>${athId}</strong></p>
</td></tr>
<tr><td style="background:#1a3a8a;padding:16px;text-align:center">
  <p style="color:rgba(255,255,255,.6);font-size:12px;margin:0">Grupo Hiperativo | contato@ghiperativo.com.br</p>
</td></tr>
</table></td></tr></table>
</body></html>`;
}


// ── Wrapper público para google.script.run ────────────────────────────────────
function gerarUrlOAuth(athId) {
  try { return _gerarUrlOAuth(athId); } catch (e) { return ''; }
}

// ── 4. TROCA CÓDIGO POR TOKEN ─────────────────────────────────────────────────

function _backupToken(athId, refreshToken, accessToken, expiresAt) {
  try {
    const props = PropertiesService.getScriptProperties();
    if (refreshToken) props.setProperty('TOK_REFRESH_' + athId, refreshToken);
    if (accessToken) props.setProperty('TOK_ACCESS_' + athId, accessToken);
    if (expiresAt) props.setProperty('TOK_EXPIRES_' + athId, String(expiresAt));
  } catch (e) { Logger.log('_backupToken erro: ' + e.message); }
}

// ── 7. IMPORTAR ATIVIDADES ────────────────────────────────────────────────────
function importarAtividades() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt('🏃 Importar Atividades', 'ID do atleta (vazio = todos):', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const athId = (r.getResponseText() || '').trim().toUpperCase();
  try {
    let msg;
    if (athId) {
      const props = PropertiesService.getScriptProperties();
      props.setProperty('Q_HIST_' + athId, '0');
      props.deleteProperty('Q_HIST_DONE_' + athId);
      const lote = _importarHistoricoPaginado(athId, 0);
      if (typeof ordenarAtividadesMaisRecentes_ === 'function') ordenarAtividadesMaisRecentes_();
      if (lote.concluido) props.deleteProperty('Q_HIST_' + athId);
      else props.setProperty('Q_HIST_' + athId, String(lote.ultimaPagina));
      msg = '✅ ' + lote.novas + ' novas atividades importadas para ' + athId +
        (lote.concluido ? '. Histórico concluído.' : '. O restante continuará pela fila automática.');
    } else {
      const ag = agendarHistoricoCompletoTodos(true);
      processarFilaStrava();
      msg = '✅ Histórico completo agendado para ' + ag.total + ' atletas. A fila continuará até a última página.';
    }
    ui.alert('📊 Resultado', msg, ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('❌ Erro', e.message, ui.ButtonSet.OK);
    _log(athId || 'TODOS', 'ERRO', 'importarAtividades', e.message, e.stack || '');
  }
}



function _importarAtividadesAtleta(athId, paginas) {
  const accessToken = _getValidAccessToken(athId);
  const nomeAtleta = _getNomeAtleta(athId);
  paginas = paginas || 3;
  let todas = [];
  for (let pg = 1; pg <= paginas; pg++) {
    if (typeof _qTemCapacidade_ === 'function' && !_qTemCapacidade_(PropertiesService.getScriptProperties())) break;
    const url = STRAVA_API_BASE + '/athlete/activities?per_page=100&page=' + pg;
    const resp = UrlFetchApp.fetch(url, {
      headers: { Authorization: 'Bearer ' + accessToken },
      muteHttpExceptions: true
    });
    if (typeof _qRegistrarRate_ === 'function') _qRegistrarRate_(resp);
    if (resp.getResponseCode() === 429 || resp.getResponseCode() !== 200) break;
    const page = JSON.parse(resp.getContentText());
    console.log('[IMPORT] pg=' + pg + ' atividades=' + page.length + ' total=' + todas.length);
    if (!page.length) break;
    todas = todas.concat(page);
  }
  const novas = _gravarAtividades(athId, nomeAtleta, todas);
  if (typeof ordenarAtividadesMaisRecentes_ === 'function') ordenarAtividadesMaisRecentes_();
  return novas;
}

function _gravarAtividades(athId, nomeAtleta, atividades) {
  if (!atividades.length) return 0;
  if (typeof _garantirEstruturaAtividadesUnificada_ === 'function') {
    _garantirEstruturaAtividadesUnificada_(false);
  }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(H.SHEETS.ATIVIDADES);
  if (!sheet) throw new Error('Aba ATIVIDADES não encontrada.');

  // A resposta original é preservada primeiro. Uma falha na camada RAW não
  // pode interromper a camada operacional, mas fica registrada para auditoria.
  try {
    if (typeof registrarAtividadesBrutasStrava_ === 'function') {
      registrarAtividadesBrutasStrava_(athId, nomeAtleta, atividades);
    }
  } catch (eRaw) {
    _log(athId, 'AVISO', '_gravarAtividades', 'Falha ao registrar RAW: ' + eRaw.message, '');
  }

  const existentes = new Map();
  const dataAtual = sheet.getDataRange().getValues();
  for (let i = 2; i < dataAtual.length; i++) {
    const sid = String(dataAtual[i][H.ATIV.STRAVA_ID - 1] || '').trim();
    if (sid && !existentes.has(sid)) existentes.set(sid, i + 1);
  }

  let count = 0;
  let reparadas = 0;
  const novasLinhas = [];
  const linhasReparadas = [];
  for (const a of atividades) {
    const sid = String(a.id || '');

    // ── Normalização oficial via Normalizar.gs ─────────────────────────────────
    const norm = normalizarAtividadeStrava(a, { ath_id: athId, nome: nomeAtleta });
    if (!norm) { console.log('[GRAVAR] norm=null strava_id=' + (a.id || '?') + ' tipo=' + a.type); continue; }


    // Campos legados: mantidos para compatibilidade com Metricas.gs e Queue.gs
    const velMps = Number(a.average_speed) > 0 ? Number(a.average_speed) : 0;
    const velKmMin = velMps > 0 ? Math.round(velMps * 0.06 * 1000) / 1000 : 0;
    const paceSegKm = Number(norm.pace_s_km) || 0;
    const paceFmt = _formatarVelocidadeDisplay(velMps, norm.tipo);
    const execId = 'ATIV_' + Utilities.getUuid().substring(0, 8).toUpperCase();

    const velKmh = velMps > 0 ? Math.round(velMps * 3.6 * 100) / 100 : 0;
    const dataAtividade = norm.data_hora ? new Date(norm.data_hora) : null;
    const linha = [
      execId,                                    // 1  EXEC_ID
      norm.ath_id,                               // 2  ATH_ID
      norm.atleta,                               // 3  NOME
      dataAtividade || '',                       // 4 DATA
      norm.tipo,                                 // 5  TIPO (normalizado PT-BR)
      'Strava',                                  // 6  FONTE
      norm.strava_id,                            // 7  STRAVA_ID
      norm.nome_atividade,                       // 8  NOME_ATIV
      norm.tempo_mov_s || 0,                     // 9  MOV_S (segundos)
      a.elapsed_time || 0,                       // 10 TOTAL_S (segundos)
      a.distance ? Math.round(a.distance) : 0,  // 11 DIST_M (metros brutos)
      norm.dist_km,                              // 12 DIST_KM (normalizado)
      velMps,                                    // 13 VEL_MPS (m/s raw — analytics)
      velKmMin,                                  // 14 VEL_KMMIN
      paceSegKm,                                 // 15 PACE_S (s/km — analytics)
      paceFmt,                                   // 16 PACE_FMT (display)
      norm.fc_media,                             // 17 FC_MED (int normalizado)
      norm.fc_max,                               // 18 FC_MAX (int normalizado)
      norm.elev_m,                               // 19 ELEV (m inteiro)
      norm.calorias,                             // 20 CAL (kcal inteiro)
      norm.cadencia,                             // 21 CADENCIA
      norm.potencia_w,                           // 22 POTENCIA (W inteiro)
      '',                                        // 23 ROTA — polyline REMOVIDA (só Supabase)
      new Date(),                                // 24 IMPORTADO
      '',                                        // 25 PSE (entrada manual 1-10)
      norm.tempo_mov_fmt || '',                  // 26 TEMPO MOVIMENTO HH:MM:SS
      norm.tempo_total_fmt || '',                // 27 TEMPO TOTAL HH:MM:SS
      velKmh,                                    // 28 VELOCIDADE KM/H
      velKmh > 0 ? velKmh.toFixed(2).replace('.', ',') + ' km/h' : '', // 29 VEL_FMT
      norm.dist_km > 0 ? norm.dist_km.toFixed(1).replace('.', ',') + ' km' : '', // 30 DIST_FMT
      norm.tipo_original || '',                  // 31 TIPO ORIGINAL
      dataAtividade || '',                       // 32 DATA FORMATADA PELA PLANILHA
      dataAtividade ? Utilities.formatDate(dataAtividade, Session.getScriptTimeZone(), 'HH:mm:ss') : '', // 33 HORA
      'Importado',                               // 34 STATUS
    ];

    const linhaExistente = existentes.get(sid);
    if (linhaExistente) {
      if (linhaExistente > dataAtual.length) continue;
      const atual = dataAtual[linhaExistente - 1].slice(0, 34);
      while (atual.length < 34) atual.push('');
      if (!_atividadePrecisaReparo_(atual, linha)) continue;

      // Preserva campos internos e manuais; recompõe apenas uma linha que já
      // está comprovadamente incompleta com a resposta oficial da Strava.
      linha[0] = atual[0] || linha[0];
      linha[23] = atual[23] || linha[23];
      linha[24] = atual[24] === undefined ? '' : atual[24];
      linhasReparadas.push({ numero: linhaExistente, valores: linha });
      dataAtual[linhaExistente - 1] = linha;
      reparadas++;
      continue;
    }

    novasLinhas.push(linha);
    existentes.set(sid, sheet.getLastRow() + novasLinhas.length);
    count++;
  }

  _gravarBlocosAtividades_(sheet, linhasReparadas);
  if (novasLinhas.length) {
    sheet.getRange(sheet.getLastRow() + 1, 1, novasLinhas.length, 34).setValues(novasLinhas);
  }

  if (count > 0 || reparadas > 0) _atividadesAlteradasExecucao_ = true;

  _log(athId, 'INFO', '_gravarAtividades',
    count + ' novas; ' + reparadas + ' reparadas de ' + atividades.length + ' recebidas', '');
  console.log('[GRAVAR] athId=' + athId + ' novas=' + count +
    ' reparadas=' + reparadas + ' de ' + atividades.length);

  // Recalcular métricas se houver atividade nova ou reparo de dado antigo.
  if (count > 0 || reparadas > 0) {
    try {
      recalcularMetricasAposAtividade(athId);
    } catch (e) {
      // Não bloquear importação se métricas falharem
      _log(athId, 'AVISO', '_gravarAtividades', 'Recálculo de métricas falhou: ' + e.message, '');
    }
  }

  return count;
}

/** Uma atividade existente só é regravada quando faltam campos estruturais. */
function _atividadePrecisaReparo_(linha, oficial) {
  const dataAtual = linha[3] instanceof Date ? linha[3].getTime() : new Date(linha[3] || 0).getTime();
  const dataOficial = oficial && oficial[3] instanceof Date ? oficial[3].getTime() : dataAtual;
  const nomeAtual = String(linha[7] || '').trim();
  const nomeTraduzido = typeof traduzirNomeAtividadeStrava_ === 'function'
    ? traduzirNomeAtividadeStrava_(nomeAtual) : nomeAtual;
  return !String(linha[2] || '').trim() ||
    !linha[3] ||
    !String(linha[7] || '').trim() ||
    (nomeTraduzido !== nomeAtual && oficial && nomeTraduzido === String(oficial[7] || '').trim()) ||
    (Number(linha[10]) > 0 && !(Number(linha[11]) > 0)) ||
    (Number(linha[8]) > 0 && !String(linha[25] || '').trim()) ||
    (Number(linha[12]) > 0 && !String(linha[28] || '').trim()) ||
    (isFinite(dataAtual) && isFinite(dataOficial) && Math.abs(dataAtual - dataOficial) >= 60000);
}

/** Agrupa linhas contíguas para reduzir chamadas à planilha. */
function _gravarBlocosAtividades_(sheet, itens) {
  if (!itens.length) return;
  itens.sort((a, b) => a.numero - b.numero);
  let inicio = itens[0].numero;
  let anterior = inicio;
  let valores = [itens[0].valores];

  for (let i = 1; i < itens.length; i++) {
    const item = itens[i];
    if (item.numero === anterior + 1) {
      valores.push(item.valores);
    } else {
      sheet.getRange(inicio, 1, valores.length, valores[0].length).setValues(valores);
      inicio = item.numero;
      valores = [item.valores];
    }
    anterior = item.numero;
  }
  sheet.getRange(inicio, 1, valores.length, valores[0].length).setValues(valores);
}

// ── Formatar pace em segundos para "5:30 /km" ─────────────────────────────────
/**
 * _formatarPace — mantida para compatibilidade (migração)
 * Para novos dados, usar _formatarVelocidadeDisplay()
 */
function _formatarPace(paceSegKm) {
  if (!paceSegKm || paceSegKm <= 0 || paceSegKm > 3600) return '';
  const min = Math.floor(paceSegKm / 60);
  const seg = Math.round(paceSegKm % 60);
  return min + ':' + String(seg).padStart(2, '0') + ' /km';
}

/**
 * _formatarVelocidadeDisplay — display inteligente por esporte
 *   Corrida / Caminhada / Trail / Funcional → "5:30 /km"   (min:ss por km)
 *   Ciclismo / Ergométrica                  → "28.5 km/h"  (1 decimal)
 *   Natação                                 → "1:45 /100m" (min:ss por 100m)
 *   Outros                                  → "5:30 /km"   (fallback corrida)
 */
function _formatarVelocidadeDisplay(velMps, tipo) {
  if (!velMps || velMps <= 0) return '';
  const t = (tipo || '').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '');

  if (t === 'ciclismo' || t === 'ergometrica' || t === 'hibrido') {
    const kmh = Math.round(velMps * 3.6 * 10) / 10;
    return kmh.toFixed(1) + ' km/h';
  }

  if (t === 'natacao' || t === 'natacao') {
    const seg100m = Math.round(100 / velMps);
    if (seg100m <= 0 || seg100m > 3600) return '';
    const min = Math.floor(seg100m / 60);
    const seg = seg100m % 60;
    return min + ':' + String(seg).padStart(2, '0') + ' /100m';
  }

  // Default: corrida, caminhada, trail, funcional → min/km
  const paceSegKm = Math.round(1000 / velMps);
  if (paceSegKm <= 0 || paceSegKm > 3600) return '';
  const min = Math.floor(paceSegKm / 60);
  const seg = paceSegKm % 60;
  return min + ':' + String(seg).padStart(2, '0') + ' /km';
}

/**
 * _calcularPaceDisplay — retorna o valor numérico de pace/velocidade
 *   Corrida/Caminhada → pace em s/km
 *   Ciclismo          → velocidade em km/h × 10 (para preservar 1 decimal)
 *   Natação           → pace em s/100m
 */
function _calcularPaceNumerico(velMps, tipo) {
  if (!velMps || velMps <= 0) return 0;
  const t = (tipo || '').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '');
  if (t === 'ciclismo' || t === 'ergometrica' || t === 'hibrido') {
    return Math.round(velMps * 3.6 * 10) / 10; // km/h
  }
  if (t === 'natacao') {
    return Math.round(100 / velMps); // s/100m
  }
  return Math.round(1000 / velMps); // s/km
}

/**
 * _calcularDistanciaDisplay — arredondamento por esporte
 *   Corrida/Trail/Caminhada → 2 decimais km  (ex: 5.23)
 *   Ciclismo                → 1 decimal km   (ex: 42.5)
 *   Natação                 → inteiro metros  (ex: 1500) — armazenado como km mas exibido diferente
 */
function _calcularDistKm(distM, tipo) {
  if (!distM || distM <= 0) return 0;
  const t = (tipo || '').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '');
  if (t === 'ciclismo' || t === 'ergometrica' || t === 'hibrido') {
    return Math.round(distM / 100) / 10; // 1 decimal km
  }
  if (t === 'natacao') {
    return Math.round(distM) / 1000; // metros precisos → km
  }
  return Math.round(distM / 10) / 100; // 2 decimais km
}

// ── 8. PERFIL DO ATLETA ───────────────────────────────────────────────────────
function importarPerfilAtleta(athId) {
  if (!athId) {
    const ui = SpreadsheetApp.getUi();
    const r = ui.prompt('👤 Perfil', 'ID do atleta:', ui.ButtonSet.OK_CANCEL);
    if (r.getSelectedButton() !== ui.Button.OK) return;
    athId = (r.getResponseText() || '').trim().toUpperCase();
  }
  try {
    const accessToken = _getValidAccessToken(athId);
    const resp = UrlFetchApp.fetch(STRAVA_API_BASE + '/athlete', {
      headers: { Authorization: 'Bearer ' + accessToken },
      muteHttpExceptions: true
    });
    if (resp.getResponseCode() !== 200) throw new Error('Strava retornou ' + resp.getResponseCode());
    const perfil = JSON.parse(resp.getContentText());
    _atualizarPerfilNoCadastro(athId, perfil);
    _log(athId, 'INFO', 'importarPerfilAtleta', 'Perfil atualizado. ID: ' + perfil.id, '');
    return perfil;
  } catch (e) {
    _log(athId, 'ERRO', 'importarPerfilAtleta', e.message, e.stack || '');
    throw e;
  }
}

function _atualizarPerfilNoCadastro(athId, perfil) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(H.SHEETS.CADASTRO);
  if (!sheet) return;
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][H.CAD.ID - 1] || '').trim().toUpperCase() === athId.toUpperCase()) {
      if (perfil.weight) sheet.getRange(i + 1, H.CAD.PESO).setValue(perfil.weight);
      sheet.getRange(i + 1, H.CAD.STRAVA_OK).setValue('Sim');
      sheet.getRange(i + 1, H.CAD.STRAVA_ID).setValue(perfil.id);
      return;
    }
  }
}

// ── 9. TRIGGER E HELPERS ──────────────────────────────────────────────────────
function triggerImportacaoAutomatica() {
  _log('SISTEMA', 'INFO', 'triggerImportacaoAutomatica', 'Iniciando ciclo automático...', '');

  // 0. Tarefas rápidas e críticas de novos cadastros primeiro. Assim elas
  // concluem mesmo quando os cálculos de métricas/rankings levam mais tempo.
  try {
    if (typeof sincronizarCadastrosParaSupabaseSeguro === 'function') {
      sincronizarCadastrosParaSupabaseSeguro();
    }
    if (typeof sincronizarFilaWhatsAppCadastros === 'function') {
      sincronizarFilaWhatsAppCadastros();
    }
  } catch (e) {
    _log('SISTEMA', 'AVISO', 'triggerImportacaoAutomatica',
      'Pré-sincronização de cadastros: ' + e.message, '');
  }

  // 1. Processar fila inteligente com rate-limiting (Queue.gs)
  try {
    processarFilaStrava();
    _log('SISTEMA', 'INFO', 'triggerImportacaoAutomatica', 'Fila Strava processada.', '');
  } catch (e) {
    _log('SISTEMA', 'ERRO', 'triggerImportacaoAutomatica', 'Erro na fila: ' + e.message, '');
  }

  // 2. Recalcular métricas (VO2máx, zonas, pace) — silencioso
  try {
    calcularMetricasTodos();
    _log('SISTEMA', 'INFO', 'triggerImportacaoAutomatica', 'Métricas recalculadas.', '');
  } catch (e) {
    _log('SISTEMA', 'ERRO', 'triggerImportacaoAutomatica', 'Erro nas métricas: ' + e.message, '');
  }

  // 3. Atualizar rankings de forma independente: uma aba não bloqueia a outra.
  try {
    atualizarRankingSheet();
    _log('SISTEMA', 'INFO', 'triggerImportacaoAutomatica', 'RANKING atualizado.', '');
  } catch (e) {
    _log('SISTEMA', 'AVISO', 'triggerImportacaoAutomatica', 'RANKING não atualizado: ' + e.message, '');
  }
  try {
    atualizarRankingExpandido();
    _log('SISTEMA', 'INFO', 'triggerImportacaoAutomatica', 'RANKING COMPLETO atualizado.', '');
  } catch (e) {
    _log('SISTEMA', 'AVISO', 'triggerImportacaoAutomatica', 'RANKING COMPLETO não atualizado: ' + e.message, '');
  }

  // 3b. Atualizar análise científica (CTL/ATL/TSB)
  try {
    atualizarAnaliseSheet();
    _log('SISTEMA', 'INFO', 'triggerImportacaoAutomatica', 'ANÁLISE científica atualizada.', '');
  } catch (e) {
    _log('SISTEMA', 'AVISO', 'triggerImportacaoAutomatica', 'ANÁLISE não atualizada: ' + e.message, '');
  }

  // 4. Atualizar aba STRAVA STATUS (se existir)
  try {
    atualizarStravaStatusSheet();
    _log('SISTEMA', 'INFO', 'triggerImportacaoAutomatica', 'STRAVA STATUS atualizado.', '');
  } catch (e) {
    _log('SISTEMA', 'AVISO', 'triggerImportacaoAutomatica', 'STATUS não atualizado: ' + e.message, '');
  }

  // 5. Sincronizar atletas do CADASTRO em todas as abas
  try {
    const novos = sincronizarAtletasEmTodasAbas();
    if (novos > 0) _log('SISTEMA', 'INFO', 'triggerImportacaoAutomatica', novos + ' novos atletas sincronizados.', '');
  } catch (e) {
    _log('SISTEMA', 'AVISO', 'triggerImportacaoAutomatica', 'Sync atletas: ' + e.message, '');
  }

  // 6. Atualizar timestamp no PAINEL
  try {
    _atualizarPainelInterno();
  } catch (e) {
    _log('SISTEMA', 'AVISO', 'triggerImportacaoAutomatica', 'Painel não atualizado: ' + e.message, '');
  }

  _log('SISTEMA', 'INFO', 'triggerImportacaoAutomatica', 'Ciclo completo concluído.', '');
}

function _atualizarStatusCadastro(athId, conectado, stravaId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(H.SHEETS.CADASTRO);
  if (!sheet) return;
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][H.CAD.ID - 1] || '').trim().toUpperCase() === athId.toUpperCase()) {
      sheet.getRange(i + 1, H.CAD.STRAVA_OK).setValue(conectado ? 'Sim' : 'Não');
      if (stravaId) sheet.getRange(i + 1, H.CAD.STRAVA_ID).setValue(stravaId);
      return;
    }
  }
}

function _getNomeAtleta(athId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(H.SHEETS.CADASTRO);
    if (!sheet) return athId;
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][H.CAD.ID - 1] || '').trim().toUpperCase() === athId.toUpperCase()) {
        return String(data[i][H.CAD.NOME - 1] || athId).trim();
      }
    }
  } catch (_) { }
  return athId;
}

function _calcPace(tempoSeg, distKm) {
  if (!distKm || distKm < 0.01) return '';
  const totalSeg = tempoSeg / distKm;
  const min = Math.floor(totalSeg / 60);
  const seg = Math.round(totalSeg % 60);
  return min + ':' + String(seg).padStart(2, '0') + ' /km';
}

function _calcPaceDecimal(tempoSeg, distKm) {
  if (!distKm || distKm < 0.01) return 0;
  return Math.round(tempoSeg / distKm);
}

function _traduzirTipo(tipo) {
  const map = {
    Run: 'Corrida', Ride: 'Ciclismo', Swim: 'Natação', Walk: 'Caminhada',
    TrailRun: 'Trail Run', VirtualRide: 'Ciclismo Virtual', VirtualRun: 'Corrida Virtual',
    WeightTraining: 'Musculação', Yoga: 'Yoga', Workout: 'Treino', Hike: 'Trilha',
    AlpineSki: 'Ski', Rowing: 'Remo', Kayaking: 'Canoagem', Soccer: 'Futebol',
    Crossfit: 'CrossFit', Elliptical: 'Elíptico', StairStepper: 'Escada',
    RockClimbing: 'Escalada', Surfing: 'Surf', Skateboard: 'Skate',
    Badminton: 'Badminton', Tennis: 'Tênis', Volleyball: 'Vôlei', Basketball: 'Basquete'
  };
  return map[tipo] || tipo || 'Outro';
}

// ── 10. PÁGINAS HTML ──────────────────────────────────────────────────────────
function _paginaCadastro(athId) {
  return `<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Cadastro — Hiperativo</title>
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:'Segoe UI',Arial,sans-serif;background:linear-gradient(135deg,#050d1a 0%,#0b1f3a 60%,#071428 100%);min-height:100vh;padding:20px 16px}
.wrapper{max-width:680px;margin:0 auto}
.header{text-align:center;padding:28px 0 20px}
.logo-wrap{margin:0 auto 14px;display:flex;align-items:center;justify-content:center;gap:16px}
.logo-svg{width:64px;height:64px;flex-shrink:0}
.brand-text{text-align:left}
.brand-name{font-size:2.2rem;font-weight:900;color:#fff;letter-spacing:2px;line-height:1;text-transform:uppercase}
.brand-sub{font-size:.72rem;color:#4fc3f7;letter-spacing:4px;text-transform:uppercase;margin-top:2px}
.header-desc{color:#90caf9;font-size:.9rem;margin-top:10px}
.card{background:rgba(255,255,255,.04);backdrop-filter:blur(8px);border:1px solid rgba(255,255,255,.09);border-radius:18px;padding:28px 24px;margin-bottom:14px}
.card-title{font-size:.95rem;font-weight:700;color:#4fc3f7;text-transform:uppercase;letter-spacing:1.2px;margin-bottom:18px;padding-bottom:10px;border-bottom:1px solid rgba(79,195,247,.2);display:flex;align-items:center;gap:8px}
.grid2{display:grid;grid-template-columns:1fr 1fr;gap:14px}
.grid3{display:grid;grid-template-columns:1fr 1fr 1fr;gap:14px}
@media(max-width:480px){.grid2,.grid3{grid-template-columns:1fr}}
.field{display:flex;flex-direction:column;gap:5px;margin-top:0}
.field+.field{margin-top:0}
.gap{margin-top:14px}
.field label{font-size:.75rem;font-weight:600;color:#90caf9;text-transform:uppercase;letter-spacing:.6px}
.req{color:#ef5350}
.field input,.field select,.field textarea{background:rgba(255,255,255,.07);border:1px solid rgba(255,255,255,.12);border-radius:9px;padding:11px 13px;color:#fff;font-size:.93rem;transition:border-color .2s,background .2s;outline:none;width:100%}
.field input::placeholder{color:rgba(255,255,255,.3)}
.field input:focus,.field select:focus,.field textarea:focus{border-color:#4fc3f7;background:rgba(79,195,247,.07)}
.field select option{background:#0b1f3a;color:#fff}
.field textarea{resize:vertical;min-height:76px}
.check-group{display:flex;flex-wrap:wrap;gap:7px;margin-top:6px}
.chk{display:flex;align-items:center;gap:6px;background:rgba(255,255,255,.05);border:1px solid rgba(255,255,255,.1);border-radius:8px;padding:7px 11px;cursor:pointer;color:#e0e0e0;font-size:.84rem;transition:all .18s;user-select:none}
.chk input{accent-color:#00c853;width:15px;height:15px;cursor:pointer}
.chk:hover{border-color:#4fc3f7;background:rgba(79,195,247,.09)}
.planos{display:grid;grid-template-columns:repeat(auto-fill,minmax(130px,1fr));gap:10px;margin-top:6px}
.plano{border:2px solid rgba(255,255,255,.1);border-radius:13px;padding:14px 10px;cursor:pointer;text-align:center;transition:all .2s;position:relative;background:rgba(255,255,255,.03)}
.plano:hover{border-color:#4fc3f7;transform:translateY(-2px)}
.plano input{position:absolute;opacity:0;pointer-events:none}
.plano.sel{border-color:#00c853;background:rgba(0,200,83,.09)}
.plano-ico{font-size:1.6rem;margin-bottom:5px}
.plano-nm{font-size:.9rem;font-weight:700;color:#fff}
.plano-dc{font-size:.72rem;color:#90caf9;margin-top:3px;line-height:1.4}
.badge{font-size:.65rem;font-weight:700;border-radius:20px;padding:2px 8px;margin-top:5px;display:inline-block}
.b-free{background:#37474f;color:#b0bec5}
.b-start{background:#1a237e;color:#90caf9}
.b-pro{background:#003d00;color:#69f0ae}
.b-elite{background:#4a148c;color:#ea80fc}
.b-corp{background:#bf360c;color:#ffccbc}
.lgpd{background:rgba(255,193,7,.07);border:1px solid rgba(255,193,7,.25);border-radius:11px;padding:14px;margin-top:10px}
.lgpd p{color:#ffd54f;font-size:.8rem;line-height:1.65}
.lgpd-chk{display:flex;align-items:flex-start;gap:9px;margin-top:11px;color:#e0e0e0;font-size:.83rem;cursor:pointer;line-height:1.5}
.lgpd-chk input{accent-color:#00c853;margin-top:2px;flex-shrink:0;cursor:pointer}
.btn-submit{width:100%;padding:16px;background:linear-gradient(135deg,#00c853,#00897b);border:none;border-radius:13px;font-size:1rem;font-weight:700;color:#fff;cursor:pointer;letter-spacing:.8px;text-transform:uppercase;margin-top:6px;transition:all .2s}
.btn-submit:hover{transform:translateY(-2px);box-shadow:0 8px 28px rgba(0,200,83,.35)}
.strava-hint{text-align:center;color:#90caf9;font-size:.8rem;margin-top:14px;display:flex;align-items:center;justify-content:center;gap:7px}
.s-ico{background:#fc4c02;border-radius:50%;width:18px;height:18px;display:inline-flex;align-items:center;justify-content:center;font-size:.65rem;color:#fff;font-weight:900;flex-shrink:0}
.req-note{color:#90caf9;font-size:.76rem;margin-bottom:14px;text-align:right}
</style>
</head>
<body>
<div class="wrapper">
<div class="header">
  <div class="logo-wrap">
    <svg class="logo-svg" viewBox="0 0 200 200" xmlns="http://www.w3.org/2000/svg">
      <circle cx="100" cy="100" r="82" fill="none" stroke="#1e3a8a" stroke-width="14" stroke-dasharray="195 65" stroke-dashoffset="32" stroke-linecap="round"/>
      <polygon points="118,18 74,105 100,105 82,182 132,82 106,82" fill="#00c853"/>
    </svg>
    <div class="brand-text">
      <div class="brand-name">Hiper<br>ativo</div>
      <div class="brand-sub">Performance &amp; Tech</div>
    </div>
  </div>
  <p class="header-desc">Preencha o formulário abaixo para começar sua jornada de performance</p>
</div>

<form id="f" onsubmit="return enviar(event)">
<input type="hidden" name="salvar" value="true">
<input type="hidden" name="athId" value="${athId}">
<p class="req-note">Campos com <span class="req">*</span> são obrigatórios</p>

<div class="card">
  <div class="card-title">👤 Dados Pessoais</div>
  <div class="grid2">
    <div class="field"><label>Nome completo <span class="req">*</span></label><input type="text" name="nome" placeholder="Seu nome" required></div>
    <div class="field"><label>E-mail <span class="req">*</span></label><input type="email" name="email" placeholder="seu@email.com" required></div>
  </div>
  <div class="grid2 gap">
    <div class="field"><label>WhatsApp <span class="req">*</span></label><input type="tel" name="whats" placeholder="(11) 99999-9999" required></div>
    <div class="field"><label>Data de nascimento <span class="req">*</span></label><input type="date" name="nasc" required></div>
  </div>
  <div class="grid3 gap">
    <div class="field"><label>Sexo <span class="req">*</span></label>
      <select name="sexo" required>
        <option value="">Selecionar</option>
        <option>Masculino</option><option>Feminino</option><option>Outro</option><option>Prefiro não informar</option>
      </select>
    </div>
    <div class="field"><label>Peso (kg)</label><input type="number" name="peso" placeholder="70" min="30" max="300" step="0.1"></div>
    <div class="field"><label>Altura (cm)</label><input type="number" name="altura" placeholder="170" min="100" max="250"></div>
  </div>
  <div class="grid2 gap">
    <div class="field"><label>Cidade</label><input type="text" name="cidade" placeholder="Sua cidade"></div>
    <div class="field"><label>Estado</label>
      <select name="estado">
        <option value="">UF</option>
        <option>AC</option><option>AL</option><option>AM</option><option>AP</option><option>BA</option><option>CE</option>
        <option>DF</option><option>ES</option><option>GO</option><option>MA</option><option>MG</option><option>MS</option>
        <option>MT</option><option>PA</option><option>PB</option><option>PE</option><option>PI</option><option>PR</option>
        <option>RJ</option><option>RN</option><option>RO</option><option>RR</option><option>RS</option><option>SC</option>
        <option>SE</option><option>SP</option><option>TO</option>
      </select>
    </div>
  </div>
  <div class="field gap"><label>CPF (opcional, para emissão de NF)</label><input type="text" name="cpf" placeholder="000.000.000-00" maxlength="14"></div>
</div>

<div class="card">
  <div class="card-title">🏃 Esporte &amp; Objetivos</div>
  <div class="field"><label>Modalidades de interesse <span class="req">*</span></label>
    <div class="check-group">
      <label class="chk"><input type="checkbox" name="mod" value="Corrida"> 🏃 Corrida</label>
      <label class="chk"><input type="checkbox" name="mod" value="Ciclismo"> 🚴 Ciclismo</label>
      <label class="chk"><input type="checkbox" name="mod" value="Natação"> 🏊 Natação</label>
      <label class="chk"><input type="checkbox" name="mod" value="Triathlon"> 🔱 Triathlon</label>
      <label class="chk"><input type="checkbox" name="mod" value="Trail Run"> 🏔️ Trail Run</label>
      <label class="chk"><input type="checkbox" name="mod" value="Caminhada"> 🚶 Caminhada</label>
      <label class="chk"><input type="checkbox" name="mod" value="Musculação"> 💪 Musculação</label>
      <label class="chk"><input type="checkbox" name="mod" value="CrossFit"> ⚡ CrossFit</label>
      <label class="chk"><input type="checkbox" name="mod" value="Funcional"> 🤸 Funcional</label>
      <label class="chk"><input type="checkbox" name="mod" value="Yoga/Pilates"> 🧘 Yoga/Pilates</label>
      <label class="chk"><input type="checkbox" name="mod" value="Futebol"> ⚽ Futebol</label>
      <label class="chk"><input type="checkbox" name="mod" value="Tênis"> 🎾 Tênis</label>
      <label class="chk"><input type="checkbox" name="mod" value="Outro"> ➕ Outro</label>
    </div>
  </div>
  <div class="grid2 gap">
    <div class="field"><label>Nível atual <span class="req">*</span></label>
      <select name="nivel" required>
        <option value="">Selecionar</option>
        <option>Iniciante (até 6 meses)</option>
        <option>Básico (6m – 2 anos)</option>
        <option>Intermediário (2 – 5 anos)</option>
        <option>Avançado (5+ anos)</option>
        <option>Competidor amador</option>
        <option>Atleta profissional</option>
      </select>
    </div>
    <div class="field"><label>Frequência semanal atual</label>
      <select name="freq">
        <option value="">Selecionar</option>
        <option>1-2x por semana</option>
        <option>3-4x por semana</option>
        <option>5-6x por semana</option>
        <option>Todos os dias</option>
        <option>Varia muito</option>
      </select>
    </div>
  </div>
  <div class="field gap"><label>Objetivos principais <span class="req">*</span></label>
    <div class="check-group">
      <label class="chk"><input type="checkbox" name="obj" value="Perda de peso"> ⚖️ Perda de peso</label>
      <label class="chk"><input type="checkbox" name="obj" value="Ganho de massa"> 💪 Ganho de massa</label>
      <label class="chk"><input type="checkbox" name="obj" value="Condicionamento"> ❤️ Condicionamento</label>
      <label class="chk"><input type="checkbox" name="obj" value="Preparação para prova"> 🏆 Preparação prova</label>
      <label class="chk"><input type="checkbox" name="obj" value="5km"> 🎯 Correr 5km</label>
      <label class="chk"><input type="checkbox" name="obj" value="10km"> 🎯 Correr 10km</label>
      <label class="chk"><input type="checkbox" name="obj" value="Meia maratona"> 🏅 Meia maratona</label>
      <label class="chk"><input type="checkbox" name="obj" value="Maratona"> 🏅 Maratona</label>
      <label class="chk"><input type="checkbox" name="obj" value="Ironman"> 🔱 Ironman</label>
      <label class="chk"><input type="checkbox" name="obj" value="Saúde e bem-estar"> 🌱 Saúde/bem-estar</label>
      <label class="chk"><input type="checkbox" name="obj" value="Reabilitação"> 🩺 Reabilitação</label>
      <label class="chk"><input type="checkbox" name="obj" value="Performance"> 📈 Performance geral</label>
    </div>
  </div>
  <div class="grid2 gap">
    <div class="field"><label>Horário preferido</label>
      <select name="horario">
        <option value="">Selecionar</option>
        <option>Manhã cedo (antes 7h)</option>
        <option>Manhã (7h–10h)</option>
        <option>Meio-dia (10h–13h)</option>
        <option>Tarde (13h–18h)</option>
        <option>Noite (18h–21h)</option>
        <option>Noite tarde (após 21h)</option>
        <option>Varia muito</option>
      </select>
    </div>
    <div class="field"><label>Como nos conheceu?</label>
      <select name="origem">
        <option value="">Selecionar</option>
        <option>Instagram</option><option>Facebook</option><option>YouTube</option>
        <option>TikTok</option><option>Indicação de amigo</option>
        <option>Google</option><option>Strava</option><option>Outro</option>
      </select>
    </div>
  </div>
  <div class="field gap"><label>Próxima prova / evento (se houver)</label><input type="text" name="prova" placeholder="Ex: Maratona de SP — Nov/2026"></div>
</div>

<div class="card">
  <div class="card-title">🩺 Saúde &amp; Histórico</div>
  <div class="field"><label>Condições de saúde</label>
    <div class="check-group">
      <label class="chk"><input type="checkbox" name="saude" value="Hipertensão"> 💊 Hipertensão</label>
      <label class="chk"><input type="checkbox" name="saude" value="Diabetes"> 💉 Diabetes</label>
      <label class="chk"><input type="checkbox" name="saude" value="Cardíaco"> ❤️ Cardíaco</label>
      <label class="chk"><input type="checkbox" name="saude" value="Asma"> 🫁 Asma</label>
      <label class="chk"><input type="checkbox" name="saude" value="Ortopédico"> 🦴 Ortopédico</label>
      <label class="chk"><input type="checkbox" name="saude" value="Nenhuma"> ✅ Nenhuma</label>
      <label class="chk"><input type="checkbox" name="saude" value="Outra"> ➕ Outra</label>
    </div>
  </div>
  <div class="field gap"><label>Lesões ou restrições físicas</label>
    <textarea name="lesao" placeholder="Descreva lesões, cirurgias ou limitações físicas..."></textarea>
  </div>
  <div class="field gap"><label>Uso de medicamentos contínuos?</label>
    <input type="text" name="medicamento" placeholder="Ex: anti-hipertensivo, insulina (ou Não)">
  </div>
</div>

<div class="card">
  <div class="card-title">💳 Plano de Interesse</div>
  <div class="planos">
    <div class="plano" onclick="selP(this,'Free')"><input type="radio" name="plano" value="Free"><div class="plano-ico">🆓</div><div class="plano-nm">Free</div><div class="plano-dc">Análise básica e acompanhamento inicial</div><div class="badge b-free">GRATUITO</div></div>
    <div class="plano" onclick="selP(this,'Start')"><input type="radio" name="plano" value="Start"><div class="plano-ico">🚀</div><div class="plano-nm">Start</div><div class="plano-dc">Plano semanal + zonas CCC</div><div class="badge b-start">BÁSICO</div></div>
    <div class="plano" onclick="selP(this,'Pro')"><input type="radio" name="plano" value="Pro"><div class="plano-ico">⚡</div><div class="plano-nm">Pro</div><div class="plano-dc">Planilha personalizada + análise IA</div><div class="badge b-pro">POPULAR</div></div>
    <div class="plano" onclick="selP(this,'Elite')"><input type="radio" name="plano" value="Elite"><div class="plano-ico">🏆</div><div class="plano-nm">Elite</div><div class="plano-dc">Coach dedicado + relatórios avançados</div><div class="badge b-elite">PREMIUM</div></div>
    <div class="plano" onclick="selP(this,'Corporativo')"><input type="radio" name="plano" value="Corporativo"><div class="plano-ico">🏢</div><div class="plano-nm">Corporativo</div><div class="plano-dc">Equipes e grupos empresariais</div><div class="badge b-corp">EMPRESAS</div></div>
  </div>
</div>

<div class="card">
  <div class="card-title">🔒 Privacidade &amp; LGPD</div>
  <div class="lgpd">
    <p>Seus dados serão tratados conforme a <strong>Lei Geral de Proteção de Dados (LGPD — Lei 13.709/2018)</strong>. As informações são usadas exclusivamente para personalização de treinos e acompanhamento esportivo. Não compartilhamos seus dados com terceiros sem autorização prévia.</p>
    <label class="lgpd-chk"><input type="checkbox" id="c1" required><span>Li e aceito a política de privacidade da Hiperativo e autorizo o uso dos meus dados para fins esportivos e de CRM.</span></label>
    <label class="lgpd-chk"><input type="checkbox" id="c2" required><span>Entendo que serei redirecionado ao <strong>Strava</strong> para autorizar o acesso seguro às minhas atividades.</span></label>
    <label class="lgpd-chk"><input type="checkbox" id="c3"><span>Aceito receber comunicações da Hiperativo por WhatsApp e e-mail. (opcional)</span></label>
  </div>
</div>

<button type="submit" class="btn-submit">⚡ Salvar e Conectar ao Strava</button>
<div class="strava-hint"><span class="s-ico">S</span>Você será redirecionado ao Strava para autorizar o acesso seguro aos seus treinos</div>
</form>
</div>

<script>
function selP(el,v){
  document.querySelectorAll('.plano').forEach(c=>c.classList.remove('sel'));
  el.classList.add('sel');
  el.querySelector('input[type=radio]').checked=true;
}
function enviar(e){
  e.preventDefault();
  const f=document.getElementById('f');
  const mods=[...document.querySelectorAll('input[name=mod]:checked')].map(i=>i.value).join(', ');
  const objs=[...document.querySelectorAll('input[name=obj]:checked')].map(i=>i.value).join(', ');
  const saudes=[...document.querySelectorAll('input[name=saude]:checked')].map(i=>i.value).join(', ');
  const plano=document.querySelector('input[name=plano]:checked');
  if(!mods){alert('Selecione ao menos uma modalidade.');return false;}
  if(!objs){alert('Selecione ao menos um objetivo.');return false;}
  if(!plano){alert('Selecione um plano de interesse.');return false;}
  if(!document.getElementById('c1').checked){alert('Aceite os termos de privacidade para continuar.');return false;}
  if(!document.getElementById('c2').checked){alert('Confirme que entende a conexão com o Strava.');return false;}
  const aH=(n,v)=>{let i=f.querySelector('input[name='+n+']');if(!i){i=document.createElement('input');i.type='hidden';i.name=n;f.appendChild(i);}i.value=v;};
  aH('modAgg',mods);aH('objAgg',objs);aH('saudeAgg',saudes);
  const btn=f.querySelector('.btn-submit');
  btn.textContent='⏳ Salvando...';btn.disabled=true;
  const fd=new FormData(f);
  fd.delete('mod');fd.delete('obj');fd.delete('saude');
  fd.set('mod',mods);fd.set('obj',objs);fd.set('saude',saudes);
  window.location.href=window.location.href.split('?')[0]+'?'+new URLSearchParams(fd).toString();
  return false;
}
</script>
</body></html>`;
}

function _paginaRedirecionando(oauthUrl, nome) {
  const primeiroNome = (nome || '').split(' ')[0] || 'Atleta';
  return `<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<meta http-equiv="refresh" content="3;url=${oauthUrl}">
<title>Redirecionando...</title>
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:'Segoe UI',Arial,sans-serif;background:linear-gradient(135deg,#050d1a,#0b1f3a);min-height:100vh;display:flex;align-items:center;justify-content:center;padding:20px}
.card{background:rgba(255,255,255,.05);border:1px solid rgba(255,255,255,.1);border-radius:20px;padding:48px 36px;max-width:420px;width:100%;text-align:center}
.ico{font-size:60px;margin-bottom:14px;animation:pulse 1.5s ease-in-out infinite}
@keyframes pulse{0%,100%{transform:scale(1)}50%{transform:scale(1.08)}}
h2{color:#fff;font-size:1.3rem;margin-bottom:8px}
p{color:#90caf9;font-size:.9rem;line-height:1.65}
.strava{color:#fc4c02;font-weight:700}
.bar{width:100%;height:4px;background:rgba(255,255,255,.08);border-radius:2px;margin-top:22px;overflow:hidden}
.bar-fill{height:100%;background:linear-gradient(90deg,#00c853,#4fc3f7);animation:fill 3s linear forwards}
@keyframes fill{0%{width:0}100%{width:100%}}
.btn{display:inline-block;margin-top:18px;padding:11px 26px;background:linear-gradient(135deg,#fc4c02,#e53935);border-radius:10px;color:#fff;font-weight:700;text-decoration:none;font-size:.9rem}
</style>
</head>
<body>
<div class="card">
  <div class="ico">⚡</div>
  <h2>Cadastro salvo, ${primeiroNome}!</h2>
  <p>Redirecionando para o <span class="strava">Strava</span> em 3 segundos...</p>
  <p style="margin-top:8px;font-size:.8rem;color:#64b5f6">Autorize o acesso às suas atividades de forma segura.</p>
  <div class="bar"><div class="bar-fill"></div></div>
  <a href="${oauthUrl}" class="btn">Ir agora ao Strava ↗</a>
</div>
</body></html>`;
}

function _paginaSucesso(athId, athlete) {
  const nome = athlete ? ((athlete.firstname || '') + ' ' + (athlete.lastname || '')).trim() : athId;
  const foto = athlete && athlete.profile_medium ? athlete.profile_medium : '';
  return `<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Conectado ✅</title>
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:'Segoe UI',Arial,sans-serif;background:linear-gradient(135deg,#050d1a,#0b1f3a);min-height:100vh;display:flex;align-items:center;justify-content:center;padding:20px}
.card{background:rgba(255,255,255,.05);border:1px solid rgba(255,255,255,.1);border-radius:20px;padding:48px 36px;max-width:420px;width:100%;text-align:center}
.avatar{width:78px;height:78px;border-radius:50%;border:3px solid #00c853;margin:0 auto 14px;object-fit:cover}
.ico{font-size:68px;margin-bottom:14px}
h2{color:#fff;font-size:1.3rem;margin-bottom:6px}
.nm{color:#69f0ae;font-size:1.05rem;font-weight:700;margin-bottom:12px}
p{color:#90caf9;font-size:.88rem;line-height:1.65}
.checks{text-align:left;margin:18px 0;background:rgba(0,200,83,.07);border-radius:11px;padding:14px}
.checks div{color:#e0f2f1;font-size:.86rem;padding:3px 0;display:flex;align-items:center;gap:7px}
.badge{display:inline-block;background:linear-gradient(135deg,#00c853,#00897b);color:#fff;border-radius:20px;padding:6px 20px;font-weight:700;font-size:.88rem;margin-top:12px}
.note{color:#64b5f6;font-size:.78rem;margin-top:14px}
</style>
</head>
<body>
<div class="card">
  ${foto ? '<img class="avatar" src="' + foto + '" alt="foto">' : '<div class="ico">✅</div>'}
  <h2>Tudo certo!</h2>
  <div class="nm">${nome}</div>
  <p>Seu Strava foi conectado com sucesso à plataforma Hiperativo.</p>
  <div class="checks">
    <div>✅ Cadastro salvo na plataforma</div>
    <div>✅ Strava autorizado e vinculado</div>
    <div>✅ Atividades serão importadas automaticamente</div>
    <div>✅ Análise de zonas CCC disponível em breve</div>
  </div>
  <p class="note">Pode fechar esta página. Seu treinador já foi notificado!</p>
  <div class="badge">⚡ Bem-vindo ao Hiperativo!</div>
</div>
</body></html>`;
}

function _paginaErro(titulo, mensagem) {
  return `<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Erro — Hiperativo</title>
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:'Segoe UI',Arial,sans-serif;background:linear-gradient(135deg,#150505,#2d1010);min-height:100vh;display:flex;align-items:center;justify-content:center;padding:20px}
.card{background:rgba(255,255,255,.04);border:1px solid rgba(239,83,80,.25);border-radius:20px;padding:48px 36px;max-width:420px;width:100%;text-align:center}
.ico{font-size:60px;margin-bottom:14px}
h2{color:#ef5350;font-size:1.2rem;margin-bottom:10px}
p{color:#ffcdd2;font-size:.88rem;line-height:1.7}
.hint{color:#90caf9;font-size:.8rem;margin-top:18px;padding:12px;background:rgba(255,255,255,.04);border-radius:9px}
</style>
</head>
<body>
<div class="card">
  <div class="ico">⚠️</div>
  <h2>${titulo}</h2>
  <p>${mensagem}</p>
  <div class="hint">Precisa de ajuda? Entre em contato com seu treinador ou com o suporte Hiperativo.</div>
</div>
</body></html>`;
}

// ══════════════════════════════════════════════════════════════════════════════
// ── VERIFICAR E ATUALIZAR STATUS STRAVA DE TODOS OS ATLETAS ──────────────────
// ══════════════════════════════════════════════════════════════════════════════


// ══════════════════════════════════════════════════════════════════════════════
// ── ENVIAR LINK STRAVA PARA ATLETAS PENDENTES (email + WhatsApp) ─────────────
// ══════════════════════════════════════════════════════════════════════════════
function _enviarLinkStravaDesconectadosLegado_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cad = ss.getSheetByName(H.SHEETS.CADASTRO);
  const ui = SpreadsheetApp.getUi();

  if (!cad) { ui.alert('❌ Aba CADASTRO não encontrada.'); return; }

  const props = PropertiesService.getScriptProperties();
  const webAppUrl = props.getProperty('WEBAPP_URL') || '';
  const adminEmail = props.getProperty('ADMIN_EMAIL') || 'contato@ghiperativo.com.br';

  if (!webAppUrl) {
    ui.alert('❌ WEBAPP_URL não configurado.',
      'Acesse ⚙️ Configurações → 🔧 Configurar credenciais Strava.', ui.ButtonSet.OK);
    return;
  }

  const confirm = ui.alert(
    '📤 Enviar links Strava',
    'Isso vai enviar email + WhatsApp para todos os atletas com status "Pendente".\n\nContinuar?',
    ui.ButtonSet.YES_NO
  );
  if (confirm !== ui.Button.YES) return;

  // Atualizar status antes de enviar
  verificarStatusStravaAtletas();

  const data = cad.getDataRange().getValues();
  let enviados = 0, erros = 0;
  const linhas = [];

  for (let i = 1; i < data.length; i++) {
    const athId = String(data[i][H.CAD.ID - 1] || '').trim().toUpperCase();
    const nome = String(data[i][H.CAD.NOME - 1] || '').trim();
    const email = String(data[i][H.CAD.EMAIL - 1] || '').trim();
    const whats = String(data[i][H.CAD.WHATS - 1] || '').trim();
    const stravaOk = String(data[i][H.CAD.STRAVA_OK - 1] || '').trim().toLowerCase();
    const status = String(data[i][H.CAD.STATUS - 1] || '').trim().toLowerCase();

    if (!athId || !/^ATH/i.test(athId)) continue;
    if (status === 'inativo') continue;
    if (stravaOk === 'conectado') continue;

    const primeiro = nome.split(' ')[0] || 'Atleta';
    let ok = false;

    try {
      const oauthUrl = _gerarUrlOAuth(athId);

      // ── Email ────────────────────────────────────────────────────────────
      if (email) {
        try {
          MailApp.sendEmail({
            to: email,
            subject: '🏃 ' + primeiro + ', conecte seu Strava ao Hiperativo!',
            htmlBody: _htmlEmailConexaoStrava(primeiro, athId, oauthUrl),
            replyTo: adminEmail
          });
          ok = true;
        } catch (eEmail) {
          _log(athId, 'AVISO', 'enviarLinkStravaDesconectados', 'Email falhou: ' + eEmail.message, '');
        }
      }

      // ── WhatsApp ─────────────────────────────────────────────────────────
      if (whats) {
        const numero = whats.replace(/\D/g, '');
        const msgWA =
          'Olá, ' + primeiro + '! 👋\n\n' +
          'Conecte sua conta *Strava* ao Hiperativo para que seu treinador acompanhe seus treinos automaticamente, sem precisar enviar nada. 🏃\n\n' +
          '🔗 Clique aqui para conectar:\n' + oauthUrl + '\n\n' +
          'Qualquer dúvida, é só chamar! ⚡';
        _enviarWhatsApp(numero, msgWA, athId);
        ok = true;
      }

      if (ok) enviados++;
      else { erros++; linhas.push('⚠️ ' + athId + ' — sem email/whats'); }
      linhas.push((ok ? '✅' : '⚠️') + ' ' + athId + ' — ' + primeiro);

    } catch (e) {
      erros++;
      linhas.push('❌ ' + athId + ' — ' + e.message.slice(0, 60));
      _log(athId, 'ERRO', 'enviarLinkStravaDesconectados', e.message, '');
    }
  }

  const resumo = '📤 Enviados: ' + enviados + '   ❌ Erros: ' + erros + '\n\n' + linhas.join('\n');
  ui.alert('📤 Links Strava Enviados', resumo.slice(0, 1000), ui.ButtonSet.OK);
}

// ── WhatsApp: Z-API → Evolution API → wa.me (fallback) ──────────────────────
function _enviarWhatsApp(numero, mensagem, athId) {
  const props = PropertiesService.getScriptProperties();
  numero = String(numero).replace(/\D/g, '');
  if (!numero) return;

  // Adicionar DDI Brasil se não tiver
  if (!numero.startsWith('55') && numero.length <= 11) numero = '55' + numero;

  // ── Opção 1: Z-API ──────────────────────────────────────────────────────
  const zapiInst = props.getProperty('ZAPI_INSTANCE') || '';
  const zapiToken = props.getProperty('ZAPI_TOKEN') || '';
  const zapiClient = props.getProperty('ZAPI_CLIENT_TOKEN') || '';
  if (zapiInst && zapiToken) {
    try {
      const url = 'https://api.z-api.io/instances/' + zapiInst + '/token/' + zapiToken + '/send-text';
      const resp = UrlFetchApp.fetch(url, {
        method: 'POST',
        contentType: 'application/json',
        headers: zapiClient ? { 'Client-Token': zapiClient } : {},
        payload: JSON.stringify({ phone: numero, message: mensagem }),
        muteHttpExceptions: true
      });
      if (resp.getResponseCode() === 200) {
        _log(athId || 'SISTEMA', 'INFO', '_enviarWhatsApp', 'Z-API OK → ' + numero, '');
        return;
      }
      _log(athId || 'SISTEMA', 'AVISO', '_enviarWhatsApp', 'Z-API status ' + resp.getResponseCode() + ': ' + resp.getContentText().slice(0, 100), '');
    } catch (e) {
      _log(athId || 'SISTEMA', 'AVISO', '_enviarWhatsApp', 'Z-API erro: ' + e.message, '');
    }
  }

  // ── Opção 2: Evolution API ───────────────────────────────────────────────
  const evoUrl = props.getProperty('EVOLUTION_URL') || '';
  const evoToken = props.getProperty('EVOLUTION_TOKEN') || '';
  const evoInstance = props.getProperty('EVOLUTION_INSTANCE') || '';
  if (evoUrl && evoToken && evoInstance) {
    try {
      const url = evoUrl.replace(/\/$/, '') + '/message/sendText/' + evoInstance;
      const resp = UrlFetchApp.fetch(url, {
        method: 'POST',
        contentType: 'application/json',
        headers: { apikey: evoToken },
        payload: JSON.stringify({ number: numero, text: mensagem }),
        muteHttpExceptions: true
      });
      if (resp.getResponseCode() === 201 || resp.getResponseCode() === 200) {
        _log(athId || 'SISTEMA', 'INFO', '_enviarWhatsApp', 'Evolution OK → ' + numero, '');
        return;
      }
      _log(athId || 'SISTEMA', 'AVISO', '_enviarWhatsApp', 'Evolution status ' + resp.getResponseCode() + ': ' + resp.getContentText().slice(0, 100), '');
    } catch (e) {
      _log(athId || 'SISTEMA', 'AVISO', '_enviarWhatsApp', 'Evolution erro: ' + e.message, '');
    }
  }

  // ── Fallback: logar o wa.me link (configurar API depois) ────────────────
  const waLink = 'https://wa.me/' + numero + '?text=' + encodeURIComponent(mensagem);
  _log(athId || 'SISTEMA', 'INFO', '_enviarWhatsApp',
    'Sem API WhatsApp configurada. wa.me manual: ' + waLink.slice(0, 200), '');
}

// ── Template email para reconexão Strava ─────────────────────────────────────
function _htmlEmailConexaoStrava(nome, athId, linkStrava) {
  return `<!DOCTYPE html>
<html><head><meta charset="UTF-8"></head>
<body style="margin:0;padding:0;background:#f4f6f9;font-family:Arial,sans-serif">
<table width="100%" cellpadding="0" cellspacing="0">
<tr><td align="center" style="padding:32px 16px">
<table width="580" cellpadding="0" cellspacing="0" style="background:#fff;border-radius:12px;overflow:hidden;box-shadow:0 4px 20px rgba(0,0,0,.12)">
<tr><td style="background:linear-gradient(135deg,#1a3a8a,#0d2560);padding:32px;text-align:center">
  <div style="font-size:48px">⚡</div>
  <h1 style="color:#fff;margin:12px 0 4px;font-size:26px;letter-spacing:2px">HIPERATIVO</h1>
  <p style="color:#00c853;margin:0;font-size:12px;letter-spacing:3px">CABEÇA • CORAÇÃO • CORPO</p>
</td></tr>
<tr><td style="padding:32px">
  <h2 style="color:#1a3a8a;margin:0 0 12px">Olá, ${nome}! 🏃</h2>
  <p style="color:#444;line-height:1.6">Seu treinador está aguardando a conexão do seu <strong>Strava</strong> para acompanhar seus treinos automaticamente.</p>
  <p style="color:#444;line-height:1.6">Com o Strava conectado, <strong>nada precisa ser enviado manualmente</strong> — suas corridas, pedais e natações aparecerão automaticamente no painel do treinador.</p>
  <div style="text-align:center;margin:28px 0">
    <a href="${linkStrava}" style="background:#fc4c02;color:#fff;padding:16px 40px;border-radius:8px;text-decoration:none;font-weight:bold;font-size:17px;display:inline-block;letter-spacing:1px">
      🔗 Conectar meu Strava agora
    </a>
  </div>
  <p style="color:#888;font-size:13px;text-align:center">É rápido — menos de 30 segundos. Seu código de atleta é <strong>${athId}</strong>.</p>
  <hr style="border:none;border-top:1px solid #eee;margin:24px 0">
  <p style="color:#aaa;font-size:12px;text-align:center">Precisa de ajuda? Responda este email ou fale pelo WhatsApp.</p>
</td></tr>
<tr><td style="background:#1a3a8a;padding:16px;text-align:center">
  <p style="color:rgba(255,255,255,.6);font-size:12px;margin:0">Grupo Hiperativo | contato@ghiperativo.com.br</p>
</td></tr>
</table></td></tr></table>
</body></html>`;
}

// ── Configurar credenciais WhatsApp API ──────────────────────────────────────
function configurarWhatsApp() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();

  const opcao = ui.alert(
    '📲 Configurar WhatsApp API',
    'Escolha a plataforma:\n\n• SIM → Z-API\n• NÃO → Evolution API\n\nCancele se não quiser configurar agora.',
    ui.ButtonSet.YES_NO_CANCEL
  );

  if (opcao === ui.Button.CANCEL) return;

  if (opcao === ui.Button.YES) {
    // Z-API
    const r1 = ui.prompt('Z-API — Passo 1/3', 'Instance ID (ex: 3ABC123...):', ui.ButtonSet.OK_CANCEL);
    if (r1.getSelectedButton() !== ui.Button.OK) return;
    if (r1.getResponseText().trim()) props.setProperty('ZAPI_INSTANCE', r1.getResponseText().trim());

    const r2 = ui.prompt('Z-API — Passo 2/3', 'Token:', ui.ButtonSet.OK_CANCEL);
    if (r2.getSelectedButton() !== ui.Button.OK) return;
    if (r2.getResponseText().trim()) props.setProperty('ZAPI_TOKEN', r2.getResponseText().trim());

    const r3 = ui.prompt('Z-API — Passo 3/3 (opcional)', 'Client-Token (se exigido pelo plano):', ui.ButtonSet.OK_CANCEL);
    if (r3.getSelectedButton() === ui.Button.OK && r3.getResponseText().trim()) {
      props.setProperty('ZAPI_CLIENT_TOKEN', r3.getResponseText().trim());
    }
    ui.alert('✅ Z-API configurado! Próximo cadastro já vai enviar WhatsApp.');
  } else {
    // Evolution API
    const r1 = ui.prompt('Evolution API — Passo 1/3', 'URL base (ex: https://evo.seusite.com):', ui.ButtonSet.OK_CANCEL);
    if (r1.getSelectedButton() !== ui.Button.OK) return;
    if (r1.getResponseText().trim()) props.setProperty('EVOLUTION_URL', r1.getResponseText().trim());

    const r2 = ui.prompt('Evolution API — Passo 2/3', 'API Key (apikey):', ui.ButtonSet.OK_CANCEL);
    if (r2.getSelectedButton() !== ui.Button.OK) return;
    if (r2.getResponseText().trim()) props.setProperty('EVOLUTION_TOKEN', r2.getResponseText().trim());

    const r3 = ui.prompt('Evolution API — Passo 3/3', 'Instance name:', ui.ButtonSet.OK_CANCEL);
    if (r3.getSelectedButton() !== ui.Button.OK) return;
    if (r3.getResponseText().trim()) props.setProperty('EVOLUTION_INSTANCE', r3.getResponseText().trim());

    ui.alert('✅ Evolution API configurada! Próximo cadastro já vai enviar WhatsApp.');
  }
}

// ── Wrapper público para google.script.run ─────────────────────────────────────
// IMPORTANTE: funções com _ são privadas e não funcionam via google.script.run
// Esta função pública permite que o cadastro.html chame _gerarUrlOAuth


// ==== PATCH_Strava_Fixes v2 — 2026-06-21T16:51:10.706Z ====
/**
 * ═══════════════════════════════════════════════════════════════════════════
 * HIPERATIVO V3 — PATCH_Strava_Fixes.gs  (v2.0 — 21/06/2026)
 *
 * COMO APLICAR:
 *   1. Abra o Editor Apps Script:
 *      https://script.google.com/d/1cm4IbBr3IsAjHtJJNOAAJQf3EnFD66ES5hMmiCaejBQCkNpVOnp5Zy_P/edit
 *   2. Se existir "PATCH_Strava_Fixes" do v1, SUBSTITUA o conteúdo completo por este.
 *      Se não existir, crie "+ Novo arquivo" → "Script" → nome: "PATCH_Strava_Fixes"
 *   3. Em Strava.gs, DELETE exatamente as funções abaixo (estão corrigidas aqui):
 *        _importarTodosAtletas         linha 482
 *        _trocarCodigoPorToken         linha 308
 *        _salvarTokensPlanilha         linha 416
 *        _backupToken                  linha 452  (substituída por persistirCredenciaisStrava)
 *        _getValidAccessToken          linha 331
 *        _refreshAccessToken           linha 393
 *        verificarStatusStravaAtletas  linha 1248
 *        gerarUrlOAuth (DUPLICATA)     linha 1546  ← manter apenas a de linha 303
 *   4. Salve tudo (Ctrl+S)
 *   5. Execute diagnosticoStravaHiperativoV3() e cole o resultado aqui
 *
 * DIFERENÇAS v1 → v2:
 *   [FIX-1] diagnosticoStravaHiperativoV3(): linha 674 chamava _refreshAccessToken(athId,'')
 *           com refresh_token vazio. Substituído por _getValidAccessToken(athId).
 *   [FIX-2] Criada persistirCredenciaisStrava() — ponto único de gravação.
 *           Elimina dupla chamada a supaUpsertToken() que existia no v1.
 *   [FIX-3] _salvarTokensPlanilha() agora SÓ grava na planilha. Supabase fica
 *           centralizado em persistirCredenciaisStrava().
 *   [FIX-4] _refreshAccessToken() valida refresh_token antes de bater na Strava.
 *           Preserva refresh_token antigo se a Strava não retornar novo.
 *           Chama persistirCredenciaisStrava(), não _salvarTokensPlanilha().
 *   [FIX-5] _trocarCodigoPorToken() valida campos obrigatórios da resposta Strava
 *           (access_token, refresh_token, expires_at, athlete.id) antes de persistir.
 *           Chama apenas persistirCredenciaisStrava(), sem duplicar lógica.
 *   [FIX-6] Instrução clara sobre gerarUrlOAuth() duplicata (linha 1546 em Strava.gs).
 *   [FIX-7] _logEvento_() unifica log de INFO, AVISO e ERRO na aba 🔴 ERROS.
 *   [FIX-8] _isRefreshTokenValido_() valida refresh_token antes de qualquer uso.
 * ═══════════════════════════════════════════════════════════════════════════
 */

// ─────────────────────────────────────────────────────────────────────────────
// CONSTANTES CENTRALIZADAS
// Ajuste os nomes de abas aqui se divergirem da sua planilha.
// ─────────────────────────────────────────────────────────────────────────────
// Constantes com prefixo PATCH_ para evitar conflito com declarações existentes no projeto
const PATCH_ABAS_ = {
  CADASTRO: '👤 CADASTRO',
  TOKENS: '🔐 TOKENS',
  STRAVA_STATUS: '📡 STRAVA STATUS',
  ERROS: '🔴 ERROS',
  ATIVIDADES: '🏃 ATIVIDADES',
};

// Colunas da aba CADASTRO (índice 0-based, A=0)
const PATCH_COL_CAD_ = {
  ATH_ID: 0,   // A — 🎫 ID Atleta
  NOME: 1,   // B — 📛 Nome Completo
  EMAIL: 2,   // C — 📧 E-mail
  WHATSAPP: 3,   // D — 📱 WhatsApp
  STRAVA_STATUS: 23,  // X — Sim / Pendente / Reconectar
  STRAVA_ID: 24,  // Y — ID numérico Strava
  STATUS_ATLETA: 25,  // Z — Ativo / Inativo
};

// Colunas da aba TOKENS (índice 0-based)
const PATCH_COL_TOK_ = {
  EXEC_ID: 0,   // A — TOK_...
  ATH_ID: 1,   // B
  NOME: 2,   // C
  ACCESS_TOKEN: 3,   // D
  REFRESH_TOKEN: 4,   // E
  EXPIRES_AT: 5,   // F — Unix seconds ou milliseconds
  SCOPE: 6,   // G
  STRAVA_ID: 7,   // H
  ULT_ATU: 8,   // I — data/hora última atualização
  STATUS: 9,   // J — Ativo / Renovado / Reconectar
};

// Termos que JAMAIS podem ser um ATH_ID real
const TERMOS_CABECALHO_ = [
  'identificação', 'identificacao', 'nome completo', 'id atleta',
  'perfil esportivo', 'saúde', 'saude', 'localização', 'localizacao',
  'emergência', 'emergencia', 'par-q', 'assinatura', 'cadastro',
  'cabeça', 'cabeca', 'coração', 'coracao', 'corpo',
  'programa', 'provas', 'plano',
];


// ─────────────────────────────────────────────────────────────────────────────
// VALIDADORES
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Valida se um valor é um ATH_ID real.
 * Aceita: ATH029112, ATH_1781116630575, ATHDE549E0A, etc.
 * Rejeita: vazio, null, cabeçalhos visuais, strings com emojis.
 */
function _isAthIdValido_(id) {
  if (id === null || id === undefined) return false;
  const limpo = String(id).trim();
  if (limpo === '') return false;
  if (!/^ATH(_)?[A-Z0-9]+$/i.test(limpo)) return false;

  const lower = limpo.toLowerCase();
  for (let i = 0; i < TERMOS_CABECALHO_.length; i++) {
    if (lower.includes(TERMOS_CABECALHO_[i])) return false;
  }
  return true;
}

/**
 * Valida se um refresh_token é utilizável.
 * Não chama a API — só verifica se o valor é uma string não vazia com comprimento mínimo.
 * Tokens Strava têm pelo menos 20 caracteres.
 */
function _isRefreshTokenValido_(rt) {
  if (!rt || typeof rt !== 'string') return false;
  const limpo = rt.trim();
  return limpo.length >= 20;
}


// ─────────────────────────────────────────────────────────────────────────────
// LOG UNIFICADO
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Registra evento na aba 🔴 ERROS.
 * nivel: 'ERRO' | 'AVISO' | 'INFO'
 * Nunca lança exceção — log não pode travar o fluxo principal.
 */
function _logEvento_(nivel, funcao, athId, mensagem, detalhe) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const aba = ss.getSheetByName(PATCH_ABAS_.ERROS);
    if (!aba) return;
    aba.appendRow([
      new Date(),
      nivel || 'INFO',
      funcao || '',
      athId || 'SISTEMA',
      mensagem || '',
      String(detalhe || '').substring(0, 500), // limita para não explodir a célula
    ]);
  } catch (e) {
    console.error('[PATCH] _logEvento_ falhou:', e.message);
  }
}

// Atalhos para compatibilidade com código anterior
function _logErro_(funcao, athId, mensagem, detalhe) {
  _logEvento_('ERRO', funcao, athId, mensagem, detalhe);
}


// ─────────────────────────────────────────────────────────────────────────────
// LEITURA DE TOKEN (helper de leitura — não escreve)
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Retorna a linha de token de um athId na aba TOKENS.
 * Retorna null se não encontrado ou ATH_ID inválido.
 */
function _getTokenRow_(athId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName(PATCH_ABAS_.TOKENS);
  if (!aba) return null;

  const dados = aba.getDataRange().getValues();
  let encontrado = null;
  for (let i = 1; i < dados.length; i++) {
    const id = String(dados[i][PATCH_COL_TOK_.ATH_ID] || '').trim();
    if (!_isAthIdValido_(id)) continue;
    if (id !== athId) continue;
    encontrado = {
      rowIdx: i + 1, // 1-based para setValues
      execId: String(dados[i][PATCH_COL_TOK_.EXEC_ID] || '').trim(),
      accessToken: String(dados[i][PATCH_COL_TOK_.ACCESS_TOKEN] || '').trim(),
      refreshToken: String(dados[i][PATCH_COL_TOK_.REFRESH_TOKEN] || '').trim(),
      expiresAt: Number(dados[i][PATCH_COL_TOK_.EXPIRES_AT]) || 0,
      scope: String(dados[i][PATCH_COL_TOK_.SCOPE] || '').trim(),
      stravaId: String(dados[i][PATCH_COL_TOK_.STRAVA_ID] || '').trim(),
      nome: String(dados[i][PATCH_COL_TOK_.NOME] || '').trim(),
      status: String(dados[i][PATCH_COL_TOK_.STATUS] || '').trim(),
    };
    break;
  }

  // ScriptProperties e Supabase podem conter um refresh_token mais novo que a
  // planilha. Isso acontece porque a Strava rotaciona o refresh_token a cada
  // renovação. Sempre escolhe a cópia com expires_at mais recente e preserva o
  // rowIdx da planilha para que a próxima persistência reconcilie as fontes.
  const props = PropertiesService.getScriptProperties();
  const propRefreshNovo = String(props.getProperty('RT_' + athId) || '').trim();
  const propExpiresNovo = Number(props.getProperty('EX_' + athId) || 0);
  const propRefreshLegado = String(props.getProperty('TOK_REFRESH_' + athId) || '').trim();
  const propExpiresLegado = Number(props.getProperty('TOK_EXPIRES_' + athId) || 0);
  const usarLegado = _isRefreshTokenValido_(propRefreshLegado) &&
    _expiresAtSeg_(propExpiresLegado) > _expiresAtSeg_(propExpiresNovo);
  const propRefresh = usarLegado ? propRefreshLegado : propRefreshNovo;
  const propExpires = usarLegado ? propExpiresLegado : propExpiresNovo;
  const propAccess = usarLegado
    ? String(props.getProperty('TOK_ACCESS_' + athId) || '').trim()
    : String(props.getProperty('AT_' + athId) || '').trim();
  if (_isRefreshTokenValido_(propRefresh) &&
      (!encontrado || _expiresAtSeg_(propExpires) >= _expiresAtSeg_(encontrado.expiresAt))) {
    encontrado = Object.assign({}, encontrado || { rowIdx: -1 }, {
      accessToken: propAccess,
      refreshToken: propRefresh,
      expiresAt: propExpires,
      scope: encontrado ? encontrado.scope : '',
      stravaId: encontrado ? encontrado.stravaId : '',
      nome: encontrado ? encontrado.nome : '',
      status: 'Renovado'
    });
  }

  if (typeof supaGetRefresh === 'function') {
    const supaRow = supaGetRefresh(athId);
    if (supaRow && _isRefreshTokenValido_(supaRow.refresh_token) &&
        (!encontrado || _expiresAtSeg_(supaRow.expires_at) >= _expiresAtSeg_(encontrado.expiresAt))) {
      encontrado = Object.assign({}, encontrado || { rowIdx: -1 }, {
        accessToken: String(supaRow.access_token || '').trim(),
        refreshToken: String(supaRow.refresh_token).trim(),
        expiresAt: Number(supaRow.expires_at) || 0,
        scope: encontrado ? encontrado.scope : '',
        stravaId: encontrado ? encontrado.stravaId : '',
        nome: encontrado ? encontrado.nome : '',
        status: 'Renovado'
      });
      console.log('[PATCH] _getTokenRow_: usando copia mais recente do Supabase para ' + athId);
    }
  }
  return encontrado;
}

function _expiresAtSeg_(valor) {
  const n = Number(valor) || 0;
  return n > 1e12 ? Math.floor(n / 1000) : Math.floor(n);
}


// ─────────────────────────────────────────────────────────────────────────────
// PERSISTÊNCIA NA PLANILHA (somente sheet — sem Supabase)
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Salva ou atualiza token na aba 🔐 TOKENS.
 * ATENÇÃO: esta função NÃO chama Supabase.
 * Para persistência completa, use persistirCredenciaisStrava().
 *
 * @param {string} athId
 * @param {object} tokenData  { accessToken, refreshToken, expiresAt, scope, stravaId, nome }
 */
function _salvarTokensPlanilha(athId, tokenData) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) {
    _logEvento_('AVISO', '_salvarTokensPlanilha', athId,
      'Outra gravação de token está em andamento; preservando cópias existentes', '');
    return false;
  }
  try {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName(PATCH_ABAS_.TOKENS);
  if (!aba) {
    _logEvento_('ERRO', '_salvarTokensPlanilha', athId,
      'Aba TOKENS não encontrada', PATCH_ABAS_.TOKENS);
    return false;
  }

  // Não salva se access_token estiver ausente
  if (!tokenData.accessToken) {
    _logEvento_('ERRO', '_salvarTokensPlanilha', athId,
      'Tentativa de salvar token sem access_token', '');
    return false;
  }

  const existing = _getTokenRow_(athId);
  const refreshSeguro = _isRefreshTokenValido_(tokenData.refreshToken)
    ? String(tokenData.refreshToken).trim()
    : (existing && _isRefreshTokenValido_(existing.refreshToken) ? existing.refreshToken : '');
  if (!_isRefreshTokenValido_(refreshSeguro)) {
    _logEvento_('ERRO', '_salvarTokensPlanilha', athId,
      'Gravação cancelada para não substituir refresh_token por valor vazio', '');
    return false;
  }
  const agora = new Date();
  const execId = (existing && existing.execId) ||
    ('TOK_' + Utilities.getUuid().substring(0, 8).toUpperCase());
  const novaLinha = [
    execId,
    athId,
    tokenData.nome || '',
    tokenData.accessToken,
    refreshSeguro,
    _expiresAtSeg_(tokenData.expiresAt),
    tokenData.scope || '',
    tokenData.stravaId || '',
    agora,
    'Ativo',
  ];

  if (existing && existing.rowIdx > 0) {
    aba.getRange(existing.rowIdx, 1, 1, novaLinha.length).setValues([novaLinha]);
    console.log('[PATCH] _salvarTokensPlanilha: UPDATE linha ' + existing.rowIdx + ' → ' + athId);
  } else {
    aba.appendRow(novaLinha);
    console.log('[PATCH] _salvarTokensPlanilha: INSERT → ' + athId);
  }
  return true;
  } finally {
    lock.releaseLock();
  }
}


// ─────────────────────────────────────────────────────────────────────────────
// PERSISTÊNCIA CENTRALIZADA (planilha + Supabase + CADASTRO + backup)
// ─────────────────────────────────────────────────────────────────────────────

/**
 * PONTO ÚNICO DE GRAVAÇÃO de credenciais Strava.
 *
 * Executa na ordem:
 *   1. Salva na aba 🔐 TOKENS (planilha)
 *   2. Upsert no Supabase public.tokens_strava
 *   3. Atualiza status no CADASTRO (strava_ok + strava_id)
 *   4. Atualiza public.atletas no Supabase (strava_ok = 'Conectado')
 *   5. Cria backup via _backupToken()
 *   6. Log de sucesso ou erro em cada etapa
 *
 * @param {string} athId
 * @param {object} tokenData  { accessToken, refreshToken, expiresAt, scope, stravaId, nome }
 * @param {string} origem     'oauth_callback' | 'refresh_token' | 'manual'
 * @returns {boolean} true se ao menos uma cópia durável foi salva com sucesso
 */
function persistirCredenciaisStrava(athId, tokenData, origem) {
  if (!_isAthIdValido_(athId)) {
    _logEvento_('ERRO', 'persistirCredenciaisStrava', athId,
      'ATH_ID inválido — persistência cancelada', '');
    return false;
  }
  if (!tokenData || !tokenData.accessToken) {
    _logEvento_('ERRO', 'persistirCredenciaisStrava', athId,
      'tokenData sem access_token — persistência cancelada', origem);
    return false;
  }

  const erros = [];
  let salvo = false;
  let backupLocalSalvo = false;
  let supabaseSalvo = false;

  // Primeira gravação: cópia de recuperação atômica no próprio projeto.
  // Nunca remove o refresh_token anterior enquanto o novo conjunto não foi
  // validado. Assim, uma falha posterior na planilha ou no Supabase não perde
  // a credencial recém-rotacionada pela Strava.
  if (_isRefreshTokenValido_(tokenData.refreshToken)) {
    try {
      PropertiesService.getScriptProperties().setProperties({
        ['AT_' + athId]: String(tokenData.accessToken),
        ['RT_' + athId]: String(tokenData.refreshToken),
        ['EX_' + athId]: String(_expiresAtSeg_(tokenData.expiresAt))
      }, false);
      backupLocalSalvo = true;
    } catch (e) {
      erros.push('ScriptProperties: ' + e.message);
      _logEvento_('ERRO', 'persistirCredenciaisStrava', athId,
        'Falha na cópia local; continuando com Planilha e Supabase', e.message);
    }
  }

  // ── 1. Planilha ─────────────────────────────────────
  try {
    const ok = _salvarTokensPlanilha(athId, tokenData);
    if (ok) {
      salvo = true;
      _logEvento_('INFO', 'persistirCredenciaisStrava', athId,
        'Token salvo na planilha [' + origem + ']', '');
    } else {
      erros.push('Planilha: _salvarTokensPlanilha retornou false');
    }
  } catch (e) {
    erros.push('Planilha: ' + e.message);
    _logEvento_('ERRO', 'persistirCredenciaisStrava', athId,
      'Falha ao salvar na planilha', e.message);
  }

  // ── 2. Supabase tokens_strava ────────────────────────
  if (typeof supaUpsertToken === 'function') {
    try {
      if (typeof supaGarantirAtleta === 'function') supaGarantirAtleta(athId);
      supabaseSalvo = supaUpsertToken(
        athId,
        tokenData.nome || '',
        tokenData.accessToken,
        tokenData.refreshToken || '',
        tokenData.expiresAt || 0
      );
      if (supabaseSalvo) {
        _logEvento_('INFO', 'persistirCredenciaisStrava', athId,
          'supaUpsertToken OK [' + origem + ']', '');
      } else {
        erros.push('Supabase tokens_strava: gravação não confirmada');
        _logEvento_('AVISO', 'persistirCredenciaisStrava', athId,
          'Supabase não confirmou o backup; outras cópias preservadas', origem);
      }
    } catch (e) {
      erros.push('Supabase tokens_strava: ' + e.message);
      _logEvento_('ERRO', 'persistirCredenciaisStrava', athId,
        'Falha em supaUpsertToken', e.message);
    }
  } else {
    console.warn('[PATCH] persistirCredenciaisStrava: Supabase não disponível para ' + athId);
  }

  // ── 3. CADASTRO: strava_ok + strava_id ──────────────
  if (typeof _atualizarStatusCadastro === 'function') {
    try {
      _atualizarStatusCadastro(athId, true, tokenData.stravaId || '');
    } catch (e) {
      erros.push('CADASTRO: ' + e.message);
      _logEvento_('ERRO', 'persistirCredenciaisStrava', athId,
        'Falha em _atualizarStatusCadastro', e.message);
    }
  }

  // ── 4. Supabase atletas: strava_ok = 'Conectado' ────
  if (typeof supaAtualizarStravaOk === 'function') {
    try {
      supaAtualizarStravaOk(athId, 'Conectado');
    } catch (e) {
      erros.push('Supabase atletas: ' + e.message);
      _logEvento_('ERRO', 'persistirCredenciaisStrava', athId,
        'Falha em supaAtualizarStravaOk', e.message);
    }
  }

  // ── 5. Backup ────────────────────────────────────────
  if (typeof _backupToken === 'function') {
    try {
      _backupToken(
        athId,
        tokenData.refreshToken || '',
        tokenData.accessToken,
        _expiresAtSeg_(tokenData.expiresAt)
      );
    } catch (e) {
      // Backup é secundário — não bloqueia, não loga como erro crítico
      console.warn('[PATCH] _backupToken falhou para ' + athId + ': ' + e.message);
    }
  }

  if (erros.length > 0) {
    console.warn('[PATCH] persistirCredenciaisStrava [' + athId + '] erros parciais: ' + erros.join(' | '));
  }

  return salvo || backupLocalSalvo || supabaseSalvo;
}


// ─────────────────────────────────────────────────────────────────────────────
// FLUXO OAUTH — troca de authorization code por token
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Troca o authorization code Strava pelo par access_token / refresh_token.
 * Valida todos os campos obrigatórios antes de persistir.
 * Chama APENAS persistirCredenciaisStrava() — sem duplicar lógica de gravação.
 *
 * @param {string} athId  ATH_ID válido
 * @param {string} code   Authorization code recebido no callback OAuth
 * @returns {object} tokenData com todos os campos preenchidos
 */
function _trocarCodigoPorToken(athId, code) {
  if (!_isAthIdValido_(athId)) {
    throw new Error('[_trocarCodigoPorToken] ATH_ID inválido: ' + athId);
  }
  if (!code || String(code).trim() === '') {
    throw new Error('[_trocarCodigoPorToken] Authorization code ausente para ' + athId);
  }

  const credenciais = _getStravaAppCredentials_();
  const clientId = credenciais.clientId;
  const clientSecret = credenciais.clientSecret;

  if (!clientId || !clientSecret) {
    _logEvento_('ERRO', '_trocarCodigoPorToken', athId,
      'Credenciais Strava não configuradas', '');
    throw new Error('Configure as credenciais Strava em Config.gs > setCredenciaisStrava()');
  }

  // 1. Requisição para a Strava
  let resp;
  try {
    resp = UrlFetchApp.fetch('https://www.strava.com/oauth/token', {
      method: 'post',
      payload: {
        client_id: clientId,
        client_secret: clientSecret,
        code: code,
        grant_type: 'authorization_code',
      },
      muteHttpExceptions: true,
    });
  } catch (e) {
    _logEvento_('ERRO', '_trocarCodigoPorToken', athId, 'Erro de rede', e.message);
    throw e;
  }

  const httpStatus = resp.getResponseCode();
  let body;
  try {
    body = JSON.parse(resp.getContentText());
  } catch (e) {
    throw new Error('[_trocarCodigoPorToken] Resposta Strava não é JSON válido (HTTP ' + httpStatus + ')');
  }

  if (httpStatus !== 200) {
    const msg = 'Strava HTTP ' + httpStatus + ': ' + (body.message || body.error || JSON.stringify(body));
    _logEvento_('ERRO', '_trocarCodigoPorToken', athId, msg, '');
    throw new Error(msg);
  }

  // 2. Valida campos obrigatórios
  const camposFaltando = [];
  if (!body.access_token) camposFaltando.push('access_token');
  if (!body.refresh_token) camposFaltando.push('refresh_token');
  if (!body.expires_at) camposFaltando.push('expires_at');
  if (!body.athlete || !body.athlete.id) camposFaltando.push('athlete.id');

  if (camposFaltando.length > 0) {
    const msg = 'Resposta Strava incompleta — campos ausentes: ' + camposFaltando.join(', ');
    _logEvento_('ERRO', '_trocarCodigoPorToken', athId, msg, JSON.stringify(body));
    throw new Error(msg);
  }

  // 3. Monta tokenData
  const nomeAtleta = typeof _getNomeAtleta === 'function' ? _getNomeAtleta(athId) : '';
  const nomeStra = body.athlete.firstname + ' ' + (body.athlete.lastname || '');

  const tokenData = {
    accessToken: body.access_token,
    refreshToken: body.refresh_token,
    expiresAt: body.expires_at * 1000, // converte Unix seconds → ms
    scope: body.scope || '',
    stravaId: String(body.athlete.id),
    nome: nomeStra.trim() || nomeAtleta,
  };

  // 4. Persistência centralizada (planilha + Supabase + CADASTRO + backup)
  const ok = persistirCredenciaisStrava(athId, tokenData, 'oauth_callback');
  if (!ok) {
    _logEvento_('AVISO', '_trocarCodigoPorToken', athId,
      'Token obtido mas persistência na planilha falhou', '');
  }

  console.log('[PATCH] _trocarCodigoPorToken OK — ' + athId + ' stravaId=' + tokenData.stravaId);
  return tokenData;
}


// ─────────────────────────────────────────────────────────────────────────────
// REFRESH TOKEN
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Obtém access_token válido para um athId.
 * Threshold de 10 minutos: se expires_at < agora+10min, faz refresh.
 * Nunca pede reconexão se refresh_token existir e for válido.
 *
 * @param {string} athId
 * @returns {string|null} access_token válido ou null
 */
function _getValidAccessToken(athId) {
  const MARGEM_MS = 10 * 60 * 1000; // 10 minutos

  if (!_isAthIdValido_(athId)) {
    console.error('[PATCH] _getValidAccessToken: ATH_ID inválido: ' + athId);
    return null;
  }

  const row = _getTokenRow_(athId);
  if (!row) {
    console.warn('[PATCH] _getValidAccessToken: nenhum token para ' + athId);
    return null;
  }

  const agora = Date.now();

  // Token ainda válido com margem? (normaliza expiresAt para seg, independente de unidade)
  const _expiresAt_s_ = row.expiresAt > 1e12 ? Math.floor(row.expiresAt / 1000) : row.expiresAt;
  const _agora_s_     = Math.floor(Date.now() / 1000);
  const _restMin_     = Math.floor((_expiresAt_s_ - _agora_s_) / 60);
  if (row.accessToken && _expiresAt_s_ > _agora_s_ + 300) {
    console.log('[PATCH] _getValidAccessToken: token válido para ' + athId +
      ' (expira em ' + _restMin_ + 'min)');
    return row.accessToken;
  }

  // Token expirado ou expirando — verifica se há refresh_token válido
  if (!_isRefreshTokenValido_(row.refreshToken)) {
    _logEvento_('AVISO', '_getValidAccessToken', athId,
      'Token expirado e refresh_token ausente nas fontes disponíveis; status preservado', '');
    return null;
  }

  // Faz refresh sob lock global. Refresh tokens da Strava são rotativos: duas
  // execuções simultâneas usando o mesmo valor podem invalidar uma à outra.
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);

    // Outra execução pode ter renovado enquanto aguardávamos o lock.
    const atualizado = _getTokenRow_(athId);
    const atualizadoExpira = atualizado ? _expiresAtSeg_(atualizado.expiresAt) : 0;
    if (atualizado && atualizado.accessToken && atualizadoExpira > Math.floor(Date.now() / 1000) + 300) {
      return atualizado.accessToken;
    }

    const refreshSeguro = atualizado && _isRefreshTokenValido_(atualizado.refreshToken)
      ? atualizado.refreshToken
      : row.refreshToken;
    console.log('[PATCH] _getValidAccessToken: token expirando, fazendo refresh para ' + athId);
    const novo = _refreshAccessToken(athId, refreshSeguro);
    return novo ? novo.accessToken : null;
  } catch (e) {
    _logEvento_('ERRO', '_getValidAccessToken', athId, 'Refresh falhou', e.message);
    return null;
  } finally {
    try { lock.releaseLock(); } catch (_) { }
  }
}

/**
 * Força uma única renovação após um HTTP 401, sem iniciar OAuth e sem alterar o
 * status do aluno. Se outra execução já tiver gravado um access_token diferente,
 * reutiliza essa cópia em vez de rotacionar o refresh_token novamente.
 */
function _forcarRefreshAccessToken_(athId, accessTokenRecusado) {
  if (!_isAthIdValido_(athId)) return null;
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    const atual = _getTokenRow_(athId);
    if (!atual) return null;

    const tokenAtual = String(atual.accessToken || '').trim();
    const expiraAtual = _expiresAtSeg_(atual.expiresAt);
    if (tokenAtual && tokenAtual !== String(accessTokenRecusado || '').trim() &&
        expiraAtual > Math.floor(Date.now() / 1000) + 60) {
      return tokenAtual;
    }

    if (!_isRefreshTokenValido_(atual.refreshToken)) {
      _logEvento_('ERRO', '_forcarRefreshAccessToken_', athId,
        'HTTP 401 e nenhum refresh_token válido nas fontes de segurança; status preservado', '');
      return null;
    }
    const novo = _refreshAccessToken(athId, atual.refreshToken);
    return novo ? novo.accessToken : null;
  } catch (e) {
    _logEvento_('ERRO', '_forcarRefreshAccessToken_', athId,
      'Renovação de recuperação após HTTP 401 falhou; status preservado', e.message);
    return null;
  } finally {
    try { lock.releaseLock(); } catch (_) { }
  }
}

/**
 * Faz refresh do access_token usando o refresh_token atual.
 *
 * Comportamento:
 *   - Valida refreshTokenAtual ANTES de bater na Strava
 *   - Se Strava retornar novo refresh_token, usa ele
 *   - Se Strava NÃO retornar refresh_token (improvável mas possível),
 *     mantém o refreshTokenAtual que foi passado como parâmetro
 *   - Chama persistirCredenciaisStrava() — NÃO chama _salvarTokensPlanilha() diretamente
 *
 * @param {string} athId
 * @param {string} refreshTokenAtual  valor lido previamente da planilha
 * @returns {object|null} tokenData atualizado ou null em caso de falha
 */
function _refreshAccessToken(athId, refreshTokenAtual) {
  // Valida antes de bater na API
  if (!_isRefreshTokenValido_(refreshTokenAtual)) {
    _logEvento_('ERRO', '_refreshAccessToken', athId,
      'refresh_token inválido ou vazio — refresh cancelado', '');
    return null;
  }

  const credenciais = _getStravaAppCredentials_();
  const clientId = credenciais.clientId;
  const clientSecret = credenciais.clientSecret;

  if (!clientId || !clientSecret) {
    _logEvento_('ERRO', '_refreshAccessToken', athId, 'Credenciais Strava ausentes', '');
    throw new Error('Credenciais Strava não configuradas.');
  }

  let resp;
  try {
    resp = UrlFetchApp.fetch('https://www.strava.com/oauth/token', {
      method: 'post',
      payload: {
        client_id: clientId,
        client_secret: clientSecret,
        refresh_token: refreshTokenAtual,
        grant_type: 'refresh_token',
      },
      muteHttpExceptions: true,
    });
  } catch (e) {
    _logEvento_('ERRO', '_refreshAccessToken', athId, 'Erro de rede no refresh', e.message);
    throw e;
  }

  const httpStatus = resp.getResponseCode();
  let body;
  try {
    body = JSON.parse(resp.getContentText());
  } catch (e) {
    throw new Error('[_refreshAccessToken] Resposta não é JSON (HTTP ' + httpStatus + ')');
  }

  if (httpStatus !== 200) {
    const msg = 'Strava retornou HTTP ' + httpStatus + ' no refresh: ' +
      (body.message || body.error || JSON.stringify(body));
    _logEvento_('ERRO', '_refreshAccessToken', athId, msg, '');
    throw new Error(msg);
  }

  // Valida access_token e expires_at na resposta
  if (!body.access_token || !body.expires_at) {
    const msg = 'Resposta de refresh incompleta — access_token ou expires_at ausentes';
    _logEvento_('ERRO', '_refreshAccessToken', athId, msg, JSON.stringify(body));
    throw new Error(msg);
  }

  // refresh_token: usa o novo se vier; mantém o anterior se não vier
  // (Strava sempre retorna, mas protege contra edge cases)
  const novoRefresh = _isRefreshTokenValido_(body.refresh_token)
    ? body.refresh_token
    : refreshTokenAtual;

  if (!_isRefreshTokenValido_(body.refresh_token)) {
    _logEvento_('AVISO', '_refreshAccessToken', athId,
      'Strava não retornou refresh_token novo — mantendo anterior', '');
  }

  // Recupera nome e stravaId da linha existente para não perder dados
  const rowAtual = _getTokenRow_(athId);
  const stravaId = (rowAtual && rowAtual.stravaId) || '';
  const nomeAtleta = (rowAtual && rowAtual.nome) || (typeof _getNomeAtleta === 'function' ? _getNomeAtleta(athId) : '');

  const tokenData = {
    accessToken: body.access_token,
    refreshToken: novoRefresh,
    expiresAt: body.expires_at * 1000, // ms
    scope: body.scope || (rowAtual ? rowAtual.scope : '') || '',
    stravaId: stravaId,
    nome: nomeAtleta,
  };

  // Persiste de forma centralizada
  persistirCredenciaisStrava(athId, tokenData, 'refresh_token');

  console.log('[PATCH] _refreshAccessToken OK para ' + athId +
    ' → expires=' + new Date(tokenData.expiresAt).toISOString());

  return tokenData;
}

/**
 * Lê as credenciais diretamente de ScriptProperties. Não usa getCfg(), pois o
 * projeto possui implementações legadas dessa função com formatos diferentes.
 */
function _getStravaAppCredentials_() {
  const props = PropertiesService.getScriptProperties();
  return {
    clientId: String(props.getProperty('STRAVA_CLIENT_ID') || '').trim(),
    clientSecret: String(props.getProperty('STRAVA_CLIENT_SECRET') || '').trim()
  };
}


// ─────────────────────────────────────────────────────────────────────────────
// IMPORTAÇÃO DE ATIVIDADES
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Importa atividades de todos os atletas com Strava conectado.
 * Valida ATH_ID antes de processar qualquer linha — nunca processa cabeçalhos.
 */
function _importarTodosAtletas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abaCad = ss.getSheetByName(PATCH_ABAS_.CADASTRO);
  if (!abaCad) {
    _logEvento_('ERRO', '_importarTodosAtletas', 'SISTEMA', 'Aba CADASTRO não encontrada', '');
    return;
  }

  const dados = abaCad.getDataRange().getValues();
  let processados = 0;
  let ignorados = 0;
  let erros = 0;

  for (let i = 0; i < dados.length; i++) {
    const linha = dados[i];
    const athId = String(linha[PATCH_COL_CAD_.ATH_ID] || '').trim();
    const strOk = String(linha[PATCH_COL_CAD_.STRAVA_STATUS] || '').trim().toLowerCase();

    // Valida ATH_ID — ignora cabeçalhos e linhas vazias
    if (!_isAthIdValido_(athId)) {
      ignorados++;
      if (athId !== '') { // só loga se não for linha vazia
        console.log('[PATCH] Linha ' + (i + 1) + ' ignorada (ATH_ID inválido): "' + athId + '"');
      }
      continue;
    }

    // Só importa atletas com Strava conectado
    const conectado = (strOk === 'sim' || strOk === 'ativo' || strOk === 'conectado');
    if (!conectado) {
      ignorados++;
      continue;
    }

    try {
      _importarAtividadesAtleta(athId, 1);
      processados++;
    } catch (e) {
      erros++;
      console.log('[IMPORT-ERRO] ' + athId + ': ' + e.message + ' | stack: ' + (e.stack || '').slice(0, 200));
      _logEvento_('ERRO', '_importarTodosAtletas', athId,
        'Erro ao importar atividades', e.message);
    }

    // Pausa após cada atleta — respeita quota Strava (100 req/15 min)
    Utilities.sleep(3000);
  }

  const resumo = 'processados=' + processados + ' ignorados=' + ignorados + ' erros=' + erros;
  console.log('[PATCH] _importarTodosAtletas: ' + resumo);
  _logEvento_('INFO', '_importarTodosAtletas', 'SISTEMA', 'Importação concluída', resumo);
}


// ─────────────────────────────────────────────────────────────────────────────
// STATUS STRAVA
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Atualiza a aba 📡 STRAVA STATUS com dados reais de atletas válidos.
 * Nunca processa cabeçalhos ou linhas com ATH_ID inválido.
 */
function verificarStatusStravaAtletas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abaCad = ss.getSheetByName(PATCH_ABAS_.CADASTRO);
  const abaStatus = ss.getSheetByName(PATCH_ABAS_.STRAVA_STATUS);

  if (!abaCad) {
    _logEvento_('ERRO', 'verificarStatusStravaAtletas', 'SISTEMA',
      'Aba CADASTRO não encontrada', PATCH_ABAS_.CADASTRO);
    return;
  }
  if (!abaStatus) {
    _logEvento_('ERRO', 'verificarStatusStravaAtletas', 'SISTEMA',
      'Aba STRAVA STATUS não encontrada', PATCH_ABAS_.STRAVA_STATUS);
    return;
  }

  const dadosCad = abaCad.getDataRange().getValues();
  let validos = 0;
  let ignorados = 0;
  const novasLinhas = [];

  for (let i = 0; i < dadosCad.length; i++) {
    const linha = dadosCad[i];
    const athId = String(linha[PATCH_COL_CAD_.ATH_ID] || '').trim();
    const nome = String(linha[PATCH_COL_CAD_.NOME] || '').trim();
    const email = String(linha[PATCH_COL_CAD_.EMAIL] || '').trim();
    const strOk = String(linha[PATCH_COL_CAD_.STRAVA_STATUS] || '').trim();
    const strId = String(linha[PATCH_COL_CAD_.STRAVA_ID] || '').trim();

    if (!_isAthIdValido_(athId)) {
      ignorados++;
      continue;
    }

    validos++;

    // Determina status real cruzando CADASTRO com aba TOKENS
    const tokenRow = _getTokenRow_(athId);
    const temToken = !!(tokenRow && tokenRow.accessToken);
    const strOkLower = strOk.toLowerCase();
    let statusReal;

    if (temToken && (strOkLower === 'sim' || strOkLower === 'ativo' || strOkLower === 'conectado')) {
      statusReal = 'Ativo';
    } else if (strOkLower === 'pendente') {
      statusReal = 'Pendente';
    } else if (strOkLower === 'reconectar') {
      statusReal = 'Reconectar';
    } else if (!temToken) {
      statusReal = 'Pendente';
    } else {
      statusReal = strOk || 'Pendente';
    }

    const ultimaSync = tokenRow && tokenRow.expiresAt
      ? new Date(_expiresAtSeg_(tokenRow.expiresAt) * 1000)
      : '';
    novasLinhas.push([athId, nome, email, statusReal, strId, ultimaSync, '', '', '', '']);
  }

  // Preserva cabeçalho da linha 1, reescreve a partir da linha 2
  const ultimaLinha = abaStatus.getLastRow();
  if (ultimaLinha > 1) {
    abaStatus.getRange(2, 1, ultimaLinha - 1, 10).clearContent();
  }
  if (novasLinhas.length > 0) {
    abaStatus.getRange(2, 1, novasLinhas.length, 10).setValues(novasLinhas);
  }

  const resumo = 'validos=' + validos + ' ignorados=' + ignorados + ' gravados=' + novasLinhas.length;
  console.log('[PATCH] verificarStatusStravaAtletas: ' + resumo);
  _logEvento_('INFO', 'verificarStatusStravaAtletas', 'SISTEMA', 'Status atualizado', resumo);
}


// ─────────────────────────────────────────────────────────────────────────────
// DIAGNÓSTICO
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Diagnóstico completo do sistema Strava.
 *
 * Retorna objeto com:
 *   atletasValidos, atletasInvalidos, tokensPlanilha, tokensSupabase,
 *   semTokenNaSupa, atletasSemStravaId, tokensVencidos, renovados, pendentes
 *
 * Renova tokens vencidos (SOMENTE se refresh_token válido existir).
 * NÃO tenta renovar se refresh_token estiver vazio/inválido.
 */
function diagnosticoStravaHiperativoV3() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const log = [];
  const agoraSeg = Math.floor(Date.now() / 1000);
  const MARGEM_SEG = 10 * 60;

  log.push('════════════════════════════════════════════════');
  log.push(' DIAGNÓSTICO STRAVA — HIPERATIVO V3  (v2)');
  log.push(' ' + new Date().toLocaleString('pt-BR'));
  log.push('════════════════════════════════════════════════');

  // ── 1. CADASTRO ────────────────────────────────────────
  const abaCad = ss.getSheetByName(PATCH_ABAS_.CADASTRO);
  const atletasValidos = [];
  const atletasInvalidos = [];

  if (!abaCad) {
    log.push('\n[ERRO CRÍTICO] Aba CADASTRO não encontrada: ' + PATCH_ABAS_.CADASTRO);
  } else {
    abaCad.getDataRange().getValues().forEach((linha, i) => {
      const id = String(linha[PATCH_COL_CAD_.ATH_ID] || '').trim();
      if (_isAthIdValido_(id)) {
        atletasValidos.push({
          athId: id,
          nome: String(linha[PATCH_COL_CAD_.NOME] || '').trim(),
          stravaOk: String(linha[PATCH_COL_CAD_.STRAVA_STATUS] || '').trim(),
          stravaId: String(linha[PATCH_COL_CAD_.STRAVA_ID] || '').trim(),
          linhaNum: i + 1,
        });
      } else {
        if (id !== '') { // não loga linhas completamente vazias
          atletasInvalidos.push({ valor: id, linhaNum: i + 1 });
        }
      }
    });
  }

  log.push('\n── CADASTRO ─────────────────────────────────────');
  log.push('Atletas válidos:   ' + atletasValidos.length);
  log.push('Linhas ignoradas:  ' + atletasInvalidos.length +
    (atletasInvalidos.length > 0
      ? '\n' + atletasInvalidos.map(x => '   linha ' + x.linhaNum + ': "' + x.valor + '"').join('\n')
      : ''));

  // ── 2. TOKENS (planilha) ────────────────────────────────
  const abaTok = ss.getSheetByName(PATCH_ABAS_.TOKENS);
  const tokensPlanilha = [];

  if (!abaTok) {
    log.push('\n[ERRO] Aba TOKENS não encontrada: ' + PATCH_ABAS_.TOKENS);
  } else {
    abaTok.getDataRange().getValues().forEach((linha, i) => {
      if (i === 0) return;
      const id = String(linha[PATCH_COL_TOK_.ATH_ID] || '').trim();
      if (!_isAthIdValido_(id)) return;
      const rt = String(linha[PATCH_COL_TOK_.REFRESH_TOKEN] || '').trim();
      tokensPlanilha.push({
        athId: id,
        temAccess: !!(String(linha[PATCH_COL_TOK_.ACCESS_TOKEN] || '').trim()),
        refreshToken: rt,
        temRefreshValido: _isRefreshTokenValido_(rt),
        expiresAt: _expiresAtSeg_(linha[PATCH_COL_TOK_.EXPIRES_AT]),
        stravaId: String(linha[PATCH_COL_TOK_.STRAVA_ID] || '').trim(),
      });
    });
  }

  log.push('\n── TOKENS NA PLANILHA ───────────────────────────');
  log.push('Total:             ' + tokensPlanilha.length);
  log.push('Com access_token:  ' + tokensPlanilha.filter(t => t.temAccess).length);
  log.push('Com refresh válido:' + tokensPlanilha.filter(t => t.temRefreshValido).length);

  // ── 3. SUPABASE ─────────────────────────────────────────
  let tokensSupabase = 0;
  let atletasSemStravaId = 0;
  log.push('\n── SUPABASE ─────────────────────────────────────');
  if (typeof diagnosticoSupabase === 'function') {
    try {
      const ds = diagnosticoSupabase();
      if (ds && typeof ds === 'object' && typeof ds.totalTokens !== 'undefined') {
        tokensSupabase = ds.totalTokens || 0;
        atletasSemStravaId = ds.atletasSemStravaId || 0;
        log.push('Tokens no Supabase:    ' + tokensSupabase);
        log.push('Atletas com strava_id: ' + atletasSemStravaId);
      } else {
        log.push('Tokens no Supabase:    0');
        log.push('Atletas com strava_id: 0');
        log.push('Status: tabela vazia ou sem conexões OAuth pós-patch');
      }
    } catch (e) {
      log.push('Tokens no Supabase:    0');
      log.push('Atletas com strava_id: 0');
      log.push('[ERRO Supabase] ' + e.message);
      _logEvento_('ERRO', 'diagnosticoStravaHiperativoV3', 'SISTEMA',
        'diagnosticoSupabase() falhou', e.message);
    }
  } else {
    log.push('Supabase não configurado ou diagnosticoSupabase() indisponível');
  }

  // ── 4. GAPS: CADASTRO vs TOKENS ─────────────────────────
  const comStravaOk = atletasValidos.filter(a => {
    const s = a.stravaOk.toLowerCase();
    return s === 'sim' || s === 'ativo' || s === 'conectado';
  });
  const idsComToken = new Set(tokensPlanilha.map(t => t.athId));
  const semTokenPlanilha = comStravaOk.filter(a => !idsComToken.has(a.athId));

  log.push('\n── GAPS ─────────────────────────────────────────');
  log.push('Strava "Sim" no CADASTRO:   ' + comStravaOk.length);
  log.push('Com token na planilha:      ' + tokensPlanilha.length);
  log.push('Strava OK mas sem token:    ' + semTokenPlanilha.length);
  if (semTokenPlanilha.length > 0) {
    semTokenPlanilha.forEach(a => log.push('   ⚠️  ' + a.athId + ' — ' + a.nome));
  }

  // ── 5. TOKENS VENCIDOS — com renovação segura ──────────
  const vencidos = tokensPlanilha.filter(t => t.expiresAt > 0 && t.expiresAt < agoraSeg + MARGEM_SEG);
  const validos_ = tokensPlanilha.filter(t => t.expiresAt > agoraSeg + MARGEM_SEG);

  log.push('\n── TOKENS VENCIDOS / EXPIRANDO ──────────────────');
  log.push('Válidos:   ' + validos_.length);
  log.push('Vencidos:  ' + vencidos.length);

  let renovados = 0;
  vencidos.forEach(t => {
    // [FIX-1] Não chama _refreshAccessToken(athId, '') — usa _getValidAccessToken
    // que lê o refresh_token correto da planilha e valida antes de chamar Strava
    if (!t.temRefreshValido) {
      log.push('   ⛔ ' + t.athId + ' — refresh_token ausente/inválido → Reconectar');
      _logEvento_('AVISO', 'diagnosticoStravaHiperativoV3', t.athId,
        'Token vencido sem refresh_token válido — reconexão necessária', '');
      // Marca como Reconectar
      if (typeof _atualizarStatusCadastro === 'function') {
        try { _atualizarStatusCadastro(t.athId, false, t.stravaId || ''); } catch (_) { }
      }
      return;
    }

    // Tenta renovar usando _getValidAccessToken (que lê o refresh_token da planilha)
    try {
      const novoToken = _getValidAccessToken(t.athId);
      if (novoToken) {
        renovados++;
        log.push('   ✅ Renovado: ' + t.athId);
      } else {
        log.push('   ❌ Renovação retornou null: ' + t.athId);
        _logEvento_('AVISO', 'diagnosticoStravaHiperativoV3', t.athId,
          '_getValidAccessToken retornou null', '');
      }
    } catch (e) {
      log.push('   ❌ Erro ao renovar ' + t.athId + ': ' + e.message);
      _logEvento_('ERRO', 'diagnosticoStravaHiperativoV3', t.athId,
        'Falha na renovação', e.message);
    }
  });
  log.push('Renovados com sucesso: ' + renovados);

  // ── 6. PENDENTES / RECONECTAR ──────────────────────────
  const pendentes = atletasValidos.filter(a => {
    const s = a.stravaOk.toLowerCase();
    return s === 'pendente' || s === 'reconectar';
  });

  log.push('\n── PENDENTES / RECONECTAR ───────────────────────');
  log.push('Total: ' + pendentes.length);
  pendentes.forEach(a => log.push('   ' + a.athId + ' — ' + a.nome + ' (' + a.stravaOk + ')'));

  // ── 7. TESTE ATH_ID INVÁLIDO ───────────────────────────
  log.push('\n── TESTE DE VALIDAÇÃO (ATH_ID) ──────────────────');
  const testes = [
    { id: '👤 IDENTIFICAÇÃO', esperado: false },
    { id: '📛 Nome Completo', esperado: false },
    { id: '🆔 ID Atleta', esperado: false },
    { id: '', esperado: false },
    { id: 'ATH029112', esperado: true },
    { id: 'ATH_1781116630575', esperado: true },
    { id: 'ATHDE549E0A', esperado: true },
  ];
  let testesOk = 0;
  let testesFail = 0;
  testes.forEach(t => {
    const resultado = _isAthIdValido_(t.id);
    const passou = resultado === t.esperado;
    if (passou) testesOk++;
    else testesFail++;
    log.push('   ' + (passou ? '✅' : '❌') +
      ' "' + t.id + '" → ' + resultado + ' (esperado: ' + t.esperado + ')');
  });
  log.push('Testes: ' + testesOk + '/' + testes.length + ' OK' +
    (testesFail > 0 ? ' — ⚠️ ' + testesFail + ' FALHOU' : ''));

  // ── 8. RESUMO ──────────────────────────────────────────
  log.push('\n════════════════════════════════════════════════');
  log.push(' RESUMO FINAL');
  log.push('  Atletas válidos no CADASTRO:     ' + atletasValidos.length);
  log.push('  Linhas ignoradas (cabeçalhos):   ' + atletasInvalidos.length);
  log.push('  Tokens na planilha:              ' + tokensPlanilha.length);
  log.push('  Tokens no Supabase:              ' + tokensSupabase);
  log.push('  Strava OK mas sem token:         ' + semTokenPlanilha.length);
  log.push('  Tokens vencidos/expirando:       ' + vencidos.length);
  log.push('  Renovados com sucesso:           ' + renovados);
  log.push('  Pendentes / Reconectar:          ' + pendentes.length);
  log.push('  Testes validação ATH_ID:         ' + testesOk + '/' + testes.length);
  log.push('════════════════════════════════════════════════');

  const relatorio = log.join('\n');
  console.log(relatorio);

  // Grava na aba ERROS como INFO
  _logEvento_('INFO', 'diagnosticoStravaHiperativoV3', 'SISTEMA',
    'Diagnóstico concluído', relatorio.substring(0, 500));

  // Exibe UI (falha silenciosa em execução headless)
  try {
    SpreadsheetApp.getUi().alert(
      'Diagnóstico Strava v2',
      relatorio,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (_) { }

  return {
    atletasValidos: atletasValidos.length,
    atletasInvalidos: atletasInvalidos.length,
    tokensPlanilha: tokensPlanilha.length,
    tokensSupabase: tokensSupabase,
    semTokenNaSupa: semTokenPlanilha.length,
    atletasSemStravaId: atletasSemStravaId,
    tokensVencidos: vencidos.length,
    renovados: renovados,
    pendentes: pendentes.length,
    testeValidacao: testesOk + '/' + testes.length,
  };
}


// ─────────────────────────────────────────────────────────────────────────────
// NOTA: gerarUrlOAuth() DUPLICATA em Strava.gs
// ─────────────────────────────────────────────────────────────────────────────
//
// Strava.gs tem duas declarações de gerarUrlOAuth():
//   LINHA 303 → manter (primeira declaração, versão principal)
//   LINHA 1546 → DELETAR (duplicata no final do arquivo)
//
// Como identificar qual é qual:
//   - Abra Strava.gs
//   - Ctrl+G (ir para linha) → 1546
//   - Você verá: function gerarUrlOAuth(athId) {
//   - Delete essa função inteira (da declaração até o } de fechamento)
//   - A versão da linha 303 é mantida intacta
//
// Por que manter a 303: é a versão original que usa _gerarUrlOAuth() internamente.
// Por que deletar a 1546: é cópia que provavelmente foi adicionada acidentalmente
//   ao final do arquivo e causa comportamento imprevisível dependendo da ordem
//   de carregamento dos arquivos pelo V8 do Apps Script.
// ─────────────────────────────────────────────────────────────────────────────


// ═══════════════════════════════════════════════════════════════════════
// DIAGNÓSTICO SEGURO DE TOKEN NAS SCRIPT PROPERTIES
// ═══════════════════════════════════════════════════════════════════════

function diagnosticarTokenScriptProperties(athId) {
  const props = PropertiesService.getScriptProperties();

  const keyAccess = 'TOK_ACCESS_' + athId;
  const keyRefresh = 'TOK_REFRESH_' + athId;
  const keyExpires = 'TOK_EXPIRES_' + athId;

  const hasAccess = !!props.getProperty(keyAccess);
  const hasRefresh = !!props.getProperty(keyRefresh);
  const expiresAt = parseInt(props.getProperty(keyExpires) || '0', 10);

  const now = Math.floor(Date.now() / 1000);
  const expirado = expiresAt > 0 ? (now >= expiresAt) : true;
  const segundos = expiresAt > 0 ? (expiresAt - now) : 0;

  const dataLegivel = expiresAt > 0
    ? new Date(expiresAt * 1000).toLocaleString('pt-BR', { timeZone: 'America/Sao_Paulo' })
    : 'N/A';

  // Verificar na aba TOKENS da planilha
  let existeNaAba = false;
  try {
    const ss = SpreadsheetApp.openById(H.SPREADSHEET_ID);
    const ws = ss.getSheets().find(s => s.getName().includes('TOKENS'));
    if (ws) {
      existeNaAba = ws.getDataRange().getValues().some(r => String(r[0]).trim() === athId);
    }
  } catch (e) {
    Logger.log('[diagnostico] erro aba TOKENS: ' + e.message);
  }

  // Verificar no Supabase
  let existeNoSupabase = false;
  let atletaStravaId = false;
  try {
    const supaUrl = PropertiesService.getScriptProperties().getProperty('SUPABASE_URL')
      || 'https://korlpbcqbknwhpnhzugr.supabase.co';
    const supaKey = PropertiesService.getScriptProperties().getProperty('SUPABASE_KEY');
    if (supaKey) {
      const hdrs = { 'apikey': supaKey, 'Authorization': 'Bearer ' + supaKey };
      const r1 = UrlFetchApp.fetch(
        supaUrl + '/rest/v1/tokens_strava?ath_id=eq.' + athId + '&select=ath_id&limit=1',
        { headers: hdrs, muteHttpExceptions: true }
      );
      existeNoSupabase = JSON.parse(r1.getContentText()).length > 0;
      const r2 = UrlFetchApp.fetch(
        supaUrl + '/rest/v1/atletas?ath_id=eq.' + athId + '&select=strava_id&limit=1',
        { headers: hdrs, muteHttpExceptions: true }
      );
      const rows = JSON.parse(r2.getContentText());
      atletaStravaId = rows.length > 0 && !!rows[0].strava_id;
    }
  } catch (e) {
    Logger.log('[diagnostico] erro Supabase: ' + e.message);
  }

  // Recomendação
  let recomendacao;
  if (!hasAccess && !hasRefresh) {
    recomendacao = 'fazer_oauth_novo';
  } else if (hasRefresh) {
    recomendacao = 'migrar_token_existente';
  } else {
    recomendacao = 'sem_token_valido';
  }

  const resultado = {
    ath_id: athId,
    existe_access_token: hasAccess,
    existe_refresh_token: hasRefresh,
    expires_at: expiresAt,
    expires_at_data_legivel: dataLegivel,
    token_expirado: expirado,
    segundos_para_expirar: segundos,
    existe_na_aba_tokens: existeNaAba,
    existe_no_supabase_tokens: existeNoSupabase,
    atleta_supabase_strava_id_preenchido: atletaStravaId,
    recomendacao: recomendacao
  };

  Logger.log('═══ DIAGNÓSTICO TOKEN SCRIPT PROPERTIES ═══');
  Logger.log(JSON.stringify(resultado, null, 2));
  return resultado;
}

// Wrapper sem parâmetro para execução direta no editor
function _runDiagnosticoATH992736() {
  diagnosticarTokenScriptProperties('ATH992736');
}

// ═══════════════════════════════════════════════════════════════════════
// MIGRAÇÃO SEGURA: Script Properties → Planilha 🔐 TOKENS + Supabase
// ═══════════════════════════════════════════════════════════════════════

function migrarTokenScriptPropertiesParaTokensESupabase(athId) {
  if (!athId || !String(athId).match(/^ATH[A-Z0-9_]+$/i)) {
    throw new Error('athId inválido: ' + athId);
  }

  const props = PropertiesService.getScriptProperties();

  let accessToken = props.getProperty('TOK_ACCESS_' + athId);
  let refreshToken = props.getProperty('TOK_REFRESH_' + athId);
  let expiresAt = parseInt(props.getProperty('TOK_EXPIRES_' + athId) || '0', 10);

  // 1. Sem refresh_token: não há como renovar — encerra com erro controlado
  if (!refreshToken) {
    Logger.log('[migracao] ' + athId + ' → sem_refresh_token_para_migrar');
    throw new Error('sem_refresh_token_para_migrar: ' + athId);
  }

  // 2. Renovar se access_token ausente ou expira em menos de 10 minutos (600s)
  const now = Math.floor(Date.now() / 1000);
  const precisaRenovar = !accessToken || (expiresAt > 0 && (expiresAt - now) < 600);

  if (precisaRenovar) {
    Logger.log('[migracao] ' + athId + ' → access_token ausente ou prestes a expirar. Renovando via refresh...');

    const supaUrl = props.getProperty('SUPABASE_URL') || 'https://korlpbcqbknwhpnhzugr.supabase.co';
    const clientId = props.getProperty('STRAVA_CLIENT_ID');
    const clientSecret = props.getProperty('STRAVA_CLIENT_SECRET');

    if (!clientId || !clientSecret) {
      throw new Error('STRAVA_CLIENT_ID ou STRAVA_CLIENT_SECRET ausente nas Script Properties');
    }

    const rRefresh = UrlFetchApp.fetch(STRAVA_TOKEN_URL, {
      method: 'post',
      contentType: 'application/x-www-form-urlencoded',
      payload: {
        client_id: clientId,
        client_secret: clientSecret,
        grant_type: 'refresh_token',
        refresh_token: refreshToken
      },
      muteHttpExceptions: true
    });

    if (rRefresh.getResponseCode() !== 200) {
      throw new Error('Refresh falhou (HTTP ' + rRefresh.getResponseCode() + '). Faça OAuth novo para ' + athId);
    }

    const refreshData = JSON.parse(rRefresh.getContentText());

    // Atualiza tokens — preserva refresh anterior se Strava não retornar novo
    accessToken = refreshData.access_token;
    refreshToken = refreshData.refresh_token || refreshToken;
    expiresAt = refreshData.expires_at || (now + (refreshData.expires_in || 21600));

    // Persiste tokens renovados nas Script Properties (sem logar valores)
    props.setProperties({
      ['TOK_ACCESS_' + athId]: accessToken,
      ['TOK_REFRESH_' + athId]: refreshToken,
      ['TOK_EXPIRES_' + athId]: String(expiresAt)
    });

    Logger.log('[migracao] ' + athId + ' → refresh concluído. Novo expires_at: ' + expiresAt);
  }

  // 3. Consulta /athlete com access_token garantidamente válido
  const rAthlete = UrlFetchApp.fetch('https://www.strava.com/api/v3/athlete', {
    headers: { 'Authorization': 'Bearer ' + accessToken },
    muteHttpExceptions: true
  });

  if (rAthlete.getResponseCode() !== 200) {
    throw new Error('Falha ao consultar /athlete (HTTP ' + rAthlete.getResponseCode() + '). athId=' + athId);
  }

  const athlete = JSON.parse(rAthlete.getContentText());
  const stravaId = String(athlete.id);
  const nome = ((athlete.firstname || '') + ' ' + (athlete.lastname || '')).trim();

  Logger.log('[migracao] ' + athId + ' → strava_id=' + stravaId + ' nome=' + nome);

  // 4. Monta tokenData sem expor tokens no log
  const tokenData = {
    accessToken: accessToken,
    refreshToken: refreshToken,
    expiresAt: expiresAt,
    scope: STRAVA_SCOPE || 'read,activity:read_all,profile:read_all',
    stravaId: stravaId,
    nome: nome
  };

  // 5. Persiste via função central
  return persistirCredenciaisStrava(athId, tokenData, 'migracao_script_properties');
}


// Wrapper sem parâmetro para executar migração de ATH992736 diretamente no editor
function _runMigracaoATH992736() {
  migrarTokenScriptPropertiesParaTokensESupabase('ATH992736');
}

// ═══════════════════════════════════════════════════════════════════════
// TAREFA 1: LISTAR E DESATIVAR TRIGGERS PERIGOSOS
// ═══════════════════════════════════════════════════════════════════════
function _t1_auditarTriggers() {
  const PERIGOSAS = [
    'triggerImportacaoAutomatica', 'processarFilaStrava',
    'importarAtividades', 'importarTodasAtividades', '_importarTodosAtletas',
    'calcularMetricas', 'atualizarRanking', 'atualizarAnalise',
    'atualizarDashboard', 'processarFila', 'triggerProcessarFila',
    'sincronizarSupabase', 'syncSupabase', 'triggerSync'
  ];
  const SEGURAS = ['onOpen', 'onEdit', 'onFormSubmit', 'onInstall'];

  const triggers = ScriptApp.getProjectTriggers();
  const resultado = [];

  for (const t of triggers) {
    const fn = t.getHandlerFunction();
    const src = t.getTriggerSource().toString();
    const evt = t.getEventType().toString();
    const ehPerigosa = PERIGOSAS.some(p => fn.toLowerCase().includes(p.toLowerCase()));
    const ehSegura = SEGURAS.some(s => fn.toLowerCase() === s.toLowerCase());
    let status = 'mantido';
    if (ehPerigosa && !ehSegura) {
      try { ScriptApp.deleteTrigger(t); status = 'DESATIVADO'; }
      catch (e) { status = 'ERRO: ' + e.message; }
    }
    resultado.push({ fn, src, evt, status });
  }

  Logger.log('═══ AUDITORIA TRIGGERS ═══');
  Logger.log(JSON.stringify(resultado, null, 2));
  return resultado;
}

// ═══════════════════════════════════════════════════════════════════════
// TAREFA 2: BACKUP PLANILHA + ABAS INTERNAS
// ═══════════════════════════════════════════════════════════════════════
function _t2_backupPlanilha() {
  const SSID = (typeof H !== 'undefined' && H.SPREADSHEET_ID) ? H.SPREADSHEET_ID : '1bI5pnt-HOAD5p8M2hqjEsU9P816hc94wy4mqx0J_xOM';
  const ss = SpreadsheetApp.openById(SSID);
  const now = new Date();
  const ts = Utilities.formatDate(now, 'America/Sao_Paulo', 'yyyyMMdd_HHmm');

  // 1. Cópia da planilha inteira
  const copia = ss.copy('HIPERATIVO V3_BACKUP_PRE_NORMALIZACAO_' + ts);
  Logger.log('[backup] Cópia criada: ' + copia.getName() + ' — ID: ' + copia.getId());

  // 2. Duplicar abas internas com sufixo
  const abasParaBackup = [
    '👤 CADASTRO', '🔐 TOKENS', '🏃 ATIVIDADES', '📈 MÉTRICAS', '📡 STRAVA STATUS'
  ];

  const erros = [];
  for (const nome of abasParaBackup) {
    const ws = ss.getSheetByName(nome);
    if (!ws) { erros.push('Aba não encontrada: ' + nome); continue; }
    const nomeBackup = nome + '_BACKUP_PRE_NORMALIZACAO';
    // Remove backup anterior se existir
    const existente = ss.getSheetByName(nomeBackup);
    if (existente) ss.deleteSheet(existente);
    const bk = ws.copyTo(ss);
    bk.setName(nomeBackup);
    Logger.log('[backup] Aba duplicada: ' + nomeBackup);
  }

  Logger.log('[backup] Concluído. Erros: ' + JSON.stringify(erros));
  return { copiaNome: copia.getName(), copiaId: copia.getId(), erros };
}

// ═══════════════════════════════════════════════════════════════════════
// TAREFA 3: AUDITAR ABA 🏃 ATIVIDADES
// ═══════════════════════════════════════════════════════════════════════
function _t3_auditarAtividades() {
  const ss = SpreadsheetApp.openById('1bI5pnt-HOAD5p8M2hqjEsU9P816hc94wy4mqx0J_xOM');
  const sheets = ss.getSheets();
  const todasAbas = sheets.map(s => ({
    nome: s.getName(), lastRow: s.getLastRow(), lastCol: s.getLastColumn(), oculta: s.isSheetHidden()
  }));
  const wsAtiv = ss.getSheetByName('\u{1F3C3} ATIVIDADES');
  if (!wsAtiv) {
    Logger.log('[auditoria] Aba ATIVIDADES nao encontrada');
    Logger.log(JSON.stringify({ erro: 'aba_nao_encontrada', todasAbas }));
    return;
  }
  const temFiltro = !!wsAtiv.getFilter();
  const lastRow = wsAtiv.getLastRow();
  const lastCol = wsAtiv.getLastColumn();
  const maxSamples = Math.min(5, lastRow);
  const amostras = maxSamples > 0 && lastCol > 0
    ? wsAtiv.getRange(1, 1, maxSamples, Math.min(10, lastCol)).getValues() : [];
  const resultado = {
    nomeAba: wsAtiv.getName(), lastRow, lastCol,
    maxRows: wsAtiv.getMaxRows(), frozenRows: wsAtiv.getFrozenRows(),
    temFiltro, amostras, todasAbas
  };
  Logger.log('=== AUDITORIA ATIVIDADES ===');
  Logger.log(JSON.stringify(resultado, null, 2));
}


// Wrappers públicos (GAS não executa funções iniciadas com _)
function runAuditarTriggers() { return _t1_auditarTriggers(); }
function runBackupPlanilha() { return _t2_backupPlanilha(); }
function runAuditarAtividades() { return _t3_auditarAtividades(); }

// ─────────────────────────────────────────────────────────────
// CORREÇÃO SUPABASE_URL — Script Property
// ─────────────────────────────────────────────────────────────

function _corrigirSupabaseUrl() {
  const URL_CORRETA = 'https://korlpbclqgmqvpbrungc.supabase.co';
  const props = PropertiesService.getScriptProperties();
  const atual = props.getProperty('SUPABASE_URL') || '';

  Logger.log('[SUPABASE_URL] Valor atual: ' + atual);

  if (atual === URL_CORRETA) {
    Logger.log('[SUPABASE_URL] Já está correto. Nenhuma ação necessária.');
    return { status: 'OK', url: atual };
  }

  props.setProperty('SUPABASE_URL', URL_CORRETA);
  Logger.log('[SUPABASE_URL] Corrigido de "' + atual + '" para "' + URL_CORRETA + '"');
  return { status: 'CORRIGIDO', de: atual, para: URL_CORRETA };
}

function runCorrigirSupabaseUrl() { return _corrigirSupabaseUrl(); }

// ─────────────────────────────────────────────────────────────
// WRAPPER SEM UI — executa importação diretamente pelo IDE
// ─────────────────────────────────────────────────────────────
function runImportarAtividades() {
  return _importarTodosAtletas();
}
// ─── DIAGNÓSTICO ATIVIDADES ─────────────────────────────────
function runContarAtividades() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName('📋 ATIVIDADES') || ss.getSheetByName('ATIVIDADES');
  if (!aba) { console.log('Aba ATIVIDADES não encontrada. Abas: ' + ss.getSheets().map(s => s.getName()).join(', ')); return; }
  const dados = aba.getDataRange().getValues();
  const header = dados[0];
  console.log('Colunas: ' + header.slice(0, 5).join(' | '));
  console.log('Total linhas (com header): ' + dados.length);
  // Contagem por atleta
  const contagem = {};
  for (let i = 1; i < dados.length; i++) {
    const athId = String(dados[i][1] || dados[i][0] || '').trim();
    if (athId) contagem[athId] = (contagem[athId] || 0) + 1;
  }
  console.log('Por atleta: ' + JSON.stringify(contagem));
  console.log('Total atividades: ' + (dados.length - 1));
}
// ── DIAGNÓSTICO API STRAVA ──────────────────────────────────────────────
function runDiagnosticoAPI() {
  const token = _getValidAccessToken('ATH992736');
  if (!token) { console.log('ERRO: sem token ATH992736'); return; }

  // Testa sem filtro de data para ver se API retorna algo
  const url = 'https://www.strava.com/api/v3/athlete/activities?per_page=5&page=1';
  const resp = UrlFetchApp.fetch(url, {
    headers: { Authorization: 'Bearer ' + token },
    muteHttpExceptions: true
  });
  const code = resp.getResponseCode();
  const body = resp.getContentText();
  let ativs = [];
  try { ativs = JSON.parse(body); } catch (e) { }
  console.log('HTTP: ' + code + ' | atividades retornadas: ' + (Array.isArray(ativs) ? ativs.length : 'erro: ' + body.slice(0, 100)));
  if (Array.isArray(ativs) && ativs.length > 0) {
    console.log('Primeira: ' + ativs[0].name + ' | ' + ativs[0].start_date + ' | tipo: ' + ativs[0].type);
  }

  // Verifica também o nome da aba ATIVIDADES
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abas = ss.getSheets().map(s => s.getName());
  console.log('Abas disponíveis: ' + abas.join(' | '));
  const abaAtiv = ss.getSheets().find(s => s.getName().includes('ATIVIDADES'));
  if (abaAtiv) {
    console.log('Aba ATIVIDADES: "' + abaAtiv.getName() + '" | linhas: ' + abaAtiv.getLastRow());
  } else {
    console.log('ATIVIDADES: não encontrada');
  }
}

// ── DIAGNÓSTICO TOKENS SHEET ─────────────────────────────────────────────────
function runDiagnosticoTokensSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName('🔐 TOKENS');
  if (!aba) { console.log('Aba TOKENS não encontrada'); return; }
  const data = aba.getDataRange().getValues();
  const agora_ms = Date.now();
  const agora_s = Math.floor(agora_ms / 1000);
  console.log('AGORA_MS=' + agora_ms + ' AGORA_S=' + agora_s);
  data.forEach(function (row, i) {
    if (i === 0) return;
    const athId = String(row[0] || '').trim();
    if (!athId || athId.toLowerCase().includes('id') || athId.toLowerCase().includes('atleta')) return;
    const expiresAt = Number(row[4]) || 0;
    const hasAccess = String(row[2] || '').length > 10;
    const hasRefresh = String(row[3] || '').length > 10;
    const valid_if_ms = expiresAt > agora_ms + 300000;
    const valid_if_s = expiresAt > agora_s + 300;
    console.log('TOKENS[' + (i + 1) + '] athId=' + athId +
      ' expiresAt=' + expiresAt +
      ' hasAccess=' + hasAccess +
      ' hasRefresh=' + hasRefresh +
      ' valid_if_ms=' + valid_if_ms +
      ' valid_if_s=' + valid_if_s);
  });
}

// ── TESTE LÓGICA EXPIRAÇÃO ──────────────────────────────────────────────────
function runTestExpiracao() {
  // Testa com ATH029112 para ver os valores reais
  const athId = 'ATH029112';
  const row = _getTokenRow_(athId);
  if (!row) { console.log('[TEST] row NULL para ' + athId); return; }
  
  const agora_ms = Date.now();
  const agora_s  = Math.floor(agora_ms / 1000);
  
  // Detecta se expiresAt está em segundos ou milissegundos
  const expiresAtRaw = row.expiresAt;
  const isMs = expiresAtRaw > 1e12;
  const expiresAt_s = isMs ? Math.floor(expiresAtRaw / 1000) : expiresAtRaw;
  
  console.log('[TEST] expiresAt_raw=' + expiresAtRaw + ' unidade=' + (isMs ? 'ms' : 's'));
  console.log('[TEST] agora_ms=' + agora_ms + ' agora_s=' + agora_s);
  console.log('[TEST] expiresAt_s=' + expiresAt_s + ' restante_s=' + (expiresAt_s - agora_s));
  console.log('[TEST] hasRefresh=' + (row.refreshToken ? row.refreshToken.length > 20 : false));
  console.log('[TEST] hasAccess=' + (row.accessToken ? row.accessToken.length > 20 : false));
  
  // Simula a comparação que _getValidAccessToken faz
  // Caso 1: agora em s, MARGEM em s (300)
  console.log('[TEST] valid_agora_s_margem_s_300: ' + (expiresAtRaw > agora_s + 300));
  // Caso 2: agora em ms, MARGEM em ms (300000)
  console.log('[TEST] valid_agora_ms_margem_ms_300k: ' + (expiresAtRaw > agora_ms + 300000));
  // Caso 3: agora em s, MARGEM em ms (300000) — BUG caso
  console.log('[TEST] valid_agora_s_margem_ms_300k_BUG: ' + (expiresAtRaw > agora_s + 300000));
}
