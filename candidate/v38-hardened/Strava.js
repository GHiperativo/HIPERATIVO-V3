function fixWebAppUrl() {
  const props = PropertiesService.getScriptProperties();
  const old = props.getProperty('WEBAPP_URL');
  const newUrl = 'https://script.google.com/macros/s/AKfycbzvJzeGQXtfpiRu0C3UI4gC7_9LRIJN0hTXZkR9h8hv3t66d4GTUpdwLIoWZk2Ke-4Mtg/exec';
  props.setProperty('WEBAPP_URL', newUrl);
  Logger.log('OLD: ' + old + '\nNEW: ' + newUrl);
  return 'Done';
}

/**
 * ═══════════════════════════════════════════════════════════════════════
 * HIPERATIVO V3 — Strava.gs  (v3.2 — fluxo único cadastro+OAuth 04/06/2026)
 * OAuth2 Strava: link único cadastro → formulário → autorização → extração
 * ═══════════════════════════════════════════════════════════════════════
 */

// ── CONSTANTES STRAVA ───────────────────────────────────────────────────────
const STRAVA_AUTH_URL   = 'https://www.strava.com/oauth/authorize';
const STRAVA_TOKEN_URL  = 'https://www.strava.com/oauth/token';
const STRAVA_API_BASE   = 'https://www.strava.com/api/v3';
const STRAVA_SCOPE      = 'read,activity:read_all,profile:read_all';

// ── 1. GERAR LINK ÚNICO (cadastro + OAuth em um só link) ──────────────────
function gerarLinkStrava() {
  const ui = SpreadsheetApp.getUi();
  const r  = ui.prompt(
    '🔗 Conectar atleta ao Strava',
    'Digite o ID do atleta (ex: ATH_001):',
    ui.ButtonSet.OK_CANCEL
  );
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const athId = (r.getResponseText() || '').trim().toUpperCase();
  if (!athId) { ui.alert('⚠️ ID inválido. Operação cancelada.'); return; }

  try {
    const url = _gerarUrlCadastro(athId);
    ui.alert(
      '✅ Link gerado — envie para o atleta',
      'Cole este link no navegador do atleta:\n\n' + url +
      '\n\nO atleta vai:\n1. Preencher o cadastro completo\n2. Clicar em Conectar Strava\n3. Autorizar o app\n4. Pronto! Dados salvos automaticamente.',
      ui.ButtonSet.OK
    );
    _log(athId, 'INFO', 'gerarLinkStrava', 'Link único de cadastro gerado', '');
  } catch (e) {
    ui.alert('❌ Erro ao gerar link', e.message, ui.ButtonSet.OK);
    _log(athId, 'ERRO', 'gerarLinkStrava', e.message, e.stack || '');
  }
}

function _gerarUrlCadastro(athId) {
  const props     = PropertiesService.getScriptProperties();
  const webAppUrl = props.getProperty('WEBAPP_URL') || '';
  if (!webAppUrl) throw new Error('WEBAPP_URL não configurado. Implante o WebApp primeiro.');
  return webAppUrl + '?cadastro=true&athId=' + encodeURIComponent(athId);
}

function _gerarUrlOAuth(athId) {
  const props     = PropertiesService.getScriptProperties();
  const clientId  = props.getProperty('STRAVA_CLIENT_ID')  || '';
  const webAppUrl = props.getProperty('WEBAPP_URL')         || '';
  if (!clientId)  throw new Error('STRAVA_CLIENT_ID não configurado.');
  if (!webAppUrl) throw new Error('WEBAPP_URL não configurado. Implante o WebApp primeiro.');
  const state  = athId + '__' + Date.now();
  const params = [
    'client_id='    + encodeURIComponent(clientId),
    'redirect_uri=' + encodeURIComponent(webAppUrl),
    'response_type=code',
    'approval_prompt=auto',
    'scope='        + encodeURIComponent(STRAVA_SCOPE),
    'state='        + encodeURIComponent(state)
  ].join('&');
  return STRAVA_AUTH_URL + '?' + params;
}

function _processarFormCadastro(p) {
  const athId = (p.athId || '').toUpperCase().trim();
  if (!athId) return HtmlService.createHtmlOutput(_paginaErro('ID inválido', 'Identificador do atleta ausente.'));

  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(H.SHEETS.CADASTRO);
    if (!sheet) throw new Error('Aba CADASTRO não encontrada.');

    const data   = sheet.getDataRange().getValues();
    let   rowIdx = -1;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][H.CAD.ID - 1] || '').trim().toUpperCase() === athId) { rowIdx = i + 1; break; }
    }

    const agora   = new Date();
    const agoraStr = agora.toLocaleString('pt-BR', { timeZone: 'America/Sao_Paulo' });
    const waLink  = p.waLink || (p.whats ? 'https://wa.me/' + String(p.whats).replace(/\D/g, '') : '');
    const obsBase = p.obs || '';
    const obsText = waLink ? (waLink + (obsBase ? ' | ' + obsBase : '')) : obsBase;

    // MAPEAMENTO CORRETO: 26 valores → cols 2–27 (H.CAD)
    const vals = [
      p.nome    || '',               // col  2  NOME
      p.email   || '',               // col  3  EMAIL
      p.whats   || '',               // col  4  WHATS
      p.nasc    || '',               // col  5  NASC
      p.sexo    || '',               // col  6  SEXO
      p.peso    || '',               // col  7  PESO
      p.altura  || '',               // col  8  ALTURA
      p.mod     || p.modAgg || '',   // col  9  MOD
      p.nivel   || '',               // col 10  NIVEL
      p.obj     || p.objAgg || '',   // col 11  OBJ
      p.freq    || '',               // col 12  FREQ
      p.horario || '',               // col 13  HORARIO
      p.saude   || p.saudeAgg || '', // col 14  SAUDE
      p.lesao   || '',               // col 15  LESAO
      p.med     || '',               // col 16  MED
      p.prova   || '',               // col 17  PROVA
      p.plano   || '',               // col 18  PLANO
      p.cidade  || '',               // col 19  CIDADE
      p.estado  || '',               // col 20  ESTADO
      p.cpf     || '',               // col 21  CPF
      p.origem  || 'Link de Cadastro Web', // col 22 ORIGEM
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
  const props    = PropertiesService.getScriptProperties();
  const webAppUrl = props.getProperty('WEBAPP_URL') || '';
  const adminEmail = props.getProperty('ADMIN_EMAIL') || 'contato@ghiperativo.com.br';

  // Gerar ATH_ID se nao veio no link
  let athId = (p.athId || '').trim().toUpperCase();
  if (!athId) {
    athId = 'ATH' + String(Date.now()).slice(-6);
  }

  try {
    const ss    = SpreadsheetApp.openById(_getSsId());
    const sheet = ss.getSheetByName(H.SHEETS.CADASTRO);
    if (!sheet) return { ok: false, erro: 'Aba CADASTRO nao encontrada. Execute o Setup primeiro.' };

    const data   = sheet.getDataRange().getValues();

    // ── Validação: CPF duplicado ────────────────────────────────────────────
    const cpfNovo = String(p.cpf || '').replace(/\D/g, '').trim();
    if (cpfNovo.length >= 11) {
      for (let ci = 1; ci < data.length; ci++) {
        const existCpf  = String(data[ci][H.CAD.CPF  - 1] || '').replace(/\D/g, '').trim();
        const existId   = String(data[ci][H.CAD.ID   - 1] || '').trim().toUpperCase();
        const existNome = String(data[ci][H.CAD.NOME - 1] || existId).trim();
        if (existCpf && existCpf === cpfNovo && existId !== athId) {
          return { ok: false, erro: '⚠️ CPF já cadastrado para "' + existNome + '". Cada atleta deve usar seu próprio CPF.' };
        }
      }
    }

    let   rowIdx = -1;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][H.CAD.ID - 1] || '').trim().toUpperCase() === athId) {
        rowIdx = i + 1; break;
      }
    }

    const agora   = new Date().toLocaleString('pt-BR', { timeZone: 'America/Sao_Paulo' });
    // wa.me link vem normalizado do formulário; fallback: derivar do número
    const waLink  = p.waLink || (p.whats ? 'https://wa.me/' + String(p.whats).replace(/\D/g, '') : '');
    const obsBase = p.obs || '';
    const obsText = waLink ? (waLink + (obsBase ? ' | ' + obsBase : '')) : obsBase;

    // MAPEAMENTO CORRETO: 26 valores → cols 2–27 (H.CAD)
    const vals = [
      p.nome    || '',               // col  2  NOME
      p.email   || '',               // col  3  EMAIL
      p.whats   || '',               // col  4  WHATS (+5561999999999)
      p.nasc    || '',               // col  5  NASC
      p.sexo    || '',               // col  6  SEXO
      p.peso    || '',               // col  7  PESO
      p.altura  || '',               // col  8  ALTURA
      p.mod     || '',               // col  9  MOD
      p.nivel   || '',               // col 10  NIVEL
      p.obj     || '',               // col 11  OBJ
      p.freq    || '',               // col 12  FREQ
      p.horario || '',               // col 13  HORARIO
      p.saude   || '',               // col 14  SAUDE
      p.lesao   || '',               // col 15  LESAO
      p.med     || '',               // col 16  MED
      p.prova   || '',               // col 17  PROVA
      p.plano   || '',               // col 18  PLANO
      p.cidade  || '',               // col 19  CIDADE
      p.estado  || '',               // col 20  ESTADO
      p.cpf     || '',               // col 21  CPF
      p.origem  || 'Formulário Web', // col 22  ORIGEM
      agora,                         // col 23  DATA_CAD
      'Pendente',                    // col 24  STRAVA_OK
      '',                            // col 25  STRAVA_ID
      'Ativo',                       // col 26  STATUS
      obsText,                       // col 27  OBS (wa.me link + obs)
    ];

    if (rowIdx === -1) {
      sheet.appendRow([athId, ...vals]);
    } else {
      sheet.getRange(rowIdx, 2, 1, vals.length).setValues([vals]);
    }

    _log(athId, 'INFO', 'salvarCadastroAjax', 'Cadastro salvo: ' + (p.nome || 'N/A'), '');

    // Enviar email de confirmacao com link para conectar Strava
    const emailAtleta = p.email || '';
    const nomeAtleta  = (p.nome || '').split(' ')[0] || 'Atleta';
    if (emailAtleta && webAppUrl) {
      try {
        const linkStrava = webAppUrl + '?athId=' + encodeURIComponent(athId);
        const assunto    = 'Bem-vindo(a) ao Hiperativo! Conecte seu Strava';
        const corpo      = _htmlEmailConfirmacaoCadastro(nomeAtleta, athId, linkStrava);
        MailApp.sendEmail({ to: emailAtleta, subject: assunto, htmlBody: corpo, replyTo: adminEmail });
        _log(athId, 'INFO', 'salvarCadastroAjax', 'Email de confirmacao enviado para ' + emailAtleta, '');
      } catch(eEmail) {
        _log(athId, 'AVISO', 'salvarCadastroAjax', 'Falha ao enviar email: ' + eEmail.message, '');
      }
    }

    // Salvar consentimento LGPD com timestamp (LGPD art. 7 - registro de consentimento)
    try {
      if (p.chkLgpd || p.lgpd || p.termos) {
        var consentTs = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
        var consentTxt = 'LGPD:' + consentTs + ' | Strava:' + (p.chkStrava ? 'sim' : 'não');
        sheet.getRange(newRow, H.CAD.OBS).setValue(consentTxt);
      }
    } catch(eLgpd) { /* não bloqueia o cadastro */ }
    return { ok: true, athId: athId, nome: p.nome || '', stravaUrl: _gerarUrlOAuth(athId) };

  } catch (err) {
    _log(athId, 'ERRO', 'salvarCadastroAjax', err.message, err.stack || '');
    return { ok: false, erro: 'Erro ao salvar: ' + err.message };
  }
}

// ── Template email confirmacao de cadastro ─────────────────────────────────────
function _htmlEmailConfirmacaoCadastro(nome, athId, linkStrava) {
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
  <hr style="border:none;border-top:1px solid #eee;margin:24px 0">
  <h3 style="color:#1a3a8a;margin:0 0 8px">🏃 Conecte seu Strava (opcional mas recomendado)</h3>
  <p style="color:#444;line-height:1.6">Conectando sua conta Strava, seu treinador poderá acompanhar seus treinos automaticamente, sem que você precise enviar nada.</p>
  <div style="text-align:center;margin:24px 0">
    <a href="${linkStrava}" style="background:#fc4c02;color:#fff;padding:14px 32px;border-radius:8px;text-decoration:none;font-weight:bold;font-size:16px;display:inline-block">
      ⚡ Conectar meu Strava agora
    </a>
  </div>
  <p style="color:#888;font-size:13px;text-align:center">Você também pode conectar mais tarde acessando o mesmo link acima.<br>Guarde seu código de atleta: <strong>${athId}</strong></p>
</td></tr>
<tr><td style="background:#1a3a8a;padding:16px;text-align:center">
  <p style="color:rgba(255,255,255,.6);font-size:12px;margin:0">Grupo Hiperativo | contato@ghiperativo.com.br</p>
</td></tr>
</table></td></tr></table>
</body></html>`;
}


// ── 4. TROCA CÓDIGO POR TOKEN ─────────────────────────────────────────────────
function _trocarCodigoPorToken(athId, code) {
  const props     = PropertiesService.getScriptProperties();
  const clientId  = props.getProperty('STRAVA_CLIENT_ID')     || '';
  const clientSec = props.getProperty('STRAVA_CLIENT_SECRET') || '';
  if (!clientId || !clientSec) throw new Error('Credenciais Strava não configuradas no sistema.');

  const resp = UrlFetchApp.fetch(STRAVA_TOKEN_URL, {
    method: 'post',
    payload: {
      client_id:     clientId,
      client_secret: clientSec,
      code:          code,
      grant_type:    'authorization_code'
    },
    muteHttpExceptions: true
  });
  if (resp.getResponseCode() !== 200) {
    throw new Error('Strava retornou erro ' + resp.getResponseCode() + ': ' + resp.getContentText().substring(0, 200));
  }
  return JSON.parse(resp.getContentText());
}

// ── 5. REFRESH AUTOMÁTICO DO ACCESS TOKEN ─────────────────────────────────────
function _getValidAccessToken(athId) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(H.SHEETS.TOKENS);
  if (!sheet) throw new Error('Aba de tokens não encontrada. Execute o Setup primeiro.');

  // ── LockService: previne race condition com refresh tokens rotativos do Strava ──
  // Ref: github.com/googleworkspace/apps-script-oauth2/issues/105
  const _lock = LockService.getScriptLock();
  _lock.waitLock(30000); // espera até 30s se outra execução está renovando
  let data;
  try {
    data = sheet.getDataRange().getValues();
  const margemSeg = 600;
  const agora     = Math.floor(Date.now() / 1000);

  for (let i = 1; i < data.length; i++) {
    const rowAthId = String(data[i][H.TOK.ATH_ID - 1] || '').trim().toUpperCase();
    if (rowAthId !== athId.toUpperCase()) continue;

    const accessToken  = String(data[i][H.TOK.ACCESS   - 1] || '').trim();
    const refreshToken = String(data[i][H.TOK.REFRESH  - 1] || '').trim();
    const expiresAt    = Number(data[i][H.TOK.EXPIRES  - 1]) || 0;
    const rowIdx       = i + 1;

    if (!refreshToken) {
      // Fallback 1: PropertiesService (backup local)
      let backupRt = PropertiesService.getScriptProperties().getProperty('RT_' + athId.toUpperCase());
      // Fallback 2: Supabase (backup remoto — último recurso)
      if (!backupRt) {
        try { backupRt = supaGetRefreshToken(athId); } catch(_) {}
      }
      if (!backupRt) throw new Error('Atleta ' + athId + ' sem refresh_token. Reconecte o Strava.');
      _log(athId, 'INFO', '_getValidAccessToken', 'Refresh_token ausente na planilha — restaurando do backup PropertiesService.', '');
      const novoToken = _refreshAccessToken(athId, backupRt);
      sheet.getRange(rowIdx, H.TOK.ACCESS  ).setValue(novoToken.access_token);
      sheet.getRange(rowIdx, H.TOK.REFRESH ).setValue(novoToken.refresh_token || backupRt);
      sheet.getRange(rowIdx, H.TOK.EXPIRES ).setValue(novoToken.expires_at);
      sheet.getRange(rowIdx, H.TOK.ULT_ATU ).setValue(new Date());
      sheet.getRange(rowIdx, H.TOK.STATUS  ).setValue('Restaurado auto');
      // Atualizar backup com o novo refresh_token retornado pelo Strava
      if (novoToken.refresh_token) {
        PropertiesService.getScriptProperties().setProperty('RT_' + athId.toUpperCase(), novoToken.refresh_token);
      }
      _log(athId, 'INFO', '_getValidAccessToken', 'Token restaurado automaticamente. Atleta NÃO precisou reconectar.', '');
      return novoToken.access_token;
    }

    if (accessToken && expiresAt > agora + margemSeg) {
      return accessToken;
    }

    _log(athId, 'INFO', '_getValidAccessToken', 'Renovando token...', '');
    const novoToken = _refreshAccessToken(athId, refreshToken);

    sheet.getRange(rowIdx, H.TOK.ACCESS  ).setValue(novoToken.access_token);
    sheet.getRange(rowIdx, H.TOK.REFRESH ).setValue(novoToken.refresh_token || refreshToken);
    sheet.getRange(rowIdx, H.TOK.EXPIRES ).setValue(novoToken.expires_at);
    sheet.getRange(rowIdx, H.TOK.ULT_ATU ).setValue(new Date());
    sheet.getRange(rowIdx, H.TOK.STATUS  ).setValue('Renovado');

    _log(athId, 'INFO', '_getValidAccessToken', 'Token renovado. Expira: ' + new Date(novoToken.expires_at * 1000).toISOString(), '');
    return novoToken.access_token;
  }
  throw new Error('Atleta ' + athId + ' não encontrado nos tokens. Conecte o Strava primeiro.');
  } finally {
    _lock.releaseLock();
  }
}

function _refreshAccessToken(athId, refreshToken) {
  const props     = PropertiesService.getScriptProperties();
  const clientId  = props.getProperty('STRAVA_CLIENT_ID')     || '';
  const clientSec = props.getProperty('STRAVA_CLIENT_SECRET') || '';
  if (!clientId || !clientSec) throw new Error('Credenciais Strava não configuradas.');

  const resp = UrlFetchApp.fetch(STRAVA_TOKEN_URL, {
    method: 'post',
    payload: {
      client_id:     clientId,
      client_secret: clientSec,
      refresh_token: refreshToken,
      grant_type:    'refresh_token'
    },
    muteHttpExceptions: true
  });
  if (resp.getResponseCode() !== 200) {
    throw new Error('Falha ao renovar token: ' + resp.getContentText().substring(0, 200));
  }
  return JSON.parse(resp.getContentText());
}

// ── 6. SALVAR TOKENS NA PLANILHA ──────────────────────────────────────────────
function _salvarTokensPlanilha(athId, tokenData) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(H.SHEETS.TOKENS);
  if (!sheet) throw new Error('Aba de TOKENS não existe. Execute o Setup.');

  // ── BACKUP DE SEGURANÇA (camada dupla) ────────────────────────────────────
  // Camada 1: PropertiesService (local, instantâneo)
  // Camada 2: Supabase (remoto, sobrevive até a deleção do projeto GAS)
  if (tokenData.refresh_token) {
    PropertiesService.getScriptProperties().setProperty(
      'RT_' + athId.toUpperCase(),
      tokenData.refresh_token
    );
    try { supaSalvarToken(athId, tokenData); } catch(_) {} // silencioso — não bloqueia
  }
  // ─────────────────────────────────────────────────────────────────────────

  const data     = sheet.getDataRange().getValues();
  const athlete  = tokenData.athlete || {};
  const nome     = _getNomeAtleta(athId) || ((athlete.firstname || '') + ' ' + (athlete.lastname || '')).trim();
  const stravaId = athlete.id || '';
  const agora    = new Date();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][H.TOK.ATH_ID - 1] || '').trim().toUpperCase() === athId.toUpperCase()) {
      sheet.getRange(i + 1, H.TOK.ACCESS   ).setValue(tokenData.access_token  || '');
      sheet.getRange(i + 1, H.TOK.REFRESH  ).setValue(tokenData.refresh_token || '');
      sheet.getRange(i + 1, H.TOK.EXPIRES  ).setValue(tokenData.expires_at    || '');
      sheet.getRange(i + 1, H.TOK.STRAVA_ID).setValue(stravaId);
      sheet.getRange(i + 1, H.TOK.ULT_ATU ).setValue(agora);
      sheet.getRange(i + 1, H.TOK.STATUS  ).setValue('Ativo');
      return;
    }
  }

  const execId = 'TOK_' + Utilities.getUuid().substring(0, 8).toUpperCase();
  sheet.appendRow([
    execId, athId, nome,
    tokenData.access_token || '', tokenData.refresh_token || '', tokenData.expires_at || '',
    STRAVA_SCOPE, stravaId, agora, 'Ativo'
  ]);
}

// ── 7. IMPORTAR ATIVIDADES ────────────────────────────────────────────────────
function importarAtividades() {
  const ui = SpreadsheetApp.getUi();
  const r  = ui.prompt('🏃 Importar Atividades', 'ID do atleta (vazio = todos):', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const athId = (r.getResponseText() || '').trim().toUpperCase();
  try {
    let msg;
    if (athId) {
      const n = _importarAtividadesAtleta(athId, 3);
      msg = '✅ ' + n + ' atividades importadas para ' + athId;
    } else {
      msg = _importarTodosAtletas();
    }
    ui.alert('📊 Resultado', msg, ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('❌ Erro', e.message, ui.ButtonSet.OK);
    _log(athId || 'TODOS', 'ERRO', 'importarAtividades', e.message, e.stack || '');
  }
}

function _importarTodosAtletas() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(H.SHEETS.TOKENS);
  if (!sheet) return 'Aba de tokens não encontrada.';
  const data = sheet.getDataRange().getValues();
  const msgs = [];
  for (let i = 1; i < data.length; i++) {
    const athId  = String(data[i][H.TOK.ATH_ID - 1] || '').trim();
    const status = String(data[i][H.TOK.STATUS  - 1] || '').toLowerCase();
    if (!athId || status === 'inativo') continue;
    try {
      const n = _importarAtividadesAtleta(athId, 3);
      msgs.push(athId + ': ' + n + ' atividades');
    } catch (e) {
      msgs.push(athId + ': ERRO — ' + e.message);
    }
  }
  return msgs.length ? msgs.join('\n') : 'Nenhum atleta ativo com Strava conectado.';
}

function _importarAtividadesAtleta(athId, paginas) {
  const accessToken = _getValidAccessToken(athId);
  const nomeAtleta  = _getNomeAtleta(athId);
  paginas = paginas || 3;
  let todas = [];
  for (let pg = 1; pg <= paginas; pg++) {
    const url  = STRAVA_API_BASE + '/athlete/activities?per_page=50&page=' + pg;
    const resp = UrlFetchApp.fetch(url, {
      headers: { Authorization: 'Bearer ' + accessToken },
      muteHttpExceptions: true
    });
    if (resp.getResponseCode() !== 200) break;
    const page = JSON.parse(resp.getContentText());
    if (!page.length) break;
    todas = todas.concat(page);
  }
  return _gravarAtividades(athId, nomeAtleta, todas);
}

function _gravarAtividades(athId, nomeAtleta, atividades) {
  if (!atividades.length) return 0;
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(H.SHEETS.ATIVIDADES);
  if (!sheet) throw new Error('Aba ATIVIDADES não encontrada.');

  const existentes = new Set();
  const dataAtual  = sheet.getDataRange().getValues();
  for (let i = 1; i < dataAtual.length; i++) {
    const sid = String(dataAtual[i][H.ATIV.STRAVA_ID - 1] || '').trim();
    if (sid) existentes.add(sid);
  }

  let count = 0;
  for (const a of atividades) {
    const sid = String(a.id || '');
    if (existentes.has(sid)) continue;

    // ── Distância e tempo ──────────────────────────────────────────────────
    const distM  = a.distance    || 0;
    const distKm = distM / 1000;
    const tempoS = a.moving_time || 0;

    // ── Tipo (definir antes de calcular métricas por esporte) ────────────────
    const tipo = _traduzirTipo(a.sport_type || a.type || '');

    // ── Velocidade ─────────────────────────────────────────────────────────
    // Strava retorna average_speed em m/s (mais preciso que distance/time)
    const velMps = (a.average_speed > 0)
      ? a.average_speed
      : (tempoS > 0 && distM > 0 ? distM / tempoS : 0);
    // km/min: mantido para compatibilidade com cálculos internos
    const velKmMin = velMps > 0 ? Math.round(velMps * 0.06 * 1000) / 1000 : 0;

    // ── Pace / Velocidade por esporte ──────────────────────────────────────
    // Col 15 (PACE_S): s/km corrida | s/100m natação | 0 ciclismo
    const paceSegKm = (velMps > 0 && tipo !== 'Ciclismo') ? Math.round(1000 / velMps) : 0;
    // Col 16 (PACE_FMT): display inteligente por esporte
    const paceFmt   = _formatarVelocidadeDisplay(velMps, tipo);
    // Col 12 (DIST_KM): arredondamento adequado por esporte
    const distKmCalc = _calcularDistKm(distM, tipo);

    // ── Cadência por esporte ───────────────────────────────────────────────
    const cadRaw   = a.average_cadence || 0;
    const cadencia = cadRaw > 0
      ? (tipo === 'Corrida' || tipo === 'Caminhada' || tipo === 'Trail Run'
          ? Math.round(cadRaw * 2)  // corrida: Strava dá 1 pé → ×2 = spm total
          : Math.round(cadRaw))     // ciclismo: rpm direto | natação: braçadas/min
      : '';

    const execId = 'ATIV_' + Utilities.getUuid().substring(0, 8).toUpperCase();
    // FC e calorias como inteiros (sem decimais)
    const fcMed  = a.average_heartrate ? Math.round(a.average_heartrate) : '';
    const fcMax  = a.max_heartrate     ? Math.round(a.max_heartrate)     : '';
    const cal    = a.calories          ? Math.round(a.calories)          : '';
    const elev   = a.total_elevation_gain ? Math.round(a.total_elevation_gain) : '';

    sheet.appendRow([
      execId,                                                    // 1  EXEC_ID
      athId,                                                     // 2  ATH_ID
      nomeAtleta,                                                // 3  NOME
      a.start_date_local ? new Date(a.start_date_local) : '',   // 4  DATA
      tipo,                                                      // 5  TIPO
      'Strava',                                                  // 6  FONTE
      sid,                                                       // 7  STRAVA_ID
      a.name || '',                                              // 8  NOME_ATIV
      tempoS / 86400,                                            // 9  MOV_S (fração de dia → formato [h]:mm:ss no Sheets)
      (a.elapsed_time || 0) / 86400,                            // 10 TOTAL_S (fração de dia → formato [h]:mm:ss no Sheets)
      Math.round(distM),                                         // 11 DIST_M (metros, inteiro)
      distKmCalc,                                                // 12 DIST_KM (1-2 dec por esporte)
      velMps,                                                    // 13 VEL_MPS (m/s raw, Strava)
      velKmMin,                                                  // 14 VEL_KMMIN (km/min, interno)
      paceSegKm,                                                 // 15 PACE_S (s/km corrida | s/100m natação | 0 bike)
      paceFmt,                                                   // 16 PACE_FMT (display inteligente por esporte)
      fcMed,                                                     // 17 FC_MED (bpm, inteiro)
      fcMax,                                                     // 18 FC_MAX (bpm, inteiro)
      elev,                                                      // 19 ELEV (m, inteiro)
      cal,                                                       // 20 CAL (kcal, inteiro)
      cadencia,                                                  // 21 CADENCIA (spm corrida | rpm bike | braçadas natação)
      a.average_watts ? Math.round(a.average_watts) : '',        // 22 POTENCIA (W, inteiro)
      a.map ? (a.map.summary_polyline || '') : '',               // 23 ROTA (polyline)
      new Date(),                                                // 24 IMPORTADO
      ''                                                         // 25 PSE (entrada manual 1-10)
    ]);
    existentes.add(sid);
    count++;
  }
  _log(athId, 'INFO', '_gravarAtividades', count + ' novas de ' + atividades.length + ' recebidas', '');

  // Recalcular métricas se houver novas atividades
  if (count > 0) {
    try {
      recalcularMetricasAposAtividade(athId);
    } catch(e) {
      // Não bloquear importação se métricas falharem
      _log(athId, 'AVISO', '_gravarAtividades', 'Recálculo de métricas falhou: ' + e.message, '');
    }
  }

  return count;
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
    const r  = ui.prompt('👤 Perfil', 'ID do atleta:', ui.ButtonSet.OK_CANCEL);
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
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
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

  // 1. Processar fila inteligente com rate-limiting (Queue.gs)
  try {
    processarFilaStrava();
    _log('SISTEMA', 'INFO', 'triggerImportacaoAutomatica', 'Fila Strava processada.', '');
  } catch(e) {
    _log('SISTEMA', 'ERRO', 'triggerImportacaoAutomatica', 'Erro na fila: ' + e.message, '');
  }

  // 2. Recalcular métricas (VO2máx, zonas, pace) — silencioso
  try {
    calcularMetricasTodos();
    _log('SISTEMA', 'INFO', 'triggerImportacaoAutomatica', 'Métricas recalculadas.', '');
  } catch(e) {
    _log('SISTEMA', 'ERRO', 'triggerImportacaoAutomatica', 'Erro nas métricas: ' + e.message, '');
  }

  // 3. Atualizar abas de ranking
  try {
    atualizarRankingSheet();
    atualizarRankingExpandido();
    _log('SISTEMA', 'INFO', 'triggerImportacaoAutomatica', 'RANKING atualizado.', '');
  } catch(e) {
    _log('SISTEMA', 'AVISO', 'triggerImportacaoAutomatica', 'RANKING não atualizado: ' + e.message, '');
  }

  // 3b. Atualizar análise científica (CTL/ATL/TSB)
  try {
    atualizarAnaliseSheet();
    _log('SISTEMA', 'INFO', 'triggerImportacaoAutomatica', 'ANÁLISE científica atualizada.', '');
  } catch(e) {
    _log('SISTEMA', 'AVISO', 'triggerImportacaoAutomatica', 'ANÁLISE não atualizada: ' + e.message, '');
  }

  // 4. Atualizar aba STRAVA STATUS (se existir)
  try {
    atualizarStravaStatusSheet();
    _log('SISTEMA', 'INFO', 'triggerImportacaoAutomatica', 'STRAVA STATUS atualizado.', '');
  } catch(e) {
    _log('SISTEMA', 'AVISO', 'triggerImportacaoAutomatica', 'STATUS não atualizado: ' + e.message, '');
  }

  // 5. Sincronizar atletas do CADASTRO em todas as abas
  try {
    const novos = sincronizarAtletasEmTodasAbas();
    if (novos > 0) _log('SISTEMA', 'INFO', 'triggerImportacaoAutomatica', novos + ' novos atletas sincronizados.', '');
  } catch(e) {
    _log('SISTEMA', 'AVISO', 'triggerImportacaoAutomatica', 'Sync atletas: ' + e.message, '');
  }

  // 6. Atualizar timestamp no PAINEL
  try {
    _atualizarPainelInterno();
  } catch(e) {
    _log('SISTEMA', 'AVISO', 'triggerImportacaoAutomatica', 'Painel não atualizado: ' + e.message, '');
  }

  _log('SISTEMA', 'INFO', 'triggerImportacaoAutomatica', 'Ciclo completo concluído.', '');
}

function _atualizarStatusCadastro(athId, conectado, stravaId) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
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
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(H.SHEETS.CADASTRO);
    if (!sheet) return athId;
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][H.CAD.ID - 1] || '').trim().toUpperCase() === athId.toUpperCase()) {
        return String(data[i][H.CAD.NOME - 1] || athId).trim();
      }
    }
  } catch (_) {}
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
    Run:'Corrida', Ride:'Ciclismo', Swim:'Natação', Walk:'Caminhada',
    TrailRun:'Trail Run', VirtualRide:'Ciclismo Virtual', VirtualRun:'Corrida Virtual',
    WeightTraining:'Musculação', Yoga:'Yoga', Workout:'Treino', Hike:'Trilha',
    AlpineSki:'Ski', Rowing:'Remo', Kayaking:'Canoagem', Soccer:'Futebol',
    Crossfit:'CrossFit', Elliptical:'Elíptico', StairStepper:'Escada',
    RockClimbing:'Escalada', Surfing:'Surf', Skateboard:'Skate',
    Badminton:'Badminton', Tennis:'Tênis', Volleyball:'Vôlei', Basketball:'Basquete'
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
  const nome = athlete ? ((athlete.firstname||'') + ' ' + (athlete.lastname||'')).trim() : athId;
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
  ${foto?'<img class="avatar" src="'+foto+'" alt="foto">':'<div class="ico">✅</div>'}
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

// ── Wrapper público para google.script.run ─────────────────────────────────────
// IMPORTANTE: funções com _ são privadas e não funcionam via google.script.run
// Esta função pública permite que o cadastro.html chame _gerarUrlOAuth
function gerarUrlOAuth(athId) {
  return _gerarUrlOAuth(athId);
}

/**
 * Gera links de reconexão OAuth para todos os atletas com STATUS = 'Reconectar'.
 * Executa via menu do Apps Script e loga os links no Logger.
 */
function gerarLinksReconexaoTodos() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetTok = ss.getSheetByName('TOKENS');
  if (!sheetTok) { Logger.log('Aba TOKENS não encontrada'); return; }
  
  const data = sheetTok.getDataRange().getValues();
  const links = [];
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const athId  = String(row[H.TOK.ATH_ID  - 1] || '').trim();
    const status = String(row[H.TOK.STATUS   - 1] || '').trim();
    const nome   = String(row[H.CAD ? H.CAD.NOME - 1 : 1] || athId).trim();
    
    if (!athId) continue;
    if (status === 'Reconectar' || status === 'Reconectar' || status === 'Inativo' || status === 'Expirado') {
      const url = _gerarUrlOAuth(athId);
      links.push({ athId: athId, nome: nome, status: status, url: url });
      Logger.log('🔗 ' + nome + ' (' + athId + '): ' + url);
    }
  }
  
  if (links.length === 0) {
    Logger.log('✅ Nenhum atleta precisa reconectar o Strava.');
  } else {
    Logger.log('⚠️ ' + links.length + ' atleta(s) precisam reconectar:');
    links.forEach(function(l) {
      Logger.log('  → ' + l.nome + ': ' + l.url);
    });
  }
  
  return links;
}



// ═══════════════════════════════════════════════════════
// DIAGNÓSTICO — remover após uso
// ═══════════════════════════════════════════════════════
function _limparRegistroTeste() {
  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const shCad    = ss.getSheetByName('CADASTRO');
  const shTok    = ss.getSheetByName('TOKENS');
  const resultado = [];

  // Colunas CADASTRO (1-based): ID=1, NOME=2, CPF=21
  const CAD_ID   = 1;
  const CAD_NOME = 2;
  const CAD_CPF  = 21;

  if (!shCad) { Logger.log('❌ Aba CADASTRO não encontrada'); return; }

  const dataCad = shCad.getDataRange().getValues();
  Logger.log('📋 CADASTRO tem ' + (dataCad.length - 1) + ' registros');

  // Verificar CPFs duplicados
  const cpfMap = {};
  for (let i = 1; i < dataCad.length; i++) {
    const cpf    = String(dataCad[i][CAD_CPF  - 1] || '').replace(/\D/g, '').trim();
    const athId  = String(dataCad[i][CAD_ID   - 1] || '').trim();
    const nome   = String(dataCad[i][CAD_NOME - 1] || '').trim();
    if (!cpf) continue;
    if (!cpfMap[cpf]) cpfMap[cpf] = [];
    cpfMap[cpf].push({ row: i + 1, athId, nome });
  }

  let hasDup = false;
  Object.entries(cpfMap).forEach(function([cpf, entries]) {
    if (entries.length > 1) {
      hasDup = true;
      Logger.log('⚠️ CPF duplicado: ' + cpf);
      entries.forEach(function(e) {
        Logger.log('   Linha ' + e.row + ': ' + e.nome + ' | ID: ' + e.athId);
        resultado.push('Linha ' + e.row + ' | ' + e.nome + ' | ' + e.athId + ' | CPF: ' + cpf);
      });
    }
  });

  if (!hasDup) Logger.log('✅ Nenhum CPF duplicado no CADASTRO');

  // Verificar TOKENS órfãos
  if (shTok) {
    const dataTok = shTok.getDataRange().getValues();
    const cadIds  = dataCad.slice(1).map(function(r){ return String(r[CAD_ID-1]||'').trim().toUpperCase(); }).filter(Boolean);
    Logger.log('🔑 TOKENS tem ' + (dataTok.length - 1) + ' registros');
    const headers = dataTok[0].map(function(h){ return String(h).trim(); });
    const athIdCol = headers.findIndex(function(h){ return h.toUpperCase().includes('ID') || h.toUpperCase() === 'ATH_ID'; });
    if (athIdCol >= 0) {
      for (let i = 1; i < dataTok.length; i++) {
        const tokId = String(dataTok[i][athIdCol] || '').trim().toUpperCase();
        if (tokId && !cadIds.includes(tokId)) {
          Logger.log('🔴 TOKEN órfão linha ' + (i+1) + ': ' + tokId);
          resultado.push('TOKEN_ORFAO | Linha ' + (i+1) + ' | ' + tokId);
        }
      }
    }
  }

  return resultado;
}

function _deleteTestEntry() {
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const shCad = ss.getSheetByName('CADASTRO');
  const shTok = ss.getSheetByName('TOKENS');
  const log   = [];

  // Delete ATH518708 from CADASTRO
  const cadData = shCad.getDataRange().getValues();
  for (let i = cadData.length - 1; i >= 1; i--) {
    if (String(cadData[i][0]).trim() === 'ATH518708') {
      shCad.deleteRow(i + 1);
      log.push('CADASTRO: deleted row ' + (i+1) + ' ATH518708');
    }
  }

  // Delete ATH518708 from TOKENS
  const tokData = shTok.getDataRange().getValues();
  for (let i = tokData.length - 1; i >= 1; i--) {
    if (String(tokData[i][0]).trim() === 'ATH518708') {
      shTok.deleteRow(i + 1);
      log.push('TOKENS: deleted row ' + (i+1) + ' ATH518708');
    }
  }

  // Show Crhystiano tokens (ATH992736)
  const tokData2 = shTok.getDataRange().getValues();
  for (let i = 1; i < tokData2.length; i++) {
    if (String(tokData2[i][0]).includes('ATH992736') || String(tokData2[i][0]).includes('ATH_1781573420732')) {
      log.push('CRHYSTIANO_TOKEN | ' + tokData2[i][0] + ' | expires=' + tokData2[i][3]);
    }
  }

  log.forEach(l => Logger.log(l));
  return log;
}


// — GERAR LINK DIRETO DE CONEXÃO STRAVA (sem formulário) ——————————————————
function gerarLinkConexaoStrava() {
  const ui    = SpreadsheetApp.getUi();
  const r     = ui.prompt(
    'Conectar Strava',
    'Digite o ID do atleta (ex: ATH_001):',
    ui.ButtonSet.OK_CANCEL
  );
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const athId = r.getResponseText().trim().toUpperCase();
  if (!athId) return;

  const props    = PropertiesService.getScriptProperties();
  const clientId = props.getProperty('STRAVA_CLIENT_ID');
  const redirect = props.getProperty('WEBAPP_URL');

  const oauthUrl = STRAVA_AUTH_URL
    + '?client_id='        + clientId
    + '&redirect_uri='     + encodeURIComponent(redirect)
    + '&response_type=code'
    + '&approval_prompt=auto'
    + '&scope='            + STRAVA_SCOPE
    + '&state='            + athId;

  ui.alert(
    'Link de Conexao Strava — ' + athId,
    'Envie este link para o atleta:\n\n' + oauthUrl,
    ui.ButtonSet.OK
  );
}
