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

    const agora = new Date();
    const vals  = [
      p.nome    || '',
      p.email   || '',
      p.whats   || '',
      p.nasc    || '',
      p.sexo    || '',
      p.peso    || '',
      p.altura  || '',
      p.mod     || p.modAgg || '',
      p.nivel   || '',
      p.obj     || p.objAgg || '',
      p.freq    || '',
      p.horario || '',
      p.saude   || p.saudeAgg || '',
      p.lesao   || '',
      p.plano   || '',
      p.cidade  || '',
      p.estado  || '',
      p.origem  || 'Link de Cadastro Web',
      agora,
      'Pendente',
      '',
      'Ativo',
      p.obs     || ''
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
    let   rowIdx = -1;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][H.CAD.ID - 1] || '').trim().toUpperCase() === athId) {
        rowIdx = i + 1; break;
      }
    }

    const agora = new Date().toLocaleString('pt-BR', { timeZone: 'America/Sao_Paulo' });
    const vals  = [
      p.nome    || '',
      p.email   || '',
      p.whats   || '',
      p.nasc    || '',
      p.sexo    || '',
      p.peso    || '',
      p.altura  || '',
      p.mod     || p.modAgg || '',
      p.nivel   || '',
      p.obj     || p.objAgg || '',
      p.freq    || '',
      p.horario || '',
      p.saude   || p.saudeAgg || '',
      p.lesao   || '',
      p.plano   || '',
      p.cidade  || '',
      p.estado  || '',
      p.origem  || 'Formulario Web',
      agora,
      'Pendente',
      '',
      'Ativo',
      p.obs     || ''
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

    return { ok: true, athId: athId, nome: p.nome || '' };

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

  const data      = sheet.getDataRange().getValues();
  const margemSeg = 600;
  const agora     = Math.floor(Date.now() / 1000);

  for (let i = 1; i < data.length; i++) {
    const rowAthId = String(data[i][H.TOK.ATH_ID - 1] || '').trim().toUpperCase();
    if (rowAthId !== athId.toUpperCase()) continue;

    const accessToken  = String(data[i][H.TOK.ACCESS   - 1] || '').trim();
    const refreshToken = String(data[i][H.TOK.REFRESH  - 1] || '').trim();
    const expiresAt    = Number(data[i][H.TOK.EXPIRES  - 1]) || 0;
    const rowIdx       = i + 1;

    if (!refreshToken) throw new Error('Atleta ' + athId + ' sem refresh_token. Reconecte o Strava.');

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

function _isAthIdValido_(athId) {
  const valor = String(athId || '').trim().toUpperCase();
  return /^ATH[A-Z0-9_ -]{3,40}$/.test(valor) &&
    valor !== 'ATH_ID' &&
    valor.indexOf('IDENTIFICA') < 0 &&
    valor.indexOf('NOME') < 0;
}

function _isRefreshTokenValido_(refreshToken) {
  const valor = String(refreshToken || '').trim();
  return valor.length >= 20 &&
    valor.indexOf(' ') < 0 &&
    valor.toLowerCase() !== 'refresh_token' &&
    valor.toLowerCase() !== 'undefined' &&
    valor.toLowerCase() !== 'null';
}

function _getTokenRow_(athId) {
  const id = String(athId || '').trim().toUpperCase();
  if (!_isAthIdValido_(id)) return null;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(H.SHEETS.TOKENS);
  if (!sheet) return null;

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const rowAthId = String(data[i][H.TOK.ATH_ID - 1] || '').trim().toUpperCase();
    if (rowAthId !== id) continue;
    return {
      row: i + 1,
      execId: data[i][H.TOK.EXEC_ID - 1] || '',
      athId: rowAthId,
      nome: data[i][H.TOK.NOME - 1] || '',
      accessToken: data[i][H.TOK.ACCESS - 1] || '',
      refreshToken: data[i][H.TOK.REFRESH - 1] || '',
      expiresAt: data[i][H.TOK.EXPIRES - 1] || '',
      scope: data[i][H.TOK.SCOPE - 1] || '',
      stravaId: data[i][H.TOK.STRAVA_ID - 1] || '',
      atualizadoEm: data[i][H.TOK.ULT_ATU - 1] || '',
      status: data[i][H.TOK.STATUS - 1] || ''
    };
  }
  return null;
}

function _logEvento_(athId, funcao, msg) {
  _log(athId || 'SYSTEM', 'INFO', funcao || '', msg || '', '');
}

function _logErro_(athId, funcao, erro) {
  const e = erro || {};
  _log(athId || 'SYSTEM', 'ERRO', funcao || '', e.message || String(erro || ''), e.stack || '');
}

function persistirCredenciaisStrava(athId, tokenData) {
  const id = String(athId || '').trim().toUpperCase();
  if (!_isAthIdValido_(id)) throw new Error('ATH_ID invalido para persistencia Strava.');
  if (!tokenData || !_isRefreshTokenValido_(tokenData.refresh_token)) {
    throw new Error('refresh_token ausente ou invalido na resposta Strava.');
  }
  if (!tokenData.access_token) throw new Error('access_token ausente na resposta Strava.');
  if (!tokenData.athlete || !tokenData.athlete.id) throw new Error('athlete.id ausente na resposta Strava.');

  _salvarTokensPlanilha(id, tokenData);
  _atualizarStatusCadastro(id, true, tokenData.athlete.id);
  _logEvento_(id, 'persistirCredenciaisStrava', 'Credenciais Strava persistidas sem expor segredos.');
  return {
    ok: true,
    athId: id,
    stravaId: String(tokenData.athlete.id),
    refreshTokenValido: true,
    accessTokenPresente: true
  };
}

function diagnosticoStravaHiperativoV3() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shCad = ss.getSheetByName(H.SHEETS.CADASTRO);
  const shTok = ss.getSheetByName(H.SHEETS.TOKENS);
  const props = PropertiesService.getScriptProperties();
  const resultado = {
    ok: true,
    timestamp: new Date(),
    credenciais: {
      clientId: !!props.getProperty('STRAVA_CLIENT_ID'),
      clientSecret: !!props.getProperty('STRAVA_CLIENT_SECRET'),
      webAppUrl: !!props.getProperty('WEBAPP_URL')
    },
    cadastro: { totalLinhas: shCad ? Math.max(shCad.getLastRow() - 1, 0) : 0, atletasValidos: 0 },
    tokens: { totalTokens: 0, refreshTokensValidos: 0, refreshTokensInvalidos: 0 },
    pendencias: []
  };

  if (!shCad) resultado.pendencias.push('Aba CADASTRO ausente.');
  if (!shTok) resultado.pendencias.push('Aba TOKENS ausente.');

  if (shCad) {
    const cad = shCad.getDataRange().getValues();
    for (let i = 1; i < cad.length; i++) {
      if (_isAthIdValido_(cad[i][H.CAD.ID - 1])) resultado.cadastro.atletasValidos++;
    }
  }

  if (shTok) {
    const tok = shTok.getDataRange().getValues();
    for (let i = 1; i < tok.length; i++) {
      const athId = tok[i][H.TOK.ATH_ID - 1];
      if (!_isAthIdValido_(athId)) continue;
      resultado.tokens.totalTokens++;
      if (_isRefreshTokenValido_(tok[i][H.TOK.REFRESH - 1])) resultado.tokens.refreshTokensValidos++;
      else resultado.tokens.refreshTokensInvalidos++;
    }
  }

  if (!resultado.tokens) resultado.tokens = { totalTokens: 0, refreshTokensValidos: 0, refreshTokensInvalidos: 0 };
  _logEvento_('SYSTEM', 'diagnosticoStravaHiperativoV3',
    'tokens=' + resultado.tokens.totalTokens +
    ' refresh_validos=' + resultado.tokens.refreshTokensValidos +
    ' atletas_validos=' + resultado.cadastro.atletasValidos);
  Logger.log(JSON.stringify(resultado));
  return resultado;
}

function testeValidadorAthId() {
  const casos = {
    ATH001: _isAthIdValido_('ATH001'),
    ATH_1781116630575: _isAthIdValido_('ATH_1781116630575'),
    cabecalhoIdentificacao: _isAthIdValido_('IDENTIFICACAO'),
    cabecalhoNome: _isAthIdValido_('Nome Completo'),
    vazio: _isAthIdValido_('')
  };
  if (!casos.ATH001 || !casos.ATH_1781116630575 || casos.cabecalhoIdentificacao || casos.cabecalhoNome || casos.vazio) {
    throw new Error('testeValidadorAthId falhou: ' + JSON.stringify(casos));
  }
  return casos;
}

function testeRefreshTokenVazio() {
  const casos = {
    vazio: _isRefreshTokenValido_(''),
    undefinedTexto: _isRefreshTokenValido_('undefined'),
    curto: _isRefreshTokenValido_('abc123'),
    valido: _isRefreshTokenValido_('refresh_token_exemplo_1234567890')
  };
  if (casos.vazio || casos.undefinedTexto || casos.curto || !casos.valido) {
    throw new Error('testeRefreshTokenVazio falhou: ' + JSON.stringify(casos));
  }
  return casos;
}

function testeTokenRowSemExporSegredos() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shTok = ss.getSheetByName(H.SHEETS.TOKENS);
  if (!shTok) return { ok: true, motivo: 'Aba TOKENS ausente.' };
  const data = shTok.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const athId = data[i][H.TOK.ATH_ID - 1];
    if (!_isAthIdValido_(athId)) continue;
    const row = _getTokenRow_(athId);
    return {
      ok: true,
      athId: row.athId,
      row: row.row,
      temAccessToken: !!row.accessToken,
      temRefreshToken: !!row.refreshToken,
      refreshTokenValido: _isRefreshTokenValido_(row.refreshToken),
      stravaId: row.stravaId || ''
    };
  }
  return { ok: true, motivo: 'Nenhuma linha de token valida encontrada.' };
}

function verificarStatusStravaAtletas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shCad = ss.getSheetByName(H.SHEETS.CADASTRO);
  if (!shCad) throw new Error('Aba CADASTRO nao encontrada.');

  const shTok = ss.getSheetByName(H.SHEETS.TOKENS);
  const tokensPorAthId = {};
  if (shTok) {
    const tok = shTok.getDataRange().getValues();
    for (let i = 1; i < tok.length; i++) {
      const athId = String(tok[i][H.TOK.ATH_ID - 1] || '').trim().toUpperCase();
      if (!_isAthIdValido_(athId)) continue;
      tokensPorAthId[athId] = {
        refreshValido: _isRefreshTokenValido_(tok[i][H.TOK.REFRESH - 1]),
        stravaId: tok[i][H.TOK.STRAVA_ID - 1] || '',
        status: tok[i][H.TOK.STATUS - 1] || '',
        atualizadoEm: tok[i][H.TOK.ULT_ATU - 1] || ''
      };
    }
  }

  const cad = shCad.getDataRange().getValues();
  const saida = [[
    'ATH_ID', 'Nome', 'Email', 'Status Cadastro', 'Strava OK',
    'Strava ID Cadastro', 'Token Encontrado', 'Refresh Token Valido',
    'Status Token', 'Ultima Atualizacao Token', 'Observacao'
  ]];
  let validos = 0;
  let ignorados = 0;

  for (let i = 1; i < cad.length; i++) {
    const athId = String(cad[i][H.CAD.ID - 1] || '').trim().toUpperCase();
    if (!_isAthIdValido_(athId)) {
      if (cad[i].join('').trim()) ignorados++;
      continue;
    }
    validos++;
    const token = tokensPorAthId[athId] || null;
    saida.push([
      athId,
      cad[i][H.CAD.NOME - 1] || '',
      cad[i][H.CAD.EMAIL - 1] || '',
      cad[i][H.CAD.STATUS - 1] || '',
      cad[i][H.CAD.STRAVA_OK - 1] || '',
      cad[i][H.CAD.STRAVA_ID - 1] || '',
      token ? 'Sim' : 'Nao',
      token && token.refreshValido ? 'Sim' : 'Nao',
      token ? token.status : '',
      token ? token.atualizadoEm : '',
      token ? '' : 'Sem token Strava persistido'
    ]);
  }

  let shStatus = ss.getSheetByName('📡 STRAVA STATUS');
  if (!shStatus) shStatus = ss.insertSheet('📡 STRAVA STATUS');
  shStatus.clearContents();
  shStatus.getRange(1, 1, saida.length, saida[0].length).setValues(saida);
  shStatus.getRange(1, 1, 1, saida[0].length).setFontWeight('bold');
  shStatus.autoResizeColumns(1, saida[0].length);

  const resultado = {
    validos: validos,
    ignorados: ignorados,
    gravados: saida.length - 1
  };
  _logEvento_('SYSTEM', 'verificarStatusStravaAtletas',
    'validos=' + resultado.validos + ' ignorados=' + resultado.ignorados + ' gravados=' + resultado.gravados);
  return resultado;
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
    const distKm = (a.distance || 0) / 1000;
    const tempoS = a.moving_time || 0;
    const velMps = tempoS > 0 ? (a.distance || 0) / tempoS : 0;
    const execId = 'ATIV_' + Utilities.getUuid().substring(0, 8).toUpperCase();
    sheet.appendRow([
      execId, athId, nomeAtleta,
      a.start_date_local ? new Date(a.start_date_local) : '',
      _traduzirTipo(a.sport_type || a.type || ''),
      'Strava', sid, a.name || '',
      tempoS, a.elapsed_time || 0, a.distance || 0, distKm,
      velMps, velMps * 3.6,
      _calcPace(tempoS, distKm), _calcPaceDecimal(tempoS, distKm),
      a.average_heartrate || '', a.max_heartrate || '',
      a.total_elevation_gain || '', a.calories || '',
      a.average_cadence || '', a.average_watts || '',
      a.map ? (a.map.summary_polyline || '') : '', new Date()
    ]);
    existentes.add(sid);
    count++;
  }
  _log(athId, 'INFO', '_gravarAtividades', count + ' novas de ' + atividades.length + ' recebidas', '');
  return count;
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
  _log('SISTEMA', 'INFO', 'triggerImportacaoAutomatica', 'Iniciando importação automática...', '');
  const resultado = _importarTodosAtletas();
  _log('SISTEMA', 'INFO', 'triggerImportacaoAutomatica', 'Concluído: ' + resultado, '');
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
