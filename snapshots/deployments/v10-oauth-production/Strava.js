/**
 * HIPERATIVO V3 - Strava
 * Fluxo limpo: cadastro -> autorizacao Strava -> refresh_token -> atividades.
 */

const STRAVA_AUTH_URL = 'https://www.strava.com/oauth/authorize';
const STRAVA_TOKEN_URL = 'https://www.strava.com/oauth/token';
const STRAVA_API_BASE = 'https://www.strava.com/api/v3';
const STRAVA_SCOPE = 'read,activity:read_all';
const STRAVA_OAUTH_STATE_PREFIX = 'STRAVA_OAUTH_STATE_';
const STRAVA_OAUTH_STATE_TTL_MS = 24 * 60 * 60 * 1000;

function _registrarChamadaStrava(qtd) {
  qtd = Number(qtd) || 1;
  const props = PropertiesService.getScriptProperties();
  const agora = new Date();
  const hojeStr = Utilities.formatDate(agora, 'America/Sao_Paulo', 'yyyy-MM-dd');
  const janela = Math.floor(agora.getTime() / (15 * 60 * 1000));
  const keyDia = 'RATE_DIA_' + hojeStr;
  const keyJanela = 'RATE_15M_' + janela;
  props.setProperty(keyDia, String((parseInt(props.getProperty(keyDia) || '0', 10) || 0) + qtd));
  props.setProperty(keyJanela, String((parseInt(props.getProperty(keyJanela) || '0', 10) || 0) + qtd));
}

function gerarLinkStrava() {
  const ui = SpreadsheetApp.getUi();
  const selecionado = _getAtletaLinhaSelecionada();
  const r = ui.prompt(
    'Conectar atleta ao Strava',
    'Digite nome, e-mail ou ID do atleta:' +
      (selecionado ? '\n\nLinha selecionada: ' + selecionado.nome + '\nDeixe em branco para usar essa linha.' : ''),
    ui.ButtonSet.OK_CANCEL
  );
  if (r.getSelectedButton() !== ui.Button.OK) return;
  try {
    const atleta = _resolverAtleta(r.getResponseText(), selecionado);
    const url = _gerarUrlConexaoStrava(atleta.athId);
    ui.alert('Link Strava para ' + atleta.nome, url, ui.ButtonSet.OK);
    _log(atleta.athId, 'INFO', 'gerarLinkStrava', 'Link de conexao Strava gerado.', '');
  } catch (err) {
    _log('SYSTEM', 'ERRO', 'gerarLinkStrava', err.message, err.stack || '');
    ui.alert('Erro ao gerar link', err.message, ui.ButtonSet.OK);
  }
}

function _gerarUrlCadastro(athId) {
  const webAppUrl = PropertiesService.getScriptProperties().getProperty('WEBAPP_URL') || '';
  if (!webAppUrl) throw new Error('WEBAPP_URL nao configurado.');
  _validarUrlWebApp(webAppUrl);
  return webAppUrl + '?cadastro=true&athId=' + encodeURIComponent(athId || '');
}

function _gerarUrlConexaoStrava(athId) {
  const webAppUrl = PropertiesService.getScriptProperties().getProperty('WEBAPP_URL') || '';
  if (!webAppUrl) throw new Error('WEBAPP_URL nao configurado.');
  _validarUrlWebApp(webAppUrl);
  const atleta = _resolverAtleta(athId);
  return webAppUrl + '?connect=strava&athId=' + encodeURIComponent(atleta.athId);
}

function _gerarUrlOAuth(athId) {
  const props = PropertiesService.getScriptProperties();
  const clientId = props.getProperty('STRAVA_CLIENT_ID') || '';
  const webAppUrl = props.getProperty('WEBAPP_URL') || '';
  if (!clientId) throw new Error('STRAVA_CLIENT_ID nao configurado.');
  if (!webAppUrl) throw new Error('WEBAPP_URL nao configurado.');
  _validarUrlWebApp(webAppUrl);

  const atleta = _resolverAtleta(athId);
  const state = Utilities.getUuid().replace(/-/g, '') + String(Date.now());
  props.setProperty(STRAVA_OAUTH_STATE_PREFIX + state, JSON.stringify({
    athId: atleta.athId,
    nome: atleta.nome,
    email: atleta.email,
    criadoEm: Date.now(),
    expiraEm: Date.now() + STRAVA_OAUTH_STATE_TTL_MS
  }));
  _limparEstadosOAuthExpirados_();
  _log(atleta.athId, 'INFO', '_gerarUrlOAuth', 'OAuth Strava iniciado para ' + atleta.nome, '');

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

function gerarUrlOAuth(athId) {
  return _gerarUrlOAuth(athId);
}

function _obterEstadoOAuth_(state) {
  const props = PropertiesService.getScriptProperties();
  const key = STRAVA_OAUTH_STATE_PREFIX + String(state || '');
  const raw = props.getProperty(key);
  if (!raw) throw new Error('Autorizacao nao encontrada ou ja utilizada. Gere um novo link.');
  const dados = JSON.parse(raw);
  if (!dados.athId || !dados.expiraEm || Number(dados.expiraEm) < Date.now()) {
    props.deleteProperty(key);
    throw new Error('A autorizacao expirou. Gere um novo link de conexao.');
  }
  return dados;
}

function _consumirEstadoOAuth_(state) {
  PropertiesService.getScriptProperties().deleteProperty(STRAVA_OAUTH_STATE_PREFIX + String(state || ''));
}

function _limparEstadosOAuthExpirados_() {
  const props = PropertiesService.getScriptProperties();
  const all = props.getProperties();
  Object.keys(all).forEach(key => {
    if (key.indexOf(STRAVA_OAUTH_STATE_PREFIX) !== 0) return;
    try {
      const dados = JSON.parse(all[key] || '{}');
      if (!dados.expiraEm || Number(dados.expiraEm) < Date.now()) props.deleteProperty(key);
    } catch (_) {
      props.deleteProperty(key);
    }
  });
}

function _validarUrlWebApp(webAppUrl) {
  if (!/^https:\/\/script\.google\.com\/macros\/s\/[^/]+\/exec$/.test(String(webAppUrl || '').trim())) {
    throw new Error('WEBAPP_URL invalida. Use a URL publicada que termina em /exec.');
  }
}

function _normalizarBuscaAtleta(valor) {
  return String(valor || '').trim().toLocaleLowerCase('pt-BR');
}

function _listarAtletasCadastro() {
  const ss = SpreadsheetApp.openById(_getSsId());
  const sh = ss.getSheetByName(H.SHEETS.CADASTRO);
  if (!sh) throw new Error('Aba CADASTRO nao encontrada.');
  const data = sh.getDataRange().getValues();
  const atletas = [];
  for (let i = 2; i < data.length; i++) {
    const athId = String(data[i][H.CAD.ID - 1] || '').trim().toUpperCase();
    if (!athId) continue;
    atletas.push({
      athId,
      nome: String(data[i][H.CAD.NOME - 1] || '').trim() || athId,
      email: String(data[i][H.CAD.EMAIL - 1] || '').trim().toLowerCase(),
      row: i + 1
    });
  }
  return atletas;
}

function _resolverAtleta(referencia, fallback) {
  const busca = _normalizarBuscaAtleta(referencia);
  if (!busca && fallback && fallback.athId) return fallback;
  if (!busca) throw new Error('Informe nome, e-mail ou ID do atleta.');

  const atletas = _listarAtletasCadastro();
  const exatos = atletas.filter(a =>
    _normalizarBuscaAtleta(a.athId) === busca ||
    _normalizarBuscaAtleta(a.nome) === busca ||
    _normalizarBuscaAtleta(a.email) === busca
  );
  if (exatos.length === 1) return exatos[0];
  if (exatos.length > 1) throw new Error('Mais de um cadastro encontrado. Use o e-mail.');

  const porNome = atletas.filter(a => _normalizarBuscaAtleta(a.nome).indexOf(busca) >= 0);
  if (porNome.length === 1) return porNome[0];
  if (porNome.length > 1) throw new Error('Mais de um atleta corresponde a esse nome. Use e-mail ou ID.');
  throw new Error('Atleta nao encontrado no CADASTRO: ' + referencia);
}

function _getAtletaLinhaSelecionada() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getActiveSheet();
    if (!sh || sh.getName() !== H.SHEETS.CADASTRO || sh.getActiveRange().getRow() < 3) return null;
    return _resolverAtleta(sh.getRange(sh.getActiveRange().getRow(), H.CAD.ID).getDisplayValue());
  } catch (_) {
    return null;
  }
}

function _gerarAthIdUnico_() {
  const existentes = {};
  _listarAtletasCadastro().forEach(a => existentes[a.athId] = true);
  for (let i = 0; i < 10; i++) {
    const candidato = 'ATH' + Utilities.getUuid().replace(/-/g, '').substring(0, 8).toUpperCase();
    if (!existentes[candidato]) return candidato;
  }
  throw new Error('Nao foi possivel gerar ID unico para o atleta.');
}

function _limparTexto(v, maxLen) {
  return String(v || '').replace(/[\u0000-\u001F\u007F]/g, ' ').trim().substring(0, maxLen);
}

function _validarNumeroOpcional(v, min, max, nomeCampo) {
  if (v === undefined || v === null || String(v).trim() === '') return '';
  const n = Number(String(v).replace(',', '.'));
  if (!isFinite(n) || n < min || n > max) throw new Error(nomeCampo + ' fora do intervalo permitido.');
  return n;
}

function _validarCadastroPayload(p) {
  p = p || {};
  const nome = _limparTexto(p.nome, 120);
  const email = _limparTexto(p.email, 160).toLowerCase();
  const whats = _limparTexto(p.whats, 30);
  if (!nome) throw new Error('Nome obrigatorio.');
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) throw new Error('E-mail invalido.');
  if (whats.replace(/\D/g, '').length < 10) throw new Error('WhatsApp invalido.');
  return {
    athId: _limparTexto(p.athId, 40).toUpperCase(),
    nome,
    email,
    whats,
    nasc: _limparTexto(p.nasc, 20),
    sexo: _limparTexto(p.sexo, 30),
    peso: _validarNumeroOpcional(p.peso, 30, 300, 'Peso'),
    altura: _validarNumeroOpcional(p.altura, 100, 250, 'Altura'),
    mod: _limparTexto(p.mod || p.modAgg, 120),
    nivel: _limparTexto(p.nivel, 40),
    obj: _limparTexto(p.obj || p.objAgg, 200),
    freq: _limparTexto(p.freq, 20),
    horario: _limparTexto(p.horario, 40),
    saude: _limparTexto(p.saude || p.saudeAgg, 300),
    lesao: _limparTexto(p.lesao, 300),
    med: _limparTexto(p.med, 200),
    prova: _limparTexto(p.prova, 160),
    plano: _limparTexto(p.plano, 80),
    cidade: _limparTexto(p.cidade, 80),
    estado: _limparTexto(p.estado, 40),
    cpf: _limparTexto(p.cpf, 20),
    origem: _limparTexto(p.origem, 80),
    obs: _limparTexto(p.obs, 500)
  };
}

function _salvarCadastroAtleta_(p) {
  p = _validarCadastroPayload(p);
  let athId = (p.athId || '').trim().toUpperCase();
  if (!athId) athId = _gerarAthIdUnico_();

  const ss = SpreadsheetApp.openById(_getSsId());
  const sh = ss.getSheetByName(H.SHEETS.CADASTRO);
  if (!sh) throw new Error('Aba CADASTRO nao encontrada.');
  const data = sh.getDataRange().getValues();
  let rowIdx = -1;
  for (let i = 2; i < data.length; i++) {
    if (String(data[i][H.CAD.ID - 1] || '').trim().toUpperCase() === athId) {
      rowIdx = i + 1;
      break;
    }
  }

  const vals = [
    p.nome, p.email, p.whats, p.nasc, p.sexo, p.peso, p.altura, p.mod, p.nivel,
    p.obj, p.freq, p.horario, p.saude, p.lesao, p.med, p.prova, p.plano,
    p.cidade, p.estado, p.cpf, p.origem || 'Formulario Web', new Date(),
    'Pendente', '', 'Ativo', p.obs
  ];

  if (rowIdx === -1) sh.appendRow([athId].concat(vals));
  else sh.getRange(rowIdx, 2, 1, vals.length).setValues([vals]);
  _log(athId, 'INFO', '_salvarCadastroAtleta_', 'Cadastro salvo: ' + p.nome, '');
  return { athId, nome: p.nome, email: p.email };
}

function _processarFormCadastro(p) {
  try {
    const atleta = _salvarCadastroAtleta_(p);
    const oauthUrl = _gerarUrlOAuth(atleta.athId);
    return HtmlService.createHtmlOutput(_paginaRedirecionando(oauthUrl, atleta.nome))
      .setTitle('Autorizar Strava');
  } catch (err) {
    _log((p && p.athId) || 'SYSTEM', 'ERRO', '_processarFormCadastro', err.message, err.stack || '');
    return HtmlService.createHtmlOutput(_paginaErro('Erro ao salvar cadastro', err.message));
  }
}

function salvarCadastroAjax(p) {
  try {
    const atleta = _salvarCadastroAtleta_(p);
    const oauthUrl = _gerarUrlOAuth(atleta.athId);
    _enviarEmailBoasVindasStrava_(atleta, oauthUrl);
    return { ok: true, athId: atleta.athId, nome: atleta.nome, oauthUrl, oauthErro: '' };
  } catch (err) {
    _log((p && p.athId) || 'SYSTEM', 'ERRO', 'salvarCadastroAjax', err.message, err.stack || '');
    return { ok: false, erro: err.message };
  }
}

function _enviarEmailBoasVindasStrava_(atleta, linkStrava) {
  if (!atleta.email || !atleta.email.includes('@')) return;
  try {
    const adminEmail = PropertiesService.getScriptProperties().getProperty('ADMIN_EMAIL') || 'contato@ghiperativo.com.br';
    const primeiroNome = atleta.nome.split(' ')[0] || 'Atleta';
    MailApp.sendEmail({
      to: atleta.email,
      replyTo: adminEmail,
      subject: primeiroNome + ', conecte seu Strava ao HIPERATIVO',
      body: 'Ola, ' + primeiroNome + '!\n\nSeu cadastro foi recebido. Conecte seu Strava neste link:\n' + linkStrava,
      htmlBody: '<p>Ola, ' + _htmlEscape(primeiroNome) + '!</p><p>Seu cadastro foi recebido.</p>' +
        '<p><a href="' + _htmlEscape(linkStrava) + '">Conectar meu Strava</a></p>'
    });
    _log(atleta.athId, 'INFO', '_enviarEmailBoasVindasStrava_', 'E-mail enviado para ' + atleta.email, '');
  } catch (err) {
    _log(atleta.athId, 'AVISO', '_enviarEmailBoasVindasStrava_', err.message, '');
  }
}

function _trocarCodigoPorToken(athId, code) {
  const props = PropertiesService.getScriptProperties();
  const clientId = props.getProperty('STRAVA_CLIENT_ID') || '';
  const clientSec = props.getProperty('STRAVA_CLIENT_SECRET') || '';
  if (!clientId || !clientSec) throw new Error('Credenciais Strava nao configuradas.');
  const resp = UrlFetchApp.fetch(STRAVA_TOKEN_URL, {
    method: 'post',
    payload: {
      client_id: clientId,
      client_secret: clientSec,
      code,
      grant_type: 'authorization_code'
    },
    muteHttpExceptions: true
  });
  _registrarChamadaStrava();
  if (resp.getResponseCode() !== 200) {
    throw new Error('Strava retornou erro ' + resp.getResponseCode() + ': ' + resp.getContentText().substring(0, 300));
  }
  return JSON.parse(resp.getContentText());
}

function _getValidAccessToken(athId) {
  const ss = SpreadsheetApp.openById(_getSsId());
  const sh = ss.getSheetByName(H.SHEETS.TOKENS);
  if (!sh) throw new Error('Aba TOKENS nao encontrada.');
  const data = sh.getDataRange().getValues();
  const agora = Math.floor(Date.now() / 1000);
  const margem = 600;
  for (let i = 2; i < data.length; i++) {
    const rowAthId = String(data[i][H.TOK.ATH_ID - 1] || '').trim().toUpperCase();
    if (rowAthId !== String(athId || '').toUpperCase()) continue;
    const access = String(data[i][H.TOK.ACCESS - 1] || '').trim();
    const refresh = String(data[i][H.TOK.REFRESH - 1] || '').trim();
    const expires = Number(data[i][H.TOK.EXPIRES - 1]) || 0;
    if (!refresh) throw new Error('Atleta sem refresh_token. Reconecte o Strava.');
    if (access && expires > agora + margem) return access;
    const novo = _refreshAccessToken(athId, refresh);
    sh.getRange(i + 1, H.TOK.ACCESS).setValue(novo.access_token || '');
    sh.getRange(i + 1, H.TOK.REFRESH).setValue(novo.refresh_token || refresh);
    sh.getRange(i + 1, H.TOK.EXPIRES).setValue(novo.expires_at || '');
    sh.getRange(i + 1, H.TOK.ULT_ATU).setValue(new Date());
    sh.getRange(i + 1, H.TOK.STATUS).setValue('Renovado');
    return novo.access_token;
  }
  throw new Error('Atleta nao encontrado na aba TOKENS.');
}

function _refreshAccessToken(athId, refreshToken) {
  const props = PropertiesService.getScriptProperties();
  const clientId = props.getProperty('STRAVA_CLIENT_ID') || '';
  const clientSec = props.getProperty('STRAVA_CLIENT_SECRET') || '';
  if (!clientId || !clientSec) throw new Error('Credenciais Strava nao configuradas.');
  const resp = UrlFetchApp.fetch(STRAVA_TOKEN_URL, {
    method: 'post',
    payload: {
      client_id: clientId,
      client_secret: clientSec,
      refresh_token: refreshToken,
      grant_type: 'refresh_token'
    },
    muteHttpExceptions: true
  });
  _registrarChamadaStrava();
  if (resp.getResponseCode() !== 200) {
    throw new Error('Falha ao renovar token: ' + resp.getContentText().substring(0, 300));
  }
  return JSON.parse(resp.getContentText());
}

function _salvarTokensPlanilha(athId, tokenData) {
  const ss = SpreadsheetApp.openById(_getSsId());
  const sh = ss.getSheetByName(H.SHEETS.TOKENS);
  if (!sh) throw new Error('Aba TOKENS nao encontrada.');
  const data = sh.getDataRange().getValues();
  const athlete = tokenData.athlete || {};
  const stravaId = athlete.id || '';
  if (!tokenData.access_token || !tokenData.refresh_token || !stravaId) {
    throw new Error('Resposta OAuth incompleta: access_token, refresh_token ou ID Strava ausente.');
  }
  const nome = _getNomeAtleta(athId) || ((athlete.firstname || '') + ' ' + (athlete.lastname || '')).trim();
  const agora = new Date();

  for (let i = 2; i < data.length; i++) {
    const outroAthId = String(data[i][H.TOK.ATH_ID - 1] || '').trim().toUpperCase();
    const outroStravaId = String(data[i][H.TOK.STRAVA_ID - 1] || '').trim();
    if (String(stravaId) === outroStravaId && outroAthId && outroAthId !== String(athId).toUpperCase()) {
      throw new Error('Esta conta Strava ja esta associada a outro atleta.');
    }
  }

  for (let i = 2; i < data.length; i++) {
    if (String(data[i][H.TOK.ATH_ID - 1] || '').trim().toUpperCase() === String(athId).toUpperCase()) {
      sh.getRange(i + 1, H.TOK.NOME).setValue(nome);
      sh.getRange(i + 1, H.TOK.ACCESS).setValue(tokenData.access_token);
      sh.getRange(i + 1, H.TOK.REFRESH).setValue(tokenData.refresh_token);
      sh.getRange(i + 1, H.TOK.EXPIRES).setValue(tokenData.expires_at || '');
      sh.getRange(i + 1, H.TOK.SCOPE).setValue(tokenData.scope || STRAVA_SCOPE);
      sh.getRange(i + 1, H.TOK.STRAVA_ID).setValue(stravaId);
      sh.getRange(i + 1, H.TOK.ULT_ATU).setValue(agora);
      sh.getRange(i + 1, H.TOK.STATUS).setValue('Ativo');
      return;
    }
  }

  sh.appendRow([
    'TOK_' + Utilities.getUuid().substring(0, 8).toUpperCase(),
    athId, nome, tokenData.access_token, tokenData.refresh_token, tokenData.expires_at || '',
    tokenData.scope || STRAVA_SCOPE, stravaId, agora, 'Ativo'
  ]);
}

function importarAtividades() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt('Importar atividades', 'Nome, e-mail ou ID do atleta. Vazio = todos:', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  try {
    const ref = String(r.getResponseText() || '').trim();
    const msg = ref ? (_importarAtividadesAtleta(_resolverAtleta(ref).athId, 3) + ' atividade(s) importada(s).') : _importarTodosAtletas();
    ui.alert('Resultado', msg, ui.ButtonSet.OK);
  } catch (err) {
    ui.alert('Erro', err.message, ui.ButtonSet.OK);
  }
}

function importarAtividadesTodosStrava() {
  const ui = SpreadsheetApp.getUi();
  try {
    ui.alert('Importacao Strava concluida', _importarTodosAtletas(), ui.ButtonSet.OK);
  } catch (err) {
    _log('TODOS', 'ERRO', 'importarAtividadesTodosStrava', err.message, err.stack || '');
    ui.alert('Erro ao importar', err.message, ui.ButtonSet.OK);
  }
}

function _importarTodosAtletas(paginas) {
  const ss = SpreadsheetApp.openById(_getSsId());
  const sh = ss.getSheetByName(H.SHEETS.TOKENS);
  if (!sh) return 'Aba TOKENS nao encontrada.';
  const data = sh.getDataRange().getValues();
  const msgs = [];
  let elegiveis = 0;
  for (let i = 2; i < data.length; i++) {
    const athId = String(data[i][H.TOK.ATH_ID - 1] || '').trim().toUpperCase();
    const refresh = String(data[i][H.TOK.REFRESH - 1] || '').trim();
    const status = String(data[i][H.TOK.STATUS - 1] || '').trim().toLowerCase();
    if (!athId || !refresh || status === 'revogado' || status === 'inativo') continue;
    elegiveis++;
    try {
      msgs.push(_getNomeAtleta(athId) + ': ' + _importarAtividadesAtleta(athId, paginas || 3) + ' nova(s)');
    } catch (err) {
      msgs.push(_getNomeAtleta(athId) + ': ERRO - ' + err.message);
    }
  }
  if (!elegiveis) return 'Nenhum atleta com refresh_token encontrado.';
  return msgs.join('\n') || 'Nenhuma atividade nova encontrada.';
}

function _importarAtividadesAtleta(athId, paginas) {
  const accessToken = _getValidAccessToken(athId);
  const nomeAtleta = _getNomeAtleta(athId);
  paginas = paginas || 3;
  let todas = [];
  for (let pg = 1; pg <= paginas; pg++) {
    const resp = UrlFetchApp.fetch(STRAVA_API_BASE + '/athlete/activities?per_page=50&page=' + pg, {
      headers: { Authorization: 'Bearer ' + accessToken },
      muteHttpExceptions: true
    });
    _registrarChamadaStrava();
    if (resp.getResponseCode() !== 200) throw new Error('Strava atividades HTTP ' + resp.getResponseCode());
    const page = JSON.parse(resp.getContentText() || '[]');
    if (!page.length) break;
    todas = todas.concat(page);
  }
  return _gravarAtividades(athId, nomeAtleta, todas);
}

function _gravarAtividades(athId, nomeAtleta, atividades) {
  if (!atividades || !atividades.length) return 0;
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss = SpreadsheetApp.openById(_getSsId());
    const sh = ss.getSheetByName(H.SHEETS.ATIVIDADES);
    if (!sh) throw new Error('Aba ATIVIDADES nao encontrada.');
    const existentes = new Set();
    const data = sh.getDataRange().getValues();
    for (let i = 2; i < data.length; i++) {
      const sid = String(data[i][H.ATIV.STRAVA_ID - 1] || '').trim();
      if (sid) existentes.add(sid);
    }
    let count = 0;
    atividades.forEach(a => {
      const sid = String(a.id || '');
      if (!sid || existentes.has(sid)) return;
      sh.appendRow(_linhaAtividadeStrava_(athId, nomeAtleta, a, ''));
      existentes.add(sid);
      count++;
    });
    _log(athId, 'INFO', '_gravarAtividades', count + ' nova(s) de ' + atividades.length + ' recebida(s).', '');
    return count;
  } finally {
    lock.releaseLock();
  }
}

function importarPerfilAtleta(athId) {
  if (!athId) {
    const r = SpreadsheetApp.getUi().prompt('Perfil Strava', 'Nome, e-mail ou ID do atleta:', SpreadsheetApp.getUi().ButtonSet.OK_CANCEL);
    if (r.getSelectedButton() !== SpreadsheetApp.getUi().Button.OK) return;
    athId = _resolverAtleta(r.getResponseText()).athId;
  }
  const token = _getValidAccessToken(athId);
  const resp = UrlFetchApp.fetch(STRAVA_API_BASE + '/athlete', {
    headers: { Authorization: 'Bearer ' + token },
    muteHttpExceptions: true
  });
  _registrarChamadaStrava();
  if (resp.getResponseCode() !== 200) throw new Error('Strava perfil HTTP ' + resp.getResponseCode());
  const perfil = JSON.parse(resp.getContentText());
  _atualizarPerfilNoCadastro(athId, perfil);
  return perfil;
}

function _atualizarPerfilNoCadastro(athId, perfil) {
  const sh = SpreadsheetApp.openById(_getSsId()).getSheetByName(H.SHEETS.CADASTRO);
  if (!sh) return;
  const data = sh.getDataRange().getValues();
  for (let i = 2; i < data.length; i++) {
    if (String(data[i][H.CAD.ID - 1] || '').trim().toUpperCase() === String(athId).toUpperCase()) {
      if (perfil.weight) sh.getRange(i + 1, H.CAD.PESO).setValue(perfil.weight);
      sh.getRange(i + 1, H.CAD.STRAVA_OK).setValue('Sim');
      if (perfil.id) sh.getRange(i + 1, H.CAD.STRAVA_ID).setValue(perfil.id);
      return;
    }
  }
}

function triggerImportacaoAutomatica() {
  _log('SISTEMA', 'INFO', 'triggerImportacaoAutomatica', 'Iniciando importacao automatica.', '');
  const resultado = _importarTodosAtletas();
  _log('SISTEMA', 'INFO', 'triggerImportacaoAutomatica', resultado, '');
}

function _atualizarStatusCadastro(athId, conectado, stravaId) {
  const sh = SpreadsheetApp.openById(_getSsId()).getSheetByName(H.SHEETS.CADASTRO);
  if (!sh) return;
  const data = sh.getDataRange().getValues();
  for (let i = 2; i < data.length; i++) {
    if (String(data[i][H.CAD.ID - 1] || '').trim().toUpperCase() === String(athId).toUpperCase()) {
      sh.getRange(i + 1, H.CAD.STRAVA_OK).setValue(conectado ? 'Sim' : 'Nao');
      if (stravaId) sh.getRange(i + 1, H.CAD.STRAVA_ID).setValue(stravaId);
      return;
    }
  }
}

function _getNomeAtleta(athId) {
  try {
    const sh = SpreadsheetApp.openById(_getSsId()).getSheetByName(H.SHEETS.CADASTRO);
    if (!sh) return athId;
    const data = sh.getDataRange().getValues();
    for (let i = 2; i < data.length; i++) {
      if (String(data[i][H.CAD.ID - 1] || '').trim().toUpperCase() === String(athId).toUpperCase()) {
        return String(data[i][H.CAD.NOME - 1] || athId).trim();
      }
    }
  } catch (_) {}
  return athId;
}

function sincronizarNomesAtividades() {
  const ui = SpreadsheetApp.getUi();
  try {
    const ss = SpreadsheetApp.openById(_getSsId());
    const sh = ss.getSheetByName(H.SHEETS.ATIVIDADES);
    if (!sh) throw new Error('Aba ATIVIDADES nao encontrada.');
    const lastRow = sh.getLastRow();
    if (lastRow < 3) {
      ui.alert('Nao ha atividades para sincronizar.');
      return;
    }
    const nomesPorId = {};
    _listarAtletasCadastro().forEach(a => nomesPorId[a.athId] = a.nome);
    const ids = sh.getRange(3, H.ATIV.ATH_ID, lastRow - 2, 1).getValues();
    const nomes = ids.map(r => [nomesPorId[String(r[0] || '').trim().toUpperCase()] || '']);
    sh.getRange(3, H.ATIV.NOME, nomes.length, 1).setValues(nomes);
    ui.alert('Nomes sincronizados', nomes.length + ' linha(s) revisada(s).', ui.ButtonSet.OK);
  } catch (err) {
    ui.alert('Erro', err.message, ui.ButtonSet.OK);
  }
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

function _calcPace(tempoSeg, distKm) {
  if (!distKm || distKm < 0.01) return '';
  const total = tempoSeg / distKm;
  return Math.floor(total / 60) + ':' + String(Math.round(total % 60)).padStart(2, '0') + ' /km';
}

function _calcPaceDecimal(tempoSeg, distKm) {
  if (!distKm || distKm < 0.01) return 0;
  return Math.round(tempoSeg / distKm);
}

function _traduzirTipo(tipo) {
  const map = {
    Run: 'Corrida',
    Ride: 'Ciclismo',
    Swim: 'Natacao',
    Walk: 'Caminhada',
    TrailRun: 'Trail Run',
    VirtualRide: 'Ciclismo Virtual',
    VirtualRun: 'Corrida Virtual',
    WeightTraining: 'Musculacao',
    Workout: 'Treino',
    Hike: 'Trilha'
  };
  return map[tipo] || tipo || 'Outro';
}

function _paginaCadastro(athId) {
  const id = _htmlEscape(athId || '');
  return '<!DOCTYPE html><html><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">' +
    '<title>Cadastro HIPERATIVO</title></head><body style="font-family:Arial;padding:24px;max-width:720px;margin:auto">' +
    '<h2>Cadastro HIPERATIVO</h2>' +
    '<form method="get">' +
    '<input type="hidden" name="salvar" value="true"><input type="hidden" name="athId" value="' + id + '">' +
    '<p><label>Nome completo<br><input name="nome" required style="width:100%"></label></p>' +
    '<p><label>E-mail<br><input name="email" type="email" required style="width:100%"></label></p>' +
    '<p><label>WhatsApp<br><input name="whats" required style="width:100%"></label></p>' +
    '<p><label>Modalidade<br><input name="mod" style="width:100%" value="Corrida"></label></p>' +
    '<p><label>Objetivo<br><input name="obj" style="width:100%"></label></p>' +
    '<button type="submit">Salvar e conectar Strava</button></form></body></html>';
}

function _paginaRedirecionando(oauthUrl, nome) {
  const safeUrl = _htmlEscape(oauthUrl || '');
  const primeiroNome = _htmlEscape((nome || '').split(' ')[0] || 'Atleta');
  return '<!DOCTYPE html><html lang="pt-BR"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">' +
    '<title>Autorizar Strava</title><style>body{font-family:Arial,sans-serif;background:#061426;color:#fff;min-height:100vh;display:flex;align-items:center;justify-content:center}.card{max-width:460px;padding:36px;border:1px solid #26384f;border-radius:18px;text-align:center}.btn{display:inline-block;margin-top:18px;padding:14px 24px;background:#fc4c02;color:#fff;text-decoration:none;border-radius:10px;font-weight:700}.copy{font-size:12px;color:#9ed1ff;word-break:break-all}</style></head><body>' +
    '<div class="card"><h2>Cadastro salvo, ' + primeiroNome + '!</h2>' +
    '<p>Para concluir, abra o Strava em uma nova aba e autorize o acesso.</p>' +
    '<a class="btn" href="' + safeUrl + '" target="_blank" rel="noopener noreferrer">Abrir Strava em nova aba</a>' +
    '<p class="copy">Se a nova aba nao abrir, copie este link:<br>' + safeUrl + '</p></div></body></html>';
}

function _paginaErro(titulo, mensagem) {
  return '<!DOCTYPE html><html><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">' +
    '<title>Erro</title></head><body style="font-family:Arial;padding:32px;background:#2a0808;color:#fff">' +
    '<h2>' + _htmlEscape(titulo) + '</h2><p>' + _htmlEscape(mensagem) + '</p></body></html>';
}

function _htmlEscape(v) {
  return String(v || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}
