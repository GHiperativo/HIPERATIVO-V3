/**
 * ═══════════════════════════════════════════════════════════════════════
 * HIPERATIVO V3 — Supabase.gs
 * Integração com Supabase como banco de dados seguro e backup real-time.
 *
 * SETUP (uma vez só):
 *   1. Abra o Supabase → Settings → API → copie "service_role" key
 *   2. No Apps Script: ⚡ HIPERATIVO → Configurar credenciais
 *      Adicione: SUPABASE_URL = https://korlpbclqgmqvpbrungc.supabase.co
 *                SUPABASE_KEY = <sua service_role key>
 *
 * POR QUE SUPABASE?
 *   - Tokens Strava sobrevivem a qualquer wipe da planilha
 *   - Atletas nunca precisam reconectar após manutenção
 *   - Histórico de atividades preservado independente do Sheets
 *   - Base para futuro app/dashboard web
 * ═══════════════════════════════════════════════════════════════════════
 */

const SUPA_PROJECT = 'korlpbclqgmqvpbrungc';
const SUPA_BASE_URL = 'https://' + SUPA_PROJECT + '.supabase.co/rest/v1';

// ── Helper: headers autenticados ───────────────────────────────────────────────
function _supaHeaders() {
  const props = PropertiesService.getScriptProperties();
  const key   = props.getProperty('SUPABASE_KEY') || '';
  if (!key) throw new Error('SUPABASE_KEY nao configurada. Va em Configurar credenciais.');
  return {
    'apikey':        key,
    'Authorization': 'Bearer ' + key,
    'Content-Type':  'application/json',
    'Prefer':        'return=minimal',
  };
}

// ── Helper: requisição REST ────────────────────────────────────────────────────
function _supaRequest(method, path, payload) {
  const url  = SUPA_BASE_URL + path;
  const opts = {
    method:             method,
    headers:            _supaHeaders(),
    muteHttpExceptions: true,
  };
  if (payload) opts.payload = JSON.stringify(payload);

  const resp = UrlFetchApp.fetch(url, opts);
  const code = resp.getResponseCode();
  if (code >= 400) {
    throw new Error('Supabase ' + method.toUpperCase() + ' ' + path +
      ' → HTTP ' + code + ': ' + resp.getContentText().substring(0, 200));
  }
  const body = resp.getContentText();
  return body ? JSON.parse(body) : null;
}

// ── 1. SINCRONIZAR ATLETA ─────────────────────────────────────────────────────
/**
 * Upsert de um atleta no Supabase.
 * Chamado em: salvarCadastroAjax(), _processarFormCadastro(), restaurarAtletas()
 */
function supaSyncAtleta(athId, dados) {
  try {
    const payload = Object.assign({ ath_id: athId }, dados);
    _supaRequest('post', '/atletas?on_conflict=ath_id', payload);
    // Atualizar headers para upsert
    PropertiesService.getScriptProperties(); // force refresh
    UrlFetchApp.fetch(SUPA_BASE_URL + '/atletas?on_conflict=ath_id', {
      method:             'post',
      headers:            Object.assign(_supaHeaders(), { 'Prefer': 'resolution=merge-duplicates,return=minimal' }),
      payload:            JSON.stringify(payload),
      muteHttpExceptions: true,
    });
    _log(athId, 'INFO', 'supaSyncAtleta', 'Atleta sincronizado no Supabase.', '');
  } catch(e) {
    _log(athId, 'AVISO', 'supaSyncAtleta', 'Falha ao sincronizar: ' + e.message, '');
    // Não bloquear fluxo principal — Supabase é backup, não bloqueador
  }
}

// ── 2. SALVAR TOKEN NO SUPABASE ───────────────────────────────────────────────
/**
 * Backup seguro do refresh_token no Supabase.
 * Se a planilha TOKENS for wipada, este backup permite restauração automática.
 * Chamado em: _salvarTokensPlanilha()
 */
function supaSalvarToken(athId, tokenData) {
  try {
    const athlete = tokenData.athlete || {};
    const nome    = _getNomeAtleta(athId) ||
                    ((athlete.firstname || '') + ' ' + (athlete.lastname || '')).trim();

    const payload = {
      ath_id:        athId,
      nome:          nome,
      access_token:  tokenData.access_token  || '',
      refresh_token: tokenData.refresh_token || '',
      expires_at:    tokenData.expires_at    || 0,
      strava_id:     String(athlete.id || ''),
      ult_atu:       new Date().toISOString(),
      status:        'Ativo',
    };

    const resp = UrlFetchApp.fetch(SUPA_BASE_URL + '/tokens_strava?on_conflict=ath_id', {
      method:             'post',
      headers:            Object.assign(_supaHeaders(), { 'Prefer': 'resolution=merge-duplicates,return=minimal' }),
      payload:            JSON.stringify(payload),
      muteHttpExceptions: true,
    });

    if (resp.getResponseCode() >= 400) {
      throw new Error('HTTP ' + resp.getResponseCode() + ': ' + resp.getContentText().substring(0, 150));
    }
    _log(athId, 'INFO', 'supaSalvarToken', 'Token salvo no Supabase (backup seguro).', '');
  } catch(e) {
    _log(athId, 'AVISO', 'supaSalvarToken', 'Falha ao salvar token no Supabase: ' + e.message, '');
    // Não bloqueia — Supabase é camada adicional, PropertiesService é o backup primário
  }
}

// ── 3. RECUPERAR TOKEN DO SUPABASE ────────────────────────────────────────────
/**
 * Último recurso de recuperação de token.
 * Ordem de fallback em _getValidAccessToken():
 *   1. Planilha TOKENS (normal)
 *   2. PropertiesService RT_<ATH_ID> (backup local)
 *   3. Supabase tokens_strava (backup remoto) ← esta função
 */
function supaGetRefreshToken(athId) {
  try {
    const resp = UrlFetchApp.fetch(
      SUPA_BASE_URL + '/tokens_strava?ath_id=eq.' + encodeURIComponent(athId) +
      '&select=refresh_token,status&limit=1',
      {
        method:             'get',
        headers:            _supaHeaders(),
        muteHttpExceptions: true,
      }
    );
    if (resp.getResponseCode() !== 200) return null;
    const rows = JSON.parse(resp.getContentText());
    if (!rows || !rows.length || !rows[0].refresh_token) return null;
    _log(athId, 'INFO', 'supaGetRefreshToken', 'Refresh_token recuperado do Supabase.', '');
    return rows[0].refresh_token;
  } catch(e) {
    _log(athId, 'AVISO', 'supaGetRefreshToken', 'Falha ao buscar token no Supabase: ' + e.message, '');
    return null;
  }
}

// ── 4. RESTAURAR TODOS OS TOKENS DO SUPABASE ─────────────────────────────────
/**
 * Varre tokens_strava no Supabase e reconstrói a aba TOKENS da planilha.
 * Execute quando a aba TOKENS for wipada e PropertiesService não tiver backups.
 */
function supaRestaurarTodosTokens() {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const ui  = SpreadsheetApp.getUi();
  try {
    const rows = _supaRequest('get', '/tokens_strava?select=*&order=ath_id');
    if (!rows || !rows.length) {
      ui.alert('Nenhum token encontrado no Supabase.');
      return;
    }

    const shTok = ss.getSheetByName(H.SHEETS.TOKENS);
    if (!shTok) throw new Error('Aba TOKENS nao encontrada. Execute Setup primeiro.');

    // Identificar IDs já existentes
    const tokData = shTok.getDataRange().getValues();
    const existIds = new Set();
    for (let i = 1; i < tokData.length; i++) {
      const v = String(tokData[i][H.TOK.ATH_ID - 1] || '').trim().toUpperCase();
      if (v) existIds.add(v);
    }

    let restaurados = 0;
    for (const t of rows) {
      if (!t.refresh_token) continue;
      if (existIds.has((t.ath_id || '').toUpperCase())) continue;

      const execId = 'TOK_' + Utilities.getUuid().substring(0, 8).toUpperCase();
      shTok.appendRow([
        execId, t.ath_id, t.nome || '',
        '',               // access_token (vazio — será renovado automaticamente)
        t.refresh_token,  // refresh_token real — atleta NÃO precisa reconectar
        t.expires_at || '',
        'read,activity:read_all,profile:read_all',
        t.strava_id || '',
        new Date().toLocaleString('pt-BR'),
        'Restaurado do Supabase',
      ]);

      // Atualizar PropertiesService também
      PropertiesService.getScriptProperties().setProperty(
        'RT_' + (t.ath_id || '').toUpperCase(), t.refresh_token
      );
      restaurados++;
    }

    SpreadsheetApp.flush();
    _log('SYSTEM', 'INFO', 'supaRestaurarTodosTokens',
      restaurados + ' token(s) restaurados do Supabase.', '');
    ui.alert('✅ Tokens Restaurados',
      restaurados + ' atleta(s) restaurados do Supabase.\n' +
      'Nenhum precisou reconectar o Strava.',
      ui.ButtonSet.OK);
  } catch(e) {
    _log('SYSTEM', 'ERRO', 'supaRestaurarTodosTokens', e.message, e.stack || '');
    ui.alert('❌ Erro', e.message, ui.ButtonSet.OK);
  }
}

// ── 5. STATUS DO SUPABASE ─────────────────────────────────────────────────────
function supaStatus() {
  try {
    const atletas = _supaRequest('get', '/atletas?select=ath_id,nome,strava_ok&order=nome');
    const tokens  = _supaRequest('get', '/tokens_strava?select=ath_id,status&order=ath_id');
    const nAtl    = atletas ? atletas.length : 0;
    const nTok    = tokens  ? tokens.length  : 0;
    const comTok  = tokens  ? tokens.filter(t => t.status === 'Ativo').length : 0;
    SpreadsheetApp.getUi().alert(
      '✅ Supabase OK',
      'Projeto: hiperativo-v3 (sa-east-1)\n' +
      'Atletas cadastrados: ' + nAtl + '\n' +
      'Tokens salvos: ' + nTok + ' (' + comTok + ' ativos)\n\n' +
      'Backups funcionando corretamente.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch(e) {
    SpreadsheetApp.getUi().alert('❌ Supabase indisponível', e.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}
