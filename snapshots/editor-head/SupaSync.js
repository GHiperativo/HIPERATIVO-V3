// ═══════════════════════════════════════════════════════════════════════════
// SupaSync.gs — Integração Supabase como backup de tokens e sync de dados
// HIPERATIVO V3 | Criado: 20/06/2026
//
// ARQUITETURA:
//   Sheet TOKENS = primário (sem mudança)
//   Supabase tokens_strava = backup + fallback + foundation do app FORMA
//
// SETUP INICIAL (fazer UMA VEZ):
//   1. Abrir GAS Editor > Projeto > Configurações (ícone de engrenagem)
//   2. Propriedades do script > Adicionar:
//        SUPABASE_URL  = https://korlpbclqgmqvpbrungc.supabase.co
//        SUPABASE_KEY  = [colar o service_role key do Supabase Settings > API Keys]
//                        ↑ NÃO o anon key — o service_role (aba "Legacy anon, service_role API keys")
//   3. Executar migrarTokensParaSupabase() UMA vez (Menu GAS > Executar)
//   4. Pronto — sync automático ativo a partir daí
// ═══════════════════════════════════════════════════════════════════════════

// ─── HELPERS INTERNOS ────────────────────────────────────────────────────────

function _supaUrl_() {
  return PropertiesService.getScriptProperties().getProperty('SUPABASE_URL')
         || 'https://korlpbclqgmqvpbrungc.supabase.co';
}

function _supaKey_() {
  return PropertiesService.getScriptProperties().getProperty('SUPABASE_KEY') || '';
}

function _supaHeaders_() {
  const key = _supaKey_();
  return {
    'apikey': key,
    'Authorization': 'Bearer ' + key,
    'Content-Type': 'application/json'
  };
}

function _supaConfigurado_() { return !!_supaKey_(); }

// ─── UPSERT TOKEN NO SUPABASE ────────────────────────────────────────────────

function supaUpsertToken(athId, nome, accessToken, refreshToken, expiresAt) {
  if (!_supaConfigurado_()) return;
  try {
    const payload = JSON.stringify({
      ath_id:        athId,
      nome:          nome || '',
      access_token:  accessToken || '',
      refresh_token: refreshToken || '',
      expires_at:    Number(expiresAt) || 0,
      ult_atu:       new Date().toISOString(),
      status:        'Renovado'
    });
    const resp = UrlFetchApp.fetch(_supaUrl_() + '/rest/v1/tokens_strava', {
      method: 'post',
      headers: Object.assign({}, _supaHeaders_(), { 'Prefer': 'resolution=merge-duplicates' }),
      payload: payload,
      muteHttpExceptions: true
    });
    if (resp.getResponseCode() >= 300) {
      _log(athId, 'AVISO', 'supaUpsertToken', 'Sync Supabase retornou ' + resp.getResponseCode(), '');
    }
  } catch (e) {
    _log(athId, 'AVISO', 'supaUpsertToken', 'Sync Supabase falhou (nao critico): ' + e.message, '');
  }
}


function supaGetRefresh(athId) {
  if (!_supaConfigurado_()) return null;
  try {
    const url = _supaUrl_()
      + '/rest/v1/tokens_strava'
      + '?ath_id=eq.' + encodeURIComponent(athId)
      + '&select=refresh_token,access_token,expires_at&limit=1';
    const resp = UrlFetchApp.fetch(url, { method: 'get', headers: _supaHeaders_(), muteHttpExceptions: true });
    if (resp.getResponseCode() !== 200) return null;
    const rows = JSON.parse(resp.getContentText());
    if (!rows || rows.length === 0) return null;
    return {
      refresh_token: rows[0].refresh_token || '',
      access_token:  rows[0].access_token  || '',
      expires_at:    Number(rows[0].expires_at) || 0
    };
  } catch (e) { return null; }
}

// ─── DIAGNÓSTICO (rodar manualmente para verificar) ──────────────────────────

function diagnosticoSupabase() {
  if (!_supaConfigurado_()) { Logger.log('ERRO: SUPABASE_KEY nao configurada.'); return; }
  try {
    const url = _supaUrl_() + '/rest/v1/tokens_strava?select=ath_id,nome,status';
    const resp = UrlFetchApp.fetch(url, { method: 'get', headers: _supaHeaders_(), muteHttpExceptions: true });
    const code = resp.getResponseCode();
    if (code !== 200) { Logger.log('ERRO: Supabase respondeu ' + code + ': ' + resp.getContentText().substring(0, 200)); return; }
    const rows = JSON.parse(resp.getContentText());
    Logger.log('OK: Supabase conectado. Tokens: ' + rows.length);
    rows.forEach(r => Logger.log('  ' + r.ath_id + ' | ' + r.nome + ' | ' + r.status));
  } catch (e) { Logger.log('ERRO de conexao: ' + e.message); }
}

// ─── ATUALIZAR STATUS DO ATLETA NO SUPABASE ──────────────────────────────────

function supaAtualizarStravaOk(athId, status) {
  if (!_supaConfigurado_()) return;
  try {
    UrlFetchApp.fetch(_supaUrl_() + '/rest/v1/atletas?ath_id=eq.' + encodeURIComponent(athId), {
      method: 'patch',
      headers: Object.assign({}, _supaHeaders_(), { 'Prefer': 'return=minimal' }),
      payload: JSON.stringify({ strava_ok: status, updated_at: new Date().toISOString() }),
      muteHttpExceptions: true
    });
  } catch (e) { }
}

// ─── MONITORAMENTO DIÁRIO DE CONEXÃO STRAVA ──────────────────────────────────
// Instalado via instalarAcionadorMonitor() — roda 1x por dia às 8h.
// Envia email se qualquer atleta estiver com Strava quebrado.

function monitorarStravaOk() {
  if (!_supaConfigurado_()) return;
  try {
    // Buscar atletas com problema
    const url = _supaUrl_() + '/rest/v1/atletas'
      + '?strava_ok=in.(Reconectar,Erro)'
      + '&select=ath_id,nome,strava_ok,updated_at'
      + '&order=nome.asc';
    const resp = UrlFetchApp.fetch(url, { method: 'get', headers: _supaHeaders_(), muteHttpExceptions: true });
    if (resp.getResponseCode() !== 200) return;

    const problemas = JSON.parse(resp.getContentText());
    if (!problemas || problemas.length === 0) return; // tudo ok, sem email

    // Montar email de alerta
    const adminEmail = PropertiesService.getScriptProperties().getProperty('ADMIN_EMAIL')
                       || 'contato@ghiperativo.com.br';

    const linhas = problemas.map(a =>
      '• ' + a.nome + ' (' + a.ath_id + ') — Status: ' + a.strava_ok
    ).join('\n');

    const corpo = [
      '⚠️ ALERTA HIPERATIVO V3 — Strava desconectado',
      '',
      problemas.length + ' atleta(s) com conexão Strava quebrada:',
      '',
      linhas,
      '',
      'Esses atletas não estão recebendo dados do Strava.',
      'Acesse o sistema e peça para reconectarem via link OAuth.',
      '',
      '— Monitor automático HIPERATIVO V3',
      'Verificado em: ' + new Date().toLocaleString('pt-BR', { timeZone: 'America/Sao_Paulo' })
    ].join('\n');

    MailApp.sendEmail({
      to: adminEmail,
      subject: '⚠️ [Hiperativo] ' + problemas.length + ' atleta(s) precisam reconectar o Strava',
      body: corpo
    });

    Logger.log('Monitor: email enviado para ' + adminEmail + ' | ' + problemas.length + ' problema(s)');
  } catch (e) {
    Logger.log('Monitor: erro — ' + e.message);
  }
}

function instalarAcionadorMonitor() {
  // Remove acionadores antigos do monitor para evitar duplicatas
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'monitorarStravaOk')
    .forEach(t => ScriptApp.deleteTrigger(t));

  // Instala novo acionador diário às 8h (horário do script — Brasília)
  ScriptApp.newTrigger('monitorarStravaOk')
    .timeBased()
    .everyDays(1)
    .atHour(8)
    .create();

  Logger.log('Acionador diário instalado: monitorarStravaOk roda todo dia às 8h');
}
// ─── RENOVAÇÃO PROATIVA DE TOKENS (CORE DA RESILIÊNCIA) ──────────────────────
// Roda a cada 4h via trigger — SEM precisar abrir a planilha.
// Fluxo: lê tokens do Supabase → renova via API Strava → salva em ScriptProps + Supabase.
// Garante que nenhum token expire passivamente, mesmo sem uso do sistema.

function renovacaoProativaTokens() {
  if (!_supaConfigurado_()) { Logger.log('Supa não configurado.'); return; }
  const props = PropertiesService.getScriptProperties();
  const clientId  = props.getProperty('STRAVA_CLIENT_ID');
  const clientSecret = props.getProperty('STRAVA_CLIENT_SECRET');
  if (!clientId || !clientSecret) {
    Logger.log('ERRO: STRAVA_CLIENT_ID ou STRAVA_CLIENT_SECRET ausentes. Renovação abortada.');
    return;
  }

  try {
    const url = _supaUrl_() + '/rest/v1/tokens_strava?select=ath_id,nome,refresh_token,expires_at&order=ath_id.asc';
    const resp = UrlFetchApp.fetch(url, { method: 'get', headers: _supaHeaders_(), muteHttpExceptions: true });
    if (resp.getResponseCode() !== 200) {
      Logger.log('ERRO Supabase: ' + resp.getResponseCode()); return;
    }
    const tokens = JSON.parse(resp.getContentText());
    if (!tokens || tokens.length === 0) {
      Logger.log('Supabase vazio — aguardando primeira reconexão dos atletas.'); return;
    }

    const agora = Math.floor(Date.now() / 1000);
    let renovados = 0, validos = 0, erros = 0;

    tokens.forEach(function(t) {
      if (!t.refresh_token) { erros++; return; }
      if (t.expires_at && (t.expires_at - agora) > 10800) { validos++; return; }

      try {
        const refreshResp = UrlFetchApp.fetch('https://www.strava.com/oauth/token', {
          method: 'post',
          payload: {
            client_id:     clientId,
            client_secret: clientSecret,
            grant_type:    'refresh_token',
            refresh_token: t.refresh_token
          },
          muteHttpExceptions: true
        });

        if (refreshResp.getResponseCode() !== 200) {
          Logger.log('ERRO renovar ' + t.ath_id + ': HTTP ' + refreshResp.getResponseCode());
          erros++; return;
        }

        const novo = JSON.parse(refreshResp.getContentText());

        props.setProperty('RT_' + t.ath_id, novo.refresh_token);
        props.setProperty('AT_' + t.ath_id, novo.access_token);
        props.setProperty('EX_' + t.ath_id, String(novo.expires_at));

        supaUpsertToken(t.ath_id, t.nome, novo.access_token, novo.refresh_token, novo.expires_at);

        renovados++;
        Logger.log('OK: ' + t.ath_id + ' (' + t.nome + ') renovado.');
      } catch (e) {
        Logger.log('ERRO ' + t.ath_id + ': ' + e.message);
        erros++;
      }
    });

    Logger.log('Renovação concluída — renovados: ' + renovados + ' | válidos: ' + validos + ' | erros: ' + erros);
  } catch (e) {
    Logger.log('ERRO renovacaoProativaTokens: ' + e.message);
  }
}

// ─── INSTALAR TRIGGER DE RENOVAÇÃO PROATIVA ──────────────────────────────────
// Rodar UMA VEZ. Depois o GAS renova automaticamente a cada 4h para sempre.

function instalarAcionadorRenovacao() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'renovacaoProativaTokens')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('renovacaoProativaTokens')
    .timeBased().everyHours(4).create();
  Logger.log('Trigger instalado: renovacaoProativaTokens a cada 4h.');
}

// ─── STATUS STRAVA NO CADASTRO ────────────────────────────────────────────────
// Adiciona/atualiza coluna "🔴 Strava" no CADASTRO com status em tempo real.
// Fonte: ScriptProperties (primário) → Supabase (fallback).
// Roda automaticamente a cada 2h via trigger.

function atualizarStatusStravaEmCadastro() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const wsCad = ss.getSheetByName(H.SHEETS.CADASTRO);
  if (!wsCad) { Logger.log('CADASTRO não encontrado.'); return; }

  const props = PropertiesService.getScriptProperties();
  const agora = Math.floor(Date.now() / 1000);

  // ── Localizar cabeçalhos (linha 3) ──
  const lastCol = wsCad.getLastColumn();
  const lastRow = wsCad.getLastRow();
  if (lastRow < 4) { Logger.log('Nenhum atleta encontrado.'); return; }

  const headers = wsCad.getRange(3, 1, 1, lastCol).getValues()[0];

  // Coluna ID Atleta (índice 0 = coluna A)
  let idColIdx = headers.findIndex(h => String(h).trim() === 'ID Atleta');
  if (idColIdx < 0) idColIdx = 0;

  // Coluna Status Strava — procurar existente ou criar ao final
  let stravaColIdx = headers.findIndex(h => String(h).includes('Strava'));
  let stravaColNum;

  if (stravaColIdx >= 0) {
    stravaColNum = stravaColIdx + 1; // 1-based
    Logger.log('Coluna Strava já existe na coluna ' + stravaColNum);
  } else {
    stravaColNum = lastCol + 1;
    const hCell = wsCad.getRange(3, stravaColNum);
    hCell.setValue('🔴 Strava');
    hCell.setBackground('#b7b7b7').setFontWeight('bold')
         .setHorizontalAlignment('center');
    Logger.log('Coluna "🔴 Strava" criada na coluna ' + stravaColNum);
  }

  // ── Processar atletas (linha 4 em diante) ──
  const idData = wsCad.getRange(4, idColIdx + 1, lastRow - 3, 1).getValues();
  const updates = [];
  const colors  = [];

  idData.forEach(function(row) {
    const athId = String(row[0]).trim();

    // Pular linhas vazias ou de cabeçalho duplicado (bug conhecido no sheet)
    if (!athId || athId === 'ID Atleta' || athId === 'IDENTIFICAÇÃO' || athId.startsWith('Nome')) {
      updates.push(['']);
      colors.push([null]);
      return;
    }

    // 1. ScriptProperties (fonte primária)
    const rt = props.getProperty('RT_' + athId);

    let statusVal, bg, fg;

    if (rt && rt.length > 10) {
      statusVal = '✅ Ativo';
      bg = '#b7e1cd'; fg = '#0d652d';
    } else {
      // 2. Supabase (fallback)
      const supaData = supaGetRefresh(athId);
      if (supaData && supaData.refresh_token && supaData.refresh_token.length > 10) {
        statusVal = '🔄 Só Supabase';
        bg = '#fce8b2'; fg = '#7d5a0b';
      } else {
        statusVal = '❌ Reconectar';
        bg = '#f28b82'; fg = '#b31412';
      }
    }

    updates.push([statusVal]);
    colors.push([{ bg, fg }]);
  });

  // ── Escrever em lote ──
  if (updates.length > 0) {
    const range = wsCad.getRange(4, stravaColNum, updates.length, 1);
    range.setValues(updates);
    range.setHorizontalAlignment('center');

    // Colorir célula a célula (não tem setBackgrounds com objeto, usa loop)
    for (let i = 0; i < colors.length; i++) {
      if (!colors[i][0]) continue;
      const cell = wsCad.getRange(4 + i, stravaColNum);
      cell.setBackground(colors[i][0].bg).setFontColor(colors[i][0].fg);
    }

    Logger.log('✅ Status Strava atualizado no CADASTRO: ' + updates.filter(u => u[0]).length + ' atletas processados.');
  }
}

// ─── INSTALAR TRIGGER AUTOMÁTICO ─────────────────────────────────────────────
// Rodar UMA VEZ manualmente. Depois atualiza o CADASTRO a cada 2h sozinho.

function instalarAcionadorStatusCadastro() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'atualizarStatusStravaEmCadastro')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('atualizarStatusStravaEmCadastro')
    .timeBased().everyHours(2).create();
  Logger.log('✅ Trigger instalado: atualizarStatusStravaEmCadastro a cada 2h.');
}


// ─────────────────────────────────────────────────────────────────────────────
// SETUP SUPABASE — HIPERATIVO V3 (adicionado auditoria 21/06/2026)
// Salva SUPABASE_KEY nas ScriptProperties (URL já tem fallback)
// Executar via: Menu ⚙️ Configurações → Configurar Supabase
// ─────────────────────────────────────────────────────────────────────────────

function setupSupabase() {
  const ui    = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();

  const atual = props.getProperty('SUPABASE_KEY') || '';
  const r = ui.prompt(
    'Configurar Supabase',
    'Cole o SERVICE ROLE KEY do Supabase (Settings -> API Keys):\n\n' +
    '(atual: ' + (atual ? atual.substring(0, 12) + '...' : 'nao configurado') + ')',
    ui.ButtonSet.OK_CANCEL
  );

  if (r.getSelectedButton() !== ui.Button.OK) return;
  const rawKey = r.getResponseText().trim();
  if (!rawKey || rawKey.length < 20) {
    ui.alert('Chave invalida', 'A chave deve ter pelo menos 20 caracteres.', ui.ButtonSet.OK);
    return;
  }

  props.setProperty('SUPABASE_KEY', rawKey);
  if (!props.getProperty('SUPABASE_URL')) {
    props.setProperty('SUPABASE_URL', 'https://korlpbclqgmqvpbrungc.supabase.co');
  }

  ui.alert(
    'Supabase configurado!',
    'Chave salva. Execute diagnosticoSupabase() para verificar.',
    ui.ButtonSet.OK
  );
}


// ─────────────────────────────────────────────────────────────────────────────
// ENVIAR LINK STRAVA PARA DESCONECTADOS (bulk)
// Chamado pelo menu: 📧 Comunicacao -> Enviar link Strava para pendentes
// Envia email para cada atleta com STRAVA_OK != 'Sim' e STATUS = 'Ativo'
// REGRA: nao executa OAuth, apenas envia o link para o atleta clicar
// ─────────────────────────────────────────────────────────────────────────────

function enviarLinkStravaDesconectados() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(H.SHEETS.CADASTRO);
  if (!sheet) { SpreadsheetApp.getUi().alert('Aba CADASTRO nao encontrada.'); return; }

  const data      = sheet.getDataRange().getValues();
  const pendentes = [];

  for (let i = 1; i < data.length; i++) {
    const row      = data[i];
    const athId    = String(row[H.CAD.ID - 1]        || '').trim();
    const nome     = String(row[H.CAD.NOME - 1]      || '').trim();
    const email    = String(row[H.CAD.EMAIL - 1]     || '').trim();
    const stravaOk = String(row[H.CAD.STRAVA_OK - 1] || '').trim();
    const status   = String(row[H.CAD.STATUS - 1]    || '').trim();

    if (!_isAthIdValido_(athId))                 continue;
    if (stravaOk === 'Sim')                      continue;
    if (status === 'Inativo' || status === 'Cancelado') continue;
    if (!email || !email.includes('@'))          continue;

    pendentes.push({ athId, nome, email });
  }

  if (pendentes.length === 0) {
    SpreadsheetApp.getUi().alert('Nenhum pendente', 'Todos os atletas ativos ja estao conectados ao Strava.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  const ui   = SpreadsheetApp.getUi();
  const conf = ui.alert(
    'Enviar emails',
    'Enviar link de conexao Strava para ' + pendentes.length + ' atleta(s):\n' +
    pendentes.map(a => a.nome + ' (' + a.email + ')').slice(0, 8).join('\n') +
    (pendentes.length > 8 ? '\n... e mais ' + (pendentes.length - 8) : ''),
    ui.ButtonSet.OK_CANCEL
  );
  if (conf !== ui.Button.OK) return;

  let enviados = 0, erros = 0;
  pendentes.forEach(function(p) {
    try {
      var stravaUrl = _gerarUrlOAuth(p.athId);
      var prNome    = p.nome.split(' ')[0];
      var htmlBody  = '<div style="font-family:Arial,sans-serif;max-width:560px;padding:32px 24px;">' +
        '<h2 style="color:#FC4C02;">Conecte seu Strava</h2>' +
        '<p>Oi, <strong>' + prNome + '</strong>!</p>' +
        '<p>Para acompanharmos sua evolucao no Hiperativo, conecte seu Strava.</p>' +
        '<p style="margin:32px 0;text-align:center;">' +
        '<a href="' + stravaUrl + '" style="background:#FC4C02;color:#fff;padding:14px 28px;border-radius:6px;text-decoration:none;font-size:16px;font-weight:bold;">Conectar Strava</a></p>' +
        '<p style="color:#888;font-size:13px;">Link valido por 1 hora. Pediu novo? So chamar o treinador.</p>' +
        '<hr style="border:none;border-top:1px solid #eee;margin:24px 0;">' +
        '<p style="color:#aaa;font-size:12px;">Grupo Hiperativo — Mais que corrida. E atitude em movimento.</p></div>';
      MailApp.sendEmail({ to: p.email, subject: 'Conecte seu Strava ao Hiperativo', htmlBody: htmlBody, name: 'Grupo Hiperativo' });
      _log(p.athId, 'INFO', 'enviarLinkStravaDesconectados', 'Email enviado: ' + p.email, '');
      enviados++;
      Utilities.sleep(300);
    } catch(e) {
      _log(p.athId, 'ERRO', 'enviarLinkStravaDesconectados', e.message, '');
      erros++;
    }
  });

  ui.alert('Envio concluido', enviados + ' emails enviados.' + (erros > 0 ? ' ' + erros + ' erro(s) — ver aba ERROS.' : ''), ui.ButtonSet.OK);
}


// ─────────────────────────────────────────────────────────────────────────────
// GERAR MENSAGEM WHATSAPP COM LINK STRAVA
// Para atleta especifico — exibe dialog com texto pronto para copiar
// ─────────────────────────────────────────────────────────────────────────────

function gerarMensagemWhatsAppStrava() {
  var ui = SpreadsheetApp.getUi();
  var r  = ui.prompt('WhatsApp Strava', 'Digite o ID do atleta (ex: ATH123456):', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;

  var athId = r.getResponseText().trim().toUpperCase();
  if (!_isAthIdValido_(athId)) { ui.alert('ID invalido', 'Use o formato ATHXXXXXX.', ui.ButtonSet.OK); return; }

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(H.SHEETS.CADASTRO);
  var data  = sheet.getDataRange().getValues();
  var nome  = athId, whats = '';

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][H.CAD.ID - 1] || '').trim().toUpperCase() === athId) {
      nome  = String(data[i][H.CAD.NOME  - 1] || '').trim() || athId;
      whats = String(data[i][H.CAD.WHATS - 1] || '').trim();
      break;
    }
  }

  var stravaUrl;
  try { stravaUrl = _gerarUrlOAuth(athId); }
  catch(e) { ui.alert('Erro ao gerar link', e.message, ui.ButtonSet.OK); return; }

  var prNome = nome.split(' ')[0];
  var msg = 'Oi, ' + prNome + '! \n\n' +
    'Para acompanharmos sua evolucao no Hiperativo, precisamos conectar o seu Strava.\n\n' +
    'Clique no link abaixo e autorize:\n' + stravaUrl + '\n\n' +
    'E rapido — menos de 1 minuto.\n' +
    'Duvidas? Me chama aqui!\n\n' +
    '— Grupo Hiperativo\nMais que corrida. E atitude em movimento.';

  ui.alert('WhatsApp — ' + nome, (whats ? 'Numero: ' + whats + '\n\n' : '') + msg, ui.ButtonSet.OK);
  _log(athId, 'INFO', 'gerarMensagemWhatsAppStrava', 'Link gerado para: ' + nome, '');
}

function gerarMensagemWhatsAppStravaTeste() {
  var msg = 'Oi, Atleta Teste! \n\nPara acompanharmos sua evolucao no Hiperativo, precisamos conectar o seu Strava.\n\nLink: https://www.strava.com/oauth/TESTE\n\n— Grupo Hiperativo';
  console.log('=== TESTE WHATSAPP STRAVA ===\n' + msg);
  Logger.log('=== TESTE WHATSAPP STRAVA ===\n' + msg);
}


// ─────────────────────────────────────────────────────────────────────────────
// VERIFICAR STATUS REAL DO STRAVA (via API)
// Faz chamada leve a API Strava para cada atleta marcado como conectado.
// Atualiza STRAVA_OK para 'Reconectar' se o access nao funcionar.
// REGRA: nao executa OAuth, nao apaga refresh, nao migra atletas.
// ─────────────────────────────────────────────────────────────────────────────

function verificarStatusRealStrava() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(H.SHEETS.CADASTRO);
  if (!sheet) { SpreadsheetApp.getUi().alert('Aba CADASTRO nao encontrada.'); return; }

  var data         = sheet.getDataRange().getValues();
  var resultados   = [];
  var verificados  = 0;
  var falhos       = 0;

  for (var i = 1; i < data.length; i++) {
    var athId    = String(data[i][H.CAD.ID - 1]        || '').trim();
    var nome     = String(data[i][H.CAD.NOME - 1]      || '').trim();
    var stravaOk = String(data[i][H.CAD.STRAVA_OK - 1] || '').trim();

    if (!_isAthIdValido_(athId)) continue;
    if (stravaOk !== 'Sim')      continue;

    verificados++;
    try {
      var accessToken = _getValidAccessToken(athId);
      if (!accessToken) {
        sheet.getRange(i + 1, H.CAD.STRAVA_OK, 1, 1).setValue('Reconectar');
        resultados.push('AVISO: ' + nome + ' (' + athId + ') — token invalido → Reconectar');
        _log(athId, 'AVISO', 'verificarStatusRealStrava', 'Token invalido — atleta precisa reconectar', '');
        falhos++;
        continue;
      }
      var resp   = UrlFetchApp.fetch('https://www.strava.com/api/v3/athlete', {
        method: 'get', headers: { 'Authorization': 'Bearer ' + accessToken }, muteHttpExceptions: true
      });
      var status = resp.getResponseCode();
      if (status === 200) {
        resultados.push('OK: ' + nome + ' (' + athId + ')');
        _log(athId, 'INFO', 'verificarStatusRealStrava', 'Token verificado OK', '');
      } else if (status === 401) {
        sheet.getRange(i + 1, H.CAD.STRAVA_OK, 1, 1).setValue('Reconectar');
        resultados.push('AVISO: ' + nome + ' (' + athId + ') — 401 → Reconectar');
        _log(athId, 'AVISO', 'verificarStatusRealStrava', 'API retornou 401', '');
        falhos++;
      } else if (status === 429) {
        resultados.push('LIMITE: rate limit (429) — tentar mais tarde');
        break;
      } else {
        resultados.push('?: ' + nome + ' (' + athId + ') — HTTP ' + status);
      }
      Utilities.sleep(500);
    } catch(e) {
      resultados.push('ERRO: ' + nome + ' (' + athId + ') — ' + e.message);
      _log(athId, 'ERRO', 'verificarStatusRealStrava', e.message, '');
      falhos++;
    }
  }

  if (verificados === 0) {
    SpreadsheetApp.getUi().alert('Verificacao', 'Nenhum atleta com STRAVA_OK = Sim encontrado.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  var resumo = 'Verificados: ' + verificados + '\nOK: ' + (verificados - falhos) + '\nReconectar: ' + falhos + '\n\n' + resultados.join('\n');
  SpreadsheetApp.getUi().alert('Status Real Strava', resumo, SpreadsheetApp.getUi().ButtonSet.OK);
}

// ─────────────────────────────────────────────────────────────────────────────
// IMPORTAR ATIVIDADES APENAS PARA CONECTADOS
// Filtra: apenas atletas com STRAVA_OK = 'Sim' e STATUS = 'Ativo'.
// REGRA: nao importa de atletas desconectados.
// ─────────────────────────────────────────────────────────────────────────────

function importarAtividadesConectados() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(H.SHEETS.CADASTRO);
  if (!sheet) { SpreadsheetApp.getUi().alert('Aba CADASTRO nao encontrada.'); return; }

  var data       = sheet.getDataRange().getValues();
  var conectados = [];

  for (var i = 1; i < data.length; i++) {
    var athId    = String(data[i][H.CAD.ID - 1]        || '').trim();
    var stravaOk = String(data[i][H.CAD.STRAVA_OK - 1] || '').trim();
    var status   = String(data[i][H.CAD.STATUS - 1]    || '').trim();
    if (!_isAthIdValido_(athId))                           continue;
    if (stravaOk !== 'Sim')                               continue;
    if (status === 'Inativo' || status === 'Cancelado')   continue;
    conectados.push(athId);
  }

  if (conectados.length === 0) {
    Logger.log('[importarAtividadesConectados] Nenhum atleta conectado ao Strava (STRAVA_OK = Sim).');
    return;
  }

  var importados = 0, erros = 0;
  conectados.forEach(function(athId) {
    try {
      _importarAtividadesAtleta(athId);
      importados++;
      _log(athId, 'INFO', 'importarAtividadesConectados', 'Importacao concluida', '');
    } catch(e) {
      erros++;
      _log(athId, 'ERRO', 'importarAtividadesConectados', e.message, '');
    }
  });

  Logger.log('[importarAtividadesConectados] OK: ' + importados + ' atleta(s).' + (erros > 0 ? ' Erros: ' + erros + ' - ver aba ERROS.' : ''));
}


// ─── SYNC TOKENS → SUPABASE ──────────────────────────────────────────────────
function sincronizarTokensParaSupabase() {
  if (!_supaConfigurado_()) {
    Logger.log('[sincronizarTokens] SUPABASE_KEY nao configurada. Abortando.');
    return;
  }
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ws = ss.getSheetByName(H.SHEETS.TOKENS);
  if (!ws) {
    Logger.log('[sincronizarTokens] Aba TOKENS nao encontrada.');
    return;
  }
  var data = ws.getDataRange().getValues();
  var ok = 0, skip = 0, erros = 0;
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var athId, nome, accessToken, refreshToken, expiresAt;
    var colA = String(row[0] || '').trim();
    if (colA.indexOf('ATH') === 0) {
      // Linha estilo direto: ATH_ID na col A (ex: row 12)
      athId = colA;
      nome = String(row[1] || '').trim();
      accessToken = String(row[2] || '').trim();
      refreshToken = String(row[3] || '').trim();
      expiresAt = Number(row[4] || 0);
    } else {
      // Padrao ou col A vazia: ATH_ID na col B
      athId = String(row[1] || '').trim();
      nome = String(row[2] || '').trim();
      accessToken = String(row[3] || '').trim();
      refreshToken = String(row[4] || '').trim();
      expiresAt = Number(row[5] || 0);
    }
    if (!athId || athId.indexOf('ATH') !== 0) {
      Logger.log('[sincronizarTokens] L' + (i+1) + ': athId invalido ("' + athId + '"). Skip.');
      skip++;
      continue;
    }
    if (!accessToken || accessToken.length < 10) {
      Logger.log('[sincronizarTokens] L' + (i+1) + ' (' + athId + '): access_token vazio. Skip.');
      skip++;
      continue;
    }
    if (!refreshToken || refreshToken.length < 5) {
      Logger.log('[sincronizarTokens] L' + (i+1) + ' (' + athId + '): refresh_token vazio. Skip (seguranca: nao sobrescrever com vazio).');
      skip++;
      continue;
    }
    try {
      supaUpsertToken(athId, nome, accessToken, refreshToken, expiresAt);
      Logger.log('[sincronizarTokens] OK: ' + athId + ' (' + nome + ')');
      ok++;
    } catch(e) {
      Logger.log('[sincronizarTokens] ERRO ' + athId + ': ' + e.message);
      erros++;
    }
  }
  Logger.log('[sincronizarTokens] Concluido. OK: ' + ok + ' | Skip: ' + skip + ' | Erros: ' + erros);
}