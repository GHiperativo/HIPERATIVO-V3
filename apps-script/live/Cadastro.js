/**
 * ═══════════════════════════════════════════════════════════════════════
 * HIPERATIVO V3 — Cadastro.gs  (v2.1 — envio email/WhatsApp 04/06/2026)
 * Cadastro de atletas, geração e envio de links
 * ═══════════════════════════════════════════════════════════════════════
 */

function limparErrosFormula() {
  // Compatibilidade com o item antigo: reparar fórmulas conhecidas é seguro;
  // apagar toda célula com erro pode destruir a fórmula que precisa ser corrigida.
  return repararFormulasOperacionais(true);
}

function setupInicial() {
  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const props    = PropertiesService.getScriptProperties();
  const problemas = [];
  const ok       = [];

  // Verificar abas obrigatórias
  for (const [key, nome] of Object.entries(H.SHEETS)) {
    if (ss.getSheetByName(nome)) {
      ok.push('✅ Aba: ' + nome);
    } else {
      problemas.push('❌ Aba ausente: ' + nome);
    }
  }

  // Verificar credenciais
  if (props.getProperty('STRAVA_CLIENT_ID'))     ok.push('✅ STRAVA_CLIENT_ID configurado');
  else problemas.push('⚠️ STRAVA_CLIENT_ID não configurado');

  if (props.getProperty('STRAVA_CLIENT_SECRET'))  ok.push('✅ STRAVA_CLIENT_SECRET configurado');
  else problemas.push('⚠️ STRAVA_CLIENT_SECRET não configurado');

  if (props.getProperty('WEBAPP_URL'))            ok.push('✅ WEBAPP_URL configurado');
  else problemas.push('⚠️ WEBAPP_URL não configurado (implante o WebApp primeiro)');

  // Verificar triggers
  const triggers = ScriptApp.getProjectTriggers();
  if (triggers.length) {
    ok.push('✅ ' + triggers.length + ' trigger(s) ativo(s)');
  } else {
    problemas.push('⚠️ Nenhum trigger ativo — use ⚡ HIPERATIVO > Criar trigger (6h)');
  }

  const msg = (problemas.length
    ? '🚨 PROBLEMAS ENCONTRADOS:\n' + problemas.join('\n') + '\n\n'
    : '🎉 Sem problemas!\n\n')
    + 'STATUS:\n' + ok.join('\n');

  try {
    SpreadsheetApp.getUi().alert('🔍 Diagnóstico HIPERATIVO V3', msg, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (_) {}

  _log('SISTEMA', problemas.length ? 'WARN' : 'INFO', 'setupInicial', 'Diagnóstico: ' + problemas.length + ' problema(s)', msg);
}

function gerarLinkCadastro() {
  const ui    = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();
  const url   = props.getProperty('WEBAPP_URL') || '';

  if (!url) {
    ui.alert(
      '⚠️ WebApp não implantado',
      'Antes de gerar o link, você precisa:\n\n' +
      '1. Clicar em "Implantar" (botão azul, canto superior direito)\n' +
      '2. Selecionar "Gerenciar implantações"\n' +
      '3. Clicar no lápis ✏️ > "Nova versão"\n' +
      '4. Quem tem acesso: "Qualquer pessoa"\n' +
      '5. Salvar e copiar a URL /exec\n' +
      '6. Ir em ⚡ HIPERATIVO > Configurar credenciais e colar a URL',
      ui.ButtonSet.OK
    );
    return;
  }

  const r = ui.prompt(
    '🔗 Gerar Link de Cadastro',
    'Digite o ID do atleta (ex: ATH_001):\n(Deixe vazio para gerar link genérico sem ID)',
    ui.ButtonSet.OK_CANCEL
  );
  if (r.getSelectedButton() !== ui.Button.OK) return;

  const athId        = (r.getResponseText() || '').trim().toUpperCase();
  const linkCadastro = url + '?cadastro=true'
    + (athId ? '&athId=' + encodeURIComponent(athId) : '')
    + '&utm_source=link_direto&utm_medium=convite&utm_campaign=cadastro_hiperativo';

  ui.alert(
    '🔗 Link de Cadastro do Atleta',
    'Envie este link para o atleta se cadastrar e conectar o Strava:\n\n' +
    linkCadastro + '\n\n' +
    '📋 O atleta vai:\n' +
    '1. Preencher nome, e-mail, altura, objetivos e plano\n' +
    '2. Aceitar os termos LGPD\n' +
    '3. Clicar em "Salvar e Conectar ao Strava"\n' +
    '4. Autorizar o app Hiperativo no Strava\n' +
    '5. Pronto! Dados salvos automaticamente na planilha',
    ui.ButtonSet.OK
  );

  if (athId) _log(athId, 'INFO', 'gerarLinkCadastro', 'Link de cadastro gerado', linkCadastro);
}

// ── GERAR E ENVIAR LINK POR EMAIL ──────────────────────────────────────────

/**
 * Gera e envia o link de cadastro por email para o atleta.
 * Pede: nome, email, (opcional) ATH_ID.
 * Envia email com link personalizado + instruções para conectar Strava.
 */
function gerarLinkCadastroEmail() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();
  const webAppUrl = props.getProperty('WEBAPP_URL');
  
  if (!webAppUrl) {
    ui.alert('❌ URL do WebApp não configurada.\nVá em ⚙️ Configurações → Configurar credenciais.');
    return;
  }
  
  // Pedir nome do atleta
  const rNome = ui.prompt('✉️ Enviar Link de Cadastro',
    'Nome completo do atleta:',
    ui.ButtonSet.OK_CANCEL);
  if (rNome.getSelectedButton() !== ui.Button.OK) return;
  const nome = rNome.getResponseText().trim();
  if (!nome) { ui.alert('❌ Nome não pode ser vazio.'); return; }
  
  // Pedir email do atleta
  const rEmail = ui.prompt('✉️ Enviar Link de Cadastro',
    'E-mail do atleta (' + nome + '):',
    ui.ButtonSet.OK_CANCEL);
  if (rEmail.getSelectedButton() !== ui.Button.OK) return;
  const email = rEmail.getResponseText().trim();
  if (!email || !email.includes('@')) { ui.alert('❌ E-mail inválido.'); return; }
  
  // Gerar ATH_ID automático ou pedir
  const rId = ui.prompt('✉️ Enviar Link de Cadastro',
    'ATH_ID (deixe em branco para gerar automaticamente):',
    ui.ButtonSet.OK_CANCEL);
  if (rId.getSelectedButton() !== ui.Button.OK) return;
  
  let athId = rId.getResponseText().trim().toUpperCase();
  if (!athId) {
    const shCad = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(H.SHEETS.CADASTRO);
    const dadosCad = shCad ? shCad.getDataRange().getValues() : [];
    athId = _gerarAthIdComBase_(dadosCad);
  }
  
  // Montar link personalizado com UTM params e nome do atleta
  const nomeEncoded = encodeURIComponent(nome);
  const link = webAppUrl + '?cadastro=true&athId=' + encodeURIComponent(athId) +
               '&ref=email' +
               '&utm_source=email&utm_medium=convite&utm_campaign=cadastro_hiperativo';
  
  // Buscar email admin para remetente
  const adminEmail = props.getProperty('ADMIN_EMAIL') || 'contato@ghiperativo.com.br';
  const primeiroNome = nome.split(' ')[0];
  
  // Corpo do email em HTML
  const corpoHtml = _htmlEmailCadastro(primeiroNome, link, athId);
  
  try {
    MailApp.sendEmail({
      to: email,
      replyTo: adminEmail,
      subject: '⚡ ' + primeiroNome + ', seu link de cadastro HIPERATIVO está aqui!',
      htmlBody: corpoHtml
    });
    
    // Log na aba ERROS (nivel INFO)
    _log(athId, 'INFO', 'gerarLinkCadastroEmail', 
         'Link de cadastro enviado para ' + email + ' | Nome: ' + nome, '');
    
    ui.alert('✅ Link enviado com sucesso!\n\n' +
             '📧 Para: ' + email + '\n' +
             '👤 Atleta: ' + nome + '\n' +
             '🆔 ATH_ID: ' + athId + '\n\n' +
             '🔗 Link: ' + link);
  } catch(e) {
    ui.alert('❌ Erro ao enviar email: ' + e.message);
    _log(athId, 'ERRO', 'gerarLinkCadastroEmail', 'Falha ao enviar email para ' + email, e.stack);
  }
}

/**
 * Gera o corpo HTML do email de convite para cadastro.
 */
function _htmlEmailCadastro(primeiroNome, link, athId) {
  return `<!DOCTYPE html>
<html lang="pt-BR">
<head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1"></head>
<body style="margin:0;padding:0;background:#f4f4f4;font-family:Arial,Helvetica,sans-serif">
<table width="100%" cellpadding="0" cellspacing="0" style="background:#f4f4f4;padding:20px 0">
<tr><td align="center">
<table width="600" cellpadding="0" cellspacing="0" style="background:#fff;border-radius:12px;overflow:hidden;box-shadow:0 4px 16px rgba(0,0,0,.12)">

  <!-- HEADER -->
  <tr><td style="background:linear-gradient(135deg,#1a3a8a 0%,#0d2560 100%);padding:32px 40px;text-align:center">
    <div style="font-size:48px;margin-bottom:8px">⚡</div>
    <div style="color:#00c853;font-size:28px;font-weight:900;letter-spacing:2px">HIPERATIVO</div>
    <div style="color:#fff;font-size:12px;letter-spacing:3px;margin-top:4px;opacity:.8">CABEÇA • CORAÇÃO • CORPO</div>
  </td></tr>

  <!-- CORPO -->
  <tr><td style="padding:40px">
    <h2 style="color:#0d2560;margin:0 0 16px">Olá, ${primeiroNome}! 👋</h2>
    <p style="color:#333;line-height:1.6;margin:0 0 24px">
      Você foi convidado(a) para fazer parte do time <strong>HIPERATIVO</strong>!<br>
      Clique no botão abaixo para completar seu cadastro e escolher seu programa de treino.
    </p>
    
    <div style="text-align:center;margin:32px 0">
      <a href="${link}" style="display:inline-block;background:linear-gradient(135deg,#00c853,#009624);color:#fff;text-decoration:none;padding:16px 40px;border-radius:30px;font-size:18px;font-weight:700;letter-spacing:1px;box-shadow:0 4px 12px rgba(0,200,83,.4)">
        ⚡ FAZER MEU CADASTRO
      </a>
    </div>
    
    <p style="color:#666;font-size:13px;line-height:1.6;margin:0 0 16px">
      Ou copie e cole este link no seu navegador:<br>
      <a href="${link}" style="color:#1a3a8a;word-break:break-all">${link}</a>
    </p>
    
    <div style="background:#f0f4ff;border-left:4px solid #1a3a8a;padding:16px;border-radius:0 8px 8px 0;margin:24px 0">
      <p style="margin:0;color:#0d2560;font-size:13px">
        <strong>📌 Seu código:</strong> <code style="background:#fff;padding:2px 8px;border-radius:4px;font-size:14px;color:#00c853;font-weight:700">${athId}</code><br>
        <span style="font-size:12px;color:#666">Guarde este código — você precisará dele para acessar seus dados.</span>
      </p>
    </div>
    
    <h3 style="color:#0d2560;margin:24px 0 12px">🏃 Nossos programas:</h3>
    <ul style="color:#333;line-height:1.8;margin:0;padding-left:20px">
      <li>⚡ Alta Voltagem</li>
      <li>🏅 Corrida CCC (Cabeça Coração Corpo)</li>
      <li>🌱 Iniciante em Movimento</li>
      <li>🎯 5k/10k em 6 Semanas</li>
      <li>🏃 Hiperativo Running Club</li>
      <li>🌟 Vida Ativa na Melhor Idade</li>
    </ul>
    
    <p style="color:#999;font-size:12px;margin:32px 0 0;padding-top:16px;border-top:1px solid #eee">
      Este email foi enviado pelo sistema HIPERATIVO V3.<br>
      Em caso de dúvidas, responda este email ou entre em contato com seu treinador.
    </p>
  </td></tr>

  <!-- FOOTER -->
  <tr><td style="background:#0d2560;padding:20px;text-align:center">
    <div style="color:#fff;font-size:11px;opacity:.7">
      ⚡ HIPERATIVO — CABEÇA • CORAÇÃO • CORPO<br>
      © 2026 Grupo Hiperativo. Todos os direitos reservados.
    </div>
  </td></tr>

</table>
</td></tr>
</table>
</body>
</html>`;
}

// ── GERAR LINK WHATSAPP ───────────────────────────────────────────────────────

/**
 * Gera um link de cadastro formatado para WhatsApp.
 * Mostra o link com mensagem pronta para copiar e colar no WhatsApp.
 */
function gerarLinkCadastroWhatsapp() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();
  const webAppUrl = props.getProperty('WEBAPP_URL');
  
  if (!webAppUrl) {
    ui.alert('❌ URL do WebApp não configurada.\nVá em ⚙️ Configurações → Configurar credenciais.');
    return;
  }
  
  const rNome = ui.prompt('📲 Link WhatsApp',
    'Nome do atleta (para personalizar a mensagem):',
    ui.ButtonSet.OK_CANCEL);
  if (rNome.getSelectedButton() !== ui.Button.OK) return;
  const nome = rNome.getResponseText().trim() || 'atleta';
  const primeiroNome = nome.split(' ')[0];
  
  const shCad = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(H.SHEETS.CADASTRO);
  const dadosCad = shCad ? shCad.getDataRange().getValues() : [];
  const athId = _gerarAthIdComBase_(dadosCad);
  const link = webAppUrl + '?cadastro=true&athId=' + encodeURIComponent(athId)
    + '&ref=whatsapp&utm_source=whatsapp&utm_medium=convite&utm_campaign=cadastro_hiperativo';
  
  const mensagem = [
    'Olá, ' + primeiroNome + '! 👋',
    '',
    'Seja bem-vindo(a) ao ⚡ *HIPERATIVO*!',
    '',
    'Clique no link abaixo para fazer seu cadastro e escolher seu programa de treino:',
    '',
    '🔗 ' + link,
    '',
    '_Seu código: *' + athId + '*_',
    '',
    '*CABEÇA • CORAÇÃO • CORPO* 💪',
  ].join('\n');
  
  // WhatsApp API link (opens web.whatsapp.com or app)
  const whatsUrl = 'https://wa.me/?text=' + encodeURIComponent(mensagem);
  
  ui.alert('📲 Mensagem para WhatsApp',
    'Copie a mensagem abaixo ou use o link wa.me:\n\n' + mensagem + 
    '\n\n─────────────────────────\n' +
    '🔗 Link WhatsApp direto:\n' + whatsUrl.substring(0, 100) + '...',
    ui.ButtonSet.OK);
}

// ── ENVIO LINK STRAVA POR EMAIL ───────────────────────────────────────────────

/**
 * Envia o link de reconexão Strava por email para um atleta específico.
 */
function enviarLinkStravaEmail() {
  const ui = SpreadsheetApp.getUi();
  
  const rId = ui.prompt('📡 Enviar Link Strava',
    'ATH_ID do atleta (ex: ATH001):',
    ui.ButtonSet.OK_CANCEL);
  if (rId.getSelectedButton() !== ui.Button.OK) return;
  const athId = rId.getResponseText().trim().toUpperCase();
  if (!athId) return;
  
  // Buscar dados do atleta na planilha
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(H.SHEETS.CADASTRO);
  const dados = sh.getDataRange().getValues();
  let atletaRow = null;
  
  for (let i = 1; i < dados.length; i++) {
    if (String(dados[i][H.CAD.ID - 1]).toUpperCase() === athId) {
      atletaRow = dados[i];
      break;
    }
  }
  
  if (!atletaRow) {
    ui.alert('❌ Atleta ' + athId + ' não encontrado na planilha.');
    return;
  }
  
  const nome = atletaRow[H.CAD.NOME - 1] || athId;
  const emailAtleta = atletaRow[H.CAD.EMAIL - 1] || '';
  
  if (!emailAtleta || !emailAtleta.includes('@')) {
    ui.alert('❌ Atleta ' + nome + ' não tem e-mail cadastrado.');
    return;
  }
  
  // Usar gerarLinkStrava para obter a URL de OAuth
  const props = PropertiesService.getScriptProperties();
  const clientId = props.getProperty('STRAVA_CLIENT_ID');
  const webAppUrl = props.getProperty('WEBAPP_URL');
  
  if (!clientId || !webAppUrl) {
    ui.alert('❌ Credenciais Strava não configuradas.');
    return;
  }
  
  // Usar o fluxo OAuth único com state assinado pelo ATH_ID. O fluxo antigo
  // montava um callback sem state e o retorno acabava caindo no cadastro.
  const oauthUrl = _gerarUrlOAuth(athId);
  
  const primeiroNome = String(nome).split(' ')[0];
  const adminEmail = props.getProperty('ADMIN_EMAIL') || 'contato@ghiperativo.com.br';
  
  const corpoHtml = `<!DOCTYPE html>
<html lang="pt-BR">
<head><meta charset="UTF-8"></head>
<body style="font-family:Arial,sans-serif;background:#f4f4f4;padding:20px">
<table width="600" style="background:#fff;border-radius:12px;overflow:hidden;margin:auto">
<tr><td style="background:linear-gradient(135deg,#1a3a8a,#0d2560);padding:24px;text-align:center">
  <div style="color:#00c853;font-size:24px;font-weight:900">⚡ HIPERATIVO</div>
  <div style="color:#fff;font-size:11px;letter-spacing:2px">CABEÇA • CORAÇÃO • CORPO</div>
</td></tr>
<tr><td style="padding:32px">
  <h2 style="color:#0d2560">Conecte seu Strava, ${primeiroNome}! 🏃</h2>
  <p style="color:#333;line-height:1.6">Para que possamos importar seus treinos automaticamente, precisamos que você conecte sua conta <strong>Strava</strong> ao sistema HIPERATIVO.</p>
  <div style="text-align:center;margin:24px 0">
    <a href="${oauthUrl}" style="display:inline-block;background:#FC4C02;color:#fff;text-decoration:none;padding:14px 32px;border-radius:24px;font-weight:700;font-size:16px">
      🔗 CONECTAR MINHA CONTA STRAVA
    </a>
  </div>
  <p style="color:#666;font-size:12px">Ao clicar, você será redirecionado para o Strava para autorizar o acesso. Seus dados ficam seguros — só lemos suas atividades, não fazemos nenhuma alteração.</p>
</td></tr>
<tr><td style="background:#0d2560;padding:16px;text-align:center">
  <div style="color:#fff;font-size:10px;opacity:.7">⚡ HIPERATIVO © 2026</div>
</td></tr>
</table>
</body></html>`;
  
  try {
    MailApp.sendEmail({
      to: emailAtleta,
      replyTo: adminEmail,
      subject: '⚡ ' + primeiroNome + ', conecte seu Strava ao HIPERATIVO!',
      htmlBody: corpoHtml
    });
    
    ui.alert('✅ Link Strava enviado!\n\n📧 Para: ' + emailAtleta + '\n👤 Atleta: ' + nome);
  } catch(e) {
    ui.alert('❌ Erro ao enviar: ' + e.message);
  }
}

// ══════════════════════════════════════════════════════
// RESTAURAR ESTRUTURA (seguro — não apaga dados)
// ══════════════════════════════════════════════════════
function restaurarEstrutura() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var corrigidas = [];

  try {
    // 1. Configurar PropertiesService com valores conhecidos
    var props = PropertiesService.getScriptProperties();
    props.setProperty('STRAVA_CLIENT_ID', '153043');

    // Tentar pegar WEBAPP_URL do CONFIG se não estiver configurado
    if (!props.getProperty('WEBAPP_URL')) {
      var cfg = ss.getSheetByName(H.SHEETS.CONFIG);
      if (cfg && cfg.getLastRow() > 1) {
        var cfgData = cfg.getRange(2, 1, cfg.getLastRow() - 1, 2).getValues();
        cfgData.forEach(function(row) {
          if (String(row[0]).trim() === 'WEBAPP_URL' && row[1]) {
            props.setProperty('WEBAPP_URL', String(row[1]).trim());
          }
        });
      }
    }

    // 2. Restaurar headers de cada aba se estiverem vazios
    var abas = [
      { nome: H.SHEETS.CADASTRO,   criarFn: _criarCadastro },
      { nome: H.SHEETS.ATIVIDADES, criarFn: _criarAtividades },
      { nome: H.SHEETS.PLANO,      criarFn: _criarPlanoSemanal },
      { nome: H.SHEETS.METRICAS,   criarFn: _criarMetricas },
      { nome: H.SHEETS.FEEDBACK,   criarFn: _criarFeedback },
      { nome: H.SHEETS.GRAFICOS,   criarFn: _criarGraficos },
      { nome: H.SHEETS.ERROS,      criarFn: _criarErros }
    ];

    abas.forEach(function(aba) {
      var ws = ss.getSheetByName(aba.nome);
      if (!ws) {
        ws = ss.insertSheet(aba.nome);
        try { aba.criarFn(ws); } catch(exFreeze) { if (String(exFreeze.message).indexOf("congelar") === -1 && String(exFreeze.message).indexOf("freeze") === -1) throw exFreeze; try { ws.setFrozenRows(2); } catch(e2) {} }
        corrigidas.push(aba.nome + ' (criada)');
        return;
      }
      // Verifica se a aba está sem headers (linha 1 vazia)
      var cel = ws.getRange(1, 1).getValue();
      if (!cel || String(cel).trim() === '') {
        try { aba.criarFn(ws); } catch(exFreeze) { if (String(exFreeze.message).indexOf("congelar") === -1 && String(exFreeze.message).indexOf("freeze") === -1) throw exFreeze; try { ws.setFrozenRows(2); } catch(e2) {} }
        corrigidas.push(aba.nome + ' (headers restaurados)');
      }
    });

    // 3. Restaurar PAINEL separado (tem lógica própria)
    var painel = ss.getSheetByName(H.SHEETS.PAINEL);
    if (!painel) {
      painel = ss.insertSheet(H.SHEETS.PAINEL, 0);
      _criarPainel(painel);
      corrigidas.push(H.SHEETS.PAINEL + ' (criado)');
    }

    _log('SYSTEM', 'INFO', 'restaurarEstrutura', 'Restaurado: ' + (corrigidas.length || 'nada necessário'), '');

    var msg = corrigidas.length > 0
      ? '✅ Estrutura restaurada!\n\nCorrigidas:\n• ' + corrigidas.join('\n• ')
      : '✅ Estrutura já estava correta. Nenhuma alteração necessária.';

    ui.alert('Restauração Concluída', msg, ui.ButtonSet.OK);

  } catch(e) {
    _log('SYSTEM', 'ERRO', 'restaurarEstrutura', e.message, e.stack || '');
    ui.alert('❌ Erro na restauração', e.message, ui.ButtonSet.OK);
  }
}


// ══════════════════════════════════════════════════════
// FORÇAR RECRIAÇÃO DO PAINEL (fórmulas + visual)
// ══════════════════════════════════════════════════════
function forcarRecriarPainel() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  try {
    var painel = ss.getSheetByName(H.SHEETS.PAINEL);
    if (!painel) {
      painel = ss.insertSheet(H.SHEETS.PAINEL, 0);
    }
    _criarPainel(painel);
    _log('SYSTEM', 'INFO', 'forcarRecriarPainel', 'PAINEL recriado com sucesso', '');
    ui.alert('✅ PAINEL', 'Painel recriado com sucesso!\nVerifique se os #ERROR! foram resolvidos.', ui.ButtonSet.OK);
  } catch(e) {
    _log('SYSTEM', 'ERRO', 'forcarRecriarPainel', e.message, e.stack || '');
    ui.alert('❌ Erro', e.message, ui.ButtonSet.OK);
  }
}


// ══════════════════════════════════════════════════════
// CORRIGIR FÓRMULAS DO PAINEL (pt-BR, colunas corretas)
// ══════════════════════════════════════════════════════
function corrigirFormulasPainel(exibirAlerta) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  if (exibirAlerta === undefined) exibirAlerta = true;
  try {
    var painel = ss.getSheetByName(H.SHEETS.PAINEL);
    if (!painel) throw new Error("Aba PAINEL nao encontrada");
    var c = H.SHEETS.CADASTRO;
    var a = H.SHEETS.ATIVIDADES;
    var cad = ss.getSheetByName(c);
    var ativ = ss.getSheetByName(a);
    var cadLast = Math.max(3, cad ? cad.getLastRow() : 3);
    var ativLast = Math.max(3, ativ ? ativ.getLastRow() : 3);
    var q = "\"";
    // Row 6: metricas (STATUS=col26=Z, STRAVA_OK=col24=X) — en-US setFormula
    painel.getRange(6,1).setFormula("=COUNTIF('" + c + "'!A3:A" + cadLast + ";" + q + "ATH*" + q + ")");
    painel.getRange(6,3).setFormula("=COUNTIFS('" + c + "'!A3:A" + cadLast + ";" + q + "ATH*" + q + ";'" + c + "'!Z3:Z" + cadLast + ";" + q + "Ativo" + q + ")");
    painel.getRange(6,5).setFormula("=COUNTIFS('" + c + "'!A3:A" + cadLast + ";" + q + "ATH*" + q + ";'" + c + "'!X3:X" + cadLast + ";" + q + "Sim" + q + ")");
    painel.getRange(6,7).setFormula("=COUNTIF('" + a + "'!D3:D" + ativLast + ";" + q + ">" + q + "&TODAY()-7)");
    painel.getRange(6,9).setFormula("=IFERROR(TEXT(AVERAGEIF('" + a + "'!E3:E" + ativLast + ";" + q + "Corrida" + q + ";'" + a + "'!O3:O" + ativLast + ")/86400;" + q + "[mm]:ss" + q + ");" + q + "--" + q + ")");
    painel.getRange(6,11).setFormula("=IFERROR(ROUND(SUMIF('" + a + "'!E3:E" + ativLast + ";" + q + "Corrida" + q + ";'" + a + "'!L3:L" + ativLast + ")/MAX(1;COUNTIF('" + a + "'!E3:E" + ativLast + ";" + q + "Corrida" + q + "));1);" + q + "--" + q + ")");
    // Rows 11-20: ultimas atividades
    for (var i = 0; i < 10; i++) {
      var r = 11 + i;
      var rk = i + 1;
      painel.getRange(r,1).setFormula("=IFERROR(TEXT(LARGE('" + a + "'!D3:D" + ativLast + ";" + rk + ");" + q + "dd/mm/yyyy" + q + ");" + q + q + ")");
      painel.getRange(r,2).setFormula("=IFERROR(INDEX('" + a + "'!C3:C" + ativLast + ";MATCH(LARGE('" + a + "'!D3:D" + ativLast + ";" + rk + ");'" + a + "'!D3:D" + ativLast + ";0));" + q + q + ")");
      painel.getRange(r,3).setFormula("=IFERROR(INDEX('" + a + "'!E3:E" + ativLast + ";MATCH(LARGE('" + a + "'!D3:D" + ativLast + ";" + rk + ");'" + a + "'!D3:D" + ativLast + ";0));" + q + q + ")");
      painel.getRange(r,4).setFormula("=IFERROR(TEXT(INDEX('" + a + "'!L3:L" + ativLast + ";MATCH(LARGE('" + a + "'!D3:D" + ativLast + ";" + rk + ");'" + a + "'!D3:D" + ativLast + ";0));" + q + "0.00" + q + ")&" + q + " km" + q + ";" + q + q + ")");
      painel.getRange(r,5).setFormula("=IFERROR(INDEX('" + a + "'!P3:P" + ativLast + ";MATCH(LARGE('" + a + "'!D3:D" + ativLast + ";" + rk + ");'" + a + "'!D3:D" + ativLast + ";0));" + q + q + ")");
      painel.getRange(r,6).setFormula("=IFERROR(IF(INDEX('" + a + "'!Q3:Q" + ativLast + ";MATCH(LARGE('" + a + "'!D3:D" + ativLast + ";" + rk + ");'" + a + "'!D3:D" + ativLast + ";0))>0;ROUND(INDEX('" + a + "'!Q3:Q" + ativLast + ";MATCH(LARGE('" + a + "'!D3:D" + ativLast + ";" + rk + ");'" + a + "'!D3:D" + ativLast + ";0));0)&" + q + " bpm" + q + ";" + q + "--" + q + ");" + q + q + ")");
    }
    _log("SYSTEM","INFO","corrigirFormulasPainel","Formulas do PAINEL corrigidas","");
    if (exibirAlerta) ui.alert("Fórmulas corrigidas!","As métricas do PAINEL foram atualizadas.",ui.ButtonSet.OK);
  } catch(e) {
    _log("SYSTEM","ERRO","corrigirFormulasPainel",e.message,"");
    if (exibirAlerta) ui.alert("Erro", e.message, ui.ButtonSet.OK);
    throw e;
  }
}

function corrigirFormulasGraficos(exibirAlerta) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  if (exibirAlerta === undefined) exibirAlerta = true;
  var graf = ss.getSheetByName(H.SHEETS.GRAFICOS);
  var ativ = ss.getSheetByName(H.SHEETS.ATIVIDADES);
  var met = ss.getSheetByName(H.SHEETS.METRICAS);
  if (!graf || !ativ) throw new Error('Abas GRÁFICOS ou ATIVIDADES não encontradas.');
  var a = H.SHEETS.ATIVIDADES;
  var m = H.SHEETS.METRICAS;
  var lastA = Math.max(3, ativ.getLastRow());
  var lastM = Math.max(3, met ? met.getLastRow() : 3);
  var q = '"';
  graf.getRange('B4').setFormula("=IFERROR(SUMIFS('" + a + "'!L3:L" + lastA + ";'" + a + "'!E3:E" + lastA + ";" + q + "Corrida" + q + ";'" + a + "'!D3:D" + lastA + ";" + q + ">=" + q + "&DATE(YEAR(TODAY());MONTH(TODAY());1);'" + a + "'!D3:D" + lastA + ";" + q + "<" + q + "&DATE(YEAR(TODAY());MONTH(TODAY())+1;1));0)");
  graf.getRange('B5').setFormula("=COUNTIFS('" + a + "'!E3:E" + lastA + ";" + q + "Corrida" + q + ";'" + a + "'!D3:D" + lastA + ";" + q + ">=" + q + "&DATE(YEAR(TODAY());MONTH(TODAY());1);'" + a + "'!D3:D" + lastA + ";" + q + "<" + q + "&DATE(YEAR(TODAY());MONTH(TODAY())+1;1))");
  graf.getRange('B6').setFormula("=IFERROR(TEXT(AVERAGEIFS('" + a + "'!O3:O" + lastA + ";'" + a + "'!E3:E" + lastA + ";" + q + "Corrida" + q + ";'" + a + "'!O3:O" + lastA + ";" + q + ">0" + q + ")/86400;" + q + "[mm]:ss" + q + ");" + q + "--" + q + ")");
  graf.getRange('B7').setFormula("=IFERROR(ROUND(AVERAGEIFS('" + a + "'!Q3:Q" + lastA + ";'" + a + "'!E3:E" + lastA + ";" + q + "Corrida" + q + ";'" + a + "'!Q3:Q" + lastA + ";" + q + ">0" + q + ");0);" + q + "--" + q + ")");
  if (met) graf.getRange('B8').setFormula("=IFERROR(ROUND(AVERAGE('" + m + "'!D3:D" + lastM + ");1);" + q + "--" + q + ")");
  _log('SISTEMA', 'INFO', 'corrigirFormulasGraficos', 'Fórmulas de GRÁFICOS reparadas até a linha ' + lastA, '');
  if (exibirAlerta) ui.alert('✅ Fórmulas de GRÁFICOS reparadas.');
  return 5;
}

function repararFormulasOperacionais(exibirAlerta) {
  if (exibirAlerta === undefined) exibirAlerta = true;
  corrigirFormulasPainel(false);
  corrigirFormulasGraficos(false);
  if (exibirAlerta) SpreadsheetApp.getUi().alert(
    '✅ Fórmulas reparadas',
    'PAINEL e GRÁFICOS foram atualizados com o tamanho real das abas. Nenhuma fórmula foi apagada.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
  return true;
}
