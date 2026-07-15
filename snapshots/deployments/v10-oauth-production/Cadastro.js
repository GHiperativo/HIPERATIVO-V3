/**
 * ═══════════════════════════════════════════════════════════════════════
 * HIPERATIVO V3 — Cadastro.gs  (v2.1 — envio email/WhatsApp 04/06/2026)
 * Cadastro de atletas, geração e envio de links
 * ═══════════════════════════════════════════════════════════════════════
 */

function limparErrosFormula() {
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const sheets  = ss.getSheets();
  const erros   = ['#ERROR!', '#N/A', '#REF!', '#VALUE!', '#NAME?', '#DIV/0!', '#NUM!', '#NULL!'];
  let   total   = 0;

  for (const sheet of sheets) {
    const range  = sheet.getDataRange();
    const values = range.getValues();
    const rows   = values.length;
    const cols   = values[0] ? values[0].length : 0;
    const batch  = [];

    for (let r = 0; r < rows; r++) {
      for (let c = 0; c < cols; c++) {
        const val = String(values[r][c] || '');
        if (erros.some(e => val.startsWith(e))) {
          batch.push({ r: r + 1, c: c + 1 });
        }
      }
    }

    for (const cell of batch) {
      sheet.getRange(cell.r, cell.c).clearContent();
      total++;
    }
  }

  try {
    SpreadsheetApp.getUi().alert(
      '✅ Limpeza concluída!',
      total + ' células com erro foram limpas em todas as abas.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (_) {}

  _log('SISTEMA', 'INFO', 'limparErrosFormula', 'Total de células limpas: ' + total, '');
  return total;
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
  const linkCadastro = url + '?cadastro=true' + (athId ? '&athId=' + encodeURIComponent(athId) : '');

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
    // Auto-generate: ATH + timestamp
    athId = 'ATH' + String(Date.now()).slice(-6);
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
      body: 'Olá, ' + primeiroNome + '!\n\nPreencha seu cadastro HIPERATIVO neste link:\n' + link,
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
  
  const athId = 'ATH' + String(Date.now()).slice(-6);
  const link = webAppUrl + '?cadastro=true&athId=' + encodeURIComponent(athId) + '&ref=whatsapp';
  
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
  const selecionado = _getAtletaLinhaSelecionada();
  const rId = ui.prompt('📡 Enviar Link Strava',
    'Digite o nome, e-mail ou ID do atleta:' +
      (selecionado ? '\n\nLinha selecionada: ' + selecionado.nome + '\nDeixe em branco para usar essa linha.' : ''),
    ui.ButtonSet.OK_CANCEL);
  if (rId.getSelectedButton() !== ui.Button.OK) return;
  let atleta;
  try {
    atleta = _resolverAtleta(rId.getResponseText(), selecionado);
  } catch (e) {
    ui.alert('❌ Atleta não localizado', e.message, ui.ButtonSet.OK);
    return;
  }
  const athId = atleta.athId;
  const nome = atleta.nome;
  const emailAtleta = atleta.email;
  if (!emailAtleta || !emailAtleta.includes('@')) {
    ui.alert('❌ Atleta ' + nome + ' não tem e-mail cadastrado.');
    return;
  }
  const props = PropertiesService.getScriptProperties();
  let oauthUrl;
  try {
    oauthUrl = _gerarUrlConexaoStrava(athId);
  } catch (e) {
    _log(athId, 'ERRO', 'enviarLinkStravaEmail', e.message, e.stack || '');
    ui.alert('❌ Não foi possível gerar a autorização', e.message, ui.ButtonSet.OK);
    return;
  }
  
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
      body: 'Olá, ' + primeiroNome + '!\n\nConecte sua conta Strava ao HIPERATIVO neste link:\n' + oauthUrl +
        '\n\nApós autorizar, suas atividades poderão ser importadas automaticamente.',
      htmlBody: corpoHtml
    });
    
    ui.alert('✅ Link Strava enviado!\n\n📧 Para: ' + emailAtleta + '\n👤 Atleta: ' + nome);
    _log(athId, 'INFO', 'enviarLinkStravaEmail', 'Autorização enviada para ' + nome + ' <' + emailAtleta + '>', '');
  } catch(e) {
    _log(athId, 'ERRO', 'enviarLinkStravaEmail', e.message, e.stack || '');
    ui.alert('❌ Erro ao enviar: ' + e.message);
  }
}

// ══════════════════════════════════════════════════════
// RESTAURAR ESTRUTURA (seguro — não apaga dados)
// ══════════════════════════════════════════════════════
function enviarLinksStravaPendentesEmail() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();
  const webAppUrl = props.getProperty('WEBAPP_URL');
  if (!webAppUrl) {
    ui.alert('❌ WEBAPP_URL não configurada. Configure o Web App antes de enviar links.');
    return;
  }

  const conf = ui.alert(
    '📤 Enviar links Strava pendentes',
    'Isso vai enviar e-mail para todos os atletas cadastrados que ainda não têm refresh_token na aba TOKENS.\n\nContinuar?',
    ui.ButtonSet.YES_NO
  );
  if (conf !== ui.Button.YES) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shCad = ss.getSheetByName(H.SHEETS.CADASTRO);
  const shTok = ss.getSheetByName(H.SHEETS.TOKENS);
  if (!shCad) {
    ui.alert('❌ Aba CADASTRO não encontrada.');
    return;
  }

  const cad = shCad.getDataRange().getValues();
  const tok = shTok ? shTok.getDataRange().getValues() : [];
  const conectados = {};
  for (let i = 2; i < tok.length; i++) {
    const athId = String(tok[i][H.TOK.ATH_ID - 1] || '').trim().toUpperCase();
    const refresh = String(tok[i][H.TOK.REFRESH - 1] || '').trim();
    if (athId && refresh) conectados[athId] = true;
  }

  const adminEmail = props.getProperty('ADMIN_EMAIL') || 'contato@ghiperativo.com.br';
  const relatorio = [['Nome', 'E-mail', 'Status', 'Link Strava', 'Observação', 'ID Atleta']];
  let enviados = 0;
  let jaConectados = 0;
  let semEmail = 0;
  let erros = 0;

  for (let i = 2; i < cad.length; i++) {
    const athId = String(cad[i][H.CAD.ID - 1] || '').trim().toUpperCase();
    if (!athId) continue;

    const nome = String(cad[i][H.CAD.NOME - 1] || athId).trim();
    const email = String(cad[i][H.CAD.EMAIL - 1] || '').trim();
    if (conectados[athId]) {
      jaConectados++;
      relatorio.push([nome, email, 'Já conectado', '', 'Refresh token encontrado', athId]);
      continue;
    }

    let link = '';
    try {
      link = _gerarUrlConexaoStrava(athId);
    } catch (e) {
      erros++;
      relatorio.push([nome, email, 'Erro ao gerar link', '', e.message, athId]);
      _log(athId, 'ERRO', 'enviarLinksStravaPendentesEmail', e.message, e.stack || '');
      continue;
    }

    if (!email || !email.includes('@')) {
      semEmail++;
      relatorio.push([nome, email, 'Não enviado', link, 'Sem e-mail válido; envie o link manualmente', athId]);
      continue;
    }

    try {
      const primeiroNome = nome.split(' ')[0] || 'Atleta';
      MailApp.sendEmail({
        to: email,
        replyTo: adminEmail,
        subject: '⚡ ' + primeiroNome + ', conecte seu Strava ao HIPERATIVO!',
        body: 'Olá, ' + primeiroNome + '!\n\nConecte sua conta Strava ao HIPERATIVO neste link:\n' + link +
          '\n\nApós autorizar, suas atividades poderão ser importadas automaticamente.',
        htmlBody: _htmlEmailLinkStrava(primeiroNome, link)
      });
      enviados++;
      shCad.getRange(i + 1, H.CAD.STRAVA_OK).setValue('Pendente');
      relatorio.push([nome, email, 'Enviado', link, 'Aguardando autorização do atleta', athId]);
      _log(athId, 'INFO', 'enviarLinksStravaPendentesEmail', 'Link Strava enviado para ' + email, '');
    } catch (e) {
      erros++;
      relatorio.push([nome, email, 'Erro', link, e.message, athId]);
      _log(athId, 'ERRO', 'enviarLinksStravaPendentesEmail', e.message, e.stack || '');
    }
  }

  let shEnvios = ss.getSheetByName('📨 STRAVA ENVIOS');
  if (!shEnvios) shEnvios = ss.insertSheet('📨 STRAVA ENVIOS');
  shEnvios.clearContents();
  shEnvios.getRange(1, 1, relatorio.length, relatorio[0].length).setValues(relatorio);
  shEnvios.getRange(1, 1, 1, relatorio[0].length).setFontWeight('bold');
  shEnvios.autoResizeColumns(1, relatorio[0].length);
  ss.setActiveSheet(shEnvios);

  ui.alert(
    '📤 Envio de links Strava concluído',
    'Enviados: ' + enviados +
      '\nJá conectados: ' + jaConectados +
      '\nSem e-mail: ' + semEmail +
      '\nErros: ' + erros +
      '\n\nA aba 📨 STRAVA ENVIOS foi atualizada com os links.',
    ui.ButtonSet.OK
  );
}

function _htmlEmailLinkStrava(primeiroNome, linkStrava) {
  return `<!DOCTYPE html>
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
  <p style="color:#333;line-height:1.6">Para importarmos seus treinos automaticamente, precisamos que você conecte sua conta <strong>Strava</strong> ao sistema HIPERATIVO.</p>
  <div style="text-align:center;margin:24px 0">
    <a href="${linkStrava}" style="display:inline-block;background:#FC4C02;color:#fff;text-decoration:none;padding:14px 32px;border-radius:24px;font-weight:700;font-size:16px">
      🔗 CONECTAR MINHA CONTA STRAVA
    </a>
  </div>
  <p style="color:#666;font-size:12px">Ao clicar, você será redirecionado para o Strava para autorizar apenas a leitura das suas atividades.</p>
</td></tr>
<tr><td style="background:#0d2560;padding:16px;text-align:center">
  <div style="color:#fff;font-size:10px;opacity:.7">⚡ HIPERATIVO © 2026</div>
</td></tr>
</table>
</body></html>`;
}

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
function corrigirFormulasPainel() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  try {
    var painel = ss.getSheetByName(H.SHEETS.PAINEL);
    if (!painel) throw new Error("Aba PAINEL nao encontrada");
    var c = H.SHEETS.CADASTRO;
    var a = H.SHEETS.ATIVIDADES;
    var q = "\"";
    // Row 6: metricas (STATUS=col26=Z, STRAVA_OK=col24=X)
    painel.getRange(6,1).setFormulaLocal("=CONT.VALORES('" + c + "'!A3:A500)");
    painel.getRange(6,3).setFormulaLocal("=CONT.SE('" + c + "'!Z3:Z500;"+q+"Ativo"+q+")");
    painel.getRange(6,5).setFormulaLocal("=CONT.SE('" + c + "'!X3:X500;"+q+"Sim"+q+")");
    painel.getRange(6,7).setFormulaLocal("=CONT.SE('" + a + "'!D3:D500;"+q+">"+q+"&HOJE()-7)");
    painel.getRange(6,9).setFormulaLocal("=SEERRO(TEXTO(MÉDIASE('" + a + "'!E3:E500;"+q+"Corrida"+q+";'" + a + "'!N3:N500)/86400;"+q+"[mm]:ss"+q+");"+q+"--"+q+")");
    painel.getRange(6,11).setFormulaLocal("=SEERRO(ARRED(SOMASE('" + a + "'!E3:E500;"+q+"Corrida"+q+";'" + a + "'!L3:L500)/MAX(1;CONT.SE('" + a + "'!E3:E500;"+q+"Corrida"+q+"));1);"+q+"--"+q+")");
    // Rows 11-20: ultimas atividades e alertas
    for (var i = 0; i < 10; i++) {
      var r = 11 + i;
      var rk = i + 1;
      painel.getRange(r,1).setFormulaLocal("=SEERRO(TEXTO(MAIOR('" + a + "'!D3:D500;"+rk+");"+q+"dd/mm/yyyy"+q+");"+q+q+")");
      painel.getRange(r,2).setFormulaLocal("=SEERRO(ÍNDICE('" + a + "'!C3:C500;CORRESP(MAIOR('" + a + "'!D3:D500;"+rk+");'" + a + "'!D3:D500;0));"+q+q+")");
      painel.getRange(r,3).setFormulaLocal("=SEERRO(ÍNDICE('" + a + "'!E3:E500;CORRESP(MAIOR('" + a + "'!D3:D500;"+rk+");'" + a + "'!D3:D500;0));"+q+q+")");
      painel.getRange(r,4).setFormulaLocal("=SEERRO(ÍNDICE('" + a + "'!L3:L500;CORRESP(MAIOR('" + a + "'!D3:D500;"+rk+");'" + a + "'!D3:D500;0));"+q+q+")");
      painel.getRange(r,5).setFormulaLocal("=SEERRO(ÍNDICE('" + a + "'!O3:O500;CORRESP(MAIOR('" + a + "'!D3:D500;"+rk+");'" + a + "'!D3:D500;0));"+q+q+")");
      painel.getRange(r,6).setFormulaLocal("=SEERRO(ÍNDICE('" + a + "'!Q3:Q500;CORRESP(MAIOR('" + a + "'!D3:D500;"+rk+");'" + a + "'!D3:D500;0));"+q+q+")");
    }
    _log("SYSTEM","INFO","corrigirFormulasPainel","Formulas do PAINEL corrigidas","");
    ui.alert("Fórmulas corrigidas!","As métricas do PAINEL foram atualizadas.",ui.ButtonSet.OK);
  } catch(e) {
    _log("SYSTEM","ERRO","corrigirFormulasPainel",e.message,"");
    ui.alert("Erro", e.message, ui.ButtonSet.OK);
  }
}
