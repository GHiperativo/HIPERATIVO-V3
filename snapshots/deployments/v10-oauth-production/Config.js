/**
 * ═══════════════════════════════════════════════════════════════════════
 * HIPERATIVO V3 — Config.gs  (v4.0 — menu expandido 04/06/2026)
 * Constantes globais, menu, log, trigger e configurações
 * ═══════════════════════════════════════════════════════════════════════
 */

// ── CONSTANTES GLOBAIS ───────────────────────────────────────────────────────
const H = {
  SHEETS: {
    PAINEL:    '📊 PAINEL',
    CADASTRO:  '👤 CADASTRO',
    ATIVIDADES:'🏃 ATIVIDADES',
    PLANO:     '📅 PLANO SEMANAL',
    METRICAS:  '📈 MÉTRICAS',
    FEEDBACK:  '💬 FEEDBACK',
    GRAFICOS:  '📉 GRÁFICOS',
    ERROS:     '🔴 ERROS',
    TOKENS:    '🔐 TOKENS',
    CONFIG:    '⚙️ CONFIG',
  },

  // ── Colunas da aba CADASTRO (1-indexed) — v3.2 expandido ──────────────────
  // Col:  1     2      3       4       5      6      7      8        9
  // Col:  ID   NOME  EMAIL  WHATS   NASC   SEXO   PESO  ALTURA   MOD
  // Col: 10     11     12      13      14     15     16     17     18
  // Col: NIVEL   OBJ  FREQ  HORARIO  SAUDE  LESAO  MED   PROVA  PLANO
  // Col: 19     20      21     22        23      24     25     26
  // Col: CIDADE ESTADO  CPF  ORIGEM  DATA_CAD STRAVA_OK STRAVA_ID STATUS OBS
  CAD: {
    ID:        1,
    NOME:      2,
    EMAIL:     3,
    WHATS:     4,
    NASC:      5,
    SEXO:      6,
    PESO:      7,
    ALTURA:    8,
    MOD:       9,
    NIVEL:    10,
    OBJ:      11,
    FREQ:     12,
    HORARIO:  13,
    SAUDE:    14,
    LESAO:    15,
    MED:      16,
    PROVA:    17,
    PLANO:    18,
    CIDADE:   19,
    ESTADO:   20,
    CPF:      21,
    ORIGEM:   22,
    DATA_CAD: 23,
    STRAVA_OK:24,
    STRAVA_ID:25,
    STATUS:   26,
    OBS:      27,
  },

  // ── Colunas da aba ATIVIDADES (1-indexed) ─────────────────────────────────
  ATIV: {
    EXEC_ID:  1,
    ATH_ID:   2,
    NOME:     3,
    DATA:     4,
    TIPO:     5,
    FONTE:    6,
    STRAVA_ID:7,
    NOME_ATIV:8,
    MOV_S:    9,
    TOTAL_S: 10,
    DIST_M:  11,
    DIST_KM: 12,
    VEL_MPS: 13,
    VEL_KMH: 14,
    PACE_S:  15,
    PACE_MS: 16,
    FC_MED:  17,
    FC_MAX:  18,
    ELEV:    19,
    CAL:     20,
    CADENCIA:21,
    POTENCIA:22,
    ROTA:    23,
    IMPORTADO:24,
  },

  PLANO: {
    ID:        1,
    ATH_ID:    2,
    NOME:      3,
    SEMANA:    4,
    CICLO:     5,
    MOD:       6,
    TITULO:    7,
    DIAG:      8,
    D1_TIPO:   9,
    D1_PRESC: 10,
    D2_TIPO:  11,
    D2_PRESC: 12,
    D3_TIPO:  13,
    D3_PRESC: 14,
    INTENCAO: 15,
    STATUS:   16,
  },

  // ── Colunas da aba TOKENS (1-indexed) ─────────────────────────────────────
  TOK: {
    EXEC_ID:  1,
    ATH_ID:   2,
    NOME:     3,
    ACCESS:   4,
    REFRESH:  5,
    EXPIRES:  6,
    SCOPE:    7,
    STRAVA_ID:8,
    ULT_ATU:  9,
    STATUS:  10,
  },

  MET: {
    ATH_ID:     1,
    NOME:       2,
    ATUALIZADO: 3,
    VO2:        4,
    PACE_MED:   5,
    PACE_RAP:   6,
    PACE_LEN:   7,
    FC_MAX:     8,
    FC_MED:     9,
    VOL_SEM:   10,
  },

  // ── Colunas da aba ERROS (1-indexed) ──────────────────────────────────────
  ERR: {
    DATA:    1,
    NIVEL:   2,
    ATH_ID:  3,
    FUNCAO:  4,
    MSG:     5,
    STACK:   6,
    RESOLVIDO: 7,
    ACAO:    8,
  },
};

// ── MENU PRINCIPAL ───────────────────────────────────────────────────────────

// ── MENU PRINCIPAL ───────────────────────────────────────────────────────────
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu('⚡ HIPERATIVO');

  // ─── ATLETAS ─────────────────────────────────────────────────────────────
  menu.addSubMenu(
    ui.createMenu('👤 Atletas')
      .addItem('➕ Cadastrar novo atleta (link por email)', 'gerarLinkCadastroEmail')
      .addItem('🔗 Gerar link de cadastro (copiar)', 'gerarLinkCadastro')
      .addItem('📋 Ver todos os atletas', 'abrirCadastro')
      .addSeparator()
      .addItem('🔄 Conectar atleta ao Strava', 'gerarLinkStrava')
      .addItem('📡 Importar perfil do atleta', 'importarPerfilAtleta')
      .addSeparator()
      .addItem('🗑️ Remover atleta (limpar linha)', 'removerAtleta')
  );

  // ─── ATIVIDADES ──────────────────────────────────────────────────────────
  menu.addSubMenu(
    ui.createMenu('🏃 Atividades')
      .addItem('⬇️ Importar atividades Strava (todos)', 'importarAtividadesTodosStrava')
      .addItem('⬇️ Importar atividades (atleta específico)', 'importarAtividadesAtleta')
      .addItem('👤 Sincronizar nomes nas atividades', 'sincronizarNomesAtividades')
      .addSeparator()
      .addItem('📡 Auditar conexões Strava', 'auditarConexoesStrava')
      .addItem('🩺 Diagnosticar integração Strava', 'diagnosticarIntegracaoStrava')
      .addSeparator()
      .addItem('📊 Atualizar painel geral', 'atualizarPainel')
      .addItem('🧹 Limpar erros de fórmula', 'limparErrosFormula')
  );

  // ─── COMUNICAÇÃO ─────────────────────────────────────────────────────────
  menu.addSubMenu(
    ui.createMenu('📧 Comunicação')
      .addItem('✉️ Enviar link de cadastro (email)', 'gerarLinkCadastroEmail')
      .addItem('📲 Gerar link WhatsApp de cadastro', 'gerarLinkCadastroWhatsapp')
      .addItem('📋 Copiar link de cadastro', 'gerarLinkCadastro')
      .addSeparator()
      .addItem('📤 Enviar link Strava por email', 'enviarLinkStravaEmail')
      .addItem('📤 Enviar links Strava pendentes', 'enviarLinksStravaPendentesEmail')
      .addItem('🧪 Gerar link teste Strava', 'gerarLinkTesteStrava')
  );

  // ─── RELATÓRIOS E EXPORTAÇÃO ─────────────────────────────────────────────
  menu.addSubMenu(
    ui.createMenu('📊 Relatórios')
      .addItem('📈 Relatório geral (resumo)', 'gerarRelatorioGeral')
      .addItem('🏅 Ranking de atletas por volume', 'gerarRankingAtletas')
      .addItem('📧 Enviar relatório por email', 'enviarRelatorioEmail')
      .addSeparator()
      .addItem('📥 Exportar CADASTRO para CSV', 'exportarCadastroCSV')
      .addItem('🔗 Sincronizar para planilha externa', 'sincronizarPlanilhaExterna')
  );

  // ─── INTEGRAÇÕES ─────────────────────────────────────────────────────────
  menu.addSubMenu(
    ui.createMenu('🔌 Integrações')
      .addItem('📋 Criar formulário Google Forms', 'criarFormularioGoogleForms')
      .addItem('🔗 Gerar WebApp de cadastro (link)', 'mostrarUrlWebApp')
      .addSeparator()
      .addItem('⚡ Configurar webhook Strava', 'configurarWebhookStrava')
      .addItem('📡 Consultar webhook Strava', 'consultarWebhookStrava')
      .addItem('📥 Processar fila Strava agora', 'processarFilaWebhookStravaAgora')
      .addItem('🗑️ Remover webhook Strava', 'removerWebhookStrava')
      .addSeparator()
      .addItem('📊 Vincular planilha de destino', 'vincularPlanilhaDestino')
      .addItem('📤 Push de dados para planilha externa', 'pushDadosPlanilhaExterna')
  );

  // ─── CONFIGURAÇÕES E SISTEMA ─────────────────────────────────────────────
  menu.addSubMenu(
    ui.createMenu('⚙️ Configurações')
      .addItem('🔧 Configurar credenciais Strava', 'configurarCredenciais')
    .addItem('⚡ Configuração Rápida / Status', 'configuracaoRapida')
      .addItem('📧 Configurar email admin', 'configurarEmailAdmin')
      .addSeparator()
      .addItem('🕐 Configurar automação Strava', 'configurarAutomacaoStrava')
      .addItem('🗑️ Remover todos os triggers', 'removerTriggers')
      .addSeparator()
      .addItem('🔴 Ver log de erros', 'abrirErros')
      .addItem('🛠️ Setup inicial da planilha', 'setupInicial')
    .addItem('🔧 Restaurar estrutura (seguro)', 'restaurarEstrutura')
      .addItem('🧰 Corrigir estrutura leve', 'corrigirEstruturaLeve')
      .addItem('♻️ Reinstalar estrutura', 'setupPlanilha')
  );

  menu.addToUi();
}

// ── CONFIGURAÇÕES ──────────────────────────────────────────────────────────

function configurarCredenciais() {
  const ui    = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();

  const r1 = ui.prompt('⚙️ Configurar — Passo 1/3', 'STRAVA_CLIENT_ID (apenas números):', ui.ButtonSet.OK_CANCEL);
  if (r1.getSelectedButton() !== ui.Button.OK) return;
  if (r1.getResponseText().trim()) props.setProperty('STRAVA_CLIENT_ID', r1.getResponseText().trim());

  const r2 = ui.prompt('⚙️ Configurar — Passo 2/3', 'STRAVA_CLIENT_SECRET:', ui.ButtonSet.OK_CANCEL);
  if (r2.getSelectedButton() !== ui.Button.OK) return;
  if (r2.getResponseText().trim()) props.setProperty('STRAVA_CLIENT_SECRET', r2.getResponseText().trim());

  const r3 = ui.prompt('⚙️ Configurar — Passo 3/3', 'WEBAPP_URL (URL /exec da implantação):', ui.ButtonSet.OK_CANCEL);
  if (r3.getSelectedButton() !== ui.Button.OK) return;
  if (r3.getResponseText().trim()) props.setProperty('WEBAPP_URL', r3.getResponseText().trim());

  ui.alert('✅ Credenciais salvas!', 'As configurações foram armazenadas com segurança nas propriedades do projeto.', ui.ButtonSet.OK);
}

function configurarEmailAdmin() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();
  const atual = props.getProperty('ADMIN_EMAIL') || '';
  const r = ui.prompt('📧 Email Admin', 
    'Email para receber relatórios e notificações:\n(atual: ' + (atual || 'não configurado') + ')',
    ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const email = r.getResponseText().trim();
  if (!email || !email.includes('@')) { ui.alert('❌ Email inválido.'); return; }
  props.setProperty('ADMIN_EMAIL', email);
  ui.alert('✅ Email admin configurado: ' + email);
}

function criarTrigger() {
  configurarAutomacaoStrava();
}

function removerTriggers() {
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  try {
    SpreadsheetApp.getUi().alert('🗑️ Triggers removidos', 'Todos os triggers foram excluídos.', SpreadsheetApp.getUi().ButtonSet.OK);
  } catch(_) {}
}

// ── NAVEGAÇÃO ────────────────────────────────────────────────────────────────

function abrirCadastro() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(H.SHEETS.CADASTRO);
  if (sh) ss.setActiveSheet(sh);
}

function mostrarUrlWebApp() {
  const url = PropertiesService.getScriptProperties().getProperty('WEBAPP_URL');
  if (!url) {
    SpreadsheetApp.getUi().alert('❌ URL do WebApp não configurada.\nVá em ⚙️ Configurações → Configurar credenciais.');
    return;
  }
  SpreadsheetApp.getUi().alert('🔗 URL do WebApp de Cadastro:\n\n' + url + '\n\nCopie esta URL para compartilhar o formulário de cadastro.');
}

// ── ATLETAS HELPERS ───────────────────────────────────────────────────────────

function removerAtleta() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt('🗑️ Remover Atleta', 
    'Digite o ATH_ID do atleta a remover (ex: ATH001):',
    ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const athId = r.getResponseText().trim().toUpperCase();
  if (!athId) return;
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(H.SHEETS.CADASTRO);
  const dados = sh.getDataRange().getValues();
  
  for (let i = 1; i < dados.length; i++) {
    if (String(dados[i][H.CAD.ID - 1]).toUpperCase() === athId) {
      const conf = ui.alert('⚠️ Confirmar remoção?', 
        'Atleta: ' + dados[i][H.CAD.NOME - 1] + '\nID: ' + athId + '\n\nIsso apagará todos os dados desta linha.',
        ui.ButtonSet.YES_NO);
      if (conf === ui.Button.YES) {
        sh.deleteRow(i + 1);
        ui.alert('✅ Atleta ' + athId + ' removido com sucesso.');
      }
      return;
    }
  }
  ui.alert('❌ Atleta ' + athId + ' não encontrado.');
}

function importarAtividadesAtleta() {
  const ui = SpreadsheetApp.getUi();
  const selecionado = _getAtletaLinhaSelecionada();
  const r = ui.prompt('⬇️ Importar Atividades',
    'Digite o nome, e-mail ou ID do atleta:' +
      (selecionado ? '\n\nLinha selecionada: ' + selecionado.nome : ''),
    ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  try {
    const atleta = _resolverAtleta(r.getResponseText(), selecionado);
    const n = _importarAtividadesAtleta(atleta.athId, 3);
    ui.alert('✅ ' + n + ' atividades de ' + atleta.nome + ' importadas com sucesso!');
  } catch(e) {
    _log('SYSTEM', 'ERRO', 'importarAtividadesAtleta', e.message, e.stack || '');
    ui.alert('❌ Erro: ' + e.message);
  }
}

// ── RELATÓRIOS ────────────────────────────────────────────────────────────────

function auditarConexoesStrava() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shCad = ss.getSheetByName(H.SHEETS.CADASTRO);
  const shTok = ss.getSheetByName(H.SHEETS.TOKENS);
  if (!shCad) {
    SpreadsheetApp.getUi().alert('❌ Aba CADASTRO não encontrada.');
    return;
  }

  const cad = shCad.getDataRange().getValues();
  const tok = shTok ? shTok.getDataRange().getValues() : [];
  const tokensPorAtleta = {};

  for (let i = 2; i < tok.length; i++) {
    const athId = String(tok[i][H.TOK.ATH_ID - 1] || '').trim().toUpperCase();
    if (!athId) continue;
    tokensPorAtleta[athId] = {
      access: String(tok[i][H.TOK.ACCESS - 1] || '').trim(),
      refresh: String(tok[i][H.TOK.REFRESH - 1] || '').trim(),
      expires: tok[i][H.TOK.EXPIRES - 1] || '',
      stravaId: tok[i][H.TOK.STRAVA_ID - 1] || '',
      atualizado: tok[i][H.TOK.ULT_ATU - 1] || '',
      status: String(tok[i][H.TOK.STATUS - 1] || '').trim(),
    };
  }

  const statusRows = [[
    'Nome', 'E-mail', 'Conexão Strava', 'Status Cadastro', 'ID Atleta',
    'ID Strava', 'Refresh Token', 'Expira Em', 'Última Atualização', 'Observação'
  ]];

  let conectados = 0;
  let naoConectados = 0;
  let pendentes = 0;

  for (let i = 2; i < cad.length; i++) {
    const athId = String(cad[i][H.CAD.ID - 1] || '').trim().toUpperCase();
    if (!athId) continue;
    const nome = cad[i][H.CAD.NOME - 1] || '';
    const email = cad[i][H.CAD.EMAIL - 1] || '';
    const statusCad = cad[i][H.CAD.STATUS - 1] || '';
    const token = tokensPorAtleta[athId];
    const temRefresh = !!(token && token.refresh);
    const conexao = temRefresh ? 'Conectado' : 'Não conectado';
    const obs = temRefresh ? 'OK para importar' : 'Enviar link de conexão Strava';

    if (temRefresh) {
      conectados++;
      shCad.getRange(i + 1, H.CAD.STRAVA_OK).setValue('Sim');
      if (token.stravaId) shCad.getRange(i + 1, H.CAD.STRAVA_ID).setValue(token.stravaId);
    } else {
      naoConectados++;
      const atual = String(cad[i][H.CAD.STRAVA_OK - 1] || '').trim().toLowerCase();
      if (atual === 'pendente') pendentes++;
      shCad.getRange(i + 1, H.CAD.STRAVA_OK).setValue(atual === 'pendente' ? 'Pendente' : 'Não');
    }

    statusRows.push([
      nome, email, conexao, statusCad, athId,
      token ? token.stravaId : '', temRefresh ? 'Sim' : 'Não',
      token ? token.expires : '', token ? token.atualizado : '', obs
    ]);
  }

  let shStatus = ss.getSheetByName('📡 STRAVA STATUS');
  if (!shStatus) shStatus = ss.insertSheet('📡 STRAVA STATUS');
  shStatus.clearContents();
  shStatus.getRange(1, 1, statusRows.length, statusRows[0].length).setValues(statusRows);
  shStatus.getRange(1, 1, 1, statusRows[0].length).setFontWeight('bold');
  shStatus.autoResizeColumns(1, statusRows[0].length);
  ss.setActiveSheet(shStatus);

  _log('SYSTEM', 'INFO', 'auditarConexoesStrava',
    'Conectados: ' + conectados + ' | Não conectados: ' + naoConectados + ' | Pendentes: ' + pendentes, '');
  SpreadsheetApp.getUi().alert(
    '📡 Auditoria Strava concluída',
    'Conectados: ' + conectados + '\nNão conectados: ' + naoConectados + '\nPendentes: ' + pendentes +
      '\n\nA aba 📡 STRAVA STATUS foi atualizada.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function gerarRelatorioGeral() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shCad = ss.getSheetByName(H.SHEETS.CADASTRO);
  const shAtv = ss.getSheetByName(H.SHEETS.ATIVIDADES);
  
  const cadDados = shCad ? shCad.getDataRange().getValues() : [];
  const atvDados = shAtv ? shAtv.getDataRange().getValues() : [];
  
  const totalAtletas = cadDados.length > 1 ? cadDados.length - 1 : 0;
  const atletasStrava = cadDados.slice(1).filter(r => {
    const v = String(r[H.CAD.STRAVA_OK - 1] || '').trim().toLowerCase();
    return v === 'sim' || v === 'true';
  }).length;
  const totalAtividades = atvDados.length > 1 ? atvDados.length - 1 : 0;
  
  const msg = [
    '📊 RELATÓRIO GERAL — HIPERATIVO V3',
    '════════════════════════════════════',
    '👥 Total de atletas: ' + totalAtletas,
    '🔗 Conectados ao Strava: ' + atletasStrava,
    '🏃 Total de atividades: ' + totalAtividades,
    '',
    'Data: ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm'),
  ].join('\n');
  
  SpreadsheetApp.getUi().alert('📊 Relatório Geral', msg, SpreadsheetApp.getUi().ButtonSet.OK);
}

function gerarRankingAtletas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shAtv = ss.getSheetByName(H.SHEETS.ATIVIDADES);
  if (!shAtv) { SpreadsheetApp.getUi().alert('❌ Aba ATIVIDADES não encontrada.'); return; }
  
  const dados = shAtv.getDataRange().getValues();
  if (dados.length < 2) { SpreadsheetApp.getUi().alert('Sem atividades registradas.'); return; }
  
  // Count activities per athlete (assuming col 1 = ATH_ID, col 5 = distance or similar)
  const ranking = {};
  dados.slice(1).forEach(row => {
    const id = String(row[0]);
    if (!ranking[id]) ranking[id] = {count: 0, nome: String(row[1] || id)};
    ranking[id].count++;
  });
  
  const sorted = Object.entries(ranking).sort((a,b) => b[1].count - a[1].count).slice(0,10);
  const lines = ['🏅 RANKING — TOP 10 ATLETAS (por nº de atividades)', '════════════════════════════════════════════════'];
  sorted.forEach(([id, data], idx) => {
    lines.push((idx+1) + 'º ' + data.nome + ' (' + id + '): ' + data.count + ' atividades');
  });
  
  SpreadsheetApp.getUi().alert('🏅 Ranking', lines.join('\n'), SpreadsheetApp.getUi().ButtonSet.OK);
}

function enviarRelatorioEmail() {
  const props = PropertiesService.getScriptProperties();
  const email = props.getProperty('ADMIN_EMAIL') || 'contato@ghiperativo.com.br';
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shCad = ss.getSheetByName(H.SHEETS.CADASTRO);
  const shAtv = ss.getSheetByName(H.SHEETS.ATIVIDADES);
  
  const totalAtletas = shCad ? Math.max(0, shCad.getLastRow() - 1) : 0;
  const totalAtividades = shAtv ? Math.max(0, shAtv.getLastRow() - 1) : 0;
  const data = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
  
  const corpo = '<h2>📊 Relatório HIPERATIVO V3 — ' + data + '</h2>' +
    '<table border="1" cellpadding="8" style="border-collapse:collapse">' +
    '<tr><td><b>Total de atletas</b></td><td>' + totalAtletas + '</td></tr>' +
    '<tr><td><b>Total de atividades</b></td><td>' + totalAtividades + '</td></tr>' +
    '</table>' +
    '<p>Acesse a planilha: <a href="' + ss.getUrl() + '">' + ss.getName() + '</a></p>';
  
  MailApp.sendEmail({
    to: email,
    subject: '📊 Relatório HIPERATIVO V3 — ' + data,
    body: 'Relatório HIPERATIVO V3 — ' + data + '\nTotal de atletas: ' + totalAtletas +
      '\nTotal de atividades: ' + totalAtividades + '\nPlanilha: ' + ss.getUrl(),
    htmlBody: corpo
  });
  SpreadsheetApp.getUi().alert('✅ Relatório enviado para: ' + email);
}

// ── EXPORTAÇÃO E INTEGRAÇÃO ───────────────────────────────────────────────────

function exportarCadastroCSV() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(H.SHEETS.CADASTRO);
  if (!sh) { SpreadsheetApp.getUi().alert('❌ Aba CADASTRO não encontrada.'); return; }
  
  const dados = sh.getDataRange().getValues();
  const csv = dados.map(row => row.map(cell => {
    const s = String(cell).replace(/"/g, '""');
    return s.includes(',') || s.includes('"') || s.includes('\n') ? '"' + s + '"' : s;
  }).join(',')).join('\n');
  
  const blob = Utilities.newBlob(csv, 'text/csv', 'CADASTRO_HIPERATIVO_' + 
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd') + '.csv');
  const file = DriveApp.createFile(blob);
  
  SpreadsheetApp.getUi().alert('✅ CSV exportado para o Drive:\n' + file.getName() + '\n\nURL: ' + file.getUrl());
}

function vincularPlanilhaDestino() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt('🔗 Vincular Planilha',
    'Cole o ID da planilha de destino (ex: 1aBcD...xyz):\n(É o código da URL entre /d/ e /edit)',
    ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const id = r.getResponseText().trim();
  if (!id) return;
  
  try {
    const dest = SpreadsheetApp.openById(id);
    PropertiesService.getScriptProperties().setProperty('PLANILHA_DESTINO_ID', id);
    ui.alert('✅ Planilha vinculada: ' + dest.getName() + '\nUse "Push de dados" para sincronizar.');
  } catch(e) {
    ui.alert('❌ Erro ao acessar planilha: ' + e.message + '\nVerifique se você tem acesso a ela.');
  }
}

function sincronizarPlanilhaExterna() {
  pushDadosPlanilhaExterna();
}

function pushDadosPlanilhaExterna() {
  const props = PropertiesService.getScriptProperties();
  const destId = props.getProperty('PLANILHA_DESTINO_ID');
  if (!destId) {
    SpreadsheetApp.getUi().alert('❌ Nenhuma planilha vinculada.\nVá em 🔌 Integrações → Vincular planilha de destino.');
    return;
  }
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dest = SpreadsheetApp.openById(destId);
    
    // Copy CADASTRO sheet
    const shCad = ss.getSheetByName(H.SHEETS.CADASTRO);
    if (shCad) {
      let destCad = dest.getSheetByName('CADASTRO_HIPERATIVO');
      if (!destCad) destCad = dest.insertSheet('CADASTRO_HIPERATIVO');
      else destCad.clearContents();
      
      const dados = shCad.getDataRange().getValues();
      if (dados.length > 0) {
        destCad.getRange(1, 1, dados.length, dados[0].length).setValues(dados);
      }
    }
    
    SpreadsheetApp.getUi().alert('✅ Dados sincronizados para: ' + dest.getName() + '\nAba criada: CADASTRO_HIPERATIVO');
  } catch(e) {
    SpreadsheetApp.getUi().alert('❌ Erro na sincronização: ' + e.message);
  }
}

function criarFormularioGoogleForms() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt('📋 Criar Formulário',
    'Nome do formulário (ex: Cadastro HIPERATIVO - Turma Janeiro):',
    ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const titulo = r.getResponseText().trim() || 'Cadastro HIPERATIVO';
  
  try {
    const form = FormApp.create(titulo);
    form.setTitle(titulo);
    form.setDescription('Formulário de cadastro HIPERATIVO. Preencha todos os campos.');
    
    // Add questions
    form.addTextItem().setTitle('Nome completo').setRequired(true);
    form.addTextItem().setTitle('E-mail').setRequired(true);
    form.addTextItem().setTitle('WhatsApp (com DDD)').setRequired(true);
    form.addDateItem().setTitle('Data de nascimento').setRequired(true);
    form.addMultipleChoiceItem().setTitle('Sexo').setRequired(true)
      .setChoiceValues(['Masculino', 'Feminino', 'Prefiro não informar']);
    form.addMultipleChoiceItem().setTitle('Programa de interesse').setRequired(true)
      .setChoiceValues([
        'Alta Voltagem ⚡',
        'Corrida CCC 🏅',
        'Iniciante em Movimento 🌱',
        '5k/10k em 6 Semanas 🎯',
        'Hiperativo Running Club 🏃',
        'Vida Ativa Melhor Idade 🌟'
      ]);
    form.addTextItem().setTitle('Objetivo principal').setRequired(false);
    form.addParagraphTextItem().setTitle('Observações / condições de saúde').setRequired(false);
    
    const url = form.getPublishedUrl();
    const editUrl = form.getEditUrl();
    
    PropertiesService.getScriptProperties().setProperty('GOOGLE_FORMS_URL', url);
    
    ui.alert('✅ Formulário criado com sucesso!\n\n📋 Link para preencher:\n' + url + 
      '\n\n✏️ Link de edição:\n' + editUrl + '\n\nO link foi salvo nas configurações.');
  } catch(e) {
    ui.alert('❌ Erro ao criar formulário: ' + e.message);
  }
}

// ── LOG E ERROS ───────────────────────────────────────────────────────────────

function _log(athId, nivel, funcao, msg, stack) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(H.SHEETS.ERROS);
    if (!sh) return;
    const ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss');
    sh.appendRow([ts, nivel || 'INFO', athId || '', funcao || '', msg || '', stack || '', 'Não', '']);
  } catch(e) { Logger.log('_log error: ' + e.message); }
}

function abrirErros() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(H.SHEETS.ERROS);
  if (sh) ss.setActiveSheet(sh);
}

function abrirPainel() {
  SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName(H.SHEETS.PAINEL));
  SpreadsheetApp.flush();
}

function getCfg() {
  const props = PropertiesService.getScriptProperties();
  return {
    clientId: props.getProperty('STRAVA_CLIENT_ID') ? '✅ OK' : '❌ Ausente',
    secret:   props.getProperty('STRAVA_CLIENT_SECRET') ? '✅ OK' : '❌ Ausente',
    webApp:   props.getProperty('WEBAPP_URL') ? '✅ OK' : '❌ Ausente',
  };
}

/**
 * Configuração Rápida — configura CLIENT_ID, verifica status e guia o usuário.
 * Acessível via menu ⚡ HIPERATIVO > Configurações > Configuração Rápida
 */
function configuracaoRapida() {
  const ui   = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();

  // 1. Configura CLIENT_ID automaticamente (valor fixo do app Strava)
  props.setProperty('STRAVA_CLIENT_ID', '153043');

  // 2. Verifica o que está faltando
  const secret  = props.getProperty('STRAVA_CLIENT_SECRET') || '';
  const webApp  = props.getProperty('WEBAPP_URL') || '';
  const clientId = props.getProperty('STRAVA_CLIENT_ID') || '153043';

  let status = '✅ CONFIGURADO:\n';
  status += '  • CLIENT_ID: ' + clientId + '\n\n';

  let pendente = '⚠️ PENDENTE (ação manual necessária):\n';
  let acoes = '';

  if (!secret) {
    pendente += '  • CLIENT_SECRET: não configurado\n';
    acoes += '1. Acesse: https://www.strava.com/settings/api\n';
    acoes += '   Copie o "Client Secret" do seu app\n\n';
  } else {
    status += '  • CLIENT_SECRET: configurado ✓\n';
  }

  if (!webApp) {
    pendente += '  • WEBAPP_URL: não configurado\n';
    acoes += '2. Implante o WebApp:\n';
    acoes += '   a) Extensões > Apps Script\n';
    acoes += '   b) Clique "Implantar" > "Nova implantação"\n';
    acoes += '   c) Tipo: "Aplicativo da Web"\n';
    acoes += '   d) Executar como: "EU (sua conta)"\n';
    acoes += '   e) Quem tem acesso: "Qualquer pessoa"\n';
    acoes += '   f) Copie a URL /exec gerada\n\n';
  } else {
    status += '  • WEBAPP_URL: ' + webApp.substring(0, 50) + '...\n';
  }

  if (!secret || !webApp) {
    acoes += '3. Execute: HIPERATIVO > Configurações > Configurar credenciais Strava\n';
    acoes += '   e cole o CLIENT_SECRET e a URL do WebApp quando solicitado.\n';

    const msg = status + '\n' + pendente + '\n\n📋 PRÓXIMAS AÇÕES:\n' + acoes;
    ui.alert('⚡ Configuração Rápida', msg, ui.ButtonSet.OK);
  } else {
    // Tudo configurado! Gerar link de teste
    const linkTeste = webApp + '?cadastro=true';
    ui.alert(
      '✅ Sistema configurado!',
      'Tudo está configurado.\n\n' + status + '\n' +
      '🔗 Link de cadastro (genérico):\n' + linkTeste + '\n\n' +
      'Use o menu "🔗 Gerar link de cadastro" para gerar links por atleta.',
      ui.ButtonSet.OK
    );
  }

  _log('SYSTEM', 'INFO', 'configuracaoRapida', 'Status verificado. Secret:' + (secret ? 'OK' : 'falta') + ' WebApp:' + (webApp ? 'OK' : 'falta'), '');
}


function setCredenciaisStrava() {
  throw new Error('Função desativada por segurança. Configure as credenciais pelo menu Configurar credenciais Strava.');
}


function diagnosticoRapido() {
  var props = PropertiesService.getScriptProperties();
  var clientId = props.getProperty('STRAVA_CLIENT_ID') || 'NAO CONFIGURADO';
  var secret = props.getProperty('STRAVA_CLIENT_SECRET') ? 'CONFIGURADO (' + props.getProperty('STRAVA_CLIENT_SECRET').length + ' chars)' : 'NAO CONFIGURADO';
  var webApp = props.getProperty('WEBAPP_URL') || 'NAO CONFIGURADO';
  Logger.log('CLIENT_ID: ' + clientId);
  Logger.log('CLIENT_SECRET: ' + secret);
  Logger.log('WEBAPP_URL: ' + webApp);
}

function diagnosticarIntegracaoStrava() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();
  const itens = [];
  const erros = [];
  const clientId = props.getProperty('STRAVA_CLIENT_ID') || '';
  const secret = props.getProperty('STRAVA_CLIENT_SECRET') || '';
  const webAppUrl = props.getProperty('WEBAPP_URL') || '';
  const webhookId = props.getProperty('STRAVA_WEBHOOK_SUBSCRIPTION_ID') || '';
  const todasProps = props.getProperties();
  const estadosPendentes = Object.keys(todasProps).filter(k => k.indexOf(STRAVA_OAUTH_STATE_PREFIX) === 0).length;

  clientId ? itens.push('✅ Client ID configurado') : erros.push('❌ STRAVA_CLIENT_ID ausente');
  secret ? itens.push('✅ Client Secret configurado') : erros.push('❌ STRAVA_CLIENT_SECRET ausente');
  try {
    _validarUrlWebApp(webAppUrl);
    itens.push('✅ URL publicada válida: ' + webAppUrl);
  } catch (e) {
    erros.push('❌ ' + e.message);
  }
  webhookId ? itens.push('✅ Webhook registrado: ' + webhookId) : erros.push('⚠️ Webhook ainda não configurado');
  itens.push('ℹ️ Autorizações OAuth aguardando retorno: ' + estadosPendentes);

  const quota = MailApp.getRemainingDailyQuota();
  quota > 0 ? itens.push('✅ Cota de e-mail disponível: ' + quota) : erros.push('❌ Cota diária de e-mail esgotada');

  try {
    const atletas = _listarAtletasCadastro();
    itens.push('✅ Atletas cadastrados: ' + atletas.length);
    const comEmail = atletas.filter(a => a.email && a.email.indexOf('@') > 0).length;
    itens.push('✅ Atletas com e-mail válido: ' + comEmail);
    const shTok = SpreadsheetApp.openById(_getSsId()).getSheetByName(H.SHEETS.TOKENS);
    const tok = shTok ? shTok.getDataRange().getValues() : [];
    const conectados = tok.slice(2).filter(r => String(r[H.TOK.REFRESH - 1] || '').trim()).length;
    itens.push('✅ Atletas com refresh token: ' + conectados);
  } catch (e) {
    erros.push('❌ Não foi possível ler o cadastro: ' + e.message);
  }

  const titulo = erros.length ? '⚠️ Diagnóstico Strava com pendências' : '✅ Integração Strava pronta';
  const mensagem = itens.concat(erros).join('\n') +
    '\n\nNo painel do app Strava, confirme:\n' +
    '• Authorization Callback Domain: script.google.com\n' +
    '• Athlete Capacity suficiente para conectar outros alunos';
  _log('SYSTEM', erros.length ? 'ERRO' : 'INFO', 'diagnosticarIntegracaoStrava', mensagem, '');
  ui.alert(titulo, mensagem, ui.ButtonSet.OK);
}

function gerarLinkTesteStrava() {
  const ui = SpreadsheetApp.getUi();
  const selecionado = _getAtletaLinhaSelecionada();
  const r = ui.prompt('🧪 Link teste Strava',
    'Digite nome, e-mail ou ID do atleta:' +
      (selecionado ? '\n\nLinha selecionada: ' + selecionado.nome + '\nDeixe em branco para usar essa linha.' : ''),
    ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  try {
    const atleta = _resolverAtleta(r.getResponseText(), selecionado);
    const url = _gerarUrlConexaoStrava(atleta.athId);
    ui.alert('Link teste Strava para ' + atleta.nome,
      'Abra este link no navegador do aluno:\n\n' + url +
      '\n\nDepois de autorizar, rode Auditar conexões Strava.',
      ui.ButtonSet.OK);
    _log(atleta.athId, 'INFO', 'gerarLinkTesteStrava', 'Link teste gerado para ' + atleta.nome, '');
  } catch (err) {
    _log('SYSTEM', 'ERRO', 'gerarLinkTesteStrava', err.message, err.stack || '');
    ui.alert('Erro ao gerar link teste', err.message, ui.ButtonSet.OK);
  }
}

function diagnosticarConfiguracao() {
  const props = PropertiesService.getScriptProperties();
    const cId   = props.getProperty('STRAVA_CLIENT_ID')     || '(NAO CONFIGURADO)';
      const cSec  = props.getProperty('STRAVA_CLIENT_SECRET') ? '*** CONFIGURADO ***' : '(NAO CONFIGURADO)';
        const wUrl  = props.getProperty('WEBAPP_URL')           || '(NAO CONFIGURADO)';
          const email = props.getProperty('ADMIN_EMAIL')          || '(NAO CONFIGURADO)';
            const ssId  = props.getProperty('SPREADSHEET_ID')       || '(usa getActiveSpreadsheet)';
              const msg = 'CLIENT_ID: ' + cId + '\n' +
                            'CLIENT_SECRET: ' + cSec + '\n' +
                                          'WEBAPP_URL: ' + wUrl + '\n' +
                                                        'ADMIN_EMAIL: ' + email + '\n' +
                                                                      'SPREADSHEET_ID: ' + ssId;
                                                                        SpreadsheetApp.getUi().alert('Diagnostico', msg, SpreadsheetApp.getUi().ButtonSet.OK);
                                                                        }
