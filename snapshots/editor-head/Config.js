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
    WHATSAPP_STRAVA: '📲 WHATSAPP STRAVA',
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
    STATUS:     26,
    OBS:        27,
    // ── Emergência & Saúde (v4.1 — vindos do formulário web) ─────────────
    EMERG_NOME: 28,
    EMERG_TEL:  29,
    EMERG_REL:  30,
    PAR_Q:      31,  // ex: "N/N/N/N/N/N/N" (7 respostas PAR-Q)
    PR_TEMPOS:  32,  // ex: "5k:25:00|10k:52:00|21k:—|42k:—"
    ASSINATURA: 33,
    DATA_ASS:   34,
    // ── CRM Avançado (preenchimento manual pelo treinador) ─────────────
    INSTAGRAM:   35,
    PROX_PROVA:  36, // ex: "São Silvestre 15k"
    DATA_PROVA:  37,
    PLANO_PAG:   38, // mensal/trimestral/semestral/anual
    DATA_INICIO: 39,
    ULTIMA_AVAL: 40,
  },

  // ── Colunas da aba ATIVIDADES (1-indexed) ─────────────────────────────────
  // Col 13: velocidade m/s (raw Strava — base para todos os cálculos)
  // Col 14: velocidade km/min (corrida/trail) | km/h (ciclismo) | m/min (natação) — armazenado como número
  // Col 15: pace em SEGUNDOS/km (corrida/caminhada) | s/100m (natação) | 0 para ciclismo
  // Col 16: DISPLAY INTELIGENTE por esporte:
  //         Corrida / Caminhada / Trail → "5:30 /km"   (pace min:ss por km)
  //         Ciclismo                   → "28.5 km/h"  (velocidade com 1 decimal)
  //         Natação                    → "1:45 /100m" (pace min:ss por 100m)
  // Col 25: PSE — Percepção Subjetiva de Esforço (escala 1-10, entrada manual)
  ATIV: {
    EXEC_ID:   1,
    ATH_ID:    2,
    NOME:      3,
    DATA:      4,
    TIPO:      5,
    FONTE:     6,
    STRAVA_ID: 7,
    NOME_ATIV: 8,
    MOV_S:     9,
    TOTAL_S:  10,
    DIST_M:   11,
    DIST_KM:  12,
    VEL_MPS:  13,
    VEL_KMMIN:14,  // km/min (era VEL_KMH — atualizado para km/min)
    PACE_S:   15,  // segundos/km (número) — para cálculos de métricas
    PACE_FMT: 16,  // pace formatado "5:30 /km" — para exibição
    FC_MED:   17,
    FC_MAX:   18,
    ELEV:     19,
    CAL:      20,
    CADENCIA: 21,  // spm duplo para corrida (Strava dá 1 pé, multiplicamos por 2)
    POTENCIA: 22,
    ROTA:     23,
    IMPORTADO:24,
    PSE:      25,  // Percepção Subjetiva de Esforço (1-10, entrada manual)
  },

  // ── Colunas da aba MÉTRICAS (1-indexed) ──────────────────────────────────
  MET: {
    ATH_ID:       1,
    NOME:         2,
    ATUALIZADO:   3,
    VO2:          4,
    PACE_MED:     5,
    PACE_RAP:     6,
    PACE_LENTO:   7,
    FC_MAX:       8,
    FC_MED:       9,
    VOL_SEM:     10,
    Z1_LENTO:    11,
    Z1_RAPIDO:   12,
    Z2_LENTO:    13,
    Z2_RAPIDO:   14,
    Z3_LENTO:    15,
    Z3_RAPIDO:   16,
    Z4_LENTO:    17,
    Z5_MIN:      18,
    PERFIL_MAN:  19,
    VOLUME_MAN:  20,
    INTENS_MAN:  21,
    ORIGEM:      22,
    CONFIANCA:   23,
    OBS:         24,
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

  // ── Colunas da aba ERROS (1-indexed) ──────────────────────────────────────
  ERR: {
    DATA:    1,
    ATH_ID:  2,
    NIVEL:   3,
    FUNCAO:  4,
    MSG:     5,
    STACK:   6,
  },
};

// Retorna a opção declarada pelo atleta no CADASTRO (Sim/Não/etc.).
// A ausência de cadastro mantém o comportamento legado; somente "Não" exclui
// explicitamente o atleta das automações Strava.
function _mapaUsoStravaCadastro_() {
  const mapa = {};
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss && ss.getSheetByName(H.SHEETS.CADASTRO);
  if (!sh || sh.getLastRow() < 4) return mapa;

  const dados = sh.getRange(4, 1, sh.getLastRow() - 3, H.CAD.STRAVA_OK).getValues();
  dados.forEach(function(row) {
    const athId = String(row[H.CAD.ID - 1] || '').trim();
    if (!_isAthIdValido_(athId)) return;
    mapa[athId] = String(row[H.CAD.STRAVA_OK - 1] || '').trim()
      .toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '');
  });
  return mapa;
}

function _cadastroNaoUsaStrava_(athId, mapa) {
  const uso = mapa || _mapaUsoStravaCadastro_();
  return String(uso[String(athId || '').trim()] || '') === 'nao';
}

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
      .addItem('🔄 Link geral de conexão com Strava', 'gerarLinkStrava')
      .addItem('📡 Verificar status Strava (todos)', 'verificarStatusStravaAtletas')
      .addItem('🔍 Verificar status REAL Strava (via API)', 'verificarStatusRealStrava')
      .addItem('📤 Enviar link Strava para pendentes', 'enviarLinkStravaDesconectados')
      .addItem('📲 Gerar mensagem WhatsApp (link Strava)', 'gerarMensagemWhatsAppStrava')
      .addItem('📋 Abrir fila WhatsApp dos cadastros', 'abrirFilaWhatsAppStrava')
      .addItem('🔄 Atualizar fila WhatsApp agora', 'sincronizarFilaWhatsAppCadastros')
      .addItem('🔗 Gerar links reconeção (tokens inválidos)', 'gerarLinksReconexaoTodos')
      .addItem('👤 Importar perfil do atleta', 'importarPerfilAtleta')
      .addSeparator()
      .addItem('⏸️ Desativar atleta (preservar histórico)', 'removerAtleta')
  );

  // ─── ATIVIDADES ──────────────────────────────────────────────────────────
  menu.addSubMenu(
    ui.createMenu('🏃 Atividades')
      .addItem('✅ Processar fila Strava agora (seguro)', 'importarUltimas20AtividadesPorAtleta')
      .addItem('⬇️ Iniciar histórico completo (todos)', 'iniciarImportacaoHistoricaCompleta')
      .addSeparator()
      .addItem('📊 Atualizar painel geral', 'atualizarPainel')
      .addItem('🧹 Reparar fórmulas operacionais', 'repararFormulasOperacionais')
      .addSeparator()
      .addItem('📦 Importar lote RAW→CONV (seguro)', 'importarLoteRawConvertidoTodosAtletas_SEGURO')
      .addItem('📈 Gerar Métricas Beta', 'gerarMetricasBeta')
      .addItem('⏰ Ativar automação principal (3h)', 'configurarAutomacaoPrincipal')
      .addItem('📊 Atualizar Ranking Completo', 'atualizarRankingExpandido')
  );

  // ─── COMUNICAÇÃO ─────────────────────────────────────────────────────────
  menu.addSubMenu(
    ui.createMenu('📧 Comunicação')
      .addItem('✉️ Enviar link de cadastro (email)', 'gerarLinkCadastroEmail')
      .addItem('📲 Gerar link WhatsApp de cadastro', 'gerarLinkCadastroWhatsapp')
      .addItem('📋 Copiar link de cadastro', 'gerarLinkCadastro')
      .addSeparator()
      .addItem('📤 Enviar link Strava (email + WhatsApp) para pendentes', 'enviarLinkStravaDesconectados')
      .addItem('📋 Fila WhatsApp pronta para copiar', 'abrirFilaWhatsAppStrava')
      .addItem('⏱️ Ativar fila automática (15 min)', 'instalarAcionadorFilaWhatsApp')
      .addSeparator()
      .addItem('📲 Configurar WhatsApp API (Z-API / Evolution)', 'configurarWhatsApp')
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
      .addItem('🔗 Exibir WebApp de cadastro', 'mostrarUrlWebApp')
      .addSeparator()
      .addItem('📊 Vincular planilha de destino', 'vincularPlanilhaDestino')
      .addItem('📤 Push de dados para planilha externa', 'pushDadosPlanilhaExterna')
  );

  // ─── ANÁLISE E RANKINGS ───────────────────────────────────────────────────
  menu.addSubMenu(
    ui.createMenu('🔬 Análise & Rankings')
      .addItem('🏆 Atualizar Ranking Completo (8 categorias)', 'atualizarRankingExpandido')
      .addItem('🔬 Atualizar Análise Científica (CTL/ATL/TSB)', 'atualizarAnaliseSheet')
      .addSeparator()
      .addItem('📊 Status da Fila Strava', 'statusFila')
  );

  // ─── CONFIGURAÇÕES E SISTEMA ─────────────────────────────────────────────
  menu.addSubMenu(
    ui.createMenu('⚙️ Configurações')
      .addItem('🔧 Configurar credenciais Strava', 'configurarCredenciais')
      .addItem('🗄️ Configurar Supabase (chave)', 'setupSupabase')
        .addItem('🔗 Configurar URL WebApp (exec)', 'configurarUrlWebApp')
      .addItem('⚡ Configuração Rápida / Status', 'configuracaoRapida')
      .addItem('📧 Configurar email admin', 'configurarEmailAdmin')
      .addSeparator()
      .addItem('🕐 Criar trigger automático (3h)', 'configurarTriggers')
      .addSeparator()
      .addItem('🩺 Diagnosticar menu e integrações', 'diagnosticarMenuHiperativo')
      .addItem('🔄 Sincronizar atletas em todas as abas', 'sincronizarAtletasEmTodasAbas')
      .addItem('🔢 Corrigir fórmulas do PAINEL', 'corrigirFormulasPainel')
      .addItem('🔧 Corrigir erros da planilha (linhas/colunas)', 'corrigirErrosDaPlanilha')
      .addSeparator()
      .addItem('🔴 Ver log de erros', 'abrirErros')
      .addItem('🔧 Restaurar estrutura (seguro)', 'restaurarEstrutura')
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
    // Alert removido para compatibilidade com editor GAS
    return;
  }
  SpreadsheetApp.getUi().alert('🔗 URL do WebApp de Cadastro:\n\n' + url + '\n\nCopie esta URL para compartilhar o formulário de cadastro.');
}

// ── ATLETAS HELPERS ───────────────────────────────────────────────────────────

function removerAtleta() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt('⏸️ Desativar Atleta',
    'Digite o ATH_ID do atleta a desativar (ex: ATH001):',
    ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const athId = r.getResponseText().trim().toUpperCase();
  if (!athId) return;
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(H.SHEETS.CADASTRO);
  const dados = sh.getDataRange().getValues();
  
  for (let i = 1; i < dados.length; i++) {
    if (String(dados[i][H.CAD.ID - 1]).toUpperCase() === athId) {
      const conf = ui.alert('⚠️ Confirmar desativação?',
        'Atleta: ' + dados[i][H.CAD.NOME - 1] + '\nID: ' + athId +
        '\n\nO cadastro, os tokens e todo o histórico serão preservados.',
        ui.ButtonSet.YES_NO);
      if (conf === ui.Button.YES) {
        sh.getRange(i + 1, H.CAD.STATUS).setValue('Inativo');
        _log(athId, 'INFO', 'removerAtleta', 'Atleta desativado; histórico e tokens preservados.', '');
        ui.alert('✅ Atleta ' + athId + ' desativado. Nenhum dado foi apagado.');
      }
      return;
    }
  }
  ui.alert('❌ Atleta ' + athId + ' não encontrado.');
}

function importarAtividadesAtleta() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt('⬇️ Importar Atividades',
    'Digite o ATH_ID do atleta (ex: ATH001):',
    ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const athId = r.getResponseText().trim().toUpperCase();
  if (!athId) return;
  try {
    importarPerfilAtleta(athId);
    ui.alert('✅ Atividades de ' + athId + ' importadas com sucesso!');
  } catch(e) {
    ui.alert('❌ Erro: ' + e.message);
  }
}

// ── RELATÓRIOS ────────────────────────────────────────────────────────────────

function gerarRelatorioGeral() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shCad = ss.getSheetByName(H.SHEETS.CADASTRO);
  const shAtv = ss.getSheetByName(H.SHEETS.ATIVIDADES);
  
  const cadDados = shCad ? shCad.getDataRange().getValues() : [];
  const atvDados = shAtv ? shAtv.getDataRange().getValues() : [];
  
  const atletasValidos = cadDados.slice(2).filter(r => _isAthIdValido_(r[H.CAD.ID - 1]));
  const totalAtletas = atletasValidos.length;
  const atletasStrava = atletasValidos.filter(r =>
    String(r[H.CAD.STRAVA_OK - 1] || '').trim().toLowerCase() === 'sim'
  ).length;
  const totalAtividades = atvDados.slice(2).filter(r =>
    _isAthIdValido_(r[H.ATIV.ATH_ID - 1]) && r[H.ATIV.DATA - 1] instanceof Date
  ).length;
  
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
  
  // Colunas operacionais: B=ATH_ID, C=nome. A é EXEC_ID e não identifica atleta.
  const ranking = {};
  dados.slice(2).forEach(row => {
    const id = String(row[H.ATIV.ATH_ID - 1] || '').trim();
    if (!_isAthIdValido_(id)) return;
    if (!ranking[id]) ranking[id] = {count: 0, nome: String(row[H.ATIV.NOME - 1] || id)};
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
  
  MailApp.sendEmail({to: email, subject: '📊 Relatório HIPERATIVO V3 — ' + data, htmlBody: corpo});
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

  const ui = SpreadsheetApp.getUi();
  const conf = ui.alert(
    '📤 Sincronizar cadastro operacional',
    'Serão enviados apenas identificação e dados operacionais. CPF, saúde, emergência, PAR-Q e assinatura não serão copiados.\n\nContinuar?',
    ui.ButtonSet.YES_NO
  );
  if (conf !== ui.Button.YES) return;
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dest = SpreadsheetApp.openById(destId);
    
    // Copy CADASTRO sheet
    const shCad = ss.getSheetByName(H.SHEETS.CADASTRO);
    if (shCad) {
      let destCad = dest.getSheetByName('CADASTRO_HIPERATIVO');
      if (!destCad) destCad = dest.insertSheet('CADASTRO_HIPERATIVO');
      else destCad.clearContents();
      
      const origem = shCad.getDataRange().getValues();
      const colunas = [
        H.CAD.ID, H.CAD.NOME, H.CAD.EMAIL, H.CAD.WHATS, H.CAD.MOD,
        H.CAD.PLANO, H.CAD.CIDADE, H.CAD.ESTADO, H.CAD.DATA_CAD,
        H.CAD.STRAVA_OK, H.CAD.STRAVA_ID, H.CAD.STATUS
      ];
      const dados = [[
        'ATH_ID', 'Nome', 'E-mail', 'WhatsApp', 'Modalidade', 'Plano',
        'Cidade', 'Estado', 'Data Cadastro', 'Strava', 'ID Strava', 'Status'
      ]];
      origem.slice(2).forEach(row => {
        if (_isAthIdValido_(row[H.CAD.ID - 1])) {
          dados.push(colunas.map(col => row[col - 1]));
        }
      });
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
    sh.appendRow([ts, athId || '', nivel || 'INFO', funcao || '', msg || '', stack || '']);
  } catch(e) { Logger.log('_log error: ' + e.message); }
}

function abrirErros() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(H.SHEETS.ERROS);
  if (sh) ss.setActiveSheet(sh);
}

// atualizarPainel está definida em Metricas.gs (versão completa com cálculo de métricas)

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
  // Compatibilidade com atalhos antigos: nunca grava segredo no código-fonte.
  configurarCredenciais();
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

function salvarConfig() { configurarCredenciais(); }


// ── GERAR LINKS DE RECONEXAO STRAVA ──────────────────────────────────────
function gerarLinksReconexaoTodos() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var wsCad = ss.getSheetByName(H.SHEETS.CADASTRO);
  if (!wsCad || wsCad.getLastRow() < 3) { SpreadsheetApp.getUi().alert("Sem atletas."); return; }
  var rows = wsCad.getDataRange().getValues().slice(2);
  var inv = rows.filter(function(r) {
    var id = String(r[H.CAD.ID-1]||"").trim();
    var ok = String(r[H.CAD.STRAVA_OK-1]||"").trim();
    var st = String(r[H.CAD.STATUS-1]||"").trim();
    return id && st !== "Inativo" && ok !== "Sim";
  });
  if (!inv.length) { SpreadsheetApp.getUi().alert("Todos ja conectados!"); return; }
  var txt = "";
  inv.forEach(function(r) {
    var athId = String(r[H.CAD.ID-1]||"").trim();
    var nome  = String(r[H.CAD.NOME-1]||athId).trim();
    var whats = String(r[H.CAD.WHATS-1]||"").trim();
    var fn    = nome.split(" ")[0];
    var url   = "";
    try { url = _gerarUrlOAuth(athId); } catch(e) { url = "ERRO: "+e.message; }
    txt += "=== " + nome + " (WhatsApp: " + (whats||athId) + ") ===\n";
    txt += "Ola, " + fn + "! Para conectar seu Strava ao HIPERATIVO clique:\n";
    txt += url + "\n";
    txt += "Apos clicar, autorize no Strava. Pode aparecer aviso - ignore. Voce recebera e-mail de confirmacao.\n\n";
  });
  var html = "<style>body{font-family:Arial;padding:12px}</style>" +
    "<h3 style=\"color:#FC4C02\">" + inv.length + " atletas para reconectar</h3>" +
    "<p><b>Clique na area abaixo, Ctrl+A para selecionar tudo, depois Ctrl+C para copiar:</b></p>" +
    "<textarea style=\"width:100%;height:380px;font-size:12px\" onclick=\"this.select()\">" + txt + "</textarea>";
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(html).setWidth(680).setHeight(500),
    "Mensagens WhatsApp - Reconexao Strava"
  );
}

// ── AUDITORIA DE TOKENS STRAVA ─────────────────────────────────────────────
function auditarTokensStrava() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var wsTok = ss.getSheetByName(H.SHEETS.TOKENS);
  if (!wsTok) { Logger.log("Aba TOKENS nao encontrada."); return; }
  var data = wsTok.getDataRange().getValues();
  var agora = new Date();
  var ok = [], semRefresh = [], expirados = [], revogados = [];
  for (var i = 1; i < data.length; i++) {
    var r = data[i];
    var athId   = String(r[0] || "").trim();
    if (!athId) continue;
    var refresh = String(r[3] || "").trim();
    var expires = r[4];
    var status  = String(r[7] || "").trim();
    if (status === "Revogado" || status === "Expirado") {
      revogados.push(athId + " [" + status + "]");
    } else if (!refresh) {
      semRefresh.push(athId);
    } else if (expires instanceof Date && expires < agora) {
      expirados.push(athId);
    } else {
      ok.push(athId + " (expira: " + (expires instanceof Date ? Utilities.formatDate(expires, "America/Sao_Paulo", "dd/MM HH:mm") : "?") + ")");
    }
  }
  Logger.log("=== AUDITORIA TOKENS STRAVA ===");
  Logger.log("OK (com refresh_token): " + ok.length + " | " + ok.join(", "));
  Logger.log("SEM refresh_token: " + semRefresh.length + " | " + semRefresh.join(", "));
  Logger.log("ACCESS expirado: " + expirados.length + " | " + expirados.join(", "));
  Logger.log("REVOGADOS: " + revogados.length + " | " + revogados.join(", "));
  Logger.log("TOTAL: " + (ok.length+semRefresh.length+expirados.length+revogados.length));
}

function checarSemStrava() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var wsCad = ss.getSheetByName(H.SHEETS.CADASTRO);
  var data = wsCad.getDataRange().getValues().slice(2);
  var semStrava = data.filter(function(r) {
    return String(r[H.CAD.ID-1]||"").trim() &&
           String(r[H.CAD.STATUS-1]||"").trim() !== "Inativo" &&
           String(r[H.CAD.STRAVA_OK-1]||"").trim() !== "Sim";
  });
  Logger.log("Atletas SEM Strava: " + semStrava.length);
  semStrava.forEach(function(r) {
    var id = String(r[H.CAD.ID-1]||"").trim();
    var nome = String(r[H.CAD.NOME-1]||"").trim();
    var ok = String(r[H.CAD.STRAVA_OK-1]||"Vazio").trim();
    Logger.log(id + " | " + nome + " | STRAVA_OK: " + ok);
  });
}

// ── CONFIGURAR URL DA WEBAPP ──────────────────────────────────────────────────
function configurarUrlWebApp() {
  const ui    = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();
  const atual = props.getProperty('WEBAPP_URL') || '(não configurado)';
  const r = ui.prompt(
    '🔗 Configurar URL do WebApp',
    'Valor atual: ' + atual + '\n\nCole a URL de implantação (termina em /exec):' +
    '\n(Apps Script → Implantar → Gerenciar implantações)',
    ui.ButtonSet.OK_CANCEL
  );
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const url = r.getResponseText().trim();
  if (!url) { ui.alert('⚠️ Nenhuma URL informada. Cancelado.'); return; }
  if (url.indexOf('https://') !== 0) { ui.alert('❌ URL inválida. Deve começar com https://'); return; }
  props.setProperty('WEBAPP_URL', url);
  ui.alert('✅ WEBAPP_URL salva com sucesso!\n\n' + url);
}
