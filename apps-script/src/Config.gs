/**
 * ═══════════════════════════════════════════════════════════════════════
 * HIPERATIVO V3 — CONFIG.GS
 * Menu, constantes globais e setup inicial
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
    RANKING:   '🏆 RANKING',
    INPUT_MANUAL: '📝 INPUT MANUAL',
    DESAFIOS:  '🎯 DESAFIOS',
    DESAFIOS_PROGRESSO: '📊 DESAFIOS_PROGRESSO',
  },

  // Colunas da aba CADASTRO (1-indexed)
  CAD: {
    ID: 1, NOME: 2, EMAIL: 3, WHATS: 4, NASC: 5,
    SEXO: 6, PESO: 7, MOD: 8, NIVEL: 9, OBJ: 10,
    FREQ: 11, HORARIO: 12, SAUDE: 13, LESAO: 14,
    PLANO: 15, ORIGEM: 16, DATA_CAD: 17,
    STRAVA_OK: 18, STRAVA_ID: 19, STATUS: 20, OBS: 21,
    INTERVALS_ID: 22,
  },

  // Colunas da aba ATIVIDADES (1-indexed)
  ATIV: {
    // ── Identificação ──────────────────────────────────────────────────────
    EXEC_ID: 1, ATH_ID: 2, NOME: 3, DATA: 4, TIPO: 5,
    FONTE: 6, STRAVA_ID: 7, NOME_ATIV: 8,
    // ── Tempo (hh:mm:ss) ──────────────────────────────────────────────────
    MOV_HMS: 9,     // Tempo de Movimentação (hh:mm:ss)
    TOTAL_HMS: 10,  // Tempo Total          (hh:mm:ss)
    // ── Distância ─────────────────────────────────────────────────────────
    DIST_M: 11,     // metros (1 decimal)
    DIST_KM: 12,    // km (1 decimal)
    // ── Velocidade e Pace ─────────────────────────────────────────────────
    VEL_KMH: 13,    // km/h (1 decimal)
    PACE_S: 14,     // s/km  (usado nas fórmulas do painel)
    PACE_MS: 15,    // min:ss/km (texto)
    PACE_RAPIDO: 16,
    // ── Cardíaco ──────────────────────────────────────────────────────────
    FC_MED: 17, FC_MAX: 18,
    // ── Esforço físico ────────────────────────────────────────────────────
    ELEV: 19, CAL: 20,
    // ── Campos manuais ────────────────────────────────────────────────────
    RPE: 21, OBS: 22,
    // ── Extras Strava ─────────────────────────────────────────────────────
    SPORT_TYPE: 23, // sport_type (ex: TrailRun, VirtualRide)
    CADENCIA: 24,   // ppm/rpm médio
    ESFORCO_REL: 25, // suffer_score (Strava relative effort 0-100)
    DATA_IMPORT: 26, // timestamp da importação
  },

  // Colunas da aba PLANO SEMANAL (1-indexed)
  PLANO: {
    ID: 1, ATH_ID: 2, NOME: 3, SEMANA: 4, CICLO: 5,
    MOD: 6, TITULO: 7, DIAG: 8,
    D1_TIPO: 9, D1_PRESC: 10,
    D2_TIPO: 11, D2_PRESC: 12,
    D3_TIPO: 13, D3_PRESC: 14,
    INTENCAO: 15, STATUS: 16,
  },

  // Colunas da aba TOKENS (1-indexed)
  TOKENS: {
    ATH_ID: 1, STRAVA_ATH_ID: 2, ACCESS: 3,
    REFRESH: 4, EXPIRES: 5, DATA_CON: 6,
    ULT_ATU: 7, STATUS: 8,
  },

  // Colunas da aba MÉTRICAS (1-indexed)
  MET: {
    ATH_ID: 1, NOME: 2, ATUALIZADO: 3, VO2: 4,
    PACE_MED: 5, PACE_RAP: 6, PACE_LEN: 7,
    FC_MAX: 8, FC_MED: 9, VOL_SEM: 10,
    CTL: 11, ATL: 12, TSB: 13, RAMP_RATE: 14, FONTE_VO2MAX: 15,
  },

  // Colunas da aba ERROS (1-indexed)
  ERROS: {
    DATA: 1, NIVEL: 2, ATH_ID: 3, FUNCAO: 4,
    MSG: 5, JSON: 6, RESOLVIDO: 7, ACAO: 8,
  },
};

// ── LER CONFIG DA PLANILHA ───────────────────────────────────────────────────
function getCfg(chave) {
  try {
    const ss = SpreadsheetApp.getActive();
    const ws = ss.getSheetByName(H.SHEETS.CONFIG);
    if (!ws) return null;
    const dados = ws.getRange('A3:B20').getValues();
    for (const [k, v] of dados) {
      if (k === chave) return v ? String(v).trim() : null;
    }
    return null;
  } catch (e) {
    return null;
  }
}

// ── MENU ─────────────────────────────────────────────────────────────────────
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('⚡ HIPERATIVO')
    .addSubMenu(
      ui.createMenu('🛠️ Setup')
        .addItem('Criar / Recriar todas as abas', 'setupPlanilha')
        .addItem('Configurar credenciais (Strava + WebApp URL)', 'setupPropriedades')
        .addItem('Criar trigger automático (diário 06h)', 'criarTrigger')
        .addItem('🔍 Diagnóstico rápido (credenciais + tokens)', 'diagnosticoRapido')
        .addItem('🔧 Corrigir IDs na aba TOKENS', 'corrigirTokenIds')
        .addItem('🧹 Desativar duplicados Trial', 'limparDuplicadosTrial')
        .addItem('🔗 Gerar link Strava — Amanda', 'gerarLinkStravaAmanda')
    )
    .addSeparator()
    .addSubMenu(
      ui.createMenu('👤 Atletas')
        .addItem('Gerar link Strava para atleta', 'gerarLinkStrava')
        .addItem('🔗 Gerar links reconexão (tokens inválidos)', 'gerarLinksReconexaoTodos')
        .addSeparator()
        .addItem('📥 Importar Strava — mês vigente (todos)', 'importarStravaAll')
        .addItem('📥 Importar Strava — período personalizado', 'importarStravaPeriodo')
        .addItem('📥 Importar Strava — atleta específico', 'importarStravaUm')
        .addSeparator()
        .addItem('Recalcular métricas de todos', 'calcularMetricasTodos')
        .addSeparator()
        .addItem('🔍 Diagnóstico de conexão Strava', 'diagnosticarConexaoStrava')
        .addItem('🧹 Limpar tokens duplicados/revogados', 'limparTokensInvalidos')
    )
    .addSeparator()
    .addItem('📊 Atualizar painel agora', 'atualizarPainel')
    .addItem('🏆 Atualizar ranking mensal', 'atualizarRanking')
    .addItem('📋 Ver log de erros', 'abrirErros')
    .addToUi();

  // Sincroniza status Strava ao abrir (leve — sem chamadas externas)
  try { sincronizarStatusStrava(); } catch(e) {}
}

// ── SETUP DE PROPRIEDADES GUIADO ─────────────────────────────────────────────
function setupPropriedades() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();

  const clientId = ui.prompt(
    '⚙️ Setup — Passo 1/3',
    'Cole seu STRAVA_CLIENT_ID\n(encontrado em strava.com/settings/api):',
    ui.ButtonSet.OK_CANCEL
  );
  if (clientId.getSelectedButton() !== ui.Button.OK) return;

  const clientSecret = ui.prompt(
    '⚙️ Setup — Passo 2/3',
    'Cole seu STRAVA_CLIENT_SECRET:',
    ui.ButtonSet.OK_CANCEL
  );
  if (clientSecret.getSelectedButton() !== ui.Button.OK) return;

  const webAppUrl = ui.prompt(
    '⚙️ Setup — Passo 3/3',
    'Cole a URL do Web App publicado\n(Implantar → Gerenciar → copiar URL /exec):',
    ui.ButtonSet.OK_CANCEL
  );
  if (webAppUrl.getSelectedButton() !== ui.Button.OK) return;

  props.setProperties({
    'STRAVA_CLIENT_ID':     clientId.getResponseText().trim(),
    'STRAVA_CLIENT_SECRET': clientSecret.getResponseText().trim(),
    'WEBAPP_URL':           webAppUrl.getResponseText().trim(),
  });

  // Salvar também na aba CONFIG
  const ws = SpreadsheetApp.getActive().getSheetByName(H.SHEETS.CONFIG);
  if (ws) {
    const dados = ws.getRange('A3:B20').getValues();
    dados.forEach(([k], i) => {
      if (k === 'STRAVA_CLIENT_ID') ws.getRange(3 + i, 2).setValue(clientId.getResponseText().trim());
      if (k === 'WEBAPP_URL')       ws.getRange(3 + i, 2).setValue(webAppUrl.getResponseText().trim());
    });
  }

  ui.alert('✅ Credenciais salvas!',
    'Próximo passo: Implantar → Nova implantação → Web App\n' +
    'Executar como: Você | Acesso: Qualquer pessoa\n\n' +
    'Depois use: Menu → Criar trigger de importação (2h)',
    ui.ButtonSet.OK);
}

// ── TRIGGER AUTOMÁTICO ───────────────────────────────────────────────────────
function criarTrigger() {
  // Remove triggers antigos do mesmo tipo para não duplicar
  ScriptApp.getProjectTriggers().forEach(t => {
    const fn = t.getHandlerFunction();
    if (fn === 'importarStravaAll' || fn === 'triggerImportacaoAutomatica') {
      ScriptApp.deleteTrigger(t);
    }
  });

  // Trigger diário às 06h — importa só o mês corrente (evita timeout)
  ScriptApp.newTrigger('triggerImportacaoAutomatica')
    .timeBased()
    .everyDays(1)
    .atHour(6)
    .create();

  SpreadsheetApp.getUi().alert(
    '✅ Trigger diário criado',
    'O sistema vai importar atividades do Strava todos os dias às 06h.\n\n' +
    '📅 Importa apenas o MÊS CORRENTE (evita timeout de 6 min).\n\n' +
    'Para importar período específico use:\nMenu → Atletas → Importar Strava — período personalizado',
    SpreadsheetApp.getUi().ButtonSet.OK
  );

  _log('SYSTEM', 'INFO', 'Trigger', 'Trigger diário 06h criado (mês corrente)', '');
}

// ── ATUALIZAR WEBAPP_URL APÓS NOVO DEPLOYMENT ────────────────────────────────
function atualizarWebAppUrl() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.prompt(
    '🔗 Novo Deployment — Atualizar URL',
    'Cole a nova URL do Web App (/exec):\n\n' +
    '(Implantar → Gerenciar implantações → copiar URL da versão ativa)',
    ui.ButtonSet.OK_CANCEL
  );
  if (resp.getSelectedButton() !== ui.Button.OK) return;

  const novaUrl = resp.getResponseText().trim();
  if (!novaUrl || !novaUrl.startsWith('https://script.google.com/macros/s/')) {
    ui.alert('❌ URL inválida', 'A URL deve começar com:\nhttps://script.google.com/macros/s/', ui.ButtonSet.OK);
    return;
  }

  // Salvar no PropertiesService
  PropertiesService.getScriptProperties().setProperty('WEBAPP_URL', novaUrl);

  // Atualizar aba CONFIG
  const ws = SpreadsheetApp.getActive().getSheetByName(H.SHEETS.CONFIG);
  if (ws) {
    const dados = ws.getRange('A3:B25').getValues();
    let webAppRow = -1;
    let linkRow   = -1;
    dados.forEach(([k], i) => {
      if (k === 'WEBAPP_URL')    webAppRow = 3 + i;
      if (k === 'LINK_CADASTRO') linkRow   = 3 + i;
    });
    if (webAppRow > 0) {
      ws.getRange(webAppRow, 2).setValue(novaUrl);
      // Garantir que LINK_CADASTRO exista logo abaixo com fórmula
      if (linkRow < 0) {
        ws.insertRowAfter(webAppRow);
        linkRow = webAppRow + 1;
        ws.getRange(linkRow, 1).setValue('LINK_CADASTRO')
          .setFontFamily('Courier New').setFontSize(10).setFontWeight('bold')
          .setFontColor(COR.verde).setBackground(COR.cinza_claro)
          .setHorizontalAlignment('left').setVerticalAlignment('middle');
        ws.getRange(linkRow, 3).setValue('Link a compartilhar com novos alunos — copie desta célula')
          .setFontFamily('Arial').setFontSize(10).setFontStyle('italic')
          .setFontColor('#388E3C').setBackground('#F1F8E9')
          .setHorizontalAlignment('left').setVerticalAlignment('middle');
        ws.setRowHeight(linkRow, 26);
      }
      ws.getRange(linkRow, 2).setFormula(`=B${webAppRow}&"?cadastro=true"`)
        .setFontFamily('Courier New').setFontSize(10)
        .setBackground('#F1F8E9').setFontColor('#2E7D32')
        .setHorizontalAlignment('left').setVerticalAlignment('middle');
    }
  }

  _log('SYSTEM', 'INFO', 'atualizarWebAppUrl', 'WEBAPP_URL atualizada para: ' + novaUrl, '');
  ui.alert('✅ URL atualizada', 'Nova URL salva com sucesso.\n\nCopie o LINK_CADASTRO da aba CONFIG e envie aos alunos.', ui.ButtonSet.OK);
}

// ── PATCH: ADICIONAR LINK_CADASTRO NO CONFIG EXISTENTE (executar uma vez) ────
function adicionarLinkCadastroPatch() {
  const ws = SpreadsheetApp.getActive().getSheetByName(H.SHEETS.CONFIG);
  if (!ws) { Logger.log('CONFIG não encontrado'); return; }

  const dados = ws.getRange('A3:B25').getValues();
  let webAppRow = -1;
  let linkRow   = -1;
  dados.forEach(([k], i) => {
    if (k === 'WEBAPP_URL')    webAppRow = 3 + i;
    if (k === 'LINK_CADASTRO') linkRow   = 3 + i;
  });

  if (webAppRow < 0) { Logger.log('WEBAPP_URL não encontrada'); return; }
  if (linkRow > 0)   { Logger.log('LINK_CADASTRO já existe na linha ' + linkRow); return; }

  ws.insertRowAfter(webAppRow);
  linkRow = webAppRow + 1;
  ws.getRange(linkRow, 1).setValue('LINK_CADASTRO')
    .setFontFamily('Courier New').setFontSize(10).setFontWeight('bold')
    .setFontColor(COR.verde).setBackground(COR.cinza_claro)
    .setHorizontalAlignment('left').setVerticalAlignment('middle');
  ws.getRange(linkRow, 2).setFormula(`=B${webAppRow}&"?cadastro=true"`)
    .setFontFamily('Courier New').setFontSize(10)
    .setBackground('#F1F8E9').setFontColor('#2E7D32')
    .setHorizontalAlignment('left').setVerticalAlignment('middle');
  ws.getRange(linkRow, 3).setValue('Link para novos alunos — copie e compartilhe')
    .setFontFamily('Arial').setFontSize(10).setFontStyle('italic')
    .setFontColor('#388E3C').setBackground('#F1F8E9')
    .setHorizontalAlignment('left').setVerticalAlignment('middle');
  ws.setRowHeight(linkRow, 26);
  Logger.log('✅ LINK_CADASTRO adicionado na linha ' + linkRow + ' — fórmula: =B' + webAppRow + '&"?cadastro=true"');
}

// ── DEFINIR TOKENS AUXILIARES (sem UI, roda no editor) ──────────────────────
function setTokensAuxiliares() {
  PropertiesService.getScriptProperties().setProperty('STRAVA_VERIFY_TOKEN', 'HIPERATIVO_STRAVA_2024');
  Logger.log('✅ STRAVA_VERIFY_TOKEN definido');
}

// ── INIT CONFIG SHEET ONLY (seguro: não apaga outras abas) ──────────────────
function initConfigOnly() {
  const ws = SpreadsheetApp.getActive().getSheetByName(H.SHEETS.CONFIG);
  if (!ws) { Logger.log('CONFIG não encontrado'); return; }
  _criarConfig(ws);
  Logger.log('✅ CONFIG inicializado com sucesso');
}

// ── LOG INTERNO ──────────────────────────────────────────────────────────────
function _log(athId, nivel, funcao, msg, jsonStr) {
  try {
    const ws = SpreadsheetApp.getActive().getSheetByName(H.SHEETS.ERROS);
    if (!ws) return;
    ws.appendRow([
      new Date(), nivel, athId || 'SYSTEM', funcao,
      msg, jsonStr || '', 'Não', '',
    ]);
  } catch (e) { /* silencioso */ }
}

// ── CONFIGURAR CABEÇALHOS DA ABA ATIVIDADES ──────────────────────────────────
function configurarCabecalhosAtividades() {
  const ss = SpreadsheetApp.getActive();
  const ws = ss.getSheetByName(H.SHEETS.ATIVIDADES);
  if (!ws) {
    SpreadsheetApp.getUi().alert('❌ Aba "' + H.SHEETS.ATIVIDADES + '" não encontrada.');
    return;
  }

  const cabecalhos = [
    // Cols 1-8: Identificação
    'ID Execução', 'ID Atleta', 'Nome Atleta', 'Data', 'Tipo de Atividade',
    'Fonte', 'ID Strava', 'Nome da Atividade',
    // Cols 9-10: Tempo
    'Tempo Mov. (hh:mm:ss)', 'Tempo Total (hh:mm:ss)',
    // Cols 11-12: Distância
    'Distância (m)', 'Distância (km)',
    // Cols 13-16: Velocidade e Pace
    'Vel. Média (km/h)', 'Pace Médio (s/km)', 'Pace Médio (min:ss)', 'Pace Mais Rápido',
    // Cols 17-18: Cardíaco
    'FC Média', 'FC Máxima',
    // Cols 19-20: Esforço
    'Elevação (m)', 'Calorias',
    // Cols 21-22: Manuais
    'RPE (1-10)', 'Observações',
    // Cols 23-26: Extras Strava
    'Tipo Esporte', 'Cadência (ppm)', 'Esforço Relativo', 'Data Importação',
  ];

  ws.getRange(2, 1, 1, cabecalhos.length).setValues([cabecalhos]);

  // Formatar cabeçalho (negrito + fundo cinza)
  const rng = ws.getRange(2, 1, 1, cabecalhos.length);
  rng.setFontWeight('bold');
  rng.setBackground('#EFEFEF');

  // Larguras sugeridas
  const larguras = [120, 90, 130, 130, 120, 60, 110, 180,
                    110, 110, 90, 90, 100, 100, 110, 100,
                    70, 70, 85, 80, 70, 130, 110, 90, 110, 120];
  larguras.forEach((w, i) => ws.setColumnWidth(i + 1, w));

  _log('SYSTEM', 'INFO', 'configurarCabecalhosAtividades', 'Cabeçalhos atualizados: 26 colunas', '');
  SpreadsheetApp.getUi().alert('✅ Cabeçalhos atualizados!',
    '26 colunas configuradas na aba ATIVIDADES.\n\n' +
    'Novos campos adicionados:\n' +
    '• Tipo Esporte (col W)\n• Cadência ppm (col X)\n' +
    '• Esforço Relativo (col Y)\n• Data Importação (col Z)',
    SpreadsheetApp.getUi().ButtonSet.OK);
}

// ── ABRIR ERROS ──────────────────────────────────────────────────────────────
function abrirErros() {
  const ss = SpreadsheetApp.getActive();
  ss.setActiveSheet(ss.getSheetByName(H.SHEETS.ERROS));
}

// ── CORRIGIR LINHAS DESLOCADAS ────────────────────────────────────────────────
// As atividades e métricas foram gravadas abaixo das linhas de fórmula (503+/203+)
// por causa do appendRow. Esta função move tudo para as linhas corretas (3+).
function corrigirLinhasDeslocadas() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();
  let msg  = '';

  // ── ATIVIDADES: mover de linhas 503+ para linhas 3+ ──
  const wsAtiv = ss.getSheetByName(H.SHEETS.ATIVIDADES);
  if (wsAtiv) {
    const lastAtiv = wsAtiv.getLastRow();
    if (lastAtiv >= 503) {
      const total  = lastAtiv - 502;
      const dados  = wsAtiv.getRange(503, 1, total, 22).getValues()
        .filter(r => r[H.ATIV.EXEC_ID - 1]);
      if (dados.length > 0) {
        wsAtiv.getRange(3, 1, dados.length, 22).setValues(dados);
        wsAtiv.getRange(503, 1, total, 22).clearContent();
        msg += '✅ ATIVIDADES: ' + dados.length + ' linhas movidas para linha 3+\n';
      }
    } else {
      msg += 'ℹ️ ATIVIDADES: sem dados deslocados\n';
    }
  }

  // ── MÉTRICAS: mover de linhas 203+ para linhas 3+ ──
  const wsMet = ss.getSheetByName(H.SHEETS.METRICAS);
  if (wsMet) {
    const lastMet = wsMet.getLastRow();
    if (lastMet >= 203) {
      const total = lastMet - 202;
      const dados = wsMet.getRange(203, 1, total, 10).getValues()
        .filter(r => r[H.MET.ATH_ID - 1]);
      if (dados.length > 0) {
        wsMet.getRange(3, 1, dados.length, 10).setValues(dados);
        wsMet.getRange(203, 1, total, 10).clearContent();
        msg += '✅ MÉTRICAS: ' + dados.length + ' linhas movidas para linha 3+\n';
      }
    } else {
      msg += 'ℹ️ MÉTRICAS: sem dados deslocados\n';
    }
  }

  if (!msg) msg = 'Nenhuma correção necessária.';
  _log('SYSTEM', 'INFO', 'corrigirLinhasDeslocadas', msg.replace(/\n/g, ' | '), '');
  ui.alert('🔧 Correção concluída', msg, ui.ButtonSet.OK);
}

// ── CORRIGIR FÓRMULAS pt-BR ───────────────────────────────────────────────────
// O setup gerou fórmulas com vírgula (sintaxe US). A localidade pt-BR exige
// ponto-e-vírgula como separador. Esta função reescreve todas as fórmulas
// do PAINEL, GRÁFICOS e zonas de MÉTRICAS com o separador correto.
function corrigirFormulas() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();

  // ── PAINEL ─────────────────────────────────────────────────────────────────
  const wsP = ss.getSheetByName(H.SHEETS.PAINEL);
  if (wsP) {
    // KPI linha 6 — corrige fórmulas e o bug do Total (COUNTA - COUNTBLANK = negativo)
    const kpiCols = [1, 3, 5, 7, 9, 11];
    const kpiFmls = [
      "=COUNTA('👤 CADASTRO'!A3:A500)",
      "=COUNTIF('👤 CADASTRO'!T3:T500;\"Ativo\")",
      "=COUNTIF('👤 CADASTRO'!R3:R500;\"Sim\")",
      "=COUNTIFS('🏃 ATIVIDADES'!D3:D500;\">=\"&TODAY()-7)",
      "=IFERROR(TEXT(INT(AVERAGEIF('🏃 ATIVIDADES'!E3:E500;\"Corrida\";'🏃 ATIVIDADES'!N3:N500)/60);\"0\")&\":\"&TEXT(MOD(ROUND(AVERAGEIF('🏃 ATIVIDADES'!E3:E500;\"Corrida\";'🏃 ATIVIDADES'!N3:N500);0);60);\"00\");\"--\")",
      "=IFERROR(ROUND(SUMIF('🏃 ATIVIDADES'!E3:E500;\"Corrida\";'🏃 ATIVIDADES'!L3:L500)/MAX(1;COUNTIF('👤 CADASTRO'!T3:T500;\"Ativo\"));1);\"--\")",
    ];
    kpiCols.forEach((col, i) => wsP.getRange(6, col).setFormula(kpiFmls[i]));

    // Últimas atividades (linhas 11-20, colunas 1-6) e alertas (colunas 8-11)
    const fmlsAtiv = [], fmlsAlert = [];
    for (let r = 11; r <= 20; r++) {
      const idx = r - 10;
      fmlsAtiv.push([
        `=IFERROR(TEXT(LARGE('🏃 ATIVIDADES'!D$3:D$500;${idx});"DD/MM/AAAA");"--")`,
        `=IFERROR(INDEX('🏃 ATIVIDADES'!C$3:C$500;MATCH(LARGE('🏃 ATIVIDADES'!D$3:D$500;${idx});'🏃 ATIVIDADES'!D$3:D$500;0));"--")`,
        `=IFERROR(INDEX('🏃 ATIVIDADES'!E$3:E$500;MATCH(LARGE('🏃 ATIVIDADES'!D$3:D$500;${idx});'🏃 ATIVIDADES'!D$3:D$500;0));"--")`,
        `=IFERROR(ROUND(INDEX('🏃 ATIVIDADES'!L$3:L$500;MATCH(LARGE('🏃 ATIVIDADES'!D$3:D$500;${idx});'🏃 ATIVIDADES'!D$3:D$500;0));2);"--")`,
        `=IFERROR(INDEX('🏃 ATIVIDADES'!O$3:O$500;MATCH(LARGE('🏃 ATIVIDADES'!D$3:D$500;${idx});'🏃 ATIVIDADES'!D$3:D$500;0));"--")`,
        `=IFERROR(ROUND(INDEX('🏃 ATIVIDADES'!Q$3:Q$500;MATCH(LARGE('🏃 ATIVIDADES'!D$3:D$500;${idx});'🏃 ATIVIDADES'!D$3:D$500;0));0);"--")`,
      ]);
      fmlsAlert.push([
        `=IFERROR(INDEX('👤 CADASTRO'!B$3:B$500;${idx});"--")`,
        `=IFERROR(TEXT(MAXIFS('🏃 ATIVIDADES'!D$3:D$500;'🏃 ATIVIDADES'!C$3:C$500;H${r});"DD/MM/AA");"Nunca")`,
        `=IFERROR(TODAY()-MAXIFS('🏃 ATIVIDADES'!D$3:D$500;'🏃 ATIVIDADES'!C$3:C$500;H${r});"--")`,
        `=IFERROR(IF(J${r}>14;"⚠️ Verificar";"✅ OK");"--")`,
      ]);
    }
    wsP.getRange(11, 1, 10, 6).setFormulas(fmlsAtiv);
    wsP.getRange(11, 8, 10, 4).setFormulas(fmlsAlert);
  }

  // ── GRÁFICOS ───────────────────────────────────────────────────────────────
  const wsG = ss.getSheetByName(H.SHEETS.GRAFICOS);
  if (wsG) {
    const gFmls = [
      ["=SUMPRODUCT((MONTH('🏃 ATIVIDADES'!D3:D500)=MONTH(TODAY()))*(YEAR('🏃 ATIVIDADES'!D3:D500)=YEAR(TODAY()))*('🏃 ATIVIDADES'!E3:E500=\"Corrida\")*('🏃 ATIVIDADES'!L3:L500))"],
      ["=COUNTIFS('🏃 ATIVIDADES'!E3:E500;\"Corrida\";'🏃 ATIVIDADES'!D3:D500;\">=\"&DATE(YEAR(TODAY());MONTH(TODAY());1))"],
      ["=IFERROR(TEXT(INT(AVERAGEIF('🏃 ATIVIDADES'!E3:E500;\"Corrida\";'🏃 ATIVIDADES'!N3:N500)/60);\"0\")&\":\"&TEXT(MOD(ROUND(AVERAGEIF('🏃 ATIVIDADES'!E3:E500;\"Corrida\";'🏃 ATIVIDADES'!N3:N500);0);60);\"00\");\"--\")"],
      ["=IFERROR(ROUND(AVERAGEIF('🏃 ATIVIDADES'!E3:E500;\"Corrida\";'🏃 ATIVIDADES'!Q3:Q500);0);\"--\")"],
      ["=IFERROR(ROUND(AVERAGE('📈 MÉTRICAS'!D3:D202);1);\"--\")"],
      ["=COUNTIF('📊 PAINEL'!K11:K20;\"⚠️ Verificar\")"],
    ];
    wsG.getRange(4, 2, 6, 1).setFormulas(gFmls);
  }

  // ── MÉTRICAS: zonas Z1–Z5 (cols 11–18) em batch ───────────────────────────
  const wsMet = ss.getSheetByName(H.SHEETS.METRICAS);
  if (wsMet) {
    const fim = Math.min(wsMet.getLastRow(), 202);
    if (fim >= 3) {
      const zoneFmls = [];
      for (let r = 3; r <= fim; r++) {
        zoneFmls.push([
          `=IFERROR(TEXT(INT(E${r}*1.20/60);"0")&":"&TEXT(MOD(ROUND(E${r}*1.20;0);60);"00");"")`,
          `=IFERROR(TEXT(INT(E${r}*1.08/60);"0")&":"&TEXT(MOD(ROUND(E${r}*1.08;0);60);"00");"")`,
          `=IFERROR(TEXT(INT(E${r}*1.04/60);"0")&":"&TEXT(MOD(ROUND(E${r}*1.04;0);60);"00");"")`,
          `=IFERROR(TEXT(INT(E${r}*0.96/60);"0")&":"&TEXT(MOD(ROUND(E${r}*0.96;0);60);"00");"")`,
          `=IFERROR(TEXT(INT(E${r}*0.94/60);"0")&":"&TEXT(MOD(ROUND(E${r}*0.94;0);60);"00");"")`,
          `=IFERROR(TEXT(INT(E${r}*0.87/60);"0")&":"&TEXT(MOD(ROUND(E${r}*0.87;0);60);"00");"")`,
          `=IFERROR(TEXT(INT(E${r}*0.84/60);"0")&":"&TEXT(MOD(ROUND(E${r}*0.84;0);60);"00");"")`,
          `=IFERROR(TEXT(INT(F${r}/60);"0")&":"&TEXT(MOD(F${r};60);"00");"")`,
        ]);
      }
      wsMet.getRange(3, 11, fim - 2, 8).setFormulas(zoneFmls);
    }
  }

  _log('SYSTEM', 'INFO', 'corrigirFormulas', 'Separador pt-BR (;) aplicado: PAINEL, GRÁFICOS, MÉTRICAS', '');
  ui.alert('✅ Fórmulas corrigidas!',
    'PAINEL, GRÁFICOS e zonas de MÉTRICAS atualizados.\n\n' +
    'Próximos passos:\n' +
    '1. Menu → Setup → 🔧 Corrigir linhas deslocadas\n' +
    '2. Menu → 📊 Atualizar painel agora',
    ui.ButtonSet.OK);
}
