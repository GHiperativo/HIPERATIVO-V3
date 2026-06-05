/**
 * ═══════════════════════════════════════════════════════════════════════════
 * HIPERATIVO V3 — SETUP DA PLANILHA
 * Cria todas as abas, cabeçalhos, dropdowns, formatação e fórmulas
 * Execute: Menu ⚡ HIPERATIVO → 🛠️ Setup Inicial da Planilha
 * ═══════════════════════════════════════════════════════════════════════════
 */

// ── PALETA DE CORES ──────────────────────────────────────────────────────────
const COR = {
  azul_escuro:  '#001F3F',
  azul_medio:   '#003D7A',
  azul_claro:   '#00B4FF',
  azul_suave:   '#D6EEFF',
  preto:        '#0A0A0A',
  cinza_escuro: '#2C2C2C',
  cinza_claro:  '#F2F2F2',
  branco:       '#FFFFFF',
  verde:        '#1D9E75',
  verde_claro:  '#D6F5EC',
  vermelho:     '#E24B4A',
  vermelho_cl:  '#FDECEA',
  laranja:      '#FC4C02',
  amarelo:      '#FFC107',
  amarelo_cl:   '#FFF8E1',
  roxo:         '#6C3DC7',
  roxo_claro:   '#EDE9FB',
};

// ── UTILITÁRIOS ──────────────────────────────────────────────────────────────

/** Aplica cabeçalho de linha com cor de fundo e texto branco */
function _cabecalho(ws, linha, colunas, bgHex, fontSize) {
  fontSize = fontSize || 10;
  const range = ws.getRange(linha, 1, 1, colunas.length);
  range.setValues([colunas]);
  range.setFontFamily('Arial');
  range.setFontSize(fontSize);
  range.setFontWeight('bold');
  range.setFontColor('#FFFFFF');
  range.setBackground(bgHex);
  range.setHorizontalAlignment('center');
  range.setVerticalAlignment('middle');
  range.setWrap(true);
  range.setBorder(true, true, true, true, true, true, '#1A3A5C',
    SpreadsheetApp.BorderStyle.SOLID);
  ws.setRowHeight(linha, 26);
}

/** Aplica zebra stripe e bordas em um bloco de dados */
function _zebra(ws, linhaInicio, linhaFim, numCols, cor1, cor2) {
  cor1 = cor1 || COR.branco;
  cor2 = cor2 || COR.azul_suave;
  for (let r = linhaInicio; r <= linhaFim; r++) {
    const rg = ws.getRange(r, 1, 1, numCols);
    rg.setBackground(r % 2 === 0 ? cor1 : cor2);
    rg.setFontFamily('Arial');
    rg.setFontSize(10);
    rg.setVerticalAlignment('middle');
    rg.setBorder(true, true, true, true, true, true, '#DDDDDD',
      SpreadsheetApp.BorderStyle.SOLID);
  }
}

/** Cabeçalho principal de aba */
function _tituloPrincipal(ws, texto, numCols, bgHex, fgHex) {
  fgHex = fgHex || COR.branco;
  ws.getRange(1, 1, 1, numCols)
    .mergeAcross()
    .setValue(texto)
    .setFontFamily('Arial').setFontSize(14).setFontWeight('bold')
    .setFontColor(fgHex).setBackground(bgHex)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  ws.setRowHeight(1, 40);
}

/** Dropdown numa coluna inteira */
function _dropdown(ws, col, linhaInicio, linhaFim, opcoes) {
  const rg = ws.getRange(linhaInicio, col, linhaFim - linhaInicio + 1, 1);
  const regra = SpreadsheetApp.newDataValidation()
    .requireValueInList(opcoes, true)
    .setAllowInvalid(false)
    .build();
  rg.setDataValidation(regra);
}

/** Congelar painel */
function _congelar(ws, linhas, colunas) {
  ws.setFrozenRows(linhas);
  ws.setFrozenColumns(colunas);
}

/** Formatação condicional por valor de texto */
function _condTexto(ws, range, valor, bgHex) {
  const regra = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo(valor)
    .setBackground(bgHex)
    .setRanges([ws.getRange(range)])
    .build();
  const regras = ws.getConditionalFormatRules();
  regras.push(regra);
  ws.setConditionalFormatRules(regras);
}

/** Escala de cor gradiente */
function _escala(ws, range, minCor, midCor, maxCor) {
  const regra = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMinpointWithValue(minCor, SpreadsheetApp.InterpolationType.NUMBER, '1')
    .setGradientMidpointWithValue(midCor, SpreadsheetApp.InterpolationType.PERCENTILE, '50')
    .setGradientMaxpointWithValue(maxCor, SpreadsheetApp.InterpolationType.MAX, '')
    .setRanges([ws.getRange(range)])
    .build();
  const regras = ws.getConditionalFormatRules();
  regras.push(regra);
  ws.setConditionalFormatRules(regras);
}

// ════════════════════════════════════════════════════════════════════════════
// FUNÇÃO PRINCIPAL — executar pelo menu
// ════════════════════════════════════════════════════════════════════════════
function setupPlanilha() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const resp = ui.alert(
    '⚡ HIPERATIVO V3 — Setup da Planilha',
    'Isso vai criar/recriar todas as abas com layout completo.\n' +
    'Abas existentes com os mesmos nomes serão substituídas.\n\nContinuar?',
    ui.ButtonSet.YES_NO
  );
  if (resp !== ui.Button.YES) return;

  // Ordem das abas
  const ABAS = [
    '📊 PAINEL', '👤 CADASTRO', '🏃 ATIVIDADES', '📅 PLANO SEMANAL',
    '📈 MÉTRICAS', '💬 FEEDBACK', '📉 GRÁFICOS', '🔴 ERROS',
    '🔐 TOKENS', '⚙️ CONFIG'
  ];

  // Remover abas existentes com esses nomes
  ABAS.forEach(nome => {
    const existente = ss.getSheetByName(nome);
    if (existente) ss.deleteSheet(existente);
  });

  // Criar abas na ordem
  const sheets = {};
  ABAS.forEach((nome, i) => {
    const ws = i === 0 ? ss.insertSheet(nome, 0) : ss.insertSheet(nome, i);
    sheets[nome] = ws;
    SpreadsheetApp.flush();
  });

  // Remover Sheet1 padrão se existir
  const padrao = ss.getSheetByName('Sheet1') || ss.getSheetByName('Plan1');
  if (padrao) {
    try { ss.deleteSheet(padrao); } catch(e) {}
  }

  _criarPainel(sheets['📊 PAINEL']);
  _criarCadastro(sheets['👤 CADASTRO']);
  _criarAtividades(sheets['🏃 ATIVIDADES']);
  _criarPlanoSemanal(sheets['📅 PLANO SEMANAL']);
  _criarMetricas(sheets['📈 MÉTRICAS']);
  _criarFeedback(sheets['💬 FEEDBACK']);
  _criarGraficos(sheets['📉 GRÁFICOS']);
  _criarErros(sheets['🔴 ERROS']);
  _criarTokens(sheets['🔐 TOKENS']);
  _criarConfig(sheets['⚙️ CONFIG']);

  // Ativar o painel
  ss.setActiveSheet(sheets['📊 PAINEL']);

  ui.alert('✅ Planilha HIPERATIVO V3 criada com sucesso!',
    'Próximos passos:\n' +
    '1. Configure ⚙️ CONFIG com seu Client ID e URL do Web App\n' +
    '2. Use o menu para fazer Setup do Strava OAuth\n' +
    '3. Registre seu primeiro atleta em 👤 CADASTRO',
    ui.ButtonSet.OK);
}


// ════════════════════════════════════════════════════════════════════════════
// 1. PAINEL
// ════════════════════════════════════════════════════════════════════════════
function _criarPainel(ws) {
  ws.setTabColor(COR.azul_claro.replace('#',''));
  ws.clearContents();

  _tituloPrincipal(ws, '⚡ GRUPO HIPERATIVO — PAINEL DE CONTROLE', 12, COR.azul_escuro);

  ws.getRange('A2:L2').mergeAcross()
    .setValue('Cabeça • Coração • Corpo   |   Sistema de Monitoramento e Prescrição')
    .setFontFamily('Arial').setFontSize(10).setFontStyle('italic')
    .setFontColor(COR.azul_claro).setBackground(COR.azul_medio)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  ws.setRowHeight(2, 20);

  ws.getRange('A3').setValue('Atualizado em: ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy'))
    .setFontFamily('Arial').setFontSize(9).setFontStyle('italic').setFontColor('#888888');
  ws.setRowHeight(3, 18);
  ws.setRowHeight(4, 8);

  // KPI Cards
  const kpis = [
    ['Total de Atletas',     '=COUNTA(\'👤 CADASTRO\'!A3:A500)-COUNTBLANK(\'👤 CADASTRO\'!A3:A500)'],
    ['Atletas Ativos',       '=COUNTIF(\'👤 CADASTRO\'!T3:T500,"Ativo")'],
    ['Strava Conectados',    '=COUNTIF(\'👤 CADASTRO\'!R3:R500,"Sim")'],
    ['Treinos esta semana',  '=COUNTIFS(\'🏃 ATIVIDADES\'!D3:D500,">="&TODAY()-7)'],
    ['Pace médio (corrida)', '=IFERROR(TEXT(INT(AVERAGEIF(\'🏃 ATIVIDADES\'!E3:E500,"Corrida",\'🏃 ATIVIDADES\'!N3:N500)/60),"0")&":"&TEXT(MOD(ROUND(AVERAGEIF(\'🏃 ATIVIDADES\'!E3:E500,"Corrida",\'🏃 ATIVIDADES\'!N3:N500),0),60),"00"),"--")'],
    ['km médios/sem',        '=IFERROR(ROUND(SUMIF(\'🏃 ATIVIDADES\'!E3:E500,"Corrida",\'🏃 ATIVIDADES\'!L3:L500)/MAX(1,COUNTIF(\'👤 CADASTRO\'!T3:T500,"Ativo")),1),"--")'],
  ];
  const kpiCores = [COR.azul_escuro, COR.verde, COR.laranja, COR.roxo, COR.azul_claro, COR.azul_medio];
  const kpiCols  = [1, 3, 5, 7, 9, 11];

  kpis.forEach(([titulo, formula], i) => {
    const col = kpiCols[i];
    const cor  = kpiCores[i];

    ws.getRange(5, col, 1, 2).mergeAcross()
      .setValue(titulo.toUpperCase())
      .setFontFamily('Arial').setFontSize(8).setFontWeight('bold')
      .setFontColor(COR.branco).setBackground(cor)
      .setHorizontalAlignment('center').setVerticalAlignment('middle');

    ws.getRange(6, col, 1, 2).mergeAcross()
      .setFormula(formula)
      .setFontFamily('Arial').setFontSize(22).setFontWeight('bold')
      .setFontColor(cor).setBackground(COR.cinza_claro)
      .setHorizontalAlignment('center').setVerticalAlignment('middle');

    ws.getRange(7, col, 1, 2).mergeAcross()
      .setBackground(COR.cinza_claro);

    [5,6,7].forEach(r => ws.setRowHeight(r, 28));
  });

  ws.setRowHeight(8, 10);

  // Últimas atividades
  ws.getRange('A9:F9').mergeAcross()
    .setValue('📋  ÚLTIMAS ATIVIDADES IMPORTADAS DO STRAVA')
    .setFontFamily('Arial').setFontSize(11).setFontWeight('bold')
    .setFontColor(COR.branco).setBackground(COR.azul_medio)
    .setVerticalAlignment('middle');
  ws.setRowHeight(9, 24);

  _cabecalho(ws, 10, ['Data','Atleta','Tipo','Dist. (km)','Pace Médio','FC Média'], COR.azul_escuro, 9);

  for (let r = 11; r <= 20; r++) {
    const idx = r - 10;
    ws.getRange(r, 1).setFormula(`=IFERROR(TEXT(LARGE('🏃 ATIVIDADES'!D$3:D$500,${idx}),"DD/MM/AAAA"),"--")`);
    ws.getRange(r, 2).setFormula(`=IFERROR(INDEX('🏃 ATIVIDADES'!C$3:C$500,MATCH(LARGE('🏃 ATIVIDADES'!D$3:D$500,${idx}),'🏃 ATIVIDADES'!D$3:D$500,0)),"--")`);
    ws.getRange(r, 3).setFormula(`=IFERROR(INDEX('🏃 ATIVIDADES'!E$3:E$500,MATCH(LARGE('🏃 ATIVIDADES'!D$3:D$500,${idx}),'🏃 ATIVIDADES'!D$3:D$500,0)),"--")`);
    ws.getRange(r, 4).setFormula(`=IFERROR(ROUND(INDEX('🏃 ATIVIDADES'!L$3:L$500,MATCH(LARGE('🏃 ATIVIDADES'!D$3:D$500,${idx}),'🏃 ATIVIDADES'!D$3:D$500,0)),2),"--")`);
    ws.getRange(r, 5).setFormula(`=IFERROR(INDEX('🏃 ATIVIDADES'!O$3:O$500,MATCH(LARGE('🏃 ATIVIDADES'!D$3:D$500,${idx}),'🏃 ATIVIDADES'!D$3:D$500,0)),"--")`);
    ws.getRange(r, 6).setFormula(`=IFERROR(ROUND(INDEX('🏃 ATIVIDADES'!Q$3:Q$500,MATCH(LARGE('🏃 ATIVIDADES'!D$3:D$500,${idx}),'🏃 ATIVIDADES'!D$3:D$500,0)),0),"--")`);
    const bg = r % 2 === 0 ? COR.branco : COR.azul_suave;
    ws.getRange(r, 1, 1, 6).setBackground(bg).setFontSize(9).setHorizontalAlignment('center').setVerticalAlignment('middle');
    ws.setRowHeight(r, 18);
  }

  // Alertas
  ws.getRange('H9:L9').mergeAcross()
    .setValue('⚠️  ALERTAS — SEM TREINAR HÁ MAIS DE 14 DIAS')
    .setFontFamily('Arial').setFontSize(11).setFontWeight('bold')
    .setFontColor(COR.branco).setBackground(COR.vermelho)
    .setVerticalAlignment('middle');

  const alertHeaders = ['Atleta','Último Treino','Dias','Alerta'];
  alertHeaders.forEach((h, i) => {
    ws.getRange(10, 8 + i)
      .setValue(h).setFontFamily('Arial').setFontSize(9).setFontWeight('bold')
      .setFontColor(COR.branco).setBackground(COR.cinza_escuro)
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
  });

  for (let r = 11; r <= 20; r++) {
    const idx = r - 10;
    ws.getRange(r, 8).setFormula(`=IFERROR(INDEX('👤 CADASTRO'!B$3:B$500,${idx}),"--")`);
    ws.getRange(r, 9).setFormula(`=IFERROR(TEXT(MAXIFS('🏃 ATIVIDADES'!D$3:D$500,'🏃 ATIVIDADES'!C$3:C$500,H${r}),"DD/MM/AA"),"Nunca")`);
    ws.getRange(r, 10).setFormula(`=IFERROR(TODAY()-MAXIFS('🏃 ATIVIDADES'!D$3:D$500,'🏃 ATIVIDADES'!C$3:C$500,H${r}),"--")`);
    ws.getRange(r, 11).setFormula(`=IFERROR(IF(J${r}>14,"⚠️ Verificar","✅ OK"),"--")`);
    const bg = r % 2 === 0 ? COR.branco : COR.amarelo_cl;
    ws.getRange(r, 8, 1, 4).setBackground(bg).setFontSize(9).setHorizontalAlignment('center').setVerticalAlignment('middle');
    ws.setRowHeight(r, 18);
  }

  ws.setColumnWidth(1, 12 * 7);
  [1,14,14,12,12,12].forEach((w, i) => ws.setColumnWidth(i+1, w*7));
  ws.setFrozenRows(3);
  ws.setHiddenGridlines(true);
}


// ════════════════════════════════════════════════════════════════════════════
// 2. CADASTRO
// ════════════════════════════════════════════════════════════════════════════
function _criarCadastro(ws) {
  ws.setTabColor(COR.verde.replace('#',''));

  _tituloPrincipal(ws, '👤  CADASTRO DE ATLETAS — GRUPO HIPERATIVO', 21, COR.azul_escuro);

  const cols = [
    'ID Atleta','Nome Completo','E-mail','WhatsApp','Data Nasc.',
    'Sexo','Peso (kg)','Modalidade','Nível','Objetivo',
    'Freq./Semana','Horário Pref.','Condições de Saúde','Histórico de Lesão',
    'Plano','Canal de Origem','Data Cadastro','Strava Conectado',
    'ID Strava','Status','Observações'
  ];
  _cabecalho(ws, 2, cols, COR.azul_medio);
  _zebra(ws, 3, 202, 21, COR.branco, COR.azul_suave);
  _congelar(ws, 2, 2);

  // Dropdowns
  _dropdown(ws, 6,  3, 500, ['Masculino','Feminino','Outro']);
  _dropdown(ws, 8,  3, 500, ['Corrida','Academia','Ergométrica','Híbrido']);
  _dropdown(ws, 9,  3, 500, ['Iniciante','Intermediário','Avançado']);
  _dropdown(ws, 10, 3, 500, ['Emagrecimento','Condicionamento','Performance','Saúde Mental','Força']);
  _dropdown(ws, 11, 3, 500, ['2x','3x','4x','5x']);
  _dropdown(ws, 12, 3, 500, ['Manhã','Tarde','Noite','Variável']);
  _dropdown(ws, 15, 3, 500, ['Experimental','Online R$100','Presencial R$120','Híbrido R$160']);
  _dropdown(ws, 16, 3, 500, ['Instagram','Indicação','Strava','Google','Outro']);
  _dropdown(ws, 18, 3, 500, ['Sim','Não','Pendente']);
  _dropdown(ws, 20, 3, 500, ['Ativo','Inativo','Suspenso','Trial']);

  // Formatação condicional Status
  _condTexto(ws, 'T3:T500', 'Ativo',    COR.verde_claro);
  _condTexto(ws, 'T3:T500', 'Inativo',  COR.vermelho_cl);
  _condTexto(ws, 'T3:T500', 'Trial',    COR.amarelo_cl);
  _condTexto(ws, 'T3:T500', 'Suspenso', '#E8E8E8');
  _condTexto(ws, 'R3:R500', 'Sim',      COR.verde_claro);
  _condTexto(ws, 'R3:R500', 'Não',      COR.vermelho_cl);

  const larguras = [12,22,24,16,12,9,9,13,13,18,11,12,22,22,20,16,14,15,12,10,26];
  larguras.forEach((w, i) => ws.setColumnWidth(i+1, w*7));
  ws.setRowHeight(2, 28);
  ws.setHiddenGridlines(true);
}


// ════════════════════════════════════════════════════════════════════════════
// 3. ATIVIDADES
// ════════════════════════════════════════════════════════════════════════════
function _criarAtividades(ws) {
  ws.setTabColor(COR.laranja.replace('#',''));

  _tituloPrincipal(ws, '🏃  ATIVIDADES — DADOS IMPORTADOS DO STRAVA', 22, COR.laranja);

  const cols = [
    'ID Execução','ID Atleta','Nome Atleta','Data','Tipo',
    'Fonte','ID Strava','Nome da Atividade','Tempo Mov. (s)',
    'Tempo Total (s)','Distância (m)','Distância (km)','Vel. Média (m/s)',
    'Pace Médio (s/km)','Pace Médio (min:s)','Pace Mais Rápido',
    'FC Média','FC Máxima','Elevação (m)','Calorias',
    'RPE (1-10)','Observações'
  ];
  _cabecalho(ws, 2, cols, COR.cinza_escuro);
  _zebra(ws, 3, 502, 22, COR.branco, '#FFF5EE');
  _congelar(ws, 2, 4);

  // Fórmulas calculadas para linhas 3..502
  for (let r = 3; r <= 502; r++) {
    // Distância km
    ws.getRange(r, 12).setFormula(`=IFERROR(K${r}/1000,"")`);
    // Pace médio s/km
    ws.getRange(r, 14).setFormula(`=IFERROR(ROUND(1000/M${r},0),"")`);
    // Pace médio min:s
    ws.getRange(r, 15).setFormula(`=IFERROR(TEXT(INT(N${r}/60),"0")&":"&TEXT(MOD(N${r},60),"00"),"")`);
    // Pace mais rápido estimado (~82% do pace médio)
    ws.getRange(r, 16).setFormula(`=IFERROR(TEXT(INT(N${r}*0.82/60),"0")&":"&TEXT(MOD(ROUND(N${r}*0.82,0),60),"00"),"")`);
  }

  // Dropdowns
  _dropdown(ws, 5,  3, 502, ['Corrida','Academia','Ergométrica','Caminhada','Ciclismo','Natação']);
  _dropdown(ws, 6,  3, 502, ['Strava','Manual','App']);
  _dropdown(ws, 21, 3, 502, ['1','2','3','4','5','6','7','8','9','10']);

  // Formatação condicional Tipo
  _condTexto(ws, 'E3:E502', 'Corrida',    '#D6EEFF');
  _condTexto(ws, 'E3:E502', 'Academia',   COR.verde_claro);
  _condTexto(ws, 'E3:E502', 'Ergométrica',COR.amarelo_cl);

  const larguras = [13,11,18,12,12,9,13,22,14,13,12,12,13,14,14,14,10,10,11,10,10,22];
  larguras.forEach((w, i) => ws.setColumnWidth(i+1, w*7));
  ws.setRowHeight(2, 28);
  ws.setHiddenGridlines(true);
}


// ════════════════════════════════════════════════════════════════════════════
// 4. PLANO SEMANAL
// ════════════════════════════════════════════════════════════════════════════
function _criarPlanoSemanal(ws) {
  ws.setTabColor(COR.roxo.replace('#',''));

  _tituloPrincipal(ws, '📅  PLANO SEMANAL — PRESCRIÇÕES CCC', 16, COR.roxo);

  const cols = [
    'ID Plano','ID Atleta','Nome Atleta','Semana Início',
    'Semana do Ciclo','Modalidade','Título da Semana (Remédio Simbólico)',
    'Diagnóstico CCC','Dia 1 — Tipo','Dia 1 — Prescrição',
    'Dia 2 — Tipo','Dia 2 — Prescrição','Dia 3 — Tipo','Dia 3 — Prescrição',
    'Intenção Simbólica','Status Execução'
  ];
  _cabecalho(ws, 2, cols, COR.roxo);
  _zebra(ws, 3, 202, 16, COR.branco, COR.roxo_claro);
  _congelar(ws, 2, 3);

  for (let r = 3; r <= 202; r++) {
    ws.setRowHeight(r, 45);
    [7,8,10,12,14,15].forEach(col => ws.getRange(r,col).setWrap(true).setVerticalAlignment('top'));
  }

  _dropdown(ws, 6,  3, 502, ['Corrida','Academia','Ergométrica','Híbrido']);
  _dropdown(ws, 9,  3, 502, ['Corrida Z1/Z2','Corrida Z3','Intervalado','Progressivo','Fartlek','Força','Mobilidade','Descanso']);
  _dropdown(ws, 11, 3, 502, ['Corrida Z1/Z2','Corrida Z3','Intervalado','Progressivo','Fartlek','Força','Mobilidade','Descanso']);
  _dropdown(ws, 13, 3, 502, ['Corrida Z1/Z2','Corrida Z3','Intervalado','Progressivo','Fartlek','Força','Mobilidade','Descanso']);
  _dropdown(ws, 16, 3, 502, ['Pendente','Em andamento','Concluído','Parcial','Cancelado']);

  _condTexto(ws, 'P3:P502', 'Concluído',   COR.verde_claro);
  _condTexto(ws, 'P3:P502', 'Pendente',    COR.amarelo_cl);
  _condTexto(ws, 'P3:P502', 'Cancelado',   COR.vermelho_cl);
  _condTexto(ws, 'P3:P502', 'Parcial',     '#FFF0E0');

  const larguras = [12,11,18,13,13,13,32,40,14,44,14,44,14,44,38,16];
  larguras.forEach((w, i) => ws.setColumnWidth(i+1, w*7));
  ws.setRowHeight(2, 28);
  ws.setHiddenGridlines(true);
}


// ════════════════════════════════════════════════════════════════════════════
// 5. MÉTRICAS
// ════════════════════════════════════════════════════════════════════════════
function _criarMetricas(ws) {
  ws.setTabColor(COR.verde.replace('#',''));

  _tituloPrincipal(ws, '📈  MÉTRICAS E ZONAS DE TREINO POR ATLETA', 18, COR.verde);

  const cols = [
    'ID Atleta','Nome Atleta','Atualizado em','VO2máx Est.',
    'Pace Médio (s/km)','Pace Rápido (s/km)','Pace Lento (s/km)',
    'FC Máxima','FC Média','Vol./Semana (km)',
    'Z1 Lento','Z1 Rápido','Z2 Lento','Z2 Rápido',
    'Z3 Lento','Z3 Rápido','Z4 Lento','Z5 Mín'
  ];
  _cabecalho(ws, 2, cols, COR.verde);
  _zebra(ws, 3, 202, 18, COR.branco, COR.verde_claro);
  _congelar(ws, 2, 2);

  for (let r = 3; r <= 202; r++) {
    // Z1
    ws.getRange(r,11).setFormula(`=IFERROR(TEXT(INT(E${r}*1.20/60),"0")&":"&TEXT(MOD(ROUND(E${r}*1.20,0),60),"00"),"")`);
    ws.getRange(r,12).setFormula(`=IFERROR(TEXT(INT(E${r}*1.08/60),"0")&":"&TEXT(MOD(ROUND(E${r}*1.08,0),60),"00"),"")`);
    // Z2
    ws.getRange(r,13).setFormula(`=IFERROR(TEXT(INT(E${r}*1.04/60),"0")&":"&TEXT(MOD(ROUND(E${r}*1.04,0),60),"00"),"")`);
    ws.getRange(r,14).setFormula(`=IFERROR(TEXT(INT(E${r}*0.96/60),"0")&":"&TEXT(MOD(ROUND(E${r}*0.96,0),60),"00"),"")`);
    // Z3
    ws.getRange(r,15).setFormula(`=IFERROR(TEXT(INT(E${r}*0.94/60),"0")&":"&TEXT(MOD(ROUND(E${r}*0.94,0),60),"00"),"")`);
    ws.getRange(r,16).setFormula(`=IFERROR(TEXT(INT(E${r}*0.87/60),"0")&":"&TEXT(MOD(ROUND(E${r}*0.87,0),60),"00"),"")`);
    // Z4
    ws.getRange(r,17).setFormula(`=IFERROR(TEXT(INT(E${r}*0.84/60),"0")&":"&TEXT(MOD(ROUND(E${r}*0.84,0),60),"00"),"")`);
    // Z5
    ws.getRange(r,18).setFormula(`=IFERROR(TEXT(INT(F${r}/60),"0")&":"&TEXT(MOD(F${r},60),"00"),"")`);
  }

  // Escala de cor VO2máx
  _escala(ws, 'D3:D202', COR.vermelho, COR.amarelo, COR.verde);

  const larguras = [11,18,14,14,16,16,16,10,10,14,13,13,12,12,12,12,12,12];
  larguras.forEach((w, i) => ws.setColumnWidth(i+1, w*7));
  ws.setRowHeight(2, 28);
  ws.setHiddenGridlines(true);
}


// ════════════════════════════════════════════════════════════════════════════
// 6. FEEDBACK
// ════════════════════════════════════════════════════════════════════════════
function _criarFeedback(ws) {
  ws.setTabColor(COR.amarelo.replace('#',''));

  _tituloPrincipal(ws, '💬  FEEDBACK SEMANAL — CABEÇA, CORAÇÃO E CORPO', 14, COR.amarelo, COR.preto);

  const cols = [
    'ID Feedback','ID Atleta','Nome Atleta','Data Semana',
    'Humor (1-10)','Qualidade do Sono','Fadiga (1-10)',
    'Resultado vs Plano','Treinos Realizados','Treinos Planejados',
    'Alertas Físicos','Feedback Livre','Ajuste Recomendado','Respondido IA'
  ];
  _cabecalho(ws, 2, cols, COR.cinza_escuro);
  _zebra(ws, 3, 202, 14, COR.branco, COR.amarelo_cl);
  _congelar(ws, 2, 3);

  for (let r = 3; r <= 202; r++) {
    ws.setRowHeight(r, 40);
    [12,13].forEach(col => ws.getRange(r,col).setWrap(true).setVerticalAlignment('top'));
  }

  _dropdown(ws, 6,  3, 502, ['Excelente (7h+)','Bom (6-7h)','Regular (5-6h)','Ruim (menos 5h)']);
  _dropdown(ws, 8,  3, 502, ['Acima do esperado','Dentro do plano','Abaixo - carga alta','Abaixo - faltou','Não treinou']);
  _dropdown(ws, 14, 3, 502, ['Sim','Não','Pendente']);

  // Escala humor (1=ruim → 10=ótimo)
  _escala(ws, 'E3:E502', COR.vermelho, COR.amarelo, COR.verde);
  // Escala fadiga invertida (1=ok → 10=esgotado)
  _escala(ws, 'G3:G502', COR.verde, COR.amarelo, COR.vermelho);

  _condTexto(ws, 'N3:N502', 'Sim',      COR.verde_claro);
  _condTexto(ws, 'N3:N502', 'Não',      COR.vermelho_cl);
  _condTexto(ws, 'N3:N502', 'Pendente', COR.amarelo_cl);

  const larguras = [13,11,18,12,11,18,13,22,14,14,24,38,32,14];
  larguras.forEach((w, i) => ws.setColumnWidth(i+1, w*7));
  ws.setRowHeight(2, 28);
  ws.setHiddenGridlines(true);
}


// ════════════════════════════════════════════════════════════════════════════
// 7. GRÁFICOS
// ════════════════════════════════════════════════════════════════════════════
function _criarGraficos(ws) {
  ws.setTabColor(COR.azul_claro.replace('#',''));

  _tituloPrincipal(ws, '📉  GRÁFICOS DE PERFORMANCE — GRUPO HIPERATIVO', 6, COR.azul_escuro);

  ws.getRange('A3:F3').setValues([['PARÂMETRO','VALOR','','','','']]);
  ws.getRange('A3:F3').setFontWeight('bold').setBackground(COR.azul_medio).setFontColor(COR.branco).setHorizontalAlignment('center');

  const metricas = [
    ['Total km (mês atual)',  "=SUMPRODUCT((MONTH('🏃 ATIVIDADES'!D3:D500)=MONTH(TODAY()))*(YEAR('🏃 ATIVIDADES'!D3:D500)=YEAR(TODAY()))*('🏃 ATIVIDADES'!E3:E500=\"Corrida\")*('🏃 ATIVIDADES'!L3:L500))"],
    ['Qtd treinos (mês)',     "=COUNTIFS('🏃 ATIVIDADES'!E3:E500,\"Corrida\",'🏃 ATIVIDADES'!D3:D500,\">=\"&DATE(YEAR(TODAY()),MONTH(TODAY()),1))"],
    ['Pace médio geral',      "=IFERROR(TEXT(INT(AVERAGEIF('🏃 ATIVIDADES'!E3:E500,\"Corrida\",'🏃 ATIVIDADES'!N3:N500)/60),\"0\")&\":\"&TEXT(MOD(ROUND(AVERAGEIF('🏃 ATIVIDADES'!E3:E500,\"Corrida\",'🏃 ATIVIDADES'!N3:N500),0),60),\"00\"),\"--\")"],
    ['FC média geral',        "=IFERROR(ROUND(AVERAGEIF('🏃 ATIVIDADES'!E3:E500,\"Corrida\",'🏃 ATIVIDADES'!Q3:Q500),0),\"--\")"],
    ['VO2máx médio (atletas)','=IFERROR(ROUND(AVERAGE(\'📈 MÉTRICAS\'!D3:D202),1),"--")'],
    ['Atletas sem treinar 14d','=COUNTIF(A12:A21,"⚠️ Verificar")'],
  ];

  metricas.forEach(([label, formula], i) => {
    ws.getRange(4 + i, 1).setValue(label).setFontFamily('Arial').setFontSize(10).setBackground(COR.cinza_claro).setFontWeight('bold');
    ws.getRange(4 + i, 2).setFormula(formula).setFontFamily('Arial').setFontSize(10).setHorizontalAlignment('center');
    ws.setRowHeight(4 + i, 22);
  });

  ws.getRange('A3:B9').setBorder(true,true,true,true,true,true,'#DDDDDD', SpreadsheetApp.BorderStyle.SOLID);
  ws.setColumnWidth(1, 200);
  ws.setColumnWidth(2, 120);
  ws.setRowHeight(2, 28);
  ws.setHiddenGridlines(true);
}


// ════════════════════════════════════════════════════════════════════════════
// 8. ERROS
// ════════════════════════════════════════════════════════════════════════════
function _criarErros(ws) {
  ws.setTabColor(COR.vermelho.replace('#',''));

  _tituloPrincipal(ws, '🔴  LOG DE ERROS E EVENTOS DO SISTEMA', 8, COR.vermelho);

  const cols = ['Data/Hora','Nível','ID Atleta','Função de Origem','Mensagem','Detalhes JSON','Resolvido','Ação Tomada'];
  _cabecalho(ws, 2, cols, COR.cinza_escuro);
  _zebra(ws, 3, 1002, 8, COR.branco, COR.vermelho_cl);
  _congelar(ws, 2, 1);

  for (let r = 3; r <= 1002; r++) {
    [5,6,8].forEach(col => ws.getRange(r,col).setWrap(true));
    ws.setRowHeight(r, 28);
  }

  _dropdown(ws, 2, 3, 1002, ['INFO','AVISO','ERRO','CRÍTICO']);
  _dropdown(ws, 7, 3, 1002, ['Sim','Não','Em análise']);

  _condTexto(ws, 'B3:B1002', 'ERRO',    COR.vermelho_cl);
  _condTexto(ws, 'B3:B1002', 'CRÍTICO', '#FF6B6B');
  _condTexto(ws, 'B3:B1002', 'AVISO',   COR.amarelo_cl);
  _condTexto(ws, 'B3:B1002', 'INFO',    COR.verde_claro);
  _condTexto(ws, 'G3:G1002', 'Sim',     COR.verde_claro);
  _condTexto(ws, 'G3:G1002', 'Não',     COR.vermelho_cl);

  const larguras = [18,10,12,22,42,42,12,36];
  larguras.forEach((w, i) => ws.setColumnWidth(i+1, w*7));
  ws.setRowHeight(2, 28);
  ws.setHiddenGridlines(true);
}


// ════════════════════════════════════════════════════════════════════════════
// 9. TOKENS
// ════════════════════════════════════════════════════════════════════════════
function _criarTokens(ws) {
  ws.setTabColor(COR.cinza_escuro.replace('#',''));

  _tituloPrincipal(ws, '🔐  TOKENS STRAVA — NÃO COMPARTILHAR', 8, COR.cinza_escuro);

  const cols = ['ID Atleta','ID Strava','Access Token','Refresh Token','Expira Em (Unix)','Data Conexão','Últ. Atualização','Status Token'];
  _cabecalho(ws, 2, cols, COR.preto);
  _zebra(ws, 3, 502, 8, COR.branco, COR.cinza_claro);
  _congelar(ws, 2, 2);

  for (let r = 3; r <= 502; r++) {
    ws.getRange(r, 1, 1, 8).setFontFamily('Courier New').setFontSize(9);
    ws.setRowHeight(r, 20);
  }

  _dropdown(ws, 8, 3, 502, ['Válido','Expirado','Revogado','Pendente']);
  _condTexto(ws, 'H3:H502', 'Válido',   COR.verde_claro);
  _condTexto(ws, 'H3:H502', 'Expirado', COR.vermelho_cl);
  _condTexto(ws, 'H3:H502', 'Revogado', '#E8E8E8');

  const larguras = [12,15,46,46,18,16,18,14];
  larguras.forEach((w, i) => ws.setColumnWidth(i+1, w*7));
  ws.setRowHeight(2, 28);
  ws.setHiddenGridlines(true);
}


// ════════════════════════════════════════════════════════════════════════════
// 10. CONFIG
// ════════════════════════════════════════════════════════════════════════════
function _criarConfig(ws) {
  ws.setTabColor(COR.cinza_escuro.replace('#',''));

  _tituloPrincipal(ws, '⚙️  CONFIGURAÇÕES DO SISTEMA HIPERATIVO V3', 4, COR.cinza_escuro);

  _cabecalho(ws, 2, ['Parâmetro','Valor','Descrição',''], COR.cinza_escuro);

  const configs = [
    ['STRAVA_CLIENT_ID',        '',      'ID do aplicativo Strava (strava.com/settings/api)'],
    ['WEBAPP_URL',              '',      'URL do Web App publicado no Apps Script (exec)'],
    ['PIN_TREINADOR',           '1234',  'PIN de acesso ao painel do treinador'],
    ['EMAIL_ADMIN',             '',      'E-mail do administrador para notificações automáticas'],
    ['CARGA_CRITICA_RPE',       '8',     'RPE acima deste valor gera alerta de sobrecarga'],
    ['DIAS_SEM_TREINO_ALERTA',  '14',    'Dias sem atividade para gerar alerta no painel'],
    ['CICLO_SEMANAS',           '12',    'Duração do ciclo de treinamento (semanas)'],
    ['PERCENTUAL_Z2',           '80',    '% do volume semanal em Z1/Z2 (modelo Seiler 80/20)'],
    ['VERSAO_SISTEMA',          '3.0',   'Versão atual — não alterar'],
  ];

  configs.forEach(([param, val, desc], i) => {
    const row = i + 3;
    ws.getRange(row, 1).setValue(param).setFontFamily('Courier New').setFontSize(10)
      .setFontWeight('bold').setFontColor(COR.azul_medio).setBackground(COR.cinza_claro)
      .setHorizontalAlignment('left').setVerticalAlignment('middle');
    ws.getRange(row, 2).setValue(val).setFontFamily('Courier New').setFontSize(10)
      .setBackground('#FFFDE7').setHorizontalAlignment('center').setVerticalAlignment('middle');
    ws.getRange(row, 3).setValue(desc).setFontFamily('Arial').setFontSize(10)
      .setFontStyle('italic').setFontColor('#888888').setBackground(COR.branco)
      .setHorizontalAlignment('left').setVerticalAlignment('middle');
    ws.setRowHeight(row, 26);
  });

  ws.getRange('A2:C11').setBorder(true,true,true,true,true,true,'#DDDDDD', SpreadsheetApp.BorderStyle.SOLID);

  [30,28,50,10].forEach((w, i) => ws.setColumnWidth(i+1, w*7));
  ws.setRowHeight(2, 28);
  ws.setHiddenGridlines(true);
}