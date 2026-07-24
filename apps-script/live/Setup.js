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
  try { ws.setFrozenColumns(colunas); } catch(e) {
    // Ignora erro de célula mesclada no freeze de colunas
  }
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
    ['Total de Atletas',     '=COUNTA(\'👤 CADASTRO\'!B3:B500)'],
    ['Atletas Ativos',       '=COUNTIF(\'👤 CADASTRO\'!Z3:Z500;"Ativo")'],
    ['Strava Conectados',    '=COUNTIF(\'👤 CADASTRO\'!X3:X500;"Sim")'],
    ['Treinos esta semana',  '=COUNTIFS(\'🏃 ATIVIDADES\'!D3:D500;">="&TODAY()-7)'],
    ['Pace médio (corrida)', '=IFERROR(TEXT(AVERAGEIF(\'🏃 ATIVIDADES\'!E3:E500;"Corrida";\'🏃 ATIVIDADES\'!O3:O500)/86400;"[mm]:ss");"--")'],
    ['km médios/sem',        '=IFERROR(ROUND(SUMIF(\'🏃 ATIVIDADES\'!E3:E500;"Corrida";\'🏃 ATIVIDADES\'!L3:L500)/MAX(1;COUNTIF(\'👤 CADASTRO\'!Z3:Z500;"Ativo"));1);"--")'],
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
    ws.getRange(r, 1).setFormula(`=IFERROR(TEXT(LARGE('🏃 ATIVIDADES'!D$3:D$500;${idx});"DD/MM/AAAA");"--")`);
    ws.getRange(r, 2).setFormula(`=IFERROR(INDEX('🏃 ATIVIDADES'!C$3:C$500;MATCH(LARGE('🏃 ATIVIDADES'!D$3:D$500;${idx});'🏃 ATIVIDADES'!D$3:D$500;0));"--")`);
    ws.getRange(r, 3).setFormula(`=IFERROR(INDEX('🏃 ATIVIDADES'!E$3:E$500;MATCH(LARGE('🏃 ATIVIDADES'!D$3:D$500;${idx});'🏃 ATIVIDADES'!D$3:D$500;0));"--")`);
    ws.getRange(r, 4).setFormula(`=IFERROR(ROUND(INDEX('🏃 ATIVIDADES'!L$3:L$500;MATCH(LARGE('🏃 ATIVIDADES'!D$3:D$500;${idx});'🏃 ATIVIDADES'!D$3:D$500;0));2);"--")`);
    ws.getRange(r, 5).setFormula(`=IFERROR(INDEX('🏃 ATIVIDADES'!O$3:O$500;MATCH(LARGE('🏃 ATIVIDADES'!D$3:D$500;${idx});'🏃 ATIVIDADES'!D$3:D$500;0));"--")`);
    ws.getRange(r, 6).setFormula(`=IFERROR(ROUND(INDEX('🏃 ATIVIDADES'!Q$3:Q$500;MATCH(LARGE('🏃 ATIVIDADES'!D$3:D$500;${idx});'🏃 ATIVIDADES'!D$3:D$500;0));0);"--")`);
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
    ws.getRange(r, 8).setFormula(`=IFERROR(INDEX('👤 CADASTRO'!B$3:B$500;${idx});"--")`);
    ws.getRange(r, 9).setFormula(`=IFERROR(TEXT(MAXIFS('🏃 ATIVIDADES'!D$3:D$500;'🏃 ATIVIDADES'!C$3:C$500;H${r});"DD/MM/AA");"Nunca")`);
    ws.getRange(r, 10).setFormula(`=IFERROR(TODAY()-MAXIFS('🏃 ATIVIDADES'!D$3:D$500;'🏃 ATIVIDADES'!C$3:C$500;H${r});"--")`);
    ws.getRange(r, 11).setFormula(`=IFERROR(IF(J${r}>14;"⚠️ Verificar";"✅ OK");"--")`);
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
// 2. CADASTRO — 40 colunas (v4.1 — alinhado com H.CAD expandido)
// Grupos visuais em 3 linhas de cabeçalho:
//   Linha 1 = Título principal (mesclado)
//   Linha 2 = Labels de grupo (coloridos, mesclados por grupo)
//   Linha 3 = Nomes de coluna
// ════════════════════════════════════════════════════════════════════════════

const _CAD_GRUPOS = [
  { label: '👤  IDENTIFICAÇÃO',      de: 1,  ate: 8,  cor: '#001F3F', txt: '#FFFFFF' },
  { label: '🏃  PERFIL ESPORTIVO',   de: 9,  ate: 13, cor: '#1D9E75', txt: '#FFFFFF' },
  { label: '❤️  SAÚDE & HISTÓRICO',  de: 14, ate: 16, cor: '#E24B4A', txt: '#FFFFFF' },
  { label: '⚡  PROGRAMA & PROVAS',  de: 17, ate: 18, cor: '#FC4C02', txt: '#FFFFFF' },
  { label: '📍  LOCALIZAÇÃO & CRM',  de: 19, ate: 23, cor: '#6C3DC7', txt: '#FFFFFF' },
  { label: '🔗  STRAVA & SISTEMA',   de: 24, ate: 27, cor: '#003D7A', txt: '#FFFFFF' },
  { label: '🚨  EMERGÊNCIA',         de: 28, ate: 30, cor: '#B71C1C', txt: '#FFFFFF' },
  { label: '🩺  PAR-Q & MELHORES TEMPOS', de: 31, ate: 32, cor: '#F57F17', txt: '#000000' },
  { label: '✍️  ASSINATURA DIGITAL', de: 33, ate: 34, cor: '#37474F', txt: '#FFFFFF' },
  { label: '📊  CRM AVANÇADO',       de: 35, ate: 40, cor: '#4A148C', txt: '#FFFFFF' },
];

const _CAD_COLS = [
  // ── Identificação (1-8) ──────────────────────────────────────────────────
  '🆔 ID Atleta',   '📛 Nome Completo',  '📧 E-mail',      '📱 WhatsApp',
  '🎂 Data Nasc.',  '⚧ Sexo',            '⚖️ Peso (kg)',   '📏 Altura (cm)',
  // ── Perfil Esportivo (9-13) ──────────────────────────────────────────────
  '🏅 Modalidade(s)', '🎯 Nível',         '🏆 Objetivo(s)', '📅 Freq./Sem.', '🕐 Horário',
  // ── Saúde (14-16) ────────────────────────────────────────────────────────
  '🏥 Condições de Saúde', '🦴 Lesões/Limitações', '💊 Medicamentos',
  // ── Programa & Provas (17-18) ────────────────────────────────────────────
  '🏁 Provas Realizadas',  '⚡ Plano/Programa',
  // ── Localização & CRM (19-23) ────────────────────────────────────────────
  '🏙️ Cidade', '🗺️ Estado', '🪪 CPF', '📣 Canal de Origem', '📆 Data Cadastro',
  // ── Strava & Sistema (24-27) ─────────────────────────────────────────────
  '🔗 Strava',      '🏃 ID Strava',      '🔄 Status',  '💬 Obs / Link WA',
  // ── Emergência (28-30) ───────────────────────────────────────────────────
  '🚨 Emerg. Nome', '📞 Emerg. Tel.',    '👨‍👩‍👧 Parentesco',
  // ── PAR-Q & PRs (31-32) ─────────────────────────────────────────────────
  '🩺 PAR-Q (7 resp.)', '⏱️ PRs (Melhores Tempos)',
  // ── Assinatura (33-34) ───────────────────────────────────────────────────
  '✍️ Assinatura Digital', '📅 Data Assinatura',
  // ── CRM Avançado (35-40) — preenchimento manual ───────────────────────────
  '📸 Instagram',   '🏁 Próxima Prova',  '📆 Data Prova', '💳 Plano Pgto.', '🗓️ Início Plano', '📋 Últ. Avaliação',
];

function _aplicarCabecalhoCadastro(ws, numCols) {
  numCols = numCols || 40;

  // ── Linha 1: Título mesclado ─────────────────────────────────────────────
  ws.getRange(1, 1, 1, numCols).mergeAcross()
    .setValue('👤  CADASTRO DE ATLETAS — GRUPO HIPERATIVO  |  CABEÇA • CORAÇÃO • CORPO')
    .setFontFamily('Arial').setFontSize(13).setFontWeight('bold')
    .setFontColor('#FFFFFF').setBackground('#001F3F')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  ws.setRowHeight(1, 38);

  // ── Linha 2: Grupos coloridos ─────────────────────────────────────────────
  _CAD_GRUPOS.forEach(g => {
    const largura = g.ate - g.de + 1;
    const rg = ws.getRange(2, g.de, 1, largura);
    if (largura > 1) rg.mergeAcross();
    rg.setValue(g.label)
      .setFontFamily('Arial').setFontSize(9).setFontWeight('bold')
      .setFontColor(g.txt).setBackground(g.cor)
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
  });
  ws.setRowHeight(2, 22);

  // ── Linha 3: Nomes de coluna ──────────────────────────────────────────────
  ws.getRange(3, 1, 1, _CAD_COLS.length).setValues([_CAD_COLS])
    .setFontFamily('Arial').setFontSize(9).setFontWeight('bold')
    .setFontColor('#FFFFFF').setBackground('#2C2C2C')
    .setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setWrap(true);
  ws.setRowHeight(3, 30);
}

function _aplicarDropdownsCadastro(ws) {
  // Campos de seleção única
  _dropdown(ws,  6, 4, 500, ['Masculino','Feminino','Prefiro não informar']); // F=SEXO
  _dropdown(ws, 10, 4, 500, ['Iniciante','Intermediário','Avançado','Elite']); // J=NIVEL
  _dropdown(ws, 12, 4, 500, ['1-2x por semana','3-4x por semana','5-6x por semana','Todos os dias']); // L=FREQ
  _dropdown(ws, 13, 4, 500, ['Manhã (5h–9h)','Manhã (9h–12h)','Tarde (12h–17h)','Noite (17h–21h)','Noite (21h+)','Variável']); // M=HORARIO
  _dropdown(ws, 18, 4, 500, ['Alta Voltagem','Corrida CCC','Iniciante em Movimento','5k10k em 6 Semanas','Hiperativo Running Club','Vida Ativa Melhor Idade']); // R=PLANO
  _dropdown(ws, 24, 4, 500, ['Sim','Não','Pendente','Reconectar']); // X=STRAVA_OK
  _dropdown(ws, 26, 4, 500, ['Ativo','Inativo','Suspenso','Trial','Cancelado']); // Z=STATUS
  _dropdown(ws, 30, 4, 500, ['Cônjuge / Companheiro(a)','Pai / Mãe','Filho(a)','Irmão / Irmã','Amigo(a)','Outro']); // AD=EMERG_REL
  _dropdown(ws, 38, 4, 500, ['Mensal','Trimestral','Semestral','Anual','Cortesia']); // AL=PLANO_PAG

  // Campos com múltiplos valores possíveis — lista para referência, aceita texto livre
  const multiHint = (ws, col, opcoes) => {
    const rg = ws.getRange(4, col, 497, 1);
    const regra = SpreadsheetApp.newDataValidation()
      .requireValueInList(opcoes, true)
      .setAllowInvalid(true)   // permite digitar combinações
      .setHelpText('Selecione ou digite valores separados por vírgula')
      .build();
    rg.setDataValidation(regra);
  };
  multiHint(ws,  9, ['Corrida de rua','Ciclismo','Natação','Triathlon','Caminhada','Musculação','Funcional','Trail Running','Outra']); // I=MOD
  multiHint(ws, 11, ['Emagrecer / Reduzir gordura','Ganhar massa muscular','Melhorar desempenho','Completar primeira prova','Saúde e bem-estar','Bater recorde pessoal (PR)','Reabilitação / Retorno ao esporte']); // K=OBJ
  multiHint(ws, 17, ['Nunca participei','5k','10k','Meia Maratona','Maratona','Triathlon','Ultra','Trail']); // Q=PROVAS
  multiHint(ws, 22, ['Instagram','Facebook','Indicação de amigo','Google','WhatsApp','Evento esportivo','Strava','YouTube','TikTok','Outro']); // V=ORIGEM
}

function _aplicarFormatacaoCadastroCond(ws) {
  ws.clearConditionalFormatRules();
  // Status (col 26 = Z)
  _condTexto(ws, 'Z4:Z500', 'Ativo',     COR.verde_claro);
  _condTexto(ws, 'Z4:Z500', 'Trial',      COR.amarelo_cl);
  _condTexto(ws, 'Z4:Z500', 'Inativo',   COR.vermelho_cl);
  _condTexto(ws, 'Z4:Z500', 'Suspenso',  '#E8E8E8');
  _condTexto(ws, 'Z4:Z500', 'Cancelado', '#FFCCCC');
  // Strava (col 24 = X)
  _condTexto(ws, 'X4:X500', 'Sim',        COR.verde_claro);
  _condTexto(ws, 'X4:X500', 'Pendente',   COR.amarelo_cl);
  _condTexto(ws, 'X4:X500', 'Reconectar', '#FFE0B2');
  _condTexto(ws, 'X4:X500', 'Não',        COR.vermelho_cl);
  // Conexão Strava — estado operacional real (col 50 = AX)
  _condTexto(ws, 'AX4:AX500', 'Conectado',          COR.verde_claro);
  _condTexto(ws, 'AX4:AX500', 'Aguardando conexão',  COR.amarelo_cl);
  _condTexto(ws, 'AX4:AX500', 'Não utiliza',         '#E8E8E8');
  _condTexto(ws, 'AX4:AX500', 'Reconectar',          COR.vermelho_cl);
  // PAR-Q alerta (col 31 = AE): se contiver S → destaque
  const regraParq = SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains('S')
    .setBackground('#FFF8E1').setFontColor('#E65100')
    .setRanges([ws.getRange('AE4:AE500')])
    .build();
  const regras = ws.getConditionalFormatRules();
  regras.push(regraParq);
  ws.setConditionalFormatRules(regras);
}

const _CAD_LARGURAS = [
  // Identificação (1-8)
  14, 26, 26, 18, 12, 10, 9, 9,
  // Esportivo (9-13)
  20, 12, 22, 12, 14,
  // Saúde (14-16)
  22, 22, 18,
  // Programa & Provas (17-18)
  22, 20,
  // Localização & CRM (19-23)
  14, 9, 14, 16, 16,
  // Strava & Sistema (24-27)
  12, 14, 10, 28,
  // Emergência (28-30)
  20, 16, 16,
  // PAR-Q & PRs (31-32)
  18, 22,
  // Assinatura (33-34)
  22, 16,
  // CRM Avançado (35-40)
  16, 22, 14, 14, 14, 16,
];

function _criarCadastro(ws) {
  ws.setTabColor(COR.verde.replace('#',''));
  while (ws.getMaxColumns() < 40) ws.insertColumnAfter(ws.getMaxColumns());

  _aplicarCabecalhoCadastro(ws, 40);
  _zebra(ws, 4, 203, 40, COR.branco, COR.azul_suave);
  _aplicarDropdownsCadastro(ws);
  _aplicarFormatacaoCadastroCond(ws);

  _CAD_LARGURAS.forEach((w, i) => ws.setColumnWidth(i + 1, w * 7));
  _congelar(ws, 3, 2);
  ws.setHiddenGridlines(true);
}


// ════════════════════════════════════════════════════════════════════════════
// REFORMATAR ABA CADASTRO EXISTENTE SEM APAGAR DADOS (v4.1 — 40 colunas)
// Menu: ⚙️ Reformatar cabeçalho do CADASTRO
// ════════════════════════════════════════════════════════════════════════════
function reformatarCabecalhoCadastro() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let ws   = ss.getSheetByName(H.SHEETS.CADASTRO);
  if (!ws) {
    SpreadsheetApp.getUi().alert('Aba CADASTRO não encontrada.');
    return;
  }

  // Garantir 40 colunas
  while (ws.getMaxColumns() < 40) ws.insertColumnAfter(ws.getMaxColumns());

  // Inserir linhas de cabeçalho se necessário (migração de 27-col → 40-col)
  const linha1Val = ws.getRange(1, 1).getValue();
  const linha2Val = ws.getRange(2, 1).getValue();

  // Se linha 2 contém dados de atleta (não é cabeçalho de grupo), inserir linha extra
  if (linha2Val && !String(linha2Val).includes('IDENTIFICAÇÃO') && !String(linha2Val).includes('Atleta')) {
    ws.insertRowBefore(2);
    ws.insertRowBefore(2);
  } else if (linha2Val && String(linha2Val).includes('Atleta')) {
    // Estrutura antiga (2 linhas): inserir 1 linha de grupos
    ws.insertRowBefore(2);
  }

  _aplicarCabecalhoCadastro(ws, 40);
  _zebra(ws, 4, 203, 40, COR.branco, COR.azul_suave);
  _aplicarDropdownsCadastro(ws);
  _aplicarFormatacaoCadastroCond(ws);
  _CAD_LARGURAS.forEach((w, i) => ws.setColumnWidth(i + 1, w * 7));
  _congelar(ws, 3, 2);
  ws.setHiddenGridlines(true);
  ws.setRowHeight(3, 30);

  try {
    SpreadsheetApp.getUi().alert(
      '✅ CADASTRO reformatado com 40 colunas!',
      'Novos grupos adicionados:\n' +
      '• Emergência (cols 28-30)\n' +
      '• PAR-Q & Melhores Tempos (cols 31-32)\n' +
      '• Assinatura Digital (cols 33-34)\n' +
      '• CRM Avançado — preenchimento manual (cols 35-40)\n\n' +
      'Dados existentes preservados.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch(_) {}
}


/**
 * Aplica a cor condicional da Conexão Strava (AX) e reafirma o
 * congelamento de cabeçalho + coluna Nome no CADASTRO, sem tocar em
 * dados, dropdowns ou zebra já existentes. Seguro para rodar quantas
 * vezes quiser — só reaplica regras de formatação.
 * Menu: ⚡ HIPERATIVO ▸ 🏃 Atividades ▸ Congelar cabeçalho/Nome + cor Conexão Strava
 */
function aplicarCorConexaoStrava() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ws = ss.getSheetByName(H.SHEETS.CADASTRO);
  if (!ws) { SpreadsheetApp.getUi().alert('Aba CADASTRO não encontrada.'); return; }
  _aplicarFormatacaoCadastroCond(ws);
  _congelar(ws, 3, 2);
  try {
    SpreadsheetApp.getUi().alert(
      '✅ Cor da Conexão Strava aplicada',
      'Coluna AX (Conexão Strava) agora tem cor por status.\n' +
      'Cabeçalho (3 linhas) e colunas ID+Nome permanecem congelados ao rolar.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch(_) {}
}

// COMPAT: manter _criarCadastroLegacy para referência — não usado no setup principal
function _criarCadastroLegado_27cols(ws) {
  // versão antiga preservada como fallback
  _tituloPrincipal(ws, '👤  CADASTRO DE ATLETAS — GRUPO HIPERATIVO', 27, COR.azul_escuro);
  const cols = [
    'ID Atleta','Nome Completo','E-mail','WhatsApp',
    'Data Nasc.','Sexo','Peso (kg)','Altura (cm)',
    'Modalidade','Nível','Objetivo','Freq./Semana','Horário Pref.',
    'Condições de Saúde','Lesões/Limitações','Medicamentos',
    'Provas Anteriores','Plano/Programa',
    // Col 19–21: Localização e CPF
    'Cidade','Estado','CPF',
    // Col 22–23: CRM
    'Canal de Origem','Data Cadastro',
    // Col 24–26: Strava / Sistema
    'Strava Conectado','ID Strava','Status',
    // Col 27: Obs
    'Observações / Link WA'
  ];
  _cabecalho(ws, 2, cols, COR.azul_medio);
  _zebra(ws, 3, 202, 27, COR.branco, COR.azul_suave);
  _congelar(ws, 2, 2);

  // Dropdowns alinhados com H.CAD
  _dropdown(ws,  6, 3, 500, ['Masculino','Feminino','Prefiro não informar']);          // F=SEXO
  _dropdown(ws,  9, 3, 500, ['Corrida de rua','Ciclismo','Natação','Triathlon',        // I=MOD
    'Caminhada','Musculação','Funcional','Trail Running','Outra']);
  _dropdown(ws, 10, 3, 500, ['Iniciante','Intermediário','Avançado','Elite']);          // J=NIVEL
  _dropdown(ws, 11, 3, 500, ['Emagrecer / Reduzir gordura','Ganhar massa muscular',    // K=OBJ
    'Melhorar desempenho','Completar primeira prova','Saúde e bem-estar',
    'Bater recorde pessoal (PR)','Reabilitação / Retorno ao esporte']);
  _dropdown(ws, 12, 3, 500, ['1-2x por semana','3-4x por semana',                     // L=FREQ
    '5-6x por semana','Todos os dias']);
  _dropdown(ws, 13, 3, 500, ['Manhã (5h–9h)','Manhã (9h–12h)',                        // M=HORARIO
    'Tarde (12h–17h)','Noite (17h–21h)','Noite (21h+)','Variável']);
  _dropdown(ws, 17, 3, 500, ['Nunca participei','Sim, 5k','Sim, 10k',                  // Q=PROVA
    'Sim, meia maratona','Sim, maratona','Sim, triathlon','Sim, outras distâncias']);
  _dropdown(ws, 18, 3, 500, ['Alta Voltagem','Corrida CCC','Iniciante em Movimento',   // R=PLANO
    '5k10k em 6 Semanas','Hiperativo Running Club','Vida Ativa Melhor Idade']);
  _dropdown(ws, 22, 3, 500, ['Instagram','Facebook','Indicação de amigo','Google',     // V=ORIGEM
    'WhatsApp','Evento esportivo','Strava','Outro']);
  _dropdown(ws, 24, 3, 500, ['Sim','Não','Pendente']);                                  // X=STRAVA_OK
  _dropdown(ws, 26, 3, 500, ['Ativo','Inativo','Suspenso','Trial']);                    // Z=STATUS

  // bloco legado encerrado — novo setup usa _criarCadastro com 40 colunas
}


// ════════════════════════════════════════════════════════════════════════════
// 3. ATIVIDADES
// ════════════════════════════════════════════════════════════════════════════
function _criarAtividades(ws) {
  ws.setTabColor(COR.laranja.replace('#',''));

  _tituloPrincipal(ws, '🏃  ATIVIDADES — DADOS IMPORTADOS DO STRAVA', 22, COR.laranja);

  const cols = [
    'ID Exec.','Atleta ID','Atleta','Data','Tipo',
    'Fonte','Strava ID','Atividade','Tempo Mov.',
    'Tempo Total','Dist. (m)','Dist. (km)',
    'Vel. m/s','Vel. km/min','Pace s/km','⚡ Ritmo/Vel.',
    'FC Méd ♥','FC Máx ♥','Elevação m','Kcal',
    'Cadência','Potência W','Rota','Importado','PSE'
  ];
  _cabecalho(ws, 2, cols, COR.cinza_escuro);
  _zebra(ws, 3, 502, 25, COR.branco, '#FFF5EE');
  _congelar(ws, 2, 4);

  // Formato de tempo [h]:mm:ss nas colunas Tempo Mov. (9) e Tempo Total (10)
  const fmtTempo = SpreadsheetApp.newDataValidation ? null : null; // apenas formato
  ws.getRange(3, 9, 500, 1).setNumberFormat('[h]:mm:ss'); // MOV_S
  ws.getRange(3, 10, 500, 1).setNumberFormat('[h]:mm:ss'); // TOTAL_S

  // Fórmulas derivadas para entradas manuais (dados Strava já vêm preenchidos)
  for (let r = 3; r <= 502; r++) {
    // Col 12: Dist. km — 2 dec corrida/trail, 1 dec ciclismo, 3 dec natação
    ws.getRange(r, 12).setFormula(
      `=IFERROR(IF(K${r}<=0;"";` +
      `IF(OR(E${r}="Ciclismo",E${r}="Ergométrica",E${r}="Híbrido"),ROUND(K${r}/1000,1),` +
      `IF(OR(E${r}="Natação"),ROUND(K${r}/1000,3),` +
      `ROUND(K${r}/1000,2)))),"")`
    );
    // Col 14: Vel. km/min (interno, para cálculos)
    ws.getRange(r, 14).setFormula(`=IFERROR(IF(M${r}>0;ROUND(M${r}*0.06;3);"");"")`);
    // Col 15: Pace numérico — s/km (corrida) | s/100m (natação) | 0 (ciclismo)
    ws.getRange(r, 15).setFormula(
      `=IFERROR(IF(M${r}<=0;"";` +
      `IF(OR(E${r}="Ciclismo",E${r}="Ergométrica",E${r}="Híbrido"),0,` +
      `IF(OR(E${r}="Natação"),ROUND(100/M${r},0),` +
      `ROUND(1000/M${r},0)))),"")`
    );
    // Col 16: ⚡ Ritmo/Vel. — display inteligente por esporte
    ws.getRange(r, 16).setFormula(
      `=IFERROR(IF(M${r}<=0;"";` +
      // Ciclismo → km/h com 1 decimal
      `IF(OR(E${r}="Ciclismo",E${r}="Ergométrica",E${r}="Híbrido"),` +
      `TEXT(M${r}*3.6,"0.0")&" km/h",` +
      // Natação → min:ss /100m
      `IF(OR(E${r}="Natação"),` +
      `INT(ROUND(100/M${r},0)/60)&":"&TEXT(MOD(ROUND(100/M${r},0),60),"00")&" /100m",` +
      // Corrida / Caminhada / Trail → min:ss /km
      `INT(ROUND(1000/M${r},0)/60)&":"&TEXT(MOD(ROUND(1000/M${r},0),60),"00")&" /km"` +
      `))),"")`
    );
  }

  // Dropdowns
  _dropdown(ws, 5,  3, 502, ['Corrida','Caminhada','Ciclismo','Natação','Trail Run','Musculação','Funcional','Outro']);
  _dropdown(ws, 6,  3, 502, ['Strava','Manual','App']);

  // Formatação condicional Tipo
  _condTexto(ws, 'E3:E502', 'Corrida',   '#D6EEFF');
  _condTexto(ws, 'E3:E502', 'Ciclismo',  COR.verde_claro);
  _condTexto(ws, 'E3:E502', 'Caminhada', COR.amarelo_cl);
  _condTexto(ws, 'E3:E502', 'Trail Run', '#F0E6FF');

  // PSE dropdown 1-10 com descrições
  _dropdown(ws, 25, 3, 502, [
    '1 — Repouso','2 — Muito leve','3 — Leve','4 — Moderado leve',
    '5 — Moderado','6 — Moderado intenso','7 — Intenso',
    '8 — Muito intenso','9 — Extremamente intenso','10 — Máximo'
  ]);
  // Formatação condicional PSE por cor (verde→amarelo→vermelho)
  const pse = ws.getRange('Y3:Y502');
  [[['1','2','3'],'#D6F5EC'],[['4','5','6'],'#FFF8E1'],[['7','8'],'#FFE0B2'],[['9','10'],'#FDECEA']].forEach(([vals, bg]) => {
    vals.forEach(v => _condTexto(ws, 'Y3:Y502', v.split(' ')[0] + ' —', bg));
  });
  const larguras = [11,9,16,10,11,7,11,20,10,10,9,9,8,9,9,13,8,8,8,7,9,8,18,12,10];
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
  _garantirEstruturaMetricas_(ws);
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
  _garantirEstruturaFeedbackEstudos_();
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
    ['Qtd treinos (mês)',     "=COUNTIFS('🏃 ATIVIDADES'!E3:E500;\"Corrida\";'🏃 ATIVIDADES'!D3:D500;\">=\"&DATE(YEAR(TODAY());MONTH(TODAY());1))"],
    ['Pace médio geral',      "=IFERROR(TEXT(INT(AVERAGEIF('🏃 ATIVIDADES'!E3:E500;\"Corrida\";'🏃 ATIVIDADES'!N3:N500)/60);\"0\")&\":\"&TEXT(MOD(ROUND(AVERAGEIF('🏃 ATIVIDADES'!E3:E500;\"Corrida\";'🏃 ATIVIDADES'!N3:N500);0);60);\"00\");\"--\")"],
    ['FC média geral',        "=IFERROR(ROUND(AVERAGEIF('🏃 ATIVIDADES'!E3:E500;\"Corrida\";'🏃 ATIVIDADES'!Q3:Q500);0);\"--\")"],
    ['VO₂ com teste validado (média)','=IFERROR(ROUND(AVERAGE(\'📈 MÉTRICAS\'!D3:D202);1);"--")'],
    ['Atletas sem treinar 14d','=COUNTIF(A12:A21;"⚠️ Verificar")'],
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
