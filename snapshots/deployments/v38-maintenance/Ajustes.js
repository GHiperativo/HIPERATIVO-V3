/**
 * ═══════════════════════════════════════════════════════════════════════
 * HIPERATIVO V3 — Ajustes.gs
 * Correções e ajustes estruturais da planilha (checklist COWORK-001 a 014)
 * Execute: ajustesPlanilha() no Apps Script editor
 * ═══════════════════════════════════════════════════════════════════════
 */

// ── [001] Limpar linhas vazias com #ERROR! na aba ATIVIDADES ──────────────────
/**
 * Deleta todas as linhas da aba ATIVIDADES onde a coluna ATH_ID (col B) está vazia.
 * Essas são linhas pré-seeded de fórmula sem dados reais.
 * Mantém apenas o header + linhas com dados.
 */
function limparAtividades() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(H.SHEETS.ATIVIDADES);
  if (!sh) return 'ATIVIDADES: aba não encontrada.';

  const lastRow = sh.getLastRow();
  if (lastRow <= 1) return 'ATIVIDADES: nenhuma linha de dados.';

  // Lê coluna B (ATH_ID) de todas as linhas de dados
  const athIds = sh.getRange(2, 2, lastRow - 1, 1).getValues();
  const toDelete = [];

  for (let i = athIds.length - 1; i >= 0; i--) {
    const val = String(athIds[i][0] || '').trim();
    if (!val) toDelete.push(i + 2); // +2 porque começa na linha 2
  }

  let deletadas = 0;
  for (const row of toDelete) {
    sh.deleteRow(row);
    deletadas++;
  }

  SpreadsheetApp.flush();
  _log('SYSTEM', 'INFO', 'limparAtividades', deletadas + ' linhas vazias removidas de ATIVIDADES.', '');
  return 'ATIVIDADES: ' + deletadas + ' linhas vazias removidas.';
}

// ── [002 + 003] Corrigir PAINEL: alertas de data e contador de ativos ─────────
/**
 * Varre o PAINEL procurando:
 * (a) células que mostram "46187" (serial 0 interpretado como 30/12/1900) → limpa
 * (b) a seção de alertas "Sem Treinar há mais de 14 dias" → protege contra blank
 *
 * Abordagem: corrigir via fórmulas nas células problemáticas identificadas,
 * ou simplesmente limpar células que mostram o valor serial inválido.
 *
 * Para o contador de Atletas Ativos: ajustar a fórmula para incluir Trial.
 */
function corrigirPainel() {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const sh  = ss.getSheetByName(H.SHEETS.PAINEL);
  if (!sh) return 'PAINEL: aba não encontrada.';

  const data   = sh.getDataRange().getValues();
  const formulas = sh.getDataRange().getFormulas();
  const log    = [];
  let fixes    = 0;

  for (let r = 0; r < data.length; r++) {
    for (let c = 0; c < data[r].length; c++) {
      const val = data[r][c];
      const formula = formulas[r][c] || '';

      // (a) Valor numérico 46187 em células que deveriam ser datas → célula mostrando
      //     data serial 0 (30/12/1900) porque fonte está vazia
      if (typeof val === 'number' && val === 46187 && !formula) {
        sh.getRange(r + 1, c + 1).clearContent();
        fixes++;
      }

      // (b) Fórmulas que calculam "HOJE() - célula > 14" sem checar blank
      //     → substituir por versão segura
      if (formula && formula.includes('HOJE()') && formula.includes('>14') && !formula.includes('ISBLANK')) {
        // Extrai a referência de célula de data da fórmula original
        // e reconstrói com ISBLANK guard
        const safeFormula = formula.replace(
          /=IF\(HOJE\(\)\s*-\s*([^>]+)>14/gi,
          '=IF(ISBLANK($1),"",IF(HOJE()-$1>14'
        ).replace(/("✅[^"]*")\)(\s*)$/, '$1),"")$2');
        // Se a substituição falhou (fórmula muito complexa), apenas adiciona IFERROR
        const finalFormula = safeFormula.startsWith('=IFERROR') ? safeFormula : '=IFERROR(' + formula.substring(1) + ',"")';
        sh.getRange(r + 1, c + 1).setFormula(finalFormula);
        fixes++;
        log.push('Alerta de data corrigido: linha ' + (r+1) + ' col ' + (c+1));
      }
    }
  }

  SpreadsheetApp.flush();
  _log('SYSTEM', 'INFO', 'corrigirPainel', 'PAINEL: ' + fixes + ' correções aplicadas.', log.join(' | '));
  return 'PAINEL: ' + fixes + ' correções aplicadas.';
}

// ── [003] Corrigir COUNTIF de Atletas Ativos para incluir Trial ────────────────
/**
 * Procura no PAINEL fórmulas COUNTIF que contam apenas "Ativo" na coluna de status
 * do CADASTRO e substitui por versão que inclui "Trial" e "Reativado".
 */
function corrigirContadorAtivos() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(H.SHEETS.PAINEL);
  if (!sh) return 'PAINEL: aba não encontrada.';

  const formulas = sh.getDataRange().getFormulas();
  let fixes = 0;

  for (let r = 0; r < formulas.length; r++) {
    for (let c = 0; c < formulas[r].length; c++) {
      const f = formulas[r][c];
      // Detecta COUNTIF/COUNTIFS que filtram por "Ativo" no CADASTRO
      if (f && f.toUpperCase().includes('COUNTIF') &&
          f.includes('CADASTRO') && f.includes('"Ativo"') &&
          !f.includes('"Trial"')) {

        // Substitui COUNTIF simples por COUNTIFS multi-status
        // =COUNTIF(CADASTRO!Z:Z,"Ativo") → =COUNTIFS(CADASTRO!Z:Z,"Ativo")+COUNTIFS(CADASTRO!Z:Z,"Trial")+COUNTIFS(CADASTRO!Z:Z,"Reativado")
        const newFormula = f.replace(
          /COUNTIF\(([^,]+),"Ativo"\)/gi,
          'COUNTIFS($1,"Ativo")+COUNTIFS($1,"Trial")+COUNTIFS($1,"Reativado")'
        );
        if (newFormula !== f) {
          sh.getRange(r + 1, c + 1).setFormula(newFormula);
          fixes++;
        }
      }
    }
  }

  SpreadsheetApp.flush();
  _log('SYSTEM', 'INFO', 'corrigirContadorAtivos', fixes + ' fórmulas de contador ajustadas.', '');
  return 'PAINEL contador ativos: ' + fixes + ' fórmula(s) ajustada(s) (Trial + Reativado incluídos).';
}

// ── [004] Popular aba CONFIG com parâmetros globais ───────────────────────────
function popularConfig() {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const sh  = ss.getSheetByName(H.SHEETS.CONFIG);
  if (!sh) return 'CONFIG: aba não encontrada.';

  // Só popula se estiver vazia
  if (sh.getLastRow() > 1) return 'CONFIG: já possui dados — pulado.';

  const params = [
    ['CHAVE',                            'VALOR',                                        'DESCRIÇÃO'],
    ['NOME_GRUPO',                       'Grupo Hiperativo',                             'Nome do grupo/academia'],
    ['EMAIL_TREINADOR',                  'contato@ghiperativo.com.br',                   'E-mail principal do treinador'],
    ['CIDADE',                           'Brasília',                                     'Cidade sede'],
    ['ESTADO',                           'DF',                                           'Estado'],
    ['DIAS_ALERTA_INATIVIDADE',          '14',                                           'Dias sem treino para disparar alerta'],
    ['DIAS_ALERTA_MENSALIDADE',          '5',                                            'Dias antes do vencimento para alertar'],
    ['DIAS_ALERTA_AVALIACAO',            '90',                                           'Dias sem avaliação física para alertar'],
    ['HORAS_ALERTA_FEEDBACK',            '48',                                           'Horas sem feedback visto para alertar'],
    ['URL_WEBHOOK_STRAVA',               '',                                             'URL do webhook Strava (preencher após deploy)'],
    ['STRAVA_CLIENT_ID',                 '',                                             'Client ID do app Strava'],
    ['STRAVA_VERIFY_TOKEN',              '',                                             'Token de verificação do webhook Strava'],
    ['SUPABASE_URL',                     'https://korlpbclqgmqvpbrungc.supabase.co',    'URL do projeto Supabase'],
    ['SUPABASE_KEY',                     '(configurar via menu Credenciais)',             'Service role key — NÃO inserir aqui, usar PropertiesService'],
    ['VERSAO_SISTEMA',                   'V3.1.0',                                       'Versão atual do HIPERATIVO'],
    ['DATA_ULTIMA_ATUALIZACAO',          new Date().toLocaleDateString('pt-BR'),          'Data do último ajuste estrutural'],
    ['SLOGAN',                           'Cabeça • Coração • Corpo',                     'Slogan do sistema'],
    ['STATUS_ATIVOS_VALIDOS',            'Ativo,Trial,Reativado',                        'Status considerados como atleta ativo'],
  ];

  sh.clearContents();
  sh.getRange(1, 1, params.length, 3).setValues(params);

  // Formatar header
  const header = sh.getRange(1, 1, 1, 3);
  header.setBackground('#1a1a2e').setFontColor('#ffffff').setFontWeight('bold');

  // Largura das colunas
  sh.setColumnWidth(1, 240);
  sh.setColumnWidth(2, 380);
  sh.setColumnWidth(3, 320);

  // Linhas alternadas
  for (let i = 2; i <= params.length; i++) {
    if (i % 2 === 0) sh.getRange(i, 1, 1, 3).setBackground('#f8f9fa');
  }

  // Destacar campos que precisam ser preenchidos
  for (let i = 2; i <= params.length; i++) {
    const val = sh.getRange(i, 2).getValue();
    if (!val || String(val).trim() === '') {
      sh.getRange(i, 2).setBackground('#fff3cd'); // Amarelo: campo vazio
    }
  }

  sh.setFrozenRows(1);
  SpreadsheetApp.flush();
  _log('SYSTEM', 'INFO', 'popularConfig', 'CONFIG populada com ' + (params.length - 1) + ' parâmetros.', '');
  return 'CONFIG: ' + (params.length - 1) + ' parâmetros configurados.';
}

// ── [TOKENS] Limpar linhas duplicadas/corrompidas no TOKENS ──────────────────
function limparTokensCorrompidos() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(H.SHEETS.TOKENS);
  if (!sh) return 'TOKENS: aba não encontrada.';

  const lastRow = sh.getLastRow();
  if (lastRow <= 1) return 'TOKENS: vazio.';

  const data = sh.getRange(2, 1, lastRow - 1, sh.getLastColumn()).getValues();
  const toDelete = [];

  // Deleta linhas onde col A (EXEC_ID) contém texto de header ou é "TOK_ID Atleta" etc.
  const headerPatterns = ['TOK_ID', 'EXEC_ID', 'ID ATLETA', 'NOME COMPLETO', 'ID Atleta'];
  for (let i = data.length - 1; i >= 0; i--) {
    const colA = String(data[i][0] || '').trim().toUpperCase();
    const isHeader = headerPatterns.some(p => colA.includes(p.toUpperCase()));
    if (isHeader) toDelete.push(i + 2);
  }

  for (const row of toDelete) sh.deleteRow(row);

  SpreadsheetApp.flush();
  _log('SYSTEM', 'INFO', 'limparTokensCorrompidos', toDelete.length + ' linha(s) corrompida(s) removidas de TOKENS.', '');
  return 'TOKENS: ' + toDelete.length + ' linha(s) de header duplicado removida(s).';
}

// ── [MÉTRICAS] Limpar linhas duplicadas de header em MÉTRICAS ────────────────
function limparMetricasCorrompidas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(H.SHEETS.METRICAS);
  if (!sh) return 'MÉTRICAS: aba não encontrada.';

  const lastRow = sh.getLastRow();
  if (lastRow <= 1) return 'MÉTRICAS: vazio.';

  const data = sh.getRange(2, 1, lastRow - 1, 2).getValues();
  const toDelete = [];

  const headerPatterns = ['ID ATLETA', 'ATH_ID', 'NOME COMPLETO', 'ID Atleta'];
  for (let i = data.length - 1; i >= 0; i--) {
    const colA = String(data[i][0] || '').trim().toUpperCase();
    const isHeader = headerPatterns.some(p => colA.includes(p.toUpperCase()));
    if (isHeader) toDelete.push(i + 2);
  }

  for (const row of toDelete) sh.deleteRow(row);

  SpreadsheetApp.flush();
  _log('SYSTEM', 'INFO', 'limparMetricasCorrompidas', toDelete.length + ' linha(s) de header duplicado removidas de MÉTRICAS.', '');
  return 'MÉTRICAS: ' + toDelete.length + ' linha(s) de header duplicado removida(s).';
}

// ── [ERROS] Trimmar aba ERROS mantendo apenas as últimas N linhas ─────────────
function trimmarErros(limite) {
  limite = limite || 200;
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const sh  = ss.getSheetByName(H.SHEETS.ERROS);
  if (!sh) return 'ERROS: aba não encontrada.';

  const lastRow = sh.getLastRow();
  if (lastRow <= limite + 1) return 'ERROS: ' + (lastRow - 1) + ' linhas — dentro do limite.';

  const remover = lastRow - limite - 1; // -1 para preservar header
  sh.deleteRows(2, remover); // deleta das mais antigas (logo após o header)

  SpreadsheetApp.flush();
  _log('SYSTEM', 'INFO', 'trimmarErros', 'ERROS: ' + remover + ' linhas antigas removidas. Mantendo últimas ' + limite + '.', '');
  return 'ERROS: ' + remover + ' linhas antigas removidas (mantidas as últimas ' + limite + ').';
}

// Wrapper para o menu (trimmar com limite padrão 200)
function trimmarErrosPadrao() { trimmarErros(200); }

// ── [SUPABASE_KEY] Configurar via prompt seguro ───────────────────────────────
function configurarSupabaseKey() {
  const ui = SpreadsheetApp.getUi();
  const r  = ui.prompt(
    '🔑 Configurar SUPABASE_KEY',
    'Cole aqui o service_role key do Supabase:\n' +
    '(Dashboard → korlpbclqgmqvpbrungc → Settings → API → service_role)\n\n' +
    '⚠️ Não compartilhe esta chave com ninguém.',
    ui.ButtonSet.OK_CANCEL
  );
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const key = r.getResponseText().trim();
  if (!key || key.length < 20) { ui.alert('❌ Chave inválida. Tente novamente.'); return; }
  PropertiesService.getScriptProperties().setProperty('SUPABASE_KEY', key);
  _log('SYSTEM', 'INFO', 'configurarSupabaseKey', 'SUPABASE_KEY configurada com sucesso.', '');
  ui.alert('✅ SUPABASE_KEY configurada!\n\nO backup automático de tokens no Supabase está ativo.');
}

// ── CORRIGIR WEB APP URL + STRAVA ────────────────────────────────────────────
/**
 * Restaura o WEBAPP_URL para a URL estável do deployment v18.
 * Execute uma vez no editor para corrigir os links do Strava.
 * Depois de criar uma nova implantação via UI, rode definirWebAppUrl(url).
 */
function fixarWebApp() {
  const ui    = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();

  // URL do deployment v18 — ANYONE_ANONYMOUS — confirmada funcionando
  const URL_ESTAVEL = 'https://script.google.com/macros/s/AKfycbzvJzeGQXtfpiRu0C3UI4gC7_9LRIJN0hTXZkR9h8hv3t66d4GTUpdwLIoWZk2Ke-4Mtg/exec';

  const atual = props.getProperty('WEBAPP_URL') || '(não configurado)';
  const ehQuebrado = atual.includes('AKfycby6') || atual.includes('AKfycbwMcWH') || !atual;

  props.setProperty('WEBAPP_URL', URL_ESTAVEL);
  props.setProperty('STRAVA_CLIENT_ID', '153043');

  _log('SYSTEM', 'INFO', 'fixarWebApp',
    'WEBAPP_URL restaurado. Anterior: ' + atual.slice(-40), '');

  ui.alert(
    '✅ Web App Corrigido',
    'WEBAPP_URL atualizado com sucesso!\n\n' +
    '• URL anterior: ' + (ehQuebrado ? '❌ QUEBRADO' : '...' + atual.slice(-30)) + '\n' +
    '• URL nova: ✅ Deployment v18 (estável)\n' +
    '• CLIENT_ID: 153043\n\n' +
    '⚠️ PRÓXIMO PASSO (para URL com código mais recente):\n' +
    'Clique em "Implantar" → "Nova implantação"\n' +
    '→ Tipo: Aplicativo da Web\n' +
    '→ Executar como: Eu (contato@ghiperativo.com.br)\n' +
    '→ Quem pode acessar: Qualquer pessoa\n' +
    '→ Copie a URL /exec e rode: definirWebAppUrl()\n\n' +
    'Por agora os links do Strava já estão funcionando!',
    ui.ButtonSet.OK
  );
}

/**
 * Atualiza o WEBAPP_URL para uma nova URL após criar implantação via UI.
 * Execute após criar nova implantação (Implantar → Nova implantação → Web App).
 */
function definirWebAppUrl() {
  const ui    = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();

  // Tentar detectar automaticamente a URL do serviço atual
  let urlAuto = '';
  try {
    urlAuto = ScriptApp.getService().getUrl() || '';
  } catch(e) { urlAuto = ''; }

  const hint = urlAuto || props.getProperty('WEBAPP_URL') || '';
  const r = ui.prompt(
    '🔗 Definir URL do Web App',
    'Cole aqui a URL /exec do novo deployment:\n' +
    '(termine em .../exec)\n\n' +
    'URL atual: ' + (hint || '(não configurado)'),
    ui.ButtonSet.OK_CANCEL
  );
  if (r.getSelectedButton() !== ui.Button.OK) return;

  let url = r.getResponseText().trim();
  if (!url) url = urlAuto;
  if (!url || !url.includes('script.google.com')) {
    ui.alert('❌ URL inválida. Deve ser uma URL do script.google.com');
    return;
  }

  // Remover parâmetros se houver
  url = url.split('?')[0];

  props.setProperty('WEBAPP_URL', url);
  _log('SYSTEM', 'INFO', 'definirWebAppUrl', 'WEBAPP_URL definido: ' + url, '');

  ui.alert(
    '✅ URL Configurada',
    'WEBAPP_URL salvo:\n' + url + '\n\n' +
    'Os links do Strava agora apontam para esta URL.',
    ui.ButtonSet.OK
  );
}

// ── DIAGNÓSTICO COMPLETO DO SISTEMA ──────────────────────────────────────────
function diagnosticoSistema() {
  const ui    = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();

  const clientId  = props.getProperty('STRAVA_CLIENT_ID') || '❌ não configurado';
  const secret    = props.getProperty('STRAVA_CLIENT_SECRET') ? '✅ configurado' : '❌ não configurado';
  const webApp    = props.getProperty('WEBAPP_URL') || '❌ não configurado';
  const supaKey   = props.getProperty('SUPABASE_KEY') ? '✅ configurado' : '❌ não configurado';
  const adminEmail = props.getProperty('ADMIN_EMAIL') || '❌ não configurado';

  const urlOk = webApp.includes('AKfycbxguMbkP') ? '✅ URL estável v18' :
                webApp.includes('AKfycby6') ? '❌ URL quebrada' :
                webApp ? '⚠️ URL desconhecida' : '❌ vazia';

  const msg = [
    '📊 DIAGNÓSTICO DO SISTEMA\n',
    '🔑 STRAVA_CLIENT_ID: ' + clientId,
    '🔒 STRAVA_CLIENT_SECRET: ' + secret,
    '🌐 WEBAPP_URL: ' + urlOk,
    '   ' + (webApp ? '...' + webApp.slice(-45) : ''),
    '☁️ SUPABASE_KEY: ' + supaKey,
    '📧 ADMIN_EMAIL: ' + adminEmail,
    '',
    webApp.includes('AKfycby6') || !webApp
      ? '⚠️ AÇÃO NECESSÁRIA: rode fixarWebApp() para corrigir URL'
      : '✅ Configuração OK',
  ].join('\n');

  ui.alert('🔬 Diagnóstico', msg, ui.ButtonSet.OK);
}

// ══════════════════════════════════════════════════════════════════════════════
// FUNÇÃO MASTER — execute esta para aplicar todos os ajustes de uma vez
// ══════════════════════════════════════════════════════════════════════════════
function ajustesPlanilha() {
  const ui  = SpreadsheetApp.getUi();
  const log = [];

  try {
    log.push(limparAtividades());
    log.push(limparTokensCorrompidos());
    log.push(limparMetricasCorrompidas());
    log.push(corrigirPainel());
    log.push(corrigirContadorAtivos());
    log.push(popularConfig());
    log.push(trimmarErros(200));

    // Tentar atualizar STRAVA STATUS com a função existente
    try { atualizarStravaStatus(); log.push('STRAVA STATUS: atualizado.'); }
    catch(e) { log.push('STRAVA STATUS: ' + e.message); }

    _log('SYSTEM', 'INFO', 'ajustesPlanilha', 'Todos os ajustes concluídos.', log.join(' | '));

    ui.alert(
      '✅ Ajustes Concluídos',
      log.join('\n'),
      ui.ButtonSet.OK
    );
    return { ok: true, log: log };

  } catch(e) {
    _log('SYSTEM', 'ERRO', 'ajustesPlanilha', e.message, e.stack || '');
    try { ui.alert('❌ Erro em ajustesPlanilha', e.message, ui.ButtonSet.OK); } catch(_) {}
    return { ok: false, erro: e.message };
  }
}
