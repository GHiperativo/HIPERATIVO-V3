/**
 * HIPERATIVO V3 — Manutencao.gs
 * Rotinas seguras para corrigir estrutura da planilha real sem apagar alunos.
 */

function corrigirEstruturaLeve() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.alert(
    '🧰 Corrigir estrutura leve',
    'Isso vai corrigir cabeçalhos, migrar colunas deslocadas e reduzir linhas/colunas sobrando.\n\nUm backup da aba CADASTRO será criado antes.\n\nContinuar?',
    ui.ButtonSet.YES_NO
  );
  if (resp !== ui.Button.YES) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', ss.getId());

  const resumo = [];
  try {
    resumo.push(_corrigirCadastroLeve_(ss));
    resumo.push(_corrigirTokensLeve_(ss));
    resumo.push(_corrigirAtividadesLeve_(ss));
    resumo.push(_corrigirErrosLeve_(ss));
    resumo.push(_corrigirPainelLeve_(ss));
    resumo.push(_limparAbasExtrasLeve_(ss));
    _reduzirTamanhoPlanilhaLeve_(ss);
    try { auditarConexoesStrava(); } catch (eAudit) { resumo.push('Auditoria Strava: ' + eAudit.message); }
    _log('SYSTEM', 'INFO', 'corrigirEstruturaLeve', resumo.join(' | '), '');
    ui.alert('✅ Estrutura corrigida', resumo.join('\n'), ui.ButtonSet.OK);
  } catch (e) {
    _log('SYSTEM', 'ERRO', 'corrigirEstruturaLeve', e.message, e.stack || '');
    ui.alert('❌ Erro ao corrigir estrutura', e.message, ui.ButtonSet.OK);
  }
}

function _corrigirCadastroLeve_(ss) {
  const sh = ss.getSheetByName(H.SHEETS.CADASTRO);
  if (!sh) throw new Error('Aba CADASTRO não encontrada.');

  const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
  sh.copyTo(ss).setName(('BACKUP_CADASTRO_' + stamp).substring(0, 99));

  const headers27 = [
    'ID Atleta','Nome Completo','E-mail','WhatsApp','Data Nasc.','Sexo','Peso (kg)','Altura (cm)',
    'Modalidade','Nível','Objetivo','Freq./Semana','Horário Pref.','Condições de Saúde',
    'Histórico de Lesão','Medicamentos','Prova/Meta','Plano','Cidade','Estado','CPF',
    'Canal de Origem','Data Cadastro','Strava Conectado','ID Strava','Status','Observações'
  ];

  const lastRow = Math.max(sh.getLastRow(), 2);
  const old = sh.getRange(1, 1, lastRow, Math.max(sh.getLastColumn(), 27)).getValues();
  const migrated = [];
  const titulo = new Array(27).fill('');
  titulo[0] = '👤  CADASTRO DE ATLETAS — GRUPO HIPERATIVO';
  migrated.push(titulo);
  migrated.push(headers27);

  for (let r = 2; r < old.length; r++) {
    const row = old[r];
    const athId = String(row[0] || '').trim();
    if (!athId) continue;
    const out = new Array(27).fill('');
    out[0] = athId;
    out[1] = row[1] || '';
    out[2] = row[2] || '';
    out[3] = row[3] || '';
    out[4] = row[4] || '';
    out[5] = row[5] || '';
    out[6] = row[6] || '';
    out[7] = row[7] || '';
    out[8] = row[8] || '';
    out[9] = row[9] || '';
    out[10] = row[10] || '';
    out[11] = row[11] || '';
    out[12] = row[12] || '';
    out[13] = row[13] || '';
    out[14] = row[14] || '';

    // Layout antigo parcial: plano/cidade/estado/origem/data/strava/id/status/obs ficaram deslocados.
    out[15] = '';
    out[16] = '';
    out[17] = row[15] || '';
    out[18] = row[16] || '';
    out[19] = row[17] || '';
    out[20] = '';
    out[21] = row[18] || '';
    out[22] = row[19] || '';
    out[23] = _normalizarStatusStrava_(row[20] || row[23] || '');
    out[24] = row[21] || row[24] || '';
    out[25] = row[22] || row[25] || 'Ativo';
    out[26] = row[23] || row[26] || '';
    migrated.push(out);
  }

  sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).breakApart();
  sh.clearContents();
  sh.getRange(1, 1, migrated.length, 27).setValues(migrated);
  sh.getRange(1, 1, 1, 27).mergeAcross().setFontWeight('bold').setHorizontalAlignment('center');
  sh.getRange(2, 1, 1, 27).setFontWeight('bold').setWrap(true);
  sh.setFrozenRows(2);
  sh.setFrozenColumns(2);
  sh.autoResizeColumns(1, 27);
  _garantirDimensoes_(sh, Math.max(100, migrated.length + 20), 27);
  return 'CADASTRO migrado para 27 colunas: ' + (migrated.length - 2) + ' atletas.';
}

function _corrigirTokensLeve_(ss) {
  const sh = ss.getSheetByName(H.SHEETS.TOKENS) || ss.insertSheet(H.SHEETS.TOKENS);
  const headers = ['ID Execução','ID Atleta','Nome Atleta','Access Token','Refresh Token','Expira Em (Unix)','Scope','ID Strava','Últ. Atualização','Status Token'];
  sh.getRange(1, 1).setValue('🔐  TOKENS STRAVA — NÃO COMPARTILHAR');
  sh.getRange(2, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  sh.setFrozenRows(2);
  sh.setFrozenColumns(2);
  _garantirDimensoes_(sh, Math.max(100, sh.getLastRow() + 20), headers.length);
  return 'TOKENS alinhada para 10 colunas.';
}

function _corrigirAtividadesLeve_(ss) {
  const sh = ss.getSheetByName(H.SHEETS.ATIVIDADES) || ss.insertSheet(H.SHEETS.ATIVIDADES);
  const headers = [
    'ID Execução','ID Atleta','Nome Atleta','Data','Tipo','Fonte','ID Strava','Nome da Atividade',
    'Tempo Mov. (s)','Tempo Total (s)','Distância (m)','Distância (km)','Vel. Média (m/s)',
    'Vel. Média (km/h)','Pace Médio (s/km)','Pace Médio (min:s)','FC Média','FC Máxima',
    'Elevação (m)','Calorias','Cadência','Potência','Rota','Importado em'
  ];
  sh.getRange(1, 1).setValue('🏃  ATIVIDADES — DADOS IMPORTADOS DO STRAVA');
  sh.getRange(2, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  sh.setFrozenRows(2);
  sh.setFrozenColumns(4);
  _garantirDimensoes_(sh, Math.max(100, sh.getLastRow() + 50), headers.length);
  return 'ATIVIDADES alinhada para 24 colunas.';
}

function _corrigirErrosLeve_(ss) {
  const sh = ss.getSheetByName(H.SHEETS.ERROS) || ss.insertSheet(H.SHEETS.ERROS);
  const headers = ['Data/Hora','Nível','ID Atleta','Função de Origem','Mensagem','Detalhes JSON','Resolvido','Ação Tomada'];
  sh.getRange(1, 1).setValue('🔴  LOG DE ERROS E EVENTOS DO SISTEMA');
  sh.getRange(2, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  sh.setFrozenRows(2);
  _garantirDimensoes_(sh, Math.max(100, sh.getLastRow() + 50), headers.length);
  return 'ERROS alinhada para 8 colunas.';
}

function _corrigirPainelLeve_(ss) {
  try {
    corrigirFormulasPainel();
    return 'PAINEL com fórmulas corrigidas.';
  } catch (e) {
    return 'PAINEL não corrigido automaticamente: ' + e.message;
  }
}

function _limparAbasExtrasLeve_(ss) {
  ['Página1', 'STRAVA_TOKEN_STATUS'].forEach(nome => {
    const sh = ss.getSheetByName(nome);
    if (sh) sh.hideSheet();
  });
  return 'Abas extras Página1/STRAVA_TOKEN_STATUS ocultadas quando existentes.';
}

function _reduzirTamanhoPlanilhaLeve_(ss) {
  ss.getSheets().forEach(sh => {
    if (sh.getName().indexOf('BACKUP_') === 0) return;
    const keepRows = Math.max(100, sh.getLastRow() + 50);
    const keepCols = Math.max(8, sh.getLastColumn());
    _garantirDimensoes_(sh, keepRows, keepCols);
  });
}

function _garantirDimensoes_(sh, keepRows, keepCols) {
  const maxRows = sh.getMaxRows();
  const maxCols = sh.getMaxColumns();
  if (maxRows > keepRows) sh.deleteRows(keepRows + 1, maxRows - keepRows);
  if (maxCols > keepCols) sh.deleteColumns(keepCols + 1, maxCols - keepCols);
}

function _normalizarStatusStrava_(valor) {
  const v = String(valor || '').trim().toLowerCase();
  if (v === 'sim' || v === 'true' || v === 'conectado') return 'Sim';
  if (v === 'pendente') return 'Pendente';
  return 'Não';
}
