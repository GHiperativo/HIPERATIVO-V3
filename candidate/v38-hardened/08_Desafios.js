/**
 * ═══════════════════════════════════════════════════════════════════════
 * HIPERATIVO V3 — 08_Desafios.gs
 * Sistema de desafios mensais com rastreamento de progresso por atleta
 * ═══════════════════════════════════════════════════════════════════════
 */

// ── CRIAR ABAS DE DESAFIOS ────────────────────────────────────────────────────
function criarAbasDesafios() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  // ── Aba DESAFIOS ─────────────────────────────────────────────────────────
  let wsD = ss.getSheetByName(H.SHEETS.DESAFIOS);
  if (wsD) ss.deleteSheet(wsD);
  wsD = ss.insertSheet(H.SHEETS.DESAFIOS);

  wsD.getRange(1, 1, 1, 10).mergeAcross()
    .setValue('🎯 DESAFIOS — GRUPO HIPERATIVO')
    .setFontFamily('Arial').setFontSize(13).setFontWeight('bold')
    .setFontColor('#FFFFFF').setBackground('#001F3F')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  wsD.setRowHeight(1, 36);

  const cabsD = [
    'ID Desafio','Nome','Descrição','Tipo',
    'Meta (valor)','Unidade','Data Início','Data Fim',
    'Status','Premiação'
  ];
  wsD.getRange(2, 1, 1, cabsD.length).setValues([cabsD])
    .setFontFamily('Arial').setFontSize(10).setFontWeight('bold')
    .setBackground('#003D7A').setFontColor('#FFFFFF')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  wsD.setRowHeight(2, 26);

  for (var r = 3; r <= 102; r++) {
    wsD.getRange(r, 1, 1, 10).setBackground(r % 2 === 0 ? '#FFFFFF' : '#FFF8F0');
    wsD.setRowHeight(r, 22);
  }

  _dropdown(wsD, 4, 3, 102, ['Volume_km','Consistência_treinos','Pace_melhoria','Distância_única']);
  _dropdown(wsD, 9, 3, 102, ['Rascunho','Ativo','Encerrado','Cancelado']);

  wsD.getRange(3, 7, 100, 2).setNumberFormat('DD/MM/YYYY');

  // Formatação condicional de status
  [
    ['Ativo',    '#D6F5EC'],
    ['Encerrado','#E8E8E8'],
    ['Cancelado','#FDECEA'],
    ['Rascunho', '#FFF8E1'],
  ].forEach(function(pair) { _condTexto(wsD, 'I3:I102', pair[0], pair[1]); });

  var largsD = [16, 28, 40, 22, 12, 12, 13, 13, 14, 30];
  largsD.forEach(function(w, i) { wsD.setColumnWidth(i + 1, w * 7); });
  wsD.setFrozenRows(2);

  // ── Aba DESAFIOS_PROGRESSO ───────────────────────────────────────────────
  var wsP = ss.getSheetByName(H.SHEETS.DESAFIOS_PROGRESSO);
  if (wsP) ss.deleteSheet(wsP);
  wsP = ss.insertSheet(H.SHEETS.DESAFIOS_PROGRESSO);

  wsP.getRange(1, 1, 1, 10).mergeAcross()
    .setValue('📊 PROGRESSO DOS DESAFIOS — GRUPO HIPERATIVO')
    .setFontFamily('Arial').setFontSize(13).setFontWeight('bold')
    .setFontColor('#FFFFFF').setBackground('#1D9E75')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  wsP.setRowHeight(1, 36);

  var cabsP = [
    'ID','Desafio ID','Atleta ID','Nome Atleta',
    'Progresso Atual','Meta','Percentual',
    'Última Atualização','Concluído','Data Conclusão'
  ];
  wsP.getRange(2, 1, 1, cabsP.length).setValues([cabsP])
    .setFontFamily('Arial').setFontSize(10).setFontWeight('bold')
    .setBackground('#1D9E75').setFontColor('#FFFFFF')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  wsP.setRowHeight(2, 26);

  for (var rp = 3; rp <= 502; rp++) {
    wsP.getRange(rp, 1, 1, 10).setBackground(rp % 2 === 0 ? '#FFFFFF' : '#D6F5EC');
    wsP.setRowHeight(rp, 20);
  }

  _condTexto(wsP, 'I3:I502', 'SIM', '#D6F5EC');
  _condTexto(wsP, 'I3:I502', 'NÃO', '#FFFFFF');

  var largsP = [22, 18, 14, 22, 16, 12, 14, 20, 12, 18];
  largsP.forEach(function(w, i) { wsP.setColumnWidth(i + 1, w * 7); });
  wsP.setFrozenRows(2);

  _log('SYSTEM', 'INFO', 'criarAbasDesafios', 'Abas DESAFIOS e DESAFIOS_PROGRESSO criadas', '');
  ui.alert('✅ Abas de Desafios criadas!',
    '"' + H.SHEETS.DESAFIOS + '" — cadastre desafios com Status "Ativo".\n' +
    '"' + H.SHEETS.DESAFIOS_PROGRESSO + '" — calculado automaticamente.\n\n' +
    'Use: Menu → 🎯 Desafios → Atualizar progresso dos desafios.',
    ui.ButtonSet.OK);
}

// ── ATUALIZAR PROGRESSO DOS DESAFIOS ─────────────────────────────────────────
function atualizarProgressoDesafiosUI() {
  atualizarProgressoDesafios();
  SpreadsheetApp.getUi().alert('✅ Progresso dos desafios atualizado com sucesso!');
}

function atualizarProgressoDesafios() {
  var ss              = SpreadsheetApp.getActiveSpreadsheet();
  var sheetDesafios   = ss.getSheetByName(H.SHEETS.DESAFIOS);
  var sheetProgresso  = ss.getSheetByName(H.SHEETS.DESAFIOS_PROGRESSO);
  var sheetAtividades = ss.getSheetByName(H.SHEETS.ATIVIDADES);
  var sheetCadastro   = ss.getSheetByName(H.SHEETS.CADASTRO);

  if (!sheetDesafios || !sheetProgresso || !sheetAtividades || !sheetCadastro) {
    _log('SYSTEM', 'ERRO', 'atualizarProgressoDesafios',
      'Uma ou mais abas necessárias não encontradas. Execute "Criar abas de desafios" primeiro.', '');
    return;
  }

  var desafios   = sheetDesafios.getDataRange().getValues();
  var atividades = sheetAtividades.getDataRange().getValues().slice(2)
    .filter(function(r) { return r[H.ATIV.ATH_ID - 1]; });
  var atletas    = sheetCadastro.getDataRange().getValues().slice(2)
    .filter(function(r) { return r[H.CAD.ID - 1] && r[H.CAD.STATUS - 1] !== 'Inativo'; });

  // Limpar progresso anterior (mantém cabeçalho nas linhas 1-2)
  if (sheetProgresso.getLastRow() > 2) {
    sheetProgresso.getRange(3, 1, sheetProgresso.getLastRow() - 2, 10).clearContent();
  }

  var linhaProgresso = 3;

  for (var d = 2; d < desafios.length; d++) {
    var desafio = desafios[d];
    if (String(desafio[8]) !== 'Ativo') continue; // col 9 = Status

    var desafioId   = String(desafio[0]);
    var tipo        = String(desafio[3]); // col 4 = Tipo
    var meta        = parseFloat(desafio[4]) || 0;
    var dataInicio  = desafio[6] instanceof Date ? desafio[6] : new Date(desafio[6]);
    var dataFim     = desafio[7] instanceof Date ? desafio[7] : new Date(desafio[7]);

    if (isNaN(dataInicio) || isNaN(dataFim) || meta <= 0) continue;

    for (var a = 0; a < atletas.length; a++) {
      var atletaId   = String(atletas[a][H.CAD.ID - 1]);
      var nomeAtleta = String(atletas[a][H.CAD.NOME - 1] || atletaId);

      // Filtrar atividades do atleta dentro do período do desafio
      var atividadesDesafio = atividades.filter(function(act) {
        if (String(act[H.ATIV.ATH_ID - 1]) !== atletaId) return false;
        var dataAct = act[H.ATIV.DATA - 1];
        if (!(dataAct instanceof Date)) return false;
        return dataAct >= dataInicio && dataAct <= dataFim;
      });

      var progresso = 0;

      if (tipo === 'Volume_km') {
        progresso = atividadesDesafio.reduce(function(sum, act) {
          return sum + (parseFloat(act[H.ATIV.DIST_KM - 1]) || 0);
        }, 0);
      } else if (tipo === 'Consistência_treinos') {
        progresso = atividadesDesafio.length;
      } else if (tipo === 'Distância_única') {
        progresso = atividadesDesafio.reduce(function(max, act) {
          return Math.max(max, parseFloat(act[H.ATIV.DIST_KM - 1]) || 0);
        }, 0);
      } else if (tipo === 'Pace_melhoria') {
        // Melhor (menor) pace médio em corridas
        var paces = atividadesDesafio
          .filter(function(act) {
            return String(act[H.ATIV.TIPO - 1]) === 'Corrida' && Number(act[H.ATIV.PACE_S - 1]) > 0;
          })
          .map(function(act) { return Number(act[H.ATIV.PACE_S - 1]); });
        progresso = paces.length > 0 ? Math.min.apply(null, paces) : 0;
      }

      progresso = Math.round(progresso * 100) / 100;
      var percentual = meta > 0 ? Math.min(Math.round((progresso / meta) * 100), 100) : 0;
      var concluido  = percentual >= 100;
      var agora      = new Date();

      sheetProgresso.getRange(linhaProgresso, 1, 1, 10).setValues([[
        'PROG_' + desafioId + '_' + atletaId,
        desafioId,
        atletaId,
        nomeAtleta,
        progresso,
        meta,
        percentual + '%',
        agora,
        concluido ? 'SIM' : 'NÃO',
        concluido ? agora : '',
      ]]);

      linhaProgresso++;
    }
  }

  var totalRegistros = linhaProgresso - 3;
  _log('SYSTEM', 'INFO', 'atualizarProgressoDesafios',
    'Progresso atualizado: ' + totalRegistros + ' registros', '');
}
