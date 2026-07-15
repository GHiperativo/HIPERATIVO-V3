/**
 * PainelFix.gs — Correção das fórmulas quebradas do PAINEL
 * Lê dados de ATIVIDADES_CONVERTIDAS e atualiza o PAINEL diretamente
 * NÃO interfere na captura do Strava nem nos scripts existentes
 */

// Nomes das abas
var SHEET_CONV_FIX = '🏃 ATIVIDADES_CONVERTIDAS';
var SHEET_PAINEL_FIX = '📊 PAINEL';
var SHEET_CADASTRO_FIX = '👤 CADASTRO';

/**
 * Atualiza a aba PAINEL com os dados mais recentes de ATIVIDADES_CONVERTIDAS
 * Pode ser chamada manualmente ou por trigger
 */
function corrigirPainel() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var wsConv = ss.getSheetByName(SHEET_CONV_FIX);
  var wsPainel = ss.getSheetByName(SHEET_PAINEL_FIX);
  var wsCad = ss.getSheetByName(SHEET_CADASTRO_FIX);

  if (!wsConv || !wsPainel) {
    Logger.log('Erro: aba ATIVIDADES_CONVERTIDAS ou PAINEL não encontrada');
    return;
  }

  // ── 1. Ler dados de ATIVIDADES_CONVERTIDAS ──────────────────────────────
  // Colunas: A=ID, B=NomeAtleta, C=Atleta, D=Data/Hora, E=Data, L=Tipo, P=DistKm, V=PaceS, W=Pace, Z=FCMedia
  // Linha 1 = header1, Linha 2 = header2, dados a partir da linha 3
  var lastRow = wsConv.getLastRow();
  if (lastRow < 3) {
    Logger.log('ATIVIDADES_CONVERTIDAS sem dados');
    return;
  }

  var dados = wsConv.getRange(3, 1, lastRow - 2, 26).getValues(); // cols A-Z

  // Ordenar por data descendente (col E = índice 4)
  dados.sort(function(a, b) {
    var dA = a[4] ? new Date(a[4]) : new Date(0);
    var dB = b[4] ? new Date(b[4]) : new Date(0);
    return dB - dA;
  });

  // ── 2. Atualizar ÚLTIMAS 10 ATIVIDADES (linhas 11-20, colunas 1-6) ──────
  var tz = Session.getScriptTimeZone();
  for (var i = 0; i < 10; i++) {
    var row = wsPainel.getRange(11 + i, 1, 1, 6);
    if (i < dados.length) {
      var d = dados[i];
      var dataAtiv = d[4]; // col E = Data
      var atleta   = d[2]; // col C = Atleta (nome completo)
      var tipo     = d[11]; // col L = Tipo
      var distKm   = d[15] || 0; // col P = Distancia km
      var pace     = d[22]; // col W = Pace formatado
      var fc       = d[25] ? Math.round(d[25]) : ''; // col Z = FC Media

      var dataFmt = '';
      if (dataAtiv) {
        try {
          dataFmt = Utilities.formatDate(new Date(dataAtiv), tz, 'dd/MM/yy');
        } catch(e) {
          dataFmt = String(dataAtiv).substring(0, 10);
        }
      }

      var distFmt = distKm > 0 ? distKm.toFixed(2) + ' km' : '--';
      var paceFmt = pace || '--';
      var fcFmt   = fc ? fc + ' bpm' : '--';

      row.setValues([[dataFmt, atleta, tipo, distFmt, paceFmt, fcFmt]]);
    } else {
      row.setValues([['', '', '', '', '', '']]);
    }
  }

  // ── 3. Atualizar TREINOS ESTA SEMANA (célula G6:H6) ─────────────────────
  var hoje = new Date();
  hoje.setHours(0, 0, 0, 0);
  var semanaAtras = new Date(hoje.getTime() - 7 * 24 * 60 * 60 * 1000);
  var treinsNaSemana = dados.filter(function(d) {
    if (!d[4]) return false;
    var dt = new Date(d[4]);
    dt.setHours(0, 0, 0, 0);
    return dt >= semanaAtras;
  }).length;
  wsPainel.getRange('G6:H6').setValue(treinsNaSemana);

  // ── 4. Atualizar PACE MÉDIO (célula I6:J6) ───────────────────────────────
  // Pace médio da semana para atividades de corrida com pace válido
  var atividadesSemana = dados.filter(function(d) {
    if (!d[4]) return false;
    var dt = new Date(d[4]);
    dt.setHours(0, 0, 0, 0);
    return dt >= semanaAtras && d[21] > 0; // col V = pace em s/km
  });

  var paceMedioStr = '--';
  if (atividadesSemana.length > 0) {
    var totalPace = atividadesSemana.reduce(function(acc, d) { return acc + (d[21] || 0); }, 0);
    var mediaPace = totalPace / atividadesSemana.length;
    var min = Math.floor(mediaPace / 60);
    var sec = Math.round(mediaPace % 60);
    paceMedioStr = min + "'" + (sec < 10 ? '0' : '') + sec + '"';
  }
  try { wsPainel.getRange('I6:J6').setValue(paceMedioStr); } catch(e) {}

  // ── 5. Atualizar KM MÉDIOS (célula K6:L6) ───────────────────────────────
  var kmMedioStr = '--';
  if (atividadesSemana.length > 0) {
    var totalKm = atividadesSemana.reduce(function(acc, d) { return acc + (d[15] || 0); }, 0);
    var mediaKm = totalKm / atividadesSemana.length;
    kmMedioStr = mediaKm.toFixed(1) + ' km';
  }
  try { wsPainel.getRange('K6:L6').setValue(kmMedioStr); } catch(e) {}

  // ── 6. Atualizar ALERTAS — atletas sem treinar (linhas 11-20, colunas 8-11) ─
  if (wsCad) {
    var lastCadRow = wsCad.getLastRow();
    // CADASTRO: col 1=ID, col 2=Nome, col 26=Status, dados a partir da linha 3
    var cadRows = wsCad.getRange(3, 1, Math.max(1, lastCadRow - 2), 27).getValues();
    var atletas = cadRows.filter(function(r) {
      return r[0] && r[25] !== 'Inativo'; // col 1=ID presente, col 26=Status != Inativo
    }).map(function(r) {
      return String(r[1]).trim(); // col 2 = NOME
    }).filter(function(n) { return n; });

    // Pegar os que aparecem na aba PAINEL já (H11:H20)
    // Para respeitar a lista já estabelecida, ler da planilha
    var alertasAtletasExist = wsPainel.getRange(11, 8, 10, 1).getValues();
    var alertasAtletas = alertasAtletasExist.map(function(r) { return String(r[0]).trim(); }).filter(function(n) { return n; });
    if (alertasAtletas.length === 0) alertasAtletas = atletas.slice(0, 10);

    // Construir mapa: atleta -> data do último treino
    var ultimoTreino = {};
    dados.forEach(function(d) {
      var nome = String(d[2]).trim();
      if (!nome) return;
      var dt = d[4] ? new Date(d[4]) : null;
      if (!dt) return;
      if (!ultimoTreino[nome] || dt > ultimoTreino[nome]) {
        ultimoTreino[nome] = dt;
      }
    });

    for (var j = 0; j < 10; j++) {
      var atleta = alertasAtletas[j] || '';
      if (!atleta) {
        wsPainel.getRange(11 + j, 9, 1, 3).setValues([['', '', '']]);
        continue;
      }

      var ultimaDt = ultimoTreino[atleta] || null;
      var ultimaStr = '';
      var diasStr = '';
      var alertaStr = '';

      if (ultimaDt) {
        ultimaStr = Utilities.formatDate(ultimaDt, tz, 'dd/MM/yy');
        var diffMs = hoje.getTime() - ultimaDt.getTime();
        var dias = Math.floor(diffMs / (24 * 60 * 60 * 1000));
        diasStr = dias;
        if (dias >= 14) {
          alertaStr = '⚠️ ALERTA: ' + dias + 'd sem treinar';
        } else if (dias >= 7) {
          alertaStr = '⚠️ Verificar (' + dias + 'd)';
        } else {
          alertaStr = '✅ Ativo';
        }
      } else {
        ultimaStr = 'Nunca';
        diasStr = '--';
        alertaStr = '⚠️ Verificar';
      }

      wsPainel.getRange(11 + j, 9, 1, 3).setValues([[ultimaStr, diasStr, alertaStr]]);
    }
  }

  // ── 7. Atualizar timestamp ────────────────────────────────────────────────
  var agora = Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy HH:mm');
  wsPainel.getRange('A3').setValue('Atualizado em: ' + agora);

  Logger.log('PAINEL atualizado com sucesso em ' + agora);
  try {
    SpreadsheetApp.getUi().alert('✅ PAINEL atualizado com sucesso!', '', SpreadsheetApp.getUi().ButtonSet.OK);
  } catch(e) {}
}

/**
 * Versão sem UI para uso em triggers automáticos
 */
function corrigirPainelSilencioso() {
  try {
    corrigirPainel();
  } catch(e) {
    Logger.log('Erro em corrigirPainelSilencioso: ' + e.message);
  }
}
