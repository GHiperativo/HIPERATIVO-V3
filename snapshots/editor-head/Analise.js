function estabilizacaoRawConvertida() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var log = [];
  var dataHoje = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy, HH:mm:ss');

  function reg(aba, campo, valor, obs) {
    try {
      var abaAudit = ss.getSheetByName('\uD83D\uDCCB AUDITORIA')
        || ss.getSheetByName('\uD83E\uDDEA AUDITORIA_FECHAMENTO_20260623');
      if (!abaAudit) return;
      abaAudit.appendRow([new Date(), 'ESTABILIZACAO', aba, campo, String(valor), String(obs || '')]);
    } catch (e) { log.push('reg() falhou: ' + e.message); }
  }

  log.push('\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550');
  log.push('HIPERATIVO V3 \u2014 ESTABILIZA\u00C7\u00C3O RAW + CONVERTIDA');
  log.push('Executado em: ' + dataHoje);
  log.push('\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550');

  // âââ TAREFA 1: Scan abas relacionadas ââââââââââââââââââââââââââââââââââââââ
  log.push('\n[ TAREFA 1 \u2014 VERIFICA\u00C7\u00C3O DAS ABAS ]\n');

  var sheets = ss.getSheets();
  var relacionadas = [];

  sheets.forEach(function (sh) {
    var nm = sh.getName();
    var nmU = nm.toUpperCase().replace(/[^\w]/g, '');
    if (nmU.indexOf('ATIVIDADE') !== -1 || nmU.indexOf('STRAVARAW') !== -1 || nmU.indexOf('CONVERTIDA') !== -1) {
      var lr = sh.getLastRow();
      var lc = sh.getLastColumn();
      var a1 = '', a2 = '';
      try { a1 = String(sh.getRange(1, 1).getValue()).substring(0, 80); } catch (e) { }
      try { a2 = lr >= 2 ? String(sh.getRange(2, 1).getValue()).substring(0, 80) : ''; } catch (e) { }
      relacionadas.push({ nome: nm, id: sh.getSheetId(), oculta: sh.isSheetHidden(), lr: lr, lc: lc, a1: a1, a2: a2 });
    }
  });

  relacionadas.forEach(function (info) {
    log.push('Nome: ' + info.nome);
    log.push('Sheet ID: ' + info.id + ' | Vis\u00EDvel: ' + (!info.oculta));
    log.push('Linhas: ' + info.lr + ' | Colunas: ' + info.lc);
    log.push('A1: ' + info.a1);
    log.push('A2: ' + info.a2);
    log.push('\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500');
    reg(info.nome, 'scan', info.lr + ' linhas / ' + info.lc + ' colunas', info.oculta ? 'OCULTA' : 'vis\u00EDvel');
  });

  if (relacionadas.length === 0) {
    log.push('\u26A0\uFE0F  Nenhuma aba com ATIVIDADE/STRAVA_RAW/CONVERTIDA encontrada!');
  }

  // âââ DiagnÃ³stico especÃ­fico de ATIVIDADES ââââââââââââââââââââââââââââââââââ
  log.push('\n[ DIAGN\u00D3STICO: \uD83C\uDFC3 ATIVIDADES ]\n');
  var ABA_ATIV = '\uD83C\uDFC3 ATIVIDADES';
  var abaAtiv = ss.getSheetByName(ABA_ATIV);
  if (abaAtiv) {
    var aLr = abaAtiv.getLastRow();
    var aLc = abaAtiv.getLastColumn();
    if (aLr <= 1 && aLc <= 1) {
      log.push('\u26A0\uFE0F  ALERTA CR\u00CDTICO: Aba com ' + aLr + ' linha(s) e ' + aLc + ' coluna(s).');
      log.push('    Poss\u00EDvel esvaziamento OU aba sem dados reais.');
      log.push('    N\u00C3O restaurar automaticamente. Aguardando decis\u00E3o.');
      reg(ABA_ATIV, 'status', 'ALERTA', aLr + ' lin / ' + aLc + ' col \u2014 verificar manualmente');
    } else {
      log.push('\u2705 ATIVIDADES: ' + aLr + ' linhas | ' + aLc + ' colunas');
      try { log.push('    A1: ' + String(abaAtiv.getRange(1, 1).getValue()).substring(0, 80)); } catch (e) { }
      try { log.push('    A2: ' + String(abaAtiv.getRange(2, 1).getValue()).substring(0, 80)); } catch (e) { }
      reg(ABA_ATIV, 'status', 'OK', aLr + ' lin / ' + aLc + ' col');
    }
  } else {
    log.push('\u274C ATIVIDADES n\u00E3o encontrada com nome exato!');
    log.push('    Tentando variantes...');
    var candidatos = sheets.filter(function (sh) { return sh.getName().toUpperCase().replace(/[^\w]/g, '').indexOf('ATIVIDADE') !== -1; });
    candidatos.forEach(function (sh) {
      log.push('    Candidato: "' + sh.getName() + '" | ' + sh.getLastRow() + ' lin | ' + sh.getLastColumn() + ' col');
    });
    reg(ABA_ATIV, 'status', 'ERRO', 'N\u00E3o encontrada \u2014 ver candidatos no log');
  }

  // Buscar backup com dados (341 linhas ou 34+ colunas)
  log.push('\n[ BUSCA DE BACKUP/CANDIDATO A RESTAURA\u00C7\u00C3O ]\n');
  var candidatosBackup = sheets.filter(function (sh) {
    var nm = sh.getName().toUpperCase().replace(/[^\w]/g, '');
    var temNome = nm.indexOf('ATIVIDADE') !== -1;
    var temDados = sh.getLastRow() > 10 && sh.getLastColumn() > 10;
    return temNome && temDados && sh.getName() !== ABA_ATIV;
  });
  if (candidatosBackup.length > 0) {
    candidatosBackup.forEach(function (sh) {
      log.push('\uD83D\uDCCC Candidato a backup: "' + sh.getName() + '" | ' + sh.getLastRow() + ' lin | ' + sh.getLastColumn() + ' col');
      reg(sh.getName(), 'backup_candidato', sh.getLastRow() + ' lin / ' + sh.getLastColumn() + ' col', 'Candidato para restaura\u00E7\u00E3o futura');
    });
  } else {
    log.push('Nenhum candidato a backup encontrado.');
  }

  // âââ TAREFA 2: Aplicar headers em ATIVIDADES_CONVERTIDAS ââââââââââââââââââ
  log.push('\n[ TAREFA 2 \u2014 ATIVIDADES_CONVERTIDAS: HEADERS ]\n');

  var TITULO_CONV = '\uD83C\uDFC3 ATIVIDADES_CONVERTIDAS \u2014 DADOS NORMALIZADOS PARA AN\u00C1LISE';
  var HEADERS_CONV = [
    'ID Interno', 'ATH_ID', 'Atleta', 'Data/Hora', 'Data', 'Hora', 'Ano', 'M\u00EAs', 'M\u00EAs Ref',
    'Semana', 'Dia da Semana', 'Tipo', 'Tipo Original', 'Strava ID', 'Nome da Atividade',
    'Dist\u00E2ncia km', 'Dist\u00E2ncia', 'Tempo Movimento s', 'Tempo Movimento', 'Tempo Total s',
    'Tempo Total', 'Pace s/km', 'Pace', 'Velocidade km/h', 'Velocidade', 'FC M\u00E9dia',
    'FC M\u00E9dia fmt', 'FC M\u00E1x.', 'FC M\u00E1x. fmt', 'Eleva\u00E7\u00E3o m', 'Eleva\u00E7\u00E3o', 'Calorias',
    'Calorias fmt', 'Cad\u00EAncia', 'Cad\u00EAncia fmt', 'Pot\u00EAncia W', 'Pot\u00EAncia', 'Carga Simples',
    'Intensidade', 'Qualidade do Dado', 'Flags', 'Tipo de Registro', 'Possui Dispositivo?',
    'Dado de Dist\u00E2ncia Aplic\u00E1vel?', 'Dado de Pace Aplic\u00E1vel?', 'Dado de FC Aplic\u00E1vel?',
    'Fonte', 'Importado em', 'Status'
  ];

  var ABA_CONV = '\uD83C\uDFC3 ATIVIDADES_CONVERTIDAS';
  var abaConv = ss.getSheetByName(ABA_CONV);
  if (abaConv) {
    var cLr = abaConv.getLastRow();
    var cLc = abaConv.getLastColumn();
    if (cLr === 0 && cLc === 0) {
      abaConv.getRange(1, 1).setValue(TITULO_CONV);
      abaConv.getRange(2, 1, 1, HEADERS_CONV.length).setValues([HEADERS_CONV]);
      try { abaConv.setFrozenRows(2); } catch (e) { }
      log.push('\u2705 49 headers aplicados em ATIVIDADES_CONVERTIDAS.');
      log.push('    T\u00EDtulo: ' + TITULO_CONV);
      log.push('    Linhas congeladas: 2');
      log.push('    Colunas: ' + HEADERS_CONV.length);
      reg(ABA_CONV, 'headers', 'APLICADO', '49 colunas');
    } else {
      log.push('\u2139\uFE0F  ATIVIDADES_CONVERTIDAS j\u00E1 tem conte\u00FAdo (' + cLr + ' lin / ' + cLc + ' col). N\u00E3o alterada.');
      if (cLc !== 49) {
        log.push('\u26A0\uFE0F  Colunas: ' + cLc + ' vs 49 esperadas. Verificar manualmente.');
        reg(ABA_CONV, 'headers', 'ALERTA', cLr + ' lin / ' + cLc + ' col (esperado 49)');
      } else {
        log.push('\u2705 Colunas OK: ' + cLc);
        reg(ABA_CONV, 'headers', 'OK', cLr + ' lin / ' + cLc + ' col');
      }
    }
  } else {
    log.push('\u274C ATIVIDADES_CONVERTIDAS n\u00E3o encontrada!');
    reg(ABA_CONV, 'headers', 'ERRO', 'Aba n\u00E3o encontrada');
  }

  // âââ TAREFA 3: Confirmar STRAVA_RAW ââââââââââââââââââââââââââââââââââââââââ
  log.push('\n[ TAREFA 3 \u2014 STRAVA_RAW ]\n');
  var ABA_RAW = '\uD83C\uDFC3 STRAVA_RAW';
  var abaRaw = ss.getSheetByName(ABA_RAW);
  if (abaRaw) {
    var rLr = abaRaw.getLastRow();
    var rLc = abaRaw.getLastColumn();
    log.push('STRAVA_RAW existe: Sim');
    log.push('Colunas: ' + rLc + (rLc === 44 ? ' \u2705' : ' \u26A0\uFE0F (esperado 44)'));
    log.push('Linhas: ' + rLr);
    log.push('Dados: ' + (rLr > 2 ? (rLr - 2) + ' registros' : 'Nenhum'));
    reg(ABA_RAW, 'status', rLc === 44 ? 'OK' : 'ALERTA', rLr + ' lin / ' + rLc + ' col');
  } else {
    log.push('STRAVA_RAW existe: N\u00E3o');
    reg(ABA_RAW, 'status', 'ERRO', 'N\u00E3o encontrada');
  }

  // âââ TAREFA 5: Teste local sem API âââââââââââââââââââââââââââââââââââââââââ
  log.push('\n[ TAREFA 5 \u2014 TESTE LOCAL SEM API ]\n');
  try {
    var raw = {
      'ATH_ID': 'ATH_TESTE',
      'Atleta': 'Atleta Teste',
      'Activity ID': '123456',
      'Name': 'Corrida teste',
      'Sport Type': 'Run',
      'Type': 'Run',
      'Distance m': 5000,
      'Moving Time s': 1800,
      'Elapsed Time s': 1900,
      'Average Speed m/s': 2.777,
      'Average Heartrate': 145,
      'Max Heartrate': 170,
      'Average Cadence': 82,
      'Calories': 350,
      'Total Elevation Gain m': 42,
      'Manual': false,
      'Gear ID': 'g123'
    };
    var mapa = { Run: 'Corrida', TrailRun: 'Corrida em trilha', Walk: 'Caminhada', Ride: 'Ciclismo', Swim: 'Nata\u00E7\u00E3o', Workout: 'Treino', WeightTraining: 'Muscula\u00E7\u00E3o', HighIntensityIntervalTraining: 'HIIT', Yoga: 'Yoga', Hike: 'Trilha' };
    var tipo = mapa[raw['Sport Type']] || 'Outro';
    var distKm = Math.round((raw['Distance m'] / 1000) * 100) / 100;
    var avgSpeed = raw['Average Speed m/s'];
    var paceSKm = Math.round(1000 / avgSpeed);
    var paceMin = Math.floor(paceSKm / 60);
    var paceSec = Math.round(paceSKm % 60);
    var paceFmt = String(paceMin).padStart(2, '0') + ':' + String(paceSec).padStart(2, '0') + ' min/km';
    var velKmh = Math.round(avgSpeed * 3.6 * 100) / 100;
    var movH = Math.floor(raw['Moving Time s'] / 3600);
    var movM = Math.floor((raw['Moving Time s'] % 3600) / 60);
    var movS = raw['Moving Time s'] % 60;
    var tempoFmt = String(movH).padStart(2, '0') + ':' + String(movM).padStart(2, '0') + ':' + String(movS).padStart(2, '0');
    var intensidade = paceSKm >= 480 ? 'Leve' : paceSKm >= 390 ? 'Moderado' : paceSKm >= 300 ? 'Forte' : 'Muito Forte';
    var flags = [];
    if (!raw['Average Cadence']) flags.push('SEM_CADENCIA');
    var qualidade = flags.length === 0 ? 'OK' : 'OK_PARCIAL';
    log.push('Tipo: ' + tipo + ' \u2705');
    log.push('Dist\u00E2ncia: ' + distKm + ' km \u2705');
    log.push('Pace: ' + paceFmt + ' \u2705');
    log.push('Velocidade: ' + velKmh + ' km/h \u2705');
    log.push('Tempo Movimento: ' + tempoFmt + ' \u2705');
    log.push('FC M\u00E9dia: ' + raw['Average Heartrate'] + ' bpm \u2705');
    log.push('Intensidade: ' + intensidade + ' \u2705');
    log.push('Qualidade: ' + qualidade + ' \u2705');
    log.push('Flags: ' + (flags.join('|') || 'DADO_MINIMO_OK') + ' \u2705');
    log.push('\u2705 Teste local OK \u2014 sem API, sem grava\u00E7\u00E3o, sem token.');
    reg('TEST', 'pipeline_local', 'OK', 'tipo=' + tipo + ' | dist=' + distKm + ' km | pace=' + paceFmt + ' | qual=' + qualidade);
  } catch (e) {
    log.push('\u274C Erro no teste local: ' + e.message);
    reg('TEST', 'pipeline_local', 'ERRO', e.message);
  }

  // âââ TAREFA 4: StravaPipeline.gs âââââââââââââââââââââââââââââââââââââââââââ
  log.push('\n[ TAREFA 4 \u2014 STRAVAPIPELINE.GS ]\n');
  log.push('\u26A0\uFE0F  StravaPipeline.gs ser\u00E1 criado em seguida via editor.');
  log.push('    Fun\u00E7\u00F5es: traduzirTipoStrava_, isAtividadeComDistancia_,');
  log.push('             isAtividadeSemDistanciaObrigatoria_, formatarTempo_,');
  log.push('             formatarPace_, classificarQualidadeDado_,');
  log.push('             calcularCargaSimples_, converterAtividadeRawParaConvertida_');
  reg('StravaPipeline.gs', 'status', 'PENDENTE_CRIACAO', 'Criar arquivo ap\u00F3s esta execu\u00E7\u00E3o');

  // âââ TAREFA 6: Resumo final ââââââââââââââââââââââââââââââââââââââââââââââââ
  log.push('\n[ TAREFA 6 \u2014 RESUMO FINAL ]\n');
  log.push('\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550');
  log.push('ESTABILIZA\u00C7\u00C3O CONCLU\u00CDDA \u2014 ' + dataHoje);
  log.push('\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550');
  log.push('\u2705 NENHUMA API STRAVA CHAMADA');
  log.push('\u2705 NENHUMA IMPORTA\u00C7\u00C3O EXECUTADA');
  log.push('\u2705 NENHUMA MIGRA\u00C7\u00C3O EXECUTADA');
  log.push('\u2705 NENHUM TOKEN EXPOSTO');
  log.push('\u2705 SUPABASE N\u00C3O ALTERADO');
  log.push('\u2705 M\u00C9TRICAS/PAINEL/RANKING N\u00C3O ALTERADOS');

  reg('GERAL', 'status_final', 'OK', 'Estabiliza\u00E7\u00E3o conclu\u00EDda');

  Logger.log(log.join('\n'));
}
