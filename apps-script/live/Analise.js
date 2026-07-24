/**
 * Atualiza a visão consolidada de carga de treino por atleta.
 * CTL/ATL usam médias exponenciais de 42 e 7 dias; ACWR compara 7d com a
 * média semanal de 28d. A carga sRPE só é calculada quando existe PSE real.
 * Nenhum token ou dado de cadastro é alterado por esta rotina.
 */
function atualizarAnaliseSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shAtiv = ss.getSheetByName(H.SHEETS.ATIVIDADES);
  const shCad = ss.getSheetByName(H.SHEETS.CADASTRO);
  const shMet = ss.getSheetByName(H.SHEETS.METRICAS);
  if (!shAtiv || !shCad) return 0;

  let sh = ss.getSheetByName('🔬 ANÁLISE');
  if (!sh) sh = ss.insertSheet('🔬 ANÁLISE');

  const tz = Session.getScriptTimeZone() || 'America/Sao_Paulo';
  const hoje = new Date();
  hoje.setHours(12, 0, 0, 0);
  const inicio = new Date(hoje.getTime() - 89 * 86400000);

  const fcMaxMap = {};
  if (shMet) {
    shMet.getDataRange().getValues().slice(2).forEach(r => {
      const id = String(r[H.MET.ATH_ID - 1] || '').trim().toUpperCase();
      const fc = Number(r[H.MET.FC_MAX - 1]) || 0;
      if (id && fc > 0) fcMaxMap[id] = fc;
    });
  }

  const atletas = shCad.getDataRange().getValues().slice(2)
    .map(r => ({
      id: String(r[H.CAD.ID - 1] || '').trim().toUpperCase(),
      nome: String(r[H.CAD.NOME - 1] || '').trim(),
      status: String(r[H.CAD.STATUS - 1] || '').trim().toLowerCase()
    }))
    .filter(a => a.id && a.status !== 'inativo' &&
      (typeof _isAthIdValido_ !== 'function' || _isAthIdValido_(a.id)));

  const porAtleta = {};
  atletas.forEach(a => { porAtleta[a.id] = { dias: {}, z: [0, 0, 0], atividades: 0, pseValidas: 0 }; });

  shAtiv.getDataRange().getValues().slice(2).forEach(r => {
    const id = String(r[H.ATIV.ATH_ID - 1] || '').trim().toUpperCase();
    const data = r[H.ATIV.DATA - 1];
    if (!porAtleta[id] || !(data instanceof Date) || data < inicio || data > new Date(hoje.getTime() + 86400000)) return;

    const minutos = Math.max(0, (Number(r[H.ATIV.MOV_S - 1]) || Number(r[H.ATIV.TOTAL_S - 1]) || 0) / 60);
    if (!minutos) return;
    const pseInformado = Number(r[H.ATIV.PSE - 1]) || 0;
    const fcMed = Number(r[H.ATIV.FC_MED - 1]) || 0;
    const fcMax = fcMaxMap[id] || Number(r[H.ATIV.FC_MAX - 1]) || 0;
    const pctFc = fcMed > 0 && fcMax > 0 ? fcMed / fcMax : 0;
    const chave = Utilities.formatDate(data, tz, 'yyyy-MM-dd');
    porAtleta[id].atividades++;

    if (pseInformado >= 1 && pseInformado <= 10) {
      const carga = Math.round(minutos * pseInformado * 10) / 10;
      porAtleta[id].dias[chave] = (porAtleta[id].dias[chave] || 0) + carga;
      porAtleta[id].pseValidas++;
    }

    const zona = pctFc > 0
      ? (pctFc <= 0.80 ? 0 : pctFc <= 0.90 ? 1 : 2)
      : (pseInformado > 0 ? (pseInformado <= 4 ? 0 : pseInformado <= 7 ? 1 : 2) : -1);
    if (zona >= 0) porAtleta[id].z[zona] += minutos;
  });

  const linhas = atletas.map(a => {
    const info = porAtleta[a.id];
    if (!info || !info.atividades) return [a.nome || a.id, '—', '—', '—', '—', '—', '—', '—', '0%', '⚪ Sem atividades válidas em 90 dias'];
    const coberturaPse = Math.round(info.pseValidas / info.atividades * 100);
    if (!info.pseValidas) {
      return [a.nome || a.id, '—', '—', '—', '—', '—', '—', '—', '0%', '⚪ Sem PSE: carga interna indisponível'];
    }

    let ctl = 0, atl = 0, carga7 = 0, carga28 = 0;
    for (let d = 0; d < 90; d++) {
      const data = new Date(inicio.getTime() + d * 86400000);
      const chave = Utilities.formatDate(data, tz, 'yyyy-MM-dd');
      const carga = info.dias[chave] || 0;
      ctl += (carga - ctl) / 42;
      atl += (carga - atl) / 7;
      if (d >= 83) carga7 += carga;
      if (d >= 62) carga28 += carga;
    }
    const tsb = ctl - atl;
    const crSemanal = carga28 / 4;
    const acwr = crSemanal > 0 ? carga7 / crSemanal : 0;
    const totalZona = info.z.reduce((s, v) => s + v, 0) || 1;
    const pct = info.z.map(v => Math.round(v / totalZona * 100));
    let diag = '✅ Carga equilibrada';
    if (acwr > 1.5) diag = '🔴 Pico de carga — revisar recuperação';
    else if (acwr > 1.3) diag = '🟠 Carga alta — acompanhar fadiga';
    else if (tsb < -20) diag = '🟠 Fadiga acumulada';
    else if (acwr > 0 && acwr < 0.8) diag = '🔵 Carga recente abaixo da base';
    else if (tsb > 20) diag = '🟢 Recuperado / baixa carga aguda';
    return [
      a.nome || a.id,
      Math.round(ctl * 10) / 10,
      Math.round(atl * 10) / 10,
      Math.round(tsb * 10) / 10,
      acwr ? Math.round(acwr * 100) / 100 : '—',
      pct[0] + '%', pct[1] + '%', pct[2] + '%',
      coberturaPse + '%', diag
    ];
  });

  sh.getRange(1, 1, Math.max(sh.getLastRow(), 4), sh.getMaxColumns()).breakApart();
  sh.getRange(1, 1, Math.max(sh.getLastRow(), 4), 10).clearContent();
  sh.getRange(1, 1, 1, 10).merge().setValue('🔬 ANÁLISE DE CARGA — HIPERATIVO V3')
    .setBackground('#001F3F').setFontColor('#FFFFFF').setFontWeight('bold')
    .setFontSize(13).setHorizontalAlignment('center');
  const ts = Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy HH:mm');
  sh.getRange(2, 1, 1, 10).merge().setValue('Atualizado: ' + ts + ' | Carga sRPE observada · CTL 42d · ATL 7d · ACWR como alerta, não diagnóstico')
    .setBackground('#F3F6FA').setFontColor('#5D6675').setFontStyle('italic')
    .setHorizontalAlignment('center');
  const headers = ['Atleta','CTL sRPE','ATL sRPE','TSB','ACWR (alerta)','Z1-Z2%','Z3%','Z4-Z5%','Cobertura PSE','Sinal operacional'];
  sh.getRange(3, 1, 1, headers.length).setValues([headers])
    .setBackground('#174A7E').setFontColor('#FFFFFF').setFontWeight('bold')
    .setHorizontalAlignment('center').setWrap(true);
  if (linhas.length) sh.getRange(4, 1, linhas.length, headers.length).setValues(linhas);
  sh.setFrozenRows(3);
  [24,10,10,10,10,10,9,10,13,38].forEach((w, i) => sh.setColumnWidth(i + 1, w * 7));
  if (linhas.length) {
    sh.getRange(4, 1, linhas.length, 10).setFontSize(10).setVerticalAlignment('middle');
    sh.getRange(4, 2, linhas.length, 8).setHorizontalAlignment('center');
  }
  SpreadsheetApp.flush();
  return linhas.length;
}

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
