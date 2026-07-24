/**
 * HIPERATIVO V3 — visual operacional e atualização orientada a eventos.
 *
 * Atualiza somente camadas derivadas (PAINEL, MÉTRICAS, ANÁLISE e RANKING).
 * Não altera atividades brutas, cadastro original nem qualquer token.
 */

var VIS_CORES_ = {
  marinho: '#071A2E',
  azul: '#0B3A67',
  ciano: '#00AFC8',
  verde: '#1D9E75',
  laranja: '#FC4C02',
  roxo: '#6D5BD0',
  amarelo: '#F4B400',
  vermelho: '#D93025',
  cinza: '#5F6B7A',
  claro: '#F4F7FA',
  azulClaro: '#EAF4FB',
  verdeClaro: '#E4F5EF',
  amareloClaro: '#FFF5D6',
  vermelhoClaro: '#FCE8E6',
  branco: '#FFFFFF'
};

function aplicarAtualizacaoVisualOperacional() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  _visCriarBackupPainelUmaVez_(ss);
  _atualizarPainelInteligente_();
  aplicarTemaAbasOperacionais_();
  try { atualizarRankingSheet(); } catch (e1) { _log('SISTEMA', 'AVISO', 'aplicarAtualizacaoVisualOperacional', e1.message, ''); }
  try { atualizarAnaliseSheet(); } catch (e2) { _log('SISTEMA', 'AVISO', 'aplicarAtualizacaoVisualOperacional', e2.message, ''); }
  SpreadsheetApp.flush();
  return { ok: true, painel: 'atualizado', tema: 'aplicado' };
}

function _visCriarBackupPainelUmaVez_(ss) {
  var nome = '📊 PAINEL_BACKUP_VISUAL_20260722';
  if (ss.getSheetByName(nome)) return;
  var origem = ss.getSheetByName(H.SHEETS.PAINEL);
  if (!origem) return;
  var backup = origem.copyTo(ss).setName(nome);
  backup.hideSheet();
}

/** Disparador leve para edição humana; não é necessário instalar trigger. */
function onEdit(e) {
  try {
    if (!e || !e.range) return;
    var sh = e.range.getSheet();
    var nome = sh.getName();
    var linha = e.range.getRow();

    // A declaração "Usa Strava?" é um input humano. A conexão real é
    // recalculada imediatamente em AX/AY sem disparar métricas, ranking ou
    // outras varreduras pesadas.
    if (nome === H.SHEETS.CADASTRO && linha >= 4 &&
        e.range.getColumn() <= H.CAD.STRAVA_OK &&
        e.range.getLastColumn() >= H.CAD.STRAVA_OK) {
      var athStrava = _visAthIdDaLinha_(sh, linha);
      if (_isAthIdValido_(athStrava) &&
          typeof atualizarStatusConexaoStravaPorEdicao_ === 'function') {
        atualizarStatusConexaoStravaPorEdicao_(e);
      }
      return;
    }

    var monitoradas = {};
    monitoradas[H.SHEETS.CADASTRO] = 'cadastro_manual';
    monitoradas[H.SHEETS.FEEDBACK] = 'feedback';
    monitoradas[H.SHEETS.PLANO] = 'plano';
    monitoradas['📝 INPUT MANUAL'] = 'input_manual';
    if (!monitoradas[nome]) return;
    var athId = _visAthIdDaLinha_(sh, linha);
    sincronizarVisoesAposInput_(monitoradas[nome], athId, { permitirGlobal: false });
  } catch (erro) {
    try { _log('SISTEMA', 'AVISO', 'onEdit.visual', erro.message, ''); } catch (_) {}
  }
}

function _visAthIdDaLinha_(sh, linha) {
  if (linha < 3) return '';
  if (sh.getName() === H.SHEETS.CADASTRO) {
    return String(sh.getRange(linha, H.CAD.ID).getValue() || '').trim().toUpperCase();
  }
  var limite = Math.min(sh.getLastColumn(), 40);
  for (var hr = 1; hr <= Math.min(3, sh.getLastRow()); hr++) {
    var cab = sh.getRange(hr, 1, 1, limite).getDisplayValues()[0];
    for (var c = 0; c < cab.length; c++) {
      var h = _textoNormalizado_(cab[c]);
      if (h === 'id atleta' || h === 'ath id' || h === 'ath_id') {
        return String(sh.getRange(linha, c + 1).getValue() || '').trim().toUpperCase();
      }
    }
  }
  return '';
}

function sincronizarVisoesAposInput_(origem, athId, opcoes) {
  opcoes = opcoes || {};
  var lock = LockService.getDocumentLock();
  if (!lock.tryLock(3000)) return { ok: false, adiado: true };
  try {
    if (_isAthIdValido_(athId) && !opcoes.metricasJaAtualizadas &&
        typeof _calcularMetricasAtleta === 'function') {
      try { _calcularMetricasAtleta(athId); } catch (eM) {
        _log(athId, 'AVISO', 'sincronizarVisoesAposInput_', 'Métricas: ' + eM.message, '');
      }
    }
    _atualizarPainelInteligente_();

    if (opcoes.permitirGlobal !== false) {
      var props = PropertiesService.getScriptProperties();
      var agora = Date.now();
      var ultima = Number(props.getProperty('VIS_ULTIMA_ATUALIZACAO_GLOBAL') || 0);
      if (opcoes.forcarGlobal || !ultima || agora - ultima >= 5 * 60 * 1000) {
        props.setProperty('VIS_ULTIMA_ATUALIZACAO_GLOBAL', String(agora));
        try { atualizarRankingSheet(); } catch (eR) { _log('SISTEMA', 'AVISO', 'visual.ranking', eR.message, ''); }
        try { atualizarAnaliseSheet(); } catch (eA) { _log('SISTEMA', 'AVISO', 'visual.analise', eA.message, ''); }
      }
    }
    return { ok: true, origem: origem || '', ath_id: athId || '' };
  } finally {
    lock.releaseLock();
  }
}

function _atualizarPainelInteligente_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ws = ss.getSheetByName(H.SHEETS.PAINEL);
  var cad = ss.getSheetByName(H.SHEETS.CADASTRO);
  var ativ = ss.getSheetByName(H.SHEETS.ATIVIDADES);
  var met = ss.getSheetByName(H.SHEETS.METRICAS);
  var tok = ss.getSheetByName(H.SHEETS.TOKENS);
  if (!ws || !cad || !ativ || !met || !tok) return;

  var modelo = _visMontarModelo_(cad, ativ, met, tok);
  _visDesenharPainel_(ss, ws, modelo);
}

function _visMontarModelo_(cad, ativ, met, tok) {
  var agora = new Date();
  var hoje = new Date(agora.getFullYear(), agora.getMonth(), agora.getDate());
  var inicio7 = new Date(agora.getTime() - 7 * 86400000);
  var inicio14 = new Date(hoje.getTime() - 13 * 86400000);
  var cads = cad.getDataRange().getValues().slice(2);
  var atividades = ativ.getDataRange().getValues().slice(2).filter(function(r) {
    return _isAthIdValido_(String(r[H.ATIV.ATH_ID - 1] || '').trim()) &&
      r[H.ATIV.DATA - 1] instanceof Date && !isNaN(r[H.ATIV.DATA - 1].getTime());
  });
  var metricas = met.getDataRange().getValues().slice(2);
  var tokens = tok.getDataRange().getValues();
  var props = PropertiesService.getScriptProperties();

  var atletas = cads.map(function(r) {
    return {
      id: String(r[H.CAD.ID - 1] || '').trim().toUpperCase(),
      nome: String(r[H.CAD.NOME - 1] || '').trim(),
      status: _textoNormalizado_(r[H.CAD.STATUS - 1]),
      strava: _textoNormalizado_(r[H.CAD.STRAVA_OK - 1])
    };
  }).filter(function(a) {
    return _isAthIdValido_(a.id) && a.status !== 'inativo' && a.status !== 'cancelado';
  });
  var ativosSet = {};
  atletas.forEach(function(a) { ativosSet[a.id] = true; });

  atividades.sort(function(a, b) { return b[H.ATIV.DATA - 1] - a[H.ATIV.DATA - 1]; });
  var ultimas = {};
  var treinaram7 = {};
  var treinos7 = 0, km7 = 0, segundos7 = 0, corridaSeg = 0, corridaKm = 0;
  var tendencia = [];
  for (var d = 0; d < 14; d++) {
    var dia = new Date(inicio14.getTime() + d * 86400000);
    tendencia.push({ data: dia, chave: Utilities.formatDate(dia, 'GMT', 'yyyy-MM-dd'), atividades: 0, km: 0 });
  }
  var trendMap = {};
  tendencia.forEach(function(x) { trendMap[x.chave] = x; });

  atividades.forEach(function(r) {
    var id = String(r[H.ATIV.ATH_ID - 1] || '').trim().toUpperCase();
    var data = r[H.ATIV.DATA - 1];
    if (!ultimas[id]) ultimas[id] = data;
    var seg = _segundosAtividade_(r);
    var km = Number(r[H.ATIV.DIST_KM - 1]) || (Number(r[H.ATIV.DIST_M - 1]) || 0) / 1000;
    if (data >= inicio7 && ativosSet[id]) {
      treinos7++;
      treinaram7[id] = true;
      segundos7 += seg;
      if (_ehCorridaMetrica_(r)) {
        km7 += km;
        corridaKm += km;
        corridaSeg += seg;
      }
    }
    var chave = Utilities.formatDate(data, 'GMT', 'yyyy-MM-dd');
    if (trendMap[chave]) {
      trendMap[chave].atividades++;
      trendMap[chave].km += km;
    }
  });

  var metPorId = {};
  metricas.forEach(function(r) {
    var id = String(r[H.MET.ATH_ID - 1] || '').trim().toUpperCase();
    if (_isAthIdValido_(id)) metPorId[id] = r;
  });
  var sinais = atletas.map(function(a) {
    var r = metPorId[a.id] || [];
    var dias = r.length ? r[34] : (ultimas[a.id]
      ? Math.max(0, Math.floor((hoje - new Date(ultimas[a.id].getFullYear(), ultimas[a.id].getMonth(), ultimas[a.id].getDate())) / 86400000)) : '');
    var treinos = r.length ? Number(r[24]) || 0 : 0;
    var sinal = String(r[35] || (dias === '' ? '⚪ Sem dados' : dias >= 14 ? '🔴 Retorno após pausa' : dias >= 7 ? '🟡 Atenção' : '🟢 Em movimento'));
    var acao = String(r[36] || (dias === '' ? 'Conferir cadastro e fonte de atividades.' : dias >= 7 ? 'Contatar e entender a pausa.' : 'Manter acompanhamento.'));
    return { id: a.id, nome: a.nome || a.id, treinos: treinos, dias: dias, sinal: sinal, acao: acao };
  });
  sinais.sort(function(a, b) { return _visPrioridadeSinal_(a.sinal) - _visPrioridadeSinal_(b.sinal) || Number(b.dias || 0) - Number(a.dias || 0); });

  var tokenPorId = {};
  tokens.forEach(function(r, idx) {
    var id = String(r[H.TOK.ATH_ID - 1] || '').trim().toUpperCase();
    if (!_isAthIdValido_(id)) return;
    var exp = Number(r[H.TOK.EXPIRES - 1]) || 0;
    if (!tokenPorId[id] || exp > tokenPorId[id].exp || (exp === tokenPorId[id].exp && idx > tokenPorId[id].idx)) {
      tokenPorId[id] = { row: r, exp: exp, idx: idx };
    }
  });
  var conexoes = atletas.map(function(a) {
    var item = tokenPorId[a.id];
    var refresh = item ? String(item.row[H.TOK.REFRESH - 1] || '') : '';
    var tokenStatus = item ? _textoNormalizado_(item.row[H.TOK.STATUS - 1]) : '';
    var revogado = !!props.getProperty('STRAVA_REVOGADO_' + a.id) || tokenStatus.indexOf('revog') >= 0;
    var erroTemp = _visLerAlertaLocal_(a.id);
    if (a.strava === 'nao') return { id:a.id, nome:a.nome, tipo:'nao_usa', rotulo:'⚪ Não utiliza', detalhe:'Cadastro informado sem Strava' };
    if (revogado) return { id:a.id, nome:a.nome, tipo:'revogado', rotulo:'🔴 Desconectado', detalhe:'Revogação confirmada pela Strava' };
    if (_isRefreshTokenValido_(refresh)) {
      if (erroTemp && erroTemp.tipo === 'ERRO_TEMPORARIO') return { id:a.id, nome:a.nome, tipo:'erro', rotulo:'🟡 Falha temporária', detalhe:erroTemp.detalhe || 'Nova tentativa automática' };
      return { id:a.id, nome:a.nome, tipo:'conectado', rotulo:'🟢 Conectado', detalhe:'Refresh token protegido' };
    }
    return { id:a.id, nome:a.nome, tipo:'pendente', rotulo:'🟠 Aguardando conexão', detalhe:'Ainda não concluiu o OAuth' };
  });
  var ordemConexao = { revogado:0, erro:1, pendente:2, conectado:3, nao_usa:4 };
  conexoes.sort(function(a,b) { return ordemConexao[a.tipo] - ordemConexao[b.tipo] || a.nome.localeCompare(b.nome); });

  var recentes = atividades.slice(0, 10).map(function(r) {
    var pace = Number(r[H.ATIV.PACE_S - 1]) || 0;
    return [
      r[H.ATIV.DATA - 1],
      String(r[H.ATIV.NOME - 1] || r[H.ATIV.ATH_ID - 1] || ''),
      String(r[H.ATIV.TIPO - 1] || ''),
      Number(r[H.ATIV.DIST_KM - 1]) || (Number(r[H.ATIV.DIST_M - 1]) || 0) / 1000,
      pace ? _visPace_(pace) : '--',
      Number(r[H.ATIV.FC_MED - 1]) || ''
    ];
  });

  return {
    atualizado: agora,
    kpis: [
      { titulo:'ATLETAS ATIVOS', valor:String(atletas.length), nota:'cadastros em acompanhamento', cor:VIS_CORES_.azul },
      { titulo:'ADESÃO 7 DIAS', valor:atletas.length ? Math.round(Object.keys(treinaram7).length / atletas.length * 100) + '%' : '--', nota:Object.keys(treinaram7).length + ' de ' + atletas.length + ' treinaram', cor:VIS_CORES_.verde },
      { titulo:'ATIVIDADES 7 DIAS', valor:String(treinos7), nota:'todas as modalidades', cor:VIS_CORES_.roxo },
      { titulo:'CORRIDA 7 DIAS', valor:(Math.round(km7 * 10) / 10).toFixed(1).replace('.', ',') + ' km', nota:'distância registrada', cor:VIS_CORES_.laranja },
      { titulo:'TEMPO DE TREINO', valor:_visDuracao_(segundos7), nota:'últimos 7 dias', cor:VIS_CORES_.ciano },
      { titulo:'PACE PONDERADO', valor:corridaKm > 0 ? _visPace_(Math.round(corridaSeg / corridaKm)) : '--', nota:'min/km • corridas 7d', cor:VIS_CORES_.marinho }
    ],
    recentes: recentes,
    sinais: sinais.slice(0, 10),
    conexoes: conexoes,
    tendencia: tendencia
  };
}

function _visDesenharPainel_(ss, ws, m) {
  var c = VIS_CORES_;
  var area = ws.getRange('A1:L42');
  area.breakApart();
  area.clear({ contentsOnly: false });
  ws.getRange('A1:L1').merge().setValue('⚡ HIPERATIVO — CENTRAL DE PERFORMANCE E OPERAÇÃO')
    .setBackground(c.marinho).setFontColor(c.branco).setFontWeight('bold').setFontSize(16)
    .setHorizontalAlignment('left').setVerticalAlignment('middle');
  ws.getRange('A2:L2').merge().setValue('Treino • recuperação • aderência • conexão Strava  |  visão para decisão rápida')
    .setBackground(c.azul).setFontColor('#BFEAF2').setFontStyle('italic').setFontSize(10)
    .setHorizontalAlignment('left');
  ws.getRange('A3:L3').merge().setValue('Atualizado automaticamente em ' + Utilities.formatDate(m.atualizado, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss'))
    .setFontColor(c.cinza).setFontSize(9).setHorizontalAlignment('right');
  ws.setRowHeight(1, 42); ws.setRowHeight(2, 24); ws.setRowHeight(3, 20); ws.setRowHeight(4, 8);

  m.kpis.forEach(function(k, i) {
    var col = 1 + i * 2;
    ws.getRange(5, col, 1, 2).merge().setValue(k.titulo).setBackground(k.cor)
      .setFontColor(c.branco).setFontWeight('bold').setFontSize(8).setHorizontalAlignment('center');
    ws.getRange(6, col, 1, 2).merge().setValue(k.valor).setBackground(c.claro)
      .setFontColor(k.cor).setFontWeight('bold').setFontSize(20).setHorizontalAlignment('center');
    ws.getRange(7, col, 1, 2).merge().setValue(k.nota).setBackground(c.claro)
      .setFontColor(c.cinza).setFontSize(8).setHorizontalAlignment('center');
  });
  ws.setRowHeight(5, 24); ws.setRowHeight(6, 34); ws.setRowHeight(7, 20); ws.setRowHeight(8, 8);

  _visTituloBloco_(ws.getRange('A9:F9'), '🏃 ATIVIDADES MAIS RECENTES', c.azul);
  _visCabecalho_(ws.getRange('A10:F10'), ['Data / hora','Atleta','Tipo','Distância','Pace','FC média']);
  var recentes = [];
  for (var i = 0; i < 10; i++) recentes.push(m.recentes[i] || ['', '', '', '', '', '']);
  ws.getRange(11, 1, 10, 6).setValues(recentes).setFontSize(9).setVerticalAlignment('middle');
  ws.getRange('A11:A20').setNumberFormat('dd/MM/yy HH:mm');
  ws.getRange('D11:D20').setNumberFormat('0.0" km"');
  _visZebra_(ws, 11, 20, 1, 6);

  _visTituloBloco_(ws.getRange('H9:L9'), '🚦 SEMÁFORO DE TREINO E AÇÃO', c.marinho);
  _visCabecalho_(ws.getRange('H10:L10'), ['Atleta','Treinos 7d','Dias sem','Sinal','Próxima ação']);
  var sinais = [];
  for (var s = 0; s < 10; s++) {
    var x = m.sinais[s];
    sinais.push(x ? [x.nome, x.treinos, x.dias === '' ? '--' : x.dias, x.sinal, x.acao] : ['', '', '', '', '']);
  }
  ws.getRange(11, 8, 10, 5).setValues(sinais).setFontSize(8).setVerticalAlignment('middle').setWrap(true);
  for (var sr = 11; sr <= 20; sr++) {
    var sinal = String(ws.getRange(sr, 11).getDisplayValue() || '');
    var fundo = sinal.indexOf('🔴') >= 0 ? c.vermelhoClaro : sinal.indexOf('🟡') >= 0 ? c.amareloClaro : sinal.indexOf('🟢') >= 0 ? c.verdeClaro : c.claro;
    ws.getRange(sr, 8, 1, 5).setBackground(fundo);
  }

  var graf = ss.getSheetByName('📉 GRÁFICOS');
  if (graf) {
    var dadosTrend = [['Data','Atividades','Distância (km)']].concat(m.tendencia.map(function(t) {
      return [t.data, t.atividades, Math.round(t.km * 10) / 10];
    }));
    graf.getRange('N1:P15').clearContent().setValues(dadosTrend);
    graf.getRange('N2:N15').setNumberFormat('dd/MM');
    ws.getCharts().forEach(function(ch) { ws.removeChart(ch); });
    var chart = ws.newChart().asComboChart()
      .addRange(graf.getRange('N1:P15'))
      .setNumHeaders(1)
      .setPosition(22, 1, 0, 0)
      .setOption('title', 'Pulso dos últimos 14 dias')
      .setOption('subtitle', 'Atividades e distância registradas')
      .setOption('fontName', 'Arial')
      .setOption('backgroundColor', c.branco)
      .setOption('legend', { position: 'bottom' })
      .setOption('series', { 0: { type: 'bars', color: c.ciano }, 1: { type: 'line', targetAxisIndex: 1, color: c.laranja, lineWidth: 3 } })
      .setOption('hAxis', { format: 'dd/MM', textStyle: { fontSize: 9 } })
      .setOption('vAxes', { 0: { title: 'Atividades', minValue: 0 }, 1: { title: 'km', minValue: 0 } })
      .setOption('width', 620).setOption('height', 330).build();
    ws.insertChart(chart);
  }

  _visTituloBloco_(ws.getRange('H22:L22'), '🔌 CENTRAL STRAVA — SITUAÇÃO REAL', c.laranja);
  var cont = { conectado:0, pendente:0, erro:0, revogado:0, nao_usa:0 };
  m.conexoes.forEach(function(x) { cont[x.tipo] = (cont[x.tipo] || 0) + 1; });
  ws.getRange('H23:L23').setValues([['🟢 Conectados','🟠 Aguardando','🟡 Temporários','🔴 Revogados','⚪ Não usam']])
    .setFontWeight('bold').setFontSize(8).setHorizontalAlignment('center').setBackground(c.claro);
  ws.getRange('H24:L24').setValues([[cont.conectado,cont.pendente,cont.erro,cont.revogado,cont.nao_usa]])
    .setFontWeight('bold').setFontSize(16).setHorizontalAlignment('center');
  _visCabecalho_(ws.getRange('H26:L26'), ['Atleta','Status','Motivo','Ação','ID']);
  var problemas = m.conexoes.filter(function(x) { return x.tipo !== 'conectado' && x.tipo !== 'nao_usa'; }).slice(0, 10);
  var linhasCon = [];
  for (var q = 0; q < 10; q++) {
    var p = problemas[q];
    var acao = p ? (p.tipo === 'revogado' ? 'Reconectar' : p.tipo === 'erro' ? 'Aguardar fila' : 'Enviar link') : '';
    linhasCon.push(p ? [p.nome, p.rotulo, p.detalhe, acao, p.id] : ['', '', '', '', '']);
  }
  ws.getRange(27, 8, 10, 5).setValues(linhasCon).setFontSize(8).setWrap(true).setVerticalAlignment('middle');
  _visZebra_(ws, 27, 36, 8, 5);
  ws.getRange('H38:L39').merge().setValue('Regra de segurança: falha temporária nunca gera pedido de reconexão. Reconexão só aparece após revogação oficial ou ausência real de refresh token.')
    .setBackground(c.azulClaro).setFontColor(c.azul).setFontSize(8).setFontStyle('italic').setWrap(true).setVerticalAlignment('middle');

  [86,170,98,88,82,78,18,160,70,58,155,220].forEach(function(w, idx) { ws.setColumnWidth(idx + 1, w); });
  for (var row = 10; row <= 39; row++) ws.setRowHeight(row, row === 10 || row === 26 ? 28 : 30);
  ws.setFrozenRows(3);
  ws.setHiddenGridlines(true);
  ws.setTabColor(c.ciano);
  ws.getRange('A1:L42').setFontFamily('Arial');
}

function _visTituloBloco_(range, titulo, cor) {
  range.merge().setValue(titulo).setBackground(cor).setFontColor('#FFFFFF')
    .setFontWeight('bold').setFontSize(10).setHorizontalAlignment('left').setVerticalAlignment('middle');
}

function _visCabecalho_(range, valores) {
  range.setValues([valores]).setBackground(VIS_CORES_.azul).setFontColor('#FFFFFF')
    .setFontWeight('bold').setFontSize(8).setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true);
}

function _visZebra_(ws, inicio, fim, col, largura) {
  for (var r = inicio; r <= fim; r++) {
    ws.getRange(r, col, 1, largura).setBackground((r - inicio) % 2 ? '#F7FAFC' : '#FFFFFF');
  }
}

function _visPace_(seg) {
  seg = Math.max(0, Math.round(Number(seg) || 0));
  return Math.floor(seg / 60) + ':' + String(seg % 60).padStart(2, '0') + '/km';
}

function _visDuracao_(seg) {
  seg = Math.max(0, Math.round(Number(seg) || 0));
  var h = Math.floor(seg / 3600);
  var m = Math.floor((seg % 3600) / 60);
  return h + 'h ' + String(m).padStart(2, '0') + 'min';
}

function _visPrioridadeSinal_(sinal) {
  sinal = String(sinal || '');
  if (sinal.indexOf('🔴') >= 0) return 0;
  if (sinal.indexOf('🟡') >= 0) return 1;
  if (sinal.indexOf('⚪') >= 0) return 2;
  return 3;
}

function aplicarTemaAbasOperacionais_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var defs = [
    { nome:H.SHEETS.CADASTRO, titulo:1, cab:3, congelar:3, cor:VIS_CORES_.azul },
    { nome:'📲 WHATSAPP STRAVA', titulo:0, cab:1, congelar:1, cor:VIS_CORES_.verde },
    { nome:H.SHEETS.ATIVIDADES, titulo:1, cab:2, congelar:2, cor:VIS_CORES_.laranja },
    { nome:H.SHEETS.METRICAS, titulo:1, cab:2, congelar:2, cor:VIS_CORES_.roxo },
    { nome:H.SHEETS.PLANO, titulo:1, cab:2, congelar:2, cor:VIS_CORES_.ciano },
    { nome:H.SHEETS.FEEDBACK, titulo:1, cab:2, congelar:2, cor:VIS_CORES_.verde },
    { nome:'🔬 ANÁLISE', titulo:1, cab:3, congelar:3, cor:VIS_CORES_.roxo },
    { nome:'🏆 RANKING', titulo:1, cab:3, congelar:3, cor:VIS_CORES_.amarelo },
    { nome:'📝 INPUT MANUAL', titulo:1, cab:2, congelar:2, cor:VIS_CORES_.azul }
  ];
  defs.forEach(function(d) {
    var sh = ss.getSheetByName(d.nome);
    if (!sh) return;
    var lastRow = Math.max(sh.getLastRow(), d.cab);
    var lastCol = Math.max(sh.getLastColumn(), 1);
    sh.getRange(1, 1, lastRow, lastCol).setFontFamily('Arial').setVerticalAlignment('middle');
    if (d.titulo) sh.getRange(d.titulo, 1, 1, lastCol).setBackground(VIS_CORES_.marinho).setFontColor('#FFFFFF').setFontWeight('bold');
    sh.getRange(d.cab, 1, 1, lastCol).setBackground(d.cor).setFontColor('#FFFFFF')
      .setFontWeight('bold').setHorizontalAlignment('center').setWrap(true);
    sh.setRowHeight(d.cab, 32);
    sh.setFrozenRows(d.congelar);
    sh.setHiddenGridlines(true);
    sh.setTabColor(d.cor);
    if (lastRow > d.cab && sh.getBandings().length === 0) {
      sh.getRange(d.cab, 1, lastRow - d.cab + 1, lastCol)
        .applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false);
      sh.getRange(d.cab, 1, 1, lastCol).setBackground(d.cor).setFontColor('#FFFFFF').setFontWeight('bold');
    }
  });
  _visFormatosEspecificos_(ss);
}

function _visFormatosEspecificos_(ss) {
  var at = ss.getSheetByName(H.SHEETS.ATIVIDADES);
  if (at && at.getLastRow() >= 3) {
    at.getRange(3, H.ATIV.DATA, at.getLastRow() - 2, 1).setNumberFormat('dd/MM/yyyy HH:mm');
    at.getRange(3, H.ATIV.DIST_KM, at.getLastRow() - 2, 1).setNumberFormat('0.0');
    at.setColumnWidth(H.ATIV.NOME, 180); at.setColumnWidth(H.ATIV.TIPO, 120);
  }
  var met = ss.getSheetByName(H.SHEETS.METRICAS);
  if (met && met.getLastRow() >= 3) {
    var regras = met.getConditionalFormatRules().filter(function(regra) {
      return !regra.getRanges().some(function(r) { return r.getColumn() === 36; });
    });
    var sem = met.getRange(3, 36, Math.max(1, met.getMaxRows() - 2), 1);
    regras.push(SpreadsheetApp.newConditionalFormatRule().whenTextContains('🔴').setBackground(VIS_CORES_.vermelhoClaro).setFontColor(VIS_CORES_.vermelho).setRanges([sem]).build());
    regras.push(SpreadsheetApp.newConditionalFormatRule().whenTextContains('🟡').setBackground(VIS_CORES_.amareloClaro).setFontColor('#8A5A00').setRanges([sem]).build());
    regras.push(SpreadsheetApp.newConditionalFormatRule().whenTextContains('🟢').setBackground(VIS_CORES_.verdeClaro).setFontColor('#0B6B4B').setRanges([sem]).build());
    met.setConditionalFormatRules(regras);
    met.setColumnWidth(H.MET.NOME, 180); met.setColumnWidth(36, 180); met.setColumnWidth(37, 360);
  }
}

function registrarAlertaStravaOperacional_(athId, tipo, detalhe, notificar) {
  athId = String(athId || '').trim().toUpperCase();
  if (!_isAthIdValido_(athId)) return;
  var props = PropertiesService.getScriptProperties();
  var agora = new Date();
  var alerta = { athId:athId, tipo:String(tipo || 'ERRO_TEMPORARIO'), detalhe:String(detalhe || '').substring(0, 500), detectadoEm:agora.toISOString(), status:'aberto' };
  props.setProperty('VIS_ALERTA_STRAVA_' + athId, JSON.stringify(alerta));
  if (tipo === 'REVOGADO_CONFIRMADO') props.setProperty('STRAVA_REVOGADO_' + athId, agora.toISOString());
  if (tipo === 'ERRO_TEMPORARIO') props.setProperty('STRAVA_ERRO_TEMP_' + athId, agora.toISOString());
  try {
    if (typeof supaAtualizarStravaOk === 'function') {
      supaAtualizarStravaOk(athId, tipo === 'REVOGADO_CONFIRMADO' ? 'Reconectar' : tipo === 'PENDENTE_CADASTRO' ? 'Pendente' : 'Erro');
    }
  } catch (_) {}
  if (notificar) _visNotificarAdmin_(alerta);
}

function resolverAlertasStravaOperacionais_(athId) {
  athId = String(athId || '').trim().toUpperCase();
  if (!_isAthIdValido_(athId)) return;
  var props = PropertiesService.getScriptProperties();
  props.deleteProperty('VIS_ALERTA_STRAVA_' + athId);
  props.deleteProperty('STRAVA_ERRO_TEMP_' + athId);
  props.deleteProperty('STRAVA_REVOGADO_' + athId);
  try { if (typeof supaAtualizarStravaOk === 'function') supaAtualizarStravaOk(athId, 'Conectado'); } catch (_) {}
}

function _visLerAlertaLocal_(athId) {
  try {
    var raw = PropertiesService.getScriptProperties().getProperty('VIS_ALERTA_STRAVA_' + athId);
    return raw ? JSON.parse(raw) : null;
  } catch (_) { return null; }
}

function _visNotificarAdmin_(alerta) {
  var props = PropertiesService.getScriptProperties();
  var chave = 'VIS_EMAIL_' + alerta.athId + '_' + alerta.tipo;
  var ultima = Number(props.getProperty(chave) || 0);
  if (ultima && Date.now() - ultima < 12 * 60 * 60 * 1000) return;
  var nome = typeof _getNomeAtleta === 'function' ? _getNomeAtleta(alerta.athId) : alerta.athId;
  var assunto = alerta.tipo === 'REVOGADO_CONFIRMADO'
    ? '🔴 [Hiperativo] Strava desconectado — ' + nome
    : alerta.tipo === 'PENDENTE_CADASTRO'
      ? '🟠 [Hiperativo] Novo atleta aguardando Strava — ' + nome
      : '🟡 [Hiperativo] Falha temporária no Strava — ' + nome;
  var acao = alerta.tipo === 'REVOGADO_CONFIRMADO'
    ? 'A revogação foi confirmada pela própria Strava. Somente neste caso a reconexão deve ser solicitada.'
    : alerta.tipo === 'PENDENTE_CADASTRO'
      ? 'O cadastro foi concluído, mas a autorização do Strava ainda não foi finalizada. O link já está na fila de WhatsApp.'
      : 'O refresh token foi preservado. Aguarde a fila automática; não peça reconexão ao atleta.';
  MailApp.sendEmail({
    to: props.getProperty('ADMIN_EMAIL') || 'contato@ghiperativo.com.br',
    subject: assunto,
    body: assunto + '\n\nAtleta: ' + nome + ' (' + alerta.athId + ')\nDetalhe: ' + alerta.detalhe + '\n\n' + acao + '\n\nAbra o PAINEL do HIPERATIVO V3 para acompanhar.',
    name: 'Monitor HIPERATIVO V3'
  });
  props.setProperty(chave, String(Date.now()));
}
