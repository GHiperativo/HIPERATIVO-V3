/**
 * HIPERATIVO V3 — Métricas clínicas-operacionais por atleta
 *
 * Princípios:
 * - pace é calculado por tempo/distância e não pela média simples dos treinos;
 * - VO2 não é inventado a partir do perfil ou de um único pace;
 * - sRPE só existe quando o atleta informou PSE/RPE;
 * - perfil, frequência e objetivo entram como contexto, nunca como medição;
 * - zonas sem teste validado são hipóteses provisórias e exigem base mínima.
 */

const METRICAS_HEADERS_ = [
  'ID Atleta','Nome Atleta','Atualizado em','VO₂ (teste validado)',
  'Pace Ponderado (s/km)','Pace Rápido P20 (s/km)','Pace Lento P80 (s/km)',
  'FC Máx. Observada','FC Média Ponderada','Volume Médio Semanal 28d (km)',
  'Z1 Lento','Z1 Rápido','Z2 Lento','Z2 Rápido',
  'Z3 Lento','Z3 Rápido','Z4 Lento','Z5 Mín',
  'Perfil Manual','Frequência Planejada','Intensidade Manual',
  'Origem dos Dados','Confiança','Observações',
  'Treinos 7d','Treinos 28d','Volume Corrida 7d (km)','Volume Corrida 28d (km)',
  'Duração 7d (min)','Carga sRPE 7d','Cobertura PSE 28d (%)',
  'Pace Mediano 28d (s/km)','Tendência Pace 14d (%)','Última Atividade',
  'Dias sem Treino','Semáforo','Decisão de Carga','Qualidade dos Dados (%)',
  'Objetivo do Cadastro','Saúde / Lesões Declaradas'
];

const METRICAS_NOTAS_ = [
  'Identificador interno estável do atleta.',
  'Nome vindo do cadastro.',
  'Data e hora do último recálculo.',
  'Preencher somente com teste de campo/laboratório validado. O sistema não estima VO₂ por perfil ou por um treino isolado.',
  'Tempo total em movimento dividido pela distância total das corridas válidas dos últimos 28 dias.',
  'Percentil 20 dos paces por atividade. É mais robusto que usar um único treino extremo.',
  'Percentil 80 dos paces por atividade. É mais robusto que usar o treino mais lento isolado.',
  'Percentil 95 das FC máximas observadas (máximo quando há menos de 3 registros). Não é FC máxima fisiológica testada.',
  'FC média ponderada pelo tempo em movimento.',
  'Distância de corrida dos últimos 28 dias dividida por quatro semanas.',
  'Hipótese provisória; confirmar com pace, FC, PSE e contexto.',
  'Hipótese provisória; confirmar com pace, FC, PSE e contexto.',
  'Hipótese provisória; confirmar com pace, FC, PSE e contexto.',
  'Hipótese provisória; confirmar com pace, FC, PSE e contexto.',
  'Hipótese provisória; confirmar com pace, FC, PSE e contexto.',
  'Hipótese provisória; confirmar com pace, FC, PSE e contexto.',
  'Hipótese provisória; confirmar com pace, FC, PSE e contexto.',
  'Hipótese provisória; confirmar com pace, FC, PSE e contexto.',
  'Contexto manual preservado; não gera números fisiológicos.',
  'Frequência prevista no cadastro; não substitui aderência observada.',
  'Contexto manual preservado; não gera números fisiológicos.',
  'Resumo da fonte efetivamente usada no cálculo.',
  'Confiança derivada da completude e quantidade dos dados.',
  'Limitações e critérios usados.',
  'Número de atividades válidas de todas as modalidades nos últimos 7 dias.',
  'Número de atividades válidas de todas as modalidades nos últimos 28 dias.',
  'Distância de corrida nos últimos 7 dias.',
  'Distância de corrida nos últimos 28 dias.',
  'Tempo em movimento de todas as modalidades nos últimos 7 dias.',
  'Soma de duração em minutos × PSE para sessões com PSE informado.',
  'Percentual das atividades de 28 dias que têm PSE informado.',
  'Mediana dos paces de corrida válidos dos últimos 28 dias.',
  'Variação do pace ponderado: últimos 14 dias versus 14 dias anteriores. Positivo = mais rápido.',
  'Data da atividade mais recente de qualquer modalidade.',
  'Dias completos desde a última atividade.',
  'Sinal operacional. Não é diagnóstico médico.',
  'Regra semanal baseada em treinos observados e recuperação disponível.',
  'Completude da base: quantidade, distância/tempo, FC e PSE.',
  'Objetivo declarado pelo atleta no cadastro.',
  'Condições e lesões declaradas no cadastro; revisar antes de prescrever.'
];

function recalcularMetricasAposAtividade(athId) {
  if (!_isAthIdValido_(athId)) return;
  try {
    _calcularMetricasAtleta(athId);
    _log(athId, 'INFO', 'recalcularMetricasAposAtividade', 'Métricas atualizadas após nova atividade.', '');
  } catch (e) {
    _log(athId, 'AVISO', 'recalcularMetricasAposAtividade', 'Erro ao recalcular: ' + e.message, '');
  }
}

function calcularMetricasTodos(silencioso) {
  const contexto = _criarContextoMetricas_();
  if (!contexto) return 0;
  _garantirEstruturaMetricas_(contexto.wsMet);
  _garantirEstruturaFeedbackEstudos_();

  const linhas = contexto.atletas.map(a => _montarLinhaMetrica_(a, contexto));
  const lastRow = contexto.wsMet.getLastRow();
  if (lastRow >= 3) {
    contexto.wsMet.getRange(3, 1, lastRow - 2, METRICAS_HEADERS_.length).clearContent();
  }
  if (linhas.length) {
    contexto.wsMet.getRange(3, 1, linhas.length, METRICAS_HEADERS_.length).setValues(linhas);
  }
  _formatarMetricas_(contexto.wsMet, linhas.length);
  SpreadsheetApp.flush();

  _log('SISTEMA', 'INFO', 'calcularMetricasTodos', linhas.length + ' atleta(s) recalculados sem estimativas por perfil.', '');
  if (!silencioso) {
    try {
      SpreadsheetApp.getUi().alert('✅ Métricas atualizadas',
        linhas.length + ' atletas processados. Medições ausentes permanecem em branco.',
        SpreadsheetApp.getUi().ButtonSet.OK);
    } catch (_) {}
  }
  return linhas.length;
}

function _calcularMetricasAtleta(athId) {
  const contexto = _criarContextoMetricas_();
  if (!contexto) return;
  _garantirEstruturaMetricas_(contexto.wsMet);
  const atleta = contexto.atletas.find(a => a.id === String(athId).trim().toUpperCase());
  if (!atleta) return;
  const linha = _montarLinhaMetrica_(atleta, contexto);
  const dados = contexto.wsMet.getDataRange().getValues();
  let destino = -1;
  for (let i = 2; i < dados.length; i++) {
    if (String(dados[i][H.MET.ATH_ID - 1] || '').trim().toUpperCase() === atleta.id) {
      destino = i + 1;
      break;
    }
  }
  if (destino < 0) destino = _primeiraLinhaVazia(contexto.wsMet, H.MET.ATH_ID);
  contexto.wsMet.getRange(destino, 1, 1, METRICAS_HEADERS_.length).setValues([linha]);
  _formatarMetricas_(contexto.wsMet, Math.max(contexto.wsMet.getLastRow() - 2, 1));
}

function _criarContextoMetricas_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const wsCad = ss.getSheetByName(H.SHEETS.CADASTRO);
  const wsAtiv = ss.getSheetByName(H.SHEETS.ATIVIDADES);
  const wsMet = ss.getSheetByName(H.SHEETS.METRICAS);
  if (!wsCad || !wsAtiv || !wsMet) return null;

  const agora = new Date();
  const inicio28 = new Date(agora.getTime() - 28 * 86400000);
  const inicio14 = new Date(agora.getTime() - 14 * 86400000);
  const inicio7 = new Date(agora.getTime() - 7 * 86400000);

  const metricasAtuais = {};
  wsMet.getDataRange().getValues().slice(2).forEach(r => {
    const id = String(r[H.MET.ATH_ID - 1] || '').trim().toUpperCase();
    if (_isAthIdValido_(id)) metricasAtuais[id] = r;
  });

  const atletas = wsCad.getDataRange().getValues().slice(2).map(r => ({
    id: String(r[H.CAD.ID - 1] || '').trim().toUpperCase(),
    nome: String(r[H.CAD.NOME - 1] || '').trim(),
    nivel: String(r[H.CAD.NIVEL - 1] || '').trim(),
    objetivo: String(r[H.CAD.OBJ - 1] || '').trim(),
    frequencia: String(r[H.CAD.FREQ - 1] || '').trim(),
    saude: String(r[H.CAD.SAUDE - 1] || '').trim(),
    lesao: String(r[H.CAD.LESAO - 1] || '').trim(),
    usaStrava: String(r[H.CAD.STRAVA_OK - 1] || '').trim(),
    status: String(r[H.CAD.STATUS - 1] || '').trim().toLowerCase()
  })).filter(a => _isAthIdValido_(a.id) && a.status !== 'inativo');

  const porAtleta = {};
  const ultimaAtividade = {};
  atletas.forEach(a => { porAtleta[a.id] = []; });
  wsAtiv.getDataRange().getValues().slice(2).forEach(r => {
    const id = String(r[H.ATIV.ATH_ID - 1] || '').trim().toUpperCase();
    const data = r[H.ATIV.DATA - 1];
    if (!porAtleta[id] || !(data instanceof Date) || isNaN(data.getTime()) || data > new Date(agora.getTime() + 86400000)) return;
    if (!ultimaAtividade[id] || data > ultimaAtividade[id]) ultimaAtividade[id] = data;
    if (data >= inicio28) porAtleta[id].push(r);
  });

  return {
    ss, wsCad, wsAtiv, wsMet, agora, inicio28, inicio14, inicio7,
    atletas, porAtleta, ultimaAtividade, metricasAtuais,
    feedback: _lerFeedbackRecente_(ss)
  };
}

function _montarLinhaMetrica_(atleta, ctx) {
  const todas = (ctx.porAtleta[atleta.id] || []).filter(r => _segundosAtividade_(r) > 0);
  const ult7 = todas.filter(r => r[H.ATIV.DATA - 1] >= ctx.inicio7);
  const corridas = todas.filter(_ehCorridaMetrica_).map(r => _normalizarCorridaMetrica_(r)).filter(Boolean);
  const corr7 = corridas.filter(x => x.data >= ctx.inicio7);
  const corr14 = corridas.filter(x => x.data >= ctx.inicio14);
  const corr14Anterior = corridas.filter(x => x.data < ctx.inicio14);

  const dist28 = _soma_(corridas.map(x => x.distKm));
  const dist7 = _soma_(corr7.map(x => x.distKm));
  const seg28 = _soma_(corridas.map(x => x.seg));
  const pacePond = dist28 > 0 ? Math.round(seg28 / dist28) : 0;
  const paces = corridas.map(x => x.pace).filter(v => v > 0).sort((a, b) => a - b);
  const paceRap = paces.length ? Math.round(_percentil_(paces, 0.20)) : 0;
  const paceLento = paces.length ? Math.round(_percentil_(paces, 0.80)) : 0;
  const paceMediano = paces.length ? Math.round(_percentil_(paces, 0.50)) : 0;

  const fcMaxArr = corridas.map(x => x.fcMax).filter(v => v >= 60 && v <= 240).sort((a, b) => a - b);
  const fcMaxObs = fcMaxArr.length ? Math.round(_percentil_(fcMaxArr, fcMaxArr.length < 3 ? 1 : 0.95)) : 0;
  const fcComTempo = corridas.filter(x => x.fcMed >= 60 && x.fcMed <= 230 && x.seg > 0);
  const fcPeso = _soma_(fcComTempo.map(x => x.seg));
  const fcMedPond = fcPeso ? Math.round(_soma_(fcComTempo.map(x => x.fcMed * x.seg)) / fcPeso) : 0;

  const pace14 = _pacePonderado_(corr14);
  const pace14Ant = _pacePonderado_(corr14Anterior);
  const tendencia = pace14 > 0 && pace14Ant > 0
    ? Math.round(((pace14Ant - pace14) / pace14Ant) * 1000) / 10 : '';

  const dur7Min = Math.round(_soma_(ult7.map(_segundosAtividade_)) / 60);
  const pse28 = todas.filter(r => {
    const v = Number(r[H.ATIV.PSE - 1]);
    return v >= 1 && v <= 10;
  });
  const carga7 = Math.round(ult7.reduce((s, r) => {
    const pse = Number(r[H.ATIV.PSE - 1]);
    return s + (pse >= 1 && pse <= 10 ? _segundosAtividade_(r) / 60 * pse : 0);
  }, 0));
  const coberturaPse = todas.length ? Math.round(pse28.length / todas.length * 100) : 0;
  const coberturaFc = corridas.length ? Math.round(fcComTempo.length / corridas.length * 100) : 0;
  const completudePace = corridas.length ? Math.round(corridas.filter(x => x.distKm > 0 && x.seg > 0).length / corridas.length * 100) : 0;

  const ultima = ctx.ultimaAtividade[atleta.id] || '';
  const diasSem = ultima ? Math.max(0, Math.floor((ctx.agora.getTime() - ultima.getTime()) / 86400000)) : '';
  const feedback = ctx.feedback[atleta.id] || null;

  let qualidade = 0;
  if (todas.length) qualidade += Math.min(25, todas.length * 5);
  if (corridas.length) qualidade += Math.min(25, corridas.length * 5);
  qualidade += Math.round(completudePace * 0.20);
  qualidade += Math.round(coberturaFc * 0.15);
  qualidade += Math.round(coberturaPse * 0.15);
  qualidade = Math.min(100, qualidade);

  const confianca = qualidade >= 80 ? 'Alta' : qualidade >= 55 ? 'Média' : qualidade > 0 ? 'Baixa' : 'Indisponível';
  let origem = '⚠️ Sem atividades válidas em 28 dias';
  if (_textoNormalizado_(atleta.usaStrava) === 'nao' && !todas.length) origem = 'Cadastro informa: não usa Strava';
  else if (todas.length && !corridas.length) origem = todas.length + ' atividade(s), sem corrida válida para pace';
  else if (corridas.length) origem = corridas.length + ' corrida(s) real(is) em 28 dias';

  const zonas = corridas.length >= 6 ? _calcularZonasPace_(paceMediano || pacePond, paceRap) : ['', '', '', '', '', '', '', ''];
  const avaliacao = _avaliarSemaforoMetrica_(atleta, todas, ult7, tendencia, coberturaPse, feedback, diasSem);
  const atual = ctx.metricasAtuais[atleta.id] || [];
  const perfil = String(atual[H.MET.PERFIL_MAN - 1] || atleta.nivel || '').trim();
  const frequencia = String(atual[H.MET.VOLUME_MAN - 1] || atleta.frequencia || '').trim();
  const intensidade = String(atual[H.MET.INTENS_MAN - 1] || '').trim();

  const observacoes = [];
  observacoes.push('VO₂ em branco até existir teste validado.');
  if (corridas.length < 3) observacoes.push('Base de pace insuficiente (mínimo operacional: 3 corridas).');
  else if (corridas.length < 6) observacoes.push('Pace real disponível; zonas aguardam 6 corridas válidas.');
  else observacoes.push('Zonas são provisórias e devem ser cruzadas com FC, PSE e contexto.');
  if (coberturaPse < 50 && todas.length) observacoes.push('PSE pouco preenchido (' + coberturaPse + '%).');
  if (feedback && feedback.pse >= 0) observacoes.push('Último PSE do feedback: ' + feedback.pse + '/10.');
  if (feedback && feedback.fcRepouso > 0) {
    observacoes.push('FC de repouso informada: ' + feedback.fcRepouso + ' bpm' +
      (feedback.fcRepousoAlto ? ' (acima da mediana recente).' : '.'));
  }

  return [
    atleta.id, atleta.nome, new Date(), '',
    pacePond || '', paceRap || '', paceLento || '',
    fcMaxObs || '', fcMedPond || '', dist28 ? Math.round(dist28 / 4 * 10) / 10 : '',
    zonas[0], zonas[1], zonas[2], zonas[3], zonas[4], zonas[5], zonas[6], zonas[7],
    perfil, frequencia, intensidade, origem, confianca, observacoes.join(' '),
    ult7.length, todas.length, Math.round(dist7 * 10) / 10, Math.round(dist28 * 10) / 10,
    dur7Min, carga7 || '', coberturaPse,
    paceMediano || '', tendencia, ultima,
    diasSem, avaliacao.semaforo, avaliacao.decisao, qualidade,
    atleta.objetivo, [atleta.saude, atleta.lesao].filter(Boolean).join(' | ')
  ];
}

function _avaliarSemaforoMetrica_(atleta, todas, ult7, tendencia, coberturaPse, feedback, diasSem) {
  const naoUsaStrava = _textoNormalizado_(atleta.usaStrava) === 'nao';
  if (!todas.length) {
    if (!naoUsaStrava && diasSem !== '' && diasSem > 14) {
      return { semaforo: '🔴 Retorno após pausa', decisao: 'Retornar abaixo da última carga consolidada; sem progressão nesta semana.' };
    }
    return {
      semaforo: naoUsaStrava ? '⚪ Avaliação manual' : '⚪ Dados insuficientes',
      decisao: 'Retomada/diagnóstico; não progredir carga sem observação recente.'
    };
  }
  if (feedback && feedback.alerta) {
    return { semaforo: '🔴 Reduzir / adaptar', decisao: 'Dor ou alerta físico registrado; reduzir/adaptar e revisar antes do próximo estímulo.' };
  }
  if (feedback && feedback.atencao) {
    return {
      semaforo: '🟡 Recuperação em atenção',
      decisao: 'Manter ou reduzir; PSE alto, recuperação ruim ou FC de repouso elevada no feedback recente.'
    };
  }
  if (diasSem !== '' && diasSem > 14) {
    return { semaforo: '🔴 Retorno após pausa', decisao: 'Retornar abaixo da última carga consolidada; sem progressão nesta semana.' };
  }
  if (ult7.length === 0) {
    return { semaforo: '⚪ Retomada', decisao: 'Nenhum treino em 7 dias; reiniciar sem progressão.' };
  }
  if (ult7.length === 1) {
    return { semaforo: '🟡 Atenção', decisao: 'Manter ou reduzir; confirmar recuperação e motivo da baixa frequência.' };
  }
  if (ult7.length === 2) {
    return { semaforo: '🟡 Controle', decisao: 'Manter; progressão leve somente se dor ausente e recuperação boa.' };
  }
  if (feedback && feedback.recuperacaoBoa && coberturaPse >= 50 && !(Number(tendencia) < -5)) {
    return { semaforo: '🟢 Progressão possível', decisao: 'Pode progredir 5–10% em uma variável, preservando recuperação.' };
  }
  return {
    semaforo: '🟡 Dados de recuperação incompletos',
    decisao: 'Manter carga; coletar PSE, dor, sono e fadiga antes de progredir.'
  };
}

function _lerFeedbackRecente_(ss) {
  const sh = ss.getSheetByName(H.SHEETS.FEEDBACK);
  const saida = {};
  if (!sh || sh.getLastRow() < 3) return saida;
  const valores = sh.getDataRange().getValues();
  const headers = valores[1].map(_textoNormalizado_);
  const idxId = _indiceCabecalho_(headers, ['id atleta', 'ath id']);
  const idxData = _indiceCabecalho_(headers, ['data da resposta', 'data resposta', 'data semana', 'semana']);
  const idxSono = _indiceCabecalho_(headers, ['qualidade do sono', 'sono']);
  const idxFadiga = _indiceCabecalho_(headers, ['fadiga']);
  const idxAlerta = _indiceCabecalho_(headers, ['alertas fisicos', 'alerta fisico']);
  const idxDor = _indiceCabecalho_(headers, ['dor 0-10', 'dor (0-10)', 'dor']);
  const idxPse = _indiceCabecalho_(headers, ['pse/rpe ultimo treino (0-10)', 'pse/rpe ultimo treino 0-10', 'pse/rpe ultimo treino', 'pse', 'rpe']);
  const idxFcRepouso = _indiceCabecalho_(headers, ['fc repouso', 'frequencia cardiaca repouso']);
  if (idxId < 0) return saida;

  const fcHistorico = {};
  valores.slice(2).forEach(function(r) {
    const id = String(r[idxId] || '').trim().toUpperCase();
    const fc = idxFcRepouso >= 0 ? Number(r[idxFcRepouso]) : NaN;
    const data = idxData >= 0 && r[idxData] instanceof Date ? r[idxData] : new Date(0);
    if (_isAthIdValido_(id) && fc >= 30 && fc <= 220) {
      if (!fcHistorico[id]) fcHistorico[id] = [];
      fcHistorico[id].push({ fc: fc, data: data });
    }
  });

  valores.slice(2).forEach(r => {
    const id = String(r[idxId] || '').trim().toUpperCase();
    if (!_isAthIdValido_(id)) return;
    const data = idxData >= 0 && r[idxData] instanceof Date ? r[idxData] : new Date(0);
    if (saida[id] && saida[id].data > data) return;
    const alertaTxt = idxAlerta >= 0 ? _textoNormalizado_(r[idxAlerta]) : '';
    const dor = idxDor >= 0 && r[idxDor] !== '' ? Number(r[idxDor]) : NaN;
    const sono = idxSono >= 0 ? _textoNormalizado_(r[idxSono]) : '';
    const fadigaRaw = idxFadiga >= 0 ? r[idxFadiga] : '';
    const fadiga = Number(fadigaRaw);
    const pseRaw = idxPse >= 0 ? r[idxPse] : '';
    const pse = pseRaw !== '' ? Number(pseRaw) : NaN;
    const fcRepousoRaw = idxFcRepouso >= 0 ? r[idxFcRepouso] : '';
    const fcRepouso = fcRepousoRaw !== '' ? Number(fcRepousoRaw) : NaN;
    const anteriores = (fcHistorico[id] || [])
      .filter(function(x) { return x.data < data; })
      .sort(function(a, b) { return b.data - a.data; })
      .slice(0, 7)
      .map(function(x) { return x.fc; })
      .sort(function(a, b) { return a - b; });
    const medianaFc = anteriores.length >= 3 ? _percentil_(anteriores, 0.5) : 0;
    const fcRepousoAlto = !isNaN(fcRepouso) && medianaFc > 0 &&
      fcRepouso >= medianaFc + 5 && fcRepouso >= medianaFc * 1.08;
    const semAlerta = !alertaTxt || /^(nao|nenhum|sem dor|ok)$/.test(alertaTxt);
    const sonoBom = /excelente|bom|boa|7h|8h/.test(sono);
    const fadigaBoa = fadigaRaw !== '' && !isNaN(fadiga) ? fadiga <= 4 : /baixa|leve|boa/.test(_textoNormalizado_(fadigaRaw));
    const pseAlto = !isNaN(pse) && pse >= 8;
    const sonoRuim = /ruim|menos 5h/.test(sono);
    const fadigaAlta = fadigaRaw !== '' && !isNaN(fadiga) ? fadiga >= 7 : /alta|intensa/.test(_textoNormalizado_(fadigaRaw));
    saida[id] = {
      data,
      alerta: (!isNaN(dor) && dor > 0) || !semAlerta,
      atencao: pseAlto || sonoRuim || fadigaAlta || fcRepousoAlto,
      recuperacaoBoa: semAlerta && (isNaN(dor) || dor === 0) && sonoBom && fadigaBoa &&
        !pseAlto && !fcRepousoAlto,
      pse: isNaN(pse) ? -1 : pse,
      fcRepouso: isNaN(fcRepouso) ? 0 : fcRepouso,
      fcRepousoMediana: medianaFc || 0,
      fcRepousoAlto: fcRepousoAlto
    };
  });
  return saida;
}

function _garantirEstruturaFeedbackEstudos_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ws = ss.getSheetByName(H.SHEETS.FEEDBACK);
  if (!ws) return;
  const extras = ['PSE/RPE Último Treino (0-10)','Dor (0-10)','Local da Dor','Sono (horas)','FC Repouso','Data da Resposta'];
  const total = 14 + extras.length;
  if (ws.getMaxColumns() < total) ws.insertColumnsAfter(ws.getMaxColumns(), total - ws.getMaxColumns());
  ws.getRange(1, 1, 1, ws.getMaxColumns()).breakApart();
  ws.getRange(1, 1, 1, total).merge().setValue('💬 FEEDBACK SEMANAL — CABEÇA, CORAÇÃO E CORPO');
  ws.getRange(2, 15, 1, extras.length).setValues([extras])
    .setBackground('#F9A825').setFontColor('#111111').setFontWeight('bold')
    .setHorizontalAlignment('center').setWrap(true);
  try { _dropdown(ws, 15, 3, 502, ['0','1','2','3','4','5','6','7','8','9','10']); } catch (_) {}
  try { _dropdown(ws, 16, 3, 502, ['0','1','2','3','4','5','6','7','8','9','10']); } catch (_) {}
  ws.getRange(3, 18, Math.max(1, ws.getMaxRows() - 2), 1).setNumberFormat('0.0');
  ws.getRange(3, 20, Math.max(1, ws.getMaxRows() - 2), 1).setNumberFormat('dd/mm/yyyy hh:mm');
  [16,12,18,14,12,17].forEach((w, i) => ws.setColumnWidth(15 + i, w * 7));
}

function _garantirEstruturaMetricas_(ws) {
  if (ws.getMaxColumns() < METRICAS_HEADERS_.length) {
    ws.insertColumnsAfter(ws.getMaxColumns(), METRICAS_HEADERS_.length - ws.getMaxColumns());
  }
  // Uma faixa mesclada no título atravessa todas as colunas; o Google Sheets
  // não permite congelar apenas parte dessa faixa.
  ws.setFrozenColumns(0);
  ws.getRange(1, 1, 1, ws.getMaxColumns()).breakApart();
  ws.getRange(1, 1, 1, METRICAS_HEADERS_.length).merge()
    .setValue('📈 MÉTRICAS E DECISÃO DE CARGA — HIPERATIVO V3')
    .setBackground('#001F3F').setFontColor('#FFFFFF').setFontWeight('bold')
    .setFontSize(13).setHorizontalAlignment('center');
  ws.getRange(2, 1, 1, METRICAS_HEADERS_.length).setValues([METRICAS_HEADERS_])
    .setNotes([METRICAS_NOTAS_])
    .setBackground('#174A7E').setFontColor('#FFFFFF').setFontWeight('bold')
    .setHorizontalAlignment('center').setWrap(true);
  ws.setFrozenRows(2);
  ws.setHiddenGridlines(true);
  try { _dropdown(ws, H.MET.PERFIL_MAN, 3, 1000, ['Iniciante','Intermediário','Avançado','Retorno/lesão']); } catch (_) {}
  try { _dropdown(ws, H.MET.INTENS_MAN, 3, 1000, ['Leve','Moderado','Forte','Competitivo']); } catch (_) {}
}

function _formatarMetricas_(ws, quantidade) {
  const n = Math.max(1, quantidade);
  ws.setRowHeight(1, 34);
  ws.setRowHeight(2, 44);
  ws.getRange(3, 1, n, METRICAS_HEADERS_.length)
    .setFontSize(10).setVerticalAlignment('middle');
  ws.getRange(3, 3, n, 1).setNumberFormat('dd/mm/yyyy hh:mm');
  ws.getRange(3, 34, n, 1).setNumberFormat('dd/mm/yyyy');
  ws.getRange(3, 5, n, 6).setNumberFormat('0.0');
  ws.getRange(3, 25, n, 8).setNumberFormat('0.0');
  ws.getRange(3, 33, n, 1).setNumberFormat('0.0;-0.0');
  ws.getRange(3, 35, n, 1).setNumberFormat('0');
  ws.getRange(3, 38, n, 1).setNumberFormat('0');

  const larguras = [12,24,17,17,18,19,18,17,18,21,11,11,11,11,11,11,11,11,16,18,17,27,13,48,12,13,20,21,16,15,18,20,19,15,14,24,48,19,30,36];
  larguras.forEach((w, i) => ws.setColumnWidth(i + 1, w * 7));
  ws.getRange(3, 22, n, 3).setWrap(true);
  ws.getRange(3, 36, n, 5).setWrap(true);

  const semaforos = ws.getRange(3, 36, n, 1).getDisplayValues();
  const fundos = semaforos.map(r => {
    const v = String(r[0] || '');
    if (v.indexOf('🟢') >= 0) return ['#D9EAD3'];
    if (v.indexOf('🟡') >= 0) return ['#FFF2CC'];
    if (v.indexOf('🔴') >= 0) return ['#F4CCCC'];
    return ['#E7E6E6'];
  });
  ws.getRange(3, 36, n, 1).setBackgrounds(fundos).setFontWeight('bold');
}

function _ehCorridaMetrica_(r) {
  const t = _textoNormalizado_(r[H.ATIV.TIPO - 1]);
  return t.indexOf('corrida') >= 0 || t.indexOf('trail') >= 0 || t.indexOf('esteira') >= 0;
}

function _normalizarCorridaMetrica_(r) {
  const seg = _segundosAtividade_(r);
  let distKm = Number(r[H.ATIV.DIST_KM - 1]) || 0;
  if (!distKm) {
    const distM = Number(r[H.ATIV.DIST_M - 1]) || 0;
    if (distM > 0) distKm = distM / 1000;
  }
  let pace = distKm > 0 && seg > 0 ? seg / distKm : 0;
  if (!(pace > 120 && pace < 1800)) {
    const p1 = Number(r[H.ATIV.PACE_S - 1]);
    const p2 = Number(r[H.ATIV.PACE_FMT - 1]);
    pace = p1 > 120 && p1 < 1800 ? p1 : (p2 > 120 && p2 < 1800 ? p2 : 0);
  }
  if (!(pace > 120 && pace < 1800) || !(distKm > 0)) return null;
  return {
    data: r[H.ATIV.DATA - 1], seg: seg || pace * distKm, distKm, pace,
    fcMed: Number(r[H.ATIV.FC_MED - 1]) || 0,
    fcMax: Number(r[H.ATIV.FC_MAX - 1]) || 0
  };
}

function _segundosAtividade_(r) {
  let v = Number(r[H.ATIV.MOV_S - 1]) || Number(r[H.ATIV.TOTAL_S - 1]) || 0;
  if (v > 0 && v < 2) v *= 86400;
  return v > 0 && v < 7 * 86400 ? v : 0;
}

function _pacePonderado_(corridas) {
  const km = _soma_(corridas.map(x => x.distKm));
  return km > 0 ? _soma_(corridas.map(x => x.seg)) / km : 0;
}

function _soma_(arr) {
  return arr.reduce((s, v) => s + (Number(v) || 0), 0);
}

function _percentil_(ordenados, p) {
  if (!ordenados.length) return 0;
  if (p <= 0) return ordenados[0];
  if (p >= 1) return ordenados[ordenados.length - 1];
  const pos = (ordenados.length - 1) * p;
  const base = Math.floor(pos);
  const resto = pos - base;
  return ordenados[base + 1] === undefined
    ? ordenados[base]
    : ordenados[base] + resto * (ordenados[base + 1] - ordenados[base]);
}

function _calcularZonasPace_(paceRef, paceRap) {
  if (!paceRef || paceRef <= 0) return ['', '', '', '', '', '', '', ''];
  const fmt = s => {
    const total = Math.max(1, Math.round(s));
    return Math.floor(total / 60) + ':' + String(total % 60).padStart(2, '0');
  };
  return [
    fmt(paceRef * 1.20), fmt(paceRef * 1.08),
    fmt(paceRef * 1.04), fmt(paceRef * 0.96),
    fmt(paceRef * 0.94), fmt(paceRef * 0.87),
    fmt(paceRef * 0.84), fmt(paceRap || paceRef * 0.80)
  ];
}

function _indiceCabecalho_(headersNormalizados, candidatos) {
  const alvos = candidatos.map(_textoNormalizado_);
  for (let i = 0; i < headersNormalizados.length; i++) {
    if (alvos.indexOf(headersNormalizados[i]) >= 0) return i;
  }
  return -1;
}

function _textoNormalizado_(valor) {
  return String(valor || '').trim().toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '');
}

function _primeiraLinhaVazia(ws, col) {
  const lastRow = Math.max(ws.getLastRow(), 3);
  const values = ws.getRange(3, col, lastRow - 2, 1).getValues();
  for (let i = 0; i < values.length; i++) if (!values[i][0]) return i + 3;
  return lastRow + 1;
}

function _getRegistroMetrica(wsMet, athId) {
  const rows = wsMet.getDataRange().getValues();
  for (let i = 2; i < rows.length; i++) {
    if (String(rows[i][H.MET.ATH_ID - 1]) === athId) return rows[i];
  }
  return [];
}

// Mantida por compatibilidade com rotinas antigas. Não produz números fisiológicos.
function _resolverPerfilManual(athId, registroAtual) {
  return {
    perfil: String(registroAtual[H.MET.PERFIL_MAN - 1] || _getNivelAtleta(athId) || '').trim(),
    volume: String(registroAtual[H.MET.VOLUME_MAN - 1] || _getFreqAtleta(athId) || '').trim(),
    intensidade: String(registroAtual[H.MET.INTENS_MAN - 1] || '').trim()
  };
}

function _normalizarNivel(nivel) {
  const n = _textoNormalizado_(nivel);
  if (n.indexOf('inic') >= 0) return 'Iniciante';
  if (n.indexOf('avan') >= 0) return 'Avançado';
  return 'Intermediário';
}

function _normalizarPerfilManual(perfil) {
  const p = _textoNormalizado_(perfil);
  if (p.indexOf('retorno') >= 0 || p.indexOf('les') >= 0) return 'Retorno/lesão';
  return _normalizarNivel(perfil);
}

function _getFreqAtleta(athId) {
  const ws = SpreadsheetApp.getActive().getSheetByName(H.SHEETS.CADASTRO);
  if (!ws) return '';
  const r = ws.getDataRange().getValues().slice(2)
    .find(row => String(row[H.CAD.ID - 1]).trim().toUpperCase() === String(athId).trim().toUpperCase());
  return r ? String(r[H.CAD.FREQ - 1] || '') : '';
}

function _getNivelAtleta(athId) {
  const ws = SpreadsheetApp.getActive().getSheetByName(H.SHEETS.CADASTRO);
  if (!ws) return '';
  const r = ws.getDataRange().getValues().slice(2)
    .find(row => String(row[H.CAD.ID - 1]).trim().toUpperCase() === String(athId).trim().toUpperCase());
  return r ? String(r[H.CAD.NIVEL - 1] || '') : '';
}

function atualizarPainel() {
  const etapas = [];
  const executar = (nome, fn) => {
    try { fn(); etapas.push('✅ ' + nome); }
    catch (e) { etapas.push('⚠️ ' + nome + ': ' + e.message); _log('SISTEMA', 'AVISO', 'atualizarPainel', nome + ': ' + e.message, ''); }
  };
  executar('Fórmulas', () => repararFormulasOperacionais(false));
  executar('Métricas', () => calcularMetricasTodos(true));
  executar('Ranking', atualizarRankingSheet);
  executar('Ranking completo', atualizarRankingExpandido);
  executar('Análise', atualizarAnaliseSheet);
  executar('Status Strava', atualizarStravaStatusSheet);
  executar('Painel', _atualizarPainelInterno);
  executar('Organização das abas', organizarAbasOperacionais);
  try { SpreadsheetApp.getUi().alert('📊 Atualização concluída', etapas.join('\n'), SpreadsheetApp.getUi().ButtonSet.OK); } catch (_) {}
}

function _atualizarPainelInterno() {
  if (typeof _atualizarPainelInteligente_ === 'function') {
    _atualizarPainelInteligente_();
    return;
  }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ws = ss.getSheetByName(H.SHEETS.PAINEL);
  const wsCad = ss.getSheetByName(H.SHEETS.CADASTRO);
  const wsAtiv = ss.getSheetByName(H.SHEETS.ATIVIDADES);
  if (!ws || !wsCad || !wsAtiv) return;

  const agora = new Date();
  const hoje = new Date(agora.getFullYear(), agora.getMonth(), agora.getDate());
  const inicio7 = new Date(agora.getTime() - 7 * 86400000);
  const inicio28 = new Date(agora.getTime() - 28 * 86400000);
  const tz = Session.getScriptTimeZone() || 'America/Sao_Paulo';
  const atividades = wsAtiv.getDataRange().getValues().slice(2);
  ws.getRange('A3').setValue('Atualizado em: ' + Utilities.formatDate(agora, tz, 'dd/MM/yyyy HH:mm'));

  let treinos7 = 0;
  let segCorrida7 = 0;
  let kmCorrida7 = 0;
  let kmCorrida28 = 0;
  const ultimaPorId = {};

  atividades.forEach(r => {
    const id = String(r[H.ATIV.ATH_ID - 1] || '').trim().toUpperCase();
    const data = r[H.ATIV.DATA - 1];
    if (!_isAthIdValido_(id) || !(data instanceof Date) || isNaN(data.getTime())) return;
    if (!ultimaPorId[id] || data > ultimaPorId[id]) ultimaPorId[id] = data;
    const seg = _segundosAtividade_(r);
    if (data >= inicio7 && seg > 0) treinos7++;
    if (!_ehCorridaMetrica_(r)) return;
    let km = Number(r[H.ATIV.DIST_KM - 1]) || 0;
    if (!km) km = (Number(r[H.ATIV.DIST_M - 1]) || 0) / 1000;
    if (!(km > 0)) return;
    if (data >= inicio28) kmCorrida28 += km;
    if (data >= inicio7 && seg > 0) {
      kmCorrida7 += km;
      segCorrida7 += seg;
    }
  });

  const pace7 = kmCorrida7 > 0 ? Math.round(segCorrida7 / kmCorrida7) : 0;
  const paceFmt = pace7 > 0
    ? Math.floor(pace7 / 60) + ':' + String(pace7 % 60).padStart(2, '0')
    : '--';
  ws.getRange('G6').setValue(treinos7);
  ws.getRange('I6').setValue(paceFmt);
  ws.getRange('K6').setValue(Math.round(kmCorrida28 / 4 * 10) / 10);

  const atletas = wsCad.getDataRange().getValues().slice(2).map(r => ({
    id: String(r[H.CAD.ID - 1] || '').trim().toUpperCase(),
    nome: String(r[H.CAD.NOME - 1] || '').trim(),
    status: _textoNormalizado_(r[H.CAD.STATUS - 1]),
    usaStrava: _textoNormalizado_(r[H.CAD.STRAVA_OK - 1])
  })).filter(a => _isAthIdValido_(a.id) && a.status !== 'inativo' && a.usaStrava !== 'nao');

  const status = atletas.map(a => {
    const ultima = ultimaPorId[a.id] || null;
    const dias = ultima ? Math.max(0, Math.floor((hoje.getTime() - new Date(ultima.getFullYear(), ultima.getMonth(), ultima.getDate()).getTime()) / 86400000)) : null;
    return { id: a.id, nome: a.nome, ultima, dias };
  }).sort((a, b) => {
    if (a.ultima === null && b.ultima !== null) return -1;
    if (a.ultima !== null && b.ultima === null) return 1;
    return (b.dias || 0) - (a.dias || 0);
  }).slice(0, 10);

  const linhas = [];
  for (let i = 0; i < 10; i++) {
    const a = status[i];
    if (!a) {
      linhas.push(['', '', '', '']);
    } else if (!a.ultima) {
      linhas.push([a.nome, 'Nunca', '--', '⚪ Sem atividade registrada']);
    } else {
      const sinal = a.dias >= 14
        ? '🔴 ' + a.dias + 'd sem treinar'
        : a.dias >= 7 ? '🟡 Verificar (' + a.dias + 'd)' : '🟢 Ativo';
      linhas.push([a.nome, Utilities.formatDate(a.ultima, tz, 'dd/MM/yy'), a.dias, sinal]);
    }
  }
  // O layout original mescla H:L; selecionar o intervalo completo evita erro
  // de mesclagem parcial e preserva a coluna L usada apenas como margem.
  const tituloStatus = ws.getRange('H9:L9');
  tituloStatus.breakApart();
  tituloStatus.merge().setValue('📡 STATUS DE RECÊNCIA — ATLETAS QUE USAM STRAVA');
  ws.getRange(11, 8, linhas.length, 4).setValues(linhas);
}
