/**
 * ═══════════════════════════════════════════════════════════════════════
 * HIPERATIVO V3 — 03_Metricas.gs
 * Cálculo de VO2máx, zonas Z1-Z5 e métricas consolidadas por atleta
 * ═══════════════════════════════════════════════════════════════════════
 */


// ── RECALCULAR MÉTRICAS APÓS NOVA ATIVIDADE (chamado por _gravarAtividades) ──
/**
 * Chamada automaticamente sempre que uma nova atividade é gravada.
 * Garante que MÉTRICAS refletem o estado mais recente sem precisar de trigger manual.
 */
function recalcularMetricasAposAtividade(athId) {
  if (!athId) return;
  try {
    _calcularMetricasAtleta(athId);
    _log(athId, 'INFO', 'recalcularMetricasAposAtividade', 'Métricas atualizadas após nova atividade.', '');
  } catch(e) {
    _log(athId, 'AVISO', 'recalcularMetricasAposAtividade', 'Erro ao recalcular: ' + e.message, '');
  }
}

// ── CALCULAR MÉTRICAS DE TODOS OS ATLETAS ────────────────────────────────────
function calcularMetricasTodos() {
  const ss = SpreadsheetApp.getActive();
  const wsCad = ss.getSheetByName(H.SHEETS.CADASTRO);
  const wsMet = ss.getSheetByName(H.SHEETS.METRICAS);
  if (!wsCad || !wsMet) return;

  _garantirEstruturaMetricas_(wsMet);

  const atletas = wsCad.getDataRange().getValues().slice(2)
    .filter(r => r[H.CAD.ID - 1] && r[H.CAD.STATUS - 1] !== 'Inativo');

  let ok = 0;
  atletas.forEach(r => {
    try {
      _calcularMetricasAtleta(String(r[H.CAD.ID - 1]));
      ok++;
    } catch (e) {
      _log(String(r[H.CAD.ID - 1]), 'ERRO', 'calcularMetricasTodos', e.message, '');
    }
  });

  try { SpreadsheetApp.getUi().alert('✅ Métricas atualizadas', `${ok} atletas processados`, SpreadsheetApp.getUi().ButtonSet.OK); } catch(_) {}
  return ok;
}

// ── CALCULAR MÉTRICAS DE UM ATLETA ────────────────────────────────────────────
function _calcularMetricasAtleta(athId) {
  const wsAtiv = SpreadsheetApp.getActive().getSheetByName(H.SHEETS.ATIVIDADES);
  const wsMet  = SpreadsheetApp.getActive().getSheetByName(H.SHEETS.METRICAS);
  if (!wsAtiv || !wsMet) return;

  // Filtrar atividades de corrida dos últimos 28 dias
  const cutoff = new Date(Date.now() - 28 * 86400000);
  const rows   = wsAtiv.getDataRange().getValues().slice(2)
    .filter(r => {
      const psNovo = Number(r[H.ATIV.PACE_S - 1]);
      const psAntigo = Number(r[H.ATIV.PACE_FMT - 1]);
      const paceValido = (psNovo > 10 && psNovo < 3600) || (psAntigo > 10 && psAntigo < 3600);
      return String(r[H.ATIV.ATH_ID - 1]) === athId
        && r[H.ATIV.DATA - 1] instanceof Date
        && r[H.ATIV.DATA - 1] >= cutoff
        && String(r[H.ATIV.TIPO - 1]) === 'Corrida'
        && paceValido;
    });

  const registroAtual = _getRegistroMetrica(wsMet, athId);
  const manual = _resolverPerfilManual(athId, registroAtual);

  let paceMed = 0, paceRap = 9999, paceLento = 0;
  let fcMaxArr = [], fcMedArr = [], distArr = [];

  rows.forEach(r => {
    // Compatibilidade: dados antigos tinham string em col 15 e número em col 16.
    // Dados novos têm número em col 15 (PACE_S) e string em col 16 (PACE_FMT).
    const psNovo   = Number(r[H.ATIV.PACE_S   - 1]);  // col 15 — número nos novos dados
    const psAntigo = Number(r[H.ATIV.PACE_FMT - 1]);  // col 16 — número nos dados antigos
    const ps = (psNovo > 10 && psNovo < 3600) ? psNovo : (psAntigo > 10 && psAntigo < 3600 ? psAntigo : 0);
    if (ps > 0) { paceMed += ps; paceRap = Math.min(paceRap, ps); paceLento = Math.max(paceLento, ps); }
    if (Number(r[H.ATIV.FC_MAX - 1]) > 0) fcMaxArr.push(Number(r[H.ATIV.FC_MAX - 1]));
    if (Number(r[H.ATIV.FC_MED - 1]) > 0) fcMedArr.push(Number(r[H.ATIV.FC_MED - 1]));
    if (Number(r[H.ATIV.DIST_KM - 1]) > 0) distArr.push(Number(r[H.ATIV.DIST_KM - 1]));
  });

  const nRows = rows.length;
  if (nRows > 0) paceMed = Math.round(paceMed / nRows); else { paceMed = 0; paceRap = 0; paceLento = 0; }

  if (nRows === 0) {
    paceMed   = manual.paceMed;
    paceRap   = manual.paceRap;
    paceLento = manual.paceLento;
  }

  const fcMax = fcMaxArr.length ? Math.round(Math.max(...fcMaxArr)) : manual.fcMax;
  const fcMed = fcMedArr.length ? Math.round(fcMedArr.reduce((a, b) => a + b, 0) / fcMedArr.length) : Math.round(fcMax * 0.83);
  const volSem = distArr.length ? Math.round(distArr.reduce((a, b) => a + b, 0) / (28 / 7) * 10) / 10 : manual.volSem;

  // VO2máx estimado pela equação metabólica de corrida (ACSM), usando o
  // melhor pace médio recente. Evita o erro anterior de unidade km/min, que
  // achatava todos os resultados no piso de 20.
  let vo2 = manual.vo2;
  if (paceRap > 0) {
    const velocidadeMMin = 60000 / paceRap;
    vo2 = Math.max(20, Math.min(80, Math.round(3.5 + 0.2 * velocidadeMMin)));
  }

  // Salvar na aba MÉTRICAS
  // ── Determinar status de dados ────────────────────────────────────────────
  let origem, confianca, obsExibir;
  if (nRows === 0) {
    origem    = '⚠️ Sem atividades recentes (28d)';
    confianca = '⚠️ Sem dados suficientes';
    obsExibir = 'Aguardando treinos para gerar métricas reais. Estimativa por perfil manual.';
  } else if (nRows < 3) {
    origem    = `⚡ Poucos dados (${nRows} treino${nRows > 1 ? 's' : ''} em 28d)`;
    confianca = '⚠️ Dados insuficientes';
    obsExibir = `Apenas ${nRows} treino(s) nos últimos 28 dias. Mínimo recomendado: 3.`;
  } else if (nRows < 6) {
    origem    = `📊 Dados parciais (${nRows} treinos em 28d)`;
    confianca = 'Média';
    obsExibir = 'Métricas baseadas em dados reais, mas o volume ainda é baixo.';
  } else {
    origem    = `✅ Dados reais (${nRows} treinos em 28d)`;
    confianca = 'Alta';
    obsExibir = 'Métricas calculadas com atividades reais dos últimos 28 dias.';
  }

  // VO2max: armazenar como inteiro (sem decimal — precisão real não justifica decimal)
  const vo2Int = vo2 ? Math.round(vo2) : 0;
  const metricasBase = [athId, _getNomeAtleta(athId), new Date(), vo2Int, paceMed, paceRap, paceLento, fcMax, fcMed, volSem];
  const zonas = _calcularZonasPace_(paceMed, paceRap);
  const metricasManuais = [manual.perfil, manual.volume, manual.intensidade, origem, confianca, obsExibir];
  const rowsMet = wsMet.getDataRange().getValues();

  for (let i = 2; i < rowsMet.length; i++) {
    if (String(rowsMet[i][H.MET.ATH_ID - 1]) === athId) {
      wsMet.getRange(i + 1, 1, 1, metricasBase.length).setValues([metricasBase]);
      wsMet.getRange(i + 1, H.MET.Z1_LENTO, 1, zonas.length).setValues([zonas]);
      wsMet.getRange(i + 1, H.MET.PERFIL_MAN, 1, metricasManuais.length).setValues([metricasManuais]);
      _log(athId, 'INFO', '_calcularMetricasAtleta', `VO2máx: ${vo2}, Pace médio: ${paceMed}s/km`, '');
      return;
    }
  }
  const linhaVazia = _primeiraLinhaVazia(wsMet, H.MET.ATH_ID);
  wsMet.getRange(linhaVazia, 1, 1, metricasBase.length).setValues([metricasBase]);
  wsMet.getRange(linhaVazia, H.MET.Z1_LENTO, 1, zonas.length).setValues([zonas]);
  wsMet.getRange(linhaVazia, H.MET.PERFIL_MAN, 1, metricasManuais.length).setValues([metricasManuais]);
  _log(athId, 'INFO', '_calcularMetricasAtleta', `Novo registro — VO2máx: ${vo2}`, '');
}

function _garantirEstruturaMetricas_(ws) {
  const headers = [
    'ID Atleta','Nome Atleta','Atualizado em','VO2máx Est.',
    'Pace Médio (s/km)','Pace Rápido (s/km)','Pace Lento (s/km)',
    'FC Máxima','FC Média','Vol./Semana (km)',
    'Z1 Lento','Z1 Rápido','Z2 Lento','Z2 Rápido',
    'Z3 Lento','Z3 Rápido','Z4 Lento','Z5 Mín',
    'Perfil Manual','Volume Manual','Intensidade Manual',
    'Origem dos Dados','Confiança','Observações'
  ];
  ws.getRange(1, 1, 1, ws.getMaxColumns()).breakApart();
  ws.getRange(1, 1, 1, headers.length).merge().setValue('📈 MÉTRICAS E ZONAS DE TREINO — HIPERATIVO V3');
  ws.getRange(2, 1, 1, headers.length).setValues([headers])
    .setBackground('#1D9E75').setFontColor('#FFFFFF').setFontWeight('bold')
    .setHorizontalAlignment('center').setWrap(true);
  ws.setFrozenRows(2);
}

function _calcularZonasPace_(paceMed, paceRap) {
  if (!paceMed || paceMed <= 0) return ['', '', '', '', '', '', '', ''];
  const fmt = s => {
    const total = Math.max(1, Math.round(s));
    return Math.floor(total / 60) + ':' + String(total % 60).padStart(2, '0');
  };
  return [
    fmt(paceMed * 1.20), fmt(paceMed * 1.08),
    fmt(paceMed * 1.04), fmt(paceMed * 0.96),
    fmt(paceMed * 0.94), fmt(paceMed * 0.87),
    fmt(paceMed * 0.84), fmt(paceRap || paceMed * 0.80)
  ];
}

function _getRegistroMetrica(wsMet, athId) {
  const rows = wsMet.getDataRange().getValues();
  for (let i = 2; i < rows.length; i++) {
    if (String(rows[i][H.MET.ATH_ID - 1]) === athId) return rows[i];
  }
  return [];
}

function _resolverPerfilManual(athId, registroAtual) {
  const nivel = _getNivelAtleta(athId);
  const perfil = _normalizarPerfilManual(registroAtual[H.MET.PERFIL_MAN - 1] || nivel);
  const volume = String(registroAtual[H.MET.VOLUME_MAN - 1] || _getFreqAtleta(athId) || 'Moderado').trim();
  const intensidade = String(registroAtual[H.MET.INTENS_MAN - 1] || 'Moderado').trim();

  const perfilBase = {
    'Iniciante':      { vo2: 32, fcMax: 175, pace: 480 },
    'Intermediário':  { vo2: 42, fcMax: 182, pace: 390 },
    'Avançado':       { vo2: 52, fcMax: 188, pace: 330 },
    'Retorno/lesão':  { vo2: 30, fcMax: 172, pace: 520 },
  };
  const volumeBase = {
    'Baixo': 8,
    'Moderado': 18,
    'Alto': 32,
    'Muito alto': 45,
  };
  const intensidadeAjuste = {
    'Leve': 1.08,
    'Moderado': 1,
    'Forte': 0.94,
    'Competitivo': 0.88,
  };

  const base = perfilBase[perfil] || perfilBase[_normalizarNivel(nivel)] || perfilBase['Intermediário'];
  const ajuste = intensidadeAjuste[intensidade] || 1;
  const paceMed = Math.round(base.pace * ajuste);

  return {
    perfil,
    volume,
    intensidade,
    vo2: base.vo2,
    fcMax: base.fcMax,
    paceMed,
    paceRap: Math.round(paceMed * 0.90),
    paceLento: Math.round(paceMed * 1.18),
    volSem: volumeBase[volume] || 18,
    origem: 'Perfil manual sem dados recentes',
    obs: 'Estimativa por múltipla escolha; revisar após 3 treinos válidos.'
  };
}

function _normalizarNivel(nivel) {
  const n = String(nivel || '').toLowerCase();
  if (n.indexOf('inic') >= 0) return 'Iniciante';
  if (n.indexOf('avan') >= 0 || n.indexOf('avanç') >= 0) return 'Avançado';
  return 'Intermediário';
}

function _normalizarPerfilManual(perfil) {
  const p = String(perfil || '').toLowerCase();
  if (p.indexOf('retorno') >= 0 || p.indexOf('les') >= 0) return 'Retorno/lesão';
  return _normalizarNivel(perfil);
}

function _getFreqAtleta(athId) {
  const ws = SpreadsheetApp.getActive().getSheetByName(H.SHEETS.CADASTRO);
  if (!ws) return 'Moderado';
  const rows = ws.getDataRange().getValues().slice(2);
  const r = rows.find(row => String(row[H.CAD.ID - 1]) === athId);
  const freq = r ? String(r[H.CAD.FREQ - 1]) : '';
  if (/1|2|baixo/i.test(freq)) return 'Baixo';
  if (/5|6|alto/i.test(freq)) return 'Alto';
  return 'Moderado';
}

function _primeiraLinhaVazia(ws, col) {
  const lastRow = Math.max(ws.getLastRow(), 3);
  const values = ws.getRange(3, col, lastRow - 2, 1).getValues();
  for (let i = 0; i < values.length; i++) {
    if (!values[i][0]) return i + 3;
  }
  return lastRow + 1;
}

function _getNivelAtleta(athId) {
  const ws   = SpreadsheetApp.getActive().getSheetByName(H.SHEETS.CADASTRO);
  if (!ws) return 'intermediário';
  const rows = ws.getDataRange().getValues().slice(2);
  const r    = rows.find(row => String(row[H.CAD.ID - 1]) === athId);
  return r ? String(r[H.CAD.NIVEL - 1]).toLowerCase() : 'intermediário';
}

// ── ATUALIZAR PAINEL ───────────────────────────────────────────────────────────
function atualizarPainel() {
  _atualizarPainelInterno();
  try { SpreadsheetApp.getUi().alert('✅ Painel atualizado com sucesso.', '', SpreadsheetApp.getUi().ButtonSet.OK); } catch(_) {}
}

// Versão sem UI para uso em triggers automáticos
function _atualizarPainelInterno() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ws = ss.getSheetByName(H.SHEETS.PAINEL);
  if (ws) {
    ws.getRange('A3').setValue(
      'Atualizado em: ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm')
    );
  }
  calcularMetricasTodos();
}
