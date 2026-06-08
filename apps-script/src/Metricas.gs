/**
 * ═══════════════════════════════════════════════════════════════════════
 * HIPERATIVO V3 — 03_Metricas.gs
 * Cálculo de VO2máx, zonas Z1-Z5 e métricas consolidadas por atleta
 * ═══════════════════════════════════════════════════════════════════════
 */

// ── CALCULAR MÉTRICAS DE TODOS OS ATLETAS ────────────────────────────────────
function calcularMetricasTodos() {
  const wsCad = SpreadsheetApp.getActive().getSheetByName(H.SHEETS.CADASTRO);
  if (!wsCad) return;

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

  SpreadsheetApp.getUi().alert('✅ Métricas atualizadas', `${ok} atletas processados`, SpreadsheetApp.getUi().ButtonSet.OK);
}

// ── CALCULAR MÉTRICAS DE UM ATLETA ────────────────────────────────────────────
function _calcularMetricasAtleta(athId) {
  const wsAtiv = SpreadsheetApp.getActive().getSheetByName(H.SHEETS.ATIVIDADES);
  const wsMet  = SpreadsheetApp.getActive().getSheetByName(H.SHEETS.METRICAS);
  if (!wsAtiv || !wsMet) return;

  // Filtrar atividades de corrida dos últimos 28 dias
  const cutoff = new Date(Date.now() - 28 * 86400000);
  const rows   = wsAtiv.getDataRange().getValues().slice(2)
    .filter(r => String(r[H.ATIV.ATH_ID - 1]) === athId
               && r[H.ATIV.DATA - 1] instanceof Date
               && r[H.ATIV.DATA - 1] >= cutoff
               && String(r[H.ATIV.TIPO - 1]) === 'Corrida'
               && Number(r[H.ATIV.PACE_S - 1]) > 0);

  const registroAtual = _getRegistroMetrica(wsMet, athId);
  const manual = _resolverPerfilManual(athId, registroAtual);

  let paceMed = 0, paceRap = 9999, paceLento = 0;
  let fcMaxArr = [], fcMedArr = [], distArr = [];

  rows.forEach(r => {
    const ps = Number(r[H.ATIV.PACE_S - 1]);
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

  // VO2máx estimado (fórmula simplificada: velocidade × coeficiente de FC)
  let vo2 = manual.vo2;
  if (paceMed > 0 && fcMed && fcMax) {
    const vel = 60 / paceMed; // km/min
    const pct = (fcMed - 60) / (fcMax - 60);
    vo2 = Math.max(20, Math.min(80, Math.round((vel * 3.5 + 3.5) * pct * 0.9)));
  }

  // Salvar na aba MÉTRICAS
  const origem = nRows > 0 ? `Atividades 28d (${nRows})` : manual.origem;
  const confianca = nRows >= 3 ? 'Alta' : nRows > 0 ? 'Média' : 'Baixa';
  const metricasBase = [athId, _getNomeAtleta(athId), new Date(), vo2, paceMed, paceRap, paceLento, fcMax, fcMed, volSem];
  const metricasManuais = [manual.perfil, manual.volume, manual.intensidade, origem, confianca, manual.obs];
  const rowsMet = wsMet.getDataRange().getValues();

  for (let i = 2; i < rowsMet.length; i++) {
    if (String(rowsMet[i][H.MET.ATH_ID - 1]) === athId) {
      wsMet.getRange(i + 1, 1, 1, metricasBase.length).setValues([metricasBase]);
      wsMet.getRange(i + 1, H.MET.PERFIL_MAN, 1, metricasManuais.length).setValues([metricasManuais]);
      _log(athId, 'INFO', '_calcularMetricasAtleta', `VO2máx: ${vo2}, Pace médio: ${paceMed}s/km`, '');
      return;
    }
  }
  const linhaVazia = _primeiraLinhaVazia(wsMet, H.MET.ATH_ID);
  wsMet.getRange(linhaVazia, 1, 1, metricasBase.length).setValues([metricasBase]);
  wsMet.getRange(linhaVazia, H.MET.PERFIL_MAN, 1, metricasManuais.length).setValues([metricasManuais]);
  _log(athId, 'INFO', '_calcularMetricasAtleta', `Novo registro — VO2máx: ${vo2}`, '');
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
  // Força recalculação de fórmulas atualizando data na célula A3
  const ws = SpreadsheetApp.getActive().getSheetByName(H.SHEETS.PAINEL);
  if (ws) {
    ws.getRange('A3').setValue(
      'Atualizado em: ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm')
    );
  }
  calcularMetricasTodos();
  SpreadsheetApp.getUi().alert('✅ Painel atualizado com sucesso.', '', SpreadsheetApp.getUi().ButtonSet.OK);
}
