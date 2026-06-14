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

// ── UTILITÁRIOS ───────────────────────────────────────────────────────────────

// Retorna a linha após o último dado (simples e sem dependência de colIdx)
function primeiraLinhaVazia(sheet) {
  return sheet.getLastRow() + 1;
}

// Limpa validação antes de gravar para evitar erro de dropdown
function gravarNivelAptidao(sheet, celula, valor) {
  const range = sheet.getRange(celula);
  range.clearDataValidations();
  range.setValue(valor);
}

// ── VO2MAX — FÓRMULA JACK DANIELS ─────────────────────────────────────────────
// paceSegundosPorKm: pace em segundos (ex: 5:39/km = 339 s)
// Usa o pace de esforço alto (melhores corridas), não o pace médio de treino
function estimarVO2max(paceSegundosPorKm) {
  const velocidadeMetrosPorMin = 1000 / (paceSegundosPorKm / 60);
  const v = velocidadeMetrosPorMin;
  const vo2max = -4.60 + 0.182258 * v + 0.000193 * v * v;
  return Math.round(Math.max(vo2max, 15));
}

// ── CALCULAR MÉTRICAS DE UM ATLETA ────────────────────────────────────────────
function _calcularMetricasAtleta(athId) {
  const wsAtiv = SpreadsheetApp.getActive().getSheetByName(H.SHEETS.ATIVIDADES);
  const wsMet  = SpreadsheetApp.getActive().getSheetByName(H.SHEETS.METRICAS);
  if (!wsAtiv || !wsMet) return;

  const todasAtividades = wsAtiv.getDataRange().getValues().slice(2)
    .filter(r => String(r[H.ATIV.ATH_ID - 1]) === athId
               && r[H.ATIV.DATA - 1] instanceof Date);

  // ── VO2max: mediana dos 3 melhores paces em corridas >= 3 km ──────────────
  const runsLongas = todasAtividades.filter(r => {
    const tipo   = String(r[H.ATIV.TIPO - 1]);
    const distKm = Number(r[H.ATIV.DIST_KM - 1]);
    const paceS  = Number(r[H.ATIV.PACE_S - 1]);
    return (tipo === 'Corrida' || tipo === 'Corrida (Trail)')
        && distKm >= 3
        && paceS > 60 && paceS < 1200; // sanidade: 1:00/km a 20:00/km
  });

  let vo2;
  if (runsLongas.length > 0) {
    runsLongas.sort((a, b) => Number(a[H.ATIV.PACE_S - 1]) - Number(b[H.ATIV.PACE_S - 1]));
    const top3 = runsLongas.slice(0, Math.min(3, runsLongas.length));
    const medianIdx = Math.floor((top3.length - 1) / 2);
    const paceMediano = Number(top3[medianIdx][H.ATIV.PACE_S - 1]);
    vo2 = estimarVO2max(paceMediano);
  } else {
    const nivel = _getNivelAtleta(athId);
    const fallback = { iniciante: 32, intermediário: 42, avançado: 52 };
    vo2 = fallback[nivel.toLowerCase()] || 42;
  }

  // ── Métricas de treino dos últimos 28 dias ─────────────────────────────────
  const cutoff = new Date(Date.now() - 28 * 86400000);
  const rows28 = todasAtividades.filter(r =>
    r[H.ATIV.DATA - 1] >= cutoff
    && String(r[H.ATIV.TIPO - 1]) === 'Corrida'
    && Number(r[H.ATIV.PACE_S - 1]) > 0
  );

  let paceMed = 0, paceRap = 9999, paceLento = 0;
  let fcMaxArr = [], fcMedArr = [], distArr = [];

  rows28.forEach(r => {
    const ps = Number(r[H.ATIV.PACE_S - 1]);
    if (ps > 0) {
      paceMed  += ps;
      paceRap   = Math.min(paceRap, ps);
      paceLento = Math.max(paceLento, ps);
    }
    if (Number(r[H.ATIV.FC_MAX - 1]) > 0) fcMaxArr.push(Number(r[H.ATIV.FC_MAX - 1]));
    if (Number(r[H.ATIV.FC_MED - 1]) > 0) fcMedArr.push(Number(r[H.ATIV.FC_MED - 1]));
    if (Number(r[H.ATIV.DIST_KM - 1]) > 0) distArr.push(Number(r[H.ATIV.DIST_KM - 1]));
  });

  const nRows = rows28.length;
  if (nRows > 0) {
    paceMed = Math.round(paceMed / nRows);
  } else {
    paceMed = 0; paceRap = 0; paceLento = 0;
  }

  const fcMax  = fcMaxArr.length ? Math.round(Math.max(...fcMaxArr)) : '';
  const fcMed  = fcMedArr.length
    ? Math.round(fcMedArr.reduce((a, b) => a + b, 0) / fcMedArr.length)
    : '';
  const volSem = distArr.length
    ? Math.round(distArr.reduce((a, b) => a + b, 0) / (28 / 7) * 10) / 10
    : '';

  // ── Salvar na aba MÉTRICAS ──────────────────────────────────────────────────
  const linha = [
    athId, _getNomeAtleta(athId), new Date(), vo2,
    nRows > 0 ? paceMed  : '',
    nRows > 0 ? paceRap  : '',
    nRows > 0 ? paceLento: '',
    fcMax, fcMed, volSem,
  ];
  const rowsMet = wsMet.getDataRange().getValues();

  for (let i = 2; i < rowsMet.length; i++) {
    if (String(rowsMet[i][H.MET.ATH_ID - 1]) === athId) {
      wsMet.getRange(i + 1, 1, 1, linha.length).setValues([linha]);
      _log(athId, 'INFO', '_calcularMetricasAtleta', 'VO2máx: ' + vo2 + ', Pace médio: ' + paceMed + 's/km', '');
      return;
    }
  }

  const linhaVazia = primeiraLinhaVazia(wsMet);
  wsMet.getRange(linhaVazia, 1, 1, linha.length).setValues([linha]);
  _log(athId, 'INFO', '_calcularMetricasAtleta', 'Novo registro — VO2máx: ' + vo2, '');
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
  // Sincroniza status dos tokens (sem API externa)
  try { sincronizarStatusStrava(); } catch(e) {}

  // Recalcula métricas de todos
  calcularMetricasTodos();

  // Atualiza timestamp no painel
  const ws = SpreadsheetApp.getActive().getSheetByName(H.SHEETS.PAINEL);
  if (ws) {
    ws.getRange('A3').setValue(
      'Atualizado em: ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm')
    );
  }
  SpreadsheetApp.getUi().alert('✅ Painel atualizado.', '', SpreadsheetApp.getUi().ButtonSet.OK);
}
