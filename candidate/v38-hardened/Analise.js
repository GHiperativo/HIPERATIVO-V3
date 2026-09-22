/**
 * ═══════════════════════════════════════════════════════════════════════
 * HIPERATIVO V3 — Analise.gs (v1.0 — 14/06/2026)
 * Motor de análise científica para prescrição de treinos
 *
 * Referências:
 *  - CTL/ATL/TSB: Banister (1991), Coggan/Friel (TrainingPeaks)
 *  - Polarized Training (80/20): Seiler (2010) — IJSPP
 *  - ACWR / Injury Risk: Gabbett (2016) — BJSM
 *  - Aerobic Decoupling: Friel (2009)
 *  - TRIMP (Training Impulse): Morton et al. (1990)
 * ═══════════════════════════════════════════════════════════════════════
 */

const A_CTL_DIAS = 42; // janela CTL (forma crônica)
const A_ATL_DIAS = 7;  // janela ATL (fadiga aguda)

// ── CALCULAR ANÁLISE COMPLETA DE UM ATLETA ───────────────────────────────────
function calcularAnaliseAtleta(athId) {
  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const wsAtiv = ss.getSheetByName(H.SHEETS.ATIVIDADES);
  const wsMet  = ss.getSheetByName(H.SHEETS.METRICAS);
  if (!wsAtiv || !wsMet) return null;

  const dados = wsAtiv.getDataRange().getValues().slice(2)
    .filter(r => String(r[H.ATIV.ATH_ID - 1]) === athId && r[H.ATIV.DATA - 1] instanceof Date)
    .sort((a, b) => a[H.ATIV.DATA - 1] - b[H.ATIV.DATA - 1]);

  if (dados.length === 0) return null;

  const metRow = _getRegistroMetrica(wsMet, athId);
  const fcMax  = Number(metRow[H.MET.FC_MAX - 1]) || 190;

  const ctl  = _calcularCTL(dados, fcMax);
  const atl  = _calcularATL(dados, fcMax);
  const tsb  = Math.round((ctl - atl) * 10) / 10;

  const analise = {
    athId,
    ctl,
    atl,
    tsb,
    distribuicao: _calcularDistribuicaoZonas(dados, fcMax),
    decoupling:   _calcularDecoupling(dados),
    acwr:         _calcularACWR(dados),
  };

  analise.prescricao = _gerarPrescricao(analise, metRow);
  return analise;
}

// ── CTL: Carga Crônica (forma) ────────────────────────────────────────────────
function _calcularCTL(dados, fcMax) {
  const hoje   = new Date();
  const dias90 = new Date(hoje.getTime() - 90 * 86400000);
  const recent = dados.filter(r => r[H.ATIV.DATA - 1] >= dias90);
  const k      = 1 - Math.exp(-1 / A_CTL_DIAS);

  const porDia = _agruparPorDia(recent, fcMax);
  let ctl = 0;
  for (let i = 89; i >= 0; i--) {
    const d  = Utilities.formatDate(new Date(hoje.getTime() - i * 86400000), 'America/Sao_Paulo', 'yyyy-MM-dd');
    const tl = porDia[d] || 0;
    ctl = ctl + k * (tl - ctl);
  }
  return Math.round(ctl * 10) / 10;
}

// ── ATL: Carga Aguda (fadiga) ─────────────────────────────────────────────────
function _calcularATL(dados, fcMax) {
  const hoje   = new Date();
  const dias21 = new Date(hoje.getTime() - 21 * 86400000);
  const recent = dados.filter(r => r[H.ATIV.DATA - 1] >= dias21);
  const k      = 1 - Math.exp(-1 / A_ATL_DIAS);

  const porDia = _agruparPorDia(recent, fcMax);
  let atl = 0;
  for (let i = 20; i >= 0; i--) {
    const d  = Utilities.formatDate(new Date(hoje.getTime() - i * 86400000), 'America/Sao_Paulo', 'yyyy-MM-dd');
    const tl = porDia[d] || 0;
    atl = atl + k * (tl - atl);
  }
  return Math.round(atl * 10) / 10;
}

// ── AGRUPAR ATIVIDADES POR DIA (TRIMP) ────────────────────────────────────────
function _agruparPorDia(dados, fcMax) {
  const porDia = {};
  dados.forEach(r => {
    const d  = Utilities.formatDate(r[H.ATIV.DATA - 1], 'America/Sao_Paulo', 'yyyy-MM-dd');
    const tl = _calcularTRIMP(r, fcMax);
    porDia[d] = (porDia[d] || 0) + tl;
  });
  return porDia;
}

// ── TRIMP: Training Impulse ───────────────────────────────────────────────────
// TRIMP = duração (min) × intensidade relativa
function _calcularTRIMP(row, fcMax) {
  const tempoFrac = Number(row[H.ATIV.MOV_S   - 1]) || 0;
  const tempoMin  = tempoFrac > 0 && tempoFrac < 1 ? tempoFrac * 1440 : tempoFrac / 60;
  const fcMed     = Number(row[H.ATIV.FC_MED  - 1]) || 0;
  const tipo      = String(row[H.ATIV.TIPO     - 1] || '');

  let intensidade;
  if (fcMed > 0 && fcMax > 0) {
    intensidade = fcMed / fcMax;
  } else {
    const base = { 'Corrida':0.75,'Trail Run':0.78,'Ciclismo':0.70,'Natação':0.78,'Caminhada':0.60,'Musculação':0.55 };
    intensidade = base[tipo] || 0.68;
  }
  return Math.round(tempoMin * intensidade * 10) / 10;
}

// ── DISTRIBUIÇÃO DE ZONAS (modelo Seiler 3 zonas) ─────────────────────────────
function _calcularDistribuicaoZonas(dados, fcMax) {
  const dias90 = new Date(Date.now() - 90 * 86400000);
  const recent = dados.filter(r => r[H.ATIV.DATA - 1] >= dias90 && Number(r[H.ATIV.FC_MED - 1]) > 0);
  if (recent.length === 0) return { z1z2:0, z3:0, z4z5:0, amostra:0, modelo:null };

  let z1z2 = 0, z3 = 0, z4z5 = 0;
  recent.forEach(r => {
    const fcRel   = Number(r[H.ATIV.FC_MED - 1]) / fcMax;
    const tempFrac = Number(r[H.ATIV.MOV_S - 1]) || 0;
    const mins    = tempFrac > 0 && tempFrac < 1 ? tempFrac * 1440 : tempFrac / 60;
    // Seiler Z limiar: <77% aeróbico | 77–91% limiar | >91% anaeróbico
    if      (fcRel < 0.77) z1z2 += mins;
    else if (fcRel < 0.91) z3   += mins;
    else                   z4z5 += mins;
  });

  const total = z1z2 + z3 + z4z5 || 1;
  const pZ1z2 = Math.round(z1z2 / total * 100);
  const pZ3   = Math.round(z3   / total * 100);
  const pZ4z5 = Math.round(z4z5 / total * 100);

  let modelo;
  if (pZ1z2 >= 70 && pZ4z5 >= 10) modelo = '✅ Polarizado (Seiler)';
  else if (pZ3 >= 40)              modelo = '⚠️ Excesso Z3 (zona cinza)';
  else                             modelo = '📊 Piramidal';

  return { z1z2:pZ1z2, z3:pZ3, z4z5:pZ4z5, amostra:recent.length, modelo };
}

// ── ACWR: Acute:Chronic Workload Ratio (Gabbett, 2016) ────────────────────────
// Zona segura: 0.8–1.3 | >1.5 = alto risco de lesão
function _calcularACWR(dados) {
  const hoje  = new Date();
  const dias7 = new Date(hoje.getTime() - 7  * 86400000);
  const dias28= new Date(hoje.getTime() - 28 * 86400000);

  const vol7  = dados.filter(r => r[H.ATIV.DATA - 1] >= dias7)
    .reduce((s, r) => s + (Number(r[H.ATIV.DIST_KM - 1]) || 0), 0);
  const vol28 = dados.filter(r => r[H.ATIV.DATA - 1] >= dias28)
    .reduce((s, r) => s + (Number(r[H.ATIV.DIST_KM - 1]) || 0), 0);

  const mediaSem = vol28 / 4;
  const acwr     = mediaSem > 0 ? Math.round(vol7 / mediaSem * 100) / 100 : null;

  let risco;
  if (!acwr)          risco = '— Sem dados';
  else if (acwr < 0.8) risco = '⚠️ Volume baixo (destreinamento)';
  else if (acwr <= 1.3)risco = '✅ Zona segura';
  else if (acwr <= 1.5)risco = '⚠️ Carga elevada — atenção';
  else                 risco = '🔴 Alto risco de lesão — reduzir';

  return { acwr, vol7: Math.round(vol7 * 10)/10, mediaSem: Math.round(mediaSem * 10)/10, risco };
}

// ── AEROBIC DECOUPLING (proxy via variabilidade de FC) ────────────────────────
function _calcularDecoupling(dados) {
  const longas = dados.filter(r => {
    const tipo  = String(r[H.ATIV.TIPO - 1] || '');
    const frac  = Number(r[H.ATIV.MOV_S  - 1]) || 0;
    const mins  = frac > 0 && frac < 1 ? frac * 1440 : frac / 60;
    return (tipo === 'Corrida' || tipo === 'Trail Run') && mins >= 45 && Number(r[H.ATIV.FC_MED - 1]) > 0;
  }).slice(-6);

  if (longas.length < 2) return { valor:null, status:'Poucos dados' };

  const fcs   = longas.map(r => Number(r[H.ATIV.FC_MED - 1]));
  const fcMed = fcs.reduce((a, b) => a + b, 0) / fcs.length;
  const desvio= Math.sqrt(fcs.reduce((s, f) => s + (f - fcMed) ** 2, 0) / fcs.length);
  const cv    = Math.round(desvio / fcMed * 100 * 10) / 10;

  let status;
  if (cv < 4)       status = '✅ Excelente (base aeróbica sólida)';
  else if (cv < 8)  status = '📊 Bom';
  else if (cv < 15) status = '⚠️ Moderado — trabalhe mais Z1-Z2';
  else              status = '🔴 Alta variação — priorizar base';

  return { valor: cv, status };
}

// ── GERAR PRESCRIÇÃO BASEADA NA ANÁLISE ──────────────────────────────────────
function _gerarPrescricao(analise, metRow) {
  const { ctl, atl, tsb, distribuicao, acwr } = analise;
  const paceMed = Number(metRow[H.MET.PACE_MED - 1]) || 0;
  const items   = [];

  // Estado geral (TSB)
  if      (tsb < -20) items.push({ tipo:'🔴 Recuperação',   msg:'TSB ' + tsb + ': corpo sobrecarregado. Priorize Z1, sono e alimentação. Sem treinos intensos.' });
  else if (tsb < -10) items.push({ tipo:'🟡 Adaptação',      msg:'TSB ' + tsb + ': em adaptação. Reduza volume 20% e mantenha 1 sessão de qualidade.' });
  else if (tsb > 15)  items.push({ tipo:'✅ Forma ideal',    msg:'TSB +' + tsb + ': ótima janela para treino de qualidade ou competição.' });
  else                items.push({ tipo:'📊 Equilíbrio',     msg:'TSB ' + tsb + ': carga balanceada. Continue o planejamento.' });

  // Zonas (Seiler)
  if (distribuicao.z1z2 < 60) {
    items.push({ tipo:'⚠️ Polarização baixa', msg:'Apenas ' + distribuicao.z1z2 + '% em Z1-Z2 (ideal ≥80%). Adicione corridas fáceis ao ritmo de conversação.' });
  }
  if (distribuicao.z3 > 35) {
    items.push({ tipo:'📊 Excesso Z3',        msg:distribuicao.z3 + '% em zona de limiar. "Zona cinza": esforço que fatiga sem adaptar. Polarize: mais fácil ou mais intenso.' });
  }

  // ACWR
  if (acwr.acwr && acwr.acwr > 1.3) {
    items.push({ tipo: acwr.risco, msg: 'ACWR ' + acwr.acwr + ' — semana (' + acwr.vol7 + 'km) vs média 4sem (' + acwr.mediaSem + 'km). Reduza volume ou adie treino longo.' });
  }

  // Zonas de pace
  if (paceMed > 0) {
    const z = [
      'Z1 (recuperação): ' + _fmtPace(Math.round(paceMed * 1.30)) + '–' + _fmtPace(Math.round(paceMed * 1.18)) + '/km',
      'Z2 (aeróbico):    ' + _fmtPace(Math.round(paceMed * 1.18)) + '–' + _fmtPace(Math.round(paceMed * 1.05)) + '/km',
      'Z3 (limiar):      ' + _fmtPace(Math.round(paceMed * 0.98)) + '/km',
      'Z4 (VO₂):         ' + _fmtPace(Math.round(paceMed * 0.93)) + '/km',
      'Z5 (máximo):      abaixo de ' + _fmtPace(Math.round(paceMed * 0.88)) + '/km',
    ].join(' | ');
    items.push({ tipo:'🏃 Zonas de Treino', msg: z });
  }

  return items;
}

// ── ATUALIZAR ABA 🔬 ANÁLISE ─────────────────────────────────────────────────
function atualizarAnaliseSheet() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const wsCad = ss.getSheetByName(H.SHEETS.CADASTRO);
  if (!wsCad) return;

  const atletas = wsCad.getDataRange().getValues().slice(2)
    .filter(r => r[H.CAD.ID - 1] && String(r[H.CAD.STATUS - 1]).toLowerCase() !== 'inativo')
    .map(r => String(r[H.CAD.ID - 1]));

  let shAnal = ss.getSheetByName('🔬 ANÁLISE');
  if (!shAnal) shAnal = ss.insertSheet('🔬 ANÁLISE');
  shAnal.clearContents();

  const corAzul = '#001F3F';
  shAnal.getRange(1, 1, 1, 10).merge()
    .setValue('🔬 ANÁLISE CIENTÍFICA — HIPERATIVO V3')
    .setFontWeight('bold').setFontColor('#FFFFFF').setBackground(corAzul)
    .setFontSize(13).setHorizontalAlignment('center');
  shAnal.setRowHeight(1, 36);

  const ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
  shAnal.getRange(2, 1, 1, 10).merge()
    .setValue('Atualizado: ' + ts + '  |  CTL/ATL/TSB (Banister) · ACWR (Gabbett) · Zonas (Seiler)')
    .setFontStyle('italic').setFontSize(9).setFontColor('#666').setHorizontalAlignment('center').setBackground('#F9F9F9');

  const cab = ['Atleta', 'CTL', 'ATL', 'TSB', 'ACWR', 'Z1-Z2%', 'Z3%', 'Z4-Z5%', 'Decoupling', 'Diagnóstico'];
  shAnal.getRange(3, 1, 1, 10).setValues([cab])
    .setFontWeight('bold').setFontColor('#FFFFFF').setBackground('#003D7A').setHorizontalAlignment('center');
  shAnal.setRowHeight(3, 24);

  const linhas = [];
  atletas.forEach(athId => {
    try {
      const a = calcularAnaliseAtleta(athId);
      if (!a) { linhas.push([_getNomeAtleta(athId), '-', '-', '-', '-', '-', '-', '-', '-', '⚠️ Sem dados']); return; }
      const diag   = a.prescricao && a.prescricao.length ? a.prescricao[0].tipo : '—';
      const acwrStr = a.acwr.acwr ? String(a.acwr.acwr) : 'N/A';
      const decStr  = a.decoupling.valor != null ? a.decoupling.valor + '%' : 'N/A';
      const tsbStr  = a.tsb > 0 ? '+' + a.tsb : String(a.tsb);
      linhas.push([_getNomeAtleta(athId), a.ctl, a.atl, tsbStr, acwrStr, a.distribuicao.z1z2 + '%', a.distribuicao.z3 + '%', a.distribuicao.z4z5 + '%', decStr, diag]);
    } catch(e) {
      linhas.push([_getNomeAtleta(athId), '-', '-', '-', '-', '-', '-', '-', '-', '❌ ' + e.message]);
    }
  });

  if (linhas.length > 0) {
    shAnal.getRange(4, 1, linhas.length, 10).setValues(linhas).setFontSize(10).setHorizontalAlignment('center');
    for (let i = 0; i < linhas.length; i++) {
      const row = i + 4;
      const tsb = Number(String(linhas[i][3]).replace('+', '')) || 0;
      const bg  = i % 2 === 0 ? '#FFFFFF' : '#F0F7FF';
      shAnal.getRange(row, 1, 1, 10).setBackground(bg);
      const tsbBg = tsb < -20 ? '#FDECEA' : tsb < -10 ? '#FFF8E1' : tsb > 10 ? '#E8F5E9' : bg;
      shAnal.getRange(row, 4).setBackground(tsbBg).setFontWeight('bold');
      shAnal.setRowHeight(row, 22);
    }
  } else {
    shAnal.getRange(4, 1).setValue('Nenhum atleta com dados suficientes para análise.');
  }

  [20, 9, 9, 9, 10, 9, 8, 9, 12, 32].forEach((w, i) => shAnal.setColumnWidth(i + 1, w * 7));
  SpreadsheetApp.flush();
  _log('SISTEMA', 'INFO', 'atualizarAnaliseSheet', atletas.length + ' atletas analisados', '');
}

// ── HELPERS ───────────────────────────────────────────────────────────────────
function _fmtPace(seg) {
  if (!seg || seg <= 0) return '--';
  const m = Math.floor(seg / 60);
  const s = seg % 60;
  return m + ':' + String(s).padStart(2, '0');
}
