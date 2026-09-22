/**
 * ═══════════════════════════════════════════════════════════════════════
 * HIPERATIVO V3 — 04_SISRUN.gs
 * Exportação para SISRUN, envio WhatsApp e E-mail
 * ═══════════════════════════════════════════════════════════════════════
 */

// ── EXPORTAR PARA O SISRUN ────────────────────────────────────────────────────
function exportarSISRUN() {
  const ui = SpreadsheetApp.getUi();

  const rAtleta = ui.prompt('Exportar → SISRUN', 'ID do atleta (ex: ATH_001):', ui.ButtonSet.OK_CANCEL);
  if (rAtleta.getSelectedButton() !== ui.Button.OK) return;
  const athId = rAtleta.getResponseText().trim();

  const rSemana = ui.prompt('Exportar → SISRUN',
    'Data de início da semana (DD/MM/AAAA):\nEx: 28/04/2026', ui.ButtonSet.OK_CANCEL);
  if (rSemana.getSelectedButton() !== ui.Button.OK) return;
  const dataStr = rSemana.getResponseText().trim();

  try {
    const texto = _gerarTextoSISRUN(athId, dataStr);
    const html  = HtmlService.createHtmlOutput(
      `<html><head><meta charset="UTF-8">
      <style>body{font-family:Arial;padding:20px}
      pre{background:#0A1628;color:#A8D5FF;padding:16px;border-radius:8px;font-size:12px;
          line-height:1.7;white-space:pre-wrap;overflow:auto}
      button{padding:10px 20px;background:#00B4FF;color:#fff;border:none;
             border-radius:6px;font-size:14px;cursor:pointer;margin-bottom:12px}
      </style></head><body>
      <button onclick="navigator.clipboard.writeText(document.getElementById('txt').innerText)
        .then(()=>{this.innerText='✅ Copiado!';setTimeout(()=>this.innerText='📋 Copiar para área de transferência',2000)})">
        📋 Copiar para área de transferência</button>
      <pre id="txt">${_escapeHtml(texto)}</pre>
      </body></html>`
    ).setWidth(700).setHeight(560);

    ui.showModalDialog(html, `SISRUN — ${athId} — Semana ${dataStr}`);
    _log(athId, 'INFO', 'exportarSISRUN', `Exportado semana ${dataStr}`, '');

  } catch (e) {
    ui.alert('❌ Erro', e.message, ui.ButtonSet.OK);
    _log(athId, 'ERRO', 'exportarSISRUN', e.message, '');
  }
}

// ── GERAR TEXTO SISRUN ────────────────────────────────────────────────────────
function _gerarTextoSISRUN(athId, dataStr) {
  const ws    = SpreadsheetApp.getActive().getSheetByName(H.SHEETS.PLANO);
  const wsMet = SpreadsheetApp.getActive().getSheetByName(H.SHEETS.METRICAS);
  if (!ws) throw new Error('Aba PLANO SEMANAL não encontrada');

  // Converter data
  const partes  = dataStr.split('/');
  if (partes.length < 3) throw new Error('Data inválida. Use DD/MM/AAAA');
  const semanaInicio = new Date(partes[2], partes[1] - 1, partes[0]);
  const semanaFim    = new Date(semanaInicio.getTime() + 6 * 86400000);

  // Buscar plano da semana
  const rows = ws.getDataRange().getValues().slice(2);
  const plano = rows.find(r => {
    const s = r[H.PLANO.ATH_ID - 1];
    const d = r[H.PLANO.SEMANA - 1];
    if (String(s) !== athId) return false;
    if (!(d instanceof Date)) return false;
    return Math.abs(d.getTime() - semanaInicio.getTime()) < 86400000;
  });

  if (!plano) throw new Error(`Nenhum plano encontrado para ${athId} na semana de ${dataStr}. Salve a prescrição primeiro na aba 📅 PLANO SEMANAL.`);

  // Buscar métricas
  const rowsMet = wsMet ? wsMet.getDataRange().getValues().slice(2) : [];
  const met     = rowsMet.find(r => String(r[H.MET.ATH_ID - 1]) === athId);

  const nomeAtleta = plano[H.PLANO.NOME  - 1] || athId;
  const titulo     = plano[H.PLANO.TITULO - 1] || '';
  const diag       = plano[H.PLANO.DIAG  - 1] || '';
  const ciclo      = plano[H.PLANO.CICLO - 1] || '';
  const intencao   = plano[H.PLANO.INTENCAO - 1] || '';

  const d1Tipo  = plano[H.PLANO.D1_TIPO  - 1] || '';
  const d1Presc = plano[H.PLANO.D1_PRESC - 1] || '';
  const d2Tipo  = plano[H.PLANO.D2_TIPO  - 1] || '';
  const d2Presc = plano[H.PLANO.D2_PRESC - 1] || '';
  const d3Tipo  = plano[H.PLANO.D3_TIPO  - 1] || '';
  const d3Presc = plano[H.PLANO.D3_PRESC - 1] || '';

  // Zonas
  const fmtPace = (s) => s > 0 ? `${Math.floor(s/60)}:${String(s%60).padStart(2,'0')}` : '--';
  const paceMed  = met ? Number(met[H.MET.PACE_MED - 1]) : 0;
  const fcMax    = met ? Number(met[H.MET.FC_MAX   - 1]) : 0;
  const z = {
    z1: [Math.round(paceMed * 1.20), Math.round(paceMed * 1.08)],
    z2: [Math.round(paceMed * 1.04), Math.round(paceMed * 0.96)],
    z3: [Math.round(paceMed * 0.94), Math.round(paceMed * 0.87)],
    z4: [Math.round(paceMed * 0.84), Math.round(paceMed * 0.79)],
    z5: [Math.round(paceMed * 0.78), 0],
  };

  const fmtData = (d) => Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd/MM/yyyy');
  const linha   = '─'.repeat(46);
  const linha2  = '═'.repeat(46);

  const diasSemana = ['DOM','SEG','TER','QUA','QUI','SEX','SÁB'];
  const nomeDia = (offset) => {
    const d = new Date(semanaInicio.getTime() + offset * 86400000);
    return diasSemana[d.getDay()] + ' ' + fmtData(d);
  };

  let txt = '';
  txt += `ATLETA: ${nomeAtleta} | ID: ${athId}\n`;
  txt += `SEMANA: ${fmtData(semanaInicio)} – ${fmtData(semanaFim)}\n`;
  txt += ciclo ? `CICLO:  Semana ${ciclo}/12\n` : '';
  txt += titulo ? `TÍTULO: ${titulo}\n` : '';
  txt += `${linha2}\n\n`;

  if (diag) {
    txt += `DIAGNÓSTICO CCC:\n${diag}\n\n${linha}\n\n`;
  }

  if (d1Tipo && d1Presc) {
    txt += `${nomeDia(0)} — ${d1Tipo.toUpperCase()}\n`;
    txt += `${d1Presc.trim()}\n\n${linha}\n\n`;
  }

  if (d2Tipo && d2Presc) {
    txt += `${nomeDia(2)} — ${d2Tipo.toUpperCase()}\n`;
    txt += `${d2Presc.trim()}\n\n${linha}\n\n`;
  }

  if (d3Tipo && d3Presc) {
    txt += `${nomeDia(4)} — ${d3Tipo.toUpperCase()}\n`;
    txt += `${d3Presc.trim()}\n\n${linha}\n\n`;
  }

  if (met && paceMed > 0) {
    txt += `ZONAS INDIVIDUAIS (pace min:s por km):\n`;
    txt += `  Z1 Regenerativo: ${fmtPace(z.z1[0])} – ${fmtPace(z.z1[1])}  (60-70% FC)\n`;
    txt += `  Z2 Aeróbico base: ${fmtPace(z.z2[0])} – ${fmtPace(z.z2[1])}  (70-80% FC)\n`;
    txt += `  Z3 Limiar aeróbico: ${fmtPace(z.z3[0])} – ${fmtPace(z.z3[1])}  (80-87% FC)\n`;
    txt += `  Z4 Limiar anaeróbico: ${fmtPace(z.z4[0])} – ${fmtPace(z.z4[1])}  (87-93% FC)\n`;
    txt += `  Z5 VO2máx/Sprint: abaixo de ${fmtPace(z.z5[0])}  (93-100% FC)\n`;
    if (fcMax) txt += `  FC Máx referência: ${fcMax} bpm\n`;
    txt += `\n${linha}\n\n`;
  }

  if (intencao) {
    txt += `INTENÇÃO DA SEMANA:\n${intencao}\n\n${linha}\n\n`;
  }

  txt += `Gerado em: ${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm')} | HIPERATIVO V3\n`;

  return txt;
}

// ── ENVIAR POR WHATSAPP ───────────────────────────────────────────────────────
function enviarWhatsApp() {
  const ui = SpreadsheetApp.getUi();

  const rAtleta = ui.prompt('Enviar WhatsApp', 'ID do atleta:', ui.ButtonSet.OK_CANCEL);
  if (rAtleta.getSelectedButton() !== ui.Button.OK) return;
  const athId = rAtleta.getResponseText().trim();

  const rSemana = ui.prompt('Enviar WhatsApp', 'Semana (DD/MM/AAAA):', ui.ButtonSet.OK_CANCEL);
  if (rSemana.getSelectedButton() !== ui.Button.OK) return;

  try {
    const texto = _gerarTextoWpp(athId, rSemana.getResponseText().trim());
    const whats = _getWhatsAtleta(athId);

    const url = `https://wa.me/${whats.replace(/\D/g, '')}?text=${encodeURIComponent(texto)}`;

    const html = HtmlService.createHtmlOutput(
      `<html><head><meta charset="UTF-8"></head>
      <body style="font-family:Arial;padding:24px">
      <h3>📱 Enviar via WhatsApp</h3>
      <p><strong>Atleta:</strong> ${_getNomeAtleta(athId)}</p>
      <p><strong>WhatsApp:</strong> ${whats}</p>
      <p>Clique no botão abaixo para abrir o WhatsApp com o treino pré-preenchido:</p>
      <a href="${url}" target="_blank"
        style="display:inline-block;padding:12px 24px;background:#25D366;color:#fff;
               text-decoration:none;border-radius:8px;font-weight:bold;font-size:14px">
        📲 Abrir WhatsApp
      </a>
      <p style="font-size:12px;color:#888;margin-top:16px">
        Revise o texto antes de enviar. Nenhuma mensagem sai automaticamente.
      </p>
      </body></html>`
    ).setWidth(480).setHeight(300);

    ui.showModalDialog(html, 'Enviar WhatsApp');
    _log(athId, 'INFO', 'enviarWhatsApp', `Link WhatsApp gerado semana ${rSemana.getResponseText()}`, '');

  } catch (e) {
    ui.alert('❌ Erro', e.message, ui.ButtonSet.OK);
    _log(athId, 'ERRO', 'enviarWhatsApp', e.message, '');
  }
}

// ── ENVIAR POR E-MAIL ─────────────────────────────────────────────────────────
function enviarEmail() {
  const ui = SpreadsheetApp.getUi();

  const rAtleta = ui.prompt('Enviar E-mail', 'ID do atleta:', ui.ButtonSet.OK_CANCEL);
  if (rAtleta.getSelectedButton() !== ui.Button.OK) return;
  const athId = rAtleta.getResponseText().trim();

  const rSemana = ui.prompt('Enviar E-mail', 'Semana (DD/MM/AAAA):', ui.ButtonSet.OK_CANCEL);
  if (rSemana.getSelectedButton() !== ui.Button.OK) return;

  try {
    const wsCad = SpreadsheetApp.getActive().getSheetByName(H.SHEETS.CADASTRO);
    const rows  = wsCad.getDataRange().getValues().slice(2);
    const cad   = rows.find(r => String(r[H.CAD.ID - 1]) === athId);
    if (!cad) throw new Error(`Atleta ${athId} não encontrado`);

    const email  = String(cad[H.CAD.EMAIL - 1]);
    const nome   = String(cad[H.CAD.NOME  - 1]);
    const texto  = _gerarTextoWpp(athId, rSemana.getResponseText().trim());

    // Confirma antes de enviar
    const conf = ui.alert(
      'Confirmar envio?',
      `Para: ${nome} <${email}>\n\nAssunto: Treino da semana — Grupo Hiperativo\n\nClique OK para enviar.`,
      ui.ButtonSet.OK_CANCEL
    );
    if (conf !== ui.Button.OK) return;

    GmailApp.sendEmail(email,
      `⚡ Treino da Semana — Grupo Hiperativo`,
      texto,
      { name: 'Grupo Hiperativo' }
    );

    ui.alert('✅ E-mail enviado!', `Treino enviado para ${email}`, ui.ButtonSet.OK);
    _log(athId, 'INFO', 'enviarEmail', `E-mail enviado para ${email}`, '');

  } catch (e) {
    ui.alert('❌ Erro', e.message, ui.ButtonSet.OK);
    _log(athId, 'ERRO', 'enviarEmail', e.message, '');
  }
}

// ── HELPERS ───────────────────────────────────────────────────────────────────
function _gerarTextoWpp(athId, dataStr) {
  const ws    = SpreadsheetApp.getActive().getSheetByName(H.SHEETS.PLANO);
  if (!ws) throw new Error('Aba PLANO SEMANAL não encontrada');

  const partes = dataStr.split('/');
  if (partes.length < 3) throw new Error('Data inválida. Use DD/MM/AAAA');
  const semanaInicio = new Date(partes[2], partes[1] - 1, partes[0]);

  const rows  = ws.getDataRange().getValues().slice(2);
  const plano = rows.find(r => {
    const s = r[H.PLANO.ATH_ID - 1];
    const d = r[H.PLANO.SEMANA - 1];
    if (String(s) !== athId) return false;
    if (!(d instanceof Date)) return false;
    return Math.abs(d.getTime() - semanaInicio.getTime()) < 86400000;
  });

  if (!plano) throw new Error(`Plano não encontrado para ${athId} em ${dataStr}`);

  const titulo   = plano[H.PLANO.TITULO   - 1] || '';
  const diag     = plano[H.PLANO.DIAG     - 1] || '';
  const d1Tipo   = plano[H.PLANO.D1_TIPO  - 1] || '';
  const d1Presc  = plano[H.PLANO.D1_PRESC - 1] || '';
  const d2Tipo   = plano[H.PLANO.D2_TIPO  - 1] || '';
  const d2Presc  = plano[H.PLANO.D2_PRESC - 1] || '';
  const d3Tipo   = plano[H.PLANO.D3_TIPO  - 1] || '';
  const d3Presc  = plano[H.PLANO.D3_PRESC - 1] || '';
  const intencao = plano[H.PLANO.INTENCAO - 1] || '';

  const diasSemana = ['DOM','SEG','TER','QUA','QUI','SEX','SÁB'];
  const nomeDia = (offset) => {
    const d = new Date(semanaInicio.getTime() + offset * 86400000);
    return diasSemana[d.getDay()] + ' ' + Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd/MM');
  };

  let txt = '';
  if (titulo) txt += `🏷️ SEMANA: ${titulo}\n`;
  txt += '─'.repeat(38) + '\n';
  if (diag) txt += `📋 DIAGNÓSTICO CCC:\n${diag}\n${'─'.repeat(38)}\n`;

  if (d1Tipo && d1Presc) txt += `\n📅 ${nomeDia(0)} — ${d1Tipo.toUpperCase()}\n${d1Presc.trim()}\n`;
  if (d2Tipo && d2Presc) txt += `\n📅 ${nomeDia(2)} — ${d2Tipo.toUpperCase()}\n${d2Presc.trim()}\n`;
  if (d3Tipo && d3Presc) txt += `\n📅 ${nomeDia(4)} — ${d3Tipo.toUpperCase()}\n${d3Presc.trim()}\n`;

  txt += '\n' + '─'.repeat(38) + '\n';
  if (intencao) txt += `🧠 INTENÇÃO DA SEMANA:\n${intencao}\n`;
  txt += `\n⚡ Grupo Hiperativo — Cabeça, Coração e Corpo`;

  return txt;
}

function _getWhatsAtleta(athId) {
  const ws   = SpreadsheetApp.getActive().getSheetByName(H.SHEETS.CADASTRO);
  if (!ws) throw new Error('Aba CADASTRO não encontrada');
  const rows = ws.getDataRange().getValues().slice(2);
  const r    = rows.find(row => String(row[H.CAD.ID - 1]) === athId);
  if (!r) throw new Error(`Atleta ${athId} não encontrado`);
  const w = String(r[H.CAD.WHATS - 1] || '');
  if (!w) throw new Error(`Atleta ${athId} sem WhatsApp cadastrado`);
  return w;
}

function _escapeHtml(txt) {
  return txt.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}