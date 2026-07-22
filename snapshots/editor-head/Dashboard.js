/**
 * ═══════════════════════════════════════════════════════════════════════
 * HIPERATIVO V3 — Dashboard.gs
 * Atualização das abas RANKING, STRAVA STATUS e STRAVA ENVIOS
 * Chamado automaticamente pelo triggerImportacaoAutomatica (Strava.gs)
 * ═══════════════════════════════════════════════════════════════════════
 */

// ── Helper: localiza uma aba pelo nome (exato ou parcial) ─────────────────────
function _getAba(ss, nome) {
  let sh = ss.getSheetByName(nome);
  if (sh) return sh;
  const sheets = ss.getSheets();
  for (const s of sheets) {
    if (s.getName().toUpperCase().replace(/[^A-Z0-9 ]/g, '').indexOf(nome.toUpperCase().replace(/[^A-Z0-9 ]/g, '')) >= 0) {
      return s;
    }
  }
  return null;
}

// ══════════════════════════════════════════════════════════════════════════════
// RANKING — Top atletas por km (últimos 30 dias)
// ══════════════════════════════════════════════════════════════════════════════
function atualizarRankingSheet() {
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const shAtiv  = ss.getSheetByName(H.SHEETS.ATIVIDADES);
  const shRank  = _getAba(ss, 'RANKING');
  if (!shRank) return; // aba não existe, ok
  if (!shAtiv)  return;

  const corAzul   = '#001F3F';
  const corVerde  = '#1D9E75';
  const corClaro  = '#D6EEFF';
  const corBranco = '#FFFFFF';

  const cutoff = new Date(Date.now() - 30 * 86400000);
  const dados  = shAtiv.getDataRange().getValues();

  // Acumular por atleta: km, treinos, último treino
  const mapa = {};
  for (let i = 2; i < dados.length; i++) {
    const athId = String(dados[i][H.ATIV.ATH_ID   - 1] || '').trim();
    const nome  = String(dados[i][H.ATIV.NOME      - 1] || '').trim();
    const data  = dados[i][H.ATIV.DATA - 1];
    const tipo  = String(dados[i][H.ATIV.TIPO      - 1] || '').trim();
    const km    = Number(dados[i][H.ATIV.DIST_KM   - 1]) || 0;
    if (!_isAthIdValido_(athId) || !(data instanceof Date)) continue;

    if (!mapa[athId]) mapa[athId] = { nome, km30: 0, treinos30: 0, kmTotal: 0, treinosTotal: 0, ultimoTreino: null };

    mapa[athId].kmTotal += km;
    mapa[athId].treinosTotal++;
    if (!mapa[athId].ultimoTreino || data > mapa[athId].ultimoTreino) {
      mapa[athId].ultimoTreino = data;
    }
    if (data >= cutoff && tipo === 'Corrida') {
      mapa[athId].km30     += km;
      mapa[athId].treinos30++;
    }
  }

  // Ordenar por km30 desc
  const ranking = Object.entries(mapa)
    .map(([id, v]) => [id, v.nome, Math.round(v.km30 * 10) / 10, v.treinos30, v.ultimoTreino])
    .sort((a, b) => b[2] - a[2]);

  // Escrever na aba
  // clearContents não remove mesclagens antigas. Desfazê-las evita a falha
  // "selecione todas as células" sem alterar dados fora da área do ranking.
  shRank.getRange(1, 1, Math.max(shRank.getLastRow(), 4), shRank.getMaxColumns()).breakApart();
  shRank.clearContents();

  const tituloRange = shRank.getRange(1, 1, 1, 6);
  tituloRange.merge().setValue('🏆 RANKING — GRUPO HIPERATIVO')
    .setFontWeight('bold').setFontColor('#FFFFFF').setBackground(corAzul)
    .setFontSize(13).setHorizontalAlignment('center');
  shRank.setRowHeight(1, 36);

  const ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
  shRank.getRange(2, 1, 1, 6).merge()
    .setValue('Atualizado em: ' + ts + '  |  Baseado em corridas dos últimos 30 dias')
    .setFontStyle('italic').setFontSize(9).setFontColor('#666666')
    .setHorizontalAlignment('center').setBackground('#F9F9F9');
  shRank.setRowHeight(2, 20);

  const cabecalhos = ['#', 'Atleta', 'ID', 'km (30d)', 'Corridas (30d)', 'Último Treino'];
  const cabRange = shRank.getRange(3, 1, 1, 6);
  cabRange.setValues([cabecalhos]).setFontWeight('bold').setFontColor('#FFFFFF')
    .setBackground('#003D7A').setHorizontalAlignment('center');
  shRank.setRowHeight(3, 24);

  if (ranking.length === 0) {
    shRank.getRange(4, 1).setValue('Sem dados de corrida nos últimos 30 dias.');
    return;
  }

  const linhas = ranking.map(([id, nome, km, treinos, ult], idx) => [
    idx + 1,
    nome || id,
    id,
    km,
    treinos,
    ult instanceof Date ? Utilities.formatDate(ult, Session.getScriptTimeZone(), 'dd/MM/yyyy') : '--'
  ]);

  const dataRange = shRank.getRange(4, 1, linhas.length, 6);
  dataRange.setValues(linhas);

  // Formatação zebra + medalhas
  for (let i = 0; i < linhas.length; i++) {
    const row = i + 4;
    const bg  = i === 0 ? '#FFF9C4' : i === 1 ? '#F5F5F5' : i === 2 ? '#FFF3E0' : (i % 2 === 0 ? corBranco : corClaro);
    shRank.getRange(row, 1, 1, 6).setBackground(bg).setFontSize(10).setHorizontalAlignment('center');
    if (i < 3) shRank.getRange(row, 1).setFontWeight('bold');
    shRank.setRowHeight(row, 22);
  }

  // Coluna km em negrito
  shRank.getRange(4, 4, linhas.length, 1).setFontWeight('bold').setFontColor(corVerde);

  // Larguras
  [6, 22, 13, 10, 14, 14].forEach((w, i) => shRank.setColumnWidth(i + 1, w * 7));

  SpreadsheetApp.flush();
}


// ══════════════════════════════════════════════════════════════════════════════
// STRAVA STATUS — Situação de cada atleta conectado
// ══════════════════════════════════════════════════════════════════════════════
function atualizarStravaStatusSheet() {
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const shTok   = ss.getSheetByName(H.SHEETS.TOKENS);
  const shCad   = ss.getSheetByName(H.SHEETS.CADASTRO);
  const shAtiv  = ss.getSheetByName(H.SHEETS.ATIVIDADES);
  const shSt    = _getAba(ss, 'STRAVA STATUS');
  if (!shSt || !shTok) return;

  shSt.clearContents();

  const corAzul  = '#001F3F';
  const corVerde = '#D6F5EC';
  const corVerm  = '#FDECEA';
  const corAmar  = '#FFF8E1';

  shSt.getRange(1, 1, 1, 8).merge()
    .setValue('📡 STRAVA STATUS — HIPERATIVO V3')
    .setFontWeight('bold').setFontColor('#FFFFFF').setBackground(corAzul)
    .setFontSize(13).setHorizontalAlignment('center');
  shSt.setRowHeight(1, 36);

  const ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
  shSt.getRange(2, 1, 1, 8).merge()
    .setValue('Atualizado em: ' + ts)
    .setFontStyle('italic').setFontSize(9).setFontColor('#666')
    .setHorizontalAlignment('center').setBackground('#F9F9F9');
  shSt.setRowHeight(2, 20);

  const cab = ['Atleta', 'ID', 'ID Strava', 'Token Expira', 'Status', 'Último Treino', 'Qtd Atividades', 'Observação'];
  shSt.getRange(3, 1, 1, 8).setValues([cab])
    .setFontWeight('bold').setFontColor('#FFFFFF').setBackground('#003D7A')
    .setHorizontalAlignment('center');
  shSt.setRowHeight(3, 24);

  const tokDados = shTok.getDataRange().getValues();
  const cadDados = shCad ? shCad.getDataRange().getValues() : [];
  const atvDados = shAtiv ? shAtiv.getDataRange().getValues() : [];
  const agora    = Math.floor(Date.now() / 1000);
  const usoStrava = typeof _mapaUsoStravaCadastro_ === 'function'
    ? _mapaUsoStravaCadastro_() : {};

  const linhas = [];
  for (let i = 1; i < tokDados.length; i++) {
    const athId     = String(tokDados[i][H.TOK.ATH_ID   - 1] || '').trim();
    const nome      = String(tokDados[i][H.TOK.NOME      - 1] || '').trim();
    const stravaId  = String(tokDados[i][H.TOK.STRAVA_ID - 1] || '').trim();
    const expires   = Number(tokDados[i][H.TOK.EXPIRES   - 1]) || 0;
    const tokenStat = String(tokDados[i][H.TOK.STATUS    - 1] || '').trim();
    if (!athId) continue;

    // Data de expiração legível
    const naoUsaStrava = typeof _cadastroNaoUsaStrava_ === 'function'
      && _cadastroNaoUsaStrava_(athId, usoStrava);
    let expDate = expires ? Utilities.formatDate(new Date(expires * 1000), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm') : '--';

    // Status do token
    let statusLabel = tokenStat || 'Desconhecido';
    if (naoUsaStrava) statusLabel = 'Não utiliza';
    else if (expires && expires > agora + 600) statusLabel = 'Ativo';
    else if (expires && expires <= agora) statusLabel = 'Expirado';

    // Último treino do atleta
    let ultimoTreino = '--';
    let qtdAtiv = 0;
    for (let j = 2; j < atvDados.length; j++) {
      if (String(atvDados[j][H.ATIV.ATH_ID - 1] || '').trim() === athId) {
        qtdAtiv++;
        const dataAtiv = atvDados[j][H.ATIV.DATA - 1];
        if (dataAtiv instanceof Date) {
          const dataStr = Utilities.formatDate(dataAtiv, Session.getScriptTimeZone(), 'dd/MM/yyyy');
          if (ultimoTreino === '--' || dataAtiv > new Date(ultimoTreino.split('/').reverse().join('-'))) {
            ultimoTreino = dataStr;
          }
        }
      }
    }

    // Observação — diferencia "sem vínculo mas tem atividades" de "nunca conectou"
    let obs = '';
    if (naoUsaStrava) {
      expDate = '--';
      obs = 'Atleta não utiliza Strava';
    }
    else if (!stravaId && qtdAtiv > 0) obs = 'Atividades importadas — Strava ID ausente no token';
    else if (!stravaId)            obs = 'Strava não vinculado';
    else if (statusLabel === 'Expirado') obs = 'Token expirado — atleta deve reconectar';
    else if (qtdAtiv === 0)        obs = 'Sem atividades importadas ainda';

    linhas.push([nome || athId, athId, stravaId, expDate, statusLabel, ultimoTreino, qtdAtiv, obs]);
  }

  if (linhas.length === 0) {
    shSt.getRange(4, 1).setValue('Nenhum atleta conectado ao Strava ainda.');
    return;
  }

  shSt.getRange(4, 1, linhas.length, 8).setValues(linhas).setFontSize(10).setHorizontalAlignment('center');

  // Formatação condicional manual por status
  for (let i = 0; i < linhas.length; i++) {
    const row = i + 4;
    const bg  = i % 2 === 0 ? '#FFFFFF' : '#F0F7FF';
    shSt.getRange(row, 1, 1, 8).setBackground(bg);

    const status = linhas[i][4];
    let statusBg = bg;
    if (status === 'Ativo')    statusBg = corVerde;
    else if (status === 'Não utiliza') statusBg = '#EEEEEE';
    else if (status === 'Expirado') statusBg = corVerm;
    else statusBg = corAmar;
    shSt.getRange(row, 5).setBackground(statusBg).setFontWeight('bold');
    shSt.setRowHeight(row, 22);
  }

  [18, 13, 13, 16, 11, 14, 14, 30].forEach((w, i) => shSt.setColumnWidth(i + 1, w * 7));
  SpreadsheetApp.flush();
}


// ══════════════════════════════════════════════════════════════════════════════
// STRAVA ENVIOS — Registrar log de importação
// ══════════════════════════════════════════════════════════════════════════════
function logStravaEnvio(athId, nome, qtdNovas, qtdTotal, status, obs) {
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const shSE = _getAba(ss, 'STRAVA ENVIOS');
  if (!shSE) return;

  // Criar cabeçalho se aba vazia
  if (shSE.getLastRow() < 2) {
    shSE.getRange(1, 1, 1, 7).setValues([['Data/Hora', 'ATH_ID', 'Nome', 'Novas Atividades', 'Total Importado', 'Status', 'Observação']])
      .setFontWeight('bold').setBackground('#001F3F').setFontColor('#FFFFFF');
    shSE.setRowHeight(1, 24);
  }

  const ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss');
  shSE.appendRow([ts, athId || '', nome || '', qtdNovas || 0, qtdTotal || 0, status || 'OK', obs || '']);
}



// ══════════════════════════════════════════════════════════════════════════════
// MIGRAR ATIVIDADES ANTIGAS — corrige colunas 14/15/16 para novo formato
// Executa uma vez só (safe: só recalcula, não apaga dados)
// ══════════════════════════════════════════════════════════════════════════════
function migrarAtividades() {
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const sh   = ss.getSheetByName(H.SHEETS.ATIVIDADES);
  if (!sh) return;

  const dados    = sh.getDataRange().getValues();
  let migradas   = 0;

  for (let i = 2; i < dados.length; i++) {
    const row = dados[i];
    const velMps = Number(row[H.ATIV.VEL_MPS - 1]);
    if (!velMps || velMps <= 0) continue;

    // Detectar se dados antigos: col 14 tem km/h (>1) e col 15 tem string
    const col14 = row[H.ATIV.VEL_KMMIN - 1];
    const col15 = row[H.ATIV.PACE_S    - 1];
    const col16 = row[H.ATIV.PACE_FMT  - 1];

    // Se col 14 > 1.0, provavelmente é km/h (não km/min) → migrar
    const ehAntigo = (typeof col14 === 'number' && col14 > 1.0)
                  || (typeof col15 === 'string' && col15.includes(':'));

    if (!ehAntigo) continue;

    // Recalcular com novos valores
    const velKmMin  = Math.round(velMps * 0.06 * 1000) / 1000;
    const paceSegKm = Math.round(1000 / velMps);
    const paceFmt   = _formatarPace(paceSegKm);

    const linhaReal = i + 1; // 1-indexed
    sh.getRange(linhaReal, H.ATIV.VEL_KMMIN, 1, 1).setValue(velKmMin);
    sh.getRange(linhaReal, H.ATIV.PACE_S,    1, 1).setValue(paceSegKm);
    sh.getRange(linhaReal, H.ATIV.PACE_FMT,  1, 1).setValue(paceFmt);

    migradas++;
  }

  SpreadsheetApp.flush();
  _log('SISTEMA', 'INFO', 'migrarAtividades', migradas + ' atividades migradas para novo formato.', '');
  try { SpreadsheetApp.getUi().alert('✅ Migração concluída', migradas + ' atividades atualizadas para o novo formato (km/min, pace em s/km).', SpreadsheetApp.getUi().ButtonSet.OK); } catch(_) {}
}


// ══════════════════════════════════════════════════════════════════════════════
// SINCRONIZAR ATLETAS — garante que todo atleta do CADASTRO existe em todas
// as abas relevantes (MÉTRICAS, TOKENS, RANKING, STRAVA STATUS)
// ══════════════════════════════════════════════════════════════════════════════
function sincronizarAtletasEmTodasAbas() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const shCad = ss.getSheetByName(H.SHEETS.CADASTRO);
  const shMet = ss.getSheetByName(H.SHEETS.METRICAS);
  const shTok = ss.getSheetByName(H.SHEETS.TOKENS);
  if (!shCad) return;

  const cadDados = shCad.getDataRange().getValues();
  let criados = 0;

  for (let i = 1; i < cadDados.length; i++) {
    const athId  = String(cadDados[i][H.CAD.ID     - 1] || '').trim();
    const nome   = String(cadDados[i][H.CAD.NOME   - 1] || '').trim();
    const status = String(cadDados[i][H.CAD.STATUS - 1] || '').trim();
    if (!_isAthIdValido_(athId) || status.toLowerCase() === 'inativo') continue;

    // 1. Garantir linha em MÉTRICAS
    if (shMet) {
      const metDados = shMet.getDataRange().getValues();
      const existeMet = metDados.some((r, idx) => idx > 0 && String(r[H.MET.ATH_ID - 1] || '').trim() === athId);
      if (!existeMet) {
        // Inserir linha placeholder com dados mínimos
        shMet.appendRow([
          athId, nome, new Date(),
          '', '', '', '', '', '', '',
          '', '', '', '', '', '', '', '',
          'Iniciante', 'Moderado', 'Moderado', 'Sem dados ainda', 'Pendente', 'Aguardando atividades Strava'
        ]);
        criados++;
      }
    }

    // 2. Garantir linha em TOKENS (só se não existir — NÃO criar token fake)
    if (shTok) {
      const tokDados = shTok.getDataRange().getValues();
      const existeTok = tokDados.some((r, idx) => idx > 0 && String(r[H.TOK.ATH_ID - 1] || '').trim() === athId);
      if (!existeTok) {
        // Inserir linha de referência sem tokens (status: Pendente)
        shTok.appendRow([
          'TOK_' + athId, athId, nome,
          '', '', '', '', '', '', 'Pendente'
        ]);
      }
    }
  }

  // 3. Atualizar abas visuais
  try { atualizarRankingSheet(); } catch(e) {}
  try { atualizarStravaStatusSheet(); } catch(e) {}

  _log('SISTEMA', 'INFO', 'sincronizarAtletasEmTodasAbas', criados + ' atleta(s) adicionado(s) às abas.', '');
  return criados;
}



// ══════════════════════════════════════════════════════════════════════════════
// CORRIGIR ERROS DA PLANILHA — Remove linhas corrompidas, reconstrói abas visuais
// Seguro: só remove linhas sem refresh token; nunca toca em dados OAuth reais
// ══════════════════════════════════════════════════════════════════════════════
function corrigirErrosDaPlanilha() {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const log = [];

  // ── 1. TOKENS — remover linhas corrompidas ───────────────────────────────────
  // REGRA: só apaga se NÃO tiver refresh token preenchido (col 5 = H.TOK.REFRESH)
  const shTok = ss.getSheetByName(H.SHEETS.TOKENS);
  if (shTok) {
    const tok = shTok.getDataRange().getValues();
    // Varrer de baixo pra cima para não deslocar índices ao deletar
    for (let i = tok.length - 1; i >= 1; i--) {
      const execId  = String(tok[i][H.TOK.EXEC_ID  - 1] || '').trim();
      const athId   = String(tok[i][H.TOK.ATH_ID   - 1] || '').trim();
      const nome    = String(tok[i][H.TOK.NOME      - 1] || '').trim();
      const refresh = String(tok[i][H.TOK.REFRESH   - 1] || '').trim();

      // Tipo 1: linha de cabeçalho duplicado ("TOK_ID Atleta", "ID Atleta" etc.)
      const ehHeader =
        execId.toLowerCase().includes('tok_id') ||
        execId.toLowerCase() === 'exec_id'      ||
        athId.toLowerCase()  === 'id atleta'    ||
        athId.toLowerCase()  === 'ath_id'       ||
        nome.toLowerCase()   === 'nome completo';

      // Tipo 2: linha misalinhada — ATH_ID contém só dígitos (Strava ID numérico
      // vazou para a coluna errada) E o NOME contém hash hexadecimal longo
      const ehMisalinhada =
        /^\d{6,12}$/.test(athId) &&            // athId é ID Strava numérico puro
        /^[0-9a-f]{20,}$/i.test(nome);         // nome é token hash hex

      const corrompida = ehHeader || ehMisalinhada;

      if (corrompida && !refresh) {
        shTok.deleteRow(i + 1);
        const motivo = ehHeader ? 'cabeçalho duplicado' : 'colunas misalinhadas (hash no campo NOME)';
        log.push('🗑️ TOKENS linha ' + (i + 1) + ' removida (' + motivo + ': "' + (execId || athId) + '")');
      }
    }
  }

  // ── 2. MÉTRICAS — remover linhas que são cabeçalhos duplicados ───────────────
  const shMet = ss.getSheetByName(H.SHEETS.METRICAS);
  if (shMet) {
    const met = shMet.getDataRange().getValues();
    for (let i = met.length - 1; i >= 1; i--) {
      const athId = String(met[i][H.MET.ATH_ID - 1] || '').trim();
      const nome  = String(met[i][H.MET.NOME   - 1] || '').trim();

      const corrompida =
        athId.toLowerCase() === 'id atleta' ||
        athId.toLowerCase() === 'ath_id'    ||
        nome.toLowerCase()  === 'nome atleta' ||
        nome.toLowerCase()  === 'nome';

      if (corrompida) {
        shMet.deleteRow(i + 1);
        log.push('🗑️ MÉTRICAS linha ' + (i + 1) + ' removida (header duplicado: "' + athId + '")');
      }
    }
  }

  // ── 3. ATIVIDADES — cabeçalho compatível com o gravador operacional (25 cols) ─
  const shAtiv = ss.getSheetByName(H.SHEETS.ATIVIDADES);
  if (shAtiv) {
    const cabAtiv = [
      'EXEC_ID', 'ATH_ID', 'Nome', 'Data', 'Tipo', 'Fonte', 'Strava ID',
      'Nome Atividade', 'Tempo Movimento', 'Tempo Total', 'Distância (m)',
      'Distância (km)', 'Vel. (m/s)', 'Vel. (km/min)', 'Pace (s/km)',
      'Pace (min:ss)', 'FC Média', 'FC Máxima', 'Elevação (m)', 'Calorias',
      'Cadência (spm)', 'Potência (W)', 'Rota', 'Importado em', 'PSE (1-10)',
    ];
    shAtiv.getRange(2, 1, 1, cabAtiv.length)
      .setValues([cabAtiv])
      .setFontWeight('bold').setFontColor('#FFFFFF')
      .setBackground('#003D7A').setHorizontalAlignment('center');
    log.push('✅ ATIVIDADES: cabeçalho linha 2 atualizado (25 cols, km/min, pace s/km e PSE).');
  }

  // ── 4. STRAVA STATUS — reconstruir do zero (clearContents já está no fn) ──────
  // Coluna C = ID Strava (posição 3 no array de linhas)
  try {
    atualizarStravaStatusSheet();
    log.push('✅ STRAVA STATUS: reconstruído — ID Strava na coluna C.');
  } catch (e) {
    log.push('❌ STRAVA STATUS: erro — ' + e.message);
  }

  // ── 5. RANKING — reconstruir ─────────────────────────────────────────────────
  try {
    atualizarRankingSheet();
    log.push('✅ RANKING: reconstruído.');
  } catch (e) {
    log.push('❌ RANKING: erro — ' + e.message);
  }

  SpreadsheetApp.flush();
  _log('SISTEMA', 'INFO', 'corrigirErrosDaPlanilha', log.join(' | '), '');

  try {
    SpreadsheetApp.getUi().alert(
      '🔧 Correção da Planilha — Concluída',
      log.length ? log.join('\n') : '✅ Nenhum erro encontrado. Tudo limpo.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (_) {}
}


// ══════════════════════════════════════════════════════════════════════════════
// MIGRAR FORMATAÇÃO DE ATIVIDADES EXISTENTES
// Corrige distância (menos dígitos) e ritmo/vel (display por esporte)
// em todas as linhas já gravadas na aba ATIVIDADES.
// Seguro: lê dist_m (col 11) e vel_mps (col 13) como fonte da verdade.
// Menu: ⚙️ Configurações → 📐 Migrar dados de atividades (km/min)
// ══════════════════════════════════════════════════════════════════════════════
function migrarFormatacaoAtividades() {
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const sh   = ss.getSheetByName(H.SHEETS.ATIVIDADES);
  if (!sh) { SpreadsheetApp.getUi().alert('Aba ATIVIDADES não encontrada.'); return; }

  const lastRow = sh.getLastRow();
  if (lastRow < 3) { SpreadsheetApp.getUi().alert('Sem dados para migrar.'); return; }

  const dados = sh.getRange(3, 1, lastRow - 2, 25).getValues();
  const atualizacoes12    = []; // DIST_KM
  const atualizacoes16    = []; // PACE_FMT / VEL display
  const atualizacoesFc    = []; // FC/CAL/ELEV como inteiros
  const atualizacoesTempo = []; // Tempo mov/total → [h]:mm:ss

  dados.forEach((row, idx) => {
    const linhaReal = idx + 3;
    const tipo   = String(row[4]  || '').trim();  // col 5
    const distM  = Number(row[10] || 0);           // col 11
    const velMps = Number(row[12] || 0);           // col 13
    const fcMed  = row[16];
    const fcMax  = row[17];
    const elev   = row[18];
    const cal    = row[19];

    if (!tipo && !distM && !velMps) return; // linha vazia

    // ── Distância com precisão por esporte ──────────────────────────────────
    let distKm = '';
    if (distM > 0) {
      const t = tipo.toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '');
      if (t === 'ciclismo' || t === 'ergometrica' || t === 'hibrido') {
        distKm = Math.round(distM / 100) / 10;      // 1 decimal km
      } else if (t === 'natacao') {
        distKm = Math.round(distM) / 1000;           // metros → km 3 dec
      } else {
        distKm = Math.round(distM / 10) / 100;       // 2 decimais km
      }
    }
    atualizacoes12.push({ row: linhaReal, val: distKm });

    // ── Display pace/vel por esporte ────────────────────────────────────────
    let display = '';
    if (velMps > 0) {
      const t = tipo.toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '');
      if (t === 'ciclismo' || t === 'ergometrica' || t === 'hibrido') {
        display = (Math.round(velMps * 3.6 * 10) / 10).toFixed(1) + ' km/h';
      } else if (t === 'natacao') {
        const s = Math.round(100 / velMps);
        display = Math.floor(s / 60) + ':' + String(s % 60).padStart(2, '0') + ' /100m';
      } else if (velMps > 0) {
        const s = Math.round(1000 / velMps);
        if (s > 0 && s < 3600) {
          display = Math.floor(s / 60) + ':' + String(s % 60).padStart(2, '0') + ' /km';
        }
      }
    }
    atualizacoes16.push({ row: linhaReal, val: display });

    // ── Tempo: converter segundos brutos para fração de dia ([h]:mm:ss) ──────
    // Valores > 1 indicam segundos brutos (ex: 3600); < 1 já estão convertidos
    const tempoMovRaw   = Number(row[8]  || 0);  // col 9
    const tempoTotalRaw = Number(row[9]  || 0);  // col 10
    const tempoMovFrac   = tempoMovRaw   > 1 ? tempoMovRaw   / 86400 : tempoMovRaw;
    const tempoTotalFrac = tempoTotalRaw > 1 ? tempoTotalRaw / 86400 : tempoTotalRaw;
    atualizacoesTempo.push({ row: linhaReal, mov: tempoMovFrac, total: tempoTotalFrac });

    // ── FC, calorias, elevação como inteiros ────────────────────────────────
    atualizacoesFc.push({
      row: linhaReal,
      fcMed: fcMed !== '' ? Math.round(Number(fcMed)) : '',
      fcMax: fcMax !== '' ? Math.round(Number(fcMax)) : '',
      elev:  elev  !== '' ? Math.round(Number(elev))  : '',
      cal:   cal   !== '' ? Math.round(Number(cal))   : '',
    });
  });

  // Escrever em batch (linha por linha para evitar problemas de range não contíguo)
  atualizacoes12.forEach(u => { if (u.val !== '') sh.getRange(u.row, 12).setValue(u.val); });
  atualizacoes16.forEach(u => { if (u.val !== '') sh.getRange(u.row, 16).setValue(u.val); });
  atualizacoesFc.forEach(u => {
    if (u.fcMed !== '') sh.getRange(u.row, 17).setValue(u.fcMed);
    if (u.fcMax !== '') sh.getRange(u.row, 18).setValue(u.fcMax);
    if (u.elev  !== '') sh.getRange(u.row, 19).setValue(u.elev);
    if (u.cal   !== '') sh.getRange(u.row, 20).setValue(u.cal);
  });

  // Tempo: converter segundos → fração de dia e aplicar formato [h]:mm:ss
  atualizacoesTempo.forEach(u => {
    if (u.mov   > 0) { sh.getRange(u.row, 9) .setValue(u.mov);   sh.getRange(u.row, 9) .setNumberFormat('[h]:mm:ss'); }
    if (u.total > 0) { sh.getRange(u.row, 10).setValue(u.total); sh.getRange(u.row, 10).setNumberFormat('[h]:mm:ss'); }
  });

  SpreadsheetApp.flush();
  _log('SISTEMA', 'INFO', 'migrarFormatacaoAtividades',
    (lastRow - 2) + ' linhas migradas (dist, pace/vel, FC, cal)', '');
  SpreadsheetApp.getUi().alert(
    '✅ Migração concluída!',
    (lastRow - 2) + ' atividades atualizadas:\n' +
    '• Distância: menos dígitos (1-2 casas por esporte)\n' +
    '• Ritmo/Vel: min:ss /km | km/h | min:ss /100m\n' +
    '• Tempo: convertido para h:mm:ss\n' +
    '• FC, calorias, elevação: sem decimais',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}


// ══════════════════════════════════════════════════════════════════════════════
// RANKINGS EXPANDIDOS — 8 categorias consolidadas numa aba rica
// Chamado pelo trigger e pelo menu
// ══════════════════════════════════════════════════════════════════════════════
function atualizarRankingExpandido() {
  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const shAtiv = ss.getSheetByName(H.SHEETS.ATIVIDADES);
  const shMet  = ss.getSheetByName(H.SHEETS.METRICAS);
  if (!shAtiv) return;

  let shRE = ss.getSheetByName('🏆 RANKING COMPLETO');
  if (!shRE) shRE = ss.insertSheet('🏆 RANKING COMPLETO');
  shRE.getRange(1, 1, Math.max(shRE.getLastRow(), 4), shRE.getMaxColumns()).breakApart();
  shRE.clearContents();
  shRE.setFrozenRows(2);

  const corAzul  = '#001F3F';
  const dados    = shAtiv.getDataRange().getValues().slice(2);
  const metDados = shMet ? shMet.getDataRange().getValues().slice(2) : [];
  const agora    = new Date();
  const d30  = new Date(agora.getTime() - 30  * 86400000);
  const d90  = new Date(agora.getTime() - 90  * 86400000);
  const d28  = new Date(agora.getTime() - 28  * 86400000);
  const ts   = Utilities.formatDate(agora, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');

  // Cabeçalho geral
  shRE.getRange(1, 1, 1, 8).merge()
    .setValue('🏆 RANKING COMPLETO — GRUPO HIPERATIVO')
    .setFontWeight('bold').setFontColor('#FFFFFF').setBackground(corAzul)
    .setFontSize(14).setHorizontalAlignment('center');
  shRE.setRowHeight(1, 38);
  shRE.getRange(2, 1, 1, 8).merge()
    .setValue('Atualizado: ' + ts)
    .setFontStyle('italic').setFontSize(9).setFontColor('#666')
    .setHorizontalAlignment('center').setBackground('#F9F9F9');
  shRE.setRowHeight(2, 20);

  // ── Construir mapa por atleta ─────────────────────────────────────────────
  const mapa = {};
  dados.forEach(row => {
    const athId = String(row[H.ATIV.ATH_ID  - 1] || '').trim();
    const nome  = String(row[H.ATIV.NOME    - 1] || '').trim();
    const data  = row[H.ATIV.DATA   - 1];
    const tipo  = String(row[H.ATIV.TIPO    - 1] || '').trim();
    const km    = Number(row[H.ATIV.DIST_KM - 1]) || 0;
    const paceS = Number(row[H.ATIV.PACE_S  - 1]) || 0;
    const elev  = Number(row[H.ATIV.ELEV    - 1]) || 0;
    const pse   = Number(row[H.ATIV.PSE     - 1]) || 0;
    if (!_isAthIdValido_(athId) || !(data instanceof Date)) return;

    if (!mapa[athId]) mapa[athId] = {
      nome, athId,
      km30:0, km90:0, kmTotal:0,
      treinos30:0, treinos90:0, treinosTotal:0,
      tiposMap:{}, // {Corrida:N, Ciclismo:N, ...}
      melhoresPace:[], // paces válidos (s/km) das corridas
      elev30:0, elevTotal:0,
      pseTotal:0, pseCount:0,
      ultimoTreino:null,
      semanasAtivas: new Set(),
    };

    const m = mapa[athId];
    m.kmTotal += km;
    m.treinosTotal++;
    m.tiposMap[tipo] = (m.tiposMap[tipo] || 0) + 1;
    m.elevTotal += elev;
    if (!m.ultimoTreino || data > m.ultimoTreino) m.ultimoTreino = data;

    const semana = Utilities.formatDate(data, Session.getScriptTimeZone(), 'yyyy-ww');
    m.semanasAtivas.add(semana);

    if (pse > 0) { m.pseTotal += pse; m.pseCount++; }

    if (data >= d30) {
      m.km30     += km;
      m.elev30   += elev;
      m.treinos30++;
    }
    if (data >= d90) {
      m.km90     += km;
      m.treinos90++;
    }

    // Melhores paces válidos (só corrida, pace razoável 200-900s/km)
    if ((tipo === 'Corrida' || tipo === 'Trail Run') && paceS > 200 && paceS < 900) {
      m.melhoresPace.push(paceS);
    }
  });

  // Métricas VO2max da aba MÉTRICAS
  const vo2Map = {};
  metDados.forEach(row => {
    const id  = String(row[H.MET.ATH_ID - 1] || '').trim();
    const vo2 = Number(row[H.MET.VO2    - 1]) || 0;
    if (_isAthIdValido_(id) && vo2 > 0) vo2Map[id] = vo2;
  });

  // Converter set em número
  Object.values(mapa).forEach(m => {
    m.semanasAtivasN = m.semanasAtivas.size;
    m.km30     = Math.round(m.km30     * 10) / 10;
    m.km90     = Math.round(m.km90     * 10) / 10;
    m.elev30   = Math.round(m.elev30);
    m.elevTotal= Math.round(m.elevTotal);
    m.pseMed   = m.pseCount > 0 ? Math.round(m.pseTotal / m.pseCount * 10) / 10 : null;
    m.melhorPace = m.melhoresPace.length ? Math.min(...m.melhoresPace) : null;
    m.vo2      = vo2Map[m.athId] || 0;
  });

  const ats = Object.values(mapa);

  let linha = 3; // linha atual na aba

  // ── Helper: escrever uma seção de ranking ───────────────────────────────
  const escreverRanking = (titulo, itens, colunas, cor) => {
    shRE.getRange(linha, 1, 1, 8).merge()
      .setValue(titulo)
      .setFontWeight('bold').setFontColor('#FFFFFF').setBackground(cor || '#003D7A')
      .setFontSize(11).setHorizontalAlignment('center');
    shRE.setRowHeight(linha, 28);
    linha++;

    const cab = colunas.map(c => c.titulo);
    shRE.getRange(linha, 1, 1, cab.length).setValues([cab])
      .setFontWeight('bold').setBackground('#E8EEF6').setFontSize(9)
      .setHorizontalAlignment('center');
    shRE.setRowHeight(linha, 20);
    linha++;

    if (!itens.length) {
      shRE.getRange(linha, 1).setValue('Sem dados suficientes.');
      linha++;
    } else {
      itens.slice(0, 15).forEach((it, i) => {
        const vals = colunas.map(c => c.valor(it, i));
        shRE.getRange(linha, 1, 1, vals.length).setValues([vals]).setFontSize(10);
        const bg = i === 0 ? '#FFF9C4' : i === 1 ? '#F5F5F5' : i === 2 ? '#FFF3E0' : (i % 2 === 0 ? '#FFFFFF' : '#F0F7FF');
        shRE.getRange(linha, 1, 1, vals.length).setBackground(bg).setHorizontalAlignment('center');
        if (i < 3) shRE.getRange(linha, 1).setFontWeight('bold');
        shRE.setRowHeight(linha, 20);
        linha++;
      });
    }
    linha++; // espaço entre seções
  };

  const medal = ['🥇', '🥈', '🥉'];
  const pos   = (i) => medal[i] || (i + 1) + 'º';
  const fmtPaceRanking = (s) => { if (!s) return '—'; const m = Math.floor(s/60); return m + ':' + String(s%60).padStart(2,'0') + '/km'; };
  const fmtData = (d) => d instanceof Date ? Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd/MM/yy') : '—';

  // ── 1. KM nos últimos 30 dias ─────────────────────────────────────────────
  escreverRanking('🏃 TOP KM — ÚLTIMOS 30 DIAS',
    ats.filter(a => a.km30 > 0).sort((a,b) => b.km30 - a.km30),
    [
      { titulo:'#',             valor:(it,i) => pos(i) },
      { titulo:'Atleta',        valor:(it)   => it.nome || it.athId },
      { titulo:'ID',            valor:(it)   => it.athId },
      { titulo:'km (30d)',      valor:(it)   => it.km30 },
      { titulo:'Treinos (30d)', valor:(it)   => it.treinos30 },
      { titulo:'Último Treino', valor:(it)   => fmtData(it.ultimoTreino) },
    ], '#1565C0');

  // ── 2. Melhor Pace (corrida) ──────────────────────────────────────────────
  escreverRanking('⚡ MELHOR PACE (CORRIDA)',
    ats.filter(a => a.melhorPace).sort((a,b) => a.melhorPace - b.melhorPace),
    [
      { titulo:'#',           valor:(it,i) => pos(i) },
      { titulo:'Atleta',      valor:(it)   => it.nome || it.athId },
      { titulo:'Melhor Pace', valor:(it)   => fmtPaceRanking(it.melhorPace) },
      { titulo:'VO2max est.', valor:(it)   => it.vo2 || '—' },
      { titulo:'Último Treino', valor:(it) => fmtData(it.ultimoTreino) },
    ], '#00695C');

  // ── 3. Mais Consistente (semanas ativas em 90 dias) ───────────────────────
  escreverRanking('🗓️ MAIS CONSISTENTE (90 DIAS)',
    ats.filter(a => a.semanasAtivasN > 0).sort((a,b) => b.semanasAtivasN - a.semanasAtivasN || b.treinos90 - a.treinos90),
    [
      { titulo:'#',             valor:(it,i) => pos(i) },
      { titulo:'Atleta',        valor:(it)   => it.nome || it.athId },
      { titulo:'Semanas Ativas',valor:(it)   => it.semanasAtivasN },
      { titulo:'Treinos (90d)', valor:(it)   => it.treinos90 },
      { titulo:'km (90d)',      valor:(it)   => it.km90 },
    ], '#4527A0');

  // ── 4. Maior Elevação (30 dias) ───────────────────────────────────────────
  escreverRanking('⛰️ MAIOR ELEVAÇÃO ACUMULADA (30d)',
    ats.filter(a => a.elev30 > 0).sort((a,b) => b.elev30 - a.elev30),
    [
      { titulo:'#',           valor:(it,i) => pos(i) },
      { titulo:'Atleta',      valor:(it)   => it.nome || it.athId },
      { titulo:'Elev. (30d)', valor:(it)   => it.elev30 + ' m' },
      { titulo:'Treinos',     valor:(it)   => it.treinos30 },
    ], '#4E342E');

  // ── 5. Maior VO2max ───────────────────────────────────────────────────────
  escreverRanking('🫁 MAIOR VO2MÁX ESTIMADO',
    ats.filter(a => a.vo2 > 0).sort((a,b) => b.vo2 - a.vo2),
    [
      { titulo:'#',       valor:(it,i) => pos(i) },
      { titulo:'Atleta',  valor:(it)   => it.nome || it.athId },
      { titulo:'VO2max',  valor:(it)   => it.vo2 + ' ml/kg/min' },
      { titulo:'Pace Méd',valor:(it)   => fmtPaceRanking(it.melhorPace) },
    ], '#1B5E20');

  // ── 6. PSE Médio (controle de esforço) ───────────────────────────────────
  escreverRanking('💢 PSE MÉDIO (PERCEPÇÃO DE ESFORÇO)',
    ats.filter(a => a.pseMed !== null).sort((a,b) => b.pseCount - a.pseCount),
    [
      { titulo:'#',          valor:(it,i) => pos(i) },
      { titulo:'Atleta',     valor:(it)   => it.nome || it.athId },
      { titulo:'PSE Médio',  valor:(it)   => it.pseMed },
      { titulo:'Registros',  valor:(it)   => it.pseCount },
      { titulo:'Interpretação', valor:(it) => it.pseMed <= 4 ? 'Fácil/Moderado' : it.pseMed <= 7 ? 'Intenso' : 'Máximo/RPE alto' },
    ], '#E65100');

  // ── 7. Diversidade de Modalidades ────────────────────────────────────────
  escreverRanking('🏅 DIVERSIDADE DE MODALIDADES',
    ats.sort((a,b) => Object.keys(b.tiposMap).length - Object.keys(a.tiposMap).length || b.treinosTotal - a.treinosTotal),
    [
      { titulo:'#',           valor:(it,i) => pos(i) },
      { titulo:'Atleta',      valor:(it)   => it.nome || it.athId },
      { titulo:'Modalidades', valor:(it)   => Object.keys(it.tiposMap).length },
      { titulo:'Tipo Principal', valor:(it) => Object.entries(it.tiposMap).sort((a,b)=>b[1]-a[1])[0]?.[0] || '—' },
      { titulo:'Total Treinos', valor:(it) => it.treinosTotal },
    ], '#283593');

  // ── 8. Volume total (todos os tempos) ────────────────────────────────────
  escreverRanking('📦 VOLUME TOTAL (TODOS OS TEMPOS)',
    ats.filter(a => a.kmTotal > 0).sort((a,b) => b.kmTotal - a.kmTotal),
    [
      { titulo:'#',             valor:(it,i) => pos(i) },
      { titulo:'Atleta',        valor:(it)   => it.nome || it.athId },
      { titulo:'Total Treinos', valor:(it)   => it.treinosTotal },
      { titulo:'Elev. Total',   valor:(it)   => it.elevTotal + ' m' },
      { titulo:'Último Treino', valor:(it)   => fmtData(it.ultimoTreino) },
    ], '#37474F');

  // Ajustar larguras de coluna
  [6, 22, 12, 14, 14, 14, 14, 14].forEach((w, i) => shRE.setColumnWidth(i + 1, w * 7));
  SpreadsheetApp.flush();
  _log('SISTEMA', 'INFO', 'atualizarRankingExpandido', 'Rankings expandidos atualizados. ' + ats.length + ' atletas.', '');
}
