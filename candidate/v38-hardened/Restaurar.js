/**
 * ═══════════════════════════════════════════════════════════════════════
 * HIPERATIVO V3 — Restaurar.gs
 * Restauração de emergência dos atletas após wipe do restaurarEstrutura
 * Execute: restaurarAtletas() no Apps Script editor
 * ═══════════════════════════════════════════════════════════════════════
 */

function restaurarAtletas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Dados conhecidos (fonte: STRAVA ENVIOS + RANKING sobreviventes)
  const ATLETAS = [
    { id: 'ATH029112',         nome: 'Giselle Mamede',                   email: 'gisellemamede@hotmail.com',     stravaOk: 'Reconectar', stravaId: '' },
    { id: 'ATH572491',         nome: 'Cynara Costa',                      email: 'ccsrego@gmail.com',             stravaOk: 'Reconectar', stravaId: '' },
    { id: 'ATH992736',         nome: 'Crhystiano Heliodoro',              email: 'crhystianoh@gmail.com',         stravaOk: 'Reconectar', stravaId: '' },
    { id: 'ATHDE549E0A',       nome: 'Alessandra Caetano Queiroz',        email: 'alessandraqueiroz48@gmail.com', stravaOk: 'Reconectar', stravaId: '' },
    { id: 'ATHF2A39037',       nome: 'Amanda Soares de Melo',             email: 'mandenhasoares@gmail.com',      stravaOk: 'Pendente',   stravaId: '' },
    { id: 'ATH6333AE19',       nome: 'Lorrana Padua',                     email: 'lorranapadua@gmail.com',        stravaOk: 'Pendente',   stravaId: '' },
    { id: 'ATH_1781116630575', nome: 'Rachel Espindola Sales de Souza',   email: '',                              stravaOk: '',           stravaId: '' },
  ];

  const agora = new Date().toLocaleString('pt-BR', { timeZone: 'America/Sao_Paulo' });
  const log   = [];

  try {
    // ── 1. CADASTRO ────────────────────────────────────────────────────────
    const shCad = ss.getSheetByName(H.SHEETS.CADASTRO);
    if (!shCad) throw new Error('Aba CADASTRO nao encontrada. Execute Setup primeiro.');

    const cadRows  = shCad.getDataRange().getValues();
    const idsExist = new Set();
    for (let i = 1; i < cadRows.length; i++) {
      const v = String(cadRows[i][0] || '').trim().toUpperCase();
      if (v) idsExist.add(v);
    }

    let cadAdd = 0, cadUpd = 0;
    for (const atl of ATLETAS) {
      const uid = atl.id.toUpperCase();
      const linha = [
        atl.id, atl.nome, atl.email,
        '', '', '', '', '',        // whats nasc sexo peso altura
        'Corrida',                 // mod
        '', '', '', '', '', '', '', // nivel obj freq horario saude lesao med
        'Corrida CCC',             // prova
        'Corrida CCC',             // plano
        'Brasilia', 'DF', '',      // cidade estado cpf
        'Restauracao de emergencia', agora, // origem data_cad
        atl.stravaOk, atl.stravaId || '', // strava_ok strava_id
        'Ativo',                   // status
        'Restaurado em ' + agora,  // obs
      ];

      if (!idsExist.has(uid)) {
        shCad.appendRow(linha);
        cadAdd++;
      } else {
        // Apenas preenche celulas vazias + atualiza STATUS
        for (let i = 1; i < cadRows.length; i++) {
          if (String(cadRows[i][0] || '').trim().toUpperCase() === uid) {
            const r = i + 1;
            if (!cadRows[i][1]) shCad.getRange(r, 2).setValue(atl.nome);
            if (!cadRows[i][2]) shCad.getRange(r, 3).setValue(atl.email);
            if (!cadRows[i][H.CAD.STRAVA_OK - 1] && atl.stravaOk)
              shCad.getRange(r, H.CAD.STRAVA_OK).setValue(atl.stravaOk);
            shCad.getRange(r, H.CAD.STATUS).setValue('Ativo');
            cadUpd++;
            break;
          }
        }
      }
    }
    SpreadsheetApp.flush();
    log.push('CADASTRO: ' + cadAdd + ' adicionado(s), ' + cadUpd + ' atualizado(s).');

    // ── 2. TOKENS (tenta recuperar do PropertiesService antes de criar placeholder) ─
    const shTok  = ss.getSheetByName(H.SHEETS.TOKENS);
    const props  = PropertiesService.getScriptProperties();
    if (shTok) {
      const tokRows = shTok.getDataRange().getValues();
      const tokIds  = new Set();
      for (let i = 1; i < tokRows.length; i++) {
        const v = String(tokRows[i][H.TOK.ATH_ID - 1] || '').trim().toUpperCase();
        if (v) tokIds.add(v);
      }

      let tokAdd = 0, tokRestored = 0;
      for (const atl of ATLETAS) {
        if (!atl.stravaOk) continue;
        if (tokIds.has(atl.id.toUpperCase())) continue;

        // Verificar se existe backup de refresh_token no PropertiesService
        const backupRt = props.getProperty('RT_' + atl.id.toUpperCase()) || '';
        const execId   = 'TOK_' + Utilities.getUuid().substring(0, 8).toUpperCase();

        if (backupRt) {
          // Temos backup — criar linha com refresh_token real (acesso será renovado automaticamente)
          shTok.appendRow([
            execId, atl.id, atl.nome,
            '',        // access_token (será renovado em _getValidAccessToken)
            backupRt,  // refresh_token — RESTAURADO DO BACKUP
            '',        // expires_at
            'read,activity:read_all,profile:read_all',
            atl.stravaId || '',
            agora,
            'Restaurado (token backup)',
          ]);
          tokRestored++;
        } else {
          // Sem backup — placeholder: atleta precisará reconectar
          shTok.appendRow([
            execId, atl.id, atl.nome,
            '', '', '',
            'read,activity:read_all,profile:read_all',
            atl.stravaId || '',
            agora,
            atl.stravaOk === 'Reconectar' ? 'Reconectar Strava' : 'Aguardando',
          ]);
          tokAdd++;
        }
      }
      SpreadsheetApp.flush();
      log.push('TOKENS: ' + tokRestored + ' restaurado(s) com backup, ' + tokAdd + ' placeholder(s) sem backup (precisam reconectar).');
    } else {
      log.push('TOKENS: aba nao encontrada — pulado.');
    }

    // ── 3. METRICAS (estrutura basica) ────────────────────────────────────
    const shMet = ss.getSheetByName(H.SHEETS.METRICAS);
    if (shMet) {
      const metRows = shMet.getDataRange().getValues();
      const metIds  = new Set();
      for (let i = 1; i < metRows.length; i++) {
        const v = String(metRows[i][0] || '').trim().toUpperCase();
        if (v) metIds.add(v);
      }

      let metAdd = 0;
      for (const atl of ATLETAS) {
        if (metIds.has(atl.id.toUpperCase())) continue;
        const arr = new Array(24).fill('');
        arr[0] = atl.id;
        arr[1] = atl.nome;
        arr[2] = agora;
        arr[21] = 'Restauracao';
        arr[23] = 'Restaurado em ' + agora;
        shMet.appendRow(arr);
        metAdd++;
      }
      SpreadsheetApp.flush();
      log.push('METRICAS: ' + metAdd + ' linha(s) criada(s).');
    } else {
      log.push('METRICAS: aba nao encontrada — pulado.');
    }

    // ── 4. STRAVA STATUS (reconstruir via funcao existente) ───────────────
    try { atualizarStravaStatus(); log.push('STRAVA STATUS: reconstruido.'); }
    catch(e) { log.push('STRAVA STATUS: ' + e.message); }

    // ── 5. RANKING ────────────────────────────────────────────────────────
    try { atualizarRanking(); log.push('RANKING: atualizado.'); }
    catch(e) { log.push('RANKING: ' + e.message); }

    // ── Log + alerta ─────────────────────────────────────────────────────
    _log('SYSTEM', 'INFO', 'restaurarAtletas', 'Restauracao concluida.', log.join(' | '));

    // Verificar se algum token foi restaurado do backup (boa notícia)
    const tokRestMsg = log.find(m => m.startsWith('TOKENS:')) || '';
    const temBackup  = tokRestMsg.includes('restaurado');
    const aviso      = temBackup
      ? '\n\n✅ TOKENS RESTAURADOS DO BACKUP — atletas nao precisam reconectar!'
      : '\n\nATENCAO: atletas com Strava precisarao RECONECTAR (sem backup disponivel).\n' +
        'Use: menu Atletas > Enviar link Strava por email.';

    try {
      SpreadsheetApp.getUi().alert(
        'Restauracao Concluida',
        '✅ ' + log.join('\n') + aviso,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    } catch(_) { Logger.log(log.join('\n') + aviso); }

    return { ok: true, log: log };

  } catch(e) {
    _log('SYSTEM', 'ERRO', 'restaurarAtletas', e.message, e.stack || '');
    try { SpreadsheetApp.getUi().alert('Erro', e.message, SpreadsheetApp.getUi().ButtonSet.OK); } catch(_) {}
    return { ok: false, erro: e.message };
  }
}

// ═══════════════════════════════════════════════════════════════════════
// limparPlanilha() — Remove abas desconhecidas e limpa células com erro
// Execute no editor: selecione limparPlanilha > Executar
// ═══════════════════════════════════════════════════════════════════════
function limparPlanilha() {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const ui  = SpreadsheetApp.getUi();
  const log = [];

  // ── Abas legítimas do sistema ──────────────────────────────────────────────
  const ABAS_CONHECIDAS = new Set([
    H.SHEETS.PAINEL,      // 📊 PAINEL
    H.SHEETS.CADASTRO,    // 👤 CADASTRO
    H.SHEETS.ATIVIDADES,  // 🏃 ATIVIDADES
    H.SHEETS.PLANO,       // 📅 PLANO SEMANAL
    H.SHEETS.METRICAS,    // 📈 MÉTRICAS
    H.SHEETS.FEEDBACK,    // 💬 FEEDBACK
    H.SHEETS.GRAFICOS,    // 📉 GRÁFICOS
    H.SHEETS.ERROS,       // 🔴 ERROS
    H.SHEETS.TOKENS,      // 🔐 TOKENS
    H.SHEETS.CONFIG,      // ⚙️ CONFIG
    '🏆 RANKING COMPLETO',// criada por atualizarRankingExpandido
    '🔬 ANÁLISE',         // criada por atualizarAnaliseSheet
    'STRAVA STATUS',      // criada por atualizarStravaStatusSheet
  ]);

  try {
    // ── 1. Identificar abas desconhecidas ──────────────────────────────────
    const todas      = ss.getSheets();
    const remover    = [];
    const preservar  = [];

    for (const sh of todas) {
      const nome = sh.getName();
      if (ABAS_CONHECIDAS.has(nome)) {
        preservar.push(nome);
      } else {
        remover.push(nome);
      }
    }

    if (remover.length === 0) {
      log.push('Nenhuma aba estranha encontrada (' + preservar.length + ' abas OK).');
    } else {
      // Confirmar antes de deletar
      const confirmacao = ui.alert(
        '🗑️ Remover abas desconhecidas?',
        'As seguintes abas NAO sao do sistema e serao excluidas:\n\n• ' +
        remover.join('\n• ') +
        '\n\nAbas preservadas: ' + preservar.length +
        '\n\nDeseja continuar?',
        ui.ButtonSet.YES_NO
      );

      if (confirmacao === ui.Button.YES) {
        for (const nome of remover) {
          try {
            ss.deleteSheet(ss.getSheetByName(nome));
            log.push('Aba removida: "' + nome + '"');
          } catch(e) {
            log.push('Aba nao removida: "' + nome + '" — ' + e.message);
          }
        }
      } else {
        log.push('Remocao cancelada pelo usuario.');
      }
    }

    // ── 2. Limpar células com erros de fórmula ────────────────────────────
    const ERROS_FORMULA = ['#ERROR!', '#N/A', '#REF!', '#VALUE!', '#NAME?', '#DIV/0!', '#NUM!', '#NULL!'];
    let totalErros = 0;

    for (const sh of ss.getSheets()) {
      // Pular abas de dados brutos (não têm fórmulas)
      const n = sh.getName();
      if (n === H.SHEETS.ATIVIDADES || n === H.SHEETS.ERROS || n === H.SHEETS.TOKENS) continue;

      const range  = sh.getDataRange();
      const values = range.getValues();
      for (let r = 0; r < values.length; r++) {
        for (let c = 0; c < values[r].length; c++) {
          const val = String(values[r][c] || '');
          if (ERROS_FORMULA.some(e => val.startsWith(e))) {
            sh.getRange(r + 1, c + 1).clearContent();
            totalErros++;
          }
        }
      }
    }
    log.push('Celulas com erro limpas: ' + totalErros);

    SpreadsheetApp.flush();
    _log('SYSTEM', 'INFO', 'limparPlanilha', 'Limpeza concluida.', log.join(' | '));

    ui.alert(
      '✅ Limpeza Concluída',
      log.join('\n'),
      ui.ButtonSet.OK
    );

    return { ok: true, log: log };

  } catch(e) {
    _log('SYSTEM', 'ERRO', 'limparPlanilha', e.message, e.stack || '');
    try { ui.alert('Erro', e.message, ui.ButtonSet.OK); } catch(_) {}
    return { ok: false, erro: e.message };
  }
}
