/**
 * Reconcilia erros históricos da integração SHE ↔ V3.
 *
 * Regra: uma linha ERRO só é encerrada quando existe, para o mesmo e-mail,
 * uma linha posterior com vínculo confirmado. O registro é preservado para
 * auditoria; nenhuma aba do Strava, atividade, token ou cadastro é alterada.
 */
const SHE_V3_RECONCILIACAO_CONFIG = Object.freeze({
  crmSpreadsheetId: '1Sn58dggyaalkNyVWWCAqx8U-xXyx8uDtaq3phLR7_GE',
  sheetName: 'INTEGRACAO_SHE_V3',
  timezone: 'America/Sao_Paulo'
});

function reconciliarErrosHistoricosSHEV3() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return { ok: false, motivo: 'PROCESSAMENTO_EM_ANDAMENTO' };

  try {
    const ss = SpreadsheetApp.openById(SHE_V3_RECONCILIACAO_CONFIG.crmSpreadsheetId);
    const sh = ss.getSheetByName(SHE_V3_RECONCILIACAO_CONFIG.sheetName);
    if (!sh || sh.getLastRow() < 2) return { ok: true, corrigidos: 0 };

    const dados = sh.getRange(2, 1, sh.getLastRow() - 1, 12).getValues();
    const finaisPorEmail = {};
    const statusFinais = {
      VINCULADO: true,
      VINCULADO_COM_ATENCAO: true,
      CRIADO_E_VINCULADO: true
    };

    dados.forEach(function(linha) {
      const email = _sheV3ReconEmail_(linha[4]);
      const status = String(linha[6] || '').trim().toUpperCase();
      if (!email || !statusFinais[status]) return;

      const tempo = _sheV3ReconTempo_(linha[0]);
      const atual = finaisPorEmail[email];
      if (!atual || tempo >= atual.tempo) {
        finaisPorEmail[email] = { tempo: tempo, linha: linha.slice() };
      }
    });

    const agora = Utilities.formatDate(
      new Date(),
      SHE_V3_RECONCILIACAO_CONFIG.timezone,
      'dd/MM/yyyy HH:mm'
    );
    let corrigidos = 0;

    dados.forEach(function(linha, indice) {
      const status = String(linha[6] || '').trim().toUpperCase();
      if (status !== 'ERRO') return;

      const email = _sheV3ReconEmail_(linha[4]);
      const final = finaisPorEmail[email];
      if (!email || !final || final.tempo < _sheV3ReconTempo_(linha[0])) return;

      const idErro = String(linha[1] || 'ERRO_SEM_ID');
      const novaLinha = final.linha.slice(1, 12);
      novaLinha[5] = 'RESOLVIDO';
      novaLinha[6] = 'REPROCESSADO_COM_SUCESSO';
      novaLinha[10] = 'Erro histórico ' + idErro +
        ' resolvido automaticamente em ' + agora +
        '. Vínculo posterior confirmado; nenhum cadastro ou token foi alterado.';

      sh.getRange(indice + 2, 2, 1, 11).setValues([novaLinha]);
      corrigidos++;
    });

    SpreadsheetApp.flush();
    return { ok: true, corrigidos: corrigidos };
  } finally {
    lock.releaseLock();
  }
}

function instalarReconciliacaoErrosSHEV3() {
  ScriptApp.getProjectTriggers()
    .filter(function(trigger) {
      return trigger.getHandlerFunction() === 'reconciliarErrosHistoricosSHEV3';
    })
    .forEach(function(trigger) { ScriptApp.deleteTrigger(trigger); });

  ScriptApp.newTrigger('reconciliarErrosHistoricosSHEV3')
    .timeBased()
    .everyDays(1)
    .atHour(6)
    .create();

  return reconciliarErrosHistoricosSHEV3();
}

function _sheV3ReconEmail_(valor) {
  return String(valor || '').trim().toLowerCase();
}

function _sheV3ReconTempo_(valor) {
  if (valor instanceof Date && !isNaN(valor.getTime())) return valor.getTime();
  const data = new Date(valor);
  return isNaN(data.getTime()) ? 0 : data.getTime();
}
