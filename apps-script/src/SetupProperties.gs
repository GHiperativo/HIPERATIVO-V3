/**
 * HIPERATIVO V3 — configuração segura de propriedades essenciais.
 *
 * Objetivo: evitar que SPREADSHEET_ID fique dependente de valor hardcoded.
 * Esta rotina não registra nem exibe credenciais sensíveis.
 */

/**
 * Grava nas Script Properties o ID da planilha ativa.
 * Pode ser executada com segurança a partir do editor do Apps Script.
 */
function configurarSpreadsheetId() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    throw new Error('Nenhuma planilha ativa encontrada. Execute esta função no projeto vinculado ao HIPERATIVO V3.');
  }

  PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', ss.getId());

  try {
    SpreadsheetApp.getUi().alert(
      '✅ SPREADSHEET_ID configurado',
      'O projeto agora aponta para a planilha ativa. O valor foi salvo nas Script Properties e não foi exposto na interface.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (_) {}
}

/**
 * Diagnóstico sem exibir os valores das propriedades.
 * Retorna apenas CONFIGURADO / AUSENTE.
 */
function diagnosticarPropriedadesEssenciais() {
  const props = PropertiesService.getScriptProperties();
  const status = {
    ADMIN_EMAIL: props.getProperty('ADMIN_EMAIL') ? 'CONFIGURADO' : 'AUSENTE',
    SPREADSHEET_ID: props.getProperty('SPREADSHEET_ID') ? 'CONFIGURADO' : 'AUSENTE'
  };

  const mensagem =
    'ADMIN_EMAIL: ' + status.ADMIN_EMAIL + '\n' +
    'SPREADSHEET_ID: ' + status.SPREADSHEET_ID + '\n\n' +
    'Nenhum valor foi exibido.';

  Logger.log(mensagem);
  try {
    SpreadsheetApp.getUi().alert('🩺 Propriedades essenciais', mensagem, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (_) {}

  return status;
}
