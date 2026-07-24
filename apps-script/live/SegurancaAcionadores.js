/**
 * Migra as automações essenciais para a conta corporativa oficial.
 *
 * Esta função não lê nem altera tokens. Ela remove apenas os acionadores
 * gerenciados pertencentes à conta que a executar e recria o conjunto atual
 * e seguro por meio de configurarTriggers(true).
 */
function migrarAcionadoresParaContaOficial() {
  var contaOficial = 'contato@ghiperativo.com.br';
  var contaAtual = String(Session.getEffectiveUser().getEmail() || '')
    .trim()
    .toLowerCase();

  if (contaAtual !== contaOficial) {
    throw new Error(
      'Migração permitida somente para ' + contaOficial + '. Conta atual: ' +
      (contaAtual || 'não identificada')
    );
  }

  var resultado = configurarTriggers(true);
  resultado.contaOficial = contaAtual;
  resultado.rotinasLegadasDesativadas = [
    'atualizarStatusStravaEmCadastro',
    'sincronizarFilaWhatsAppCadastros (periódico)'
  ];

  Logger.log(JSON.stringify(resultado));
  return resultado;
}
