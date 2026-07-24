/**
 * HIPERATIVO V3 — estado operacional único da conexão Strava.
 *
 * X  = declaração/ação ("Sim", "Não", "Pendente", "Reconectar")
 * Z  = situação do aluno no programa ("Ativo", "Inativo"...)
 * AX = conexão Strava derivada ("Conectado", "Aguardando conexão",
 *      "Não utiliza" ou "Reconectar")
 * AY = instante da última avaliação.
 *
 * Nenhuma função deste arquivo apaga ou substitui refresh tokens.
 */

function _normalizarDeclaracaoStrava_(valor) {
  return String(valor || '').trim().toLowerCase()
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '');
}

function _calcularStatusConexaoStrava_(athId, declaracao, opcoes) {
  opcoes = opcoes || {};
  if (opcoes.forcarStatus) return String(opcoes.forcarStatus);

  var id = String(athId || '').trim().toUpperCase();
  var uso = _normalizarDeclaracaoStrava_(declaracao);

  // A declaração explícita "Não" é operacional: não importar, não alertar e
  // não pedir reconexão. Uma credencial histórica pode continuar guardada
  // apenas como backup, sem transformar o atleta em usuário do Strava.
  if (uso === 'nao' || uso === 'nao utiliza') return 'Não utiliza';

  var props = PropertiesService.getScriptProperties();
  if (id && props.getProperty('STRAVA_REVOGADO_' + id)) return 'Reconectar';

  var conectado = false;
  try { conectado = id && _temRefreshTokenValido_(id); } catch (_) {}
  if (conectado) return 'Conectado';

  if (uso === 'reconectar') return 'Reconectar';
  return 'Aguardando conexão';
}

function _dadosFilaCadastroLinha_(linha) {
  return {
    athId: String(linha[H.CAD.ID - 1] || '').trim().toUpperCase(),
    nome: String(linha[H.CAD.NOME - 1] || '').trim(),
    email: String(linha[H.CAD.EMAIL - 1] || '').trim(),
    whats: String(linha[H.CAD.WHATS - 1] || '').trim(),
    dataCadastro: linha[H.CAD.DATA_CAD - 1] || '',
    stravaOk: linha[H.CAD.STRAVA_OK - 1],
    stravaId: linha[H.CAD.STRAVA_ID - 1]
  };
}

/**
 * Atualização imediata para o gatilho simples onEdit.
 *
 * O gatilho simples não deve consultar Supabase, Properties ou outras fontes
 * que dependam de autorização. Ele trata somente a declaração recém-editada e
 * preserva um "Conectado" já confirmado. OAuth, refresh e a sincronização
 * completa continuam sendo as fontes autorizadas para confirmar credenciais.
 */
function atualizarStatusConexaoStravaPorEdicao_(e) {
  if (!e || !e.range) return { ok: false, motivo: 'Evento ausente' };
  var sh = e.range.getSheet();
  var linha = e.range.getRow();
  if (sh.getName() !== H.SHEETS.CADASTRO || linha < 4) {
    return { ok: false, motivo: 'Fora do CADASTRO' };
  }

  var largura = Math.max(H.CAD.VERIF_STRAVA, sh.getLastColumn());
  var dados = sh.getRange(linha, 1, 1, largura).getValues()[0];
  var id = String(dados[H.CAD.ID - 1] || '').trim().toUpperCase();
  if (!_isAthIdValido_(id)) return { ok: false, motivo: 'ATH_ID inválido' };

  var uso = _normalizarDeclaracaoStrava_(dados[H.CAD.STRAVA_OK - 1]);
  var anterior = String(dados[H.CAD.CONEXAO_STRAVA - 1] || '').trim();
  var status;

  if (uso === 'nao' || uso === 'nao utiliza') {
    status = 'Não utiliza';
  } else if (uso === 'reconectar') {
    status = 'Reconectar';
  } else if (anterior === 'Conectado') {
    status = 'Conectado';
  } else {
    status = 'Aguardando conexão';
  }

  var agora = new Date();
  sh.getRange(linha, H.CAD.CONEXAO_STRAVA, 1, 2)
    .setValues([[status, agora]]);
  sh.getRange(linha, H.CAD.VERIF_STRAVA)
    .setNumberFormat('dd/MM/yyyy HH:mm:ss');
  return { ok: true, athId: id, status: status, verificadoEm: agora };
}

function atualizarStatusConexaoStravaAtleta(athId, opcoes) {
  opcoes = opcoes || {};
  var id = String(athId || '').trim().toUpperCase();
  if (!_isAthIdValido_(id)) return { ok: false, motivo: 'ATH_ID inválido' };

  var ss = SpreadsheetApp.openById(_getSsId());
  var sh = ss.getSheetByName(H.SHEETS.CADASTRO);
  if (!sh) return { ok: false, motivo: 'CADASTRO não encontrado' };

  var lastRow = sh.getLastRow();
  if (lastRow < 4) return { ok: false, motivo: 'CADASTRO vazio' };
  var largura = Math.max(H.CAD.VERIF_STRAVA, sh.getLastColumn());
  var dados = sh.getRange(4, 1, lastRow - 3, largura).getValues();

  for (var i = 0; i < dados.length; i++) {
    if (String(dados[i][H.CAD.ID - 1] || '').trim().toUpperCase() !== id) continue;
    var status = _calcularStatusConexaoStrava_(
      id,
      dados[i][H.CAD.STRAVA_OK - 1],
      opcoes
    );
    var agora = new Date();
    sh.getRange(i + 4, H.CAD.CONEXAO_STRAVA, 1, 2)
      .setValues([[status, agora]]);
    sh.getRange(i + 4, H.CAD.VERIF_STRAVA).setNumberFormat('dd/MM/yyyy HH:mm:ss');

    if (opcoes.atualizarFila && typeof registrarFilaWhatsAppCadastro_ === 'function') {
      try { registrarFilaWhatsAppCadastro_(_dadosFilaCadastroLinha_(dados[i])); } catch (_) {}
    }
    return { ok: true, athId: id, status: status, verificadoEm: agora };
  }
  return { ok: false, motivo: 'Atleta não encontrado', athId: id };
}

function sincronizarStatusConexaoStravaTodos(opcoes) {
  opcoes = opcoes || {};
  var ss = SpreadsheetApp.openById(_getSsId());
  var sh = ss.getSheetByName(H.SHEETS.CADASTRO);
  if (!sh || sh.getLastRow() < 4) return { atualizados: 0 };

  var lastRow = sh.getLastRow();
  var largura = Math.max(H.CAD.VERIF_STRAVA, sh.getLastColumn());
  var dados = sh.getRange(4, 1, lastRow - 3, largura).getValues();
  var agora = new Date();
  var saida = [];
  var atualizados = 0;
  var contagem = {
    'Conectado': 0,
    'Aguardando conexão': 0,
    'Não utiliza': 0,
    'Reconectar': 0
  };

  for (var i = 0; i < dados.length; i++) {
    var id = String(dados[i][H.CAD.ID - 1] || '').trim().toUpperCase();
    if (!_isAthIdValido_(id)) {
      saida.push(['', '']);
      continue;
    }
    var status = _calcularStatusConexaoStrava_(
      id,
      dados[i][H.CAD.STRAVA_OK - 1],
      {}
    );
    saida.push([status, agora]);
    contagem[status] = (contagem[status] || 0) + 1;
    atualizados++;

    if (opcoes.atualizarFila && typeof registrarFilaWhatsAppCadastro_ === 'function') {
      try { registrarFilaWhatsAppCadastro_(_dadosFilaCadastroLinha_(dados[i])); } catch (_) {}
    }
  }

  sh.getRange(4, H.CAD.CONEXAO_STRAVA, saida.length, 2).setValues(saida);
  sh.getRange(4, H.CAD.VERIF_STRAVA, saida.length, 1)
    .setNumberFormat('dd/MM/yyyy HH:mm:ss');
  SpreadsheetApp.flush();

  try {
    _log('SISTEMA', 'INFO', 'sincronizarStatusConexaoStravaTodos',
      atualizados + ' atleta(s) atualizados', JSON.stringify(contagem));
  } catch (_) {}
  return { atualizados: atualizados, contagem: contagem, verificadoEm: agora };
}
