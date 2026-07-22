/**
 * HIPERATIVO V3 — fila operacional de mensagens do cadastro para WhatsApp.
 * Não envia mensagens automaticamente e nunca cria OAuth para quem já possui
 * token. O treinador copia a mensagem ou abre o link wa.me correspondente.
 */

const FILA_WHATSAPP_HEADERS_ = [
  'Data cadastro', 'ID Atleta', 'Nome', 'WhatsApp', 'E-mail',
  'Situação Strava', 'Link seguro de conexão', 'Mensagem pronta',
  'Abrir WhatsApp', 'Controle'
];

function _abaFilaWhatsApp_() {
  const ss = SpreadsheetApp.openById(_getSsId());
  const nomeAba = (H.SHEETS && H.SHEETS.WHATSAPP_STRAVA) || '📲 WHATSAPP STRAVA';
  let sh = ss.getSheetByName(nomeAba);
  if (!sh) sh = ss.insertSheet(nomeAba);

  const cabAtual = sh.getRange(1, 1, 1, FILA_WHATSAPP_HEADERS_.length).getValues()[0];
  if (cabAtual.join('|') !== FILA_WHATSAPP_HEADERS_.join('|')) {
    sh.getRange(1, 1, 1, FILA_WHATSAPP_HEADERS_.length).setValues([FILA_WHATSAPP_HEADERS_]);
  }
  sh.setFrozenRows(1);
  sh.getRange(1, 1, 1, FILA_WHATSAPP_HEADERS_.length)
    .setBackground('#001F3F').setFontColor('#FFFFFF').setFontWeight('bold')
    .setHorizontalAlignment('center').setWrap(true);
  sh.setRowHeight(1, 34);
  [120, 120, 210, 145, 220, 130, 320, 520, 320, 110]
    .forEach(function(largura, i) { sh.setColumnWidth(i + 1, largura); });
  return sh;
}

function _statusFilaWhatsApp_(stravaOk, stravaId) {
  const valor = String(stravaOk || '').trim().toLowerCase()
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '');
  if (valor === 'nao' || valor === 'nao utiliza') return 'Não utiliza';
  if (stravaId || valor === 'sim' || valor === 'conectado' || valor === 'ativo') return 'Conectado';
  return 'Pendente';
}

function _linkSeguroFilaWhatsApp_(athId, email) {
  const base = PropertiesService.getScriptProperties().getProperty('WEBAPP_URL') || '';
  if (!base) return '';
  // O e-mail não vai na URL: o atleta informa no formulário e o servidor
  // valida contra o cadastro antes de iniciar qualquer OAuth.
  return base + '?conectar=true&athId=' + encodeURIComponent(athId);
}

function _mensagemFilaWhatsApp_(dados, situacao, link) {
  const primeiroNome = String(dados.nome || 'Atleta').trim().split(/\s+/)[0];
  if (situacao === 'Não utiliza') {
    return 'Olá, ' + primeiroNome + '! 👋\n\nSeu cadastro no *Grupo Hiperativo* foi concluído. ✅\n\n' +
      'Registramos que você não utiliza Strava, então nenhuma conexão é necessária.\n\n' +
      'Seu código de atleta é *' + dados.athId + '*. Vamos juntos! ⚡';
  }
  if (situacao === 'Conectado') {
    return 'Olá, ' + primeiroNome + '! 👋\n\nSeu cadastro e sua conexão com o Strava estão confirmados no *Grupo Hiperativo*. ✅\n\n' +
      'Seu código de atleta é *' + dados.athId + '*. Não é necessário reconectar.';
  }
  return 'Olá, ' + primeiroNome + '! 👋\n\nSeu cadastro no *Grupo Hiperativo* foi recebido com sucesso. ✅\n\n' +
    'Seu código de atleta é *' + dados.athId + '*.\n\n' +
    'Para que seus treinos sejam importados automaticamente, conecte sua conta Strava pelo link seguro abaixo:\n' +
    link + '\n\nUse o mesmo e-mail informado no cadastro: ' + (dados.email || '') + '.\n\n' +
    'Se já estiver conectado, o sistema preservará sua conexão e não pedirá nova autorização. 🏃⚡';
}

function registrarFilaWhatsAppCadastro_(dados) {
  if (!dados || !_isAthIdValido_(dados.athId)) return false;
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return false;
  try {
    const sh = _abaFilaWhatsApp_();
    const lastRow = sh.getLastRow();
    const existentes = lastRow > 1 ? sh.getRange(2, 1, lastRow - 1, 10).getValues() : [];
    let linha = -1;
    for (let i = 0; i < existentes.length; i++) {
      if (String(existentes[i][1] || '').trim().toUpperCase() === dados.athId) {
        linha = i + 2;
        break;
      }
    }

    const situacao = _statusFilaWhatsApp_(dados.stravaOk, dados.stravaId);
    const link = situacao === 'Pendente' ? _linkSeguroFilaWhatsApp_(dados.athId, dados.email) : '';
    const mensagem = _mensagemFilaWhatsApp_(dados, situacao, link);
    const numero = String(dados.whats || '').replace(/\D/g, '');
    const abrirWhats = numero ? 'https://wa.me/' + numero + '?text=' + encodeURIComponent(mensagem) : '';
    const controleAnterior = linha > 0 ? String(sh.getRange(linha, 10).getValue() || '').trim() : '';
    let controle = situacao === 'Pendente' ? 'Enviar' : situacao;
    if (controleAnterior === 'Enviado') controle = 'Enviado';
    const dataCadastro = dados.dataCadastro || new Date();
    const valores = [[
      dataCadastro, dados.athId, dados.nome || '', dados.whats || '', dados.email || '',
      situacao, link, mensagem, abrirWhats, controle
    ]];

    if (linha > 0) sh.getRange(linha, 1, 1, 10).setValues(valores);
    else {
      linha = sh.getLastRow() + 1;
      sh.getRange(linha, 1, 1, 10).setValues(valores);
    }
    sh.getRange(linha, 1).setNumberFormat('dd/MM/yyyy HH:mm');
    sh.getRange(linha, 1, 1, 10).setVerticalAlignment('middle').setWrap(true);
    sh.setRowHeight(linha, 78);
    const cor = situacao === 'Pendente' ? '#FFF4CC' :
      (situacao === 'Conectado' ? '#D6F5EC' : '#EEEEEE');
    sh.getRange(linha, 6).setBackground(cor).setFontWeight('bold');
    return true;
  } finally {
    lock.releaseLock();
  }
}

function sincronizarFilaWhatsAppCadastros() {
  const ss = SpreadsheetApp.openById(_getSsId());
  const shCad = ss.getSheetByName(H.SHEETS.CADASTRO);
  if (!shCad) throw new Error('Aba CADASTRO não encontrada.');
  const dados = shCad.getDataRange().getValues();
  let atualizados = 0;
  for (let i = 3; i < dados.length; i++) {
    const athId = String(dados[i][H.CAD.ID - 1] || '').trim().toUpperCase();
    if (!_isAthIdValido_(athId)) continue;
    if (registrarFilaWhatsAppCadastro_({
      athId: athId,
      nome: String(dados[i][H.CAD.NOME - 1] || '').trim(),
      email: String(dados[i][H.CAD.EMAIL - 1] || '').trim(),
      whats: String(dados[i][H.CAD.WHATS - 1] || '').trim(),
      dataCadastro: dados[i][H.CAD.DATA_CAD - 1] || '',
      stravaOk: dados[i][H.CAD.STRAVA_OK - 1],
      stravaId: dados[i][H.CAD.STRAVA_ID - 1]
    })) atualizados++;
  }
  _log('SISTEMA', 'INFO', 'sincronizarFilaWhatsAppCadastros',
    atualizados + ' cadastro(s) revisados na fila WhatsApp', '');
  return atualizados;
}

function instalarAcionadorFilaWhatsApp() {
  ScriptApp.getProjectTriggers()
    .filter(function(t) { return t.getHandlerFunction() === 'sincronizarFilaWhatsAppCadastros'; })
    .forEach(function(t) { ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('sincronizarFilaWhatsAppCadastros')
    .timeBased().everyMinutes(15).create();
  sincronizarFilaWhatsAppCadastros();
  Logger.log('Fila WhatsApp instalada: atualização a cada 15 minutos e no ciclo principal.');
  return true;
}

function abrirFilaWhatsAppStrava() {
  const sh = _abaFilaWhatsApp_();
  sh.activate();
  SpreadsheetApp.getActive().setActiveSheet(sh);
}
