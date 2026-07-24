/**
 * Complemento clínico SHE -> HIPERATIVO V3.
 *
 * Regras de segurança:
 * - só processa respostas com consentimento de dados e saúde;
 * - completa apenas células vazias;
 * - nunca substitui informação já existente no V3;
 * - a auditoria registra somente nomes de campos, nunca conteúdo clínico;
 * - não toca em tokens, Strava ou atividades.
 */

function sincronizarSaudeSHEHistorico() {
  var lock = LockService.getDocumentLock();
  if (!lock.tryLock(30000)) return { ok: false, motivo: 'PROCESSAMENTO_EM_ANDAMENTO' };

  try {
    var origem = SpreadsheetApp.openById(SHECRM_CFG_.spreadsheetId);
    var respostas = _sheMapaDeclaracoesStrava_(origem);
    var ss = SpreadsheetApp.openById(_getSsId());
    var cad = ss.getSheetByName(H.SHEETS.CADASTRO);
    if (!cad || cad.getLastRow() < 4) {
      return { ok: true, vinculados: 0, atletasAtualizados: 0, camposPreenchidos: 0 };
    }

    var largura = Math.max(49, cad.getLastColumn());
    var dados = cad.getRange(4, 1, cad.getLastRow() - 3, largura).getValues();
    var atletasAtualizados = 0;
    var camposPreenchidos = 0;
    var vinculados = 0;

    dados.forEach(function(row, indice) {
      var email = _sheNormalizarEmail_(
        row[_sheSaudeColuna_('EMAIL_NORM', 42) - 1] ||
        row[_sheSaudeColuna_('EMAIL', 3) - 1]
      );
      var resposta = respostas[email];
      if (!email || !resposta) return;
      vinculados++;

      var resultado = _sheEnriquecerCadastroSaudeSHE_(row, resposta);
      if (!resultado.campos.length) return;

      resultado.campos.forEach(function(campo) {
        var coluna = _sheSaudeColunaDoCampo_(campo);
        cad.getRange(indice + 4, coluna).setValue(row[coluna - 1]);
      });
      _sheRegistrarAuditoriaSaudeV3_(
        ss,
        String(row[_sheSaudeColuna_('ID', 1) - 1] || '').trim(),
        resultado.campos,
        resultado.preservados
      );
      atletasAtualizados++;
      camposPreenchidos += resultado.campos.length;
    });

    SpreadsheetApp.flush();
    return {
      ok: true,
      vinculados: vinculados,
      atletasAtualizados: atletasAtualizados,
      camposPreenchidos: camposPreenchidos,
      tokensAlterados: false
    };
  } finally {
    lock.releaseLock();
  }
}

function _sheEnriquecerCadastroSaudeSHE_(cadastro, origem) {
  var resultado = { campos: [], preservados: [], ignorado: '' };
  if (!origem || !origem.resposta || !origem.cabecalho) {
    resultado.ignorado = 'RESPOSTA_COMPLETA_INDISPONIVEL';
    return resultado;
  }

  var consentimento = _sheSaudeResposta_(origem,
    'Você autoriza o tratamento dos seus dados pessoais e dos dados de saúde informados para cadastro, triagem, prescrição e acompanhamento esportivo?');
  if (!_sheSaudeEhSim_(consentimento)) {
    resultado.ignorado = 'SEM_CONSENTIMENTO';
    return resultado;
  }

  var respostasParQ = [
    ['Cardiovascular', 'Algum médico ou profissional de saúde já disse que você tem problema cardíaco ou cardiovascular que exige cuidado para se exercitar?'],
    ['Dor no peito', 'Você sente dor, pressão ou desconforto no peito durante esforço ou em repouso?'],
    ['Tontura/desmaio', 'Você já desmaiou ou teve tontura importante, perda de equilíbrio ou sensação de desmaio durante atividade física?'],
    ['Falta de ar/palpitação', 'Você apresenta falta de ar incomum, palpitações importantes ou cansaço desproporcional ao esforço?'],
    ['Condição clínica', 'Você tem hipertensão, diabetes, doença renal ou outra condição clínica que precisa ser considerada no treino?'],
    ['Respiratório', 'Você tem asma, doença pulmonar ou outro problema respiratório relevante para o exercício?'],
    ['Neurológico', 'Você tem alguma condição neurológica, convulsão ou alteração que possa afetar equilíbrio, coordenação ou segurança durante o exercício?'],
    ['Dor/lesão atual', 'Você sente atualmente dor, lesão ou limitação em ossos, articulações, músculos, tendões ou coluna que possa piorar com exercício?'],
    ['Cirurgia/internação', 'Você passou por cirurgia, internação ou evento de saúde importante recentemente?'],
    ['Gestação/pós-parto', 'Você está grávida, no pós-parto recente ou recebeu orientação específica sobre exercício nesse período?'],
    ['Medicamento relevante', 'Você usa medicamento que possa alterar frequência cardíaca, pressão, glicemia, atenção, equilíbrio ou tolerância ao esforço?'],
    ['Histórico familiar', 'Algum parente de primeiro grau teve morte súbita, infarto ou doença cardíaca precoce antes dos 50 anos?'],
    ['Alergia grave', 'Você tem alergia grave ou condição que possa exigir medicamento ou dispositivo de resgate durante o treino?'],
    ['Revisão necessária', 'Você respondeu SIM, NÃO SEI ou PREFIRO CONVERSAR EM PARTICULAR a alguma pergunta desta triagem?']
  ];

  var triagem = [];
  var positivos = [];
  respostasParQ.forEach(function(item) {
    var resposta = _sheSaudeResposta_(origem, item[1]);
    if (!_sheSaudeTemValor_(resposta)) return;
    triagem.push(item[0] + ': ' + String(resposta).trim());
    if (_sheSaudeExigeAtencao_(resposta)) positivos.push(item[0]);
  });

  var adaptacao = _sheSaudeResposta_(origem,
    'Existe alguma necessidade de acessibilidade, comunicação ou adaptação que ajude você a participar com segurança e autonomia?');
  var detalhesSaude = _sheSaudeResposta_(origem,
    'Descreva os alertas marcados, sintomas, diagnósticos, lesões, cirurgias ou limitações relevantes. Se preferir, escreva “conversar em particular”.');
  var outraSaude = _sheSaudeResposta_(origem,
    'Existe alguma outra informação física, de saúde ou de rotina que o treinador precisa conhecer?');
  var recomendacoes = _sheSaudeResposta_(origem,
    'Há alguma recomendação profissional sobre intensidade, movimentos que devem ser evitados ou sinais que exigem interrupção?');
  var recursoResgate = _sheSaudeResposta_(origem,
    'Você carrega inalador, glicose, adrenalina, medicação ou outro recurso de emergência durante o treino? Qual?');

  var condicoes = _sheSaudeJuntar_([
    _sheSaudeRotular_('Informações adicionais', outraSaude),
    _sheSaudeRotular_('Adaptação/segurança', adaptacao),
    _sheSaudeRotular_('Detalhes informados', detalhesSaude),
    _sheSaudeRotular_('Recomendação profissional', recomendacoes),
    _sheSaudeRotular_('Recurso de emergência', recursoResgate),
    positivos.length ? 'Triagem com atenção: ' + positivos.join(', ') : ''
  ]);

  var regioes = _sheSaudeResposta_(origem,
    'Quais regiões apresentam dor, lesão ou limitação atualmente?');
  var lesaoAtual = _sheSaudeResposta_(origem,
    'Você sente atualmente dor, lesão ou limitação em ossos, articulações, músculos, tendões ou coluna que possa piorar com exercício?');
  var lesoes = _sheSaudeJuntar_([
    _sheSaudeRotular_('Regiões', regioes),
    _sheSaudeEhSim_(lesaoAtual) && !_sheSaudeTemValor_(regioes) && !_sheSaudeTemValor_(detalhesSaude)
      ? 'Triagem SHE: informou dor, lesão ou limitação atual; revisar a resposta original.'
      : '',
    _sheSaudeTemValor_(detalhesSaude) && _sheSaudeEhSim_(lesaoAtual)
      ? _sheSaudeRotular_('Detalhes', detalhesSaude)
      : ''
  ]);

  var usaMedicamento = _sheSaudeResposta_(origem,
    'Você usa medicamento que possa alterar frequência cardíaca, pressão, glicemia, atenção, equilíbrio ou tolerância ao esforço?');
  var medicamentos = _sheSaudeResposta_(origem,
    'Informe apenas medicamentos relevantes para exercício ou segurança, como os que alteram frequência cardíaca, pressão, glicemia, atenção ou equilíbrio. Se preferir, escreva “conversar em particular”.');
  if (_sheSaudeEhSim_(usaMedicamento) && !_sheSaudeTemValor_(medicamentos)) {
    medicamentos = 'Triagem SHE: informou uso de medicamento relevante; detalhes não informados. Revisão necessária.';
  }

  var nascimento = _sheSaudeResposta_(origem, 'Data de nascimento');
  var instagram = _sheSaudeResposta_(origem, 'Instagram ou outra rede social');
  var emergenciaNome = _sheSaudeResposta_(origem, 'Nome do contato de emergência');
  var emergenciaVinculo = _sheSaudeVinculo_(_sheSaudeResposta_(origem,
    'Parentesco ou vínculo do contato de emergência'));
  var emergenciaTelefone = _sheSaudeTelefone_(_sheSaudeResposta_(origem,
    'Telefone do contato de emergência com DDD'));

  _sheSaudePreencher_(cadastro, _sheSaudeColuna_('NASC', 5), nascimento, 'NASCIMENTO', resultado);
  _sheSaudePreencherOuComplementarCondicoes_(
    cadastro,
    _sheSaudeColuna_('SAUDE', 14),
    condicoes,
    resultado
  );
  _sheSaudePreencher_(cadastro, _sheSaudeColuna_('LESAO', 15), lesoes, 'LESOES_LIMITACOES', resultado);
  _sheSaudePreencher_(cadastro, _sheSaudeColuna_('MED', 16), medicamentos, 'MEDICAMENTOS', resultado);
  _sheSaudePreencher_(cadastro, _sheSaudeColuna_('EMERG_NOME', 28), emergenciaNome, 'EMERGENCIA_NOME', resultado);
  _sheSaudePreencher_(cadastro, _sheSaudeColuna_('EMERG_TEL', 29), emergenciaTelefone, 'EMERGENCIA_TELEFONE', resultado);
  _sheSaudePreencher_(cadastro, _sheSaudeColuna_('EMERG_REL', 30), emergenciaVinculo, 'EMERGENCIA_VINCULO', resultado);
  _sheSaudePreencher_(cadastro, _sheSaudeColuna_('PAR_Q', 31), triagem.join(' | '), 'PAR_Q_TRIAGEM_SHE', resultado);
  _sheSaudePreencher_(cadastro, _sheSaudeColuna_('INSTAGRAM', 35), instagram, 'INSTAGRAM', resultado);

  return resultado;
}

function _sheRegistrarAuditoriaSaudeV3_(ss, athId, campos, preservados) {
  campos = campos || [];
  preservados = preservados || [];
  if (!campos.length) return;

  var sh = ss.getSheetByName('📋 AUDITORIA');
  if (!sh) return;
  var detalhe = 'Atleta ' + athId + '.';
  if (campos.length) detalhe += ' Preenchidos pelo SHE: ' + campos.join(', ') + '.';
  if (preservados.length) detalhe += ' Valores do V3 preservados: ' + preservados.join(', ') + '.';
  sh.appendRow([
    new Date(),
    '👤 CADASTRO',
    'SHE_SAUDE ' + athId,
    campos.length ? 'ATUALIZADO' : 'PRESERVADO',
    detalhe
  ]);
  sh.getRange(sh.getLastRow(), 1).setNumberFormat('dd/MM/yyyy HH:mm:ss');
}

function _sheSaudePreencher_(row, col, valor, campo, resultado) {
  if (!_sheSaudeTemValor_(valor)) return;
  var atual = row[col - 1];
  if (!_sheSaudeTemValor_(atual)) {
    row[col - 1] = valor;
    resultado.campos.push(campo);
    return;
  }
  if (_sheSaudeComparavel_(atual) !== _sheSaudeComparavel_(valor)) {
    resultado.preservados.push(campo);
  }
}

function _sheSaudePreencherOuComplementarCondicoes_(row, col, valor, resultado) {
  if (!_sheSaudeTemValor_(valor)) return;
  var atual = row[col - 1];
  if (!_sheSaudeTemValor_(atual)) {
    row[col - 1] = valor;
    resultado.campos.push('CONDICOES_SAUDE');
    return;
  }

  var atualNorm = _sheSaudeComparavel_(atual);
  var valorNorm = _sheSaudeComparavel_(valor);
  var generico =
    atualNorm.indexOf('alertas criticos') >= 0 ||
    atualNorm.indexOf('pontos para revisao') >= 0 ||
    atualNorm.indexOf('atencao - revisar no hiper crm') >= 0 ||
    atualNorm.indexOf('nenhum alerta declarado no she') >= 0;

  if (generico && atualNorm.indexOf(valorNorm) < 0 &&
      atualNorm.indexOf('complemento she:') < 0) {
    row[col - 1] = String(atual).trim() + ' | Complemento SHE: ' + String(valor).trim();
    resultado.campos.push('CONDICOES_SAUDE_COMPLEMENTO');
    return;
  }
  if (atualNorm !== valorNorm) resultado.preservados.push('CONDICOES_SAUDE');
}

function _sheSaudeColuna_(chave, fallback) {
  var configurada = H && H.CAD ? Number(H.CAD[chave]) : 0;
  return configurada > 0 ? configurada : fallback;
}

function _sheSaudeColunaDoCampo_(campo) {
  var colunas = {
    NASCIMENTO: _sheSaudeColuna_('NASC', 5),
    CONDICOES_SAUDE: _sheSaudeColuna_('SAUDE', 14),
    CONDICOES_SAUDE_COMPLEMENTO: _sheSaudeColuna_('SAUDE', 14),
    LESOES_LIMITACOES: _sheSaudeColuna_('LESAO', 15),
    MEDICAMENTOS: _sheSaudeColuna_('MED', 16),
    EMERGENCIA_NOME: _sheSaudeColuna_('EMERG_NOME', 28),
    EMERGENCIA_TELEFONE: _sheSaudeColuna_('EMERG_TEL', 29),
    EMERGENCIA_VINCULO: _sheSaudeColuna_('EMERG_REL', 30),
    PAR_Q_TRIAGEM_SHE: _sheSaudeColuna_('PAR_Q', 31),
    INSTAGRAM: _sheSaudeColuna_('INSTAGRAM', 35)
  };
  return colunas[campo];
}

function _sheSaudeResposta_(origem, cabecalho) {
  return _sheValor_(origem.resposta, origem.cabecalho, cabecalho);
}

function _sheSaudeTemValor_(valor) {
  return valor !== '' && valor !== null && valor !== undefined &&
    String(valor).trim() !== '';
}

function _sheSaudeEhSim_(valor) {
  return /^sim(?:\b|,)/i.test(String(valor || '').trim());
}

function _sheSaudeExigeAtencao_(valor) {
  var texto = String(valor || '').trim().toLowerCase()
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '');
  return /^sim(?:\b|,)/.test(texto) ||
    texto.indexOf('nao sei') === 0 ||
    texto.indexOf('prefiro') === 0;
}

function _sheSaudeRotular_(rotulo, valor) {
  return _sheSaudeTemValor_(valor) ? rotulo + ': ' + String(valor).trim() : '';
}

function _sheSaudeJuntar_(partes) {
  var vistos = {};
  return (partes || []).filter(function(parte) {
    var valor = String(parte || '').trim();
    if (!valor) return false;
    var chave = _sheSaudeComparavel_(valor);
    if (vistos[chave]) return false;
    vistos[chave] = true;
    return true;
  }).join(' | ');
}

function _sheSaudeTelefone_(valor) {
  var numero = String(valor || '').replace(/\D/g, '');
  if (numero.length > 11) numero = numero.slice(-11);
  return numero;
}

function _sheSaudeVinculo_(valor) {
  var original = String(valor || '').trim();
  var texto = original.toLowerCase()
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '');
  if (!texto) return '';
  if (texto.indexOf('conjuge') >= 0 || texto.indexOf('companheir') >= 0 ||
      texto.indexOf('espos') >= 0 || texto.indexOf('marido') >= 0) {
    return 'Cônjuge / Companheiro(a)';
  }
  if (texto.indexOf('pai') >= 0 || texto.indexOf('mae') >= 0) return 'Pai / Mãe';
  if (texto.indexOf('filh') >= 0) return 'Filho(a)';
  if (texto.indexOf('irma') >= 0) return 'Irmão / Irmã';
  if (texto.indexOf('amig') >= 0) return 'Amigo(a)';
  return 'Outro';
}

function _sheSaudeComparavel_(valor) {
  if (valor instanceof Date && !isNaN(valor.getTime())) return String(valor.getTime());
  return String(valor || '').trim().toLowerCase()
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .replace(/\s+/g, ' ');
}
