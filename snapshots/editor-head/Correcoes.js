/**
 * HIPERATIVO V3 - Correcoes.gs
 * Correcoes de formatacao visual da planilha
 * Nao altera dados de atletas, tokens ou refresh tokens
 */

function corrigirFormatacaoCompleta() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  // const ui = SpreadsheetApp.getUi();
  try {
    corrigirAbaMetricas(ss);
    corrigirAbaTokens(ss);
    corrigirAbaRankingCompleto(ss);
    corrigirAbaRanking(ss);
    corrigirAbaErros(ss);
    corrigirAbaConfig(ss);
    deletarPaginasVazias(ss);
    Logger.log('Correcoes concluidas!');
  } catch(e) {
    Logger.log('Erro: ' + e.message);
  }
}

function corrigirAbaMetricas(ss) {
  const ws = ss.getSheetByName('📈 MÉTRICAS');
  if (!ws) return;
  const dados = ws.getDataRange().getValues();
  const linhasValidas = [];
  const idsVistos = new Set();
  for (let i = 0; i < dados.length; i++) {
    const p = String(dados[i][0] || '').trim();
    if (i === 0) { linhasValidas.push(dados[i]); continue; }
    if (p === 'ID Atleta' || p.includes('Nome Completo') || p === '') continue;
    if (p.startsWith('ATH')) {
      if (idsVistos.has(p)) continue;
      idsVistos.add(p);
    }
    linhasValidas.push(dados[i]);
  }
  const maxCols = Math.max(dados[0] ? dados[0].length : 10, ws.getLastColumn());
  ws.clearContents();
  if (linhasValidas.length > 0) ws.getRange(1, 1, linhasValidas.length, linhasValidas[0].length).setValues(linhasValidas);
  ws.insertRowBefore(1);
  ws.getRange(1, 1, 1, maxCols).merge()
    .setValue('📈 MÉTRICAS DOS ATLETAS — HIPERATIVO V3')
    .setBackground('#001F3F').setFontColor('#FFFFFF')
    .setFontSize(13).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  ws.setRowHeight(1, 35);
  ws.getRange(2, 1, 1, maxCols)
    .setBackground('#003D7A').setFontColor('#FFFFFF')
    .setFontWeight('bold').setHorizontalAlignment('center');
  ws.setRowHeight(2, 28);
  const ul = ws.getLastRow();
  for (let r = 3; r <= ul; r++) {
    ws.getRange(r, 1, 1, maxCols).setBackground(r % 2 === 0 ? '#F0F4F8' : '#FFFFFF').setFontColor('#000000');
  }
  if (ul >= 2) ws.getRange(2, 1, ul - 1, maxCols).setBorder(true,true,true,true,true,true,'#CCCCCC', SpreadsheetApp.BorderStyle.SOLID);
  ws.setFrozenRows(2);
}

function corrigirAbaTokens(ss) {
  const ws = ss.getSheetByName('🔐 TOKENS');
  if (!ws) return;
  const dados = ws.getDataRange().getValues();
  const linhasValidas = [];
  for (let i = 0; i < dados.length; i++) {
    const p = String(dados[i][0] || '').trim();
    if (i === 0) { linhasValidas.push(dados[i]); continue; }
    if (p === 'TOK_ID Atleta' || p.includes('ID Atleta') || p.includes('IDENT') || p === '') continue;
    if (p.startsWith('TOK_')) linhasValidas.push(dados[i]);
  }
  const maxCols = Math.max(dados[0] ? dados[0].length : 10, ws.getLastColumn());
  ws.clearContents();
  if (linhasValidas.length > 0) ws.getRange(1, 1, linhasValidas.length, linhasValidas[0].length).setValues(linhasValidas);
  ws.insertRowBefore(1);
  ws.getRange(1, 1, 1, maxCols).merge()
    .setValue('🔐 TOKENS DE ACESSO STRAVA — HIPERATIVO V3')
    .setBackground('#0A0A0A').setFontColor('#FFD700')
    .setFontSize(13).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  ws.setRowHeight(1, 35);
  ws.getRange(2, 1, 1, maxCols)
    .setBackground('#2C2C2C').setFontColor('#FFFFFF')
    .setFontWeight('bold').setHorizontalAlignment('center');
  ws.setRowHeight(2, 28);
  const ul = ws.getLastRow();
  for (let r = 3; r <= ul; r++) {
    ws.getRange(r, 1, 1, maxCols).setBackground(r % 2 === 0 ? '#F5F5F5' : '#FFFFFF').setFontColor('#000000');
  }
  if (ul >= 2) ws.getRange(2, 1, ul - 1, maxCols).setBorder(true,true,true,true,true,true,'#CCCCCC', SpreadsheetApp.BorderStyle.SOLID);
  ws.setFrozenRows(2);
}

function corrigirAbaRankingCompleto(ss) {
  const ws = ss.getSheetByName('🏆 RANKING COMPLETO');
  if (!ws) return;
  const ul = ws.getLastRow();
  const uc = ws.getLastColumn();
  if (ul === 0 || uc === 0) return;
  const range = ws.getRange(1, 1, ul, uc);
  const bgs = range.getBackgrounds();
  const vals = range.getValues();
  const newColors = bgs.map((row, r) => row.map((bg, c) => {
    if (!vals[r][c] && String(vals[r][c]) === '') return '#000000';
    const hex = (bg || '').replace('#','');
    if (hex.length !== 6) return '#000000';
    const lum = (parseInt(hex.substr(0,2),16)*299 + parseInt(hex.substr(2,2),16)*587 + parseInt(hex.substr(4,2),16)*114) / 1000;
    return lum < 128 ? '#FFFFFF' : '#000000';
  }));
  range.setFontColors(newColors).setFontSize(11);
}

function corrigirAbaRanking(ss) {
  const ws = ss.getSheetByName('🏆 RANKING');
  if (!ws) return;
  const ul = ws.getLastRow();
  const uc = ws.getLastColumn();
  if (ul === 0) return;
  const vals = ws.getRange(1, 1, ul, uc).getValues();
  let ultComDados = 0;
  for (let r = 0; r < vals.length; r++) {
    if (vals[r].some(v => String(v || '').trim() !== '')) ultComDados = r + 1;
  }
  if (ul > ultComDados) {
    ws.getRange(ultComDados + 1, 1, ul - ultComDados, uc)
      .setBackground(null).setFontColor('#000000')
      .setBorder(false,false,false,false,false,false)
      .clearContent();
  }
}

function corrigirAbaErros(ss) {
  const ws = ss.getSheetByName('🔴 ERROS');
  if (!ws) return;
  const bg1 = ws.getRange('A1').getBackground();
  if (bg1 !== '#ffffff' && bg1 !== 'white' && bg1 !== null) return;
  ws.insertRowBefore(1);
  const nc = ws.getLastColumn() || 6;
  ws.getRange(1, 1, 1, nc).merge()
    .setValue('🔴 LOG DE ERROS E EVENTOS — HIPERATIVO V3')
    .setBackground('#0A0A0A').setFontColor('#FF4444')
    .setFontSize(13).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  ws.setRowHeight(1, 35);
  ws.getRange(2, 1, 1, nc).setBackground('#2C2C2C').setFontColor('#FFFFFF').setFontWeight('bold').setHorizontalAlignment('center');
  ws.setRowHeight(2, 28);
  ws.setFrozenRows(2);
}

function corrigirAbaConfig(ss) {
  const ws = ss.getSheetByName('⚙️ CONFIG');
  if (!ws) return;
  const bg1 = ws.getRange('A1').getBackground();
  if (bg1 !== '#ffffff' && bg1 !== 'white' && bg1 !== null) return;
  const v1 = String(ws.getRange('A1').getValue()).trim();
  if (v1 !== 'CHAVE') return;
  ws.insertRowBefore(1);
  const nc = ws.getLastColumn() || 4;
  ws.getRange(1, 1, 1, nc).merge()
    .setValue('⚙️ CONFIGURAÇÕES DO SISTEMA — HIPERATIVO V3')
    .setBackground('#2C2C2C').setFontColor('#FFFFFF')
    .setFontSize(13).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  ws.setRowHeight(1, 35);
  ws.getRange(2, 1, 1, nc).setBackground('#4A4A4A').setFontColor('#FFFFFF').setFontWeight('bold');
  ws.setRowHeight(2, 28);
  ws.setFrozenRows(2);
}

function deletarPaginasVazias(ss) {
  const sheets = ss.getSheets();
  for (const sh of sheets) {
    const nome = sh.getName();
    if (!/^Páginad+$/.test(nome)) continue;
    const ul = sh.getLastRow();
    const uc = sh.getLastColumn();
    if (ul <= 3 && uc <= 5) {
      const texto = sh.getDataRange().getValues().flat().join('').trim();
      if (texto.length < 100) {
        ss.deleteSheet(sh);
        Utilities.sleep(300);
      }
    }
  }
}


function limparCabecalhosDuplicados() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Limpar TOKENS - remover linhas onde col A contem "TOK_🧑" ou "TOK_📊"
  const wsT = ss.getSheetByName('🔐 TOKENS');
  if (wsT) {
    const dT = wsT.getDataRange().getValues();
    for (let i = dT.length - 1; i >= 0; i--) {
      const v = String(dT[i][0] || '');
      if (v.includes('IDENT') || v.includes('ID Atleta') || 
          (v.startsWith('TOK_') && !v.startsWith('TOK_ATH') && !v.startsWith('TOK_FC') && 
           !v.startsWith('TOK_C2') && !v.startsWith('TOK_12') && !v.startsWith('TOK_A') && 
           !v.startsWith('TOK_6') && !v.startsWith('TOK_ID Atleta'))) {
        // Check se e cabecalho repetido (nao tem ID de atleta real)
        const v2 = String(dT[i][1] || '').trim();
        const v3 = String(dT[i][2] || '').trim();
        if (v2 === 'ID Atleta' || v2 === '' || v3 === 'Nome Completo' || v3 === 'IDENTIFICAÇÃO' || v2.includes('ID Atleta')) {
          wsT.deleteRow(i + 1);
          Logger.log('TOKENS: deletada linha ' + (i+1));
        }
      }
    }
  }
  
  // Limpar METRICAS - remover linhas onde col A contem texto de cabecalho
  const wsM = ss.getSheetByName('📈 MÉTRICAS');
  if (wsM) {
    const dM = wsM.getDataRange().getValues();
    for (let i = dM.length - 1; i >= 0; i--) {
      const v = String(dM[i][0] || '').trim();
      const v2 = String(dM[i][1] || '').trim();
      if (i > 1 && (v === 'ID Atleta' || v.includes('IDENTIFICAÇÃO') || 
          (v.includes('ID Atleta') && v2.includes('Nome')))) {
        wsM.deleteRow(i + 1);
        Logger.log('METRICAS: deletada linha ' + (i+1));
      }
    }
  }
  
  Logger.log('Limpeza concluida!');
}


function resetAbaMetricas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ws = ss.getSheetByName('📈 MÉTRICAS');
  if (!ws) return;
  
  // Ler todos os dados
  const dados = ws.getDataRange().getValues();
  const linhasValidas = [];
  const idsVistos = new Set();
  let cabecalhoEncontrado = false;
  
  for (let i = 0; i < dados.length; i++) {
    const p = String(dados[i][0] || '').trim();
    const p2 = String(dados[i][1] || '').trim();
    
    // Pular linhas de titulo repetidas (contem 'METRICAS DOS ATLETAS')
    if (p.includes('MÉTRICAS DOS ATLETAS') || p.includes('METRICAS DOS ATLETAS')) continue;
    
    // Pular linhas de cabecalho repetidas (ID Atleta / Nome Completo)
    if (p === 'ID Atleta' && (p2 === 'Nome Completo' || p2 === '')) {
      if (cabecalhoEncontrado) continue; // pular duplicados
      cabecalhoEncontrado = true;
      linhasValidas.push(dados[i]);
      continue;
    }
    
    // Pular linhas vazias do topo
    if (p === '' && linhasValidas.length === 0) continue;
    
    // Pular linhas que sao apenas cabecalhos repetidos (icones + texto de cabecalho)
    if (p.includes('ID Atleta') && p2.includes('Nome')) {
      if (cabecalhoEncontrado) continue;
      cabecalhoEncontrado = true;
      linhasValidas.push(dados[i]);
      continue;
    }
    
    // Dados reais de atletas (ATH...)
    if (p.startsWith('ATH') || p.startsWith('ATH_')) {
      if (idsVistos.has(p)) continue;
      idsVistos.add(p);
      linhasValidas.push(dados[i]);
      continue;
    }
    
    // Outros dados
    if (p !== '' && !p.includes('IDENTIFICA')) {
      linhasValidas.push(dados[i]);
    }
  }
  
  // Limpar e reescrever
  const maxCols = ws.getLastColumn();
  ws.clearContents();
  ws.clearFormats();
  
  // Inserir titulo
  const numTitleCols = maxCols || 20;
  ws.getRange(1, 1, 1, numTitleCols).merge()
    .setValue('📈 MÉTRICAS DOS ATLETAS — HIPERATIVO V3')
    .setBackground('#001F3F').setFontColor('#FFFFFF')
    .setFontSize(13).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  ws.setRowHeight(1, 35);
  
  // Inserir dados limpos a partir da linha 2
  if (linhasValidas.length > 0) {
    const numCols = linhasValidas[0].length;
    ws.getRange(2, 1, linhasValidas.length, numCols).setValues(linhasValidas);
    
    // Formatar cabecalho (linha 2)
    ws.getRange(2, 1, 1, numTitleCols)
      .setBackground('#003D7A').setFontColor('#FFFFFF')
      .setFontWeight('bold').setHorizontalAlignment('center');
    ws.setRowHeight(2, 28);
    
    // Alternar cores
    const ul = ws.getLastRow();
    for (let r = 3; r <= ul; r++) {
      ws.getRange(r, 1, 1, numCols).setBackground(r % 2 === 0 ? '#F0F4F8' : '#FFFFFF').setFontColor('#000000');
    }
    
    // Bordas
    if (ul >= 2) ws.getRange(2, 1, ul - 1, numCols).setBorder(true,true,true,true,true,true,'#CCCCCC', SpreadsheetApp.BorderStyle.SOLID);
  }
  
  ws.setFrozenRows(2);
  Logger.log('METRICAS resetada. Linhas validas: ' + linhasValidas.length);
}


function adicionarCabecalhoMetricas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ws = ss.getSheetByName('📈 MÉTRICAS');
  if (!ws) return;
  
  // Verificar se linha 2 tem cabecalho (deve comecar com 'ID Atleta')
  const val2 = String(ws.getRange('A2').getValue() || '').trim();
  if (val2 === 'ID Atleta') {
    Logger.log('Cabecalho ja existe na linha 2');
    return;
  }
  
  // Inserir linha de cabecalho na linha 2
  ws.insertRowBefore(2);
  
  // Buscar o cabecalho da aba MetricasFallbackSetup ou Setup
  // Usar o cabecalho padrao baseado nas constantes H.MET do Config.gs
  const headerValues = [
    'ID Atleta', 'Nome Completo', 'Ultima Atualizacao',
    'VO2max', 'FC Max', 'Z1 max (bpm)', 'Z2 max (bpm)', 
    'Z3 max (bpm)', 'Z4 max (bpm)', 'Z5 max (bpm)',
    'Pace Ref (s/km)', 'Pace Z1 (s/km)', 'Pace Z2 (s/km)',
    'Total Km 30d', 'Total Treinos 30d', 'Ultimo Treino', 'Nivel'
  ];
  
  const numCols = ws.getLastColumn();
  
  // Se tiver mais colunas que o header, completar com vazio
  while (headerValues.length < numCols) headerValues.push('');
  
  ws.getRange(2, 1, 1, Math.min(headerValues.length, numCols))
    .setValues([headerValues.slice(0, numCols)])
    .setBackground('#003D7A')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  ws.setRowHeight(2, 28);
  ws.setFrozenRows(2);
  
  Logger.log('Cabecalho adicionado na linha 2 de METRICAS');
}


function limparLinhasLixoRanking() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ws = ss.getSheetByName('🏆 RANKING');
  if (!ws) return;
  
  // Limpar linhas 17-20 que tem formatacao suja
  const range = ws.getRange('A17:F25');
  range.setBackground(null);
  range.setFontColor('#000000');
  range.setBorder(false, false, false, false, false, false);
  range.clearContent();
  
  // Remover mesclagem se houver
  try { range.breakApart(); } catch(e) {}
  
  Logger.log('Linhas 17-25 do RANKING limpas');
}


function executarDeletarPaginasVazias() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  let deletadas = 0;
  
  for (const sh of sheets) {
    const nome = sh.getName();
    
    // Verificar se e uma aba PaginaN generica
    if (!nome.match(/^P[áa]ginad+$/)) continue;
    
    const ultimaLinha = sh.getLastRow();
    const ultimaCol = sh.getLastColumn();
    
    // Aba com pouco conteudo
    if (ultimaLinha <= 3 && ultimaCol <= 6) {
      const valores = sh.getDataRange().getValues();
      const textoTotal = valores.flat().join('').trim();
      
      // Se tiver apenas ranking mensal antigo sem dados de atletas, deletar
      if (textoTotal.length < 150) {
        Logger.log('Deletando aba: ' + nome + ' (chars: ' + textoTotal.length + ')');
        ss.deleteSheet(sh);
        deletadas++;
        Utilities.sleep(500);
      }
    }
  }
  
  Logger.log('Total deletadas: ' + deletadas);
  return deletadas;
}


function inspecionarPaginas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  const resultado = [];
  
  for (const sh of sheets) {
    const nome = sh.getName();
    if (!nome.match(/^P[áa]ginad+$/)) continue;
    
    const ul = sh.getLastRow();
    const uc = sh.getLastColumn();
    const vals = sh.getDataRange().getValues();
    const texto = vals.flat().join('').trim();
    
    resultado.push(nome + ': linhas=' + ul + ', cols=' + uc + ', chars=' + texto.length + ', preview=' + texto.substring(0, 50));
  }
  
  Logger.log(resultado.join('\n'));
  return resultado.length;
}
