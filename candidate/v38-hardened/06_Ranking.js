/**
 * ═══════════════════════════════════════════════════════════════════════
 * HIPERATIVO V3 — 06_Ranking.gs
 * Ranking mensal de atletas por modalidade e categoria
 * Top 10 por categoria — atualizado automaticamente ao importar Strava
 * ═══════════════════════════════════════════════════════════════════════
 */

// ============================================================
// RANKING MENSAL — Top 10 atletas por categoria
// ============================================================
const RANK_TOP = 10;

/**
 * Entry point: calcula e renderiza o ranking do mês corrente
 * na aba H.SHEETS.RANKING ('🏆 RANKING').
 */
function atualizarRanking() {
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const wsA  = ss.getSheetByName(H.SHEETS.ATIVIDADES);
  if (!wsA) { SpreadsheetApp.getUi().alert('Aba ATIVIDADES não encontrada.'); return; }

  const now    = new Date();
  const ano    = now.getFullYear();
  const mes    = now.getMonth();       // 0-indexed
  const dtIni  = new Date(ano, mes, 1);
  const dtFim  = new Date(ano, mes + 1, 0, 23, 59, 59);
  const mesLabel = Utilities.formatDate(dtIni, Session.getScriptTimeZone(), 'MMMM/yyyy');

  // Lê ATIVIDADES
  const lastRowA = wsA.getLastRow();
  if (lastRowA < 2) { _log('RANKING','WARN','atualizarRanking','Sem dados',''); return; }
  const dados = wsA.getRange(2, 1, lastRowA - 1, wsA.getLastColumn()).getValues();

  // Filtra mês corrente
  const H_DATA = H.ATIV.DATA - 1;   // 0-indexed
  const dadosMes = dados.filter(row => {
    const d = row[H_DATA];
    return d instanceof Date && d >= dtIni && d <= dtFim;
  });

  // Garante/cria aba RANKING
  let wsR = ss.getSheetByName(H.SHEETS.RANKING);
  if (!wsR) {
    wsR = ss.insertSheet(H.SHEETS.RANKING);
    ss.setActiveSheet(wsR);
    ss.moveActiveSheet(ss.getNumSheets());
  }
  wsR.clearContents();
  wsR.clearFormats();

  _renderizarRanking(wsR, dadosMes, mesLabel);
  _log('RANKING','INFO','atualizarRanking','Ranking atualizado: ' + mesLabel + ' (' + dadosMes.length + ' atividades)','');
}

/**
 * Renderiza todas as categorias em layout 2 colunas.
 */
function _renderizarRanking(ws, dados, mesLabel) {
  // Cabeçalho
  ws.getRange(1, 1, 1, 4).merge()
    .setValue('🏆 RANKING MENSAL — ' + mesLabel)
    .setBackground('#1a1a2e')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold')
    .setFontSize(14)
    .setHorizontalAlignment('center');
  ws.setRowHeight(1, 40);

  ws.getRange(2, 1, 1, 4).merge()
    .setValue('Atualizado em: ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm'))
    .setBackground('#16213e')
    .setFontColor('#AAAAAA')
    .setFontSize(10)
    .setHorizontalAlignment('center');

  ws.freezeRows(2);

  const cats = _getCategorias();
  let linha   = 3;
  let colBase = 1;

  for (let i = 0; i < cats.length; i++) {
    if (i > 0 && i % 2 === 0) {
      linha = 3;
      colBase = 3;
    }
    _renderSecao(ws, linha, colBase, cats[i], dados);
    linha += 2 + RANK_TOP + 2; // header(2) + top10 + espacamento(2)
  }

  // Ajusta larguras
  ws.setColumnWidth(1, 30);
  ws.setColumnWidth(2, 200);
  ws.setColumnWidth(3, 30);
  ws.setColumnWidth(4, 200);
}

/**
 * Renderiza uma seção de ranking (uma categoria).
 */
function _renderSecao(ws, linhaInicio, colInicio, cat, dados) {
  const cor    = cat.cor || '#333333';
  const corTxt = '#FFFFFF';

  // Título da categoria
  ws.getRange(linhaInicio, colInicio, 1, 2).merge()
    .setValue(cat.titulo)
    .setBackground(cor)
    .setFontColor(corTxt)
    .setFontWeight('bold')
    .setFontSize(11)
    .setHorizontalAlignment('center');
  ws.setRowHeight(linhaInicio, 28);

  // Sub-header
  ws.getRange(linhaInicio + 1, colInicio)
    .setValue('#')
    .setBackground(_darken(cor))
    .setFontColor(corTxt)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  ws.getRange(linhaInicio + 1, colInicio + 1)
    .setValue('Atleta — ' + cat.unidade)
    .setBackground(_darken(cor))
    .setFontColor(corTxt)
    .setFontWeight('bold');

  // Calcula ranking
  const resultado = cat.fn(dados);
  const top       = resultado.slice(0, RANK_TOP);

  const emojis = ['🥇','🥈','🥉','4️⃣','5️⃣','6️⃣','7️⃣','8️⃣','9️⃣','🔟'];

  for (let j = 0; j < RANK_TOP; j++) {
    const row = top[j];
    const lRow = linhaInicio + 2 + j;
    const bg   = j % 2 === 0 ? '#F8F9FA' : '#FFFFFF';
    ws.getRange(lRow, colInicio)
      .setValue(row ? emojis[j] : '')
      .setBackground(bg)
      .setHorizontalAlignment('center');
    ws.getRange(lRow, colInicio + 1)
      .setValue(row ? row.nome + ' — ' + row.valor : '')
      .setBackground(bg);
    ws.setRowHeight(lRow, 22);
  }
}

/** Escurece levemente uma cor hex. */
function _darken(hex) {
  try {
    const r = parseInt(hex.slice(1,3),16);
    const g = parseInt(hex.slice(3,5),16);
    const b = parseInt(hex.slice(5,7),16);
    const d = n => Math.max(0, n - 30).toString(16).padStart(2,'0');
    return '#' + d(r) + d(g) + d(b);
  } catch(e) { return hex; }
}

/** Retorna as 14 categorias de ranking. */
function _getCategorias() {
  const H_ = H.ATIV;
  return [
    {
      titulo: '🏃 Corrida — Dist. acumulada', cor: '#FC4C02', unidade: 'km',
      fn: function(dados) {
        return _somarPorAtleta(dados, [H_.TIPO-1], ['Corrida','Corrida (Trail)','Run','TrailRun'], H_.DIST_KM-1, false);
      }
    },
    {
      titulo: '🏃 Corrida — Melhor pace médio', cor: '#E84700', unidade: 'min/km',
      fn: function(dados) {
        return _mediaFiltradaPorAtleta(dados, [H_.TIPO-1], ['Corrida','Corrida (Trail)','Run','TrailRun'], H_.PACE_S-1, true);
      }
    },
    {
      titulo: '🏃 Corrida — Maior dist. única', cor: '#D44000', unidade: 'km',
      fn: function(dados) {
        return _maxPorAtleta(dados, [H_.TIPO-1], ['Corrida','Corrida (Trail)','Run','TrailRun'], H_.DIST_KM-1, false);
      }
    },
    {
      titulo: '🏃 Corrida — Nº de treinos', cor: '#C83B00', unidade: 'saídas',
      fn: function(dados) {
        return _contarPorAtleta(dados, [H_.TIPO-1], ['Corrida','Corrida (Trail)','Run','TrailRun']);
      }
    },
    {
      titulo: '🚶 Caminhada — Dist. acumulada', cor: '#27AE60', unidade: 'km',
      fn: function(dados) {
        return _somarPorAtleta(dados, [H_.TIPO-1], ['Caminhada','Trilha','Walk','Hike'], H_.DIST_KM-1, false);
      }
    },
    {
      titulo: '🚶 Caminhada — Mais saídas', cor: '#1E8449', unidade: 'saídas',
      fn: function(dados) {
        return _contarPorAtleta(dados, [H_.TIPO-1], ['Caminhada','Trilha','Walk','Hike']);
      }
    },
    {
      titulo: '⏱️ Tempo em movimento — Corrida+Caminhada', cor: '#8E44AD', unidade: 'h:mm',
      fn: function(dados) {
        return _somarHmsPorAtleta(dados, [H_.TIPO-1], ['Corrida','Corrida (Trail)','Run','TrailRun','Caminhada','Trilha','Walk','Hike'], H_.MOV_HMS-1);
      }
    },
    {
      titulo: '💪 Academia — Mais sessões', cor: '#2471A3', unidade: 'sessões',
      fn: function(dados) {
        return _contarPorAtleta(dados, [H_.TIPO-1], ['Academia','Musculação','WeightTraining','Workout']);
      }
    },
    {
      titulo: '🚴 Ciclismo — Dist. acumulada', cor: '#F39C12', unidade: 'km',
      fn: function(dados) {
        return _somarPorAtleta(dados, [H_.TIPO-1], ['Ciclismo','Ergométrica','Ride','VirtualRide','EBikeRide'], H_.DIST_KM-1, false);
      }
    },
    {
      titulo: '🏊 Natação — Dist. acumulada', cor: '#1ABC9C', unidade: 'km',
      fn: function(dados) {
        return _somarPorAtleta(dados, [H_.TIPO-1], ['Natação','Swim'], H_.DIST_KM-1, false);
      }
    },
    {
      titulo: '⭐ Consistência — Dias ativos', cor: '#C0392B', unidade: 'dias',
      fn: function(dados) {
        return _diasAtivosPorAtleta(dados);
      }
    },
    {
      titulo: '🏅 Geral — Maior volume total', cor: '#922B21', unidade: 'km',
      fn: function(dados) {
        return _somarPorAtleta(dados, null, null, H_.DIST_KM-1, false);
      }
    },
    {
      titulo: '⛰️ Geral — Maior elevação acumulada', cor: '#7F8C8D', unidade: 'm',
      fn: function(dados) {
        return _somarPorAtleta(dados, null, null, H_.ELEV-1, false);
      }
    },
    {
      titulo: '❤️ FC Média mais alta (Corrida)', cor: '#E74C3C', unidade: 'bpm',
      fn: function(dados) {
        return _mediaFiltradaPorAtleta(dados, [H_.TIPO-1], ['Corrida','Corrida (Trail)','Run','TrailRun'], H_.FC_MED-1, false);
      }
    }
  ];
}

// ---- Funções auxiliares de agregação ----

function _somarPorAtleta(dados, colTipos, tipos, colValor, crescente) {
  const map = {};
  dados.forEach(row => {
    if (colTipos && tipos) {
      const tipo = String(row[colTipos[0]] || '');
      if (!tipos.some(t => tipo.includes(t))) return;
    }
    const nome = String(row[H.ATIV.NOME - 1] || '').trim();
    if (!nome) return;
    const val = _n(row, colValor);
    if (!map[nome]) map[nome] = 0;
    map[nome] += val;
  });
  return _sortMap(map, crescente);
}

function _contarPorAtleta(dados, colTipos, tipos) {
  const map = {};
  dados.forEach(row => {
    if (colTipos && tipos) {
      const tipo = String(row[colTipos[0]] || '');
      if (!tipos.some(t => tipo.includes(t))) return;
    }
    const nome = String(row[H.ATIV.NOME - 1] || '').trim();
    if (!nome) return;
    if (!map[nome]) map[nome] = 0;
    map[nome]++;
  });
  return _sortMap(map, false);
}

function _maxPorAtleta(dados, colTipos, tipos, colValor) {
  const map = {};
  dados.forEach(row => {
    if (colTipos && tipos) {
      const tipo = String(row[colTipos[0]] || '');
      if (!tipos.some(t => tipo.includes(t))) return;
    }
    const nome = String(row[H.ATIV.NOME - 1] || '').trim();
    if (!nome) return;
    const val = _n(row, colValor);
    if (!map[nome] || val > map[nome]) map[nome] = val;
  });
  return _sortMap(map, false);
}

function _mediaFiltradaPorAtleta(dados, colTipos, tipos, colValor, crescente) {
  const map = {};
  dados.forEach(row => {
    if (colTipos && tipos) {
      const tipo = String(row[colTipos[0]] || '');
      if (!tipos.some(t => tipo.includes(t))) return;
    }
    const nome = String(row[H.ATIV.NOME - 1] || '').trim();
    if (!nome) return;
    const val = _n(row, colValor);
    if (val <= 0) return;
    if (!map[nome]) map[nome] = {sum: 0, cnt: 0};
    map[nome].sum += val;
    map[nome].cnt++;
  });
  const result = {};
  Object.keys(map).forEach(k => { result[k] = map[k].sum / map[k].cnt; });
  return _sortMap(result, crescente);
}

function _somarHmsPorAtleta(dados, colTipos, tipos, colHms) {
  const map = {};
  const displayMap = {};
  dados.forEach(row => {
    if (colTipos && tipos) {
      const tipo = String(row[colTipos[0]] || '');
      if (!tipos.some(t => tipo.includes(t))) return;
    }
    const nome = String(row[H.ATIV.NOME - 1] || '').trim();
    if (!nome) return;
    const s = _hms2s(row[colHms]);
    if (!map[nome]) map[nome] = 0;
    map[nome] += s;
  });
  Object.keys(map).forEach(k => {
    const totalS = map[k];
    const h = Math.floor(totalS / 3600);
    const m = Math.floor((totalS % 3600) / 60);
    displayMap[k] = h + 'h' + String(m).padStart(2,'0');
  });
  return _sortMap(displayMap, false, map);
}

function _diasAtivosPorAtleta(dados) {
  const seen = {};
  dados.forEach(row => {
    const nome = String(row[H.ATIV.NOME - 1] || '').trim();
    if (!nome) return;
    const d = row[H.ATIV.DATA - 1];
    if (!(d instanceof Date)) return;
    seen[nome + '|' + d.toDateString()] = nome;
  });
  const count = {};
  Object.values(seen).forEach(nome => {
    count[nome] = (count[nome] || 0) + 1;
  });
  return _sortMap(count, false);
}

function _sortMap(map, crescente, sortMap) {
  const sm = sortMap || map;
  return Object.keys(sm)
    .sort((a, b) => crescente ? sm[a] - sm[b] : sm[b] - sm[a])
    .map(nome => {
      let val = map[nome];
      if (typeof val === 'number') {
        val = Math.round(val * 100) / 100;
      }
      return {nome, valor: val};
    });
}

function _n(row, colIdx) {
  const v = row[colIdx];
  if (v === null || v === undefined || v === '') return 0;
  const n = parseFloat(String(v).replace(',','.'));
  return isNaN(n) ? 0 : n;
}

function _hms2s(val) {
  if (!val) return 0;
  const s = String(val).trim();
  const parts = s.split(':');
  if (parts.length === 3) {
    return parseInt(parts[0])*3600 + parseInt(parts[1])*60 + parseInt(parts[2]);
  } else if (parts.length === 2) {
    return parseInt(parts[0])*60 + parseInt(parts[1]);
  }
  return 0;
}
