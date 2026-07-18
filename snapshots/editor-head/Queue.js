/**
 * ═══════════════════════════════════════════════════════════════════════
 * HIPERATIVO V3 — Queue.gs (v1.0 — 14/06/2026)
 * Sistema inteligente de fila e rate-limiting para Strava API
 *
 * Limites: obtidos dos headers de cada resposta; fallback conservador local.
 *
 * Lógica:
 *  - Round-robin: cada ciclo continua de onde parou (ponteiro PropertiesService)
 *  - Atletas novos têm flag Q_HIST_{athId} = próxima página de histórico
 *  - Histórico importa lotes pequenos por ciclo, retomável entre execuções
 *  - O cursor avança pelo tamanho da resposta da API, não pela quantidade nova
 *  - Os limites são lidos dos headers oficiais retornados pela Strava
 *  - 429 Strava interrompe o ciclo imediatamente
 * ═══════════════════════════════════════════════════════════════════════
 */

const Q_LIMITE_15MIN_FALLBACK = 80;
const Q_LIMITE_DIA_FALLBACK   = 800;
const Q_HIST_PAGINAS          = 2;
const Q_PER_PAGE              = 50;
const Q_MAX_ATLETAS           = 20;
const Q_MAX_EXEC_MS           = 4 * 60 * 1000;
const Q_LEASE_MS              = 8 * 60 * 1000;

// ── PROCESSADOR PRINCIPAL DA FILA ────────────────────────────────────────────
function processarFilaStrava() {
  const execId = _qAdquirirExecucao_();
  if (!execId) {
    _log('SISTEMA', 'INFO', 'processarFilaStrava', 'Outro ciclo da fila ainda está em execução', '');
    return;
  }

  const inicio = Date.now();
  const props = PropertiesService.getScriptProperties();
  try {
    if (!_qTemCapacidade_(props)) {
      const rate = _qEstadoRate_(props);
      _log('SISTEMA', 'AVISO', 'processarFilaStrava',
        'Margem de rate limit atingida: ' + rate.uso15 + '/' + rate.limite15 +
        ' em 15min | ' + rate.usoDia + '/' + rate.limiteDia + ' no dia', '');
      return;
    }

  // Carregar atletas ativos
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const shTok = ss.getSheetByName(H.SHEETS.TOKENS);
  if (!shTok) return;

    const tokDados = shTok.getDataRange().getValues();
    const idsUnicos = new Set();
    for (let i = 1; i < tokDados.length; i++) {
      const athId  = String(tokDados[i][H.TOK.ATH_ID - 1] || '').trim();
      const status = String(tokDados[i][H.TOK.STATUS  - 1] || '').trim().toLowerCase();
      if (!_isAthIdValido_(athId) || status === 'inativo' || status === 'pendente') continue;
      idsUnicos.add(athId);
    }
    const atletasAtivos = Array.from(idsUnicos);
    if (atletasAtivos.length === 0) return;

  // Ponteiro round-robin
    let posAtual = parseInt(props.getProperty('Q_POS_FILA') || '0', 10);
    if (posAtual >= atletasAtivos.length) posAtual = 0;

    let reqUsadas    = 0;
    let processados  = 0;
    let interrompido = false;

    for (let i = 0; i < Math.min(Q_MAX_ATLETAS, atletasAtivos.length); i++) {
      if (interrompido || Date.now() - inicio >= Q_MAX_EXEC_MS || !_qTemCapacidade_(props)) break;
      const idx   = (posAtual + i) % atletasAtivos.length;
      const athId = atletasAtivos[idx];

      const histKey = 'Q_HIST_' + athId;
      const histVal = props.getProperty(histKey);
      const histPag = histVal === null ? -1 : Math.max(0, parseInt(histVal || '0', 10));
      const isNovo  = histPag >= 0;

      try {
        const resultado = _importarHistoricoPaginado(athId, isNovo ? histPag : 0);
        reqUsadas += resultado.requests;

        if (isNovo && resultado.ultimaPagina > histPag) {
          // Persiste apenas páginas integralmente gravadas. Se o ciclo cair, a
          // próxima execução continua sem pular nenhuma atividade.
          props.setProperty(histKey, String(resultado.ultimaPagina));
        }
        if (isNovo && resultado.concluido) {
          props.deleteProperty(histKey);
          props.setProperty('Q_HIST_DONE_' + athId, String(resultado.ultimaPagina));
          _log(athId, 'INFO', 'processarFilaStrava',
            'Histórico completo até a página ' + resultado.ultimaPagina, '');
        }
        if (resultado.novas > 0) {
          _log(athId, 'INFO', 'processarFilaStrava', resultado.novas + ' novas atividades', '');
        }
        if (resultado.erro) {
          _log(athId, 'ERRO', 'processarFilaStrava', resultado.erro, '');
        }
        if (resultado.rateLimitado) interrompido = true;
        processados++;
      } catch (e) {
        _log(athId, 'ERRO', 'processarFilaStrava', e.message, '');
      }
    }

  // Avançar ponteiro
    const novaPos = (posAtual + processados) % Math.max(atletasAtivos.length, 1);
    props.setProperty('Q_POS_FILA', String(novaPos));
    const rateFinal = _qEstadoRate_(props);
    _log('SISTEMA', 'INFO', 'processarFilaStrava',
      'Ciclo: ' + processados + ' atletas | ' + reqUsadas +
      ' req | Strava: ' + rateFinal.uso15 + '/' + rateFinal.limite15 +
      ' em 15min, ' + rateFinal.usoDia + '/' + rateFinal.limiteDia + ' no dia', '');
  } finally {
    _qLiberarExecucao_(execId);
  }
}

// ── IMPORTAÇÃO HISTÓRICA PAGINADA ─────────────────────────────────────────────
function _importarHistoricoPaginado(athId, paginaInicio) {
  let   accessToken = _getValidAccessToken(athId);
  if (!accessToken) throw new Error('Nenhum access_token válido encontrado nas fontes de segurança.');
  const nomeAtleta  = _getNomeAtleta(athId);
  let   totalNovas  = 0;
  let   rateLimitado = false;
  let   requests = 0;
  let   ultimaPagina = paginaInicio;
  let   concluido = false;
  let   erro = '';

  for (let pg = paginaInicio + 1; pg <= paginaInicio + Q_HIST_PAGINAS; pg++) {
    if (!_qTemCapacidade_(PropertiesService.getScriptProperties())) break;
    const url  = STRAVA_API_BASE + '/athlete/activities?per_page=' + Q_PER_PAGE + '&page=' + pg;
    let resp = UrlFetchApp.fetch(url, {
      headers: { Authorization: 'Bearer ' + accessToken },
      muteHttpExceptions: true
    });
    requests++;
    _qRegistrarRate_(resp);

    let code = resp.getResponseCode();
    if (code === 401 && typeof _forcarRefreshAccessToken_ === 'function') {
      // Um access_token pode ser recusado antes do expires_at. Tenta uma única
      // renovação usando o refresh mais novo (Planilha/Properties/Supabase).
      accessToken = _forcarRefreshAccessToken_(athId, accessToken);
      if (accessToken) {
        resp = UrlFetchApp.fetch(url, {
          headers: { Authorization: 'Bearer ' + accessToken },
          muteHttpExceptions: true
        });
        requests++;
        _qRegistrarRate_(resp);
        code = resp.getResponseCode();
      }
    }
    if (code === 429) { rateLimitado = true; break; }
    if (code !== 200) {
      erro = 'Strava HTTP ' + code + ' na página ' + pg + '; cursor preservado.';
      break;
    }

    const page = JSON.parse(resp.getContentText());
    if (!Array.isArray(page)) {
      erro = 'Resposta inesperada da Strava na página ' + pg + '; cursor preservado.';
      break;
    }
    if (page.length === 0) {
      concluido = true;
      break;
    }

    totalNovas += _gravarAtividades(athId, nomeAtleta, page);
    ultimaPagina = pg;
    if (page.length < Q_PER_PAGE) {
      concluido = true;
      break;
    }
  }

  return { novas: totalNovas, rateLimitado, requests, ultimaPagina, concluido, erro };
}

// ── REGISTRAR ATLETA NOVO PARA IMPORTAÇÃO HISTÓRICA ──────────────────────────
// Chamado automaticamente no callback OAuth (WebApp.gs)
function registrarAtletaParaHistorico(athId) {
  if (!_isAthIdValido_(athId)) return;
  const props = PropertiesService.getScriptProperties();
  props.setProperty('Q_HIST_' + athId, '0');
  props.deleteProperty('Q_HIST_DONE_' + athId);
  _log(athId, 'INFO', 'registrarAtletaParaHistorico', 'Importação histórica agendada', '');
}

/**
 * Agenda uma única varredura completa para todos os tokens válidos existentes.
 * Não reinicia atletas já concluídos e não altera nenhuma credencial.
 */
function agendarHistoricoCompletoTodos() {
  const props = PropertiesService.getScriptProperties();
  const shTok = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(H.SHEETS.TOKENS);
  if (!shTok) throw new Error('Aba TOKENS não encontrada.');

  const dados = shTok.getDataRange().getValues();
  const ids = new Set();
  for (let i = 1; i < dados.length; i++) {
    const athId = String(dados[i][H.TOK.ATH_ID - 1] || '').trim();
    const status = String(dados[i][H.TOK.STATUS - 1] || '').trim().toLowerCase();
    if (_isAthIdValido_(athId) && status !== 'inativo' && status !== 'pendente') ids.add(athId);
  }

  let agendados = 0;
  let preservados = 0;
  ids.forEach(athId => {
    const histKey = 'Q_HIST_' + athId;
    const doneKey = 'Q_HIST_DONE_' + athId;
    if (props.getProperty(histKey) !== null || props.getProperty(doneKey) !== null) {
      preservados++;
      return;
    }
    props.setProperty(histKey, '0');
    agendados++;
  });

  _log('SISTEMA', 'INFO', 'agendarHistoricoCompletoTodos',
    agendados + ' atletas agendados; ' + preservados + ' cursores existentes preservados', '');
  return { agendados, preservados, total: ids.size };
}

// ── STATUS DA FILA (menu → diagnóstico) ──────────────────────────────────────
function statusFila() {
  const props   = PropertiesService.getScriptProperties();
  const rate      = _qEstadoRate_(props);
  const pos        = props.getProperty('Q_POS_FILA') || '0';
  const allProps   = props.getProperties();

  const histPendentes = Object.keys(allProps)
    .filter(k => k.startsWith('Q_HIST_'))
    .map(k => '  • ' + k.replace('Q_HIST_', '') + ' (pág. ' + allProps[k] + ')');

  const linhas = [
    '📊 STATUS DA FILA STRAVA',
    '─────────────────────────',
    'Req hoje:       ' + rate.usoDia + '/' + rate.limiteDia +
      ' (' + Math.round(rate.usoDia / rate.limiteDia * 100) + '%)',
    'Req (15min):    ' + rate.uso15 + '/' + rate.limite15,
    'Posição fila:   ' + pos,
    '',
    'Histórico pendente (' + histPendentes.length + '):',
    histPendentes.length ? histPendentes.join('\n') : '  Nenhum.',
    '',
    'A fila respeita a margem de segurança informada pela própria Strava.',
  ];

  try { SpreadsheetApp.getUi().alert(linhas.join('\n')); } catch(_) { Logger.log(linhas.join('\n')); }
}

// ── RESETAR CONTADORES (debug) ────────────────────────────────────────────────
function resetarContadoresFila() {
  const props = PropertiesService.getScriptProperties();
  const all   = props.getProperties();
  Object.keys(all)
    .filter(k => k.startsWith('Q_DIA_') || k.startsWith('Q_15M_') || k.startsWith('Q_RATE_'))
    .forEach(k => props.deleteProperty(k));
  _log('SISTEMA', 'INFO', 'resetarContadoresFila', 'Contadores zerados', '');
  try { SpreadsheetApp.getUi().alert('✅ Contadores de rate limit zerados.'); } catch(_) {}
}

// ── RATE LIMIT REAL (HEADERS STRAVA) ────────────────────────────────────────

function _qEstadoRate_(props) {
  const agora = Date.now();
  const slotAtual = String(Math.floor(agora / (15 * 60 * 1000)));
  const diaAtual = Utilities.formatDate(new Date(agora), 'UTC', 'yyyy-MM-dd');
  const mesmoSlot = props.getProperty('Q_RATE_SLOT') === slotAtual;
  const mesmoDia = props.getProperty('Q_RATE_DAY') === diaAtual;

  return {
    limite15: Math.max(1, parseInt(props.getProperty('Q_RATE_LIMIT_15') || Q_LIMITE_15MIN_FALLBACK, 10)),
    limiteDia: Math.max(1, parseInt(props.getProperty('Q_RATE_LIMIT_DIA') || Q_LIMITE_DIA_FALLBACK, 10)),
    uso15: mesmoSlot ? Math.max(0, parseInt(props.getProperty('Q_RATE_USAGE_15') || '0', 10)) : 0,
    usoDia: mesmoDia ? Math.max(0, parseInt(props.getProperty('Q_RATE_USAGE_DIA') || '0', 10)) : 0,
    slot: slotAtual,
    dia: diaAtual,
  };
}

function _qTemCapacidade_(props) {
  const rate = _qEstadoRate_(props);
  const margem15 = Math.max(5, Math.ceil(rate.limite15 * 0.10));
  const margemDia = Math.max(50, Math.ceil(rate.limiteDia * 0.10));
  return rate.uso15 < rate.limite15 - margem15 &&
    rate.usoDia < rate.limiteDia - margemDia;
}

function _qRegistrarRate_(resp) {
  const props = PropertiesService.getScriptProperties();
  const rateAtual = _qEstadoRate_(props);
  let headers = {};
  try { headers = resp.getAllHeaders ? resp.getAllHeaders() : resp.getHeaders(); } catch (_) { }

  const limites = _qParseParHeader_(headers, 'X-ReadRateLimit-Limit') ||
    _qParseParHeader_(headers, 'X-RateLimit-Limit');
  const usos = _qParseParHeader_(headers, 'X-ReadRateLimit-Usage') ||
    _qParseParHeader_(headers, 'X-RateLimit-Usage');

  props.setProperties({
    Q_RATE_SLOT: rateAtual.slot,
    Q_RATE_DAY: rateAtual.dia,
    Q_RATE_LIMIT_15: String(limites ? limites[0] : rateAtual.limite15),
    Q_RATE_LIMIT_DIA: String(limites ? limites[1] : rateAtual.limiteDia),
    Q_RATE_USAGE_15: String(usos ? usos[0] : rateAtual.uso15 + 1),
    Q_RATE_USAGE_DIA: String(usos ? usos[1] : rateAtual.usoDia + 1),
  }, false);
}

function _qParseParHeader_(headers, nome) {
  const chave = Object.keys(headers || {}).find(k => String(k).toLowerCase() === nome.toLowerCase());
  if (!chave) return null;
  const partes = String(headers[chave]).split(',').map(v => parseInt(v.trim(), 10));
  return partes.length >= 2 && partes.every(Number.isFinite) ? partes : null;
}

// Evita que dois gatilhos percorram a mesma página ao mesmo tempo. O lease tem
// validade limitada, então uma execução interrompida não trava a fila para sempre.
function _qAdquirirExecucao_() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return null;
  try {
    const props = PropertiesService.getScriptProperties();
    const agora = Date.now();
    const ate = Number(props.getProperty('Q_EXEC_UNTIL') || 0);
    if (ate > agora) return null;
    const execId = Utilities.getUuid();
    props.setProperties({
      Q_EXEC_OWNER: execId,
      Q_EXEC_UNTIL: String(agora + Q_LEASE_MS),
    }, false);
    return execId;
  } finally {
    lock.releaseLock();
  }
}

function _qLiberarExecucao_(execId) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return;
  try {
    const props = PropertiesService.getScriptProperties();
    if (props.getProperty('Q_EXEC_OWNER') === execId) {
      props.deleteProperty('Q_EXEC_OWNER');
      props.deleteProperty('Q_EXEC_UNTIL');
    }
  } finally {
    lock.releaseLock();
  }
}

// ── LIMPAR CHAVES ANTIGAS ─────────────────────────────────────────────────────
function _limparChavesAnteriores(props, hojeStr, janelaAtual) {
  try {
    const all = props.getProperties();
    Object.keys(all).forEach(k => {
      if (k.startsWith('Q_DIA_') && !k.endsWith(hojeStr)) props.deleteProperty(k);
      if (k.startsWith('Q_15M_')) {
        const j = parseInt(k.replace('Q_15M_', ''));
        if (janelaAtual - j > 12) props.deleteProperty(k); // mais de 3h atrás
      }
    });
  } catch(_) {}
}
