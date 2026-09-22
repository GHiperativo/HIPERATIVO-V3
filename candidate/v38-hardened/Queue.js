/**
 * ═══════════════════════════════════════════════════════════════════════
 * HIPERATIVO V3 — Queue.gs (v1.0 — 14/06/2026)
 * Sistema inteligente de fila e rate-limiting para Strava API
 *
 * Limites Strava default: 200 req/15min | 2000 req/dia
 * Usamos 80%:             160 req/15min | 1600 req/dia
 *
 * Lógica:
 *  - Round-robin: cada ciclo continua de onde parou (ponteiro PropertiesService)
 *  - Atletas novos têm flag Q_HIST_{athId} = próxima página de histórico
 *  - Histórico importa 2 páginas (100 ativ) por ciclo, retomável entre execuções
 *  - 429 Strava interrompe o ciclo imediatamente
 * ═══════════════════════════════════════════════════════════════════════
 */

const Q_LIMITE_15MIN   = 160;
const Q_LIMITE_DIA     = 1600;
const Q_REQ_POR_ATLETA = 4;
const Q_HIST_PAGINAS   = 2;
const Q_MAX_ATLETAS    = 20;

// ── PROCESSADOR PRINCIPAL DA FILA ────────────────────────────────────────────
function processarFilaStrava() {
  const props   = PropertiesService.getScriptProperties();
  const agora   = new Date();
  const hojeStr = Utilities.formatDate(agora, 'America/Sao_Paulo', 'yyyy-MM-dd');
  const janela  = Math.floor(agora.getTime() / (15 * 60 * 1000));

  let contDia    = parseInt(props.getProperty('Q_DIA_' + hojeStr) || '0');
  let contJanela = parseInt(props.getProperty('Q_15M_' + janela)  || '0');

  if (contDia >= Q_LIMITE_DIA) {
    _log('SISTEMA', 'AVISO', 'processarFilaStrava', 'Limite diário atingido: ' + contDia + '/' + Q_LIMITE_DIA, '');
    return;
  }
  if (contJanela >= Q_LIMITE_15MIN) {
    _log('SISTEMA', 'AVISO', 'processarFilaStrava', 'Limite 15min atingido: ' + contJanela + '/' + Q_LIMITE_15MIN, '');
    return;
  }

  // Capacidade restante no ciclo
  const capJanela  = Math.floor((Q_LIMITE_15MIN - contJanela) / Q_REQ_POR_ATLETA);
  const capDia     = Math.floor((Q_LIMITE_DIA   - contDia)    / Q_REQ_POR_ATLETA);
  const maxAtletas = Math.min(capJanela, capDia, Q_MAX_ATLETAS);
  if (maxAtletas <= 0) return;

  // Carregar atletas ativos
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const shTok = ss.getSheetByName(H.SHEETS.TOKENS);
  if (!shTok) return;

  const tokDados       = shTok.getDataRange().getValues();
  const atletasAtivos  = [];
  for (let i = 1; i < tokDados.length; i++) {
    const athId  = String(tokDados[i][H.TOK.ATH_ID - 1] || '').trim();
    const status = String(tokDados[i][H.TOK.STATUS  - 1] || '').trim();
    if (!athId || status.toLowerCase() === 'inativo') continue;
    atletasAtivos.push(athId);
  }
  if (atletasAtivos.length === 0) return;

  // Ponteiro round-robin
  let posAtual = parseInt(props.getProperty('Q_POS_FILA') || '0');
  if (posAtual >= atletasAtivos.length) posAtual = 0;

  let reqUsadas    = 0;
  let processados  = 0;
  let interrompido = false;

  for (let i = 0; i < Math.min(maxAtletas, atletasAtivos.length); i++) {
    if (interrompido) break;
    const idx   = (posAtual + i) % atletasAtivos.length;
    const athId = atletasAtivos[idx];

    const histKey = 'Q_HIST_' + athId;
    const histPag = parseInt(props.getProperty(histKey) || '-1');
    const isNovo  = histPag >= 0;

    try {
      if (isNovo) {
        const resultado = _importarHistoricoPaginado(athId, histPag);
        reqUsadas += Q_REQ_POR_ATLETA;

        if (resultado.rateLimitado) {
          interrompido = true;
        } else if (resultado.novas < 50 * Q_HIST_PAGINAS) {
          // Chegou no final do histórico
          props.deleteProperty(histKey);
          _log(athId, 'INFO', 'processarFilaStrava', 'Histórico completo. Total págs: ' + Math.ceil((histPag + Q_HIST_PAGINAS) / Q_HIST_PAGINAS), '');
        } else {
          // Ainda tem mais páginas
          props.setProperty(histKey, String(histPag + Q_HIST_PAGINAS));
        }
      } else {
        const novas = _importarAtividadesAtleta(athId, 2);
        reqUsadas  += Q_REQ_POR_ATLETA;
        if (novas > 0) _log(athId, 'INFO', 'processarFilaStrava', novas + ' novas atividades', '');
      }
      processados++;
    } catch (e) {
      reqUsadas += 1;
      if (String(e.message).indexOf('429') >= 0) {
        interrompido = true;
        _log(athId, 'AVISO', 'processarFilaStrava', 'Rate limit 429 — ciclo interrompido', '');
      } else {
        _log(athId, 'ERRO', 'processarFilaStrava', e.message, '');
      }
    }
  }

  // Avançar ponteiro
  const novaPos = (posAtual + processados) % Math.max(atletasAtivos.length, 1);
  props.setProperty('Q_POS_FILA', String(novaPos));

  // Atualizar contadores
  props.setProperty('Q_DIA_' + hojeStr, String(contDia + reqUsadas));
  props.setProperty('Q_15M_' + janela,  String(contJanela + reqUsadas));

  _limparChavesAnteriores(props, hojeStr, janela);
  _log('SISTEMA', 'INFO', 'processarFilaStrava',
    'Ciclo: ' + processados + ' atletas | ' + reqUsadas + ' req | total dia: ' + (contDia + reqUsadas) + '/' + Q_LIMITE_DIA, '');
}

// ── IMPORTAÇÃO HISTÓRICA PAGINADA ─────────────────────────────────────────────
function _importarHistoricoPaginado(athId, paginaInicio) {
  const accessToken = _getValidAccessToken(athId);
  const nomeAtleta  = _getNomeAtleta(athId);
  let   totalNovas  = 0;
  let   rateLimitado = false;

  for (let pg = paginaInicio + 1; pg <= paginaInicio + Q_HIST_PAGINAS; pg++) {
    const url  = STRAVA_API_BASE + '/athlete/activities?per_page=50&page=' + pg;
    const resp = UrlFetchApp.fetch(url, {
      headers: { Authorization: 'Bearer ' + accessToken },
      muteHttpExceptions: true
    });

    const code = resp.getResponseCode();
    if (code === 429) { rateLimitado = true; break; }
    if (code !== 200) break;

    const page = JSON.parse(resp.getContentText());
    if (!page || !page.length) break;

    totalNovas += _gravarAtividades(athId, nomeAtleta, page);
  }

  return { novas: totalNovas, rateLimitado };
}

// ── REGISTRAR ATLETA NOVO PARA IMPORTAÇÃO HISTÓRICA ──────────────────────────
// Chamado automaticamente no callback OAuth (WebApp.gs)
function registrarAtletaParaHistorico(athId) {
  if (!athId) return;
  PropertiesService.getScriptProperties().setProperty('Q_HIST_' + athId, '0');
  _log(athId, 'INFO', 'registrarAtletaParaHistorico', 'Importação histórica agendada', '');
}

// ── STATUS DA FILA (menu → diagnóstico) ──────────────────────────────────────
function statusFila() {
  const props   = PropertiesService.getScriptProperties();
  const agora   = new Date();
  const hojeStr = Utilities.formatDate(agora, 'America/Sao_Paulo', 'yyyy-MM-dd');
  const janela  = Math.floor(agora.getTime() / (15 * 60 * 1000));

  const contDia    = parseInt(props.getProperty('Q_DIA_' + hojeStr) || '0');
  const contJanela = parseInt(props.getProperty('Q_15M_' + janela)  || '0');
  const pos        = props.getProperty('Q_POS_FILA') || '0';
  const allProps   = props.getProperties();

  const histPendentes = Object.keys(allProps)
    .filter(k => k.startsWith('Q_HIST_'))
    .map(k => '  • ' + k.replace('Q_HIST_', '') + ' (pág. ' + allProps[k] + ')');

  const linhas = [
    '📊 STATUS DA FILA STRAVA',
    '─────────────────────────',
    'Req hoje:       ' + contDia + '/' + Q_LIMITE_DIA + ' (' + Math.round(contDia / Q_LIMITE_DIA * 100) + '%)',
    'Req (15min):    ' + contJanela + '/' + Q_LIMITE_15MIN,
    'Posição fila:   ' + pos,
    '',
    'Histórico pendente (' + histPendentes.length + '):',
    histPendentes.length ? histPendentes.join('\n') : '  Nenhum.',
    '',
    'Capacidade hoje: ~' + Math.floor((Q_LIMITE_DIA - contDia) / Q_REQ_POR_ATLETA) + ' atletas ainda',
  ];

  try { SpreadsheetApp.getUi().alert(linhas.join('\n')); } catch(_) { Logger.log(linhas.join('\n')); }
}

// ── RESETAR CONTADORES (debug) ────────────────────────────────────────────────
function resetarContadoresFila() {
  const props = PropertiesService.getScriptProperties();
  const all   = props.getProperties();
  Object.keys(all)
    .filter(k => k.startsWith('Q_DIA_') || k.startsWith('Q_15M_'))
    .forEach(k => props.deleteProperty(k));
  _log('SISTEMA', 'INFO', 'resetarContadoresFila', 'Contadores zerados', '');
  try { SpreadsheetApp.getUi().alert('✅ Contadores de rate limit zerados.'); } catch(_) {}
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
