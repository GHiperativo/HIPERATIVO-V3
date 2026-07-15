// ═══════════════════════════════════════════════════════════════════════════════
// StravaPipeline.gs — Utilitários de conversão de dados Strava
// HIPERATIVO V3 | Funções auxiliares internas (sufixo _)
// ❌ NÃO CHAMAR API STRAVA  ❌ NÃO ALTERAR SUPABASE/TOKENS
// ❌ NÃO MEXER EM MÉTRICAS/PAINEL/GRÁFICOS/RANKING
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Traduz o tipo de atividade Strava (inglês → português).
 * @param {string} tipo - Sport Type ou Type do raw Strava
 * @returns {string}
 */
function traduzirTipoStrava_(tipo) {
  var mapa = {
    'Run':                           'Corrida',
    'TrailRun':                      'Corrida em trilha',
    'VirtualRun':                    'Corrida virtual',
    'Walk':                          'Caminhada',
    'Hike':                          'Trilha',
    'Ride':                          'Ciclismo',
    'MountainBikeRide':              'Mountain Bike',
    'VirtualRide':                   'Ciclismo virtual',
    'Swim':                          'Natação',
    'Workout':                       'Treino',
    'WeightTraining':                'Musculação',
    'HighIntensityIntervalTraining': 'HIIT',
    'Yoga':                          'Yoga',
    'Crossfit':                      'CrossFit',
    'Elliptical':                    'Elíptico',
    'StairStepper':                  'Escada',
    'Rowing':                        'Remo',
    'Kayaking':                      'Caiaque',
    'Soccer':                        'Futebol',
    'Tennis':                        'Tênis',
    'Pilates':                       'Pilates',
    'Golf':                          'Golfe',
    'Skating':                       'Patinação',
    'IceSkate':                      'Patinação no gelo',
    'AlpineSki':                     'Esqui alpino',
    'NordicSki':                     'Esqui nórdico',
    'Snowboard':                     'Snowboard',
    'Surfing':                       'Surfe',
    'Kitesurf':                      'Kitesurf',
    'Windsurf':                      'Windsurf',
    'RockClimbing':                  'Escalada',
    'Wheelchair':                    'Cadeira de rodas',
    'HandCycle':                     'Handbike',
    'Other':                         'Outro'
  };
  return mapa[tipo] || tipo || 'Outro';
}

/**
 * Retorna true se o tipo de atividade envolve distância mensurável.
 * @param {string} tipo - Sport Type em inglês
 * @returns {boolean}
 */
function isAtividadeComDistancia_(tipo) {
  var comDist = [
    'Run','TrailRun','VirtualRun','Walk','Hike',
    'Ride','MountainBikeRide','VirtualRide',
    'Swim','Kayaking','Rowing','Skating',
    'IceSkate','AlpineSki','NordicSki','Wheelchair','HandCycle'
  ];
  return comDist.indexOf(tipo) !== -1;
}

/**
 * Retorna true se o tipo NÃO requer distância (yoga, musculação, etc.).
 * @param {string} tipo - Sport Type em inglês
 * @returns {boolean}
 */
function isAtividadeSemDistanciaObrigatoria_(tipo) {
  var semDist = [
    'Workout','WeightTraining','HighIntensityIntervalTraining',
    'Yoga','Crossfit','Elliptical','StairStepper',
    'Soccer','Tennis','Pilates','Golf',
    'Surfing','Kitesurf','Windsurf','RockClimbing','Snowboard'
  ];
  return semDist.indexOf(tipo) !== -1;
}

/**
 * Formata segundos totais em string HH:MM:SS.
 * @param {number} segundos
 * @returns {string}
 */
function formatarTempo_(segundos) {
  if (!segundos || isNaN(segundos) || segundos < 0) return '00:00:00';
  var s = Math.round(segundos);
  var h = Math.floor(s / 3600);
  var m = Math.floor((s % 3600) / 60);
  var seg = s % 60;
  return (
    String(h).padStart(2,'0') + ':' +
    String(m).padStart(2,'0') + ':' +
    String(seg).padStart(2,'0')
  );
}

/**
 * Formata pace em "MM:SS min/km" a partir de velocidade em m/s.
 * @param {number} velocidadeMps - velocidade em metros por segundo
 * @returns {string}
 */
function formatarPace_(velocidadeMps) {
  if (!velocidadeMps || velocidadeMps <= 0) return '--:-- min/km';
  var paceSKm = Math.round(1000 / velocidadeMps);
  var min = Math.floor(paceSKm / 60);
  var sec = Math.round(paceSKm % 60);
  return String(min).padStart(2,'0') + ':' + String(sec).padStart(2,'0') + ' min/km';
}

/**
 * Classifica qualidade do dado Strava. Retorna {qualidade, flags}.
 * @param {Object} raw - linha raw do Strava
 * @returns {{qualidade: string, flags: string[]}}
 */
function classificarQualidadeDado_(raw) {
  var flags = [];
  var tipo = raw['Sport Type'] || raw['Type'] || '';

  if (!raw['Activity ID'] && !raw['Strava ID'])            flags.push('SEM_ID');
  if (!raw['Name'] && !raw['Nome da Atividade'])           flags.push('SEM_NOME');
  if (!raw['Moving Time s'] && !raw['Tempo Movimento s'])  flags.push('SEM_TEMPO');

  if (isAtividadeComDistancia_(tipo)) {
    if (!raw['Distance m'] || Number(raw['Distance m']) <= 0)           flags.push('SEM_DISTANCIA');
    if (!raw['Average Speed m/s'] || Number(raw['Average Speed m/s']) <= 0) flags.push('SEM_VELOCIDADE');
  }
  if (!raw['Average Heartrate'])                                        flags.push('SEM_FC');
  if (!raw['Average Cadence'] && isAtividadeComDistancia_(tipo))        flags.push('SEM_CADENCIA');

  var qualidade;
  if (flags.length === 0) {
    qualidade = 'OK';
  } else if (flags.indexOf('SEM_ID') !== -1 || flags.indexOf('SEM_TEMPO') !== -1) {
    qualidade = 'INCOMPLETO';
  } else {
    qualidade = 'OK_PARCIAL';
  }
  return { qualidade: qualidade, flags: flags };
}

/**
 * Carga de treino simples — sem API, sem zonas externas.
 * Fórmula: dist × (FC/100) × 10  |  ou  tempo(min) × (FC/100)
 * @param {number} distanciaKm
 * @param {number} tempoMinutos
 * @param {number} fcMedia
 * @returns {number}
 */
function calcularCargaSimples_(distanciaKm, tempoMinutos, fcMedia) {
  if (!fcMedia || fcMedia <= 0) return 0;
  var fc   = Number(fcMedia)      || 0;
  var dist = Number(distanciaKm)  || 0;
  var t    = Number(tempoMinutos) || 0;
  if (dist > 0) return Math.round(dist * (fc / 100) * 10);
  return Math.round(t * (fc / 100));
}

/**
 * Converte um objeto raw Strava no array de 49 valores da aba CONVERTIDAS.
 * ❌ NÃO grava nada. ❌ NÃO chama API. ❌ NÃO acessa tokens.
 * @param {Object} raw - objeto com campos do STRAVA_RAW
 * @returns {Array} array de 49 posições na ordem exata dos headers
 */
function converterAtividadeRawParaConvertida_(raw) {
  // ── Identidade ───────────────────────────────────────────────────
  var athId    = raw['ATH_ID']      || '';
  var atleta   = raw['Atleta']      || '';
  var stravaId = raw['Activity ID'] || raw['Strava ID'] || '';
  var nome     = raw['Name']        || raw['Nome da Atividade'] || '';
  var tipoRaw  = raw['Sport Type']  || raw['Type'] || '';
  var tipo     = traduzirTipoStrava_(tipoRaw);
  var manual   = (raw['Manual'] === true || raw['Manual'] === 'true') ? 'Sim' : 'Não';
  var gearId   = raw['Gear ID'] || '';
  var idInterno = 'ACT_' + stravaId;

  // ── Data/hora ────────────────────────────────────────────────────
  var dataHoraRaw = raw['Start Date Local'] || raw['Data/Hora'] || '';
  var dataHoraObj = dataHoraRaw ? new Date(dataHoraRaw) : null;
  var tz          = Session.getScriptTimeZone();
  var dataFmt  = dataHoraObj ? Utilities.formatDate(dataHoraObj, tz, 'dd/MM/yyyy') : '';
  var horaFmt  = dataHoraObj ? Utilities.formatDate(dataHoraObj, tz, 'HH:mm:ss')   : '';
  var ano      = dataHoraObj ? dataHoraObj.getFullYear()  : '';
  var mes      = dataHoraObj ? dataHoraObj.getMonth() + 1 : '';
  var mesRef   = dataHoraObj ? Utilities.formatDate(dataHoraObj, tz, 'yyyy-MM')    : '';
  var diasSem  = ['Domingo','Segunda','Terça','Quarta','Quinta','Sexta','Sábado'];
  var diaSem   = dataHoraObj ? diasSem[dataHoraObj.getDay()] : '';
  var semana   = dataHoraObj
    ? Math.ceil((dataHoraObj - new Date(dataHoraObj.getFullYear(), 0, 1)) / 604800000)
    : '';

  // ── Distância ────────────────────────────────────────────────────
  var distM   = Number(raw['Distance m']) || 0;
  var distKm  = distM > 0 ? Math.round(distM / 10) / 100 : 0;
  var distFmt = distKm > 0 ? distKm + ' km' : '-';

  // ── Tempos ───────────────────────────────────────────────────────
  var movS   = Number(raw['Moving Time s'])  || 0;
  var totS   = Number(raw['Elapsed Time s']) || 0;
  var movFmt = formatarTempo_(movS);
  var totFmt = formatarTempo_(totS);

  // ── Pace e velocidade ────────────────────────────────────────────
  var avgMps  = Number(raw['Average Speed m/s']) || 0;
  var paceSKm = avgMps > 0 ? Math.round(1000 / avgMps) : 0;
  var paceFmt = avgMps > 0 ? formatarPace_(avgMps) : '--:-- min/km';
  var velKmh  = avgMps > 0 ? Math.round(avgMps * 3.6 * 100) / 100 : 0;
  var velFmt  = velKmh > 0 ? velKmh + ' km/h' : '-';

  // ── FC ───────────────────────────────────────────────────────────
  var fcMed    = Number(raw['Average Heartrate']) || 0;
  var fcMedFmt = fcMed > 0 ? fcMed + ' bpm' : '-';
  var fcMax    = Number(raw['Max Heartrate'])     || 0;
  var fcMaxFmt = fcMax > 0 ? fcMax + ' bpm' : '-';

  // ── Outros ───────────────────────────────────────────────────────
  var elev    = Number(raw['Total Elevation Gain m']) || 0;
  var elevFmt = elev > 0 ? elev + ' m'   : '-';
  var cal     = Number(raw['Calories'])  || 0;
  var calFmt  = cal  > 0 ? cal + ' kcal': '-';
  var cad     = Number(raw['Average Cadence']) || 0;
  var cadFmt  = cad  > 0 ? cad + ' rpm' : '-';
  var pot     = Number(raw['Average Watts']) || Number(raw['Potência W']) || 0;
  var potFmt  = pot  > 0 ? pot + ' W'   : '-';

  // ── Classificações ───────────────────────────────────────────────
  var carga = calcularCargaSimples_(distKm, movS / 60, fcMed);
  var intensidade;
  if (paceSKm <= 0)        intensidade = '-';
  else if (paceSKm >= 480) intensidade = 'Leve';
  else if (paceSKm >= 390) intensidade = 'Moderado';
  else if (paceSKm >= 300) intensidade = 'Forte';
  else                     intensidade = 'Muito Forte';

  var qObj      = classificarQualidadeDado_(raw);
  var qualidade = qObj.qualidade;
  var flags     = qObj.flags.length > 0 ? qObj.flags.join('|') : 'DADO_MINIMO_OK';

  // ── Flags de aplicabilidade ───────────────────────────────────────
  var temDist = isAtividadeComDistancia_(tipoRaw)                 ? 'Sim' : 'Não';
  var temPace = (isAtividadeComDistancia_(tipoRaw) && avgMps > 0) ? 'Sim' : 'Não';
  var temFC   = fcMed > 0 ? 'Sim' : 'Não';
  var tipoReg = manual === 'Sim' ? 'Manual' : (gearId ? 'Dispositivo' : 'App');
  var possDis = gearId ? 'Sim' : 'Não';

  // ── Metadados ────────────────────────────────────────────────────
  var fonte    = 'STRAVA_RAW';
  var importEm = Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy HH:mm:ss');

  // ── Retorno: 49 valores na ordem exata dos headers CONVERTIDAS ────
  return [
    idInterno, athId, atleta, dataHoraRaw, dataFmt, horaFmt, ano, mes, mesRef,
    semana, diaSem, tipo, tipoRaw, stravaId, nome,
    distKm, distFmt, movS, movFmt, totS,
    totFmt, paceSKm, paceFmt, velKmh, velFmt, fcMed,
    fcMedFmt, fcMax, fcMaxFmt, elev, elevFmt, cal,
    calFmt, cad, cadFmt, pot, potFmt, carga,
    intensidade, qualidade, flags, tipoReg, possDis,
    temDist, temPace, temFC,
    fonte, importEm, 'ATIVO'
  ];
}
