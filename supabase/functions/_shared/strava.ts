import "jsr:@supabase/functions-js/edge-runtime.d.ts";

export type TokenRow = {
  ath_id: string;
  nome: string | null;
  access_token?: string | null;
  refresh_token?: string | null;
  expires_at: number | null;
  scope: string | null;
  strava_id: string | null;
  ult_atu: string | null;
  status: string | null;
  token_version?: number | null;
};

type RefreshClaim = {
  claimed: boolean;
  reason?: string;
  lease_id?: string;
  refresh_token?: string;
  token_version?: number;
  retry_after_seconds?: number;
};

export type WebhookConfig = {
  id: number;
  verify_token: string;
  callback_url: string;
  subscription_id: number | null;
  modo: "shadow" | "ativo" | "pausado";
};

const SUPABASE_URL = Deno.env.get("SUPABASE_URL") ?? "";
const SERVICE_KEY = Deno.env.get("SUPABASE_SERVICE_ROLE_KEY") ?? Deno.env.get("SUPABASE_SECRET_KEY") ?? "";
const TOKEN_META_SELECT = "ath_id,nome,expires_at,scope,strava_id,ult_atu,status,token_version";

const BASE_HEADERS = {
  apikey: SERVICE_KEY,
  Authorization: `Bearer ${SERVICE_KEY}`,
  "Content-Type": "application/json",
};

export function json(body: unknown, status = 200): Response {
  return new Response(JSON.stringify(body), {
    status,
    headers: { "Content-Type": "application/json", "Cache-Control": "no-store" },
  });
}

export async function rest(path: string, init: RequestInit = {}): Promise<Response> {
  if (!SUPABASE_URL || !SERVICE_KEY) throw new Error("Supabase server credentials unavailable");
  return await fetch(`${SUPABASE_URL}/rest/v1/${path}`, {
    ...init,
    headers: { ...BASE_HEADERS, ...(init.headers ?? {}) },
  });
}

export async function rpc<T = unknown>(name: string, body: Record<string, unknown> = {}): Promise<T> {
  const response = await rest(`rpc/${name}`, { method: "POST", body: JSON.stringify(body) });
  const text = await response.text();
  if (!response.ok) throw new Error(`RPC ${name} HTTP ${response.status}: ${text.slice(0, 220)}`);
  return (text ? JSON.parse(text) : null) as T;
}

export async function loadConfig(): Promise<WebhookConfig> {
  const response = await rest("strava_webhook_config?id=eq.1&select=id,verify_token,callback_url,subscription_id,modo&limit=1");
  if (!response.ok) throw new Error(`config HTTP ${response.status}`);
  const rows = await response.json() as WebhookConfig[];
  if (!rows.length) throw new Error("webhook config not initialized");
  return rows[0];
}

export async function runtimeCredentials(): Promise<{client_id: string; client_secret: string}> {
  const config = await rpc<{client_id?: string; client_secret?: string}>("strava_runtime_config");
  const client_id = String(config?.client_id ?? "");
  const client_secret = String(config?.client_secret ?? "");
  if (!client_id || !client_secret) throw new Error("Strava runtime credentials unavailable in Vault");
  return { client_id, client_secret };
}

export async function internalAuthorized(req: Request): Promise<boolean> {
  const token = req.headers.get("x-strava-internal-secret") ?? "";
  if (!token) return false;
  return await rpc<boolean>("strava_internal_token_matches", { p_token: token });
}

export async function getTokenByOwner(ownerId: number): Promise<TokenRow | null> {
  const response = await rest(`tokens_strava?strava_id=eq.${encodeURIComponent(String(ownerId))}&select=${TOKEN_META_SELECT}&limit=1`);
  if (!response.ok) throw new Error(`token lookup HTTP ${response.status}`);
  const rows = await response.json() as TokenRow[];
  return rows[0] ?? null;
}

export async function listTokens(): Promise<TokenRow[]> {
  const response = await rest(`tokens_strava?select=${TOKEN_META_SELECT}&order=ath_id.asc`);
  if (!response.ok) throw new Error(`tokens list HTTP ${response.status}`);
  return await response.json() as TokenRow[];
}

async function getTokenCredentials(athId: string): Promise<TokenRow> {
  const row = await rpc<TokenRow | null>("strava_token_get_access", { p_ath_id: athId });
  if (!row?.ath_id) throw new Error(`token row unavailable for ${athId}`);
  return row;
}

async function patchToken(athId: string, values: Record<string, unknown>): Promise<void> {
  const response = await rest(`tokens_strava?ath_id=eq.${encodeURIComponent(athId)}`, {
    method: "PATCH",
    headers: { Prefer: "return=minimal" },
    body: JSON.stringify(values),
  });
  if (!response.ok) throw new Error(`token update HTTP ${response.status}`);
}

function copyTokenState(target: TokenRow, source: TokenRow): void {
  target.access_token = source.access_token ?? null;
  target.expires_at = source.expires_at;
  target.status = source.status;
  target.ult_atu = source.ult_atu;
  target.token_version = source.token_version;
}

function sleep(ms: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

async function waitForConcurrentRefresh(row: TokenRow): Promise<string | null> {
  const initialVersion = Number(row.token_version ?? 0);
  for (let attempt = 0; attempt < 5; attempt++) {
    await sleep(400);
    const fresh = await getTokenCredentials(row.ath_id);
    const now = Math.floor(Date.now() / 1000);
    if (Number(fresh.token_version ?? 0) > initialVersion && fresh.access_token && Number(fresh.expires_at ?? 0) > now + 60) {
      copyTokenState(row, fresh);
      return fresh.access_token;
    }
  }
  return null;
}

async function commitRefreshWithRecovery(
  row: TokenRow,
  claim: RefreshClaim,
  access: string,
  refresh: string,
  expires: number,
): Promise<void> {
  const leaseId = String(claim.lease_id ?? "");
  const previousVersion = Number(claim.token_version ?? row.token_version ?? 0);
  let lastError = "";

  for (let attempt = 0; attempt < 3; attempt++) {
    try {
      const committed = await rpc<boolean>("strava_token_commit_refresh", {
        p_ath_id: row.ath_id,
        p_lease_id: leaseId,
        p_access_token: access,
        p_refresh_token: refresh,
        p_expires_at: expires,
      });
      if (committed) return;
    } catch (error) {
      lastError = error instanceof Error ? error.message : String(error);
    }

    const current = await getTokenCredentials(row.ath_id).catch(() => null);
    if (current && Number(current.token_version ?? 0) > previousVersion && Number(current.expires_at ?? 0) >= expires - 5) {
      copyTokenState(row, current);
      return;
    }
    await sleep(300 * (attempt + 1));
  }

  throw new Error(`token refresh commit failed${lastError ? `: ${lastError.slice(0, 120)}` : ""}`);
}

export async function validAccessToken(row: TokenRow, force = false): Promise<string> {
  const current = await getTokenCredentials(row.ath_id);
  copyTokenState(row, current);
  const now = Math.floor(Date.now() / 1000);

  // Strava documents a one-hour refresh window. Staying ahead of it avoids an expiry gap.
  if (!force && current.access_token && Number(current.expires_at ?? 0) > now + 3600) return current.access_token;

  const claim = await rpc<RefreshClaim>("strava_token_claim_refresh", {
    p_ath_id: row.ath_id,
    p_lease_seconds: 90,
  });

  if (!claim.claimed) {
    if (claim.reason === "busy") {
      const concurrent = await waitForConcurrentRefresh(row);
      if (concurrent) return concurrent;
    }
    throw new Error(`refresh unavailable for ${row.ath_id}: ${claim.reason ?? "unknown"}`);
  }

  const leaseId = String(claim.lease_id ?? "");
  const refreshToken = String(claim.refresh_token ?? "");
  if (!leaseId || refreshToken.length < 10) {
    await rpc<boolean>("strava_token_fail_refresh", {
      p_ath_id: row.ath_id,
      p_lease_id: leaseId,
      p_error: "invalid refresh lease payload",
    }).catch(() => false);
    throw new Error(`refresh lease invalid for ${row.ath_id}`);
  }

  try {
    const cred = await runtimeCredentials();
    const form = new URLSearchParams({
      client_id: cred.client_id,
      client_secret: cred.client_secret,
      grant_type: "refresh_token",
      refresh_token: refreshToken,
    });
    const response = await fetch("https://www.strava.com/oauth/token", {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: form.toString(),
    });
    const text = await response.text();
    let body: Record<string, unknown> = {};
    try { body = JSON.parse(text) as Record<string, unknown>; } catch { /* no-op */ }
    if (!response.ok || !body.access_token || !body.expires_at) {
      throw new Error(`Strava token refresh HTTP ${response.status}: ${String(body.message ?? body.error ?? "").slice(0, 140)}`);
    }

    const access = String(body.access_token);
    const refresh = String(body.refresh_token ?? refreshToken);
    const expires = Number(body.expires_at);
    await commitRefreshWithRecovery(row, claim, access, refresh, expires);

    row.access_token = access;
    row.refresh_token = refresh;
    row.expires_at = expires;
    row.status = "Renovado";
    row.ult_atu = new Date().toISOString();
    row.token_version = Number(claim.token_version ?? 0) + 1;
    return access;
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    await rpc<boolean>("strava_token_fail_refresh", {
      p_ath_id: row.ath_id,
      p_lease_id: leaseId,
      p_error: message.slice(0, 500),
    }).catch(() => false);
    throw error;
  }
}

export async function stravaGet(path: string, row: TokenRow): Promise<{response: Response; body: unknown}> {
  let access = await validAccessToken(row);
  let response = await fetch(`https://www.strava.com/api/v3${path}`, { headers: { Authorization: `Bearer ${access}` } });
  if (response.status === 401) {
    access = await validAccessToken(row, true);
    response = await fetch(`https://www.strava.com/api/v3${path}`, { headers: { Authorization: `Bearer ${access}` } });
  }
  const text = await response.text();
  let body: unknown = null;
  try { body = text ? JSON.parse(text) : null; } catch { body = text; }
  return { response, body };
}

const SPORT_TYPE: Record<string, string> = {
  AlpineSki:"Esqui alpino",BackcountrySki:"Esqui fora de pista",Badminton:"Badminton",Basketball:"Basquete",Canoeing:"Canoagem",Cricket:"Críquete",Crossfit:"Crossfit",Dance:"Dança",EBikeRide:"Ciclismo elétrico",Elliptical:"Elíptico",EMountainBikeRide:"Mountain bike elétrico",Golf:"Golfe",GravelRide:"Gravel",Handcycle:"Handcycle",HighIntensityIntervalTraining:"HIIT",Hike:"Trilha",IceSkate:"Patinação no gelo",InlineSkate:"Patins",Kayaking:"Caiaque",Kitesurf:"Kitesurf",MountainBikeRide:"Mountain bike",NordicSki:"Esqui nórdico",Padel:"Padel",PhysicalTherapy:"Fisioterapia",Pickleball:"Pickleball",Pilates:"Pilates",Racquetball:"Raquetebol",Run:"Corrida",TrailRun:"Corrida em trilha",Walk:"Caminhada",Ride:"Ciclismo",VirtualRide:"Ciclismo virtual",RockClimbing:"Escalada",RollerSki:"Esqui sobre rodas",Rowing:"Remo",Sail:"Vela",Skateboard:"Skate",Snowboard:"Snowboard",Snowshoe:"Caminhada com raquetes de neve",Soccer:"Futebol",Squash:"Squash",StairStepper:"Escada",StandUpPaddling:"Stand up paddle",Surfing:"Surfe",Swim:"Natação",TableTennis:"Tênis de mesa",Tennis:"Tênis",Velomobile:"Velomóvel",VirtualRow:"Remo virtual",VirtualRun:"Corrida virtual",Volleyball:"Vôlei",WeightTraining:"Musculação",Wheelchair:"Cadeira de rodas",Windsurf:"Windsurf",Workout:"Treino",Yoga:"Yoga",
};
const RUN_TYPES = new Set(["Run","TrailRun","VirtualRun","Walk","Hike"]);
const CYCLE_TYPES = new Set(["Ride","VirtualRide","MountainBikeRide","EMountainBikeRide","GravelRide","EBikeRide","Handcycle","Velomobile"]);

function translatedName(value: unknown): string {
  const raw = String(value ?? "").trim();
  const match = raw.match(/^(Morning|Lunch|Afternoon|Evening|Night) (Run|Trail Run|Virtual Run|Ride|Virtual Ride|Mountain Bike Ride|Walk|Hike|Swim|Workout|Weight Training|Row|Yoga)$/);
  if (!match) return raw;
  const acts: Record<string,string> = {"Run":"Corrida","Trail Run":"Corrida em trilha","Virtual Run":"Corrida virtual","Ride":"Pedalada","Virtual Ride":"Pedalada virtual","Mountain Bike Ride":"Pedalada de mountain bike","Walk":"Caminhada","Hike":"Trilha","Swim":"Natação","Workout":"Treino","Weight Training":"Musculação","Row":"Remo","Yoga":"Yoga"};
  const periods: Record<string,string> = {Morning:"matinal",Lunch:"na hora do almoço",Afternoon:"à tarde",Evening:"ao entardecer",Night:"noturna"};
  return `${acts[match[2]] ?? match[2]} ${periods[match[1]] ?? ""}`.trim();
}

function numberOrNull(value: unknown): number | null {
  const n = Number(value);
  return Number.isFinite(n) ? n : null;
}
function intOrNull(value: unknown): number | null {
  const n = numberOrNull(value);
  return n === null ? null : Math.round(n);
}
function paceFormat(seconds: number): string {
  const minutes = Math.floor(seconds / 60);
  const rest = Math.max(0, seconds - minutes * 60);
  return `${minutes}:${String(rest).padStart(2, "0")} /km`;
}

export function activityToDb(activity: Record<string, unknown>, row: TokenRow): Record<string, unknown> {
  const original = String(activity.sport_type ?? activity.type ?? "");
  const speed = numberOrNull(activity.average_speed) ?? 0;
  const pace = RUN_TYPES.has(original) && speed > 0 ? Math.round(1000 / speed) : 0;
  const pace_fmt = pace > 0 ? paceFormat(pace) : (CYCLE_TYPES.has(original) && speed > 0 ? `${(speed * 3.6).toFixed(1)} km/h` : "");
  const distance = numberOrNull(activity.distance) ?? 0;
  return {
    ath_id: row.ath_id,
    nome: row.nome ?? "",
    data: String(activity.start_date ?? activity.start_date_local ?? ""),
    tipo: SPORT_TYPE[original] ?? original ?? "Outro",
    strava_id: String(activity.id ?? ""),
    nome_ativ: translatedName(activity.name),
    mov_s: intOrNull(activity.moving_time),
    total_s: intOrNull(activity.elapsed_time),
    dist_m: intOrNull(distance),
    dist_km: Math.round((distance / 1000) * 1000) / 1000,
    vel_mps: Math.round(speed * 1000) / 1000,
    pace_s_km: pace,
    pace_fmt,
    fc_med: intOrNull(activity.average_heartrate),
    fc_max: intOrNull(activity.max_heartrate),
    elev: intOrNull(activity.total_elevation_gain),
    cal: intOrNull(activity.calories),
    cadencia: intOrNull(activity.average_cadence),
  };
}

export async function upsertActivities(items: Record<string, unknown>[]): Promise<Record<string, unknown>> {
  if (!items.length) return { inserted: 0, updated: 0, identity_conflicts: 0 };
  return await rpc<Record<string, unknown>>("strava_upsert_activities", { p_items: items });
}

export async function deleteActivity(stravaId: string, athId: string): Promise<number> {
  return await rpc<number>("strava_delete_activity", { p_strava_id: stravaId, p_ath_id: athId });
}

export async function markTokenRevoked(row: TokenRow): Promise<void> {
  await patchToken(row.ath_id, { status: "Revogado pela Strava", ult_atu: new Date().toISOString() });
}

export function rateLimit(response: Response): {short?: number; day?: number} {
  const parts = (response.headers.get("X-RateLimit-Usage") ?? "").split(",").map(Number);
  return { short: Number.isFinite(parts[0]) ? parts[0] : undefined, day: Number.isFinite(parts[1]) ? parts[1] : undefined };
}
