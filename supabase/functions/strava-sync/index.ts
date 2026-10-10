import "jsr:@supabase/functions-js/edge-runtime.d.ts";
import {
  activityToDb,
  internalAuthorized,
  json,
  listTokens,
  loadConfig,
  rateLimit,
  rest,
  runtimeCredentials,
  stravaGet,
  upsertActivities,
  validAccessToken,
} from "../_shared/strava.ts";

async function updateConfig(values: Record<string, unknown>): Promise<void> {
  const response = await rest("strava_webhook_config?id=eq.1", {
    method: "PATCH",
    headers: { Prefer: "return=minimal" },
    body: JSON.stringify({ ...values, atualizado_at: new Date().toISOString() }),
  });
  if (!response.ok) throw new Error(`config update HTTP ${response.status}`);
}

async function subscription(): Promise<Response> {
  const credentials = await runtimeCredentials();
  const config = await loadConfig();
  const query = new URLSearchParams({ client_id: credentials.client_id, client_secret: credentials.client_secret });
  const response = await fetch(`https://www.strava.com/api/v3/push_subscriptions?${query.toString()}`);
  const text = await response.text();
  if (!response.ok) return json({ ok: false, stage: "list", http: response.status }, 502);
  const list = JSON.parse(text || "[]") as Array<Record<string, unknown>>;
  const expected = config.callback_url.replace(/\/$/, "");

  if (list.length > 1) return json({ ok: false, reason: "multiple_subscriptions", count: list.length }, 409);
  if (list.length === 1) {
    const current = list[0];
    const callback = String(current.callback_url ?? "").replace(/\/$/, "");
    if (callback !== expected) {
      return json({ ok: false, reason: "unknown_callback_preserved", subscription_id: current.id, callback_url: callback }, 409);
    }
    const id = Number(current.id);
    await updateConfig({ subscription_id: id, modo: "ativo" });
    return json({ ok: true, created: false, subscription_id: id, callback_url: callback });
  }

  const form = new URLSearchParams({
    client_id: credentials.client_id,
    client_secret: credentials.client_secret,
    callback_url: config.callback_url,
    verify_token: config.verify_token,
  });
  const create = await fetch("https://www.strava.com/api/v3/push_subscriptions", {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: form.toString(),
  });
  const createText = await create.text();
  let body: Record<string, unknown> = {};
  try { body = JSON.parse(createText) as Record<string, unknown>; } catch { /* no-op */ }
  if (!create.ok || !body.id) {
    return json({ ok: false, stage: "create", http: create.status, message: String(body.message ?? body.error ?? "").slice(0, 160) }, 502);
  }
  const id = Number(body.id);
  await updateConfig({ subscription_id: id, modo: "ativo" });
  return json({ ok: true, created: true, subscription_id: id, callback_url: config.callback_url });
}

type SyncOptions = { after?: number; before?: number; max_pages?: number; per_page?: number; athlete_limit?: number };

function activeTokens(tokens: Awaited<ReturnType<typeof listTokens>>, athleteLimit = 1000) {
  return tokens
    .filter((token) => String(token.status ?? "").toLowerCase() !== "inativo" && !String(token.status ?? "").toLowerCase().startsWith("revogado"))
    .slice(0, Math.max(1, Math.min(Number(athleteLimit || 1000), 1000)));
}

async function maintainTokens(athleteLimit = 1000): Promise<Response> {
  const tokens = activeTokens(await listTokens(), athleteLimit);
  const result = {
    ok: true,
    athletes_total: tokens.length,
    athletes_ok: 0,
    athletes_failed: 0,
    refreshed: 0,
    already_valid: 0,
    errors: [] as Array<{ath_id: string; error: string}>,
  };

  for (const token of tokens) {
    const beforeVersion = Number(token.token_version ?? 0);
    try {
      await validAccessToken(token);
      if (Number(token.token_version ?? 0) > beforeVersion) result.refreshed++;
      else result.already_valid++;
      result.athletes_ok++;
    } catch (error) {
      result.athletes_failed++;
      result.errors.push({ ath_id: token.ath_id, error: (error instanceof Error ? error.message : String(error)).slice(0, 220) });
    }
  }

  if (result.athletes_failed > 0) result.ok = false;
  return json(result, 200);
}

async function syncActivities(options: SyncOptions): Promise<Response> {
  const after = Number.isFinite(options.after) ? Number(options.after) : Math.floor(Date.now() / 1000) - 2 * 86400;
  const before = Number.isFinite(options.before) ? Number(options.before) : Math.floor(Date.now() / 1000) + 60;
  const maxPages = Math.max(1, Math.min(Number(options.max_pages ?? 3), 10));
  const perPage = Math.max(1, Math.min(Number(options.per_page ?? 100), 100));
  const athleteLimit = Math.max(1, Math.min(Number(options.athlete_limit ?? 1000), 1000));
  const tokens = activeTokens(await listTokens(), athleteLimit);

  const result = {
    ok: true,
    after,
    before,
    athletes_total: tokens.length,
    athletes_ok: 0,
    athletes_failed: 0,
    pages: 0,
    activities_received: 0,
    inserted: 0,
    updated: 0,
    identity_conflicts: 0,
    rate_limit_stop: false,
    errors: [] as Array<{ath_id: string; error: string}>,
  };

  for (const token of tokens) {
    try {
      await validAccessToken(token);
      for (let page = 1; page <= maxPages; page++) {
        const path = `/athlete/activities?after=${after}&before=${before}&per_page=${perPage}&page=${page}`;
        const { response, body } = await stravaGet(path, token);
        const usage = rateLimit(response);
        if (response.status === 429) {
          result.rate_limit_stop = true;
          throw new Error("Strava rate limit HTTP 429");
        }
        if (!response.ok || !Array.isArray(body)) throw new Error(`activities HTTP ${response.status}`);

        result.pages++;
        result.activities_received += body.length;
        const stored = await upsertActivities(body.map((activity) => activityToDb(activity as Record<string, unknown>, token)));
        result.inserted += Number(stored.inserted ?? 0);
        result.updated += Number(stored.updated ?? 0);
        result.identity_conflicts += Number(stored.identity_conflicts ?? 0);

        if (body.length < perPage) break;
        if ((usage.short ?? 0) >= 180) {
          result.rate_limit_stop = true;
          break;
        }
      }
      result.athletes_ok++;
    } catch (error) {
      result.athletes_failed++;
      result.errors.push({ ath_id: token.ath_id, error: (error instanceof Error ? error.message : String(error)).slice(0, 220) });
      if (result.rate_limit_stop) break;
    }
  }

  if (result.athletes_failed > 0 || result.identity_conflicts > 0) result.ok = false;
  return json(result, result.rate_limit_stop ? 429 : 200);
}

async function health(): Promise<Response> {
  const config = await loadConfig();
  let credentials = false;
  try { await runtimeCredentials(); credentials = true; } catch { credentials = false; }
  const tokens = await listTokens();
  const now = Math.floor(Date.now() / 1000);
  const auditResponse = await rest("tokens_strava?select=last_refresh_ok_at,last_refresh_error_at,refresh_lease_until");
  const auditRows = auditResponse.ok ? await auditResponse.json() as Array<Record<string, unknown>> : [];
  const nowMs = Date.now();
  const errors24h = auditRows.filter((row) => {
    const errorAt = Date.parse(String(row.last_refresh_error_at ?? ""));
    const okAt = Date.parse(String(row.last_refresh_ok_at ?? ""));
    return Number.isFinite(errorAt) && nowMs - errorAt <= 86400000 && (!Number.isFinite(okAt) || errorAt > okAt);
  }).length;
  const activeLeases = auditRows.filter((row) => {
    const until = Date.parse(String(row.refresh_lease_until ?? ""));
    return Number.isFinite(until) && until > nowMs;
  }).length;
  const active = activeTokens(tokens).length;
  const expiring1h = activeTokens(tokens).filter((token) => Number(token.expires_at ?? 0) <= now + 3600).length;
  const ok = credentials && config.subscription_id !== null;

  return json({
    ok,
    credentials_configured: credentials,
    subscription_id: config.subscription_id,
    mode: config.modo,
    tokens: tokens.length,
    active_tokens: active,
    expired_tokens: tokens.filter((token) => Number(token.expires_at ?? 0) <= now).length,
    expiring_within_1h: expiring1h,
    inactive_tokens: tokens.filter((token) => String(token.status ?? "").toLowerCase() === "inativo").length,
    refresh_errors_24h: errors24h,
    active_refresh_leases: activeLeases,
  }, ok ? 200 : 503);
}

Deno.serve(async (req: Request) => {
  try {
    if (req.method !== "POST") return json({ erro: "method not allowed" }, 405);
    if (!(await internalAuthorized(req))) return json({ erro: "unauthorized" }, 401);
    const body = await req.json().catch(() => ({})) as Record<string, unknown>;
    const action = String(body.action ?? "health");

    if (action === "health") return await health();
    if (action === "subscription") return await subscription();
    if (action === "refresh_tokens") return await maintainTokens(Number(body.athlete_limit ?? 1000));
    if (action === "sync" || action === "backfill") {
      const afterValue = body.after;
      const beforeValue = body.before;
      const after = typeof afterValue === "string" ? Math.floor(new Date(afterValue).getTime() / 1000) : Number(afterValue);
      const before = typeof beforeValue === "string" ? Math.floor(new Date(beforeValue).getTime() / 1000) : Number(beforeValue);
      return await syncActivities({
        after: Number.isFinite(after) ? after : undefined,
        before: Number.isFinite(before) ? before : undefined,
        max_pages: Number(body.max_pages ?? 3),
        per_page: Number(body.per_page ?? 100),
        athlete_limit: Number(body.athlete_limit ?? 1000),
      });
    }
    return json({ erro: "unknown action" }, 400);
  } catch (error) {
    console.error("strava-sync", error instanceof Error ? error.message : String(error));
    return json({ erro: "temporary sync failure" }, 503);
  }
});
