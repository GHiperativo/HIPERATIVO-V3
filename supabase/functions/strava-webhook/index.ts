import "jsr:@supabase/functions-js/edge-runtime.d.ts";

type WebhookConfig = {
  id: number;
  verify_token: string;
  apps_script_secret: string;
  apps_script_url: string;
  callback_url: string;
  subscription_id: number | null;
  modo: "shadow" | "ativo" | "pausado";
};

type StravaEvent = {
  object_type: "activity" | "athlete";
  object_id: number;
  aspect_type: "create" | "update" | "delete";
  owner_id: number;
  subscription_id: number;
  event_time: number;
  updates?: Record<string, unknown>;
};

type StoredEvent = StravaEvent & {
  id: number;
  updates: Record<string, unknown>;
};

const SUPABASE_URL = Deno.env.get("SUPABASE_URL") ?? "";
const SERVICE_KEY = Deno.env.get("SUPABASE_SERVICE_ROLE_KEY") ??
  Deno.env.get("SUPABASE_SECRET_KEY") ?? "";

const BASE_HEADERS = {
  apikey: SERVICE_KEY,
  Authorization: `Bearer ${SERVICE_KEY}`,
  "Content-Type": "application/json",
};

function json(body: unknown, status = 200): Response {
  return new Response(JSON.stringify(body), {
    status,
    headers: { "Content-Type": "application/json", "Cache-Control": "no-store" },
  });
}

async function rest(path: string, init: RequestInit = {}): Promise<Response> {
  if (!SUPABASE_URL || !SERVICE_KEY) throw new Error("Supabase server credentials unavailable");
  return await fetch(`${SUPABASE_URL}/rest/v1/${path}`, {
    ...init,
    headers: { ...BASE_HEADERS, ...(init.headers ?? {}) },
  });
}

async function loadConfig(): Promise<WebhookConfig> {
  const response = await rest(
    "strava_webhook_config?id=eq.1&select=id,verify_token,apps_script_secret,apps_script_url,callback_url,subscription_id,modo&limit=1",
  );
  if (!response.ok) throw new Error(`config HTTP ${response.status}`);
  const rows = await response.json() as WebhookConfig[];
  if (!rows.length) throw new Error("webhook config not initialized");
  return rows[0];
}

function validEvent(value: unknown): value is StravaEvent {
  if (!value || typeof value !== "object") return false;
  const event = value as Record<string, unknown>;
  return (event.object_type === "activity" || event.object_type === "athlete") &&
    (event.aspect_type === "create" || event.aspect_type === "update" || event.aspect_type === "delete") &&
    Number.isSafeInteger(event.object_id) && Number(event.object_id) > 0 &&
    Number.isSafeInteger(event.owner_id) && Number(event.owner_id) > 0 &&
    Number.isSafeInteger(event.subscription_id) && Number(event.subscription_id) >= 0 &&
    Number.isSafeInteger(event.event_time) && Number(event.event_time) > 0;
}

function recentEvent(eventTime: number): boolean {
  const now = Math.floor(Date.now() / 1000);
  return Math.abs(now - eventTime) <= 3600;
}

async function hmacHex(secret: string, message: string): Promise<string> {
  const key = await crypto.subtle.importKey(
    "raw",
    new TextEncoder().encode(secret),
    { name: "HMAC", hash: "SHA-256" },
    false,
    ["sign"],
  );
  const signature = await crypto.subtle.sign("HMAC", key, new TextEncoder().encode(message));
  return Array.from(new Uint8Array(signature))
    .map((byte) => byte.toString(16).padStart(2, "0"))
    .join("");
}

async function updateEvent(id: number, values: Record<string, unknown>): Promise<void> {
  const response = await rest(`strava_eventos_webhook?id=eq.${id}`, {
    method: "PATCH",
    headers: { Prefer: "return=minimal" },
    body: JSON.stringify(values),
  });
  if (!response.ok) throw new Error(`event update HTTP ${response.status}`);
}

async function dispatchToAppsScript(event: StoredEvent, config: WebhookConfig): Promise<void> {
  try {
    await updateEvent(event.id, {
      status: "processando",
      tentativas: 1,
      ultimo_erro: null,
      proxima_tentativa: null,
    });

    const timestamp = Math.floor(Date.now() / 1000);
    const nonce = crypto.randomUUID();
    const eventPayload = {
      queue_id: event.id,
      object_type: event.object_type,
      object_id: event.object_id,
      aspect_type: event.aspect_type,
      owner_id: event.owner_id,
      subscription_id: event.subscription_id,
      event_time: event.event_time,
      updates: event.updates ?? {},
    };
    const canonical = `${timestamp}.${nonce}.${JSON.stringify(eventPayload)}`;
    const signature = await hmacHex(config.apps_script_secret, canonical);
    const response = await fetch(config.apps_script_url, {
      method: "POST",
      redirect: "follow",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        source: "supabase-strava-webhook",
        timestamp,
        nonce,
        signature,
        event: eventPayload,
      }),
    });
    const text = await response.text();
    let result: Record<string, unknown> = {};
    try {
      result = JSON.parse(text) as Record<string, unknown>;
    } catch {
      throw new Error(`Apps Script returned non-JSON HTTP ${response.status}`);
    }
    if (!response.ok || result.ok !== true) {
      throw new Error(String(result.erro ?? `Apps Script HTTP ${response.status}`));
    }
    await updateEvent(event.id, {
      status: "processado",
      processado_at: new Date().toISOString(),
      ultimo_erro: null,
      proxima_tentativa: null,
    });
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    try {
      await updateEvent(event.id, {
        status: "falha",
        ultimo_erro: message.slice(0, 1000),
        proxima_tentativa: new Date(Date.now() + 15 * 60 * 1000).toISOString(),
      });
    } catch (updateError) {
      console.error("failed to persist webhook error", updateError);
    }
  }
}

async function receiveEvent(event: StravaEvent, config: WebhookConfig): Promise<StoredEvent | null> {
  const payload = {
    subscription_id: event.subscription_id,
    object_type: event.object_type,
    object_id: event.object_id,
    aspect_type: event.aspect_type,
    owner_id: event.owner_id,
    updates: event.updates ?? {},
    event_time: event.event_time,
    payload: event,
    status: config.modo === "ativo" ? "recebido" : "espelho",
  };
  const conflict = "subscription_id,object_type,object_id,aspect_type,event_time";
  const response = await rest(`strava_eventos_webhook?on_conflict=${conflict}`, {
    method: "POST",
    headers: { Prefer: "resolution=ignore-duplicates,return=representation" },
    body: JSON.stringify(payload),
  });
  if (!response.ok) throw new Error(`event insert HTTP ${response.status}: ${(await response.text()).slice(0, 300)}`);
  const rows = await response.json() as StoredEvent[];
  return rows[0] ?? null;
}

Deno.serve(async (req: Request) => {
  try {
    const config = await loadConfig();
    const url = new URL(req.url);

    if (req.method === "GET") {
      const mode = url.searchParams.get("hub.mode");
      const token = url.searchParams.get("hub.verify_token");
      const challenge = url.searchParams.get("hub.challenge");
      if (mode !== "subscribe" || !challenge || token !== config.verify_token) {
        return json({ erro: "verification rejected" }, 403);
      }
      return json({ "hub.challenge": challenge });
    }

    if (req.method !== "POST") return json({ erro: "method not allowed" }, 405);
    const contentLength = Number(req.headers.get("content-length") ?? 0);
    if (contentLength > 16384) return json({ erro: "payload too large" }, 413);

    const event = await req.json() as unknown;
    if (!validEvent(event)) return json({ erro: "invalid event" }, 400);
    if (!recentEvent(event.event_time)) return json({ erro: "stale event" }, 400);
    if (config.subscription_id !== null && event.subscription_id !== config.subscription_id) {
      return json({ erro: "unknown subscription" }, 403);
    }

    const stored = await receiveEvent(event, config);
    if (stored && config.modo === "ativo") {
      EdgeRuntime.waitUntil(dispatchToAppsScript(stored, config));
    }
    return json({ ok: true, accepted: Boolean(stored), mode: config.modo });
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    console.error("strava-webhook", message);
    return json({ erro: "temporary webhook failure" }, 503);
  }
});
