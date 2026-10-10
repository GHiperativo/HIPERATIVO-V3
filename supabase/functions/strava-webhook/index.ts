import "jsr:@supabase/functions-js/edge-runtime.d.ts";
import {
  activityToDb,
  deleteActivity,
  getTokenByOwner,
  json,
  loadConfig,
  markTokenRevoked,
  rest,
  stravaGet,
  upsertActivities,
  type WebhookConfig,
} from "../_shared/strava.ts";

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
  tentativas?: number;
};

function validEvent(value: unknown): value is StravaEvent {
  if (!value || typeof value !== "object") return false;
  const event = value as Record<string, unknown>;
  return (event.object_type === "activity" || event.object_type === "athlete") &&
    (event.aspect_type === "create" || event.aspect_type === "update" || event.aspect_type === "delete") &&
    Number.isSafeInteger(event.object_id) && Number(event.object_id) > 0 &&
    Number.isSafeInteger(event.owner_id) && Number(event.owner_id) > 0 &&
    Number.isSafeInteger(event.subscription_id) && Number(event.subscription_id) > 0 &&
    Number.isSafeInteger(event.event_time) && Number(event.event_time) > 0;
}

function recentEvent(eventTime: number): boolean {
  const now = Math.floor(Date.now() / 1000);
  return Math.abs(now - eventTime) <= 3600;
}

async function updateEvent(id: number, values: Record<string, unknown>): Promise<void> {
  const response = await rest(`strava_eventos_webhook?id=eq.${id}`, {
    method: "PATCH",
    headers: { Prefer: "return=minimal" },
    body: JSON.stringify(values),
  });
  if (!response.ok) throw new Error(`event update HTTP ${response.status}`);
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
  if (!response.ok) throw new Error(`event insert HTTP ${response.status}: ${(await response.text()).slice(0, 220)}`);
  const rows = await response.json() as StoredEvent[];
  return rows[0] ?? null;
}

async function processEvent(event: StoredEvent): Promise<void> {
  const attempts = Number(event.tentativas ?? 0) + 1;
  try {
    await updateEvent(event.id, {
      status: "processando",
      tentativas: attempts,
      ultimo_erro: null,
      proxima_tentativa: null,
    });

    const token = await getTokenByOwner(event.owner_id);
    if (!token) throw new Error(`unknown Strava owner ${event.owner_id}`);

    if (event.object_type === "athlete") {
      const authorized = event.updates?.authorized;
      if (authorized === false || authorized === "false") await markTokenRevoked(token);
    } else if (event.aspect_type === "delete") {
      await deleteActivity(String(event.object_id), token.ath_id);
    } else {
      const { response, body } = await stravaGet(`/activities/${event.object_id}`, token);
      if (response.status === 404) {
        await deleteActivity(String(event.object_id), token.ath_id);
      } else if (!response.ok || !body || typeof body !== "object") {
        throw new Error(`Strava activity HTTP ${response.status}`);
      } else {
        const activity = body as Record<string, unknown>;
        const athlete = activity.athlete as Record<string, unknown> | undefined;
        if (Number(athlete?.id ?? 0) !== event.owner_id) throw new Error("activity owner mismatch");
        await upsertActivities([activityToDb(activity, token)]);
      }
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
        tentativas: attempts,
        ultimo_erro: message.slice(0, 1000),
        proxima_tentativa: new Date(Date.now() + 15 * 60 * 1000).toISOString(),
      });
    } catch (updateError) {
      console.error("failed to persist webhook processing error", updateError);
    }
  }
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
    if (stored && config.modo === "ativo") EdgeRuntime.waitUntil(processEvent(stored));
    return json({ ok: true, accepted: Boolean(stored), mode: config.modo });
  } catch (error) {
    console.error("strava-webhook", error instanceof Error ? error.message : String(error));
    return json({ erro: "temporary webhook failure" }, 503);
  }
});
