import "jsr:@supabase/functions-js/edge-runtime.d.ts";

type UserRecord = { id?: string };
type UserTicket = {
  ok: boolean;
  reason?: string;
  state?: string;
  scope?: string;
  expires_at?: string;
  client_id?: string;
};

const SUPABASE_URL = Deno.env.get("SUPABASE_URL") ?? "";
const ANON_KEY = Deno.env.get("SUPABASE_ANON_KEY") ?? Deno.env.get("SB_PUBLISHABLE_KEY") ?? "";
const SERVICE_KEY = Deno.env.get("SUPABASE_SERVICE_ROLE_KEY") ?? Deno.env.get("SUPABASE_SECRET_KEY") ?? "";
const CALLBACK_URL = `${SUPABASE_URL}/functions/v1/strava-oauth`;
const STRAVA_AUTHORIZE_URL = "https://www.strava.com/oauth/authorize";
const REQUIRED_SCOPE = "read,activity:read_all";

const CORS_HEADERS = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, apikey, content-type, x-client-info",
  "Access-Control-Allow-Methods": "POST, OPTIONS",
  "Access-Control-Max-Age": "86400",
};

function json(body: unknown, status = 200): Response {
  return new Response(JSON.stringify(body), {
    status,
    headers: {
      ...CORS_HEADERS,
      "Content-Type": "application/json",
      "Cache-Control": "no-store",
      "X-Content-Type-Options": "nosniff",
    },
  });
}

function statusForReason(reason: string): number {
  switch (reason) {
    case "unauthenticated":
      return 401;
    case "access_context_missing":
    case "access_inactive":
      return 403;
    case "participant_not_found":
    case "identity_not_verified":
    case "v3_identity_missing":
      return 409;
    case "athlete_not_found":
      return 404;
    default:
      return 409;
  }
}

async function authenticatedUserId(req: Request): Promise<string | null> {
  const authorization = req.headers.get("Authorization") ?? "";
  if (!authorization.toLowerCase().startsWith("bearer ")) return null;
  if (!SUPABASE_URL || !ANON_KEY) throw new Error("Supabase public runtime config unavailable");

  const response = await fetch(`${SUPABASE_URL}/auth/v1/user`, {
    headers: { apikey: ANON_KEY, Authorization: authorization },
  });
  if (response.status === 401 || response.status === 403) return null;
  if (!response.ok) throw new Error(`auth_user_http_${response.status}`);

  const user = await response.json() as UserRecord;
  const id = String(user.id ?? "").trim();
  return id || null;
}

async function issueUserTicket(userId: string): Promise<UserTicket> {
  if (!SUPABASE_URL || !SERVICE_KEY) throw new Error("Supabase server runtime config unavailable");

  const response = await fetch(`${SUPABASE_URL}/rest/v1/rpc/strava_oauth_issue_user_ticket`, {
    method: "POST",
    headers: {
      apikey: SERVICE_KEY,
      Authorization: `Bearer ${SERVICE_KEY}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify({ p_auth_user_id: userId, p_ttl_seconds: 900 }),
  });

  const text = await response.text();
  if (!response.ok) {
    console.error("strava-connect user-ticket RPC failed", response.status);
    throw new Error(`user_ticket_http_${response.status}`);
  }

  try {
    return (text ? JSON.parse(text) : {}) as UserTicket;
  } catch {
    throw new Error("user_ticket_invalid_json");
  }
}

Deno.serve(async (req: Request) => {
  try {
    if (req.method === "OPTIONS") return new Response(null, { status: 204, headers: CORS_HEADERS });
    if (req.method !== "POST") return json({ ok: false, error: "method_not_allowed" }, 405);

    const userId = await authenticatedUserId(req);
    if (!userId) return json({ ok: false, error: "unauthenticated" }, 401);

    const ticket = await issueUserTicket(userId);
    if (!ticket.ok || !ticket.state) {
      const reason = ticket.reason ?? "ticket_failed";
      return json({ ok: false, error: reason }, statusForReason(reason));
    }

    const clientId = String(ticket.client_id ?? "").trim();
    if (!clientId) throw new Error("strava_client_id_missing");

    const authUrl = new URL(STRAVA_AUTHORIZE_URL);
    authUrl.searchParams.set("client_id", clientId);
    authUrl.searchParams.set("redirect_uri", CALLBACK_URL);
    authUrl.searchParams.set("response_type", "code");
    authUrl.searchParams.set("approval_prompt", "force");
    authUrl.searchParams.set("scope", ticket.scope || REQUIRED_SCOPE);
    authUrl.searchParams.set("state", ticket.state);

    return json({
      ok: true,
      authorization_url: authUrl.toString(),
      expires_at: ticket.expires_at,
    });
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    console.error("strava-connect", message.replace(/[A-Za-z0-9_-]{20,}/g, "[redacted]"));
    return json({ ok: false, error: "temporary_failure" }, 503);
  }
});
