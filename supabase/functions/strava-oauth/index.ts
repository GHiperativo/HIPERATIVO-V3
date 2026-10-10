import "jsr:@supabase/functions-js/edge-runtime.d.ts";
import { internalAuthorized, json, rpc, runtimeCredentials } from "../_shared/strava.ts";

type TicketResponse = {
  ok: boolean;
  reason?: string;
  state?: string;
  scope?: string;
  expires_at?: string;
};

type ConsumeResponse = {
  ok: boolean;
  reason?: string;
  ath_id?: string;
  scope?: string;
};

type OAuthCommitResponse = {
  ok: boolean;
  ath_id?: string;
  strava_id?: string;
  expires_at?: number;
  scope?: string;
  storage?: string;
};

type StravaTokenResponse = {
  access_token?: string;
  refresh_token?: string;
  expires_at?: number;
  scope?: string;
  athlete?: { id?: number | string };
};

const SUPABASE_URL = Deno.env.get("SUPABASE_URL") ?? "";
const CALLBACK_URL = `${SUPABASE_URL}/functions/v1/strava-oauth`;
const STRAVA_AUTHORIZE_URL = "https://www.strava.com/oauth/authorize";
const STRAVA_TOKEN_URL = "https://www.strava.com/oauth/token";
const REQUIRED_SCOPE = "read,activity:read_all";

function html(title: string, message: string, status = 200): Response {
  const safeTitle = title.replace(/[<>&]/g, "");
  const safeMessage = message.replace(/[<>&]/g, "");
  const body = `<!doctype html><html lang="pt-BR"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><meta name="robots" content="noindex,nofollow"><title>${safeTitle}</title><style>body{font-family:system-ui,-apple-system,sans-serif;max-width:560px;margin:64px auto;padding:0 24px;line-height:1.5;color:#18201d}h1{font-size:1.5rem}p{font-size:1rem}</style></head><body><h1>${safeTitle}</h1><p>${safeMessage}</p></body></html>`;
  return new Response(body, {
    status,
    headers: {
      "Content-Type": "text/html; charset=utf-8",
      "Cache-Control": "no-store",
      "Referrer-Policy": "no-referrer",
      "X-Content-Type-Options": "nosniff",
      "X-Frame-Options": "DENY",
    },
  });
}

function normalizeScope(value: unknown): string {
  return String(value ?? "")
    .split(/[\s,]+/)
    .map((part) => part.trim())
    .filter(Boolean)
    .join(",");
}

function hasRequiredScope(scope: string): boolean {
  const set = new Set(scope.split(",").map((part) => part.trim()).filter(Boolean));
  return set.has("activity:read_all");
}

async function createLink(req: Request): Promise<Response> {
  if (!(await internalAuthorized(req))) return json({ ok: false, error: "unauthorized" }, 401);
  const body = await req.json().catch(() => ({})) as Record<string, unknown>;
  const athId = String(body.ath_id ?? "").trim().toUpperCase();
  if (!athId) return json({ ok: false, error: "ath_id_required" }, 400);

  const ticket = await rpc<TicketResponse>("strava_oauth_issue_ticket", {
    p_ath_id: athId,
    p_ttl_seconds: 900,
  });
  if (!ticket?.ok || !ticket.state) {
    return json({ ok: false, error: ticket?.reason ?? "ticket_failed" }, ticket?.reason === "athlete_not_found" ? 404 : 409);
  }

  const credentials = await runtimeCredentials();
  const authUrl = new URL(STRAVA_AUTHORIZE_URL);
  authUrl.searchParams.set("client_id", credentials.client_id);
  authUrl.searchParams.set("redirect_uri", CALLBACK_URL);
  authUrl.searchParams.set("response_type", "code");
  authUrl.searchParams.set("approval_prompt", "force");
  authUrl.searchParams.set("scope", ticket.scope || REQUIRED_SCOPE);
  authUrl.searchParams.set("state", ticket.state);

  return json({
    ok: true,
    authorization_url: authUrl.toString(),
    expires_at: ticket.expires_at,
    callback_url: CALLBACK_URL,
  });
}

async function callback(url: URL): Promise<Response> {
  const state = url.searchParams.get("state") ?? "";
  if (!state) return html("Conexão inválida", "O vínculo de conexão está ausente ou expirou. Solicite um novo link.", 400);

  const ticket = await rpc<ConsumeResponse>("strava_oauth_consume_ticket", { p_state: state });
  if (!ticket?.ok || !ticket.ath_id) {
    return html("Link expirado", "Este link já foi usado ou expirou. Solicite uma nova conexão.", 400);
  }

  const denied = url.searchParams.get("error");
  if (denied) {
    return html("Conexão não autorizada", "A autorização do Strava não foi concluída. Nenhuma credencial foi salva.", 200);
  }

  const code = url.searchParams.get("code") ?? "";
  if (!code) return html("Conexão inválida", "O Strava não devolveu um código de autorização válido.", 400);

  const credentials = await runtimeCredentials();
  const form = new URLSearchParams({
    client_id: credentials.client_id,
    client_secret: credentials.client_secret,
    code,
    grant_type: "authorization_code",
  });

  const exchange = await fetch(STRAVA_TOKEN_URL, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: form.toString(),
  });
  if (!exchange.ok) {
    console.error("strava-oauth exchange failed", exchange.status);
    return html("Falha na conexão", "O Strava recusou a troca de autorização. Solicite um novo link e tente novamente.", 502);
  }

  let tokenBody: StravaTokenResponse = {};
  try {
    tokenBody = await exchange.json() as StravaTokenResponse;
  } catch {
    return html("Falha na conexão", "A resposta do Strava não pôde ser validada.", 502);
  }

  const access = String(tokenBody.access_token ?? "");
  const refresh = String(tokenBody.refresh_token ?? "");
  const expiresAt = Number(tokenBody.expires_at ?? 0);
  const stravaId = String(tokenBody.athlete?.id ?? "");
  const callbackScope = normalizeScope(url.searchParams.get("scope"));
  const responseScope = normalizeScope(tokenBody.scope);
  const grantedScope = responseScope || callbackScope;

  if (!access || !refresh || !Number.isFinite(expiresAt) || !stravaId) {
    return html("Falha na conexão", "A autorização foi recebida, mas os dados necessários para concluir o vínculo vieram incompletos.", 502);
  }
  if (!hasRequiredScope(grantedScope)) {
    return html("Permissão insuficiente", "A permissão para leitura completa das atividades não foi concedida. Gere um novo link e mantenha essa permissão habilitada.", 409);
  }

  try {
    const committed = await rpc<OAuthCommitResponse>("strava_oauth_commit", {
      p_ath_id: ticket.ath_id,
      p_access_token: access,
      p_refresh_token: refresh,
      p_expires_at: expiresAt,
      p_scope: grantedScope,
      p_strava_id: stravaId,
    });
    if (!committed?.ok) throw new Error("commit_not_ok");
  } catch (error) {
    console.error("strava-oauth secure commit failed", error instanceof Error ? error.message.replace(/[A-Za-z0-9_-]{20,}/g, "[redacted]") : "unknown");
    return html("Vínculo não concluído", "A autorização foi recebida, mas o vínculo seguro não pôde ser concluído. Solicite um novo link antes de tentar novamente.", 409);
  }

  return html("Strava conectado", "A conexão com o Hiperativo foi concluída com segurança. Você já pode fechar esta página.", 200);
}

Deno.serve(async (req: Request) => {
  try {
    const url = new URL(req.url);
    if (req.method === "POST") {
      const body = await req.clone().json().catch(() => ({})) as Record<string, unknown>;
      const action = String(body.action ?? "create_link");
      if (action === "health") {
        if (!(await internalAuthorized(req))) return json({ ok: false, error: "unauthorized" }, 401);
        return json({ ok: true, callback_url: CALLBACK_URL, required_scope: REQUIRED_SCOPE, storage: "vault" });
      }
      if (action === "create_link") return await createLink(req);
      return json({ ok: false, error: "unknown_action" }, 400);
    }
    if (req.method === "GET") return await callback(url);
    return json({ ok: false, error: "method_not_allowed" }, 405);
  } catch (error) {
    console.error("strava-oauth", error instanceof Error ? error.message.replace(/[A-Za-z0-9_-]{20,}/g, "[redacted]") : "unknown");
    return html("Falha temporária", "Não foi possível concluir esta operação agora. Solicite um novo link antes de tentar novamente.", 503);
  }
});
