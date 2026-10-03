import "jsr:@supabase/functions-js/edge-runtime.d.ts";
import { createClient } from "@supabase/supabase-js";

// Edge Function privada para acesso individual do Hiper Reserva.
// O gateway exige JWT e o servidor valida novamente a sessão no Supabase Auth.
// Convites administrativos pré-autorizados só podem ser reivindicados
// depois de o e-mail estar confirmado no Auth.
const apiUrl = Deno.env.get("SUPABASE_URL");
const serviceKey = Deno.env.get("SUPABASE_SERVICE_ROLE_KEY");

const jsonHeaders = {
  "Content-Type": "application/json; charset=utf-8",
  "Cache-Control": "no-store",
  "Vary": "Authorization",
};

function reply(status: number, body: Record<string, unknown>): Response {
  return new Response(JSON.stringify(body), { status, headers: jsonHeaders });
}

Deno.serve(async (request: Request) => {
  if (request.method !== "GET" && request.method !== "POST") {
    return reply(405, { error: "method_not_allowed" });
  }

  const match = /^Bearer\s+(\S+)$/i.exec(request.headers.get("Authorization") ?? "");
  if (!match) return reply(401, { error: "authentication_required" });
  if (!apiUrl || !serviceKey) return reply(503, { error: "service_unavailable" });

  try {
    const admin = createClient(apiUrl, serviceKey, {
      auth: { autoRefreshToken: false, persistSession: false },
    });

    const { data, error: authError } = await admin.auth.getUser(match[1]);
    const user = data?.user;
    if (authError || !user) return reply(401, { error: "invalid_session" });

    if (!user.email || !user.email_confirmed_at) {
      return reply(403, { error: "verified_email_required" });
    }

    const { data: current, error: stateError } =
      await admin.rpc("hr_get_verified_enrollment_state", {
        p_auth_user_id: user.id,
      });

    if (stateError) {
      console.error("hr_enrollment_state_failure", stateError.code ?? "unknown");
      return reply(503, { error: "service_unavailable" });
    }

    if (request.method === "GET") {
      return reply(200, { status: current ?? "not_requested" });
    }

    if (current === "active") return reply(200, { status: "active" });
    if (current === "suspended" || current === "revoked" || current === "rejected") {
      return reply(403, { status: current, error: "contact_administration" });
    }

    // Primeiro verifica se o usuário possui convite pré-autorizado de equipe.
    const { data: staffClaim, error: staffClaimError } =
      await admin.rpc("hr_claim_staff_invite", {
        p_auth_user_id: user.id,
      });

    if (staffClaimError) {
      console.error("hr_staff_claim_failure", staffClaimError.code ?? "unknown");
      return reply(503, { error: "service_unavailable" });
    }

    if (typeof staffClaim === "string" && staffClaim.startsWith("active:")) {
      const role = staffClaim.slice("active:".length);
      return reply(200, { status: "active", role });
    }

    if (typeof staffClaim === "string" &&
        (staffClaim.startsWith("suspended:") || staffClaim.startsWith("revoked:"))) {
      return reply(403, { status: staffClaim.split(":")[0], error: "contact_administration" });
    }

    if (current === "pending") return reply(202, { status: "pending" });

    // Participantes sem convite de equipe seguem para solicitação comum,
    // sempre pendente de conferência administrativa.
    const { error: requestError } =
      await admin.rpc("hr_submit_verified_enrollment", {
        p_auth_user_id: user.id,
      });

    if (requestError) {
      console.error("hr_enrollment_submit_failure", requestError.code ?? "unknown");
      return reply(503, { error: "service_unavailable" });
    }

    return reply(202, {
      status: "pending",
      message: "Aguardando conferência da administração.",
    });
  } catch {
    return reply(503, { error: "service_unavailable" });
  }
});
