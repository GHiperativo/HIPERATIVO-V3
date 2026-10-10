-- Hardens self-service OAuth so the browser never executes a SECURITY DEFINER RPC directly.
-- strava-connect verifies the Supabase Auth user, then calls this service-role-only issuer.

create or replace function public.strava_oauth_issue_user_ticket(
  p_auth_user_id uuid,
  p_ttl_seconds integer default 900
)
returns jsonb
language plpgsql
security definer
set search_path = ''
as $$
declare
  v_membership record;
  v_ath_id text;
  v_identity_status text;
  v_ticket jsonb;
  v_client_id text;
begin
  if p_auth_user_id is null then
    return jsonb_build_object('ok', false, 'reason', 'unauthenticated');
  end if;

  select am.participant_id, am.status
  into v_membership
  from hiper_reserva.access_memberships am
  where am.auth_user_id = p_auth_user_id
  order by am.created_at desc
  limit 1;

  if not found then
    return jsonb_build_object('ok', false, 'reason', 'access_context_missing');
  end if;

  if lower(coalesce(v_membership.status, '')) <> 'active' then
    return jsonb_build_object('ok', false, 'reason', 'access_inactive');
  end if;

  select p.ath_id_v3, p.identity_status
  into v_ath_id, v_identity_status
  from hiper_reserva.participants p
  where p.id = v_membership.participant_id
  limit 1;

  if not found then
    return jsonb_build_object('ok', false, 'reason', 'participant_not_found');
  end if;

  if lower(coalesce(v_identity_status, '')) <> 'verified' then
    return jsonb_build_object('ok', false, 'reason', 'identity_not_verified');
  end if;

  v_ath_id := upper(trim(coalesce(v_ath_id, '')));
  if v_ath_id = '' then
    return jsonb_build_object('ok', false, 'reason', 'v3_identity_missing');
  end if;

  v_ticket := public.strava_oauth_issue_ticket(v_ath_id, p_ttl_seconds);
  if not coalesce((v_ticket ->> 'ok')::boolean, false) then
    return v_ticket;
  end if;

  select decrypted_secret into v_client_id
  from vault.decrypted_secrets
  where name = 'strava_client_id'
  limit 1;

  if nullif(v_client_id, '') is null then
    raise exception 'STRAVA_CLIENT_ID_UNAVAILABLE';
  end if;

  return v_ticket || jsonb_build_object('client_id', v_client_id);
end
$$;

revoke all on function public.strava_oauth_issue_user_ticket(uuid, integer) from public, anon, authenticated;
grant execute on function public.strava_oauth_issue_user_ticket(uuid, integer) to service_role;

comment on function public.strava_oauth_issue_user_ticket(uuid, integer) is
'Service-only Strava OAuth ticket issuer. Resolves a verified Hiper Reserva participant to ath_id_v3. The caller must validate the Supabase Auth user before passing p_auth_user_id.';

revoke all on function public.strava_oauth_issue_self_ticket(integer) from public, anon, authenticated;
drop function public.strava_oauth_issue_self_ticket(integer);
