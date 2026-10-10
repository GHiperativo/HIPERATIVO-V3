-- Authenticated self-service Strava OAuth link issuance.
-- The browser never supplies ath_id. Identity is resolved from Hiper Reserva.

create or replace function public.strava_oauth_issue_self_ticket(p_ttl_seconds integer default 900)
returns jsonb
language plpgsql
security definer
set search_path = ''
as $$
declare
  v_ctx record;
  v_ath_id text;
  v_identity_status text;
  v_ticket jsonb;
  v_client_id text;
begin
  if auth.uid() is null then
    return jsonb_build_object('ok', false, 'reason', 'unauthenticated');
  end if;

  select * into v_ctx
  from hiper_reserva.resolve_access_context()
  limit 1;

  if not found then
    return jsonb_build_object('ok', false, 'reason', 'access_context_missing');
  end if;

  if lower(coalesce(v_ctx.status, '')) <> 'active' then
    return jsonb_build_object('ok', false, 'reason', 'access_inactive');
  end if;

  select p.ath_id_v3, p.identity_status
  into v_ath_id, v_identity_status
  from hiper_reserva.participants p
  where p.id = v_ctx.participant_id
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

revoke all on function public.strava_oauth_issue_self_ticket(integer) from public, anon;
grant execute on function public.strava_oauth_issue_self_ticket(integer) to authenticated;

comment on function public.strava_oauth_issue_self_ticket(integer) is
'Issues a one-use Strava OAuth ticket only for the authenticated user own verified Hiper Reserva participant mapped to ath_id_v3. Returns no OAuth secret.';
