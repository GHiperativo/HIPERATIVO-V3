-- P0 Strava hardening, phase A: least privilege + serialized refresh + audit metadata.
-- This migration deliberately does NOT move OAuth secret values. The platform security layer
-- blocks automated secret transport. Vault migration is a separate gate after this runtime is proven.

alter table public.tokens_strava
  add column if not exists token_version bigint not null default 0,
  add column if not exists refresh_lease_id uuid,
  add column if not exists refresh_lease_until timestamptz,
  add column if not exists last_refresh_ok_at timestamptz,
  add column if not exists last_refresh_error_at timestamptz,
  add column if not exists last_refresh_error text;

-- RLS already has only service_role policy. Remove redundant Data API table grants.
revoke all on table public.tokens_strava from anon, authenticated;
grant select, insert, update, delete on table public.tokens_strava to service_role;

create or replace function public.strava_token_claim_refresh(p_ath_id text, p_lease_seconds integer default 90)
returns jsonb
language plpgsql
security definer
set search_path = ''
as $$
declare
  v_row public.tokens_strava%rowtype;
  v_lease uuid;
  v_seconds integer := greatest(30, least(coalesce(p_lease_seconds, 90), 300));
begin
  select * into v_row
  from public.tokens_strava
  where ath_id = p_ath_id
  for update;

  if not found then
    return jsonb_build_object('claimed', false, 'reason', 'not_found');
  end if;

  if lower(coalesce(v_row.status, '')) = 'inativo'
     or lower(coalesce(v_row.status, '')) like 'revogado%' then
    return jsonb_build_object('claimed', false, 'reason', 'inactive');
  end if;

  if v_row.refresh_lease_id is not null and v_row.refresh_lease_until > now() then
    return jsonb_build_object(
      'claimed', false,
      'reason', 'busy',
      'retry_after_seconds', greatest(1, ceil(extract(epoch from (v_row.refresh_lease_until - now())))::integer)
    );
  end if;

  if nullif(v_row.refresh_token, '') is null then
    return jsonb_build_object('claimed', false, 'reason', 'refresh_token_missing');
  end if;

  v_lease := gen_random_uuid();
  update public.tokens_strava
  set refresh_lease_id = v_lease,
      refresh_lease_until = now() + make_interval(secs => v_seconds)
  where ath_id = p_ath_id;

  return jsonb_build_object(
    'claimed', true,
    'lease_id', v_lease,
    'refresh_token', v_row.refresh_token,
    'token_version', v_row.token_version
  );
end
$$;

create or replace function public.strava_token_commit_refresh(
  p_ath_id text,
  p_lease_id uuid,
  p_access_token text,
  p_refresh_token text,
  p_expires_at bigint
)
returns boolean
language plpgsql
security definer
set search_path = ''
as $$
declare
  v_row public.tokens_strava%rowtype;
begin
  select * into v_row
  from public.tokens_strava
  where ath_id = p_ath_id
  for update;

  if not found or v_row.refresh_lease_id is distinct from p_lease_id then
    return false;
  end if;

  update public.tokens_strava
  set access_token = p_access_token,
      refresh_token = p_refresh_token,
      expires_at = p_expires_at,
      status = 'Renovado',
      ult_atu = now(),
      token_version = token_version + 1,
      refresh_lease_id = null,
      refresh_lease_until = null,
      last_refresh_ok_at = now(),
      last_refresh_error_at = null,
      last_refresh_error = null
  where ath_id = p_ath_id;

  return true;
end
$$;

create or replace function public.strava_token_fail_refresh(p_ath_id text, p_lease_id uuid, p_error text)
returns boolean
language plpgsql
security definer
set search_path = ''
as $$
begin
  update public.tokens_strava
  set refresh_lease_id = null,
      refresh_lease_until = null,
      last_refresh_error_at = now(),
      last_refresh_error = left(coalesce(p_error, 'unknown refresh error'), 500)
  where ath_id = p_ath_id
    and refresh_lease_id = p_lease_id;
  return found;
end
$$;

revoke all on function public.strava_token_claim_refresh(text, integer) from public, anon, authenticated;
revoke all on function public.strava_token_commit_refresh(text, uuid, text, text, bigint) from public, anon, authenticated;
revoke all on function public.strava_token_fail_refresh(text, uuid, text) from public, anon, authenticated;

grant execute on function public.strava_token_claim_refresh(text, integer) to service_role;
grant execute on function public.strava_token_commit_refresh(text, uuid, text, text, bigint) to service_role;
grant execute on function public.strava_token_fail_refresh(text, uuid, text) to service_role;

comment on function public.strava_token_claim_refresh(text, integer) is
'Claims a short lease so webhook and cron workers cannot rotate the same Strava refresh token concurrently.';
comment on function public.strava_token_commit_refresh(text, uuid, text, text, bigint) is
'Commits the latest Strava token pair atomically only for the worker holding the refresh lease.';
