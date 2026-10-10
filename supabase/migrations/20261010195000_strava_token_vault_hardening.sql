-- P0 Strava hardening: move athlete OAuth secrets to Supabase Vault and serialize refreshes.
-- Phase 1 is non-destructive: legacy plaintext columns remain populated until the Vault runtime is validated.

alter table public.tokens_strava
  add column if not exists access_secret_id uuid,
  add column if not exists refresh_secret_id uuid,
  add column if not exists token_version bigint not null default 0,
  add column if not exists refresh_lease_id uuid,
  add column if not exists refresh_lease_until timestamptz,
  add column if not exists last_refresh_ok_at timestamptz,
  add column if not exists last_refresh_error_at timestamptz,
  add column if not exists last_refresh_error text;

-- RLS already restricts rows to service_role. Remove unnecessary table privileges as a second boundary.
revoke all on table public.tokens_strava from anon, authenticated;
grant select, insert, update, delete on table public.tokens_strava to service_role;

do $$
declare
  r record;
  v_access_id uuid;
  v_refresh_id uuid;
  v_access_name text;
  v_refresh_name text;
begin
  for r in
    select ath_id, access_token, refresh_token, access_secret_id, refresh_secret_id
    from public.tokens_strava
    for update
  loop
    v_access_name := 'strava_access_' || r.ath_id;
    v_refresh_name := 'strava_refresh_' || r.ath_id;
    v_access_id := r.access_secret_id;
    v_refresh_id := r.refresh_secret_id;

    if nullif(r.access_token, '') is not null then
      if v_access_id is null then
        select id into v_access_id from vault.secrets where name = v_access_name limit 1;
      end if;
      if v_access_id is null then
        v_access_id := vault.create_secret(r.access_token, v_access_name, 'Short-lived Strava access token for ' || r.ath_id);
      else
        perform vault.update_secret(v_access_id, r.access_token, v_access_name, 'Short-lived Strava access token for ' || r.ath_id);
      end if;
    end if;

    if nullif(r.refresh_token, '') is not null then
      if v_refresh_id is null then
        select id into v_refresh_id from vault.secrets where name = v_refresh_name limit 1;
      end if;
      if v_refresh_id is null then
        v_refresh_id := vault.create_secret(r.refresh_token, v_refresh_name, 'Rotating Strava refresh token for ' || r.ath_id);
      else
        perform vault.update_secret(v_refresh_id, r.refresh_token, v_refresh_name, 'Rotating Strava refresh token for ' || r.ath_id);
      end if;
    end if;

    update public.tokens_strava
    set access_secret_id = v_access_id,
        refresh_secret_id = v_refresh_id
    where ath_id = r.ath_id;
  end loop;
end
$$;

create or replace function public.strava_token_metadata_list()
returns jsonb
language sql
stable
security definer
set search_path = ''
as $$
  select coalesce(jsonb_agg(
    jsonb_build_object(
      'ath_id', t.ath_id,
      'nome', t.nome,
      'expires_at', t.expires_at,
      'scope', t.scope,
      'strava_id', t.strava_id,
      'ult_atu', t.ult_atu,
      'status', t.status,
      'token_version', t.token_version
    ) order by t.ath_id
  ), '[]'::jsonb)
  from public.tokens_strava t;
$$;

create or replace function public.strava_token_get(p_ath_id text)
returns jsonb
language sql
stable
security definer
set search_path = ''
as $$
  select jsonb_build_object(
    'ath_id', t.ath_id,
    'nome', t.nome,
    'access_token', a.decrypted_secret,
    'refresh_token', r.decrypted_secret,
    'expires_at', t.expires_at,
    'scope', t.scope,
    'strava_id', t.strava_id,
    'ult_atu', t.ult_atu,
    'status', t.status,
    'token_version', t.token_version
  )
  from public.tokens_strava t
  left join vault.decrypted_secrets a on a.id = t.access_secret_id
  left join vault.decrypted_secrets r on r.id = t.refresh_secret_id
  where t.ath_id = p_ath_id
  limit 1;
$$;

create or replace function public.strava_token_get_by_owner(p_owner_id bigint)
returns jsonb
language sql
stable
security definer
set search_path = ''
as $$
  select jsonb_build_object(
    'ath_id', t.ath_id,
    'nome', t.nome,
    'access_token', a.decrypted_secret,
    'refresh_token', r.decrypted_secret,
    'expires_at', t.expires_at,
    'scope', t.scope,
    'strava_id', t.strava_id,
    'ult_atu', t.ult_atu,
    'status', t.status,
    'token_version', t.token_version
  )
  from public.tokens_strava t
  left join vault.decrypted_secrets a on a.id = t.access_secret_id
  left join vault.decrypted_secrets r on r.id = t.refresh_secret_id
  where t.strava_id = p_owner_id::text
  limit 1;
$$;

create or replace function public.strava_token_claim_refresh(p_ath_id text, p_lease_seconds integer default 90)
returns jsonb
language plpgsql
security definer
set search_path = ''
as $$
declare
  v_row public.tokens_strava%rowtype;
  v_lease uuid;
  v_refresh text;
  v_seconds integer := greatest(30, least(coalesce(p_lease_seconds, 90), 300));
begin
  select * into v_row
  from public.tokens_strava
  where ath_id = p_ath_id
  for update;

  if not found then
    return jsonb_build_object('claimed', false, 'reason', 'not_found');
  end if;

  if lower(coalesce(v_row.status, '')) = 'inativo' or lower(coalesce(v_row.status, '')) like 'revogado%' then
    return jsonb_build_object('claimed', false, 'reason', 'inactive');
  end if;

  if v_row.refresh_lease_id is not null and v_row.refresh_lease_until > now() then
    return jsonb_build_object(
      'claimed', false,
      'reason', 'busy',
      'retry_after_seconds', greatest(1, ceil(extract(epoch from (v_row.refresh_lease_until - now())))::integer)
    );
  end if;

  select decrypted_secret into v_refresh
  from vault.decrypted_secrets
  where id = v_row.refresh_secret_id;

  if nullif(v_refresh, '') is null then
    return jsonb_build_object('claimed', false, 'reason', 'refresh_secret_missing');
  end if;

  v_lease := gen_random_uuid();
  update public.tokens_strava
  set refresh_lease_id = v_lease,
      refresh_lease_until = now() + make_interval(secs => v_seconds)
  where ath_id = p_ath_id;

  return jsonb_build_object(
    'claimed', true,
    'lease_id', v_lease,
    'refresh_token', v_refresh,
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

  if v_row.access_secret_id is null or v_row.refresh_secret_id is null then
    raise exception 'Strava Vault secret ids missing for %', p_ath_id;
  end if;

  perform vault.update_secret(v_row.access_secret_id, p_access_token);
  perform vault.update_secret(v_row.refresh_secret_id, p_refresh_token);

  update public.tokens_strava
  set expires_at = p_expires_at,
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

create or replace function public.strava_token_mark_revoked(p_ath_id text)
returns boolean
language plpgsql
security definer
set search_path = ''
as $$
begin
  update public.tokens_strava
  set status = 'Revogado pela Strava',
      ult_atu = now(),
      refresh_lease_id = null,
      refresh_lease_until = null
  where ath_id = p_ath_id;
  return found;
end
$$;

revoke all on function public.strava_token_metadata_list() from public, anon, authenticated;
revoke all on function public.strava_token_get(text) from public, anon, authenticated;
revoke all on function public.strava_token_get_by_owner(bigint) from public, anon, authenticated;
revoke all on function public.strava_token_claim_refresh(text, integer) from public, anon, authenticated;
revoke all on function public.strava_token_commit_refresh(text, uuid, text, text, bigint) from public, anon, authenticated;
revoke all on function public.strava_token_fail_refresh(text, uuid, text) from public, anon, authenticated;
revoke all on function public.strava_token_mark_revoked(text) from public, anon, authenticated;

grant execute on function public.strava_token_metadata_list() to service_role;
grant execute on function public.strava_token_get(text) to service_role;
grant execute on function public.strava_token_get_by_owner(bigint) to service_role;
grant execute on function public.strava_token_claim_refresh(text, integer) to service_role;
grant execute on function public.strava_token_commit_refresh(text, uuid, text, text, bigint) to service_role;
grant execute on function public.strava_token_fail_refresh(text, uuid, text) to service_role;
grant execute on function public.strava_token_mark_revoked(text) to service_role;

comment on function public.strava_token_claim_refresh(text, integer) is
'Claims a short refresh lease so concurrent webhook/cron workers cannot rotate the same Strava refresh token simultaneously.';
comment on function public.strava_token_commit_refresh(text, uuid, text, text, bigint) is
'Atomically persists the latest Strava access/refresh token pair into Vault and advances token_version.';
