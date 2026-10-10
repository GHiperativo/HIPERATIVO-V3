-- Strava OAuth Vault phase 2: irreversible plaintext cutover.
-- Run only after phase 1 + Edge Functions have proven Vault reads and refresh commits.

-- Active integrations must already have both Vault secrets.
do $$
begin
  if exists (
    select 1
    from public.tokens_strava
    where lower(coalesce(status, '')) <> 'inativo'
      and lower(coalesce(status, '')) not like 'revogado%'
      and (access_secret_id is null or refresh_secret_id is null)
  ) then
    raise exception 'STRAVA_VAULT_CUTOVER_BLOCKED_ACTIVE_SECRET_MISSING';
  end if;

  if exists (
    select 1
    from public.tokens_strava
    where (nullif(access_token, '') is not null and access_secret_id is null)
       or (nullif(refresh_token, '') is not null and refresh_secret_id is null)
  ) then
    raise exception 'STRAVA_VAULT_CUTOVER_BLOCKED_LEGACY_COPY_MISSING';
  end if;
end
$$;

create or replace function public.strava_token_get_access(p_ath_id text)
returns jsonb
language sql
stable
security definer
set search_path = ''
as $$
  select jsonb_build_object(
    'ath_id', t.ath_id,
    'nome', t.nome,
    'access_token', v.decrypted_secret,
    'expires_at', t.expires_at,
    'scope', t.scope,
    'strava_id', t.strava_id,
    'ult_atu', t.ult_atu,
    'status', t.status,
    'token_version', t.token_version
  )
  from public.tokens_strava t
  left join vault.decrypted_secrets v on v.id = t.access_secret_id
  where t.ath_id = p_ath_id
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
  select * into v_row from public.tokens_strava where ath_id = p_ath_id for update;
  if not found then return jsonb_build_object('claimed', false, 'reason', 'not_found'); end if;
  if lower(coalesce(v_row.status, '')) = 'inativo' or lower(coalesce(v_row.status, '')) like 'revogado%' then
    return jsonb_build_object('claimed', false, 'reason', 'inactive');
  end if;
  if v_row.refresh_lease_id is not null and v_row.refresh_lease_until > now() then
    return jsonb_build_object('claimed', false, 'reason', 'busy',
      'retry_after_seconds', greatest(1, ceil(extract(epoch from (v_row.refresh_lease_until - now())))::integer));
  end if;
  if v_row.refresh_secret_id is null then
    return jsonb_build_object('claimed', false, 'reason', 'refresh_secret_missing');
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

  return jsonb_build_object('claimed', true, 'lease_id', v_lease,
    'refresh_token', v_refresh, 'token_version', v_row.token_version);
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
  select * into v_row from public.tokens_strava where ath_id = p_ath_id for update;
  if not found or v_row.refresh_lease_id is distinct from p_lease_id then return false; end if;
  if v_row.access_secret_id is null or v_row.refresh_secret_id is null then
    raise exception 'STRAVA_VAULT_SECRET_ID_MISSING';
  end if;

  perform vault.update_secret(v_row.access_secret_id, p_access_token,
    'strava_access_' || p_ath_id, 'Strava short-lived access token for ' || p_ath_id);
  perform vault.update_secret(v_row.refresh_secret_id, p_refresh_token,
    'strava_refresh_' || p_ath_id, 'Strava rotating refresh token for ' || p_ath_id);

  update public.tokens_strava
  set access_token = null,
      refresh_token = null,
      expires_at = p_expires_at,
      status = 'Renovado',
      ult_atu = now(),
      token_version = token_version + 1,
      vault_only = true,
      vault_migrated_at = coalesce(vault_migrated_at, now()),
      refresh_lease_id = null,
      refresh_lease_until = null,
      last_refresh_ok_at = now(),
      last_refresh_error_at = null,
      last_refresh_error = null
  where ath_id = p_ath_id;
  return true;
end
$$;

-- The actual cutover: secret material is removed from the legacy public columns.
update public.tokens_strava
set access_token = null,
    refresh_token = null,
    vault_only = true,
    vault_migrated_at = coalesce(vault_migrated_at, now());

alter table public.tokens_strava
  drop constraint if exists tokens_strava_oauth_plaintext_null_ck;

alter table public.tokens_strava
  add constraint tokens_strava_oauth_plaintext_null_ck
  check (access_token is null and refresh_token is null);

revoke all on function public.strava_token_get_access(text) from public, anon, authenticated;
revoke all on function public.strava_token_claim_refresh(text, integer) from public, anon, authenticated;
revoke all on function public.strava_token_commit_refresh(text, uuid, text, text, bigint) from public, anon, authenticated;
grant execute on function public.strava_token_get_access(text) to service_role;
grant execute on function public.strava_token_claim_refresh(text, integer) to service_role;
grant execute on function public.strava_token_commit_refresh(text, uuid, text, text, bigint) to service_role;

comment on constraint tokens_strava_oauth_plaintext_null_ck on public.tokens_strava is
'Prevents Strava OAuth secret material from being persisted in legacy public columns after Vault cutover.';
