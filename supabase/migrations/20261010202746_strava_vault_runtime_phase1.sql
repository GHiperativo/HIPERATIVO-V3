-- Production mirror: strava_vault_runtime_phase1
-- New OAuth writes are Vault-only. Legacy rows read Vault first with temporary fallback until final cutover.

create or replace function public.strava_oauth_commit(
  p_ath_id text,
  p_access_token text,
  p_refresh_token text,
  p_expires_at bigint,
  p_scope text,
  p_strava_id text
)
returns jsonb
language plpgsql
security definer
set search_path = ''
as $$
declare
  v_nome text;
  v_atleta_strava_id text;
  v_row public.tokens_strava%rowtype;
  v_access_id uuid;
  v_refresh_id uuid;
  v_now_epoch bigint := floor(extract(epoch from now()))::bigint;
begin
  if nullif(p_access_token, '') is null or nullif(p_refresh_token, '') is null then raise exception 'STRAVA_OAUTH_TOKEN_MISSING'; end if;
  if p_expires_at is null or p_expires_at <= v_now_epoch then raise exception 'STRAVA_OAUTH_EXPIRY_INVALID'; end if;
  if nullif(p_strava_id, '') is null then raise exception 'STRAVA_OAUTH_ATHLETE_ID_MISSING'; end if;
  if position('activity:read_all' in coalesce(p_scope, '')) = 0 then raise exception 'STRAVA_OAUTH_SCOPE_INSUFFICIENT'; end if;

  select nome, strava_id into v_nome, v_atleta_strava_id
  from public.atletas where ath_id = p_ath_id for update;
  if not found then raise exception 'STRAVA_OAUTH_ATHLETE_NOT_FOUND'; end if;
  if nullif(v_atleta_strava_id, '') is not null and v_atleta_strava_id <> p_strava_id then raise exception 'STRAVA_OAUTH_IDENTITY_MISMATCH'; end if;
  if exists (select 1 from public.tokens_strava where strava_id = p_strava_id and ath_id <> p_ath_id) then raise exception 'STRAVA_OAUTH_IDENTITY_ALREADY_LINKED'; end if;

  select * into v_row from public.tokens_strava where ath_id = p_ath_id for update;
  if found then
    if nullif(v_row.strava_id, '') is not null and v_row.strava_id <> p_strava_id then raise exception 'STRAVA_OAUTH_IDENTITY_MISMATCH'; end if;
    v_access_id := v_row.access_secret_id;
    v_refresh_id := v_row.refresh_secret_id;
  end if;

  if v_access_id is null then
    v_access_id := vault.create_secret(p_access_token, 'strava_access_' || p_ath_id, 'Strava short-lived access token for ' || p_ath_id);
  else
    perform vault.update_secret(v_access_id, p_access_token, 'strava_access_' || p_ath_id, 'Strava short-lived access token for ' || p_ath_id);
  end if;
  if v_refresh_id is null then
    v_refresh_id := vault.create_secret(p_refresh_token, 'strava_refresh_' || p_ath_id, 'Strava rotating refresh token for ' || p_ath_id);
  else
    perform vault.update_secret(v_refresh_id, p_refresh_token, 'strava_refresh_' || p_ath_id, 'Strava rotating refresh token for ' || p_ath_id);
  end if;

  insert into public.tokens_strava (
    ath_id, nome, access_token, refresh_token, expires_at, scope, strava_id,
    ult_atu, status, token_version, access_secret_id, refresh_secret_id,
    vault_migrated_at, vault_only, refresh_lease_id, refresh_lease_until,
    last_refresh_ok_at, last_refresh_error_at, last_refresh_error
  ) values (
    p_ath_id, v_nome, null, null, p_expires_at, p_scope, p_strava_id,
    now(), 'Renovado', 1, v_access_id, v_refresh_id,
    now(), true, null, null, now(), null, null
  )
  on conflict (ath_id) do update set
    nome = excluded.nome,
    access_token = null,
    refresh_token = null,
    expires_at = excluded.expires_at,
    scope = excluded.scope,
    strava_id = excluded.strava_id,
    ult_atu = now(),
    status = 'Renovado',
    token_version = public.tokens_strava.token_version + 1,
    access_secret_id = excluded.access_secret_id,
    refresh_secret_id = excluded.refresh_secret_id,
    vault_migrated_at = now(),
    vault_only = true,
    refresh_lease_id = null,
    refresh_lease_until = null,
    last_refresh_ok_at = now(),
    last_refresh_error_at = null,
    last_refresh_error = null;

  update public.atletas
  set strava_ok = 'Conectado', strava_id = p_strava_id, updated_at = now()
  where ath_id = p_ath_id;

  return jsonb_build_object('ok', true, 'ath_id', p_ath_id, 'strava_id', p_strava_id,
    'expires_at', p_expires_at, 'scope', p_scope, 'storage', 'vault');
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
    'access_token', coalesce(v.decrypted_secret, t.access_token),
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
  if lower(coalesce(v_row.status, '')) = 'inativo' or lower(coalesce(v_row.status, '')) like 'revogado%' then return jsonb_build_object('claimed', false, 'reason', 'inactive'); end if;
  if v_row.refresh_lease_id is not null and v_row.refresh_lease_until > now() then
    return jsonb_build_object('claimed', false, 'reason', 'busy',
      'retry_after_seconds', greatest(1, ceil(extract(epoch from (v_row.refresh_lease_until - now())))::integer));
  end if;
  if v_row.refresh_secret_id is not null then
    select decrypted_secret into v_refresh from vault.decrypted_secrets where id = v_row.refresh_secret_id;
  end if;
  v_refresh := coalesce(v_refresh, v_row.refresh_token);
  if nullif(v_refresh, '') is null then return jsonb_build_object('claimed', false, 'reason', 'refresh_secret_missing'); end if;
  v_lease := gen_random_uuid();
  update public.tokens_strava
  set refresh_lease_id = v_lease, refresh_lease_until = now() + make_interval(secs => v_seconds)
  where ath_id = p_ath_id;
  return jsonb_build_object('claimed', true, 'lease_id', v_lease, 'refresh_token', v_refresh, 'token_version', v_row.token_version);
end
$$;

create or replace function public.strava_token_commit_refresh(
  p_ath_id text, p_lease_id uuid, p_access_token text, p_refresh_token text, p_expires_at bigint
)
returns boolean
language plpgsql
security definer
set search_path = ''
as $$
declare
  v_row public.tokens_strava%rowtype;
  v_access_id uuid;
  v_refresh_id uuid;
begin
  select * into v_row from public.tokens_strava where ath_id = p_ath_id for update;
  if not found or v_row.refresh_lease_id is distinct from p_lease_id then return false; end if;
  v_access_id := v_row.access_secret_id;
  v_refresh_id := v_row.refresh_secret_id;
  if v_access_id is null then
    v_access_id := vault.create_secret(p_access_token, 'strava_access_' || p_ath_id, 'Strava short-lived access token for ' || p_ath_id);
  else
    perform vault.update_secret(v_access_id, p_access_token, 'strava_access_' || p_ath_id, 'Strava short-lived access token for ' || p_ath_id);
  end if;
  if v_refresh_id is null then
    v_refresh_id := vault.create_secret(p_refresh_token, 'strava_refresh_' || p_ath_id, 'Strava rotating refresh token for ' || p_ath_id);
  else
    perform vault.update_secret(v_refresh_id, p_refresh_token, 'strava_refresh_' || p_ath_id, 'Strava rotating refresh token for ' || p_ath_id);
  end if;
  update public.tokens_strava
  set access_token = case when v_row.vault_only then null else p_access_token end,
      refresh_token = case when v_row.vault_only then null else p_refresh_token end,
      access_secret_id = v_access_id,
      refresh_secret_id = v_refresh_id,
      vault_migrated_at = coalesce(vault_migrated_at, now()),
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

revoke all on function public.strava_oauth_commit(text, text, text, bigint, text, text) from public, anon, authenticated;
revoke all on function public.strava_token_get_access(text) from public, anon, authenticated;
revoke all on function public.strava_token_claim_refresh(text, integer) from public, anon, authenticated;
revoke all on function public.strava_token_commit_refresh(text, uuid, text, text, bigint) from public, anon, authenticated;
grant execute on function public.strava_oauth_commit(text, text, text, bigint, text, text) to service_role;
grant execute on function public.strava_token_get_access(text) to service_role;
grant execute on function public.strava_token_claim_refresh(text, integer) to service_role;
grant execute on function public.strava_token_commit_refresh(text, uuid, text, text, bigint) to service_role;
