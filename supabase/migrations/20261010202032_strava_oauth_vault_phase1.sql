-- Strava OAuth Vault phase 1.
-- Adds one-time OAuth tickets, migrates legacy token material into Vault in-place,
-- and introduces a dual-read/dual-write compatibility window for existing legacy rows.
-- No token value leaves Postgres during the migration.

create schema if not exists private;
revoke all on schema private from public, anon, authenticated;

alter table public.tokens_strava
  add column if not exists access_secret_id uuid,
  add column if not exists refresh_secret_id uuid,
  add column if not exists vault_migrated_at timestamptz,
  add column if not exists vault_only boolean not null default false;

alter table public.tokens_strava
  alter column refresh_token drop not null;

create table if not exists private.strava_oauth_tickets (
  id uuid primary key default gen_random_uuid(),
  ath_id text not null references public.atletas(ath_id) on delete cascade,
  state_hash bytea not null unique,
  requested_scope text not null,
  created_at timestamptz not null default now(),
  expires_at timestamptz not null,
  consumed_at timestamptz,
  constraint strava_oauth_ticket_expiry_ck check (expires_at > created_at)
);

alter table private.strava_oauth_tickets enable row level security;
revoke all on table private.strava_oauth_tickets from public, anon, authenticated;

create index if not exists strava_oauth_tickets_ath_active_idx
  on private.strava_oauth_tickets (ath_id, expires_at desc)
  where consumed_at is null;

-- Copy legacy values to Vault entirely inside Postgres. Values are never selected back to the caller.
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
    v_access_id := r.access_secret_id;
    v_refresh_id := r.refresh_secret_id;
    v_access_name := 'strava_access_' || r.ath_id;
    v_refresh_name := 'strava_refresh_' || r.ath_id;

    if nullif(r.access_token, '') is not null then
      if v_access_id is null then
        select id into v_access_id from vault.secrets where name = v_access_name order by created_at asc limit 1;
      end if;
      if v_access_id is null then
        v_access_id := vault.create_secret(r.access_token, v_access_name, 'Strava short-lived access token for ' || r.ath_id);
      else
        perform vault.update_secret(v_access_id, r.access_token, v_access_name, 'Strava short-lived access token for ' || r.ath_id);
      end if;
    end if;

    if nullif(r.refresh_token, '') is not null then
      if v_refresh_id is null then
        select id into v_refresh_id from vault.secrets where name = v_refresh_name order by created_at asc limit 1;
      end if;
      if v_refresh_id is null then
        v_refresh_id := vault.create_secret(r.refresh_token, v_refresh_name, 'Strava rotating refresh token for ' || r.ath_id);
      else
        perform vault.update_secret(v_refresh_id, r.refresh_token, v_refresh_name, 'Strava rotating refresh token for ' || r.ath_id);
      end if;
    end if;

    update public.tokens_strava
    set access_secret_id = v_access_id,
        refresh_secret_id = v_refresh_id,
        vault_migrated_at = case
          when v_access_id is not null or v_refresh_id is not null then coalesce(vault_migrated_at, now())
          else vault_migrated_at
        end
    where ath_id = r.ath_id;
  end loop;

  if exists (
    select 1
    from public.tokens_strava
    where lower(coalesce(status, '')) <> 'inativo'
      and lower(coalesce(status, '')) not like 'revogado%'
      and (access_secret_id is null or refresh_secret_id is null)
  ) then
    raise exception 'STRAVA_VAULT_COPY_INCOMPLETE';
  end if;
end
$$;

create or replace function public.strava_oauth_issue_ticket(
  p_ath_id text,
  p_ttl_seconds integer default 900
)
returns jsonb
language plpgsql
security definer
set search_path = ''
as $$
declare
  v_state text;
  v_ttl integer := greatest(300, least(coalesce(p_ttl_seconds, 900), 1800));
  v_scope text := 'read,activity:read_all';
  v_strava_ok text;
  v_expires timestamptz;
begin
  select strava_ok into v_strava_ok
  from public.atletas
  where ath_id = p_ath_id;

  if not found then
    return jsonb_build_object('ok', false, 'reason', 'athlete_not_found');
  end if;

  if lower(coalesce(v_strava_ok, '')) in ('não utilizar strava', 'nao utilizar strava') then
    return jsonb_build_object('ok', false, 'reason', 'strava_disabled');
  end if;

  update private.strava_oauth_tickets
  set consumed_at = now()
  where ath_id = p_ath_id
    and consumed_at is null;

  v_state := encode(extensions.gen_random_bytes(32), 'hex');
  v_expires := now() + make_interval(secs => v_ttl);

  insert into private.strava_oauth_tickets (ath_id, state_hash, requested_scope, expires_at)
  values (p_ath_id, extensions.digest(v_state, 'sha256'), v_scope, v_expires);

  return jsonb_build_object(
    'ok', true,
    'state', v_state,
    'scope', v_scope,
    'expires_at', v_expires
  );
end
$$;

create or replace function public.strava_oauth_consume_ticket(p_state text)
returns jsonb
language plpgsql
security definer
set search_path = ''
as $$
declare
  v_id uuid;
  v_ath_id text;
  v_scope text;
  v_expires timestamptz;
begin
  if p_state is null or length(p_state) <> 64 or p_state !~ '^[0-9a-fA-F]{64}$' then
    return jsonb_build_object('ok', false, 'reason', 'invalid_state');
  end if;

  select id, ath_id, requested_scope, expires_at
  into v_id, v_ath_id, v_scope, v_expires
  from private.strava_oauth_tickets
  where state_hash = extensions.digest(lower(p_state), 'sha256')
    and consumed_at is null
  for update;

  if not found then
    return jsonb_build_object('ok', false, 'reason', 'invalid_or_used_state');
  end if;

  update private.strava_oauth_tickets
  set consumed_at = now()
  where id = v_id;

  if v_expires <= now() then
    return jsonb_build_object('ok', false, 'reason', 'expired_state');
  end if;

  return jsonb_build_object('ok', true, 'ath_id', v_ath_id, 'scope', v_scope);
end
$$;

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
  if nullif(p_access_token, '') is null or nullif(p_refresh_token, '') is null then
    raise exception 'STRAVA_OAUTH_TOKEN_MISSING';
  end if;
  if p_expires_at is null or p_expires_at <= v_now_epoch then
    raise exception 'STRAVA_OAUTH_EXPIRY_INVALID';
  end if;
  if nullif(p_strava_id, '') is null then
    raise exception 'STRAVA_OAUTH_ATHLETE_ID_MISSING';
  end if;
  if position('activity:read_all' in coalesce(p_scope, '')) = 0 then
    raise exception 'STRAVA_OAUTH_SCOPE_INSUFFICIENT';
  end if;

  select nome, strava_id into v_nome, v_atleta_strava_id
  from public.atletas
  where ath_id = p_ath_id
  for update;
  if not found then
    raise exception 'STRAVA_OAUTH_ATHLETE_NOT_FOUND';
  end if;

  if nullif(v_atleta_strava_id, '') is not null and v_atleta_strava_id <> p_strava_id then
    raise exception 'STRAVA_OAUTH_IDENTITY_MISMATCH';
  end if;

  if exists (
    select 1 from public.tokens_strava
    where strava_id = p_strava_id and ath_id <> p_ath_id
  ) then
    raise exception 'STRAVA_OAUTH_IDENTITY_ALREADY_LINKED';
  end if;

  select * into v_row
  from public.tokens_strava
  where ath_id = p_ath_id
  for update;

  if found then
    if nullif(v_row.strava_id, '') is not null and v_row.strava_id <> p_strava_id then
      raise exception 'STRAVA_OAUTH_IDENTITY_MISMATCH';
    end if;
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
    now(), true, null, null,
    now(), null, null
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
  set strava_ok = 'Conectado',
      strava_id = p_strava_id,
      updated_at = now()
  where ath_id = p_ath_id;

  return jsonb_build_object(
    'ok', true,
    'ath_id', p_ath_id,
    'strava_id', p_strava_id,
    'expires_at', p_expires_at,
    'scope', p_scope,
    'storage', 'vault'
  );
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
  if lower(coalesce(v_row.status, '')) = 'inativo' or lower(coalesce(v_row.status, '')) like 'revogado%' then
    return jsonb_build_object('claimed', false, 'reason', 'inactive');
  end if;
  if v_row.refresh_lease_id is not null and v_row.refresh_lease_until > now() then
    return jsonb_build_object('claimed', false, 'reason', 'busy',
      'retry_after_seconds', greatest(1, ceil(extract(epoch from (v_row.refresh_lease_until - now())))::integer));
  end if;

  if v_row.refresh_secret_id is not null then
    select decrypted_secret into v_refresh from vault.decrypted_secrets where id = v_row.refresh_secret_id;
  end if;
  v_refresh := coalesce(v_refresh, v_row.refresh_token);
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

revoke all on function public.strava_oauth_issue_ticket(text, integer) from public, anon, authenticated;
revoke all on function public.strava_oauth_consume_ticket(text) from public, anon, authenticated;
revoke all on function public.strava_oauth_commit(text, text, text, bigint, text, text) from public, anon, authenticated;
revoke all on function public.strava_token_get_access(text) from public, anon, authenticated;
revoke all on function public.strava_token_claim_refresh(text, integer) from public, anon, authenticated;
revoke all on function public.strava_token_commit_refresh(text, uuid, text, text, bigint) from public, anon, authenticated;

grant execute on function public.strava_oauth_issue_ticket(text, integer) to service_role;
grant execute on function public.strava_oauth_consume_ticket(text) to service_role;
grant execute on function public.strava_oauth_commit(text, text, text, bigint, text, text) to service_role;
grant execute on function public.strava_token_get_access(text) to service_role;
grant execute on function public.strava_token_claim_refresh(text, integer) to service_role;
grant execute on function public.strava_token_commit_refresh(text, uuid, text, text, bigint) to service_role;

comment on function public.strava_oauth_issue_ticket(text, integer) is
'Issues one short-lived opaque OAuth state for an existing athlete. Service role only.';
comment on function public.strava_oauth_commit(text, text, text, bigint, text, text) is
'Commits Strava OAuth credentials directly to Vault after identity/scope checks. Plaintext columns remain null for OAuth/reconnect writes.';
