-- Production mirror: strava_oauth_ticket_scaffold
-- Creates opaque, one-use OAuth tickets and prepares token metadata for Vault.

create schema if not exists private;
revoke all on schema private from public, anon, authenticated;

alter table public.tokens_strava
  add column if not exists access_secret_id uuid,
  add column if not exists refresh_secret_id uuid,
  add column if not exists vault_migrated_at timestamptz,
  add column if not exists vault_only boolean not null default false;

alter table public.tokens_strava alter column refresh_token drop not null;

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

create or replace function public.strava_oauth_issue_ticket(p_ath_id text, p_ttl_seconds integer default 900)
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
  select strava_ok into v_strava_ok from public.atletas where ath_id = p_ath_id;
  if not found then return jsonb_build_object('ok', false, 'reason', 'athlete_not_found'); end if;
  if lower(coalesce(v_strava_ok, '')) in ('não utilizar strava', 'nao utilizar strava') then
    return jsonb_build_object('ok', false, 'reason', 'strava_disabled');
  end if;
  update private.strava_oauth_tickets set consumed_at = now() where ath_id = p_ath_id and consumed_at is null;
  v_state := encode(extensions.gen_random_bytes(32), 'hex');
  v_expires := now() + make_interval(secs => v_ttl);
  insert into private.strava_oauth_tickets (ath_id, state_hash, requested_scope, expires_at)
  values (p_ath_id, extensions.digest(v_state, 'sha256'), v_scope, v_expires);
  return jsonb_build_object('ok', true, 'state', v_state, 'scope', v_scope, 'expires_at', v_expires);
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
  select id, ath_id, requested_scope, expires_at into v_id, v_ath_id, v_scope, v_expires
  from private.strava_oauth_tickets
  where state_hash = extensions.digest(lower(p_state), 'sha256') and consumed_at is null
  for update;
  if not found then return jsonb_build_object('ok', false, 'reason', 'invalid_or_used_state'); end if;
  update private.strava_oauth_tickets set consumed_at = now() where id = v_id;
  if v_expires <= now() then return jsonb_build_object('ok', false, 'reason', 'expired_state'); end if;
  return jsonb_build_object('ok', true, 'ath_id', v_ath_id, 'scope', v_scope);
end
$$;

revoke all on function public.strava_oauth_issue_ticket(text, integer) from public, anon, authenticated;
revoke all on function public.strava_oauth_consume_ticket(text) from public, anon, authenticated;
grant execute on function public.strava_oauth_issue_ticket(text, integer) to service_role;
grant execute on function public.strava_oauth_consume_ticket(text) to service_role;
