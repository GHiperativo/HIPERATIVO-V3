-- P0 RED Strava recovery: move runtime ingestion from the dead Apps Script chain to Supabase.
-- Scoped to Strava only. Does not alter Hiper Reserva, Auth, payments, or participant identity.

create extension if not exists pg_net;
create extension if not exists pg_cron;

do $$
declare
  v_id uuid;
begin
  select id into v_id from vault.secrets where name = 'strava_client_id' limit 1;
  if v_id is null then
    perform vault.create_secret('153043', 'strava_client_id', 'Strava application client id for Hiperativo V3');
  else
    perform vault.update_secret(v_id, '153043', 'strava_client_id', 'Strava application client id for Hiperativo V3');
  end if;

  select id into v_id from vault.secrets where name = 'strava_internal_token' limit 1;
  if v_id is null then
    perform vault.create_secret(
      encode(gen_random_bytes(32), 'hex'),
      'strava_internal_token',
      'Internal token used only by pg_cron -> strava-sync'
    );
  end if;
end
$$;

create or replace function public.strava_runtime_config()
returns jsonb
language sql
stable
security definer
set search_path = ''
as $$
  select jsonb_build_object(
    'client_id', coalesce((select decrypted_secret from vault.decrypted_secrets where name = 'strava_client_id' limit 1), ''),
    'client_secret', coalesce((select decrypted_secret from vault.decrypted_secrets where name = 'strava_client_secret' limit 1), '')
  );
$$;

revoke all on function public.strava_runtime_config() from public, anon, authenticated;
grant execute on function public.strava_runtime_config() to service_role;

create or replace function public.strava_internal_token_matches(p_token text)
returns boolean
language sql
stable
security definer
set search_path = ''
as $$
  select coalesce(p_token, '') <> ''
    and coalesce(p_token, '') = coalesce(
      (select decrypted_secret from vault.decrypted_secrets where name = 'strava_internal_token' limit 1),
      ''
    );
$$;

revoke all on function public.strava_internal_token_matches(text) from public, anon, authenticated;
grant execute on function public.strava_internal_token_matches(text) to service_role;

create or replace function public.strava_upsert_activities(p_items jsonb)
returns jsonb
language plpgsql
security definer
set search_path = ''
as $$
declare
  v_item jsonb;
  v_strava_id text;
  v_ath_id text;
  v_existing_ath text;
  v_inserted int := 0;
  v_updated int := 0;
  v_conflicts int := 0;
  v_exists boolean;
begin
  if jsonb_typeof(p_items) <> 'array' then
    raise exception 'p_items must be a JSON array';
  end if;

  for v_item in select value from jsonb_array_elements(p_items)
  loop
    v_strava_id := nullif(trim(v_item->>'strava_id'), '');
    v_ath_id := nullif(trim(v_item->>'ath_id'), '');
    if v_strava_id is null or v_ath_id is null then continue; end if;

    select a.ath_id into v_existing_ath
    from public.atividades a
    where a.strava_id = v_strava_id
    limit 1;

    if v_existing_ath is not null and v_existing_ath <> v_ath_id then
      v_conflicts := v_conflicts + 1;
      continue;
    end if;

    v_exists := v_existing_ath is not null;

    insert into public.atividades (
      exec_id, ath_id, nome, data, tipo, fonte, strava_id, nome_ativ,
      mov_s, total_s, dist_m, dist_km, vel_mps, pace_s_km, pace_fmt,
      fc_med, fc_max, elev, cal, cadencia, importado_at
    ) values (
      'ATIV_' || upper(substr(replace(gen_random_uuid()::text, '-', ''), 1, 8)),
      v_ath_id,
      nullif(v_item->>'nome', ''),
      nullif(v_item->>'data', '')::timestamptz,
      nullif(v_item->>'tipo', ''),
      'Strava',
      v_strava_id,
      nullif(v_item->>'nome_ativ', ''),
      nullif(v_item->>'mov_s', '')::integer,
      nullif(v_item->>'total_s', '')::integer,
      nullif(v_item->>'dist_m', '')::integer,
      nullif(v_item->>'dist_km', '')::numeric,
      nullif(v_item->>'vel_mps', '')::numeric,
      nullif(v_item->>'pace_s_km', '')::integer,
      coalesce(v_item->>'pace_fmt', ''),
      nullif(v_item->>'fc_med', '')::integer,
      nullif(v_item->>'fc_max', '')::integer,
      nullif(v_item->>'elev', '')::integer,
      nullif(v_item->>'cal', '')::integer,
      nullif(v_item->>'cadencia', '')::integer,
      now()
    )
    on conflict (strava_id) do update set
      nome = coalesce(excluded.nome, public.atividades.nome),
      data = coalesce(excluded.data, public.atividades.data),
      tipo = coalesce(excluded.tipo, public.atividades.tipo),
      fonte = 'Strava',
      nome_ativ = coalesce(excluded.nome_ativ, public.atividades.nome_ativ),
      mov_s = coalesce(excluded.mov_s, public.atividades.mov_s),
      total_s = coalesce(excluded.total_s, public.atividades.total_s),
      dist_m = coalesce(excluded.dist_m, public.atividades.dist_m),
      dist_km = coalesce(excluded.dist_km, public.atividades.dist_km),
      vel_mps = coalesce(excluded.vel_mps, public.atividades.vel_mps),
      pace_s_km = coalesce(excluded.pace_s_km, public.atividades.pace_s_km),
      pace_fmt = coalesce(excluded.pace_fmt, public.atividades.pace_fmt),
      fc_med = coalesce(excluded.fc_med, public.atividades.fc_med),
      fc_max = coalesce(excluded.fc_max, public.atividades.fc_max),
      elev = coalesce(excluded.elev, public.atividades.elev),
      cal = coalesce(excluded.cal, public.atividades.cal),
      cadencia = coalesce(excluded.cadencia, public.atividades.cadencia),
      importado_at = now();

    if v_exists then v_updated := v_updated + 1; else v_inserted := v_inserted + 1; end if;
  end loop;

  return jsonb_build_object('inserted', v_inserted, 'updated', v_updated, 'identity_conflicts', v_conflicts);
end
$$;

revoke all on function public.strava_upsert_activities(jsonb) from public, anon, authenticated;
grant execute on function public.strava_upsert_activities(jsonb) to service_role;

create or replace function public.strava_delete_activity(p_strava_id text, p_ath_id text)
returns integer
language plpgsql
security definer
set search_path = ''
as $$
declare
  v_count integer;
begin
  delete from public.atividades
  where strava_id = p_strava_id
    and ath_id = p_ath_id
    and fonte = 'Strava';
  get diagnostics v_count = row_count;
  return v_count;
end
$$;

revoke all on function public.strava_delete_activity(text, text) from public, anon, authenticated;
grant execute on function public.strava_delete_activity(text, text) to service_role;

comment on function public.strava_runtime_config() is
'P0 Strava runtime config. Service role only. Client secret lives encrypted in Supabase Vault.';
comment on function public.strava_upsert_activities(jsonb) is
'Idempotent Strava activity upsert by unique strava_id. Preserves exec_id and manual PSE.';
