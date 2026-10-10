-- Production mirror: strava_legacy_tokens_to_vault
-- Copies existing OAuth secret material into Vault entirely server-side.
-- No token value is selected to a client or written to Git/logs.

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
    select 1 from public.tokens_strava
    where lower(coalesce(status, '')) <> 'inativo'
      and lower(coalesce(status, '')) not like 'revogado%'
      and (access_secret_id is null or refresh_secret_id is null)
  ) then
    raise exception 'STRAVA_VAULT_COPY_INCOMPLETE';
  end if;
end
$$;
