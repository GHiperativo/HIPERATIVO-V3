-- Keep Strava OAuth tokens healthy independently from activity reconciliation.
-- The worker only refreshes tokens inside the one-hour safety window.

do $$
declare
  v_jobid bigint;
begin
  select jobid into v_jobid
  from cron.job
  where jobname = 'strava-token-maintenance-hourly'
  limit 1;

  if v_jobid is not null then
    perform cron.unschedule(v_jobid);
  end if;
end
$$;

select cron.schedule(
  'strava-token-maintenance-hourly',
  '7 * * * *',
  $job$
  select net.http_post(
    url := 'https://korlpbclqgmqvpbrungc.supabase.co/functions/v1/strava-sync',
    headers := jsonb_build_object(
      'Content-Type', 'application/json',
      'X-Strava-Internal-Secret', (
        select decrypted_secret
        from vault.decrypted_secrets
        where name = 'strava_internal_token'
        limit 1
      )
    ),
    body := jsonb_build_object('action', 'refresh_tokens'),
    timeout_milliseconds := 120000
  ) as request_id;
  $job$
);
