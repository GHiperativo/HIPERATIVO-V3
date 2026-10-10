-- Apply only after `strava_client_secret` exists in Supabase Vault and strava-sync is healthy.
-- One job replaces both proactive token renewal and recurring reconciliation from the retired Apps Script path.

select cron.schedule(
  'strava-sync-every-4h',
  '17 */4 * * *',
  $job$
  select net.http_post(
    url := 'https://korlpbclqgmqvpbrungc.supabase.co/functions/v1/strava-sync',
    headers := jsonb_build_object(
      'Content-Type', 'application/json',
      'X-Strava-Internal-Secret', (
        select decrypted_secret from vault.decrypted_secrets where name = 'strava_internal_token' limit 1
      )
    ),
    body := jsonb_build_object(
      'action', 'sync',
      'max_pages', 3,
      'per_page', 100
    ),
    timeout_milliseconds := 120000
  ) as request_id;
  $job$
);
