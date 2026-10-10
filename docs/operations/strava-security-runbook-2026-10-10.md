# Strava P0 Security & Continuity Runbook

Status: operational in production as of 2026-10-10. This document is a durable technical handoff. The Notion QG Hiperativo remains the official decision record.

## What this replaces

The runtime no longer depends on the retired Google Apps Script web app for Strava ingestion, token renewal or reconciliation. The old Apps Script URL returned HTTP 404 and the legacy Google project/sheet was not reliably accessible.

## What is new

Runtime path:

`Strava -> strava-webhook Edge Function -> Supabase`

Recovery/reconciliation path:

`pg_cron -> strava-sync Edge Function -> Strava API -> Supabase`

The Strava application client secret and the internal cron-to-Edge secret are stored in Supabase Vault. No application secret is committed to Git.

## Production state

- Supabase project: `hiperativo-v3` (`korlpbclqgmqvpbrungc`).
- Active Strava push subscription: `376140`.
- Callback: `https://korlpbclqgmqvpbrungc.supabase.co/functions/v1/strava-webhook`.
- `strava-webhook`: direct event processing, idempotent event queue, subscription-id validation and stale-event rejection.
- `strava-sync`: health, subscription bootstrap/validation, token maintenance, reconciliation and controlled backfill.
- Activity reconciliation cron: `strava-sync-every-4h` at `17 */4 * * *`.
- Token maintenance cron: `strava-token-maintenance-hourly` at `7 * * * *`.
- Current token population at hardening close: 28 active, 1 inactive.

## Token-safety controls

1. `public.tokens_strava` has RLS enabled and only a `service_role` policy.
2. `anon` and `authenticated` table grants were explicitly revoked. Table ACL is restricted to postgres/service_role.
3. Bulk token listings return metadata only. OAuth secret values are no longer loaded for all athletes at once.
4. The access token for one athlete is loaded only when that athlete is processed.
5. The refresh token is released to the worker only after it obtains an exclusive refresh lease.
6. Refresh leases are row-locked and time-bounded, preventing webhook and cron workers from rotating the same refresh token concurrently.
7. A successful refresh atomically persists access token, latest refresh token, expiry, version and audit timestamps.
8. `token_version` lets a waiting worker detect that another worker already completed a rotation.
9. Failed refresh attempts persist timestamp/error metadata and clear the lease when safe to retry.
10. Worker commit logic retries and verifies version/expiry after an ambiguous persistence response.
11. Tokens are proactively checked hourly and are refreshed only inside the one-hour safety window, avoiding unnecessary rotation.
12. Health telemetry exposes counts only: active/expired/expiring tokens, recent refresh errors and active leases. It never returns token values.

## Why the lease is mandatory

Strava rotates refresh tokens. A successful OAuth refresh can return a new refresh token that supersedes the previous one. Two workers refreshing the same athlete concurrently can therefore invalidate each other's credential path. The database lease serializes rotation per `ath_id`.

## Validation evidence on 2026-10-10

P0 recovery:

- 28/28 active athletes processed successfully.
- 155 activities received in recovery backfill.
- 145 inserted, 10 updated.
- 0 identity conflicts.
- 0 rate-limit stop.
- 0 duplicated `strava_id`.

Post-hardening:

- Health HTTP 200.
- 28 active tokens.
- 0 active tokens expiring within one hour at validation time.
- 0 unresolved refresh errors.
- 0 active/stuck refresh leases.
- Manual `refresh_tokens`: 28/28 OK, 0 failures, 28 already valid, 0 unnecessary refreshes.
- Controlled activity sync: 28/28 OK, 0 failures, 10 existing activities updated, 0 inserted, 0 identity conflicts.
- Invalid webhook verification token: HTTP 403.

## Applied production migrations

The following P0 versions are intentionally mirrored in Git with the same production version numbers:

- `20261010192900_strava_p0_runtime.sql`
- `20261010194408_strava_p0_schedule.sql`
- `20261010194421_strava_p0_health_after_secret.sql`
- `20261010194436_strava_p0_subscription_bootstrap.sql`
- `20261010194451_strava_p0_backfill_20260922.sql`
- `20261010194609_strava_p0_final_health.sql`
- `20261010195744_strava_token_refresh_hardening.sql`
- `20261010200617_strava_token_maintenance_hourly.sql`

The four one-time operational invocation files are historical no-op markers in Git. This is deliberate: health calls, subscription creation and backfills must never execute automatically during a database restore or `db push`.

## Known limitations and next security gate

### Athlete OAuth secrets are not yet physically migrated to Vault

`access_token` and `refresh_token` still exist in `public.tokens_strava`. Their exposure is now constrained by RLS plus explicit service-role-only ACLs and minimized runtime reads, but the target state is stronger:

- metadata/version/lease remain in the database;
- access token and refresh token become separate encrypted Vault secrets per athlete;
- the OAuth callback/reconnect path writes directly to that secure store;
- legacy plaintext columns are scrubbed only after dual-read validation proves the Vault runtime.

The automated migration of the existing token values was intentionally not bypassed when the execution security layer blocked secret transport. Do not copy tokens through chat, logs, temporary files or Git to finish this migration.

### `pg_net` advisor warning

Supabase currently reports `pg_net` as an extension installed in `public`. In this project the extension is marked `extrelocatable = false`, while callable HTTP functions live under schema `net`. Do not drop/recreate it merely to silence the advisor while the Strava pipeline depends on it. Treat extension relocation as a separate infrastructure gate with restore/rollback testing.

### Legacy migration drift outside this P0

`supabase migration list` shows older remote migrations that predate the current repository migration directory. All Strava P0 migrations listed above are aligned local/remote. Do not conflate the older V3 migration-history cleanup with Strava token recovery.

## Incident checks

Use metadata-only checks first:

- Edge Functions active and expected versions.
- `strava-sync` health returns HTTP 200.
- `subscription_id` matches the active Strava subscription.
- active token count, expired count, expiring-within-one-hour count.
- unresolved refresh errors.
- active refresh leases.
- cron jobs active.
- latest activity/import timestamps.
- duplicate `strava_id` count must remain zero.

Never print or export `access_token`, `refresh_token`, Vault decrypted secrets, client secret or internal cron secret during diagnosis.

## Recovery order

1. Read health and logs without secret disclosure.
2. Verify subscription and callback.
3. Verify cron jobs and recent executions.
4. Check token metadata and refresh error audit.
5. Run controlled token maintenance if necessary.
6. Run a narrow reconciliation window before any broad backfill.
7. Only use a backfill after confirming idempotency and rate-limit headroom.
8. Record result in Notion and the active PR before leaving the incident.

## Continuity checkpoint

Branch: `fix/strava-p0-red-20261010`  
PR: `#14 fix(strava): reparar ingestão P0 sem Apps Script`  
Production code was deployed from this branch under the explicit Strava P0 RED authorization. The PR remains draft and unmerged. Merging `main` is a separate gate.

Next executable security step: design and validate a provider-safe athlete-token Vault write path for new OAuth/reconnects, then migrate existing credentials without ever surfacing their values to an assistant or repository.
