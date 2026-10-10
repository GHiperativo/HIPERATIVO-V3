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
PR: `#14 fix(strava): reparar ingestÃ£o P0 sem Apps Script`
Production code was deployed from this branch under the explicit Strava P0 RED authorization. The PR remains draft and unmerged. Merging `main` is a separate gate.

Next executable security step: design and validate a provider-safe athlete-token Vault write path for new OAuth/reconnects, then migrate existing credentials without ever surfacing their values to an assistant or repository.

## OAuth + Vault cutover | 2026-10-10

Status: production cutover completed for stored athlete credentials. Real new/reconnect authorization remains externally gated only by the Strava Authorization Callback Domain setting.

### What this replaces

The legacy OAuth/reconnect path in Google Apps Script and plaintext athlete OAuth columns are no longer the target runtime. `public.tokens_strava.access_token` and `public.tokens_strava.refresh_token` are retained only as compatibility columns and are constrained to remain `NULL`.

### New OAuth architecture

`trusted operator -> strava-oauth create_link -> one-use opaque state -> Strava authorize -> strava-oauth callback -> Vault -> token metadata`

Controls:

1. OAuth links can be issued only through an internally authenticated POST action.
2. The OAuth `state` is 32 cryptographically random bytes encoded as 64 hex characters.
3. Only SHA-256 of state is stored in `private.strava_oauth_tickets`.
4. Ticket lifetime is 15 minutes by default, bounded to 5-30 minutes.
5. A new ticket invalidates an older unconsumed ticket for the same athlete.
6. State is consumed exactly once before authorization-code exchange.
7. New flow requests only `read,activity:read_all`.
8. Callback verifies that `activity:read_all` was actually granted.
9. Existing `ath_id <-> strava_id` identity cannot be silently replaced.
10. A Strava identity already linked to another `ath_id` is rejected.
11. New and reconnect credentials are written directly to per-athlete Vault secrets.
12. OAuth success updates only token metadata plus `atletas.strava_ok/strava_id`.
13. Callback responses never contain OAuth credentials.

### Production migrations

The authoritative production sequence is:

- `20261010202656_strava_oauth_ticket_scaffold`
- `20261010202722_strava_legacy_tokens_to_vault`
- `20261010202746_strava_vault_runtime_phase1`
- `20261010202944_strava_oauth_vault_cutover`

Git mirrors those same four versions. Do not collapse them into one migration: the split reflects the actual production safety gates and rollback window.

### Legacy migration evidence

Before plaintext scrub:

- 28 active integrations had both Vault secret IDs.
- all 29 token rows, including the inactive integration, had both Vault secret IDs.
- server-side SHA-256 comparison produced 29/29 access-token matches and 29/29 refresh-token matches between legacy value and Vault value.
- no token value was returned to ChatGPT, terminal output, Git, Notion or PR comments.

After cutover:

- total token rows: 29.
- rows with both legacy OAuth columns NULL: 29.
- `vault_only=true`: 29.
- active integrations readable through the Vault-only access RPC: 28/28.
- refresh-secret claim test succeeded without returning secret material and was rolled back.
- active/stuck refresh leases after test: 0.
- `tokens_strava_oauth_plaintext_null_ck` enforces `access_token IS NULL AND refresh_token IS NULL` at database level.

### Edge versions after OAuth deploy

- `strava-oauth`: v1, public callback with internal authentication required for POST administration.
- `strava-sync`: v4.
- `strava-webhook`: v6.

Negative surface tests:

- unauthenticated `POST strava-oauth` -> HTTP 401.
- invalid OAuth state -> HTTP 400.

### External gate: Strava Authorization Callback Domain

A live authorization probe against the new Supabase callback currently returns HTTP 400 with Strava `redirect_uri invalid`. The Strava application settings must set the Authorization Callback Domain to:

`korlpbclqgmqvpbrungc.supabase.co`

The intended redirect URI is:

`https://korlpbclqgmqvpbrungc.supabase.co/functions/v1/strava-oauth`

This provider setting is not available through the public Strava API, so it cannot be changed by the Supabase runtime. Do not change client secret or webhook subscription merely to solve this callback-domain gate.

### Security advisor checkpoint

No new OAuth/Vault WARN was introduced. `private.strava_oauth_tickets` appears as RLS-enabled-without-policy INFO by design because it is not client-facing and no client role has privileges. Existing separate warnings remain: `pg_net` extension placement and Auth leaked-password protection disabled.

### Continuity checkpoint

Branch: `feat/strava-oauth-vault-20261010`
PR: `#15 feat(strava): OAuth seguro e Vault por atleta`
Base: `fix/strava-p0-red-20261010` / PR #14.
Merge to `main` remains a separate approval gate.

Next executable step after the callback-domain setting is updated: re-run the Strava authorization probe, issue one controlled reconnect link, complete one real OAuth round-trip, verify Vault-only persistence and then declare the new connection path operational.
