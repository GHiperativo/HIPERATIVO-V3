#!/usr/bin/env node

import { createHash } from 'node:crypto';
import { existsSync, readFileSync, readdirSync } from 'node:fs';
import { resolve } from 'node:path';

const BASELINE_PATH = 'supabase/migration-baseline.json';
const MIGRATIONS_DIR = 'supabase/migrations';

export function md5Sql(text) {
  return createHash('md5').update(text.trimEnd()).digest('hex');
}

export function readLocalMigrations({ cwd = process.cwd() } = {}) {
  const dir = resolve(cwd, MIGRATIONS_DIR);
  if (!existsSync(dir)) return [];
  return readdirSync(dir)
    .filter((name) => /^\d{14}_.+\.sql$/.test(name))
    .sort()
    .map((filename) => {
      const version = filename.slice(0, 14);
      const sql = readFileSync(resolve(dir, filename), 'utf8');
      return { version, filename, sql_md5: md5Sql(sql), sql_chars: sql.trimEnd().length };
    });
}

export function validateLocalAgainstBaseline({ baseline, localMigrations }) {
  const errors = [];
  const warnings = [];
  const byVersion = new Map(localMigrations.map((item) => [item.version, item]));
  const baselineByVersion = new Map(baseline.entries.map((item) => [item.version, item]));
  const legacyLocalOnly = new Set(baseline.legacy_local_only ?? []);

  for (const entry of baseline.entries) {
    const local = byVersion.get(entry.version);
    if (!local) continue;
    if (local.sql_md5 !== entry.sql_md5) {
      errors.push({ type: 'baseline_hash_mismatch', version: entry.version });
    }
  }

  for (const local of localMigrations) {
    if (baselineByVersion.has(local.version)) continue;
    if (local.version <= baseline.cutoff_version) {
      if (legacyLocalOnly.has(local.version)) warnings.push({ type: 'legacy_local_only', version: local.version });
      else errors.push({ type: 'unexpected_pre_cutoff_local_migration', version: local.version });
    }
  }

  return { errors, warnings };
}

export function compareRemote({ baseline, localMigrations, remoteEntries }) {
  const errors = [];
  const warnings = [];
  const baselineByVersion = new Map(baseline.entries.map((item) => [item.version, item]));
  const localByVersion = new Map(localMigrations.map((item) => [item.version, item]));
  const remoteByVersion = new Map(remoteEntries.map((item) => [item.version, item]));

  for (const expected of baseline.entries) {
    const remote = remoteByVersion.get(expected.version);
    if (!remote) {
      errors.push({ type: 'baseline_remote_missing', version: expected.version });
      continue;
    }
    if (remote.name !== expected.name || remote.sql_md5 !== expected.sql_md5) {
      errors.push({ type: 'baseline_remote_changed', version: expected.version });
    }
  }

  for (const remote of remoteEntries) {
    if (remote.version <= baseline.cutoff_version) continue;
    const local = localByVersion.get(remote.version);
    if (!local) {
      errors.push({ type: 'remote_missing_local_file', version: remote.version });
      continue;
    }
    if (local.sql_md5 !== remote.sql_md5) {
      errors.push({ type: 'remote_local_hash_mismatch', version: remote.version });
    }
  }

  for (const local of localMigrations) {
    if (local.version <= baseline.cutoff_version) continue;
    if (!remoteByVersion.has(local.version)) warnings.push({ type: 'local_pending_remote', version: local.version });
  }

  return { errors, warnings };
}

export async function fetchRemoteLedger({ supabaseUrl, serviceRoleKey, fetchImpl = fetch }) {
  if (!supabaseUrl && !serviceRoleKey) return { status: 'skipped', reason: 'credentials_not_configured', entries: [] };
  if (!supabaseUrl || !serviceRoleKey) return { status: 'error', reason: 'incomplete_credentials', entries: [] };

  const response = await fetchImpl(`${supabaseUrl.replace(/\/$/, '')}/rest/v1/rpc/platform_migration_ledger`, {
    method: 'POST',
    headers: {
      apikey: serviceRoleKey,
      Authorization: `Bearer ${serviceRoleKey}`,
      Accept: 'application/json',
      'Content-Type': 'application/json',
    },
    body: '{}',
  });

  if (!response.ok) return { status: 'error', reason: `http_${response.status}`, entries: [] };
  const entries = await response.json();
  if (!Array.isArray(entries)) return { status: 'error', reason: 'unexpected_response', entries: [] };
  return { status: 'ok', entries };
}

export async function buildMigrationAudit({ cwd = process.cwd(), env = process.env, fetchImpl = fetch } = {}) {
  const baseline = JSON.parse(readFileSync(resolve(cwd, BASELINE_PATH), 'utf8'));
  const localMigrations = readLocalMigrations({ cwd });
  const local = validateLocalAgainstBaseline({ baseline, localMigrations });
  const remoteFetch = await fetchRemoteLedger({
    supabaseUrl: env.SUPABASE_URL,
    serviceRoleKey: env.SUPABASE_SERVICE_ROLE_KEY,
    fetchImpl,
  });

  let remote = { errors: [], warnings: [] };
  if (remoteFetch.status === 'ok') remote = compareRemote({ baseline, localMigrations, remoteEntries: remoteFetch.entries });
  else if (remoteFetch.status === 'error') remote.errors.push({ type: remoteFetch.reason });

  const errors = [...local.errors, ...remote.errors];
  const warnings = [...local.warnings, ...remote.warnings];
  return {
    status: errors.length ? 'error' : 'ok',
    cutoff_version: baseline.cutoff_version,
    baseline_remote_count: baseline.remote_count,
    local_file_count: localMigrations.length,
    remote_check: remoteFetch.status,
    errors,
    warnings,
  };
}

async function main() {
  const report = await buildMigrationAudit();
  process.stdout.write(`${JSON.stringify(report, null, 2)}\n`);
  if (report.status !== 'ok') process.exitCode = 1;
}

if (import.meta.url === `file://${process.argv[1]}`) {
  main().catch((error) => {
    console.error(`Auditoria de migrations falhou: ${error.message}`);
    process.exit(1);
  });
}
