#!/usr/bin/env node

import { mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { buildHealthReport } from '../audit/platform-health.mjs';

const SNAPSHOT_SCHEMA_VERSION = 2;

export function sanitizeHealthReport(report) {
  return {
    schema_version: SNAPSHOT_SCHEMA_VERSION,
    kind: 'platform_health',
    captured_at: report.checked_at ?? new Date().toISOString(),
    overall: report.overall ?? 'unknown',
    verification: report.verification ?? 'unknown',
    coverage: {
      local: report.coverage?.local ?? 'unknown',
      supabase: report.coverage?.supabase ?? 'unknown',
      strava: report.coverage?.strava ?? 'unknown',
    },
    checks: {
      runtime: {
        status: report.local?.runtime?.status ?? 'unknown',
        node_major: report.local?.runtime?.node_major ?? null,
        expected_node_major: report.local?.runtime?.expected_node_major ?? null,
      },
      repository: {
        status: report.local?.repository?.status ?? 'unknown',
        required_paths_missing: Array.isArray(report.local?.repository?.required_paths_missing)
          ? report.local.repository.required_paths_missing
          : [],
        migration_count: report.local?.repository?.migration_count ?? null,
      },
      supabase: {
        status: report.supabase?.status ?? 'unknown',
        reason: report.supabase?.reason ?? null,
        http_status: report.supabase?.http_status ?? null,
      },
      strava_credentials: {
        status: report.strava_credentials?.status ?? 'unknown',
        reason: report.strava_credentials?.reason ?? null,
        credential_status: report.strava_credentials?.credential_status ?? null,
        vault_only: report.strava_credentials?.vault_only ?? null,
        token_version: report.strava_credentials?.token_version ?? null,
        vault_reference: report.strava_credentials?.vault_reference ?? null,
      },
    },
  };
}

export function safeTimestamp(value = new Date().toISOString()) {
  return value.replaceAll(':', '-').replace(/\.\d{3}Z$/, 'Z');
}

export function writeSnapshot(snapshot, { cwd = process.cwd() } = {}) {
  const directory = resolve(cwd, 'snapshots/platform-health');
  mkdirSync(directory, { recursive: true });
  const filename = `${safeTimestamp(snapshot.captured_at)}.json`;
  const path = resolve(directory, filename);
  writeFileSync(path, `${JSON.stringify(snapshot, null, 2)}\n`, { flag: 'wx' });
  return path;
}

export function diffSnapshots(before, after, prefix = '') {
  const ignored = new Set(['captured_at']);
  const keys = new Set([...Object.keys(before ?? {}), ...Object.keys(after ?? {})]);
  const changes = [];

  for (const key of [...keys].sort()) {
    if (ignored.has(key)) continue;
    const path = prefix ? `${prefix}.${key}` : key;
    const left = before?.[key];
    const right = after?.[key];

    const bothObjects = left && right
      && typeof left === 'object'
      && typeof right === 'object'
      && !Array.isArray(left)
      && !Array.isArray(right);

    if (bothObjects) {
      changes.push(...diffSnapshots(left, right, path));
      continue;
    }

    if (JSON.stringify(left) !== JSON.stringify(right)) {
      changes.push({ path, before: left ?? null, after: right ?? null });
    }
  }

  return changes;
}

function loadJson(path) {
  return JSON.parse(readFileSync(resolve(path), 'utf8'));
}

async function capture() {
  const report = await buildHealthReport();
  const snapshot = sanitizeHealthReport(report);
  const path = writeSnapshot(snapshot);
  process.stdout.write(`${path}\n`);
  if (snapshot.overall !== 'healthy') process.exitCode = 1;
}

function compare(beforePath, afterPath) {
  if (!beforePath || !afterPath) {
    throw new Error('Uso: npm run snapshot:compare -- <snapshot-anterior.json> <snapshot-atual.json>');
  }

  const before = loadJson(beforePath);
  const after = loadJson(afterPath);
  const changes = diffSnapshots(before, after);
  process.stdout.write(`${JSON.stringify({ changes_count: changes.length, changes }, null, 2)}\n`);
}

async function main() {
  const [command = 'capture', ...args] = process.argv.slice(2);
  if (command === 'capture') return capture();
  if (command === 'compare') return compare(args[0], args[1]);
  throw new Error(`Comando desconhecido: ${command}`);
}

if (import.meta.url === `file://${process.argv[1]}`) {
  main().catch((error) => {
    console.error(`Snapshot falhou: ${error.message}`);
    process.exit(1);
  });
}
