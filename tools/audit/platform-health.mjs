#!/usr/bin/env node

import { existsSync, readdirSync } from 'node:fs';
import { resolve } from 'node:path';
import {
  fetchCredentialMetadata,
  summarizeCredentialMetadata,
  validateSafeOutput,
} from './strava-credentials-status.mjs';

const REQUIRED_NODE_MAJOR = 24;
const REQUIRED_PATHS = [
  '.nvmrc',
  'package.json',
  'tools/README.md',
  'tools/audit/strava-credentials-status.mjs',
  '.github/workflows/tooling-check.yml',
];

export function inspectLocalFoundation({ cwd = process.cwd(), exists = existsSync, readDir = readdirSync } = {}) {
  const nodeMajor = Number(process.versions.node.split('.')[0]);
  const missingPaths = REQUIRED_PATHS.filter((path) => !exists(resolve(cwd, path)));

  const migrationsDir = resolve(cwd, 'supabase/migrations');
  let migrationCount = 0;
  if (exists(migrationsDir)) {
    migrationCount = readDir(migrationsDir, { withFileTypes: true })
      .filter((entry) => entry.isFile() && entry.name.endsWith('.sql'))
      .length;
  }

  return {
    runtime: {
      status: nodeMajor === REQUIRED_NODE_MAJOR ? 'ok' : 'error',
      node_major: nodeMajor,
      expected_node_major: REQUIRED_NODE_MAJOR,
    },
    repository: {
      status: missingPaths.length === 0 ? 'ok' : 'error',
      required_paths_missing: missingPaths,
      migration_count: migrationCount,
    },
  };
}

export async function checkSupabaseReachability({ supabaseUrl, serviceRoleKey, fetchImpl = fetch }) {
  if (!supabaseUrl && !serviceRoleKey) {
    return { status: 'skipped', reason: 'credentials_not_configured' };
  }
  if (!supabaseUrl || !serviceRoleKey) {
    return { status: 'error', reason: 'incomplete_credentials' };
  }

  const baseUrl = supabaseUrl.replace(/\/$/, '');
  const response = await fetchImpl(`${baseUrl}/rest/v1/rpc/platform_health_ping`, {
    method: 'POST',
    headers: {
      apikey: serviceRoleKey,
      Authorization: `Bearer ${serviceRoleKey}`,
      Accept: 'application/json',
      'Content-Type': 'application/json',
    },
    body: '{}',
  });

  if (!response.ok) {
    return { status: 'error', http_status: response.status };
  }

  const payload = await response.json();
  if (payload !== 'ok') {
    return { status: 'error', http_status: response.status, reason: 'unexpected_probe_response' };
  }

  return { status: 'ok', http_status: response.status };
}

export async function checkStravaCredentialHealth({ supabaseUrl, serviceRoleKey, athleteId, fetchImpl = fetch }) {
  if (!athleteId) {
    return { status: 'skipped', reason: 'athlete_not_configured' };
  }
  if (!supabaseUrl || !serviceRoleKey) {
    return { status: 'error', reason: 'supabase_credentials_missing' };
  }

  try {
    const row = await fetchCredentialMetadata({
      supabaseUrl,
      serviceRoleKey,
      athleteId,
      fetchImpl,
    });
    const summary = validateSafeOutput(summarizeCredentialMetadata(row));
    const healthy = summary.vault_only && summary.vault_reference === 'complete';
    const { status: credentialStatus, ...safeMetadata } = summary;

    return {
      status: healthy ? 'ok' : 'error',
      credential_status: credentialStatus,
      ...safeMetadata,
    };
  } catch (error) {
    return { status: 'error', reason: error.message };
  }
}

export function computeOverallStatus(report) {
  const statuses = [
    report.local.runtime.status,
    report.local.repository.status,
    report.supabase.status,
    report.strava_credentials.status,
  ];

  return statuses.includes('error') ? 'degraded' : 'healthy';
}

function remoteCoverage(check, incompleteReasons = []) {
  if (check?.status === 'skipped') return 'skipped';
  if (incompleteReasons.includes(check?.reason)) return 'incomplete';
  return 'complete';
}

export function computeCoverage(report) {
  return {
    local: 'complete',
    supabase: remoteCoverage(report.supabase, ['incomplete_credentials']),
    strava: remoteCoverage(report.strava_credentials, ['supabase_credentials_missing']),
  };
}

export function computeVerification(coverage) {
  const remote = [coverage.supabase, coverage.strava];
  if (remote.every((status) => status === 'complete')) return 'full';
  if (remote.every((status) => status === 'skipped')) return 'local_only';
  return 'partial';
}

export async function buildHealthReport({ env = process.env, cwd = process.cwd(), fetchImpl = fetch } = {}) {
  const local = inspectLocalFoundation({ cwd });
  const supabase = await checkSupabaseReachability({
    supabaseUrl: env.SUPABASE_URL,
    serviceRoleKey: env.SUPABASE_SERVICE_ROLE_KEY,
    fetchImpl,
  });
  const stravaCredentials = await checkStravaCredentialHealth({
    supabaseUrl: env.SUPABASE_URL,
    serviceRoleKey: env.SUPABASE_SERVICE_ROLE_KEY,
    athleteId: env.STRAVA_ATH_ID,
    fetchImpl,
  });

  const report = {
    checked_at: new Date().toISOString(),
    local,
    supabase,
    strava_credentials: stravaCredentials,
  };
  const coverage = computeCoverage(report);

  return {
    overall: computeOverallStatus(report),
    verification: computeVerification(coverage),
    coverage,
    ...report,
  };
}

async function main() {
  const report = await buildHealthReport();
  process.stdout.write(`${JSON.stringify(report, null, 2)}\n`);
  if (report.overall !== 'healthy') process.exitCode = 1;
}

if (import.meta.url === `file://${process.argv[1]}`) {
  main().catch((error) => {
    console.error(`Platform health falhou: ${error.message}`);
    process.exit(1);
  });
}
