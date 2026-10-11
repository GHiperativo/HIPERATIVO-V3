import test from 'node:test';
import assert from 'node:assert/strict';

import { evaluateRegression } from './platform-health-regression.mjs';

function platform(overrides = {}) {
  return {
    kind: 'platform_health',
    overall: 'healthy',
    verification: 'full',
    coverage: { local: 'complete', supabase: 'complete', strava: 'complete' },
    checks: {
      runtime: { status: 'ok', node_major: 24, expected_node_major: 24 },
      repository: { status: 'ok', required_paths_missing: [], migration_count: 40 },
      supabase: { status: 'ok' },
      strava_credentials: {
        status: 'ok',
        vault_only: true,
        token_version: 3,
        vault_reference: 'complete',
      },
    },
    ...overrides,
  };
}

test('snapshot saudável permanece clean', () => {
  const result = evaluateRegression(platform(), platform());
  assert.equal(result.status, 'clean');
  assert.equal(result.errors, 0);
});

test('detecta regressão de Vault e migrations', () => {
  const before = platform();
  const after = platform({
    checks: {
      ...before.checks,
      repository: { status: 'ok', required_paths_missing: [], migration_count: 39 },
      strava_credentials: {
        status: 'ok',
        vault_only: false,
        token_version: 2,
        vault_reference: 'partial',
      },
    },
  });
  const result = evaluateRegression(before, after);
  const codes = result.issues.map((entry) => entry.code);
  assert.equal(result.status, 'regression');
  assert.ok(codes.includes('MIGRATION_COUNT_DECREASED'));
  assert.ok(codes.includes('VAULT_ONLY_DISABLED'));
  assert.ok(codes.includes('VAULT_REFERENCE_REGRESSION'));
  assert.ok(codes.includes('TOKEN_VERSION_DECREASED'));
});

test('perda de coverage vira warning explícito, não erro', () => {
  const before = platform();
  const after = platform({
    verification: 'partial',
    coverage: { local: 'complete', supabase: 'complete', strava: 'skipped' },
    checks: {
      ...before.checks,
      strava_credentials: { status: 'skipped' },
    },
  });
  const result = evaluateRegression(before, after);
  const codes = result.issues.map((entry) => entry.code);
  assert.equal(result.status, 'warning');
  assert.equal(result.errors, 0);
  assert.ok(codes.includes('COVERAGE_REGRESSION'));
  assert.ok(codes.includes('STATUS_UNVERIFIED'));
});

test('credenciais remotas skipped viram warning, não regressão', () => {
  const before = platform();
  const after = platform({
    verification: 'local_only',
    coverage: { local: 'complete', supabase: 'skipped', strava: 'skipped' },
    checks: {
      ...before.checks,
      supabase: { status: 'skipped' },
      strava_credentials: { status: 'skipped' },
    },
  });
  const result = evaluateRegression(before, after);
  assert.equal(result.status, 'warning');
  assert.equal(result.errors, 0);
  assert.equal(result.warnings, 3);
  assert.ok(result.issues.some((entry) => entry.code === 'COVERAGE_REGRESSION'));
});

test('baseline verificado aceita primeiro snapshot automático sem credenciais como warning', () => {
  const before = {
    kind: 'verified_baseline',
    overall: 'healthy',
    checks: {
      github_ci: { node_runtime_check: 'success', tooling_check: 'success' },
      supabase: { status: 'ok', migration_count: 39, strava_token_rows: 30, vault_only_rows: 30, complete_vault_refs: 30 },
    },
  };
  const after = platform({
    verification: 'local_only',
    coverage: { local: 'complete', supabase: 'skipped', strava: 'skipped' },
    checks: {
      ...platform().checks,
      supabase: { status: 'skipped' },
      strava_credentials: { status: 'skipped' },
    },
  });
  const result = evaluateRegression(before, after);
  assert.equal(result.status, 'warning');
  assert.equal(result.errors, 0);
  assert.ok(result.issues.some((entry) => entry.code === 'COVERAGE_REGRESSION'));
});

test('snapshot legado sem verification não gera falsa regressão de coverage', () => {
  const before = platform();
  delete before.verification;
  delete before.coverage;
  const result = evaluateRegression(before, platform());
  assert.equal(result.status, 'clean');
});

test('baseline agregado falha quando cobertura Vault fica incompleta', () => {
  const before = {
    kind: 'verified_baseline',
    overall: 'healthy',
    checks: {
      github_ci: { node_runtime_check: 'success', tooling_check: 'success' },
      supabase: { status: 'ok', migration_count: 39, strava_token_rows: 30, vault_only_rows: 30, complete_vault_refs: 30 },
    },
  };
  const after = structuredClone(before);
  after.checks.supabase.complete_vault_refs = 29;
  const result = evaluateRegression(before, after);
  assert.equal(result.status, 'regression');
  assert.ok(result.issues.some((entry) => entry.code === 'VAULT_REFERENCE_COVERAGE_LOST'));
});
