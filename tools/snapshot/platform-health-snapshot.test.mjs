import test from 'node:test';
import assert from 'node:assert/strict';
import {
  diffSnapshots,
  safeTimestamp,
  sanitizeHealthReport,
} from './platform-health-snapshot.mjs';

test('sanitize remove campos não autorizados e preserva coverage', () => {
  const snapshot = sanitizeHealthReport({
    checked_at: '2026-10-10T23:30:00.000Z',
    overall: 'healthy',
    verification: 'full',
    coverage: { local: 'complete', supabase: 'complete', strava: 'complete' },
    local: {
      runtime: { status: 'ok', node_major: 24, expected_node_major: 24, secret: 'x' },
      repository: { status: 'ok', required_paths_missing: [], migration_count: 39 },
    },
    supabase: { status: 'ok', http_status: 200, service_role_key: 'never' },
    strava_credentials: {
      status: 'ok',
      credential_status: 'Renovado',
      vault_only: true,
      token_version: 1,
      vault_reference: 'complete',
      access_secret_id: 'uuid',
      refresh_token: 'never',
    },
  });

  const serialized = JSON.stringify(snapshot);
  assert.equal(snapshot.schema_version, 2);
  assert.equal(serialized.includes('service_role_key'), false);
  assert.equal(serialized.includes('access_secret_id'), false);
  assert.equal(serialized.includes('refresh_token'), false);
  assert.equal(snapshot.checks.strava_credentials.vault_reference, 'complete');
  assert.equal(snapshot.verification, 'full');
  assert.deepEqual(snapshot.coverage, { local: 'complete', supabase: 'complete', strava: 'complete' });
});

test('snapshot antigo sem coverage permanece sanitizável como unknown', () => {
  const snapshot = sanitizeHealthReport({ overall: 'healthy', local: {}, supabase: {}, strava_credentials: {} });
  assert.equal(snapshot.verification, 'unknown');
  assert.deepEqual(snapshot.coverage, { local: 'unknown', supabase: 'unknown', strava: 'unknown' });
});

test('timestamp vira nome de arquivo seguro', () => {
  assert.equal(
    safeTimestamp('2026-10-10T23:30:00.123Z'),
    '2026-10-10T23-30-00Z',
  );
});

test('comparação ignora captured_at e reporta mudanças reais', () => {
  const before = {
    captured_at: '2026-10-10T10:00:00Z',
    overall: 'healthy',
    verification: 'full',
    checks: { supabase: { status: 'ok' } },
  };
  const after = {
    captured_at: '2026-10-11T10:00:00Z',
    overall: 'degraded',
    verification: 'partial',
    checks: { supabase: { status: 'error' } },
  };

  assert.deepEqual(diffSnapshots(before, after), [
    { path: 'checks.supabase.status', before: 'ok', after: 'error' },
    { path: 'overall', before: 'healthy', after: 'degraded' },
    { path: 'verification', before: 'full', after: 'partial' },
  ]);
});
