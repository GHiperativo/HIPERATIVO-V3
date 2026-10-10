import test from 'node:test';
import assert from 'node:assert/strict';
import {
  checkSupabaseReachability,
  checkStravaCredentialHealth,
  computeOverallStatus,
} from './platform-health.mjs';

test('Supabase sem credenciais fica skipped', async () => {
  const result = await checkSupabaseReachability({});
  assert.deepEqual(result, { status: 'skipped', reason: 'credentials_not_configured' });
});

test('Supabase com configuração incompleta falha', async () => {
  const result = await checkSupabaseReachability({ supabaseUrl: 'https://example.supabase.co' });
  assert.equal(result.status, 'error');
  assert.equal(result.reason, 'incomplete_credentials');
});

test('auditoria Strava preserva status do check sem expor referências internas', async () => {
  const fetchImpl = async () => ({
    ok: true,
    status: 200,
    async json() {
      return [{
        status: 'Renovado',
        vault_only: true,
        token_version: 3,
        access_secret_id: 'uuid-a',
        refresh_secret_id: 'uuid-b',
      }];
    },
  });

  const result = await checkStravaCredentialHealth({
    supabaseUrl: 'https://example.supabase.co',
    serviceRoleKey: 'test-key',
    athleteId: '123',
    fetchImpl,
  });

  assert.deepEqual(result, {
    status: 'ok',
    credential_status: 'Renovado',
    vault_only: true,
    token_version: 3,
    vault_reference: 'complete',
  });
  assert.equal('access_secret_id' in result, false);
  assert.equal('refresh_secret_id' in result, false);
});

test('overall fica degraded diante de qualquer erro', () => {
  const report = {
    local: { runtime: { status: 'ok' }, repository: { status: 'ok' } },
    supabase: { status: 'ok' },
    strava_credentials: { status: 'error' },
  };
  assert.equal(computeOverallStatus(report), 'degraded');
});
