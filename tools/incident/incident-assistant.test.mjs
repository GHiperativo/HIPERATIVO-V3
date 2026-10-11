import test from 'node:test';
import assert from 'node:assert/strict';
import { diagnoseCi, diagnoseMigrations, diagnoseStrava } from './incident-assistant.mjs';

function jsonResponse(body, { ok = true, status = 200 } = {}) {
  return {
    ok,
    status,
    async json() { return body; },
  };
}

test('strava sem athlete id fica inconclusivo e read-only', async () => {
  const report = await diagnoseStrava({ env: {} });
  assert.equal(report.mode, 'read_only');
  assert.equal(report.status, 'inconclusive');
  assert.equal(report.evidence.credential_check, 'skipped');
});

test('strava saudável expõe apenas metadados permitidos', async () => {
  const fetchImpl = async () => jsonResponse([{
    status: 'Renovado',
    vault_only: true,
    token_version: 2,
    access_secret_id: 'ref-a',
    refresh_secret_id: 'ref-b',
  }]);
  const report = await diagnoseStrava({
    env: { SUPABASE_URL: 'https://example.supabase.co', SUPABASE_SERVICE_ROLE_KEY: 'secret', STRAVA_ATH_ID: 'ath' },
    fetchImpl,
  });
  assert.equal(report.status, 'clear');
  assert.equal(report.evidence.credential_check.vault_only, true);
  assert.equal(report.evidence.credential_check.vault_reference, 'complete');
  assert.equal('access_secret_id' in report.evidence.credential_check, false);
  assert.equal('refresh_secret_id' in report.evidence.credential_check, false);
});

test('migration audit remoto ausente fica inconclusivo, não clear', async () => {
  const report = await diagnoseMigrations({ env: {}, cwd: process.cwd() });
  assert.equal(report.mode, 'read_only');
  assert.equal(report.status, 'inconclusive');
  assert.equal(report.evidence.remote_check, 'skipped');
});

test('ci saudável com cobertura incompleta fica inconclusivo', async () => {
  const report = await diagnoseCi({ env: {}, cwd: process.cwd() });
  assert.equal(report.mode, 'read_only');
  assert.equal(report.status, 'inconclusive');
  assert.equal(report.severity, 'SEV-3');
  assert.equal(report.evidence.platform_health.verification, 'local_only');
});

test('ci detecta erro remoto de Supabase como incidente', async () => {
  const fetchImpl = async (url) => {
    if (url.includes('platform_health_ping')) return jsonResponse({}, { ok: false, status: 500 });
    if (url.includes('platform_migration_ledger')) return jsonResponse([], { ok: false, status: 500 });
    throw new Error(`URL inesperada: ${url}`);
  };
  const report = await diagnoseCi({
    env: { SUPABASE_URL: 'https://example.supabase.co', SUPABASE_SERVICE_ROLE_KEY: 'secret' },
    cwd: process.cwd(),
    fetchImpl,
  });
  assert.equal(report.status, 'incident_confirmed');
  assert.equal(report.severity, 'SEV-2');
});
