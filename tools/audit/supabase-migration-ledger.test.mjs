import test from 'node:test';
import assert from 'node:assert/strict';

import {
  compareRemote,
  fetchRemoteLedger,
  validateLocalAgainstBaseline,
} from './supabase-migration-ledger.mjs';

const baseline = {
  cutoff_version: '20261010234405',
  legacy_local_only: ['20260722155300'],
  entries: [
    { version: '20261010234405', name: 'cutoff', sql_md5: 'aaa' },
  ],
};

test('legado local conhecido vira warning, não erro', () => {
  const result = validateLocalAgainstBaseline({
    baseline,
    localMigrations: [
      { version: '20260722155300', sql_md5: 'legacy' },
      { version: '20261010234405', sql_md5: 'aaa' },
    ],
  });
  assert.equal(result.errors.length, 0);
  assert.equal(result.warnings[0].type, 'legacy_local_only');
});

test('mudança em migration congelada quebra auditoria', () => {
  const result = validateLocalAgainstBaseline({
    baseline,
    localMigrations: [{ version: '20261010234405', sql_md5: 'changed' }],
  });
  assert.equal(result.errors[0].type, 'baseline_hash_mismatch');
});

test('migration remota pós-corte sem arquivo local é erro', () => {
  const result = compareRemote({
    baseline,
    localMigrations: [{ version: '20261010234405', sql_md5: 'aaa' }],
    remoteEntries: [
      { version: '20261010234405', name: 'cutoff', sql_md5: 'aaa' },
      { version: '20261011010101', name: 'remote_only', sql_md5: 'bbb' },
    ],
  });
  assert.equal(result.errors.some((item) => item.type === 'remote_missing_local_file'), true);
});

test('migration local pós-corte ainda não aplicada é warning', () => {
  const result = compareRemote({
    baseline,
    localMigrations: [
      { version: '20261010234405', sql_md5: 'aaa' },
      { version: '20261011010101', sql_md5: 'bbb' },
    ],
    remoteEntries: [{ version: '20261010234405', name: 'cutoff', sql_md5: 'aaa' }],
  });
  assert.equal(result.errors.length, 0);
  assert.equal(result.warnings.some((item) => item.type === 'local_pending_remote'), true);
});

test('RPC remoto usa apenas endpoint dedicado', async () => {
  let request;
  const result = await fetchRemoteLedger({
    supabaseUrl: 'https://example.supabase.co/',
    serviceRoleKey: 'test-key',
    fetchImpl: async (url, options) => {
      request = { url, options };
      return { ok: true, status: 200, json: async () => [] };
    },
  });
  assert.equal(result.status, 'ok');
  assert.equal(request.url, 'https://example.supabase.co/rest/v1/rpc/platform_migration_ledger');
  assert.equal(request.options.method, 'POST');
});
