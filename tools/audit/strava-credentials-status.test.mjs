import test from 'node:test';
import assert from 'node:assert/strict';

import {
  summarizeCredentialMetadata,
  validateSafeOutput,
} from './strava-credentials-status.mjs';

test('resume metadados sem expor referencias internas', () => {
  const output = summarizeCredentialMetadata({
    status: 'active',
    vault_only: true,
    token_version: 7,
    access_secret_id: 'access-ref',
    refresh_secret_id: 'refresh-ref',
  });

  assert.deepEqual(output, {
    status: 'active',
    vault_only: true,
    token_version: 7,
    vault_reference: 'complete',
  });
  assert.equal(JSON.stringify(output).includes('access-ref'), false);
  assert.equal(JSON.stringify(output).includes('refresh-ref'), false);
});

test('classifica referencia parcial e ausente', () => {
  assert.equal(summarizeCredentialMetadata({ access_secret_id: 'x' }).vault_reference, 'partial');
  assert.equal(summarizeCredentialMetadata({}).vault_reference, 'missing');
});

test('rejeita campos extras na saida publica', () => {
  assert.throws(
    () => validateSafeOutput({
      status: 'active',
      vault_only: true,
      token_version: 1,
      vault_reference: 'complete',
      forbidden_field: 'x',
    }),
    /Saída insegura/,
  );
});
