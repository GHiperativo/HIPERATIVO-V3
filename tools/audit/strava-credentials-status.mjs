#!/usr/bin/env node

const REQUIRED_ENV = ['SUPABASE_URL', 'SUPABASE_SERVICE_ROLE_KEY', 'STRAVA_ATH_ID'];

export function summarizeCredentialMetadata(row) {
  if (!row || typeof row !== 'object') {
    throw new Error('Registro de credencial Strava inválido.');
  }

  const hasAccessRef = Boolean(row.access_secret_id);
  const hasRefreshRef = Boolean(row.refresh_secret_id);

  let vaultReference = 'missing';
  if (hasAccessRef && hasRefreshRef) vaultReference = 'complete';
  else if (hasAccessRef || hasRefreshRef) vaultReference = 'partial';

  return {
    status: row.status ?? null,
    vault_only: Boolean(row.vault_only),
    token_version: Number(row.token_version ?? 0),
    vault_reference: vaultReference,
  };
}

export async function fetchCredentialMetadata({
  supabaseUrl,
  serviceRoleKey,
  athleteId,
  fetchImpl = fetch,
}) {
  const baseUrl = supabaseUrl.replace(/\/$/, '');
  const params = new URLSearchParams({
    select: 'status,vault_only,token_version,access_secret_id,refresh_secret_id',
    ath_id: `eq.${athleteId}`,
    limit: '1',
  });

  const response = await fetchImpl(`${baseUrl}/rest/v1/tokens_strava?${params.toString()}`, {
    headers: {
      apikey: serviceRoleKey,
      Authorization: `Bearer ${serviceRoleKey}`,
      Accept: 'application/json',
    },
  });

  if (!response.ok) {
    throw new Error(`Falha ao consultar metadados Strava: HTTP ${response.status}.`);
  }

  const rows = await response.json();
  if (!Array.isArray(rows) || rows.length === 0) {
    throw new Error('Nenhum registro Strava encontrado para o atleta informado.');
  }

  return rows[0];
}

export function validateSafeOutput(output) {
  const allowed = new Set(['status', 'vault_only', 'token_version', 'vault_reference']);
  const unexpected = Object.keys(output).filter((key) => !allowed.has(key));

  if (unexpected.length > 0) {
    throw new Error(`Saída insegura: campos não autorizados: ${unexpected.join(', ')}`);
  }

  return output;
}

async function main() {
  const missing = REQUIRED_ENV.filter((name) => !process.env[name]);
  if (missing.length > 0) {
    console.error(`Configuração ausente: ${missing.join(', ')}.`);
    process.exit(2);
  }

  const row = await fetchCredentialMetadata({
    supabaseUrl: process.env.SUPABASE_URL,
    serviceRoleKey: process.env.SUPABASE_SERVICE_ROLE_KEY,
    athleteId: process.env.STRAVA_ATH_ID,
  });

  const safeOutput = validateSafeOutput(summarizeCredentialMetadata(row));
  process.stdout.write(`${JSON.stringify(safeOutput, null, 2)}\n`);

  if (!safeOutput.vault_only || safeOutput.vault_reference !== 'complete') {
    process.exitCode = 1;
  }
}

if (import.meta.url === `file://${process.argv[1]}`) {
  main().catch((error) => {
    console.error(error.message);
    process.exit(1);
  });
}
