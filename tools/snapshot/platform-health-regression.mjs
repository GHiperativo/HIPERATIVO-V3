#!/usr/bin/env node

import { readdirSync, readFileSync } from 'node:fs';
import { resolve } from 'node:path';

function issue(severity, code, path, before, after, message) {
  return { severity, code, path, before: before ?? null, after: after ?? null, message };
}

function pushTransition(issues, { path, before, after, errorValues = ['error'], warningValues = ['skipped', 'unknown'] }) {
  if (before !== 'ok') return;
  if (errorValues.includes(after)) {
    issues.push(issue('error', 'STATUS_REGRESSION', path, before, after, `${path} regrediu de ok para ${after}.`));
  } else if (warningValues.includes(after)) {
    issues.push(issue('warning', 'STATUS_UNVERIFIED', path, before, after, `${path} deixou de ser verificável neste snapshot.`));
  }
}

export function evaluateRegression(before, after) {
  const issues = [];

  if (!before || !after) {
    return { status: 'invalid', errors: 1, warnings: 0, issues: [issue('error', 'INVALID_INPUT', '', null, null, 'Snapshots inválidos.')] };
  }

  if (before.overall === 'healthy' && after.overall !== 'healthy') {
    issues.push(issue('error', 'OVERALL_REGRESSION', 'overall', before.overall, after.overall, 'A plataforma deixou o estado healthy.'));
  }

  if (before.kind === 'verified_baseline' && after.kind === 'platform_health') {
    const runtime = after.checks?.runtime;
    const repository = after.checks?.repository;
    const supabase = after.checks?.supabase;
    const strava = after.checks?.strava_credentials;

    if (runtime?.status !== 'ok' || runtime?.node_major !== runtime?.expected_node_major) {
      issues.push(issue('error', 'RUNTIME_REGRESSION', 'checks.runtime', 'verified', runtime ?? null, 'Runtime Node.js não corresponde ao runtime esperado.'));
    }
    if (repository?.status !== 'ok' || (repository?.required_paths_missing?.length ?? 0) > 0) {
      issues.push(issue('error', 'REPOSITORY_REGRESSION', 'checks.repository', 'verified', repository ?? null, 'Fundação mínima do repositório está incompleta.'));
    }
    if (supabase?.status === 'error') {
      issues.push(issue('error', 'SUPABASE_REGRESSION', 'checks.supabase.status', 'ok', supabase.status, 'Supabase retornou erro.'));
    } else if (supabase?.status === 'skipped') {
      issues.push(issue('warning', 'SUPABASE_UNVERIFIED', 'checks.supabase.status', 'verified', 'skipped', 'Supabase não foi verificado neste snapshot por ausência de credenciais no ambiente.'));
    }
    if (strava?.status === 'error') {
      issues.push(issue('error', 'STRAVA_REGRESSION', 'checks.strava_credentials.status', 'verified', strava.status, 'Credenciais Strava falharam na auditoria.'));
    } else if (strava?.status === 'skipped') {
      issues.push(issue('warning', 'STRAVA_UNVERIFIED', 'checks.strava_credentials.status', 'verified', 'skipped', 'Credenciais Strava não foram verificadas neste snapshot.'));
    }
  }

  if (before.kind === 'platform_health' && after.kind === 'platform_health') {
    const b = before.checks ?? {};
    const a = after.checks ?? {};

    if (b.runtime?.status === 'ok' && (a.runtime?.status !== 'ok' || a.runtime?.node_major !== a.runtime?.expected_node_major)) {
      issues.push(issue('error', 'RUNTIME_REGRESSION', 'checks.runtime', b.runtime, a.runtime, 'Runtime Node.js regrediu.'));
    }

    const beforeMissing = b.repository?.required_paths_missing?.length ?? 0;
    const afterMissing = a.repository?.required_paths_missing?.length ?? 0;
    if (b.repository?.status === 'ok' && (a.repository?.status !== 'ok' || afterMissing > beforeMissing)) {
      issues.push(issue('error', 'REPOSITORY_REGRESSION', 'checks.repository', b.repository, a.repository, 'Arquivos obrigatórios do tooling deixaram de estar disponíveis.'));
    }

    if (Number.isFinite(b.repository?.migration_count) && Number.isFinite(a.repository?.migration_count)
      && a.repository.migration_count < b.repository.migration_count) {
      issues.push(issue('error', 'MIGRATION_COUNT_DECREASED', 'checks.repository.migration_count', b.repository.migration_count, a.repository.migration_count, 'A quantidade de migrations SQL versionadas diminuiu.'));
    }

    pushTransition(issues, { path: 'checks.supabase.status', before: b.supabase?.status, after: a.supabase?.status });
    pushTransition(issues, { path: 'checks.strava_credentials.status', before: b.strava_credentials?.status, after: a.strava_credentials?.status });

    if (a.strava_credentials?.status === 'ok') {
      if (b.strava_credentials?.vault_only === true && a.strava_credentials?.vault_only !== true) {
        issues.push(issue('error', 'VAULT_ONLY_DISABLED', 'checks.strava_credentials.vault_only', true, a.strava_credentials?.vault_only, 'vault_only deixou de estar ativo.'));
      }
      if (b.strava_credentials?.vault_reference === 'complete' && a.strava_credentials?.vault_reference !== 'complete') {
        issues.push(issue('error', 'VAULT_REFERENCE_REGRESSION', 'checks.strava_credentials.vault_reference', 'complete', a.strava_credentials?.vault_reference, 'Referências do Vault deixaram de estar completas.'));
      }
      if (Number.isFinite(b.strava_credentials?.token_version) && Number.isFinite(a.strava_credentials?.token_version)
        && a.strava_credentials.token_version < b.strava_credentials.token_version) {
        issues.push(issue('error', 'TOKEN_VERSION_DECREASED', 'checks.strava_credentials.token_version', b.strava_credentials.token_version, a.strava_credentials.token_version, 'token_version diminuiu.'));
      }
    }
  }

  if (before.kind === 'verified_baseline' && after.kind === 'verified_baseline') {
    const b = before.checks ?? {};
    const a = after.checks ?? {};
    if (b.github_ci?.node_runtime_check === 'success' && a.github_ci?.node_runtime_check !== 'success') {
      issues.push(issue('error', 'NODE_CI_REGRESSION', 'checks.github_ci.node_runtime_check', b.github_ci.node_runtime_check, a.github_ci?.node_runtime_check, 'Node Runtime Check deixou de passar.'));
    }
    if (b.github_ci?.tooling_check === 'success' && a.github_ci?.tooling_check !== 'success') {
      issues.push(issue('error', 'TOOLING_CI_REGRESSION', 'checks.github_ci.tooling_check', b.github_ci.tooling_check, a.github_ci?.tooling_check, 'Tooling Check deixou de passar.'));
    }
    if (Number.isFinite(b.supabase?.migration_count) && Number.isFinite(a.supabase?.migration_count)
      && a.supabase.migration_count < b.supabase.migration_count) {
      issues.push(issue('error', 'MIGRATION_COUNT_DECREASED', 'checks.supabase.migration_count', b.supabase.migration_count, a.supabase.migration_count, 'A quantidade de migrations registradas diminuiu.'));
    }
    if (Number.isFinite(a.supabase?.strava_token_rows)) {
      if (a.supabase.vault_only_rows < a.supabase.strava_token_rows) {
        issues.push(issue('error', 'VAULT_ONLY_COVERAGE_LOST', 'checks.supabase.vault_only_rows', b.supabase?.vault_only_rows, a.supabase.vault_only_rows, 'Nem todos os registros Strava estão em vault_only.'));
      }
      if (a.supabase.complete_vault_refs < a.supabase.strava_token_rows) {
        issues.push(issue('error', 'VAULT_REFERENCE_COVERAGE_LOST', 'checks.supabase.complete_vault_refs', b.supabase?.complete_vault_refs, a.supabase.complete_vault_refs, 'Nem todos os registros Strava possuem referências completas do Vault.'));
      }
    }
  }

  const errors = issues.filter((item) => item.severity === 'error').length;
  const warnings = issues.filter((item) => item.severity === 'warning').length;
  return { status: errors > 0 ? 'regression' : warnings > 0 ? 'warning' : 'clean', errors, warnings, issues };
}

export function loadSnapshot(path) {
  return JSON.parse(readFileSync(resolve(path), 'utf8'));
}

export function findLatestSnapshots({ cwd = process.cwd() } = {}) {
  const directory = resolve(cwd, 'snapshots/platform-health');
  const snapshots = readdirSync(directory)
    .filter((name) => name.endsWith('.json'))
    .map((name) => ({ name, path: resolve(directory, name) }))
    .map((entry) => ({ ...entry, data: loadSnapshot(entry.path) }))
    .sort((left, right) => {
      const byDate = String(left.data.captured_at ?? '').localeCompare(String(right.data.captured_at ?? ''));
      return byDate || left.name.localeCompare(right.name);
    });
  return snapshots.slice(-2);
}

function printResult(result, context = {}) {
  process.stdout.write(`${JSON.stringify({ ...context, ...result }, null, 2)}\n`);
  if (result.errors > 0) process.exitCode = 1;
}

function explicit(beforePath, afterPath) {
  if (!beforePath || !afterPath) {
    throw new Error('Uso: npm run snapshot:regression -- <snapshot-anterior.json> <snapshot-atual.json>');
  }
  printResult(evaluateRegression(loadSnapshot(beforePath), loadSnapshot(afterPath)), { before: beforePath, after: afterPath });
}

function auto() {
  const snapshots = findLatestSnapshots();
  if (snapshots.length < 2) {
    printResult({ status: 'skipped', errors: 0, warnings: 0, issues: [] }, { reason: 'fewer_than_two_snapshots' });
    return;
  }
  const [before, after] = snapshots;
  printResult(evaluateRegression(before.data, after.data), { before: before.name, after: after.name });
}

function main() {
  const args = process.argv.slice(2);
  if (args[0] === '--auto') return auto();
  return explicit(args[0], args[1]);
}

if (import.meta.url === `file://${process.argv[1]}`) {
  try {
    main();
  } catch (error) {
    console.error(`Regressão não pôde ser avaliada: ${error.message}`);
    process.exit(1);
  }
}
