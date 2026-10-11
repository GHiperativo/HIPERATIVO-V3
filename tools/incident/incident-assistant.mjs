#!/usr/bin/env node

import { buildHealthReport, checkStravaCredentialHealth } from '../audit/platform-health.mjs';
import { buildMigrationAudit } from '../audit/supabase-migration-ledger.mjs';

const READ_ONLY = 'read_only';

function baseReport(scenario) {
  return {
    mode: READ_ONLY,
    scenario,
    checked_at: new Date().toISOString(),
  };
}

function classifyMigrationAudit(audit) {
  if (audit.errors.length > 0) {
    return {
      status: 'incident_confirmed',
      severity: 'SEV-2',
      summary: 'O gate de migrations encontrou divergência que exige investigação antes de novas migrations.',
    };
  }
  if (audit.remote_check !== 'ok') {
    return {
      status: 'inconclusive',
      severity: 'SEV-3',
      summary: 'A validação local passou, mas a produção não foi completamente verificada.',
    };
  }
  if (audit.warnings.length > 0) {
    return {
      status: 'warning',
      severity: 'SEV-3',
      summary: 'Não há drift crítico confirmado, mas existem warnings que precisam ser interpretados.',
    };
  }
  return {
    status: 'clear',
    severity: 'SEV-3',
    summary: 'Nenhum drift foi detectado pelo gate de migrations.',
  };
}

export async function diagnoseStrava({ env = process.env, fetchImpl = fetch } = {}) {
  const report = baseReport('strava');

  if (!env.STRAVA_ATH_ID) {
    return {
      ...report,
      status: 'inconclusive',
      severity: 'SEV-3',
      summary: 'O check Strava não pode ser executado sem STRAVA_ATH_ID no ambiente administrativo.',
      evidence: { credential_check: 'skipped', reason: 'athlete_not_configured' },
      next_steps: [
        'Executar novamente em ambiente administrativo seguro com STRAVA_ATH_ID configurado.',
        'Não inserir IDs, tokens ou secrets em argumentos, logs ou documentação.',
      ],
      prohibited: ['Exibir access_token/refresh_token.', 'Gravar credencial fora do Vault.', 'Renovar credenciais em lote como teste.'],
    };
  }

  const check = await checkStravaCredentialHealth({
    supabaseUrl: env.SUPABASE_URL,
    serviceRoleKey: env.SUPABASE_SERVICE_ROLE_KEY,
    athleteId: env.STRAVA_ATH_ID,
    fetchImpl,
  });

  const healthy = check.status === 'ok';
  return {
    ...report,
    status: healthy ? 'clear' : 'incident_confirmed',
    severity: healthy ? 'SEV-3' : 'SEV-2',
    summary: healthy
      ? 'Os metadados seguros da credencial Strava estão coerentes.'
      : 'A credencial Strava não atende ao estado seguro esperado ou não pôde ser validada.',
    evidence: { credential_check: check },
    next_steps: healthy
      ? ['Se a renovação funcional ainda falha, investigar OAuth/código sem abrir tokens.', 'Comparar com o último commit conhecido como bom.']
      : ['Interromper qualquer renovação em lote.', 'Separar falha de Vault, configuração e fluxo OAuth antes de corrigir.'],
    prohibited: ['Exibir access_token/refresh_token.', 'Copiar secrets para chat, issue ou log.', 'Atualizar token manualmente em coluna legada.'],
  };
}

export async function diagnoseMigrations({ env = process.env, cwd = process.cwd(), fetchImpl = fetch } = {}) {
  const audit = await buildMigrationAudit({ env, cwd, fetchImpl });
  const classification = classifyMigrationAudit(audit);
  return {
    ...baseReport('migrations'),
    ...classification,
    evidence: {
      cutoff_version: audit.cutoff_version,
      baseline_remote_count: audit.baseline_remote_count,
      local_file_count: audit.local_file_count,
      remote_check: audit.remote_check,
      errors: audit.errors,
      warnings: audit.warnings,
    },
    next_steps: audit.errors.length
      ? ['Congelar novas migrations.', 'Comparar produção e repositório usando somente o ledger sanitizado.', 'Recuperar arquivo exato quando houver migration remota pós-cutoff sem local.']
      : audit.remote_check !== 'ok'
        ? ['Executar novamente com cobertura remota em ambiente administrativo seguro antes de concluir ausência de drift.']
        : ['Nenhuma ação corretiva necessária; manter o gate ativo.'],
    prohibited: ['Atualizar baseline para esconder drift.', 'Editar migration já aplicada.', 'Resetar schema_migrations.', 'Reaplicar migrations históricas.'],
  };
}

export async function diagnoseCi({ env = process.env, cwd = process.cwd(), fetchImpl = fetch } = {}) {
  const [health, migrations] = await Promise.all([
    buildHealthReport({ env, cwd, fetchImpl }),
    buildMigrationAudit({ env, cwd, fetchImpl }),
  ]);

  const migrationClass = classifyMigrationAudit(migrations);
  const critical = health.overall === 'degraded' || migrationClass.status === 'incident_confirmed';
  const incomplete = health.verification !== 'full' || migrations.remote_check !== 'ok';

  return {
    ...baseReport('ci'),
    status: critical ? 'incident_confirmed' : incomplete ? 'inconclusive' : 'clear',
    severity: critical ? 'SEV-2' : 'SEV-3',
    summary: critical
      ? 'Há regressão verificável em health ou migrations; o próximo passo é localizar o primeiro check que falhou.'
      : incomplete
        ? 'Não há falha confirmada nos checks executados, mas a cobertura é incompleta.'
        : 'Os checks estruturais disponíveis não apontam regressão.',
    evidence: {
      platform_health: {
        overall: health.overall,
        verification: health.verification,
        coverage: health.coverage,
        local: health.local,
        supabase: health.supabase,
        strava_credentials: health.strava_credentials,
      },
      migrations: {
        status: migrations.status,
        remote_check: migrations.remote_check,
        errors: migrations.errors,
        warnings: migrations.warnings,
      },
    },
    next_steps: critical
      ? ['Identificar o primeiro job/check vermelho no GitHub Actions.', 'Comparar o merge com o último commit conhecido como bom.', 'Abrir PR corretivo mínimo ou revert por novo commit quando a falha estiver localizada.']
      : incomplete
        ? ['Completar a cobertura necessária antes de declarar a plataforma totalmente saudável.', 'Se o GitHub Actions estiver vermelho, usar o nome do primeiro check falho para direcionar a investigação.']
        : ['Se o GitHub Actions ainda estiver vermelho, investigar o job específico; estes checks estruturais não reproduziram a falha.'],
    prohibited: ['Corrigir diretamente em main.', 'Desabilitar check para obter verde.', 'Alterar baseline/snapshot apenas para passar CI.', 'Force-push para apagar evidência.'],
  };
}

export async function runIncidentScenario(scenario, options = {}) {
  if (scenario === 'strava') return diagnoseStrava(options);
  if (scenario === 'migrations') return diagnoseMigrations(options);
  if (scenario === 'ci') return diagnoseCi(options);
  throw new Error(`Cenário desconhecido: ${scenario || '(vazio)'}. Use strava, migrations ou ci.`);
}

async function main() {
  const scenario = process.argv[2];
  const report = await runIncidentScenario(scenario);
  process.stdout.write(`${JSON.stringify(report, null, 2)}\n`);
}

if (import.meta.url === `file://${process.argv[1]}`) {
  main().catch((error) => {
    console.error(`Assistente de incidente falhou: ${error.message}`);
    process.exit(1);
  });
}
