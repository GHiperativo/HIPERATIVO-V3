# Incident Assistant

Triagem assistida e **read-only** para incidentes da PLATAFORMA HIPERATIVO.

## Contrato

Os comandos desta pasta podem:
- ler estado local;
- consultar checks remotos já permitidos pelos auditores existentes;
- classificar severidade;
- resumir evidência segura;
- sugerir o próximo passo operacional.

Os comandos desta pasta **não podem**:
- aplicar migration;
- alterar dados;
- renovar ou gravar credenciais;
- fazer merge/revert;
- mudar baseline/snapshot;
- revelar tokens, secrets ou referências internas de Vault.

Uma execução só falha por erro do próprio assistente. Incidente confirmado é representado no JSON por `status`, não por automação corretiva.

## Comandos

```bash
npm run incident:strava
npm run incident:migrations
npm run incident:ci
```

### `incident:strava`

Uso: suspeita de falha de renovação/reconexão Strava.

Quando o ambiente administrativo estiver completo, usa somente a auditoria segura de metadados. A saída nunca contém access token, refresh token ou UUID de Vault.

Variáveis esperadas para cobertura remota:
- `SUPABASE_URL`
- `SUPABASE_SERVICE_ROLE_KEY`
- `STRAVA_ATH_ID`

Sem `STRAVA_ATH_ID`, o resultado é `inconclusive`, não sucesso.

### `incident:migrations`

Uso: suspeita de migration drift.

Reutiliza o gate oficial de migrations e distingue:
- erro de drift;
- warning conhecido;
- produção não verificada por ausência de credenciais;
- estado sem divergência detectada.

Sem credenciais remotas, o resultado é `inconclusive`.

### `incident:ci`

Uso: GitHub Actions vermelho depois de merge.

Combina:
- Platform Health;
- coverage/verification;
- auditoria de migrations;
- fundação local.

O comando não lê automaticamente os logs do GitHub Actions. Ele reduz o espaço de investigação e orienta a localizar o primeiro job vermelho no GitHub quando os checks estruturais locais/remotos não explicarem a falha.

## Formato de saída

Campos estáveis da primeira versão:

```json
{
  "mode": "read_only",
  "scenario": "ci",
  "status": "clear | warning | inconclusive | incident_confirmed",
  "severity": "SEV-2 | SEV-3",
  "summary": "...",
  "evidence": {},
  "next_steps": [],
  "prohibited": []
}
```

## Regra operacional

`clear` significa apenas que **os checks executados** não encontraram o incidente procurado.

`inconclusive` significa que faltou cobertura para concluir com segurança.

`incident_confirmed` significa que existe evidência suficiente para interromper a ação normal e seguir o playbook correspondente no Runbook Operacional.
