# Arquitetura

> A descrição operacional atual está em [fonte-da-verdade.md](fonte-da-verdade.md).

> v0.1 — 05/06/2026

## Fluxo Principal
```
Atleta → cadastro.html → WebApp.gs (salvarCadastroAjax)
  └─ Aba CADASTRO + email de boas-vindas

Strava OAuth → _gerarUrlOAuth → callback → persistirCredenciaisStrava
  └─ Planilha/PropertiesService + cópia de segurança no Supabase

Strava Webhook → Edge Function Supabase → fila durável → doPost assinado
  └─ Busca da atividade exata → STRAVA_RAW + ATIVIDADES + MÉTRICAS
      └─ Atualização imediata → ANÁLISE + RANKING + PAINEL

Reconciliação diária → triggerImportacaoAutomatica → fila + Strava API
  └─ Recupera falhas sem repetir a varredura pesada mais de 1 vez em 20h

Entrada manual → onEdit simples (CADASTRO, PLANO, FEEDBACK ou INPUT)
  └─ Recalcula somente a visão afetada → PAINEL atualizado sem trigger instalável

Falha Strava temporária → alerta protegido + fila Supabase
  └─ Sem pedir reconexão; reconexão só após revogação oficial ou token realmente ausente
```

## Papel de Cada Ferramenta
| Ferramenta | Responsabilidade |
|---|---|
| HIPERATIVO V3 | Cadastro, Strava, atividades, métricas, prescrição |
| HUB Master | CRM, financeiro, leads, comunicação |
| SISRUN | Prescrição final |
| Notion/QG | Estratégia e documentação |

## Segurança
- Tokens no fluxo protegido do Apps Script, com aba TOKENS oculta e protegida
- Cópia de segurança dos refresh tokens no Supabase, sem acesso público
- Chamadas Edge Function → Apps Script assinadas com HMAC e proteção contra repetição
- Fila idempotente no Supabase; nenhum token é armazenado nos eventos
- Atualizações visuais atuam apenas em abas derivadas e preservam um backup oculto do painel
- Scripts: sempre try/catch + log
- Chave de integração HUB↔V3: `ID_Atleta`
