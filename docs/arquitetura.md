# Arquitetura

> v0.1 — 05/06/2026

## Fluxo Principal
```
Atleta → cadastro.html → WebApp.gs (salvarCadastroAjax)
  └─ Aba CADASTRO + email de boas-vindas

Strava OAuth → _gerarUrlOAuth → callback → PropertiesService
  └─ STRAVA_TOKEN_STATUS (status sem tokens)

Trigger 4h → triggerImportacaoAutomatica → Strava API
  └─ Aba ATIVIDADES
```

## Papel de Cada Ferramenta
| Ferramenta | Responsabilidade |
|---|---|
| HIPERATIVO V3 | Cadastro, Strava, atividades, métricas, prescrição |
| HUB Master | CRM, financeiro, leads, comunicação |
| SISRUN | Prescrição final |
| Notion/QG | Estratégia e documentação |

## Segurança
- Tokens em **PropertiesService** (nunca em células)
- Aba TOKENS: oculta e protegida
- Scripts: sempre try/catch + log
- Chave de integração HUB↔V3: `ID_Atleta`
