# Estrutura das Planilhas

> v0.2 — 08/06/2026

## V3 — Planilha Principal
ID: `1bI5pnt-HOAD5p8M2hqjEsU9P816hc94wy4mqx0J_xOM`

### CADASTRO (27 colunas)
`ID Atleta | Nome Completo | E-mail | WhatsApp | Data Nasc. | Sexo | Peso (kg) | Altura (cm) | Modalidade | Nível | Objetivo | Freq./Semana | Horário Pref. | Condições de Saúde | Histórico de Lesão | Medicamentos | Prova/Meta | Plano | Cidade | Estado | CPF | Canal de Origem | Data Cadastro | Strava Conectado | ID Strava | Status | Observações`

**Regra:** nome real da aba = `👤 CADASTRO` — sempre usar `H.SHEETS.CADASTRO`, nunca string hardcoded.

### ATIVIDADES (26 colunas)
`ID Execução | ID Atleta | Nome Atleta | Data | Tipo de Atividade | Fonte | ID Strava | Nome da Atividade | Tempo Mov. | Tempo Total | Distância (m) | Distância (km) | Vel. Média (m/s) | Vel. Média (km/h) | Pace Médio (s/km) | Pace Médio (min:ss) | FC Média | FC Máxima | Elevação (m) | Calorias | Cadência | Potência | Resumo Rota | Data Importação | RPE | Observações`

**Regra:** colunas 1-24 continuam compatíveis com a importação Strava. Campos manuais ficam depois para não interferir no fluxo de importação.

### MÉTRICAS (24 colunas)
`ID Atleta | Nome Atleta | Atualizado em | VO2máx Est. | Pace Médio (s/km) | Pace Rápido (s/km) | Pace Lento (s/km) | FC Máxima | FC Média | Vol./Semana (km) | Z1 Lento | Z1 Rápido | Z2 Lento | Z2 Rápido | Z3 Lento | Z3 Rápido | Z4 Lento | Z5 Mín | Perfil Manual | Volume Manual | Intensidade Manual | Origem dos Dados | Confiança | Observações`

**Fallback sem dados:** quando não existem corridas válidas nos últimos 28 dias, o cálculo usa os campos de múltipla escolha `Perfil Manual`, `Volume Manual` e `Intensidade Manual`. A coluna `Origem dos Dados` registra se a métrica veio de atividades recentes ou de estimativa manual.

## Locale: pt-BR
- Funções: `SE`, `CONT.SE`, `SEERRO`, `HOJE`, `MÉDIASE`, `SOMASE`
- Separador: **ponto-e-vírgula (;)**, nunca vírgula

## HUB Master
ID: `18CnUoiRGN55vwyhrMZ3lfxDAHF_1Dy33C-PyoOxQ84U`
