# HIPERATIVO V3

> Sistema de gestão de atletas — Strava OAuth, prescrição de treinos, Apps Script + Google Sheets

## Objetivo

Centralizar o ciclo técnico completo do Grupo Hiperativo: cadastro de atletas, conexão com Strava, importação de atividades, cálculo de zonas (Daniels/Seiler), prescrição de treinos (CCC) e exportação para SISRUN.

## Stack

- **Google Sheets** — banco de dados operacional
- **Google Apps Script** — backend e automações
- **Strava API (OAuth 2.0)** — integração de atividades
- **SISRUN** — exportação de prescrições
- **Notion** — documentação e QG estratégico

## Fontes Oficiais

| Recurso | Link |
|---|---|
| Planilha HIPERATIVO V3 | [Google Sheets](https://docs.google.com/spreadsheets/d/1bI5pnt-HOAD5p8M2hqjEsU9P816hc94wy4mqx0J_xOM) |
| Apps Script | [Editor](https://script.google.com/u/0/home/projects/1cm4IbBr3IsAjHtJJNOAAJQf3EnFD66ES5hMmiCaejBQCkNpVOnp5Zy_P/edit) |
| WebApp (produção) | [Deploy](https://script.google.com/macros/s/AKfycbyNrmCjUxRYjUVjKMBPK7n_qFMknas2yfEXVciAMAIcOO1dr-9zH5haSuiGxMlIIO4Fqg/exec) |
| Manual de Uso | [Notion](https://app.notion.com/p/358c485fe84381c3b2adde0a52542fce) |
| Log Técnico | [Notion](https://app.notion.com/p/374c485fe8438052aa41f57eed9d5977) |
| Drive V3 | [Google Drive](https://drive.google.com/drive/folders/1vbJbLcG5yCCjuCC3L_DCh-Fcj7QqjuJG) |

## Estrutura

```
HIPERATIVO-V3/
├── docs/               # Documentação técnica
├── apps-script/        # Scripts versionados
├── sheets/             # Schemas das abas
├── prompts/            # Prompts de IA
└── .github/            # Templates
```

## Status

| Item | Status |
|---|---|
| Planilha operacional | ✅ Ativo |
| OAuth Strava | ✅ OK |
| Importação automática (4h) | ✅ Ativo |
| Tokens em PropertiesService | ✅ OK |
| Integração SISRUN | 🔄 Em dev |
| Dashboard admin | ⏳ Pendente |

---
*Responsável: Crhystiano Heliodoro — contato@ghiperativo.com.br | 05/06/2026*
