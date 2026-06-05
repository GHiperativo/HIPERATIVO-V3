# Regras de Negócio

> v0.1 — 05/06/2026

## Atletas
- ID: `ATH` + 6 dígitos (ex: `ATH384729`)
- Strava é opcional no cadastro
- 1 atleta = 1 linha em CADASTRO (chave: ATH_ID)

## Strava OAuth
- Client ID: `153043`
- Tokens NUNCA em células
- Refresh automático antes de cada chamada à API

## Importação
- Automática: trigger a cada 4h
- Manual: menu Atividades > Importar todos

## Segurança
- Tokens em PropertiesService
- Scripts: try/catch + log obrigatório
- Nunca processar linhas sem ATH_ID válido
- Dados que NUNCA vão para HUB: access_token, refresh_token, JSON bruto
