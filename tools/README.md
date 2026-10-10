# Tools | Plataforma Hiperativo

Caixa de ferramentas técnica da Plataforma Hiperativo, executada com Node.js 24 LTS.

## Objetivo

Centralizar scripts administrativos, auditorias, migrações e utilitários que apoiam a operação da plataforma sem misturar tooling com código de produção ou com o legado Apps Script.

## Estrutura

- `platform/`: utilitários transversais da plataforma.
- `supabase/`: auditorias, migrações assistidas e verificações relacionadas ao Supabase.
- `strava/`: integrações, diagnósticos e utilitários relacionados ao Strava.
- `migration/`: ferramentas temporárias e repetíveis de migração entre sistemas.
- `audit/`: verificações somente leitura, integridade e segurança.
- `scripts/`: utilitários Node.js genéricos de apoio.

## Regras

1. Ferramentas devem ser reproduzíveis e executáveis por comando documentado.
2. Operações destrutivas não podem ser o comportamento padrão.
3. Segredos, tokens e credenciais nunca entram no repositório ou em logs.
4. Scripts que escrevem em produção devem exigir intenção explícita e validações de segurança.
5. Sempre que possível, oferecer modo somente leitura ou `dry-run` antes de qualquer alteração.
6. O legado continua funcionando durante a transição; tooling novo não deve criar dependência obrigatória prematura.

## Runtime

Use Node.js 24 LTS, conforme `.nvmrc` e `package.json` na raiz do repositório.
