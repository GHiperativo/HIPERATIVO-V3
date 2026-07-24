# Problemas Conhecidos

> 05/06/2026

## Resolvidos ✅
| Bug | Fix |
|---|---|
| PAINEL #ERROR! row 6 — fórmulas en-US | Corrigido manualmente 04/06/2026 |
| CRM J2 #NAME? — SE com vírgula | Corrigido 04/06/2026 |
| salvarCadastroAjax salvava null | H.SHEETS.CADASTRO (emoji) |
| HtmlService incompatível com google.script.run | salvarCadastroAjax retorna JSON simples |
| Auditoria de triggers podia apagar automações válidas | Transformada em diagnóstico somente leitura em 22/07/2026 |
| Fallback Supabase apontava para projeto inexistente | URL alinhada ao projeto `hiperativo-v3` |
| ATH_ID automático podia colidir em alta concorrência | UUID curto + trava de cadastro + verificação da base |
| Fila podia insistir no mesmo atleta com erro | Backoff progressivo e espera protegida com recuperação automática |
| Atualização dependia de varreduras frequentes | Webhook Strava ativo, fila durável no Supabase e atualização exata por atividade |
| Acionadores antigos de contas diferentes repetiam ciclos pesados | Trava global limita a reconciliação completa a uma vez em 20 horas |
| Dois registros legados incompletos poluíam ATIVIDADES/RAW | IDs confirmados como inexistentes (404) no Strava e removidos das duas abas |
| PAINEL rows 11-20 #ERROR! | Painel substituído por visão operacional validada em 24/07/2026; backup anterior preservado e oculto |
| Atualizações visuais dependiam de execução manual | Eventos Strava e entradas manuais passaram a atualizar as visões afetadas automaticamente |

## Pendentes ⚠️
| Bug | Impacto |
|---|---|
| STRAVA_CLIENT_SECRET não configurado | OAuth pode falhar em renovação |
| Duplicatas antigas de cadastro | Precisam de revisão humana antes de qualquer mesclagem; novas colisões estão bloqueadas |
| Uma linha histórica repetida na aba oculta TOKENS | Não reduz a cobertura: 13 atletas únicos estão protegidos no Supabase; remover só após revisão específica |
| Checkboxes em Financeiro_MP (linhas vazias) | Performance |
| Acionadores antigos pertencem a outra conta Google | Permanecem instalados, porém neutralizados pela trava diária; exclusão visual exige entrar na conta criadora |

## Ação Imediata do Admin
Configurar CLIENT_SECRET: Planilha > menu HIPERATIVO > Configurações > Configurar credenciais Strava
