# Fonte da verdade — HIPERATIVO V3

Atualizado em 24/07/2026 após a implantação visual e a validação ponta a ponta.

## Produção

- Planilha: `HIPERATIVO V3` (`1bI5pnt-HOAD5p8M2hqjEsU9P816hc94wy4mqx0J_xOM`).
- Código implantado: `apps-script/live/`.
- Web App público: implantação `AKfycbyNrmCjUxRYjUVjKMBPK7n_qFMknas2yfEXVciAMAIcOO1dr-9zH5haSuiGxMlIIO4Fqg`, atualizada para a versão 58 em 24/07/2026.
- A atualização preservou o URL público usado no cadastro, na conexão Strava e no callback OAuth.
- `apps-script/src/` e `snapshots/` são referências históricas; não são origem para publicação.
- Projeto Supabase ativo: `hiperativo-v3` (`korlpbclqgmqvpbrungc`).

## Fluxo operacional

1. `cadastro.html` chama `salvarCadastroAjax`.
2. O cadastro é salvo em `👤 CADASTRO`, com bloqueio conservador de duplicidade por código, e-mail e WhatsApp.
3. Se o atleta declarou que usa Strava, `_gerarUrlOAuth` cria o link individual.
4. O callback troca o código por tokens e chama somente `persistirCredenciaisStrava`.
5. A persistência grava a cópia local, a aba oculta `🔐 TOKENS` e o Supabase, sem substituir refresh token válido por vazio.
6. O webhook oficial do Strava entrega eventos à Edge Function `strava-webhook`, que valida, deduplica e registra a fila no Supabase.
7. A Edge Function chama o `doPost` do Apps Script com assinatura HMAC; o script busca somente a atividade indicada no evento.
8. Criações, alterações e exclusões são refletidas em `🏃 STRAVA_RAW`, `🏃 ATIVIDADES` e nas métricas do atleta.
9. A reconciliação tradicional permanece como segurança diária, com paginação, cursor e respeito aos limites da Strava.
10. A resposta original fica em `🏃 STRAVA_RAW` (oculta); a consulta operacional fica em `🏃 ATIVIDADES`.
11. Métricas, análise, ranking e painel usam `🏃 ATIVIDADES`.
12. Cada evento processado atualiza imediatamente as visões derivadas do atleta e o painel operacional.

## Hierarquia de credenciais Strava

- Primário: Apps Script/planilha, com atualização atômica pelo fluxo central.
- Segurança: Script Properties e Supabase.
- Regra inviolável: nunca apagar ou substituir um refresh token válido por valor ausente/inválido.
- Reconexão só é indicada após todas as fontes falharem e o refresh ser realmente inválido/revogado.
- Atletas marcados como “Não” em Strava não entram em monitor, fila nem alertas.
- Validação de 24/07/2026: 14 linhas locais com refresh token válido correspondem a 13 atletas únicos protegidos no Supabase; a linha histórica repetida foi preservada por segurança.

## Automações essenciais

- Atividades Strava: webhook oficial em tempo real, sem varredura por atleta.
- Fila de segurança: Supabase, com deduplicação e nova tentativa controlada.
- Novos cadastros, planos, feedbacks e inputs manuais: `onEdit` simples atualiza somente as visões afetadas, sem acionador instalável.
- Painel inteligente: atualizado após eventos Strava e entradas manuais, com indicadores, atividades recentes, semáforo de treino e central de conexão.
- Alertas Strava: falha temporária entra em espera protegida; pedido de reconexão só é gerado após revogação oficial ou ausência real do refresh token.
- `triggerImportacaoAutomatica`: reconciliação; mesmo que existam acionadores antigos, o ciclo pesado roda no máximo uma vez em 20 horas.
- `monitorarStravaOk`: monitor diário, sem declarar reconexão enquanto existir refresh token recuperável.
- `limparLogsAntigos`: manutenção semanal.

Acionadores instalados por contas Google diferentes só podem ser excluídos pela conta que os criou. O código compartilhado neutraliza execuções pesadas repetidas em todo o projeto.

## Abas

- Visíveis e operacionais: `📊 PAINEL`, `👤 CADASTRO`, `📲 WHATSAPP STRAVA`, `🏃 ATIVIDADES`, `📈 MÉTRICAS`, `📅 PLANO SEMANAL`, `💬 FEEDBACK`, `🔬 ANÁLISE`, `🏆 RANKING`, `📝 INPUT MANUAL`.
- Técnicas/ocultas: `🏃 STRAVA_RAW`, `🔐 TOKENS`, logs, filas, backups e abas legadas.
- `🏃 ATIVIDADES_CONVERTIDAS` é legado oculto; não alimenta o fluxo atual.
- O painel anterior foi preservado em `📊 PAINEL_BACKUP_VISUAL_20260722`, oculta.

## Salvaguardas

- Cadastros duplicados são bloqueados, nunca mesclados automaticamente.
- Uma conta Strava não pode ser vinculada silenciosamente a dois `ATH_ID`.
- Dados improváveis recebem alerta na coluna `STATUS`; a atividade não é descartada.
- Após três falhas seguidas, o atleta entra em espera protegida com nova tentativa automática; cursor e tokens são preservados.
- A fila do webhook não armazena tokens; as tabelas novas têm RLS e permissões exclusivas de `service_role`.
- Eventos repetidos são deduplicados e chamadas ao Apps Script exigem assinatura válida, janela de tempo e nonce inédito.
- VO₂ permanece vazio até existir teste validado.
- PSE de atividade alimenta sRPE; PSE e FC de repouso do feedback alimentam apenas o sinal de recuperação.

## Mudanças que exigem decisão explícita

- Mesclar cadastros já duplicados.
- Apagar abas, tokens, históricos ou backups.
- Liberar acesso público/cliente às tabelas do Supabase via políticas RLS.
- Trocar automaticamente o dono de um ID Strava duplicado.
