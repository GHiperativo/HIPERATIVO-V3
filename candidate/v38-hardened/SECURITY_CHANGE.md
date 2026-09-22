# Hardening da v38

Base: snapshot exato do deployment v38.

## Alteração

Foi removida exclusivamente a rota `?manutencao=...&fn=...` do `doGet`.
Essa rota aceitava chamadas anônimas e utilizava uma senha fixa presente no
próprio código.

## Preservado

- callback OAuth Strava com `code` e `state`;
- reconexão por `?reconectar=ATH_ID`;
- página individual por `?atleta=ATH_ID`;
- cadastro e páginas públicas;
- funções internas de manutenção, sem exposição pelo roteador público;
- leitura, refresh e persistência dos tokens existentes.

## Implantação

Não publicar sobre a v10 OAuth. Quando aprovado, criar uma nova versão a partir
deste candidato e atualizar somente o deployment de manutenção atualmente em
v38. Não executar funções de configuração de gatilhos durante a implantação.
