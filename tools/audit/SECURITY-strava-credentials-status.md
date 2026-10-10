# Segurança da auditoria Strava

A ferramenta consulta apenas metadados necessários para provar o estado da migração de credenciais para o Vault.

Ela não seleciona `access_token` nem `refresh_token`, não imprime as referências UUID do Vault e não faz escrita no banco.

A chave de serviço é recebida somente por variável de ambiente e deve permanecer em armazenamento seguro. Nunca registrar essa variável em logs, exemplos ou arquivos versionados.
