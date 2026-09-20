# Robô de estoque no TIMED

Lê os códigos de uma aba de uma planilha do Google e grava, para cada item,
estoque, consumo e valor unitário obtidos de um sistema TIMED.

Toda a configuração — endereço do sistema, planilha, aba e seções — vem de
variáveis de ambiente. Não há nenhum endereço ou identificador neste código.

## Variáveis

Secrets: `GOOGLE_CREDENTIALS`, `TIMED_BASE_URL`, `TIMED_USER`, `TIMED_PASS`, `SHEET_ID`
Variables: `ABA_DESTINO`, `COLUNA_GRUPO`, `SECOES_POR_GRUPO`

`SECOES_POR_GRUPO` tem o formato `Grupo=SEÇÃO|OUTRA SEÇÃO; Outro Grupo=SEÇÃO`.

## Execução

Diariamente pelo agendamento do workflow, ou manualmente em Actions → Run workflow.
