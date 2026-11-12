# VBA Extraído de Monitoramento Individual (A4UU)

## Overview
Código VBA e módulos de classes de planilha extraídos diretamente das pastas de trabalho de monitoramento individual das bases operacionais (ex: `Monitoramento individual_A4UU_v6.xlsm`). Contém funções de validação em tempo de inserção, leitura de parâmetros de rede, cálculo de prazo limite de SLA e rotina local de auditoria de preenchimento.

## Interface
- `[aFuncoes.bas:preencher_usuario_rede]` - Identifica o usuário da sessão do Windows via objeto COM `WScript.Network`
- `[aFuncoes.bas:validar_itemDescricao]` - Valida a consistência entre o código do item e a descrição da atividade cadastrada
- `[aFuncoes.bas:pesquisar_referenciaPrazo]` - Obtém o tempo regulamentar de resposta contratual para a atividade
- `[aFuncoes.bas:validar_contratoUF]` - Verifica se o contrato indicado cobre a Unidade Federativa da ordem
- `[aFuncoes.bas:verificar_localidadeAtiva]` - Checa se o centro ou ponto de atendimento consta como ativo
- `[aFuncoes.bas:calcular_dataSLA]` - Calcula data e hora exatas de vencimento do SLA em horas comerciais úteis
- `[modAnaliseMonitoramento.bas:AuditarMonitoramentoIndividual]` - Executa auditoria linha a linha na tabela estruturada do monitoramento gerando logs de inconsistência

## Invariants
- `[aFuncoes.bas:calcular_dataSLA]` Invariante de horário comercial: segunda a sexta das 08:00 às 17:00, com intervalo para almoço das 12:00 às 13:00; finais de semana e feriados não computam horas.
- `[modAnaliseMonitoramento.bas:AuditarMonitoramentoIndividual]` Quantidade solicitada deve coincidir com a atendida para conclusão válida da ordem; divergências disparam registro em log.
- `[modAnaliseMonitoramento.bas:AuditarMonitoramentoIndividual]` A relação entre Contrato e Localidade deve constar como autorizada na Tabela B.

## Side Effects
- `[aFuncoes.bas:preencher_usuario_rede]` Instancia componente externo `CreateObject("WScript.Network")`.
- `[aFuncoes.bas:validar_itemDescricao]` Acessa e lê células da `Planilha2` do monitoramento.
- `[modAnaliseMonitoramento.bas:AuditarMonitoramentoIndividual]` Grava logs de auditoria diretamente nas células da coluna de LOG da planilha ativa.

## Verification
- Run: `<test-command>`
