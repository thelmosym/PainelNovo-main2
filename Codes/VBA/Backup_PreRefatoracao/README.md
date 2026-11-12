# VBA Backup Pré-Refatoração

## Overview
Conjunto de módulos VBA legados preservados como histórico da arquitetura anterior à refatoração unificada do Painel de Controle e Auditoria. Contém rotinas de geração de cadernos de medição por contrato, atualização síncrona de conexões OLEDB Power Query e auditoria inicial de dados.

## Interface
- `[modAuditoriaLog.bas:AuditarDadosEGerarLog] Public Sub AuditarDadosEGerarLog()` - Dispara o motor de auditoria em lote sobre abas consolidadas
- `[modAuditoriaLog.bas:AuditarAbaAtiva] Public Sub AuditarAbaAtiva()` - Executa a auditoria de regras de negócio sobre a planilha ativa
- `[modAuditoriaLog.bas:LimparRelatorioLog] Public Sub LimparRelatorioLog()` - Limpa a planilha `LOG_CRITICAS` e restaura a mensagem inicial padrão
- `[modulo1.vba:GerarArquivosPorContrato] Public Sub GerarArquivosPorContrato()` - Exporta em lote os cadernos `.xlsx` oficiais dos contratos PA-LT2, IRON-LT1-RJ, IRON-LT1-SP e IRON-LT1-ES
- `[Modulo2.vba:AtualizarTodasConsultasPowerQuery] Public Sub AtualizarTodasConsultasPowerQuery()` - Força atualização síncrona de todas as consultas Power Query via conexões OLEDB da pasta de trabalho

## Invariants
- `[modAuditoriaLog.bas:AuditarDadosEGerarLog]` Requer a existência prévia de uma planilha de dados consolidados (`DADOS`, `Painel_T2M` ou `DADOS_IRON-LT1`).
- `[modAuditoriaLog.bas:AuditarAbaAtiva]` A aba ativa não pode ser `LOG_CRITICAS` ou `Painel`.
- `[modAuditoriaLog.bas:ExecutarMotorAuditoria]` A tabela ou intervalo de dados deve conter cabeçalho e pelo menos uma linha de dados (`lastRow >= 2`).
- `[modAuditoriaLog.bas:ExecutarMotorAuditoria]` Campos mandatórios (`Contrato`, `Descricao da atividade`, `Item`, `Codigo da solicitacao`) não podem ser vazios.
- `[modAuditoriaLog.bas:ExecutarMotorAuditoria]` `D. fechamento` deve ser cronologicamente maior ou igual a `D. abertura`.
- `[modAuditoriaLog.bas:ExecutarMotorAuditoria]` `Qtd. Solicitada` deve ser estritamente maior que 0 e `Qtd. Atendida` maior ou igual a 0.
- `[modAuditoriaLog.bas:ExecutarMotorAuditoria]` `Chave solicitante` deve ter exatamente 4 caracteres alfanuméricos quando informada.
- `[modulo1.vba:GerarArquivosPorContrato]` A pasta de trabalho deve estar salva em disco (`wbOrigem.Path <> ""`) e a subpasta `MEDIÇÃO` deve existir ou ser passível de criação.
- `[modulo1.vba:ListaContratos]` A ordenação dos contratos (`PA-LT2`, `IRON-LT1-RJ`, `IRON-LT1-SP`, `IRON-LT1-ES`) corresponde estritamente 1:1 com as linhas de link K7:K10 na aba `Painel`.

## Side Effects
- `[modAuditoriaLog.bas:ExecutarMotorAuditoria]` Altera e restaura `ScreenUpdating`, `EnableEvents` e `Calculation` (xlCalculationManual).
- `[modAuditoriaLog.bas:ExecutarMotorAuditoria]` Escreve texto de log diretamente na coluna `LOG_AUDITORIA` da planilha auditada.
- `[modAuditoriaLog.bas:GerarRelatorioAbaLog]` Cria ou sobrescreve a aba `LOG_CRITICAS` com cabeçalhos, zebra fill, estilos de células, bordas, AutoFilter e FreezePanes.
- `[modAuditoriaLog.bas:AtualizarStatusPainel]` Atualiza a célula `Painel!B14` com texto, cor de fundo (verde/âmbar/vermelho) e formatação negrito.
- `[modulo1.vba:ObterOuCriarSubpasta]` Cria a pasta de sistema de arquivos `MEDIÇÃO` via `MkDir`.
- `[modulo1.vba:GerarArquivoDoContrato]` Gera e salva novos arquivos de caderno `.xlsx` no disco e aplica rótulos MIP via `wb.SensitivityLabel.SetLabel`.
- `[modulo1.vba:EscreverLinkArquivo]` Grava hiperlinks para os cadernos gerados nas células `Painel!K7:K10`.
- `[Modulo2.vba:AtualizarTodasConsultasPowerQuery]` Dispara conexões de rede/banco e atualiza o status na célula `Painel!B12`.

## Verification
- Run: `<test-command>`
