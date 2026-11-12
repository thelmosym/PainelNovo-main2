# VBA Funções Novas (Produção Refatorada)

## Overview
Suíte moderna e otimizada de módulos VBA para o Painel de Controle Principal (`Painel de controle MemoriaPetrobras V.8.6.xlsm`). Desenvolvida para processamento em memória através de matrizes bidimensionais ($O(1)$ de acesso), proporcionando alta performance na auditoria de conformidade contratual, geração do Demonstrativo de FDM, auditoria de fontes de Power Query e exportação de cadernos de medição.

## Interface
- `[modAuditoriaLog.bas:AuditarDadosEGerarLog]` - Executa o motor completo de auditoria com relatório interativo de inconsistências
- `[modAuditoriaLog.bas:AuditarDadosSilencioso]` - Executa a rotina de auditoria de dados em modo batch sem bloquear com diálogos modais
- `[modAuditoriaLog.bas:AuditarAbaAtiva]` - Valida as regras de negócio diretamente sobre os registros da planilha ativa
- `[modAuditoriaLog.bas:LimparRelatorioLog]` - Reseta a planilha `LOG_CRITICAS` e restaura a mensagem padrão
- `[modAuditoriaFontesPQ.bas:ScanearFontesPowerQuery]` - Inspeciona o código M de todas as consultas Power Query e gera relatório estruturado de caminhos
- `[modDemonstrativoFDM.bas:GerarDemonstrativoCalculoFDM]` - Gera a planilha de cálculo e memória de cálculo de SLA e FDM com coluna de diagnósticos
- `[modGerarArquivos.bas:GerarArquivosPorContrato]` - Exporta os 4 cadernos oficiais (.xlsx) para a pasta `MEDIÇÃO`
- `[modLimparMonitoramento.bas:LimparAbaMonitoramentoManterNC]` - Purga atendimentos concluídos preservando linhas com Situação = 'NC'
- `[ImportarTabelaB.bas:ImportarTabelaB]` - Sincroniza tabelas de parâmetros e prazos a partir da pasta de referência `TabelaB-V2.xlsx`

## Invariants
- `[modAuditoriaLog.bas:AuditarAbaAtiva]` A aba ativa não pode ser `LOG_CRITICAS` ou `Painel`.
- `[modAuditoriaLog.bas:ExecutarMotorAuditoria]` A base de dados requer pelo menos 2 linhas (cabeçalho e uma linha de conteúdo).
- `[modAuditoriaLog.bas:ExecutarMotorAuditoria]` Campos obrigatórios (`Contrato`, `Descricao da atividade`, `Item`, `Codigo da solicitacao`) não podem estar em branco.
- `[modAuditoriaLog.bas:ExecutarMotorAuditoria]` A data de fechamento não pode ser anterior à data de abertura (`dtFech >= dtAber`).
- `[modAuditoriaLog.bas:ExecutarMotorAuditoria]` Quantidade solicitada deve ser estritamente numérica e $> 0$; quantidade atendida numérica e $\ge 0$.
- `[modAuditoriaLog.bas:ExecutarMotorAuditoria]` A chave do solicitante deve possuir exatamente 4 caracteres alfanuméricos.
- `[modLimparMonitoramento.bas:LimparAbaMonitoramentoManterNC]` Preservação exclusiva de linhas com Situação = 'NC' na tabela estruturada.

## Side Effects
- `[modAuditoriaLog.bas:ExecutarMotorAuditoria]` Otimiza e restaura `ScreenUpdating`, `EnableEvents` e `Calculation` (xlCalculationManual).
- `[modAuditoriaLog.bas:ExecutarMotorAuditoria]` Escreve o texto descritivo dos erros encontrados na coluna `LOG_AUDITORIA` da aba auditada.
- `[modAuditoriaLog.bas:ExecutarMotorAuditoria]` Cria ou sobrescreve a aba `LOG_CRITICAS` com tabela zebrada, métricas e congelamento de painéis.
- `[modAuditoriaLog.bas:ExecutarMotorAuditoria]` Atualiza status, contagem de erros e coloração na célula `Painel!F12`.
- `[modAuditoriaFontesPQ.bas:ScanearFontesPowerQuery]` Cria ou substitui a aba `Auditoria_Fontes_PQ` com cartões de indicadores e lista de fontes de dados.
- `[modGerarArquivos.bas:GerarArquivosPorContrato]` Cria a pasta de diretório `MEDIÇÃO` e grava arquivos `.xlsx` independentes por contrato.

## Verification
- Run: `<test-command>`
