# Módulos VBA (Visual Basic for Applications)

## Overview
Camada de código procedural e orientado a eventos em VBA responsável pela interface do operador no Microsoft Excel, automação de comandos do painel de controle, auditoria de regras contratuais, conciliação e geração de demonstrativos de FDM e exportação dos cadernos oficiais de faturamento.

## Interface
- `[modAuditoriaFontesPQ.bas:ScanearFontesPowerQuery] Public Sub ScanearFontesPowerQuery()` - Varre todas as conexões e consultas Power Query gerando relatório de dependências
- `[modLimparMonitoramento.bas:LimparMonitoramento] Public Sub LimparMonitoramento()` - Ponto de entrada padrão para higienização da planilha operacional
- `[modLimparMonitoramento.bas:LimparAbaMonitoramentoManterNC] Public Sub LimparAbaMonitoramentoManterNC(Optional ByVal bSilencioso As Boolean = False)` - Limpa atendimentos concluídos preservando linhas com Situação = 'NC'
- `[modAnaliseMonitoramento.bas:AuditarMonitoramentoIndividual] Public Sub AuditarMonitoramentoIndividual()` - Auditoria detalhada de preenchimento e integridade de datas, contratos e localidades
- `[modDemonstrativoFDM.bas:GerarDemonstrativoCalculoFDM] Public Sub GerarDemonstrativoCalculoFDM()` - Gera o demonstrativo detalhado de cálculo dos fatores TSPE, TSPR e IAPFARQ com memória de cálculo
- `[ImportarTabelaB.bas:ImportarTabelaB] Public Sub ImportarTabelaB()` - Importa e atualiza as tabelas de referência a partir de `TabelaB-V2.xlsx`

### Subdiretórios e Contratos Filhos
- [Backup_PreRefatoracao](./Backup_PreRefatoracao/README.md) - Módulos legados mantidos para histórico e rastreabilidade da arquitetura anterior
- [FuncoesAntigas](./FuncoesAntigas/README.md) - Biblioteca de rotinas de cálculo e manipulação de medição operacional legada
- [Funções Novas](./Funções%20Novas/README.md) - Versão de produção refatorada dos módulos de auditoria, demonstrativo FDM e exportação
- [vba_extracted_A4UU](./vba_extracted_A4UU/README.md) - Componentes e classes de validação extraídos das pastas de trabalho de monitoramento individual

## Invariants
- `[modAuditoriaFontesPQ.bas:Option Explicit]` e `[modLimparMonitoramento.bas:Option Explicit]` Declaração estrita e obrigatória de todas as variáveis.
- `[modAuditoriaFontesPQ.bas:ScanearFontesPowerQuery]` Restauração compulsória dos estados de `Calculation`, `DisplayAlerts`, `EnableEvents` e `ScreenUpdating` no bloco de saída.
- `[modLimparMonitoramento.bas:LimparAbaMonitoramentoManterNC]` Preservação exclusiva de linhas com Situação = "NC" na tabela estruturada.
- `[modAnaliseMonitoramento.bas]` Consistência obrigatória entre Contrato, UF e Localidade validada contra as matrizes da Tabela B.

## Side Effects
- `[modAuditoriaFontesPQ.bas:GerarRelatorioAbaAuditoria]` Cria ou limpa a planilha `Auditoria_Fontes_PQ` via `Cells.Clear`, recriando cabeçalhos e cartões de métricas.
- `[modLimparMonitoramento.bas:LimparAbaMonitoramentoManterNC]` Remove e recria regras de validação de dados (`Validation.Add/Delete`) nas colunas da tabela estruturada e exclui linhas faturadas.
- `[modDemonstrativoFDM.bas:GerarDemonstrativoCalculoFDM]` Cria e formata a aba `DEMONSTRATIVO_FDM` na pasta de trabalho ativa.

## Verification
- Run: `<test-command>`
