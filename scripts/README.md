# Scripts Utilitários e Automação de Engenharia

## Overview
Scripts utilitários em PowerShell, Python e Shell para suporte ao desenvolvimento, automação de tarefas operacionais no Excel via API COM, injeção/atualização de módulos VBA, sanitização de bases de monitoramento e auditoria de integridade de dados e dashboards.

## Interface
- `[build_professional_dashboard.ps1]` - Constrói abas de dashboard executivo, cartões de KPI e gráficos no Excel via automação COM
- `[limpar_planilha_monitoramento.ps1]` / `[limpar_planilha_monitoramento.py]` - Limpa dados concluídos das planilhas de monitoramento individual preservando linhas 'NC'
- `[importar_modDemonstrativoFDM.ps1]` - Remove e reimporta o módulo `modDemonstrativoFDM.bas` no projeto VBA da pasta de trabalho e dispara execução silenciosa
- `[AddFDMAnalysisModule.ps1]` / `[AddFDMAnalysisModule_Force.ps1]` - Realiza injeção forçada de módulos VBA de apuração de FDM
- `[inspect_sheets.ps1]` / `[inspect_painel.py]` - Inspeciona estrutura de abas, tabelas e conexões de pastas de trabalho
- `[inspect_criticas.py]` / `[inspect_dashboard.py]` - Audita inconsistências e checa valores calculados de cards e tabelas
- `[verify_dashboard_complete.ps1]` / `[verify_dashboard_output.py]` - Validação automatizada de consistência dos números do painel
- `[format_cards_ptbr.ps1]` - Formata máscaras monetárias e numéricas segundo o padrão pt-BR

## Invariants
- Scripts que interagem com o Excel exigem a presença do runtime local do Microsoft Office e habilitação de acesso confiável ao modelo de objeto do VBA (`AccessVBOM`).
- Os scripts de limpeza de monitoramento garantem a integridade das listas estruturadas (`ListObjects`), cabeçalhos e regras de validação nativas.
- Fechamento seguro de instâncias COM do Excel garantido em blocos `finally` via `[System.Runtime.InteropServices.Marshal]::ReleaseComObject`.

## Side Effects
- Criação e encerramento de processos `excel.exe` em segundo plano (`Visible = false`, `DisplayAlerts = false`).
- Remoção e importação de componentes no projeto VBA (`VBComponents.Remove` e `VBComponents.Import`).
- Modificação física de células, formatações e dados em arquivos `.xlsm` gravados em disco.

## Verification
- Run: `<test-command>`
