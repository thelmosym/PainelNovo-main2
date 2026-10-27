<#
.SYNOPSIS
    limpar_planilha_monitoramento.ps1
    Automação nativa em PowerShell / Excel COM para limpar e organizar a aba 'Monitoramento',
    removendo todas as linhas cuja Situação NÃO seja 'NC' (mantendo exclusivamente os registros NC),
    classificando as linhas mantidas em ordem crescente pela coluna 'Contrato', preservando
    validações de dados e formatações da tabela.

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File scripts/limpar_planilha_monitoramento.ps1 -Caminho "Monitoramento/Monitoramento individual_A4UU_v4.xlsm"
#>

param (
    [string]$Caminho = ""
)

$ErrorActionPreference = "Stop"

function Limpar-ArquivoExcel {
    param ([string]$Arquivo)

    $ArquivoAbs = (Resolve-Path $Arquivo).Path
    $NomeArquivo = Split-Path -Leaf $ArquivoAbs
    Write-Host "`n[PROCESSANDO] $NomeArquivo..." -ForegroundColor Cyan

    $xl = $null
    $wb = $null

    try {
        $xl = New-Object -ComObject Excel.Application
        $xl.Visible = $false
        $xl.DisplayAlerts = $false
        $xl.ScreenUpdating = $false
        $xl.EnableEvents = $false

        $wb = $xl.Workbooks.Open($ArquivoAbs)

        # 1. Atualizar o módulo VBA modLimparMonitoramento no workbook se disponível
        $basPath = Join-Path (Split-Path -Parent $PSScriptRoot) "Codes\VBA\modLimparMonitoramento.bas"
        if (Test-Path $basPath) {
            try {
                $vbp = $wb.VBProject
                $existing = $null
                for ($i = 1; $i -le $vbp.VBComponents.Count; $i++) {
                    if ($vbp.VBComponents.Item($i).Name -eq "modLimparMonitoramento") {
                        $existing = $vbp.VBComponents.Item($i)
                        break
                    }
                }
                if ($existing -ne $null) {
                    $vbp.VBComponents.Remove($existing)
                }
                $vbp.VBComponents.Import($basPath) | Out-Null
                Write-Host "  - Módulo VBA 'modLimparMonitoramento' atualizado no projeto."
            } catch {
                Write-Host "  - [AVISO] Não foi possível atualizar VBProject via automação: $($_.Exception.Message)" -ForegroundColor Yellow
            }
        }

        # 2. Localizar aba Monitoramento
        $ws = $null
        foreach ($s in $wb.Sheets) {
            if ($s.Name -eq "Monitoramento") {
                $ws = $s
                break
            }
        }

        if ($ws -eq $null) {
            Write-Host "  [AVISO] Aba 'Monitoramento' não encontrada em $NomeArquivo." -ForegroundColor Yellow
            $wb.Close($false)
            return
        }

        # 3. Localizar Tabela Estruturada (ListObject)
        $tbl = $null
        if ($ws.ListObjects.Count -gt 0) {
            try {
                $tbl = $ws.ListObjects.Item("Monitoramento")
            } catch {
                $tbl = $ws.ListObjects.Item(1)
            }
        }

        if ($tbl -eq $null) {
            Write-Host "  [AVISO] Nenhuma tabela estruturada encontrada na aba 'Monitoramento'." -ForegroundColor Yellow
            $wb.Close($false)
            return
        }

        $linhasIniciais = $tbl.ListRows.Count
        Write-Host "  - Tabela: '$($tbl.Name)' | Linhas iniciais: $linhasIniciais | Intervalo: $($tbl.Range.Address())"

        # Tentar executar a macro nativa de limpeza parametrizada (modo silencioso)
        $executouMacro = $false
        try {
            $xl.Run("modLimparMonitoramento.LimparAbaMonitoramentoManterNC", $true)
            $executouMacro = $true
            Write-Host "  - Rotina VBA 'LimparAbaMonitoramentoManterNC' executada com sucesso!" -ForegroundColor Green
        } catch {
            Write-Host "  - [AVISO] Execução direta da macro falhou ou módulo indisponível, executando lógica nativa COM: $($_.Exception.Message)" -ForegroundColor Yellow
        }

        # Se a macro não tiver sido executada, aplica a lógica diretamente via COM
        if (-not $executouMacro) {
            # Desativar filtros
            if ($ws.FilterMode) { $ws.ShowAllData() }
            if ($tbl.AutoFilter -ne $null -and $tbl.AutoFilter.FilterMode) { $tbl.AutoFilter.ShowAllData() }

            # Localizar colunas
            $colSituacao = 3
            $colContrato = 4
            for ($c = 1; $c -le $tbl.ListColumns.Count; $c++) {
                $nomeCol = $tbl.ListColumns.Item($c).Name
                if ($nomeCol -like "*Situa*") { $colSituacao = $c }
                if ($nomeCol -like "*Contrato*") { $colContrato = $c }
            }

            # Deletar linhas que não sejam NC (de baixo para cima)
            $removidas = 0
            $mantidas = 0
            for ($i = $tbl.ListRows.Count; $i -ge 1; $i--) {
                $val = [string]$tbl.ListRows.Item($i).Range.Cells.Item(1, $colSituacao).Value2
                if ($val.Trim().ToUpper() -eq "NC") {
                    $mantidas++
                } else {
                    $tbl.ListRows.Item($i).Delete()
                    $removidas++
                }
            }

            # Classificar por Contrato
            if ($tbl.ListRows.Count -gt 1 -and $colContrato -gt 0) {
                try {
                    $tbl.Sort.SortFields.Clear()
                    $tbl.Sort.SortFields.Add($tbl.ListColumns.Item($colContrato).Range, 0, 1, 0)
                    $tbl.Sort.Header = 1
                    $tbl.Sort.MatchCase = $false
                    $tbl.Sort.Apply()
                } catch {}
            }

            # Limpar resíduos abaixo da tabela
            $fimTabela = $tbl.Range.Row + $tbl.Range.Rows.Count - 1
            $lastRow = $ws.Cells.SpecialCells(11).Row
            if ($lastRow -gt $fimTabela) {
                $ws.Rows.Item(($fimTabela + 1).ToString() + ":" + $lastRow.ToString()).Delete() | Out-Null
            }

            # Limpar células à direita
            $colFimTabela = $tbl.Range.Column + $tbl.Range.Columns.Count - 1
            $lastCol = $ws.Cells.SpecialCells(11).Column
            if ($lastCol -gt $colFimTabela) {
                $ws.Range($ws.Cells.Item(1, $colFimTabela + 1), $ws.Cells.Item($ws.Rows.Count, $lastCol)).Clear() | Out-Null
            }
        }

        $linhasFinais = $tbl.ListRows.Count
        Write-Host "  - Linhas mantidas (Situação = NC): $linhasFinais | Novo intervalo: $($tbl.Range.Address())"

        $wb.Save()
        $wb.Close($false)
        Write-Host "  [OK] Sucesso! Planilha processada e salva com linhas NC ordenadas por Contrato." -ForegroundColor Green

    } catch {
        Write-Host "  [ERRO] Falha ao processar $NomeArquivo : $($_.Exception.Message)" -ForegroundColor Red
        if ($wb -ne $null) {
            $wb.Close($false)
        }
    } finally {
        if ($xl -ne $null) {
            $xl.Quit()
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
        }
    }
}

# Início da Execução
if ([string]::IsNullOrWhiteSpace($Caminho)) {
    $dirMonit = Join-Path (Split-Path -Parent $PSScriptRoot) "Monitoramento"
    $arquivos = Get-ChildItem -Path $dirMonit -Filter "Monitoramento individual_*.xlsm" | Where-Object { $_.Name -notlike "~$*" -and $_.Name -notlike "*backup*" }
    foreach ($arq in $arquivos) {
        Limpar-ArquivoExcel -Arquivo $arq.FullName
    }
} elseif (Test-Path $Caminho -PathType Leaf) {
    Limpar-ArquivoExcel -Arquivo $Caminho
} elseif (Test-Path $Caminho -PathType Container) {
    $arquivos = Get-ChildItem -Path $Caminho -Filter "Monitoramento individual_*.xlsm" | Where-Object { $_.Name -notlike "~$*" -and $_.Name -notlike "*backup*" }
    foreach ($arq in $arquivos) {
        Limpar-ArquivoExcel -Arquivo $arq.FullName
    }
} else {
    Write-Host "[ERRO] Caminho inválido: $Caminho" -ForegroundColor Red
}
