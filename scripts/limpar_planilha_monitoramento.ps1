<#
.SYNOPSIS
    limpar_planilha_monitoramento.ps1
    Automação nativa em PowerShell / Excel COM para limpar e organizar a aba 'Monitoramento',
    removendo todas as linhas cuja Situação seja 'AT' ou 'CO', mantendo as demais linhas preenchidas,
    classificando-as em ordem crescente pela coluna 'D. abertura', preservando todas as formatações
    e validações de dados e mantendo o mínimo de 700 linhas na tabela estruturada.

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File scripts/limpar_planilha_monitoramento.ps1 -Caminho "Monitoramento/Monitoramento individual_A4UU_v7.xlsm"
#>

param (
    [string]$Caminho = "",
    [switch]$ApenasImportarMacro
)

$ErrorActionPreference = "Stop"

function Limpar-ArquivoExcel {
    param (
        [string]$Arquivo,
        [bool]$SomenteImportar = $false
    )

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

        # 1. Atualizar o módulo VBA modLimparMonitoramento no workbook
        $basPath = Join-Path (Split-Path -Parent $PSScriptRoot) "Codes\VBA\modLimparMonitoramento.bas"
        if (Test-Path $basPath) {
            $vbp = $wb.VBProject
            $existing = $null
            for ($i = 1; $i -le $vbp.VBComponents.Count; $i++) {
                if ($vbp.VBComponents.Item($i).Name -eq "modLimparMonitoramento") {
                    $existing = $vbp.VBComponents.Item($i)
                    break
                }
            }
            if ($null -ne $existing) {
                $vbp.VBComponents.Remove($existing)
            }
            $vbp.VBComponents.Import($basPath) | Out-Null
            Write-Host "  - Módulo VBA 'modLimparMonitoramento' atualizado com sucesso."
        }

        if (-not $SomenteImportar) {
            $ws = $wb.Worksheets.Item("Monitoramento")
            $tbl = $ws.ListObjects.Item("Monitoramento")
            Write-Host "  - Antes: '$($tbl.Name)' | Linhas: $($tbl.ListRows.Count) | Intervalo: $($tbl.Range.Address())"

            $xl.Run("modLimparMonitoramento.LimparAbaMonitoramentoManterNC", $true)

            Write-Host "  - Depois: '$($tbl.Name)' | Linhas: $($tbl.ListRows.Count) | Intervalo: $($tbl.Range.Address())" -ForegroundColor Green
        }

        $wb.Save()
        $wb.Close($false)
        Write-Host "  [OK] Concluído e salvo: $NomeArquivo" -ForegroundColor Green

    } catch {
        Write-Host "  [ERRO] Falha ao processar $NomeArquivo : $($_.Exception.Message)" -ForegroundColor Red
        if ($null -ne $wb) {
            $wb.Close($false)
        }
        throw
    } finally {
        if ($null -ne $xl) {
            $xl.Quit()
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
        }
    }
}

if ([string]::IsNullOrWhiteSpace($Caminho)) {
    $dirMonit = Join-Path (Split-Path -Parent $PSScriptRoot) "Monitoramento"
    $arquivos = Get-ChildItem -Path $dirMonit -Filter "Monitoramento individual_*.xlsm" | Where-Object { $_.Name -notlike "~$*" -and $_.Name -notlike "*backup*" }
    foreach ($arq in $arquivos) {
        Limpar-ArquivoExcel -Arquivo $arq.FullName -SomenteImportar:$ApenasImportarMacro.IsPresent
    }
} elseif (Test-Path $Caminho -PathType Leaf) {
    Limpar-ArquivoExcel -Arquivo $Caminho -SomenteImportar:$ApenasImportarMacro.IsPresent
} elseif (Test-Path $Caminho -PathType Container) {
    $arquivos = Get-ChildItem -Path $Caminho -Filter "Monitoramento individual_*.xlsm" | Where-Object { $_.Name -notlike "~$*" -and $_.Name -notlike "*backup*" }
    foreach ($arq in $arquivos) {
        Limpar-ArquivoExcel -Arquivo $arq.FullName -SomenteImportar:$ApenasImportarMacro.IsPresent
    }
} else {
    Write-Host "[ERRO] Caminho inválido: $Caminho" -ForegroundColor Red
}
