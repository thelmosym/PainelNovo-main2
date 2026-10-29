<#
.SYNOPSIS
   Importa e executa o modulo oficial modDemonstrativoFDM na planilha principal
   "Painel de controle MemoriaPetrobras V84.xlsm".
#>

param(
    [string]$WorkbookPath = "d:\PainelNovo-main\Painel de controle MemoriaPetrobras V84.xlsm",
    [string]$ModulePath   = "d:\PainelNovo-main\Codes\VBA\modDemonstrativoFDM.bas",
    [switch]$SkipRun
)

Write-Host "=========================================================="
Write-Host "Iniciando importacao e execucao do modDemonstrativoFDM..."
Write-Host "=========================================================="

if (-not (Test-Path $WorkbookPath)) {
    Write-Error "Arquivo de pasta de trabalho nao encontrado: $WorkbookPath"
    exit 1
}

if (-not (Test-Path $ModulePath)) {
    Write-Error "Arquivo de modulo VBA nao encontrado: $ModulePath"
    exit 1
}

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.AutomationSecurity = 1   # msoAutomationSecurityLow

$workbook = $null

try {
    Write-Host "Abrindo pasta de trabalho: $WorkbookPath"
    $workbook = $excel.Workbooks.Open($WorkbookPath)

    $vbProj = $workbook.VBProject
    if ($null -eq $vbProj) {
        Write-Error "Nao foi possivel acessar VBProject. Verifique se 'Confiar no acesso ao modelo de objeto do projeto VBA' esta ativo nas Opcoes do Excel."
        exit 1
    }

    # Remover modulos legados ou versoes anteriores
    $nomesRemover = @("modDemonstrativoFDM", "FDMAnalysis")
    foreach ($comp in $vbProj.VBComponents) {
        if ($nomesRemover -contains $comp.Name) {
            Write-Host "Removendo componente antigo: $($comp.Name)"
            $vbProj.VBComponents.Remove($comp)
        }
    }

    # Importar novo modulo oficial
    Write-Host "Importando modulo oficial: $ModulePath"
    $novoModulo = $vbProj.VBComponents.Import($ModulePath)
    Write-Host "Modulo '$($novoModulo.Name)' importado com sucesso!"

    # Salvar workbook apos importacao
    Write-Host "Salvando alteracoes no arquivo principal..."
    $workbook.Save()
    Write-Host "[OK] Modulo integrado e salvo com sucesso!"

    # Executar geracao do demonstrativo se nao solicitado pular
    if (-not $SkipRun) {
        Write-Host "----------------------------------------------------------"
        Write-Host "Executando analise oficial de FDM/TSPE/IAPFARQ..."
        $macroName = "'" + $workbook.Name + "'!modDemonstrativoFDM.GerarDemonstrativoCalculoFDM_Silencioso"
        $sw = [System.Diagnostics.Stopwatch]::StartNew()
        $excel.Run($macroName)
        $sw.Stop()
        Write-Host "[OK] Analise concluida com sucesso em $($sw.ElapsedMilliseconds) ms!"

        # Salvar novamente com o link gravado em F18
        $workbook.Save()
        Write-Host "[OK] Pasta de trabalho principal atualizada e salva com sucesso!"
    }

} catch {
    Write-Error "Falha durante o processo: $_"
    exit 1
} finally {
    if ($workbook) {
        $workbook.Close($false)
    }
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    Write-Host "Excel COM finalizado com sucesso."
}
