$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.ScreenUpdating = $false
# 1 = msoAutomationSecurityLow
$excel.AutomationSecurity = 1

$wbPath = "D:\PainelNovo-main\Painel de controle MemoriaPetrobras V.8.5-Outubro.xlsm"
$basPath = "D:\PainelNovo-main\Codes\VBA\modDemonstrativoFDM.bas"

try {
    Write-Host "Opening workbook: $wbPath"
    $wb = $excel.Workbooks.Open($wbPath)
    $vbProj = $wb.VBProject

    # Check if modDemonstrativoFDM exists
    $existing = $null
    foreach ($comp in $vbProj.VBComponents) {
        if ($comp.Name -eq "modDemonstrativoFDM") {
            $existing = $comp
            break
        }
    }

    if ($existing -ne $null) {
        Write-Host "Removing existing component modDemonstrativoFDM..."
        $vbProj.VBComponents.Remove($existing)
    }

    Write-Host "Importing updated component from: $basPath"
    $newComp = $vbProj.VBComponents.Import($basPath)
    Write-Host "Imported component successfully: $($newComp.Name)"

    Write-Host "Saving workbook with updated module..."
    $wb.Save()

    Write-Host "Executing GerarDemonstrativoCalculoFDM_Silencioso..."
    $t0 = [System.Diagnostics.Stopwatch]::StartNew()
    $excel.Run("GerarDemonstrativoCalculoFDM_Silencioso")
    $t0.Stop()
    Write-Host "SUCCESS! Macro executed in $($t0.ElapsedMilliseconds) ms!"

    Write-Host "Saving workbook..."
    $wb.Save()
    $wb.Close($true)
    Write-Host "Finished successfully!"
} catch {
    Write-Host "ERROR: $($_.Exception.Message)"
} finally {
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}

