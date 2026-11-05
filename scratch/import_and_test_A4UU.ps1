$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.ScreenUpdating = $false
$excel.AutomationSecurity = 1

$wbPath = "D:\PainelNovo-main\Monitoramento\Monitoramento individual_A4UU_v4.xlsm"
$basPath = "D:\PainelNovo-main\Codes\VBA\modAnaliseMonitoramento.bas"

try {
    Write-Host "Opening workbook: $wbPath"
    $wb = $excel.Workbooks.Open($wbPath)
    $vbProj = $wb.VBProject

    # Check if modAnaliseMonitoramento exists
    $existing = $null
    foreach ($comp in $vbProj.VBComponents) {
        if ($comp.Name -eq "modAnaliseMonitoramento") {
            $existing = $comp
            break
        }
    }

    if ($existing -ne $null) {
        Write-Host "Removing existing component modAnaliseMonitoramento..."
        $vbProj.VBComponents.Remove($existing)
    }

    Write-Host "Importing updated component from: $basPath"
    $newComp = $vbProj.VBComponents.Import($basPath)
    Write-Host "Imported component successfully: $($newComp.Name)"

    Write-Host "Saving workbook with updated module..."
    $wb.Save()

    Write-Host "Running AnalisarErrosMonitoramento_Silencioso..."
    $t0 = [System.Diagnostics.Stopwatch]::StartNew()
    $excel.Run("AnalisarErrosMonitoramento_Silencioso")
    $t0.Stop()
    Write-Host "Macro executed successfully in $($t0.ElapsedMilliseconds) ms!"

    Write-Host "Saving workbook after audit..."
    $wb.Save()
    $wb.Close($true)
    Write-Host "Done!"
}
catch {
    Write-Host "ERROR: $($_.Exception.Message)"
}
finally {
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}

