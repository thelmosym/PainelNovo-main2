$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.ScreenUpdating = $false
$excel.AutomationSecurity = 1

$basPath = "D:\PainelNovo-main\Codes\VBA\modAnaliseMonitoramento.bas"
$files = Get-ChildItem "D:\PainelNovo-main\Monitoramento\*.xlsm"

foreach ($f in $files) {
    try {
        Write-Host "Updating workbook: $($f.Name)..."
        $wb = $excel.Workbooks.Open($f.FullName)
        $vbProj = $wb.VBProject

        $existing = $null
        foreach ($comp in $vbProj.VBComponents) {
            if ($comp.Name -eq "modAnaliseMonitoramento") {
                $existing = $comp
                break
            }
        }

        if ($existing -ne $null) {
            $vbProj.VBComponents.Remove($existing)
        }

        $newComp = $vbProj.VBComponents.Import($basPath)
        $wb.Save()

        # Run silent audit
        $excel.Run("AnalisarErrosMonitoramento_Silencioso")
        $wb.Save()
        $wb.Close($true)
        Write-Host "Successfully updated and audited $($f.Name)!"
    }
    catch {
        Write-Host "Error updating $($f.Name): $($_.Exception.Message)"
        if ($wb) { $wb.Close($false) }
    }
}

$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()
Write-Host "All monitoring workbooks updated successfully!"

