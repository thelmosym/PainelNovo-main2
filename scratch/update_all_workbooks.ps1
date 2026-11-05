$basPath = "D:\PainelNovo-main\Codes\VBA\modAnaliseMonitoramento.bas"
$files = Get-ChildItem "D:\PainelNovo-main\Monitoramento\*.xlsm" | Where-Object { -not ($_.Name.StartsWith("~")) }

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.ScreenUpdating = $false
$excel.AutomationSecurity = 1

foreach ($f in $files) {
    Write-Host "=== Updating: $($f.Name) ==="
    try {
        # 1. Replace module
        $wb = $excel.Workbooks.Open($f.FullName)
        $vbProj = $wb.VBProject
        $comp = $null
        foreach ($c in $vbProj.VBComponents) {
            if ($c.Name -eq "modAnaliseMonitoramento") {
                $comp = $c
                break
            }
        }
        if ($comp -ne $null) {
            $vbProj.VBComponents.Remove($comp)
        }
        $vbProj.VBComponents.Import($basPath) | Out-Null
        $wb.Save()
        $wb.Close($true)
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) | Out-Null

        # 2. Re-open and run audit
        $wb = $excel.Workbooks.Open($f.FullName)
        $sw = [System.Diagnostics.Stopwatch]::StartNew()
        $excel.Run("AnalisarErrosMonitoramento_Silencioso")
        $sw.Stop()
        Write-Host "  Audit ran in $($sw.ElapsedMilliseconds) ms!"
        $wb.Save()
        $wb.Close($true)
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) | Out-Null
        Write-Host "  Successfully updated and saved $($f.Name)!"
    }
    catch {
        Write-Host "  ERROR with $($f.Name): $($_.Exception.Message)"
        if ($wb) { $wb.Close($false) }
    }
}

$excel.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()
Write-Host "ALL WORKBOOKS PROCESSED SUCCESSFULLY!"

