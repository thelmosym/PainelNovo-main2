$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.ScreenUpdating = $false
$excel.AutomationSecurity = 1

$files = Get-ChildItem "D:\PainelNovo-main\Monitoramento\*.xlsm" | Where-Object { -not ($_.Name.StartsWith("~")) }

foreach ($f in $files) {
    try {
        $wb = $excel.Workbooks.Open($f.FullName)
        $sw = [System.Diagnostics.Stopwatch]::StartNew()
        $excel.Run("AnalisarErrosMonitoramento_Silencioso")
        $sw.Stop()
        Write-Host "Workbook $($f.Name): AUDIT OK in $($sw.ElapsedMilliseconds) ms!"
        $wb.Close($false)
    }
    catch {
        Write-Host "Workbook $($f.Name): ERROR - $($_.Exception.Message)"
        if ($wb) { $wb.Close($false) }
    }
}

$excel.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()
Write-Host "ALL WORKBOOKS VERIFIED!"

