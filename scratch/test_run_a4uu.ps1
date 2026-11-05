$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.AutomationSecurity = 1

$wb = $excel.Workbooks.Open('D:\PainelNovo-main\Monitoramento\Monitoramento individual_A4UU_v4.xlsm', $null, $true)

try {
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $excel.Run('AnalisarErrosMonitoramento_Silencioso')
    $sw.Stop()
    Write-Host "A4UU SUCCEEDED in $($sw.ElapsedMilliseconds) ms!"
} catch {
    Write-Host "A4UU Failed: $($_.Exception.Message)"
}

$wb.Close($false)
$excel.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null

