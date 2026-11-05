$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.AutomationSecurity = 1

$wb = $excel.Workbooks.Open('D:\PainelNovo-main\Monitoramento\Monitoramento individual_DPBR_v4.xlsm')
Write-Host "Workbook opened: $($wb.Name)"

$macro1 = "'$($wb.Name)'!AnalisarErrosMonitoramento_Silencioso"
Write-Host "Calling: $macro1"

$sw = [System.Diagnostics.Stopwatch]::StartNew()
$excel.Run($macro1)
$sw.Stop()
Write-Host "SUCCESS! Elapsed: $($sw.ElapsedMilliseconds) ms"

$wb.Save()
$wb.Close($true)
$excel.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null

