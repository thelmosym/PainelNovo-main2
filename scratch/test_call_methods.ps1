$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.AutomationSecurity = 1

$wb = $excel.Workbooks.Open('D:\PainelNovo-main\Monitoramento\Monitoramento individual_DPBR_v4.xlsm')

$namesToTry = @(
    "AnalisarErrosMonitoramento_Silencioso",
    "modAnaliseMonitoramento.AnalisarErrosMonitoramento_Silencioso",
    "'$($wb.Name)'!modAnaliseMonitoramento.AnalisarErrosMonitoramento_Silencioso",
    "AnalisarErrosMonitoramentoEGerarLog",
    "modAnaliseMonitoramento.AnalisarErrosMonitoramentoEGerarLog",
    "'$($wb.Name)'!modAnaliseMonitoramento.AnalisarErrosMonitoramentoEGerarLog"
)

foreach ($name in $namesToTry) {
    try {
        Write-Host "Trying: $name"
        $excel.Run($name)
        Write-Host "SUCCESS with $name!"
        break
    } catch {
        Write-Host "Failed: $($_.Exception.Message)"
    }
}

$wb.Close($false)
$excel.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null

