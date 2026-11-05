$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.AutomationSecurity = 1

$wb = $excel.Workbooks.Open('D:\PainelNovo-main\Monitoramento\Monitoramento individual_A4UU_v4.xlsm', $null, $true) # ReadOnly
$comp = $wb.VBProject.VBComponents.Item('modAnaliseMonitoramento')
for ($i = 40; $i -le 65; $i++) {
    Write-Host "$($i): $($comp.CodeModule.Lines($i, 1))"
}
$wb.Close($false)
$excel.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null

