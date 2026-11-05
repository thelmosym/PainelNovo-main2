$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.AutomationSecurity = 1

$wbA = $excel.Workbooks.Open('D:\PainelNovo-main\Monitoramento\Monitoramento individual_A4UU_v4.xlsm', $null, $true)
$wbD = $excel.Workbooks.Open('D:\PainelNovo-main\Monitoramento\Monitoramento individual_DPBR_v4.xlsm', $null, $true)

$wbA.VBProject.VBComponents.Item('modAnaliseMonitoramento').Export('D:\PainelNovo-main\scratch\modAnalise_A4UU.bas')
$wbD.VBProject.VBComponents.Item('modAnaliseMonitoramento').Export('D:\PainelNovo-main\scratch\modAnalise_DPBR.bas')

$wbA.Close($false)
$wbD.Close($false)
$excel.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
Write-Host "Exported both modules!"

