$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.AutomationSecurity = 1

$wb = $excel.Workbooks.Open('D:\PainelNovo-main\Monitoramento\Monitoramento individual_A4UU_v4.xlsm', $null, $true) # ReadOnly
Write-Host "VBComponents in $($wb.Name):"
foreach ($c in $wb.VBProject.VBComponents) {
    Write-Host "  Name: $($c.Name) | Type: $($c.Type) | Lines: $($c.CodeModule.CountOfLines)"
}
$wb.Close($false)
$excel.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
