$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.AutomationSecurity = 1

$wb = $excel.Workbooks.Open('D:\PainelNovo-main\Monitoramento\Monitoramento individual_DPBR_v4.xlsm')
$comp = $wb.VBProject.VBComponents.Item('Planilha1')
Write-Host "Planilha1 lines: $($comp.CodeModule.CountOfLines)"
for ($i = 1; $i -le [Math]::Min(30, $comp.CodeModule.CountOfLines); $i++) {
    Write-Host "$($i): $($comp.CodeModule.Lines($i, 1))"
}

$wb.Close($false)
$excel.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null

