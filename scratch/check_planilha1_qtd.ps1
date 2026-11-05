$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.AutomationSecurity = 1

$wb = $excel.Workbooks.Open('D:\PainelNovo-main\Monitoramento\Monitoramento individual_DPBR_v4.xlsm', $null, $true)
$comp = $wb.VBProject.VBComponents.Item('Planilha1')
for ($i = 1; $i -le $comp.CodeModule.CountOfLines; $i++) {
    $line = $comp.CodeModule.Lines($i, 1)
    if ($line -like "*solicitada*" -or $line -like "*atendida*" -or $line -like "*Coluna T*" -or $line -like "*Coluna O*") {
        Write-Host "$($i): $line"
    }
}
$wb.Close($false)
$excel.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null

