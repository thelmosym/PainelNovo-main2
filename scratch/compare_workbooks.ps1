$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.AutomationSecurity = 1

$wbA = $excel.Workbooks.Open('D:\PainelNovo-main\Monitoramento\Monitoramento individual_A4UU_v4.xlsm', $null, $true)
$wbD = $excel.Workbooks.Open('D:\PainelNovo-main\Monitoramento\Monitoramento individual_DPBR_v4.xlsm', $null, $true)

Write-Host "=== REFERENCES A4UU ==="
foreach ($r in $wbA.VBProject.References) {
    Write-Host "  Name: $($r.Name) | Path: $($r.FullPath) | Broken: $($r.IsBroken)"
}

Write-Host "=== REFERENCES DPBR ==="
foreach ($r in $wbD.VBProject.References) {
    Write-Host "  Name: $($r.Name) | Path: $($r.FullPath) | Broken: $($r.IsBroken)"
}

Write-Host "=== COMPONENTS IN A4UU ==="
foreach ($c in $wbA.VBProject.VBComponents) {
    Write-Host "  $($c.Name): Type=$($c.Type), Lines=$($c.CodeModule.CountOfLines)"
}

Write-Host "=== COMPONENTS IN DPBR ==="
foreach ($c in $wbD.VBProject.VBComponents) {
    Write-Host "  $($c.Name): Type=$($c.Type), Lines=$($c.CodeModule.CountOfLines)"
}

$wbA.Close($false)
$wbD.Close($false)
$excel.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null

