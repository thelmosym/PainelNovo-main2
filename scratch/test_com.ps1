$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
try {
    $wb = $xl.Workbooks.Open('d:\PainelNovo-main\PainelAgosto\Painel Agosto.xlsm')
    $q = $wb.Queries.Item('Calendario')
    Write-Host "Found query Calendario:"
    Write-Host $q.Formula
    $wb.Close($false)
} catch {
    Write-Host "ERROR: $($_.Exception.Message)"
} finally {
    $xl.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
}
