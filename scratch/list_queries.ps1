$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
try {
    $wb = $xl.Workbooks.Open('D:\PainelNovo-main\Painel de controle MemoriaPetrobras V82.xlsm')
    Write-Host "Total queries: $($wb.Queries.Count)"
    foreach ($q in $wb.Queries) {
        if ($q.Name -in @('Painel T2M', 'DADOS', 'Localidade Com Endereço', 'validUFGalpao', 'Localidade', 'Centro')) {
            Write-Host "--- Query: $($q.Name) ---"
            Write-Host $q.Formula.Substring(0, [Math]::Min(300, $q.Formula.Length))
        }
    }
    $wb.Close($false)
} catch {
    Write-Host "ERROR: $($_.Exception.Message)"
} finally {
    $xl.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
}
