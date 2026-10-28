$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false

$filePath = "d:\PainelNovo-main\PainelAgosto\Painel Agosto.xlsm"

try {
    Write-Host "Opening workbook $filePath..."
    $wb = $xl.Workbooks.Open($filePath)
    
    $wsVal = $null
    for ($i = 1; $i -le $wb.Sheets.Count; $i++) {
        $s = $wb.Sheets.Item($i)
        if ($s.Name -like "*Valida*") {
            $wsVal = $s
            break
        }
    }
    
    if ($wsVal -ne $null) {
        Write-Host "Found sheet: $($wsVal.Name)"
        Write-Host "D5 formula:" $wsVal.Range("D5").Formula
        Write-Host "D5 text:" $wsVal.Range("D5").Text
        Write-Host "D9 (localTabelaB) formula:" $wsVal.Range("D9").Formula
        Write-Host "D9 (localTabelaB) text:" $wsVal.Range("D9").Text
        $tabelaBPath = $wsVal.Range("D9").Text
        $exists = Test-Path $tabelaBPath
        Write-Host "Does TabelaB exist at $tabelaBPath ? $exists"
    } else {
        Write-Host "Sheet Validações not found!"
    }
    
    $wb.Close($false)
} catch {
    Write-Host "ERROR: $($_.Exception.Message)"
} finally {
    $xl.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
}
