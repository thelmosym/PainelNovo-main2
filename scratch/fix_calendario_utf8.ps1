$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false

$filePath = "d:\PainelNovo-main\PainelAgosto\Painel Agosto.xlsm"
$mFile = "d:\PainelNovo-main\Codes\CodigosM\Calendario.m"

try {
    Write-Host "Reading $mFile as UTF8..."
    $newFormula = Get-Content -Path $mFile -Raw -Encoding UTF8
    
    Write-Host "Formula to apply:"
    Write-Host $newFormula
    
    Write-Host "Opening workbook $filePath..."
    $wb = $xl.Workbooks.Open($filePath)
    
    $q = $wb.Queries.Item('Calendario')
    $q.Formula = $newFormula
    
    Write-Host "Formula updated."
    $wb.Save()
    Write-Host "Workbook saved."
    $wb.Close($false)
} catch {
    Write-Host "ERROR: $($_.Exception.Message)"
} finally {
    $xl.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
}
