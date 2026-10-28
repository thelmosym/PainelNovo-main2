$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false

$filePath = "D:\PainelNovo-main\Painel de controle MemoriaPetrobras V82.xlsm"

try {
    Write-Host "Opening workbook $filePath..."
    $wb = $xl.Workbooks.Open($filePath)
    
    Write-Host "Setting BackgroundQuery = false on all connections..."
    foreach ($conn in $wb.Connections) {
        try {
            if ($conn.OLEDBConnection -ne $null) {
                $conn.OLEDBConnection.BackgroundQuery = $false
            }
        } catch {}
    }
    
    Write-Host "Refreshing all data connections..."
    $wb.RefreshAll()
    
    Write-Host "Waiting for async queries..."
    $xl.CalculateUntilAsyncQueriesDone()
    
    Write-Host "Calculating full workbook..."
    $xl.CalculateFull()
    
    Write-Host "Saving workbook..."
    $wb.Save()
    Write-Host "Workbook saved successfully."
    $wb.Close($false)
    Write-Host "ALL DONE!"
} catch {
    Write-Host "ERROR: $($_.Exception.Message)"
} finally {
    $xl.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
}
