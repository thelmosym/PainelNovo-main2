$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false

$filePath = "d:\PainelNovo-main\PainelAgosto\Painel Agosto.xlsm"

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
    
    Write-Host "Refreshing all data connections (this may take 1-2 minutes)..."
    $wb.RefreshAll()
    
    # Wait for any async queries if needed
    $xl.CalculateUntilAsyncQueriesDone()
    
    Write-Host "Calculation completed. Calculating full workbook..."
    $xl.CalculateFull()
    
    Write-Host "Saving workbook..."
    $wb.Save()
    
    # Read values from TabelaFDMPorContrato
    $wsFDM = $null
    for ($i = 1; $i -le $wb.Sheets.Count; $i++) {
        if ($wb.Sheets.Item($i).Name -eq 'TabelaFDMPorContrato') {
            $wsFDM = $wb.Sheets.Item($i)
            break
        }
    }
    
    if ($wsFDM -ne $null) {
        Write-Host "--- TabelaFDMPorContrato rows ---"
        for ($r = 1; $r -le 4; $r++) {
            $line = ""
            for ($c = 1; $c -le 12; $c++) {
                $line += "$($wsFDM.Cells.Item($r, $c).Text) | "
            }
            Write-Host $line
        }
    }
    
    # Read first 5 rows of Painel T2M
    $wsT2M = $null
    for ($i = 1; $i -le $wb.Sheets.Count; $i++) {
        if ($wb.Sheets.Item($i).Name -eq 'Painel T2M') {
            $wsT2M = $wb.Sheets.Item($i)
            break
        }
    }
    
    if ($wsT2M -ne $null) {
        Write-Host "--- Painel T2M SLA sample ---"
        for ($r = 2; $r -le 6; $r++) {
            $os = $wsT2M.Cells.Item($r, 12).Text
            $status = $wsT2M.Cells.Item($r, 46).Text
            $log = $wsT2M.Cells.Item($r, 47).Text
            Write-Host "Row $r - OS: $os | Status Prazo: $status | LOG SLA: $log"
        }
    }
    
    $wb.Close($false)
    Write-Host "ALL DONE!"
} catch {
    Write-Host "ERROR: $($_.Exception.Message)"
} finally {
    $xl.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
}
