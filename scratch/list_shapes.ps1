$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false

try {
    $wb = $xl.Workbooks.Open('D:\PainelNovo-main\Painel de controle MemoriaPetrobras V82.xlsm')
    foreach ($ws in $wb.Worksheets) {
        if ($ws.Shapes.Count -gt 0) {
            Write-Host "Sheet: $($ws.Name) has $($ws.Shapes.Count) shapes"
            foreach ($s in $ws.Shapes) {
                if ($s.OnAction) {
                    Write-Host "  Shape '$($s.Name)' text: '$($s.TextFrame2.TextRange.Text)' OnAction: '$($s.OnAction)'"
                }
            }
        }
    }
    $wb.Close($false)
} catch {
    Write-Host "Error: $($_.Exception.Message)"
} finally {
    $xl.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
}
