$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false

$filePath = "D:\PainelNovo-main\Painel de controle MemoriaPetrobras V82.xlsm"

try {
    Write-Host "Opening workbook $filePath..."
    $wb = $xl.Workbooks.Open($filePath)
    
    Write-Host "Running AuditarDadosSilencioso..."
    $xl.Run("AuditarDadosSilencioso")
    
    Write-Host "Saving workbook..."
    $wb.Save()
    $wb.Close($false)
    Write-Host "Macro execution completed."
} catch {
    Write-Host "ERROR: $($_.Exception.Message)"
} finally {
    $xl.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
}
