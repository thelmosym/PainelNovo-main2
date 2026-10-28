$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$xl.AutomationSecurity = 1

try {
    $wb = $xl.Workbooks.Open('D:\PainelNovo-main\Painel de controle MemoriaPetrobras V82.xlsm')
    $macroName = "'" + $wb.Name + "'!AuditarDadosEGerarLog"
    Write-Host "Running $macroName..."
    $xl.Run($macroName)
    Write-Host "Macro run succeeded!"
    $wb.Save()
    $wb.Close($false)
} catch {
    Write-Host "Error: $($_.Exception.Message)"
} finally {
    $xl.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
}
