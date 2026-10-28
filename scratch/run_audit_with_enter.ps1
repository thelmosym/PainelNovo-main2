$wshell = New-Object -ComObject Wscript.Shell

$job = Start-Job -ScriptBlock {
    for ($i = 0; $i -lt 15; $i++) {
        Start-Sleep -Seconds 1
        $w = New-Object -ComObject Wscript.Shell
        $w.AppActivate("Microsoft Excel") | Out-Null
        $w.SendKeys("{ENTER}")
    }
}

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $true
$xl.DisplayAlerts = $false
$xl.AutomationSecurity = 1

try {
    $filePath = 'D:\PainelNovo-main\Painel de controle MemoriaPetrobras V82.xlsm'
    $wb = $xl.Workbooks.Open($filePath)
    $macroName = "'" + $wb.Name + "'!AuditarDadosEGerarLog"
    Write-Host "Running $macroName..."
    $xl.Run($macroName)
    Write-Host "Macro run finished!"
    $wb.Save()
    $wb.Close($false)
    Write-Host "Workbook saved."
} catch {
    Write-Host "Error: $($_.Exception.Message)"
} finally {
    $xl.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
    Stop-Job $job -ErrorAction SilentlyContinue | Out-Null
    Remove-Job $job -ErrorAction SilentlyContinue | Out-Null
}
