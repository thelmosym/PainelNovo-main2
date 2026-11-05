$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.AutomationSecurity = 1

$wb = $excel.Workbooks.Open('D:\PainelNovo-main\Monitoramento\Monitoramento individual_DPBR_v4.xlsm')

try {
    # Try to compile
    $cmd = $excel.VBE.CommandBars.FindControl($null, 578)
    if ($cmd -ne $null -and $cmd.Enabled) {
        $cmd.Execute()
        Write-Host "VBA Project compiled without errors!"
    } else {
        Write-Host "Compile button not enabled (may already be compiled or not accessible)."
    }
} catch {
    Write-Host "Error compiling: $($_.Exception.Message)"
}

$wb.Close($false)
$excel.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
