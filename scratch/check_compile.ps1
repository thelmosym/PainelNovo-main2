$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$xl.AutomationSecurity = 1

try {
    $filePath = 'D:\PainelNovo-main\Painel de controle MemoriaPetrobras V84.xlsm'
    $wb = $xl.Workbooks.Open($filePath)
    
    # Try compiling VBA Project
    Write-Host "Triggering VBA Compile..."
    # Id 578 is the Debug > Compile menu command
    $control = $xl.VBE.CommandBars.FindControl($null, 578)
    if ($control -ne $null) {
        Write-Host "Compile control found, Enabled: $($control.Enabled)"
        if ($control.Enabled) {
            $control.Execute()
            Write-Host "Compile executed!"
        } else {
            Write-Host "Project is already compiled or not dirty."
        }
    } else {
        Write-Host "Compile control not found."
    }
    $wb.Close($false)
} catch {
    Write-Host "Exception: $($_.Exception.Message)"
} finally {
    $xl.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
}

