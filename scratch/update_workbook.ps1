param(
    [Parameter(Mandatory=$true)]
    [string]$WorkbookPath
)

$basPath = "D:\PainelNovo-main\Codes\VBA\modAnaliseMonitoramento.bas"

Write-Host "--- Processing: $(Split-Path $WorkbookPath -Leaf) ---"

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.ScreenUpdating = $false
$excel.AutomationSecurity = 1 # msoAutomationSecurityLow

try {
    # 1. Open and replace module
    $wb = $excel.Workbooks.Open($WorkbookPath)
    $vbProj = $wb.VBProject
    
    $comp = $null
    foreach ($c in $vbProj.VBComponents) {
        if ($c.Name -eq "modAnaliseMonitoramento") {
            $comp = $c
            break
        }
    }
    if ($comp -ne $null) {
        $vbProj.VBComponents.Remove($comp)
    }
    $newComp = $vbProj.VBComponents.Import($basPath)
    
    $wb.Save()
    $wb.Close($true)
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) | Out-Null
    
    # 2. Re-open to register macro and run silent audit
    $wb = $excel.Workbooks.Open($WorkbookPath)
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $excel.Run("AnalisarErrosMonitoramento_Silencioso")
    $sw.Stop()
    Write-Host "Audit completed successfully in $($sw.ElapsedMilliseconds) ms!"
    
    $wb.Save()
    $wb.Close($true)
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) | Out-Null
}
catch {
    Write-Host "ERROR: $($_.Exception.Message)"
}
finally {
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}

