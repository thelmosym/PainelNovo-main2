$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.AutomationSecurity = 1

$wbPath = "D:\PainelNovo-main\Monitoramento\Monitoramento individual_A4UU_v4.xlsm"

try {
    $wb = $excel.Workbooks.Open($wbPath)
    $wsM = $wb.Worksheets('Monitoramento')
    
    # Check row 10 original values
    $origCont = $wsM.Cells(10, 4).Value2
    $origLoc = $wsM.Cells(10, 22).Value2
    $origSit = $wsM.Cells(10, 3).Value2
    Write-Host "Row 10 original: Contrato=$origCont, Loc=$origLoc, Sit=$origSit"
    
    # TEST 1: Change Contrato to PA-LT2 (divergent, because CENPES is IRON-LT1 / RJ)
    Write-Host "Setting Row 10 Contrato to PA-LT2 (mismatch test)..."
    $wsM.Cells(10, 4).Value = "PA-LT2"
    
    # Run audit
    $excel.Run("AnalisarErrosMonitoramento_Silencioso")
    $logTest1 = $wsM.Cells(10, 26).Value2
    Write-Host "Row 10 LOG after TEST 1 (Contrato mismatch):"
    Write-Host "  LOG: $logTest1"
    
    # TEST 2: Change Localidade to nonexistent
    Write-Host "Setting Row 10 Localidade to 'LOCALIDADE_TESTE_FANTASMA'..."
    $wsM.Cells(10, 22).Value = "LOCALIDADE_TESTE_FANTASMA"
    $excel.Run("AnalisarErrosMonitoramento_Silencioso")
    $logTest2 = $wsM.Cells(10, 26).Value2
    Write-Host "Row 10 LOG after TEST 2 (Unknown localidade):"
    Write-Host "  LOG: $logTest2"
    
    # RESTORE original values
    Write-Host "Restoring Row 10 original values..."
    $wsM.Cells(10, 4).Value = $origCont
    $wsM.Cells(10, 22).Value = $origLoc
    $excel.Run("AnalisarErrosMonitoramento_Silencioso")
    $logRestored = $wsM.Cells(10, 26).Value2
    Write-Host "Row 10 LOG after RESTORE (should be empty): '$logRestored'"
    
    $wb.Save()
    $wb.Close($true)
    Write-Host "Audit verification completed successfully!"
}
catch {
    Write-Host "ERROR: $($_.Exception.Message)"
}
finally {
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}

