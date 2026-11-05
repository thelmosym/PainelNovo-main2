$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.AutomationSecurity = 1

$wb = $excel.Workbooks.Open('D:\PainelNovo-main\Monitoramento\Monitoramento individual_DPBR_v4.xlsm')

$wbComp = $wb.VBProject.VBComponents.Item('EstaPasta_de_trabalho')
$testSub = @"
Sub TestCompile()
    modAnaliseMonitoramento.AnalisarErrosMonitoramento_Silencioso
End Sub
"@

$wbComp.CodeModule.AddFromString($testSub)
Write-Host "Added TestCompile to EstaPasta_de_trabalho"

try {
    $excel.Run("TestCompile")
    Write-Host "TestCompile SUCCEEDED!"
} catch {
    Write-Host "TestCompile Failed: $($_.Exception.Message)"
}

$wb.Close($false)
$excel.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null

