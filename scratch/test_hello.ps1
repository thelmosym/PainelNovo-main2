$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.AutomationSecurity = 1

$wb = $excel.Workbooks.Open('D:\PainelNovo-main\Monitoramento\Monitoramento individual_DPBR_v4.xlsm')

$newMod = $wb.VBProject.VBComponents.Add(1)
$newMod.Name = "modTestTemp"
$newMod.CodeModule.AddFromString(@"
Public Sub HelloTest()
    Debug.Print "Hello from HelloTest"
End Sub
"@)

try {
    $excel.Run("HelloTest")
    Write-Host "HelloTest SUCCEEDED!"
} catch {
    Write-Host "HelloTest Failed: $($_.Exception.Message)"
}

$wb.VBProject.VBComponents.Remove($newMod)
$wb.Close($false)
$excel.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null

