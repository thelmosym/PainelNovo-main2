$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
try {
    $path = [System.IO.Path]::GetFullPath("Demonstrativo_Calculo_FDM_TSPE_IAPFARQ.xlsx")
    Write-Host "Opening: $path"
    $wb = $xl.Workbooks.Open($path)
    Write-Host "Success opening workbook!"
    $wb.Calculate()
    $wb.Save()
    $wb.Close($false)
    Write-Host "Saved and closed!"
} catch {
    Write-Host "Error: $($_.Exception.ToString())"
} finally {
    $xl.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
}
