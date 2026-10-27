$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
try {
    $files = Get-ChildItem -Path . -Recurse -Filter "Demonstrativo_Calculo_FDM*.xlsx"
    foreach ($f in $files) {
        $wb = $xl.Workbooks.Open($f.FullName)
        $xl.Calculate()
        $wb.Save()
        $wb.Close($false)
        Write-Host "Recalculado com sucesso: $($f.FullName)"
    }
} finally {
    $xl.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
}
