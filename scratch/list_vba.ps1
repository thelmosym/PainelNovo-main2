$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
try {
    $wb = $xl.Workbooks.Open("d:\PainelNovo-main\Painel de controle MemoriaPetrobras V84.xlsm", $false, $true)
    Write-Host "Total components: $($wb.VBProject.VBComponents.Count)"
    foreach ($comp in $wb.VBProject.VBComponents) {
        Write-Host "$($comp.Name) (Type: $($comp.Type))"
    }
} finally {
    if ($wb) { $wb.Close($false) }
    $xl.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl) | Out-Null
}

