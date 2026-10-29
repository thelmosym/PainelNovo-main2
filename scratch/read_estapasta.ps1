$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
try {
    $wb = $xl.Workbooks.Open('D:\PainelNovo-main\Painel de controle MemoriaPetrobras V84.xlsm', $false, $true)
    $comp = $wb.VBProject.VBComponents.Item('EstaPastaDeTrabalho')
    Write-Host "Lines in EstaPastaDeTrabalho: $($comp.CodeModule.CountOfLines)"
    if ($comp.CodeModule.CountOfLines -gt 0) {
        Write-Host $comp.CodeModule.Lines(1, $comp.CodeModule.CountOfLines)
    }
} finally {
    if ($wb) { $wb.Close($false) }
    $xl.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
}

