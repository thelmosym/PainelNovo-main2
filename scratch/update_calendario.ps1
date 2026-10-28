$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false

$filePath = "d:\PainelNovo-main\PainelAgosto\Painel Agosto.xlsm"

try {
    Write-Host "Opening $filePath..."
    $wb = $xl.Workbooks.Open($filePath)
    
    $q = $wb.Queries.Item('Calendario')
    Write-Host "Current Calendario formula:"
    Write-Host $q.Formula
    Write-Host "----------------------------------"
    
    $newFormula = @"
let
    Fonte = Excel.Workbook(File.Contents(localTabelaB), null, true),
    Calendario_Table = Fonte{[Item="Calendario",Kind="Table"]}[Data],
    #"Tipo Alterado" = Table.TransformColumnTypes(Calendario_Table,{{"Data", type date}, {"Descrição do feriado", type text}, {"Tipo de feriado", type text}, {"Município", type text}, {"UF", type text}})
in
    #"Tipo Alterado"
"@
    
    $q.Formula = $newFormula
    Write-Host "New formula set successfully!"
    Write-Host "Updated Calendario formula:"
    Write-Host $q.Formula
    Write-Host "----------------------------------"
    
    $wb.Save()
    Write-Host "Workbook saved successfully."
    $wb.Close($false)
} catch {
    Write-Host "ERROR: $($_.Exception.Message)"
} finally {
    $xl.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
}
