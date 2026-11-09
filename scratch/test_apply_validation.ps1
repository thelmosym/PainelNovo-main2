$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.ScreenUpdating = $false
$excel.AutomationSecurity = 1

$testWb = "D:\PainelNovo-main\scratch\test_A4UU.xlsm"
Copy-Item "D:\PainelNovo-main\Monitoramento\Monitoramento individual_A4UU_v6.xlsm" $testWb -Force

$wb = $excel.Workbooks.Open($testWb)
$ws = $wb.Sheets.Item("Monitoramento")

# 1. Coluna B (Período faturamento)
$rangeB = $ws.Range("B3:B2000")
$rangeB.NumberFormatLocal = "dd/mm/aaaa"
$rangeB.Validation.Delete()
$rangeB.Validation.Add(4, 1, 1, "01/01/2020", "31/12/2035")
$rangeB.Validation.IgnoreBlank = $true
$rangeB.Validation.ErrorTitle = "Data Inválida"
$rangeB.Validation.ErrorMessage = "Insira uma data válida no formato dd/mm/aaaa (Exemplo: 26/10/2026)."
$rangeB.Validation.ShowError = $true
$rangeB.Validation.InputTitle = "Período de Faturamento"
$rangeB.Validation.InputMessage = "Digite a data no formato dd/mm/aaaa."
$rangeB.Validation.ShowInput = $true

# 2. Colunas L, Q, R (D. solicitação, D. abertura, D. fechamento)
$colunasHora = @("L", "Q", "R")
foreach ($col in $colunasHora) {
    $rng = $ws.Range("${col}3:${col}2000")
    $rng.NumberFormatLocal = "dd/mm/aaaa hh:mm"
    $rng.Validation.Delete()
    $rng.Validation.Add(4, 1, 1, "01/01/2020", "31/12/2035")
    $rng.Validation.IgnoreBlank = $true
    $rng.Validation.ErrorTitle = "Data e Hora Inválida"
    $rng.Validation.ErrorMessage = "Por favor, insira a data e horário no formato padrão: dd/mm/aaaa hh:mm (Exemplo: 26/10/2026 14:30)."
    $rng.Validation.ShowError = $true
    $rng.Validation.InputTitle = "Data e Hora"
    $rng.Validation.InputMessage = "Digite no formato: dd/mm/aaaa hh:mm."
    $rng.Validation.ShowInput = $true
}

$wb.Save()
$wb.Close($true)
$excel.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
Write-Host "Test workbook saved successfully!"
