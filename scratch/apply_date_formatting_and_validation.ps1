# Script para aplicação de Validação de Dados e Formatação Personalizada de Datas
$files = Get-ChildItem "D:\PainelNovo-main\Monitoramento\*.xlsm" | Where-Object { -not ($_.Name.StartsWith("~")) }

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.ScreenUpdating = $false
$excel.AutomationSecurity = 1 # msoAutomationSecurityLow

Write-Host "================================================================================"
Write-Host "APLICANDO VALIDAÇÃO DE DADOS E FORMATAÇÃO DE DATAS NAS PASTAS DE MONITORAMENTO"
Write-Host "================================================================================"

foreach ($f in $files) {
    Write-Host "`n>>> Processando: $($f.Name)..."
    try {
        $wb = $excel.Workbooks.Open($f.FullName)
        $ws = $wb.Sheets.Item("Monitoramento")

        # 1. Coluna B (Período faturamento - dd/mm/aaaa)
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
        Write-Host "  [OK] Coluna B: Formatação 'dd/mm/aaaa' e Validação de Data aplicadas."

        # 2. Colunas L, Q, R (D. solicitação, D. abertura, D. fechamento - dd/mm/aaaa hh:mm)
        $colunasHora = @(
            @{ Col = "L"; Nome = "Data da Solicitação" },
            @{ Col = "Q"; Nome = "Data de Abertura" },
            @{ Col = "R"; Nome = "Data de Fechamento" }
        )

        foreach ($item in $colunasHora) {
            $col = $item.Col
            $nome = $item.Nome
            $rng = $ws.Range("${col}3:${col}2000")
            $rng.NumberFormatLocal = "dd/mm/aaaa hh:mm"
            $rng.Validation.Delete()
            $rng.Validation.Add(4, 1, 1, "01/01/2020", "31/12/2035")
            $rng.Validation.IgnoreBlank = $true
            $rng.Validation.ErrorTitle = "Data e Hora Inválida"
            $rng.Validation.ErrorMessage = "Por favor, insira a data e horário no formato padrão: dd/mm/aaaa hh:mm (Exemplo: 26/10/2026 14:30)."
            $rng.Validation.ShowError = $true
            $rng.Validation.InputTitle = $nome
            $rng.Validation.InputMessage = "Digite no formato: dd/mm/aaaa hh:mm."
            $rng.Validation.ShowInput = $true
            Write-Host "  [OK] Coluna ${col} (${nome}): Formatação 'dd/mm/aaaa hh:mm' e Validação aplicadas."
        }

        # Salva o arquivo com as validações
        $wb.Save()

        # Executa a auditoria silenciosa para validar o ecossistema
        $sw = [System.Diagnostics.Stopwatch]::StartNew()
        $excel.Run("AnalisarErrosMonitoramento_Silencioso")
        $sw.Stop()
        Write-Host "  [OK] Auditoria executada com sucesso em $($sw.ElapsedMilliseconds) ms!"

        $wb.Save()
        $wb.Close($true)
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) | Out-Null
        Write-Host "  [SUCESSO] $($f.Name) salvo com sucesso!"
    }
    catch {
        Write-Host "  [ERRO] Falha ao processar $($f.Name): $($_.Exception.Message)"
        if ($wb) { $wb.Close($false) }
    }
}

$excel.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()

Write-Host "`n================================================================================"
Write-Host "TODAS AS PASTAS DE MONITORAMENTO FORAM FORMATADAS E VALIDADAS COM SUCESSO!"
Write-Host "================================================================================"
