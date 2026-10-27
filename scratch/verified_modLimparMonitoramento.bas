Attribute VB_Name = "modLimparMonitoramento"
'===================================================================================================================
' MÓDULO: modLimparMonitoramento
' OBJETIVO: Limpeza e filtragem da aba "Monitoramento", removendo todas as linhas da tabela estruturada
'           (ListObject) cuja coluna 'Situação' seja diferente de 'NC' (mantendo exclusivamente as linhas com 'NC')
'           e classificando as linhas mantidas em ordem crescente pela coluna 'Contrato'.
' APLICAÇÃO: Planilha de Monitoramento Individual (A4UU, DPBR, GPZ1, GQ6S, S2IJ, etc.)
'
' REGRAS IMPLEMENTADAS:
'   1. IDENTIFICAÇÃO DINÂMICA: Localiza as colunas 'Situação' e 'Contrato' pelo cabeçalho (com fallback de índice).
'   2. FILTRAGEM E EXCLUSÃO: Remove todas as linhas onde Situação <> "NC", preservando apenas registros NC.
'   3. CLASSIFICAÇÃO: Ordena as linhas mantidas de forma crescente pela coluna 'Contrato'.
'   4. PRESERVAÇÃO INTEGRAL: Mantém formatações, validações de dados (listas suspensas) e alinhamentos.
'   5. LIMPEZA RESIDUAL: Elimina linhas e células residuais vazias fora da tabela.
'   6. ALTA PERFORMANCE: Desabilita temporariamente atualização de tela, cálculo automático e alertas.
'
' CODIFICAÇÃO: Windows-1252 (CP1252 / ANSI) com quebras de linha CRLF para compatibilidade VBE.
'===================================================================================================================
Option Explicit

Private Const MSG_TITULO As String = "Limpeza do Monitoramento (Manter NC)"

' Estrutura para armazenamento seguro das propriedades de validação
Private Type InfoValidacao
    TemValidacao As Boolean
    Tipo As Long
    EstiloAlerta As Long
    Operador As Long
    Formula1 As String
    Formula2 As String
    IgnorarBranco As Boolean
    ListaSuspensa As Boolean
    ExibirEntrada As Boolean
    ExibirErro As Boolean
    TituloEntrada As String
    MsgEntrada As String
    TituloErro As String
    MsgErro As String
End Type

'-------------------------------------------------------------------------------------------------------------------
' PROCEDIMENTO PRINCIPAL (PONTO DE ENTRADA DO USUÁRIO)
' Vinculável a botões na aba "Menu" ou "Monitoramento", ou executável via Alt+F8.
'-------------------------------------------------------------------------------------------------------------------
Public Sub LimparMonitoramento()
    LimparAbaMonitoramentoManterNC False
End Sub

'-------------------------------------------------------------------------------------------------------------------
' ALIAS PARA COMPATIBILIDADE COM VERSÕES E SCRIPTS ANTERIORES
'-------------------------------------------------------------------------------------------------------------------
Public Sub LimparAbaMonitoramento100Linhas(Optional ByVal bSilencioso As Boolean = False)
    LimparAbaMonitoramentoManterNC bSilencioso
End Sub

'-------------------------------------------------------------------------------------------------------------------
' PROCEDIMENTO PRINCIPAL PARAMETRIZADO:
' Remove todas as linhas onde Situação <> "NC" e classifica as linhas restantes por Contrato.
'-------------------------------------------------------------------------------------------------------------------
Public Sub LimparAbaMonitoramentoManterNC(Optional ByVal bSilencioso As Boolean = False)
    Dim sEtapaAtual As String
    sEtapaAtual = "Inicialização"

    On Error GoTo TratarErro

    Dim wsMonit As Worksheet
    Dim tbl As ListObject
    Dim appCalc As XlCalculation
    Dim resposta As VbMsgBoxResult
    Dim lLinhaFimTabela As Long
    Dim lUltimaLinhaPlan As Long
    Dim lColFimTabela As Long
    Dim lUltimaColPlan As Long
    Dim vValidacoes() As InfoValidacao
    Dim sNomeTabela As String
    Dim colSituacao As Long
    Dim colContrato As Long
    Dim i As Long
    Dim sValSituacao As String
    Dim lTotalInicial As Long
    Dim lRemovidas As Long
    Dim lMantidas As Long

    ' 1. Confirmação do Usuário (se não estiver em modo silencioso)
    If Not bSilencioso Then
        resposta = MsgBox("ATENÇÃO: Este procedimento irá:" & vbNewLine & vbNewLine & _
                          "1. Analisar todas as linhas da tabela 'Monitoramento'." & vbNewLine & _
                          "2. Remover todas as linhas cuja Situação NÃO seja 'NC' (mantendo apenas 'NC')." & vbNewLine & _
                          "3. Classificar as linhas mantidas pelo Contrato em ordem crescente." & vbNewLine & _
                          "4. Preservar as validações de dados e formatações das colunas." & vbNewLine & vbNewLine & _
                          "Deseja realmente continuar?", vbQuestion + vbYesNo + vbDefaultButton2, MSG_TITULO)
        If resposta <> vbYes Then Exit Sub
    End If

    ' 2. Salvar e Desativar Recursos de Interface para Máxima Performance
    sEtapaAtual = "Configuração do ambiente"
    appCalc = Application.Calculation
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    ' 3. Localizar Planilha 'Monitoramento' (ActiveWorkbook com fallback para ThisWorkbook)
    sEtapaAtual = "Localização da aba Monitoramento"
    Set wsMonit = ObterAbaMonitoramento()
    If wsMonit Is Nothing Then
        MsgBox "[ERRO] A aba 'Monitoramento' não foi encontrada na pasta de trabalho ativa.", vbCritical, MSG_TITULO
        GoTo SairRotina
    End If

    ' 4. Localizar Tabela Estruturada (ListObject)
    sEtapaAtual = "Localização da tabela estruturada"
    Set tbl = ObterTabelaMonitoramento(wsMonit)
    If tbl Is Nothing Then
        MsgBox "[ERRO] Nenhuma tabela estruturada encontrada na aba 'Monitoramento'.", vbCritical, MSG_TITULO
        GoTo SairRotina
    End If

    sNomeTabela = tbl.Name

    ' 5. Desativar quaisquer filtros ativos (crucial para evitar Erro 1004)
    sEtapaAtual = "Limpeza de filtros ativos"
    On Error Resume Next
    If wsMonit.FilterMode Then wsMonit.ShowAllData
    If Not tbl.AutoFilter Is Nothing Then
        If tbl.AutoFilter.FilterMode Then tbl.AutoFilter.ShowAllData
    End If
    On Error GoTo TratarErro

    ' 6. Identificar Colunas 'Situação' e 'Contrato' dinamicamente
    sEtapaAtual = "Identificação das colunas Situação e Contrato"
    colSituacao = ObterIndiceColuna(tbl, "Situa", 3)
    colContrato = ObterIndiceColuna(tbl, "Contrato", 4)

    ' 7. Mapear e Salvar Validações de Dados das Colunas Existentes
    sEtapaAtual = "Mapeamento das validações de dados"
    ReDim vValidacoes(1 To tbl.ListColumns.Count)
    CapturarValidacoesTabela tbl, vValidacoes

    ' 8. Excluir Linhas Onde Situação <> 'NC' (de baixo para cima)
    sEtapaAtual = "Filtragem e exclusão de linhas não-NC"
    lTotalInicial = tbl.ListRows.Count
    lRemovidas = 0
    lMantidas = 0

    If lTotalInicial > 0 Then
        For i = lTotalInicial To 1 Step -1
            sValSituacao = UCase$(Trim$(CStr(tbl.ListRows(i).Range.Cells(1, colSituacao).Value)))
            If sValSituacao = "NC" Then
                lMantidas = lMantidas + 1
            Else
                tbl.ListRows(i).Delete
                lRemovidas = lRemovidas + 1
            End If
        Next i
    End If

    ' 9. Classificar Linhas Mantidas pelo Contrato
    sEtapaAtual = "Classificação das linhas pelo Contrato"
    If tbl.ListRows.Count > 1 And colContrato > 0 Then
        With tbl.Sort
            .SortFields.Clear
            .SortFields.Add Key:=tbl.ListColumns(colContrato).Range, _
                            SortOn:=xlSortOnValues, _
                            Order:=xlAscending, _
                            DataOption:=xlSortNormal
            .Header = xlYes
            .MatchCase = False
            .Apply
        End With
    End If

    ' 10. Excluir Linhas Residuais Abaixo da Tabela na Planilha
    sEtapaAtual = "Exclusão de linhas residuais abaixo da tabela"
    lLinhaFimTabela = tbl.Range.Row + tbl.Range.Rows.Count - 1
    lUltimaLinhaPlan = wsMonit.Cells.SpecialCells(xlCellTypeLastCell).Row
    If lUltimaLinhaPlan > lLinhaFimTabela Then
        wsMonit.Rows((lLinhaFimTabela + 1) & ":" & lUltimaLinhaPlan).Delete xlShiftUp
    End If

    ' 11. Limpar Células Residuais à Direita da Tabela (sem usar Delete)
    sEtapaAtual = "Limpeza de células à direita da tabela"
    lColFimTabela = tbl.Range.Column + tbl.Range.Columns.Count - 1
    lUltimaColPlan = wsMonit.Cells.SpecialCells(xlCellTypeLastCell).Column
    If lUltimaColPlan > lColFimTabela Then
        wsMonit.Range(wsMonit.Cells(1, lColFimTabela + 1), _
                      wsMonit.Cells(wsMonit.Rows.Count, lUltimaColPlan)).Clear
    End If

    ' 12. Restaurar e Garantir Validação de Dados nas Linhas Mantidas
    sEtapaAtual = "Reaplicação das validações de dados"
    ReaplicarValidacoesTabela tbl, vValidacoes, wsMonit

    ' 13. Formatação Padronizada de Tipos de Colunas Específicas
    sEtapaAtual = "Padronização de formatos e alinhamentos"
    FormatarColunasPadrao tbl

    ' 14. Garantir Ativação dos Estilos Nativos da Tabela e Filtros
    On Error Resume Next
    tbl.ShowTableStyleRowStripes = True
    tbl.ShowAutoFilter = True
    On Error GoTo TratarErro

    ' 15. Restaurar Ambiente e Posicionar Cursor de Forma Segura na Célula A3
    sEtapaAtual = "Finalização e posicionamento"
    Application.ScreenUpdating = True
    On Error Resume Next
    Application.GoTo wsMonit.Range("A3"), True
    On Error GoTo TratarErro

    ' 16. Mensagem de Sucesso (se não estiver em modo silencioso)
    If Not bSilencioso Then
        MsgBox "[OK] Filtragem e organização da aba 'Monitoramento' concluída com sucesso!" & vbNewLine & vbNewLine & _
               "- Tabela: '" & sNomeTabela & "'" & vbNewLine & _
               "- Linhas mantidas (Situação = NC): " & lMantidas & vbNewLine & _
               "- Linhas removidas: " & lRemovidas & vbNewLine & _
               "- Linhas classificadas em ordem crescente por 'Contrato'." & vbNewLine & _
               "- Validações de dados e formatações preservadas.", _
               vbInformation, MSG_TITULO
    End If

SairRotina:
    ' Restauração estrita dos estados da aplicação
    On Error Resume Next
    Application.Calculation = appCalc
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Exit Sub

TratarErro:
    MsgBox "[ERRO] Ocorreu um erro durante a limpeza da aba 'Monitoramento':" & vbNewLine & vbNewLine & _
           "Etapa: " & sEtapaAtual & vbNewLine & _
           "Código: " & Err.Number & vbNewLine & _
           "Descrição: " & Err.Description, vbCritical, MSG_TITULO
    Resume SairRotina
End Sub

'-------------------------------------------------------------------------------------------------------------------
' FUNÇÃO AUXILIAR: Localiza a aba 'Monitoramento' com segurança
'-------------------------------------------------------------------------------------------------------------------
Private Function ObterAbaMonitoramento() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ActiveWorkbook.Sheets("Monitoramento")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets("Monitoramento")
    End If
    On Error GoTo 0
    Set ObterAbaMonitoramento = ws
End Function

'-------------------------------------------------------------------------------------------------------------------
' FUNÇÃO AUXILIAR: Localiza o ListObject da aba 'Monitoramento'
'-------------------------------------------------------------------------------------------------------------------
Private Function ObterTabelaMonitoramento(ByVal ws As Worksheet) As ListObject
    Dim tbl As ListObject
    On Error Resume Next
    Set tbl = ws.ListObjects("Monitoramento")
    If tbl Is Nothing And ws.ListObjects.Count > 0 Then
        Set tbl = ws.ListObjects(1)
    End If
    On Error GoTo 0
    Set ObterTabelaMonitoramento = tbl
End Function

'-------------------------------------------------------------------------------------------------------------------
' FUNÇÃO AUXILIAR: Localiza o índice de uma coluna pelo nome ou parte do nome (com índice de fallback)
'-------------------------------------------------------------------------------------------------------------------
Private Function ObterIndiceColuna(ByVal tbl As ListObject, ByVal sTrechoNome As String, ByVal lIndiceFallback As Long) As Long
    Dim c As Long
    Dim sNomeLimpo As String
    Dim sTrechoLimpo As String

    sTrechoLimpo = UCase$(Trim$(sTrechoNome))

    For c = 1 To tbl.ListColumns.Count
        sNomeLimpo = Replace(Replace(UCase$(tbl.ListColumns(c).Name), vbLf, ""), vbCr, "")
        sNomeLimpo = Trim$(sNomeLimpo)
        If InStr(1, sNomeLimpo, sTrechoLimpo, vbTextCompare) > 0 Then
            ObterIndiceColuna = c
            Exit Function
        End If
    Next c

    If lIndiceFallback > 0 And lIndiceFallback <= tbl.ListColumns.Count Then
        ObterIndiceColuna = lIndiceFallback
    Else
        ObterIndiceColuna = 1
    End If
End Function

'-------------------------------------------------------------------------------------------------------------------
' SUBROTINA AUXILIAR: Mapeia as regras de validação presentes em cada coluna
'-------------------------------------------------------------------------------------------------------------------
Private Sub CapturarValidacoesTabela(ByVal tbl As ListObject, ByRef vValidacoes() As InfoValidacao)
    Dim c As Long
    Dim r As Long
    Dim rngCol As Range
    Dim cellTest As Range
    Dim vObj As Validation
    Dim lMaxLinhasTeste As Long

    For c = 1 To tbl.ListColumns.Count
        vValidacoes(c).TemValidacao = False
        Set rngCol = tbl.ListColumns(c).DataBodyRange

        If Not rngCol Is Nothing Then
            lMaxLinhasTeste = Application.WorksheetFunction.Min(rngCol.Rows.Count, 5)
            For r = 1 To lMaxLinhasTeste
                Set cellTest = rngCol.Cells(r, 1)
                On Error Resume Next
                Set vObj = cellTest.Validation
                If Not vObj Is Nothing Then
                    If vObj.Type <> 0 Then
                        vValidacoes(c).TemValidacao = True
                        vValidacoes(c).Tipo = vObj.Type
                        vValidacoes(c).EstiloAlerta = vObj.AlertStyle
                        vValidacoes(c).Operador = vObj.Operator
                        vValidacoes(c).Formula1 = vObj.Formula1
                        vValidacoes(c).Formula2 = ""
                        vValidacoes(c).Formula2 = vObj.Formula2
                        vValidacoes(c).IgnorarBranco = vObj.IgnoreBlank
                        vValidacoes(c).ListaSuspensa = vObj.InCellDropdown
                        vValidacoes(c).ExibirEntrada = vObj.ShowInput
                        vValidacoes(c).ExibirErro = vObj.ShowError
                        vValidacoes(c).TituloEntrada = vObj.InputTitle
                        vValidacoes(c).MsgEntrada = vObj.InputMessage
                        vValidacoes(c).TituloErro = vObj.ErrorTitle
                        vValidacoes(c).MsgErro = vObj.ErrorMessage
                        On Error GoTo 0
                        Exit For
                    End If
                End If
                On Error GoTo 0
            Next r
        End If
    Next c
End Sub

'-------------------------------------------------------------------------------------------------------------------
' SUBROTINA AUXILIAR: Reaplica as validações de dados capturadas e injeta regras de fallback
'-------------------------------------------------------------------------------------------------------------------
Private Sub ReaplicarValidacoesTabela(ByVal tbl As ListObject, ByRef vValidacoes() As InfoValidacao, ByVal wsMonit As Worksheet)
    Dim c As Long
    Dim rngCol As Range
    Dim sFormulaFbk As String
    Dim bTemTabelaA As Boolean

    On Error Resume Next
    bTemTabelaA = Not (wsMonit.Parent.Sheets("TabelaA") Is Nothing)
    On Error GoTo 0

    For c = 1 To tbl.ListColumns.Count
        Set rngCol = tbl.ListColumns(c).DataBodyRange
        If Not rngCol Is Nothing Then
            If vValidacoes(c).TemValidacao Then
                AplicarValidacaoNoRange rngCol, vValidacoes(c)
            ElseIf bTemTabelaA Then
                sFormulaFbk = ObterFormulaFallbackPorIndice(c)
                If Len(sFormulaFbk) > 0 Then
                    Dim infoFbk As InfoValidacao
                    infoFbk.TemValidacao = True
                    infoFbk.Tipo = 3 ' xlValidateList
                    infoFbk.EstiloAlerta = 1 ' xlValidAlertStop
                    infoFbk.Operador = 1 ' xlBetween
                    infoFbk.Formula1 = sFormulaFbk
                    infoFbk.IgnorarBranco = True
                    infoFbk.ListaSuspensa = True
                    infoFbk.ExibirEntrada = True
                    infoFbk.ExibirErro = True
                    AplicarValidacaoNoRange rngCol, infoFbk
                End If
            End If
        End If
    Next c
End Sub

'-------------------------------------------------------------------------------------------------------------------
' SUBROTINA AUXILIAR: Aplica uma estrutura de validação em um determinado Range
'-------------------------------------------------------------------------------------------------------------------
Private Sub AplicarValidacaoNoRange(ByVal rngAlvo As Range, ByRef info As InfoValidacao)
    On Error Resume Next
    rngAlvo.Validation.Delete
    If Len(info.Formula2) > 0 Then
        rngAlvo.Validation.Add Type:=info.Tipo, AlertStyle:=info.EstiloAlerta, _
                               Operator:=info.Operador, Formula1:=info.Formula1, Formula2:=info.Formula2
    Else
        rngAlvo.Validation.Add Type:=info.Tipo, AlertStyle:=info.EstiloAlerta, _
                               Operator:=info.Operador, Formula1:=info.Formula1
    End If
    rngAlvo.Validation.IgnoreBlank = info.IgnorarBranco
    rngAlvo.Validation.InCellDropdown = info.ListaSuspensa
    rngAlvo.Validation.ShowInput = info.ExibirEntrada
    rngAlvo.Validation.ShowError = info.ExibirErro
    If Len(info.TituloEntrada) > 0 Then rngAlvo.Validation.InputTitle = info.TituloEntrada
    If Len(info.MsgEntrada) > 0 Then rngAlvo.Validation.InputMessage = info.MsgEntrada
    If Len(info.TituloErro) > 0 Then rngAlvo.Validation.ErrorTitle = info.TituloErro
    If Len(info.MsgErro) > 0 Then rngAlvo.Validation.ErrorMessage = info.MsgErro
    On Error GoTo 0
End Sub

'-------------------------------------------------------------------------------------------------------------------
' FUNÇÃO AUXILIAR: Fórmulas de fallback para as colunas padrão de Monitoramento Individual
'-------------------------------------------------------------------------------------------------------------------
Private Function ObterFormulaFallbackPorIndice(ByVal c As Long) As String
    Select Case c
        Case 3  ' Coluna C: Situação
            ObterFormulaFallbackPorIndice = "=TabelaA!$B$2:$B$5"
        Case 4  ' Coluna D: Contrato
            ObterFormulaFallbackPorIndice = "=TabelaA!$E$2:$E$6"
        Case 5  ' Coluna E: UF
            ObterFormulaFallbackPorIndice = "=TabelaA!$AD$2:$AD$57"
        Case 6  ' Coluna F: Descrição da atividade
            ObterFormulaFallbackPorIndice = "=TabelaA!$AO$2:$AO$35"
        Case 7  ' Coluna G: Item
            ObterFormulaFallbackPorIndice = "=TabelaA!$AQ$2:$AQ$53"
        Case 8  ' Coluna H: Aplicação
            ObterFormulaFallbackPorIndice = "=TabelaA!$AH$2:$AH$3"
        Case 25 ' Coluna Y: OS disponibilizada? / Faturamento Direto
            ObterFormulaFallbackPorIndice = "=TabelaA!$AW$2:$AW$3"
        Case Else
            ObterFormulaFallbackPorIndice = ""
    End Select
End Function

'-------------------------------------------------------------------------------------------------------------------
' SUBROTINA AUXILIAR: Formata colunas de data/hora e alinhamentos básicos
'-------------------------------------------------------------------------------------------------------------------
Private Sub FormatarColunasPadrao(ByVal tbl As ListObject)
    On Error Resume Next
    If tbl.ListRows.Count = 0 Or tbl.DataBodyRange Is Nothing Then Exit Sub

    ' Coluna B: Período faturamento
    If tbl.ListColumns.Count >= 2 Then
        tbl.ListColumns(2).DataBodyRange.NumberFormat = "dd/mm/aaaa (ddd)"
        tbl.ListColumns(2).DataBodyRange.HorizontalAlignment = xlCenter
    End If

    ' Colunas de Códigos e Siglas Curtas (Centralizadas)
    If tbl.ListColumns.Count >= 3 Then tbl.ListColumns(3).DataBodyRange.HorizontalAlignment = xlCenter ' Situação
    If tbl.ListColumns.Count >= 4 Then tbl.ListColumns(4).DataBodyRange.HorizontalAlignment = xlCenter ' Contrato
    If tbl.ListColumns.Count >= 5 Then tbl.ListColumns(5).DataBodyRange.HorizontalAlignment = xlCenter ' UF
    If tbl.ListColumns.Count >= 8 Then tbl.ListColumns(8).DataBodyRange.HorizontalAlignment = xlCenter ' Aplicação
    If tbl.ListColumns.Count >= 25 Then tbl.ListColumns(25).DataBodyRange.HorizontalAlignment = xlCenter ' OS disponibilizada

    ' Colunas de Data/Hora (dd/mm/aaaa hh:mm)
    If tbl.ListColumns.Count >= 12 Then tbl.ListColumns(12).DataBodyRange.NumberFormat = "dd/mm/aaaa hh:mm" ' D. solicitação
    If tbl.ListColumns.Count >= 17 Then tbl.ListColumns(17).DataBodyRange.NumberFormat = "dd/mm/aaaa hh:mm" ' D. abertura
    If tbl.ListColumns.Count >= 18 Then tbl.ListColumns(18).DataBodyRange.NumberFormat = "dd/mm/aaaa hh:mm" ' D. fechamento

    ' Colunas de Quantidades (Alinhadas ao centro)
    If tbl.ListColumns.Count >= 15 Then tbl.ListColumns(15).DataBodyRange.HorizontalAlignment = xlCenter ' Qtd. Solicitada
    If tbl.ListColumns.Count >= 20 Then tbl.ListColumns(20).DataBodyRange.HorizontalAlignment = xlCenter ' Qtd. Atendida
    On Error GoTo 0
End Sub
