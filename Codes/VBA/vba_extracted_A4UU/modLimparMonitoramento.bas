Attribute VB_Name = "modLimparMonitoramento"
'===================================================================================================================
' MÓDULO: modLimparMonitoramento
' OBJETIVO: Limpeza e organização da aba "Monitoramento", removendo todas as linhas da tabela estruturada
'           (ListObject) cuja coluna 'Situação' esteja preenchida como 'AT', 'CO' ou 'CA', mantendo todas as
'           demais linhas com conteúdo agrupadas de forma contínua no início da tabela, classificadas de forma
'           crescente pela coluna 'D. abertura' (com desempate por 'D. solicitação' quando 'D. abertura' estiver
'           em branco), preservando todas as formatações/validações e garantindo no mínimo 700 linhas na tabela.
' APLICAÇÃO: Planilha de Monitoramento Individual (A4UU, DPBR, GPZ1, GQ6S, S2IJ, etc.)
'
' REGRAS IMPLEMENTADAS:
'   1. IDENTIFICAÇÃO DINÂMICA: Localiza 'Situação', 'D. abertura' e 'D. solicitação' pelo cabeçalho.
'   2. FILTRAGEM SELETIVA: Remove todas as linhas com Situação = "AT", "CO" ou "CA" e descarta linhas vazias
'      intermediárias (incluindo células contendo apenas espaços ou caracteres invisíveis como Chr(160)).
'   3. AGRUPAMENTO CONTÍNUO NO INÍCIO + CLASSIFICAÇÃO: Todas as linhas com conteúdo ficam juntas no topo da
'      tabela (a partir da linha 3, sem linhas vazias intercaladas), ordenadas por 'D. abertura' crescente.
'   4. MÍNIMO DE 700 LINHAS: Mantém ao menos 700 linhas na tabela (ou todas as mantidas, se > 700).
'   5. PRESERVAÇÃO INTEGRAL: Mantém estilo da tabela, formatos de número/data, alinhamentos, fontes,
'      regras de formatação condicional e validações de dados (listas suspensas) em todas as linhas.
'   6. BLINDAGEM CONTRA ERRO 1004: Remove tabelas secundárias órfãs abaixo do cabeçalho e células mescladas.
'
' CODIFICAÇÃO: Windows-1252 (CP1252 / ANSI) com quebras de linha CRLF para compatibilidade VBE.
'===================================================================================================================
Option Explicit

Private Const MSG_TITULO As String = "Limpeza do Monitoramento"
Private Const MIN_LINHAS_TABELA As Long = 700

' Estrutura para armazenamento das propriedades de validação de dados por coluna
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

' Estrutura para preservação fiel da formatação de cada coluna da tabela
Private Type InfoFormatoColuna
    FormatoLocal As String
    AlinhamentoH As Long
    AlinhamentoV As Long
    QuebrarTexto As Boolean
    NomeFonte As String
    TamanhoFonte As Double
    Negrito As Boolean
End Type

' Estrutura auxiliar para preservar regras de duplicidade da formatação condicional
Private Type InfoCondicionalDuplicados
    Ativo As Boolean
    CorFundo As Long
    CorFonte As Long
End Type

'-------------------------------------------------------------------------------------------------------------------
' PROCEDIMENTO PRINCIPAL (PONTO DE ENTRADA DO BOTÃO NA ABA 'MENU' / ALT+F8)
'-------------------------------------------------------------------------------------------------------------------
Public Sub LimparMonitoramento()
    LimparAbaMonitoramentoManterNC False
End Sub

'-------------------------------------------------------------------------------------------------------------------
' ALIAS SEMÂNTICO E ALIAS DE COMPATIBILIDADE COM SCRIPTS ANTERIORES
'-------------------------------------------------------------------------------------------------------------------
Public Sub LimparMonitoramentoManterPendentes(Optional ByVal bSilencioso As Boolean = False)
    LimparAbaMonitoramentoManterNC bSilencioso
End Sub

Public Sub LimparAbaMonitoramento100Linhas(Optional ByVal bSilencioso As Boolean = False)
    LimparAbaMonitoramentoManterNC bSilencioso
End Sub

'-------------------------------------------------------------------------------------------------------------------
' PROCEDIMENTO PRINCIPAL PARAMETRIZADO:
' Limpa linhas com Situação = 'AT', 'CO' ou 'CA', agrupa todas as linhas com conteúdo juntas no início da
' planilha ordenadas por 'D. abertura' e mantém no mínimo 700 linhas formatadas na tabela.
'-------------------------------------------------------------------------------------------------------------------
Public Sub LimparAbaMonitoramentoManterNC(Optional ByVal bSilencioso As Boolean = False)
    Dim sEtapaAtual As String
    sEtapaAtual = "Inicialização"

    On Error GoTo TratarErro

    Dim wsMonit As Worksheet
    Dim tbl As ListObject
    Dim appCalc As XlCalculation
    Dim resposta As VbMsgBoxResult
    Dim sNomeTabela As String
    Dim sEstiloTabela As String
    Dim bTirasLinhas As Boolean
    Dim dAlturaLinha As Double

    Dim lTotalCols As Long
    Dim lTotalInicial As Long
    Dim lMantidas As Long
    Dim lRemovidas As Long
    Dim lLinhasFinais As Long

    Dim colSituacao As Long
    Dim colDataAbertura As Long
    Dim colDataSolicitacao As Long

    Dim vValidacoes() As InfoValidacao
    Dim vFormatos() As InfoFormatoColuna
    Dim vCondDuplicados() As InfoCondicionalDuplicados

    Dim vDados As Variant
    Dim vSaida() As Variant
    Dim idxMantidas() As Long
    Dim rngNovoTamanho As Range

    Dim lLinhaHeader As Long
    Dim lColInicial As Long
    Dim lLinhaFimTabela As Long
    Dim lLinhaFimAnterior As Long
    Dim lUltimaLinhaPlan As Long
    Dim lColFimTabela As Long
    Dim lUltimaColPlan As Long

    Dim r As Long
    Dim k As Long
    Dim c As Long
    Dim sValSituacao As String
    Dim vCel As Variant

    ' 1. Confirmação do Usuário (quando acionado interativamente)
    If Not bSilencioso Then
        resposta = MsgBox("ATENÇÃO: Este procedimento irá:" & vbNewLine & vbNewLine & _
                          "1. Analisar todas as linhas da tabela 'Monitoramento'." & vbNewLine & _
                          "2. Remover todas as linhas com Situação 'AT', 'CO' ou 'CA'." & vbNewLine & _
                          "3. Agrupar todas as linhas com conteúdo juntas no início da tabela, ordenadas por 'D. abertura'." & vbNewLine & _
                          "4. Preservar todas as formatações e validações, mantendo no mínimo " & MIN_LINHAS_TABELA & " linhas na tabela." & vbNewLine & vbNewLine & _
                          "Deseja realmente continuar?", vbQuestion + vbYesNo + vbDefaultButton2, MSG_TITULO)
        If resposta <> vbYes Then Exit Sub
    End If

    ' 2. Salvar e Desativar Recursos de Interface para Alta Performance
    sEtapaAtual = "Configuração do ambiente"
    appCalc = Application.Calculation
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    ' 3. Localizar Planilha 'Monitoramento'
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
    lTotalCols = tbl.ListColumns.Count
    lLinhaHeader = tbl.HeaderRowRange.Row
    lColInicial = tbl.HeaderRowRange.Column

    ' 5. Desativar quaisquer filtros ativos (evita Erro 1004 e linhas ocultas)
    sEtapaAtual = "Limpeza de filtros ativos"
    On Error Resume Next
    If wsMonit.FilterMode Then wsMonit.ShowAllData
    If Not tbl.AutoFilter Is Nothing Then
        If tbl.AutoFilter.FilterMode Then tbl.AutoFilter.ShowAllData
    End If
    wsMonit.Rows.Hidden = False
    On Error GoTo TratarErro

    ' 6. Identificar Colunas 'Situação', 'D. abertura' e 'D. solicitação' dinamicamente
    sEtapaAtual = "Identificação das colunas de controle"
    colSituacao = ObterIndiceColuna(tbl, "Situa", 3)
    colDataAbertura = ObterIndiceColuna(tbl, "D. abertura", 0)
    If colDataAbertura = 0 Then
        colDataAbertura = ObterIndiceColuna(tbl, "abertura", 17)
    End If
    colDataSolicitacao = ObterIndiceColuna(tbl, "D. solicita", 0)
    If colDataSolicitacao = 0 Then
        colDataSolicitacao = ObterIndiceColuna(tbl, "solicita", 12)
    End If

    ' 7. Capturar Estilo da Tabela, Formatos de Coluna, Formatação Condicional e Validações
    sEtapaAtual = "Captura de estilos, formatos e validações"
    sEstiloTabela = "TableStyleMedium21"
    bTirasLinhas = True
    dAlturaLinha = 15#

    On Error Resume Next
    If Not tbl.TableStyle Is Nothing Then
        If Len(tbl.TableStyle.Name) > 0 Then sEstiloTabela = tbl.TableStyle.Name
    End If
    bTirasLinhas = tbl.ShowTableStyleRowStripes
    If Not tbl.DataBodyRange Is Nothing Then
        dAlturaLinha = tbl.DataBodyRange.Rows(1).RowHeight
        If dAlturaLinha < 12# Or dAlturaLinha > 40# Then dAlturaLinha = 15#
    End If
    On Error GoTo TratarErro

    ReDim vFormatos(1 To lTotalCols)
    CapturarFormatosColunas tbl, vFormatos

    ReDim vValidacoes(1 To lTotalCols)
    CapturarValidacoesTabela tbl, vValidacoes

    ReDim vCondDuplicados(1 To lTotalCols)
    CapturarCondicionaisDuplicados wsMonit, tbl, vCondDuplicados

    ' 8. Filtrar em memória: remover Situação = 'AT', 'CO' ou 'CA' e compactar apenas linhas com conteúdo real
    sEtapaAtual = "Filtragem em memória das linhas AT, CO e CA"
    lTotalInicial = tbl.ListRows.Count
    lLinhaFimAnterior = lLinhaHeader + lTotalInicial
    lMantidas = 0
    lRemovidas = 0

    If lTotalInicial > 0 And Not tbl.DataBodyRange Is Nothing Then
        If lTotalInicial = 1 And lTotalCols = 1 Then
            ReDim vDados(1 To 1, 1 To 1)
            vDados(1, 1) = tbl.DataBodyRange.Value2
        Else
            vDados = tbl.DataBodyRange.Value2
        End If

        ReDim idxMantidas(1 To lTotalInicial)

        For r = 1 To lTotalInicial
            sValSituacao = UCase$(LimparEspacosEspeciais(CStr(vDados(r, colSituacao) & "")))
            If sValSituacao = "AT" Or sValSituacao = "CO" Or sValSituacao = "CA" Then
                lRemovidas = lRemovidas + 1
            ElseIf Not LinhaTotalmenteVazia(vDados, r, lTotalCols) Then
                lMantidas = lMantidas + 1
                idxMantidas(lMantidas) = r
            End If
        Next r
    End If

    ' 9. Classificar exclusivamente o bloco de linhas com conteúdo (1..lMantidas) por 'D. abertura' crescente
    '    (mantendo todas as linhas com conteúdo juntas no topo, sem misturar com as linhas vazias abaixo)
    sEtapaAtual = "Classificação das linhas mantidas no topo"
    If lMantidas > 1 And colDataAbertura > 0 Then
        OrdenarIndicesPorDataAbertura vDados, idxMantidas, lMantidas, colDataAbertura, colDataSolicitacao
    End If

    ' 10. Montar matriz final: linhas 1..lMantidas contíguas no início + linhas vazias até o mínimo de 700
    sEtapaAtual = "Montagem da matriz de saída (mínimo de 700 linhas)"
    If lMantidas < MIN_LINHAS_TABELA Then
        lLinhasFinais = MIN_LINHAS_TABELA
    Else
        lLinhasFinais = lMantidas
    End If

    ReDim vSaida(1 To lLinhasFinais, 1 To lTotalCols)
    For k = 1 To lMantidas
        r = idxMantidas(k)
        For c = 1 To lTotalCols
            vCel = vDados(r, c)
            If VarType(vCel) = vbString Then
                If Len(LimparEspacosEspeciais(CStr(vCel))) = 0 Then
                    vSaida(k, c) = Empty
                ElseIf Len(vCel) > 1 And Left$(vCel, 1) = "0" And IsNumeric(vCel) Then
                    vSaida(k, c) = "'" & vCel
                Else
                    vSaida(k, c) = vCel
                End If
            Else
                vSaida(k, c) = vCel
            End If
        Next c
    Next k

    ' 11. Redimensionar a Tabela Estruturada e Gravar Dados em Bloco Único
    sEtapaAtual = "Remoção de tabelas órfãs e células mescladas conflitantes"
    RemoverTabelasSecundariasAbaixo wsMonit, tbl, lLinhaHeader

    Set rngNovoTamanho = wsMonit.Range(wsMonit.Cells(lLinhaHeader, lColInicial), _
                                       wsMonit.Cells(lLinhaHeader + lLinhasFinais, lColInicial + lTotalCols - 1))
    On Error Resume Next
    rngNovoTamanho.UnMerge
    On Error GoTo TratarErro

    sEtapaAtual = "Redimensionamento da tabela e gravação de dados"
    tbl.Resize rngNovoTamanho
    tbl.DataBodyRange.Value2 = vSaida

    ' 12. Limpar Resíduos Abaixo e à Direita da Tabela (via .Clear para preservar fórmulas externas)
    sEtapaAtual = "Limpeza de células residuais fora da tabela"
    lLinhaFimTabela = lLinhaHeader + lLinhasFinais
    lColFimTabela = lColInicial + lTotalCols - 1

    lUltimaLinhaPlan = wsMonit.Cells.SpecialCells(xlCellTypeLastCell).Row
    If lLinhaFimAnterior > lUltimaLinhaPlan Then lUltimaLinhaPlan = lLinhaFimAnterior

    If lUltimaLinhaPlan > lLinhaFimTabela Then
        wsMonit.Range(wsMonit.Cells(lLinhaFimTabela + 1, 1), _
                      wsMonit.Cells(lUltimaLinhaPlan, lColFimTabela)).Clear
    End If

    If wsMonit.ListObjects.Count = 1 Then
        lUltimaColPlan = wsMonit.Cells.SpecialCells(xlCellTypeLastCell).Column
        If lUltimaColPlan > lColFimTabela Then
            wsMonit.Range(wsMonit.Cells(1, lColFimTabela + 1), _
                          wsMonit.Cells(lUltimaLinhaPlan, lUltimaColPlan)).Clear
        End If
    End If

    ' 13. Reaplicar Estilo da Tabela, Formatos de Coluna, Formatação Condicional e Validações
    sEtapaAtual = "Restauração integral de formatações da tabela"
    On Error Resume Next
    tbl.TableStyle = sEstiloTabela
    tbl.ShowTableStyleRowStripes = bTirasLinhas
    tbl.ShowAutoFilter = True
    On Error GoTo TratarErro

    AplicarFormatosColunas tbl, vFormatos, dAlturaLinha

    sEtapaAtual = "Reaplicação das validações de dados"
    ReaplicarValidacoesTabela tbl, vValidacoes, wsMonit

    sEtapaAtual = "Restauração da formatação condicional"
    ReaplicarCondicionaisDuplicados wsMonit, tbl, vCondDuplicados

    ' 14. Limpar metadados de SortFields para não deslocar linhas vazias em reordenações futuras
    On Error Resume Next
    tbl.Sort.SortFields.Clear
    On Error GoTo TratarErro

    ' 15. Restaurar Ambiente e Posicionar Cursor em A3
    sEtapaAtual = "Finalização e posicionamento"
    Application.ScreenUpdating = True
    On Error Resume Next
    Application.Goto wsMonit.Range("A3"), True
    On Error GoTo TratarErro

    ' 16. Mensagem de Conclusão (modo interativo)
    If Not bSilencioso Then
        MsgBox "[OK] Limpeza e organização da aba 'Monitoramento' concluída com sucesso!" & vbNewLine & vbNewLine & _
               "- Tabela: '" & sNomeTabela & "' (" & lLinhasFinais & " linhas totais)" & vbNewLine & _
               "- Linhas removidas (AT, CO e CA): " & lRemovidas & vbNewLine & _
               "- Linhas com conteúdo agrupadas no início: " & lMantidas & vbNewLine & _
               "- Linhas em branco ao final: " & (lLinhasFinais - lMantidas) & vbNewLine & _
               "- Classificação: crescente por 'D. abertura'." & vbNewLine & _
               "- Todas as formatações e validações foram preservadas.", _
               vbInformation, MSG_TITULO
    End If

SairRotina:
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
' FUNÇÃO AUXILIAR: Remove espaços comuns, espaços não-quebráveis (Chr(160)), tabulações e quebras de linha
'-------------------------------------------------------------------------------------------------------------------
Private Function LimparEspacosEspeciais(ByVal sTexto As String) As String
    If Len(sTexto) = 0 Then
        LimparEspacosEspeciais = ""
        Exit Function
    End If
    sTexto = Replace(sTexto, Chr$(160), " ")
    sTexto = Replace(sTexto, vbTab, " ")
    sTexto = Replace(sTexto, vbCr, "")
    sTexto = Replace(sTexto, vbLf, "")
    LimparEspacosEspeciais = Trim$(sTexto)
End Function

'-------------------------------------------------------------------------------------------------------------------
' SUBROTINA AUXILIAR: Remove ListObjects secundários/órfãos (ex.: Tabela16 acidental abaixo dos dados) na aba
'                     Monitoramento que impediriam o tbl.Resize com Erro 1004
'-------------------------------------------------------------------------------------------------------------------
Private Sub RemoverTabelasSecundariasAbaixo(ByVal wsMonit As Worksheet, ByVal tblPrincipal As ListObject, _
                                            ByVal lLinhaHeaderPrincipal As Long)
    Dim i As Long
    Dim loExtra As ListObject
    Dim rngExtra As Range
    Dim bAbaixo As Boolean

    If wsMonit.ListObjects.Count <= 1 Then Exit Sub

    For i = wsMonit.ListObjects.Count To 1 Step -1
        Set loExtra = wsMonit.ListObjects(i)
        If Not (loExtra Is tblPrincipal) Then
            If StrComp(loExtra.Name, tblPrincipal.Name, vbTextCompare) <> 0 Then
                Set rngExtra = loExtra.Range
                bAbaixo = (rngExtra.Row > lLinhaHeaderPrincipal)
                On Error Resume Next
                loExtra.Unlist
                If bAbaixo And Not rngExtra Is Nothing Then
                    rngExtra.Clear
                End If
                On Error GoTo 0
            End If
        End If
    Next i
End Sub

'-------------------------------------------------------------------------------------------------------------------
' FUNÇÃO AUXILIAR: Verifica se uma linha da matriz está 100% vazia (ignorando espaços e Chr(160))
'-------------------------------------------------------------------------------------------------------------------
Private Function LinhaTotalmenteVazia(ByRef vDados As Variant, ByVal r As Long, ByVal lTotalCols As Long) As Boolean
    Dim c As Long
    For c = 1 To lTotalCols
        If Not IsEmpty(vDados(r, c)) Then
            If IsError(vDados(r, c)) Then
                LinhaTotalmenteVazia = False
                Exit Function
            ElseIf Len(LimparEspacosEspeciais(CStr(vDados(r, c)))) > 0 Then
                LinhaTotalmenteVazia = False
                Exit Function
            End If
        End If
    Next c
    LinhaTotalmenteVazia = True
End Function

'-------------------------------------------------------------------------------------------------------------------
' SUBROTINA AUXILIAR: Ordenação estável (Insertion Sort) de idx(1..n) mantendo todas as linhas com conteúdo juntas:
'   - Grupo 0: Linhas com 'D. abertura' preenchida (ordenadas da mais antiga para a mais recente)
'   - Grupo 1: Linhas com 'D. abertura' vazia, mas com 'D. solicitação' preenchida (ordenadas por 'D. solicitação')
'   - Grupo 2: Demais linhas com conteúdo sem data preenchida (mantendo ordem original logo após Grupo 1)
'-------------------------------------------------------------------------------------------------------------------
Private Sub OrdenarIndicesPorDataAbertura(ByRef vDados As Variant, ByRef idx() As Long, _
                                          ByVal n As Long, ByVal colDataAbert As Long, _
                                          ByVal colDataSolic As Long)
    Dim grupos() As Long
    Dim valoresNum() As Double
    Dim valoresSec() As Double
    Dim valoresTxt() As String
    Dim i As Long
    Dim j As Long
    Dim rOrig As Long

    Dim idxChave As Long
    Dim grpChave As Long
    Dim numChave As Double
    Dim secChave As Double
    Dim txtChave As String

    Dim grpSolic As Long
    Dim numSolic As Double
    Dim txtSolic As String

    ReDim grupos(1 To n)
    ReDim valoresNum(1 To n)
    ReDim valoresSec(1 To n)
    ReDim valoresTxt(1 To n)

    For i = 1 To n
        rOrig = idx(i)
        ExtrairChaveOrdenacaoData vDados(rOrig, colDataAbert), grupos(i), valoresNum(i), valoresTxt(i)

        If colDataSolic > 0 Then
            ExtrairChaveOrdenacaoData vDados(rOrig, colDataSolic), grpSolic, numSolic, txtSolic
            If grpSolic = 0 Then
                valoresSec(i) = numSolic
            Else
                valoresSec(i) = CDbl(rOrig)
            End If
        Else
            valoresSec(i) = CDbl(rOrig)
        End If

        ' Se D. abertura estiver vazia, usa D. solicitação no Grupo 1 para manter todas as linhas juntas e ordenadas
        If grupos(i) = 2 And colDataSolic > 0 Then
            If grpSolic = 0 Then
                grupos(i) = 1
                valoresNum(i) = numSolic
            End If
        End If
    Next i

    For i = 2 To n
        idxChave = idx(i)
        grpChave = grupos(i)
        numChave = valoresNum(i)
        secChave = valoresSec(i)
        txtChave = valoresTxt(i)
        j = i - 1

        Do While j >= 1
            If CompararChavesData(grupos(j), valoresNum(j), valoresSec(j), valoresTxt(j), _
                                  grpChave, numChave, secChave, txtChave) > 0 Then
                idx(j + 1) = idx(j)
                grupos(j + 1) = grupos(j)
                valoresNum(j + 1) = valoresNum(j)
                valoresSec(j + 1) = valoresSec(j)
                valoresTxt(j + 1) = valoresTxt(j)
                j = j - 1
            Else
                Exit Do
            End If
        Loop

        idx(j + 1) = idxChave
        grupos(j + 1) = grpChave
        valoresNum(j + 1) = numChave
        valoresSec(j + 1) = secChave
        valoresTxt(j + 1) = txtChave
    Next i
End Sub

'-------------------------------------------------------------------------------------------------------------------
' SUBROTINA AUXILIAR: Extrai chave tipada para ordenação cronológica (0=Data válida >0, 2=Vazio, 3=Texto)
'-------------------------------------------------------------------------------------------------------------------
Private Sub ExtrairChaveOrdenacaoData(ByVal vValor As Variant, ByRef lGrupo As Long, _
                                      ByRef dNum As Double, ByRef sTxt As String)
    Dim sLimpo As String
    dNum = 0#
    sTxt = ""

    If IsEmpty(vValor) Or IsError(vValor) Then
        lGrupo = 2
        Exit Sub
    End If

    If VarType(vValor) = vbDate Or VarType(vValor) = vbDouble Or _
       VarType(vValor) = vbSingle Or VarType(vValor) = vbLong Or _
       VarType(vValor) = vbInteger Or VarType(vValor) = vbCurrency Or VarType(vValor) = vbDecimal Then
        If CDbl(vValor) > 0# Then
            lGrupo = 0
            dNum = CDbl(vValor)
        Else
            lGrupo = 2
        End If
        Exit Sub
    End If

    sLimpo = LimparEspacosEspeciais(CStr(vValor))
    If Len(sLimpo) = 0 Then
        lGrupo = 2
    ElseIf IsNumeric(sLimpo) Then
        If CDbl(sLimpo) > 0# Then
            lGrupo = 0
            dNum = CDbl(sLimpo)
        Else
            lGrupo = 2
        End If
    ElseIf IsDate(sLimpo) Then
        lGrupo = 0
        dNum = CDbl(CDate(sLimpo))
    Else
        lGrupo = 3
        sTxt = UCase$(sLimpo)
    End If
End Sub

'-------------------------------------------------------------------------------------------------------------------
' FUNÇÃO AUXILIAR: Compara duas chaves extraídas (-1 se A < B, 0 se A = B, 1 se A > B)
'-------------------------------------------------------------------------------------------------------------------
Private Function CompararChavesData(ByVal grpA As Long, ByVal numA As Double, ByVal secA As Double, ByVal txtA As String, _
                                    ByVal grpB As Long, ByVal numB As Double, ByVal secB As Double, ByVal txtB As String) As Long
    If grpA < grpB Then
        CompararChavesData = -1
    ElseIf grpA > grpB Then
        CompararChavesData = 1
    ElseIf grpA = 0 Or grpA = 1 Then
        If numA < numB Then
            CompararChavesData = -1
        ElseIf numA > numB Then
            CompararChavesData = 1
        ElseIf secA < secB Then
            CompararChavesData = -1
        ElseIf secA > secB Then
            CompararChavesData = 1
        Else
            CompararChavesData = 0
        End If
    ElseIf grpA = 3 Then
        CompararChavesData = StrComp(txtA, txtB, vbTextCompare)
    Else
        CompararChavesData = 0
    End If
End Function

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
' FUNÇÃO AUXILIAR: Localiza o índice de uma coluna pelo nome (exato primeiro, depois parcial, com fallback)
'-------------------------------------------------------------------------------------------------------------------
Private Function ObterIndiceColuna(ByVal tbl As ListObject, ByVal sTrechoNome As String, ByVal lIndiceFallback As Long) As Long
    Dim c As Long
    Dim sNomeLimpo As String
    Dim sTrechoLimpo As String

    sTrechoLimpo = UCase$(LimparEspacosEspeciais(sTrechoNome))

    ' 1ª passagem: correspondência exata
    For c = 1 To tbl.ListColumns.Count
        sNomeLimpo = UCase$(LimparEspacosEspeciais(tbl.ListColumns(c).Name))
        If sNomeLimpo = sTrechoLimpo Then
            ObterIndiceColuna = c
            Exit Function
        End If
    Next c

    ' 2ª passagem: correspondência parcial
    For c = 1 To tbl.ListColumns.Count
        sNomeLimpo = UCase$(LimparEspacosEspeciais(tbl.ListColumns(c).Name))
        If InStr(1, sNomeLimpo, sTrechoLimpo, vbTextCompare) > 0 Then
            ObterIndiceColuna = c
            Exit Function
        End If
    Next c

    If lIndiceFallback >= 0 And lIndiceFallback <= tbl.ListColumns.Count Then
        ObterIndiceColuna = lIndiceFallback
    Else
        ObterIndiceColuna = 0
    End If
End Function

'-------------------------------------------------------------------------------------------------------------------
' SUBROTINA AUXILIAR: Captura formato de número/data, alinhamento, quebra de linha e fonte de cada coluna
'-------------------------------------------------------------------------------------------------------------------
Private Sub CapturarFormatosColunas(ByVal tbl As ListObject, ByRef vFormatos() As InfoFormatoColuna)
    Dim c As Long
    Dim cellRef As Range

    For c = 1 To tbl.ListColumns.Count
        vFormatos(c).FormatoLocal = "Geral"
        vFormatos(c).AlinhamentoH = xlGeneral
        vFormatos(c).AlinhamentoV = xlCenter
        vFormatos(c).QuebrarTexto = (c = 26)
        vFormatos(c).NomeFonte = "Calibri"
        vFormatos(c).TamanhoFonte = IIf(c = 26, 9#, 11#)
        vFormatos(c).Negrito = False

        If Not tbl.DataBodyRange Is Nothing Then
            Set cellRef = tbl.DataBodyRange.Cells(1, c)
            On Error Resume Next
            vFormatos(c).FormatoLocal = CStr(cellRef.NumberFormatLocal)
            vFormatos(c).AlinhamentoH = CLng(cellRef.HorizontalAlignment)
            vFormatos(c).AlinhamentoV = CLng(cellRef.VerticalAlignment)
            vFormatos(c).QuebrarTexto = CBool(cellRef.WrapText)
            If Len(CStr(cellRef.Font.Name)) > 0 Then vFormatos(c).NomeFonte = CStr(cellRef.Font.Name)
            If CDbl(cellRef.Font.Size) > 0 Then vFormatos(c).TamanhoFonte = CDbl(cellRef.Font.Size)
            vFormatos(c).Negrito = CBool(cellRef.Font.Bold)
            On Error GoTo 0
        End If
    Next c
End Sub

'-------------------------------------------------------------------------------------------------------------------
' SUBROTINA AUXILIAR: Aplica formatos capturados uniformemente em todas as linhas da tabela (>= 700 linhas)
'-------------------------------------------------------------------------------------------------------------------
Private Sub AplicarFormatosColunas(ByVal tbl As ListObject, ByRef vFormatos() As InfoFormatoColuna, ByVal dAltura As Double)
    Dim c As Long
    Dim rngCol As Range

    If tbl.DataBodyRange Is Nothing Then Exit Sub

    On Error Resume Next
    With tbl.DataBodyRange
        .Interior.ColorIndex = xlColorIndexNone
        .Font.ColorIndex = xlColorIndexAutomatic
        .RowHeight = dAltura
    End With

    For c = 1 To tbl.ListColumns.Count
        Set rngCol = tbl.ListColumns(c).DataBodyRange
        If Not rngCol Is Nothing Then
            If Len(vFormatos(c).FormatoLocal) > 0 Then
                rngCol.NumberFormatLocal = vFormatos(c).FormatoLocal
            End If
            rngCol.HorizontalAlignment = vFormatos(c).AlinhamentoH
            rngCol.VerticalAlignment = vFormatos(c).AlinhamentoV
            rngCol.WrapText = vFormatos(c).QuebrarTexto
            rngCol.Font.Name = vFormatos(c).NomeFonte
            rngCol.Font.Size = vFormatos(c).TamanhoFonte
            rngCol.Font.Bold = vFormatos(c).Negrito
        End If
    Next c
    On Error GoTo 0
End Sub

'-------------------------------------------------------------------------------------------------------------------
' SUBROTINA AUXILIAR: Captura regras de destaque de valores duplicados existentes na tabela
'-------------------------------------------------------------------------------------------------------------------
Private Sub CapturarCondicionaisDuplicados(ByVal wsMonit As Worksheet, ByVal tbl As ListObject, _
                                           ByRef vCond() As InfoCondicionalDuplicados)
    Dim i As Long
    Dim c As Long
    Dim fc As Object
    Dim rngInter As Range

    For c = 1 To tbl.ListColumns.Count
        vCond(c).Ativo = False
        vCond(c).CorFundo = 13551615 ' RGB(255, 199, 206) padrão Excel
        vCond(c).CorFonte = 393372   ' RGB(156, 0, 6) padrão Excel
    Next c

    If tbl.DataBodyRange Is Nothing Then Exit Sub

    On Error Resume Next
    For i = 1 To wsMonit.Cells.FormatConditions.Count
        Set fc = wsMonit.Cells.FormatConditions.Item(i)
        If Not fc Is Nothing Then
            If fc.Type = xlUniqueValues Then
                If fc.DupeUnique = xlDuplicate Then
                    For c = 1 To tbl.ListColumns.Count
                        Set rngInter = Nothing
                        Set rngInter = Application.Intersect(fc.AppliesTo, tbl.ListColumns(c).DataBodyRange)
                        If Not rngInter Is Nothing Then
                            vCond(c).Ativo = True
                            If CLng(fc.Interior.Color) > 0 Then vCond(c).CorFundo = CLng(fc.Interior.Color)
                            If CLng(fc.Font.Color) >= 0 Then vCond(c).CorFonte = CLng(fc.Font.Color)
                        End If
                    Next c
                End If
            End If
        End If
    Next i
    On Error GoTo 0
End Sub

'-------------------------------------------------------------------------------------------------------------------
' SUBROTINA AUXILIAR: Desfragmenta e reaplica regras de duplicidade em todo o DataBodyRange de cada coluna
'-------------------------------------------------------------------------------------------------------------------
Private Sub ReaplicarCondicionaisDuplicados(ByVal wsMonit As Worksheet, ByVal tbl As ListObject, _
                                            ByRef vCond() As InfoCondicionalDuplicados)
    Dim i As Long
    Dim c As Long
    Dim bAlgumaAtiva As Boolean
    Dim fcObj As Object
    Dim uvRule As UniqueValues
    Dim rngCol As Range

    If tbl.DataBodyRange Is Nothing Then Exit Sub

    bAlgumaAtiva = False
    For c = 1 To tbl.ListColumns.Count
        If vCond(c).Ativo Then
            bAlgumaAtiva = True
            Exit For
        End If
    Next c
    If Not bAlgumaAtiva Then Exit Sub

    On Error Resume Next
    For i = wsMonit.Cells.FormatConditions.Count To 1 Step -1
        Set fcObj = wsMonit.Cells.FormatConditions.Item(i)
        If Not fcObj Is Nothing Then
            If fcObj.Type = xlUniqueValues Then
                fcObj.Delete
            End If
        End If
    Next i

    For c = 1 To tbl.ListColumns.Count
        If vCond(c).Ativo Then
            Set rngCol = tbl.ListColumns(c).DataBodyRange
            If Not rngCol Is Nothing Then
                Set uvRule = rngCol.FormatConditions.AddUniqueValues
                uvRule.DupeUnique = xlDuplicate
                uvRule.Interior.Color = vCond(c).CorFundo
                uvRule.Font.Color = vCond(c).CorFonte
                uvRule.StopIfTrue = False
            End If
        End If
    Next c
    On Error GoTo 0
End Sub

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
    Dim lTipo As Long

    For c = 1 To tbl.ListColumns.Count
        vValidacoes(c).TemValidacao = False
        Set rngCol = tbl.ListColumns(c).DataBodyRange

        If Not rngCol Is Nothing Then
            lMaxLinhasTeste = Application.WorksheetFunction.Min(rngCol.Rows.Count, 5)
            For r = 1 To lMaxLinhasTeste
                Set cellTest = rngCol.Cells(r, 1)
                lTipo = 0
                On Error Resume Next
                Set vObj = cellTest.Validation
                If Not vObj Is Nothing Then lTipo = vObj.Type
                If Err.Number = 0 And lTipo <> 0 Then
                    vValidacoes(c).TemValidacao = True
                    vValidacoes(c).Tipo = lTipo
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
                Err.Clear
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
        Case 25 ' Coluna Y: OS disponibilizada?
            ObterFormulaFallbackPorIndice = "=TabelaA!$AW$2:$AW$3"
        Case Else
            ObterFormulaFallbackPorIndice = ""
    End Select
End Function
