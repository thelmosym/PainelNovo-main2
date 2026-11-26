Attribute VB_Name = "modGerarArquivos"
Option Explicit

'====================================================================================================
' MÓDULO: modGerarArquivos
' OBJETIVO:
'   1. Gera os 4 cadernos regionais limpos (PA-LT2, IRON-LT1-RJ, IRON-LT1-SP, IRON-LT1-ES),
'      cada um com 4 abas (MC, ARM, DADOS, FRETE), preservando fórmulas ativas na MC e gravando
'      links em Painel!F14:F17.
'   2. Gera junto 2 planilhas consolidadas por contrato-base (IRON-LT1 e PA-LT2, desconsiderando
'      as siglas após o 2º hífen do contrato), cada uma com as 4 abas (Painel T2M até a coluna U /
'      Observações, ARM, DADOS e FRETE), com nomenclatura:
'      "Painel de Controle Memoria Petrobras<CONTRATO><ddmmyyyyhhnnss>.xlsx"
'      e links gravados em Painel!F19:F20 (preservando F18 do Demonstrativo FDM).
'
' CODIFICAÇÃO: Windows-1252 (CP1252 / ANSI) com quebras CRLF para compatibilidade VBE.
'====================================================================================================

Private Const MSG_TITULO As String = "Geração de Arquivos por Contrato"
Private Const NOME_PLANILHA_LINKS As String = "Painel"
Private Const COLUNA_CELULA_LINK As String = "F"
Private Const LINHA_INICIAL_LINK As Long = 14
Private Const LINHA_INICIAL_LINK_CONSOLIDADO As Long = 19
Private Const NOME_SUBPASTA_DESTINO As String = "MEDIÇÃO"
Private Const PREFIXO_ARQUIVO_CONSOLIDADO As String = "Painel de Controle Memoria Petrobras"

' Paleta de cores corporativa para status dos links na aba Painel (RGB)
Private Const COR_STATUS_AMARELO     As Long = 65535        ' RGB(255, 255, 0)   - Amarelo (processando)
Private Const COR_STATUS_VERDE_CLARO As Long = 5296274      ' RGB(146, 208, 80)  - Verde suave (sucesso)
Private Const COR_STATUS_VERMELHO    As Long = 255          ' RGB(255, 0, 0)     - Vermelho (erro / falha)

' GUID do rótulo de confidencialidade "Pública" do tenant da Petrobras
Private Const PUBLICO_LABEL_ID As String = "140b9f7d-8e3a-482f-9702-4b7ffc40985a"
Private Const MSO_ASSIGNMENT_METHOD_STANDARD As Long = 1

Private Function ListaTipos() As Variant
    ListaTipos = Array("MC", "ARM", "DADOS", "FRETE")
End Function

Private Function ListaContratos() As Variant
    ListaContratos = Array("PA-LT2", "IRON-LT1-RJ", "IRON-LT1-SP", "IRON-LT1-ES")
End Function

Private Function ListaNumerosContrato() As Variant
    ListaNumerosContrato = Array("4600687006", "4600686987", "4600686987", "4600686987")
End Function

Private Function ListaSufixosEstado() As Variant
    ListaSufixosEstado = Array("BA", "", "", "")
End Function

Private Function ListaContratosConsolidados() As Variant
    ListaContratosConsolidados = Array("IRON-LT1", "PA-LT2")
End Function

'----------------------------------------------------------------------------------------------------
' PONTOS DE ENTRADA PÚBLICOS
'----------------------------------------------------------------------------------------------------
Public Sub GerarArquivosPorContrato()
    ExecutarGeracaoArquivos False
End Sub

Public Sub GerarArquivosPorContrato_Silencioso()
    ExecutarGeracaoArquivos True
End Sub

'----------------------------------------------------------------------------------------------------
' ORQUESTRADOR PRINCIPAL DE GERAÇÃO (4 CADERNOS REGIONAIS + 2 CONSOLIDADOS POR CONTRATO)
'----------------------------------------------------------------------------------------------------
Private Sub ExecutarGeracaoArquivos(ByVal bSilencioso As Boolean)
    Dim wbOrigem            As Workbook
    Dim wsLinks             As Worksheet
    Dim sPastaDestino       As String
    Dim vContratos          As Variant
    Dim vTipos              As Variant
    Dim vNumeros            As Variant
    Dim vSufixos            As Variant
    Dim vConsolidados       As Variant
    Dim iTotalEsperado      As Long
    Dim iContrato           As Long
    Dim iCons               As Long
    Dim iArquivosOK         As Long
    Dim sResumoFinal        As String
    Dim sTimestampLote      As String

    Dim appCalc             As XlCalculation
    Dim bEvents             As Boolean
    Dim bScreen             As Boolean
    Dim bAlerts             As Boolean

    Set wbOrigem = ThisWorkbook

    If Len(wbOrigem.Path) = 0 Then
        If Not bSilencioso Then
            MsgBox "Procedimento interrompido." & vbNewLine & vbNewLine & _
                   "Este arquivo ainda não foi salvo em disco, portanto não é possível " & _
                   "determinar a pasta de destino automaticamente." & vbNewLine & _
                   "Salve este arquivo (Ctrl+S) e tente novamente.", _
                   vbExclamation, MSG_TITULO
        End If
        Exit Sub
    End If

    vContratos = ListaContratos()
    vTipos = ListaTipos()
    vNumeros = ListaNumerosContrato()
    vSufixos = ListaSufixosEstado()
    vConsolidados = ListaContratosConsolidados()
    iTotalEsperado = (UBound(vContratos) - LBound(vContratos) + 1) + _
                     (UBound(vConsolidados) - LBound(vConsolidados) + 1)

    sPastaDestino = ObterOuCriarSubpasta(wbOrigem.Path, NOME_SUBPASTA_DESTINO)
    If Len(sPastaDestino) = 0 Then
        If Not bSilencioso Then
            MsgBox "Procedimento interrompido. Não foi possível criar ou acessar " & _
                   "a pasta '" & NOME_SUBPASTA_DESTINO & "'.", vbCritical, MSG_TITULO
        End If
        Exit Sub
    End If

    Set wsLinks = ObterOuCriarPlanilha(wbOrigem, NOME_PLANILHA_LINKS)

    If Not bSilencioso Then
        If MsgBox("Serão gerados " & iTotalEsperado & " arquivos na pasta:" & vbNewLine & _
                  sPastaDestino & vbNewLine & vbNewLine & _
                  "1) 4 Cadernos Regionais (" & Join(vContratos, ", ") & ") -> Abas: " & Join(vTipos, ", ") & vbNewLine & _
                  "2) 2 Planilhas Consolidadas (" & Join(vConsolidados, ", ") & ") -> Abas: Painel T2M (até col. U), ARM, DADOS, FRETE" & vbNewLine & vbNewLine & _
                  "Deseja iniciar a geração agora?", _
                  vbYesNo + vbQuestion, MSG_TITULO) = vbNo Then
            MsgBox "Procedimento cancelado.", vbInformation, MSG_TITULO
            Exit Sub
        End If
    End If

    '-- Inicializa células F14:F17 (regionais) e F19:F20 (consolidados, preservando F18) em amarelo
    Dim iInicial As Long
    For iInicial = LBound(vContratos) To UBound(vContratos)
        MarcarCelulaEmAndamento wsLinks, ObterLinhaCelulaLink(iInicial), CStr(vContratos(iInicial))
    Next iInicial

    For iCons = LBound(vConsolidados) To UBound(vConsolidados)
        CopiarEstiloCelulaModelo wsLinks, LINHA_INICIAL_LINK + 3, ObterLinhaCelulaLinkConsolidado(iCons)
        MarcarCelulaEmAndamento wsLinks, ObterLinhaCelulaLinkConsolidado(iCons), "Painel " & CStr(vConsolidados(iCons))
    Next iCons
    DoEvents

    '-- Captura e isola o estado do ambiente do Excel
    appCalc = Application.Calculation
    bEvents = Application.EnableEvents
    bScreen = Application.ScreenUpdating
    bAlerts = Application.DisplayAlerts

    On Error GoTo TratarErro
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    iArquivosOK = 0
    sResumoFinal = ""
    sTimestampLote = Format(Now, "ddmmyyyyhhnnss")

    '-- PARTE 1: Gera os 4 cadernos regionais por contrato (MC, ARM, DADOS, FRETE)
    For iContrato = LBound(vContratos) To UBound(vContratos)
        Dim sContrato As String, sNumero As String, sSufixo As String
        Dim sCaminhoGerado As String

        sContrato = CStr(vContratos(iContrato))
        sNumero = CStr(vNumeros(iContrato))
        sSufixo = CStr(vSufixos(iContrato))

        Application.StatusBar = "Gerando arquivo " & (iContrato + 1) & " de " & _
                                iTotalEsperado & ": " & sContrato & "..."

        sCaminhoGerado = GerarArquivoDoContrato(wbOrigem, sContrato, sNumero, sSufixo, vTipos, sPastaDestino)

        If Len(sCaminhoGerado) > 0 Then
            iArquivosOK = iArquivosOK + 1
            sResumoFinal = sResumoFinal & "[OK] " & sContrato & " -> " & ExtrairNomeArquivo(sCaminhoGerado) & vbNewLine
            EscreverLinkArquivoEmLinha wsLinks, ObterLinhaCelulaLink(iContrato), sCaminhoGerado
        Else
            sResumoFinal = sResumoFinal & "[FALHA] " & sContrato & " (Falha na geração)" & vbNewLine
            EscreverErroLinkEmLinha wsLinks, ObterLinhaCelulaLink(iContrato), sContrato
        End If
    Next iContrato

    '-- PARTE 2: Gera as 2 planilhas consolidadas por contrato-base (IRON-LT1 e PA-LT2)
    For iCons = LBound(vConsolidados) To UBound(vConsolidados)
        Dim sContratoBase As String
        Dim sCaminhoCons As String
        Dim iOrdemAtual As Long

        sContratoBase = CStr(vConsolidados(iCons))
        iOrdemAtual = (UBound(vContratos) - LBound(vContratos) + 1) + (iCons - LBound(vConsolidados) + 1)

        Application.StatusBar = "Gerando arquivo " & iOrdemAtual & " de " & _
                                iTotalEsperado & ": Consolidado " & sContratoBase & "..."

        sCaminhoCons = GerarArquivoConsolidadoContrato(wbOrigem, sContratoBase, sPastaDestino, sTimestampLote)

        If Len(sCaminhoCons) > 0 Then
            iArquivosOK = iArquivosOK + 1
            sResumoFinal = sResumoFinal & "[OK] Consolidado " & sContratoBase & " -> " & ExtrairNomeArquivo(sCaminhoCons) & vbNewLine
            EscreverLinkArquivoEmLinha wsLinks, ObterLinhaCelulaLinkConsolidado(iCons), sCaminhoCons
        Else
            sResumoFinal = sResumoFinal & "[FALHA] Consolidado " & sContratoBase & " (Falha na geração)" & vbNewLine
            EscreverErroLinkEmLinha wsLinks, ObterLinhaCelulaLinkConsolidado(iCons), "Consolidado " & sContratoBase
        End If
    Next iCons

SairRotina:
    Application.Calculation = appCalc
    Application.EnableEvents = bEvents
    Application.DisplayAlerts = bAlerts
    Application.ScreenUpdating = bScreen
    Application.StatusBar = False

    '-- Garante que qualquer célula de link que não tenha gerado hyperlink termine sinalizada em vermelho
    On Error Resume Next
    If Not wsLinks Is Nothing And iArquivosOK < iTotalEsperado Then
        Dim iFim As Long
        For iFim = LBound(vContratos) To UBound(vContratos)
            SinalizarVermelhoSeSemLink wsLinks, ObterLinhaCelulaLink(iFim)
        Next iFim
        For iFim = LBound(vConsolidados) To UBound(vConsolidados)
            SinalizarVermelhoSeSemLink wsLinks, ObterLinhaCelulaLinkConsolidado(iFim)
        Next iFim
    End If
    On Error GoTo 0

    If Not bSilencioso And iArquivosOK > 0 Then
        MsgBox "Processamento concluído com sucesso." & vbNewLine & vbNewLine & _
               "Arquivos gerados: " & iArquivosOK & " de " & iTotalEsperado & vbNewLine & vbNewLine & _
               sResumoFinal & vbNewLine & _
               "Links regionais em: " & wsLinks.Name & "!F14:F17" & vbNewLine & _
               "Links consolidados em: " & wsLinks.Name & "!F19:F20", _
               vbInformation, MSG_TITULO
    End If
    Exit Sub

TratarErro:
    On Error Resume Next
    If Not wsLinks Is Nothing Then
        Dim iErrIdx As Long
        For iErrIdx = LBound(vContratos) To UBound(vContratos)
            SinalizarErroInesperadoCelula wsLinks, ObterLinhaCelulaLink(iErrIdx), CStr(vContratos(iErrIdx))
        Next iErrIdx
        For iErrIdx = LBound(vConsolidados) To UBound(vConsolidados)
            SinalizarErroInesperadoCelula wsLinks, ObterLinhaCelulaLinkConsolidado(iErrIdx), CStr(vConsolidados(iErrIdx))
        Next iErrIdx
    End If
    On Error GoTo 0

    If Not bSilencioso Then
        MsgBox "Erro inesperado durante a geração dos arquivos:" & vbNewLine & _
               Err.Number & " - " & Err.Description, vbCritical, MSG_TITULO
    End If
    Resume SairRotina
End Sub

'----------------------------------------------------------------------------------------------------
' NOVA FUNÇÃO: GERA PLANILHA CONSOLIDADA POR CONTRATO-BASE (IRON-LT1 OU PA-LT2)
' Abas geradas: Painel T2M (até coluna U / Observações), ARM, DADOS e FRETE
' Filtro: desconsidera siglas após o 2º '-' do contrato (ex.: IRON-LT1-RJ -> IRON-LT1; PA-LT2-BA -> PA-LT2)
' Nome: "Painel de Controle Memoria Petrobras<CONTRATO><ddmmyyyyhhnnss>.xlsx"
'----------------------------------------------------------------------------------------------------
Private Function GerarArquivoConsolidadoContrato(ByVal wbOrigem As Workbook, _
                                                 ByVal sContratoBase As String, _
                                                 ByVal sPastaDestino As String, _
                                                 ByVal sTimestamp As String) As String
    Dim wbDestino        As Workbook
    Dim wsOrigT2M        As Worksheet
    Dim wsOrigARM        As Worksheet
    Dim wsOrigDADOS      As Worksheet
    Dim wsOrigFRETE      As Worksheet
    Dim wsDest           As Worksheet
    Dim lColLimiteT2M    As Long
    Dim iAbasOK          As Long
    Dim sNomeArquivo     As String
    Dim sCaminhoCompleto As String
    Dim sCarimbo         As String

    On Error GoTo TratarErroLocal
    GerarArquivoConsolidadoContrato = ""

    Set wbDestino = Workbooks.Add(xlWBATWorksheet)
    Do While wbDestino.Worksheets.Count > 1
        wbDestino.Worksheets(wbDestino.Worksheets.Count).Delete
    Loop

    iAbasOK = 0

    '-- 1. Aba 'Painel T2M' (limitada até a coluna U / Observações)
    Set wsOrigT2M = ObterPlanilha(wbOrigem, "Painel T2M")
    If wsOrigT2M Is Nothing Then Set wsOrigT2M = ObterPlanilha(wbOrigem, "Painel_T2M")

    Set wsDest = wbDestino.Worksheets(1)
    wsDest.Name = "Painel T2M"

    If Not wsOrigT2M Is Nothing Then
        lColLimiteT2M = ObterColunaLimiteObservacoesT2M(wsOrigT2M)
        ExportarAbaFiltradaPorContratoBase wsOrigT2M, wsDest, "Painel_T2M", sContratoBase, lColLimiteT2M
        iAbasOK = iAbasOK + 1
    Else
        wsDest.Range("A1").Value = "Aba de origem 'Painel T2M' não encontrada."
    End If

    '-- 2. Aba 'ARM'
    Set wsOrigARM = ObterPlanilha(wbOrigem, "ARM")
    Set wsDest = wbDestino.Worksheets.Add(After:=wbDestino.Worksheets(wbDestino.Worksheets.Count))
    wsDest.Name = "ARM"

    If Not wsOrigARM Is Nothing Then
        ExportarAbaFiltradaPorContratoBase wsOrigARM, wsDest, "ARM", sContratoBase, 0
        iAbasOK = iAbasOK + 1
    Else
        wsDest.Range("A1").Value = "Aba de origem 'ARM' não encontrada."
    End If

    '-- 3. Aba 'DADOS'
    Set wsOrigDADOS = ObterPlanilha(wbOrigem, "DADOS")
    Set wsDest = wbDestino.Worksheets.Add(After:=wbDestino.Worksheets(wbDestino.Worksheets.Count))
    wsDest.Name = "DADOS"

    If Not wsOrigDADOS Is Nothing Then
        ExportarAbaFiltradaPorContratoBase wsOrigDADOS, wsDest, "DADOS", sContratoBase, 0
        iAbasOK = iAbasOK + 1
    Else
        wsDest.Range("A1").Value = "Aba de origem 'DADOS' não encontrada."
    End If

    '-- 4. Aba 'FRETE'
    Set wsOrigFRETE = ObterPlanilha(wbOrigem, "FRETE")
    If wsOrigFRETE Is Nothing Then Set wsOrigFRETE = ObterPlanilha(wbOrigem, "Frete")
    Set wsDest = wbDestino.Worksheets.Add(After:=wbDestino.Worksheets(wbDestino.Worksheets.Count))
    wsDest.Name = "FRETE"

    If Not wsOrigFRETE Is Nothing Then
        ExportarAbaFiltradaPorContratoBase wsOrigFRETE, wsDest, "FRETE", sContratoBase, 0
        iAbasOK = iAbasOK + 1
    Else
        wsDest.Range("A1").Value = "Aba de origem 'FRETE' não encontrada."
    End If

    If iAbasOK = 0 Then
        wbDestino.Close SaveChanges:=False
        Exit Function
    End If

    ' Posiciona na primeira aba para apresentação limpa ao abrir
    On Error Resume Next
    wbDestino.Worksheets(1).Activate
    On Error GoTo TratarErroLocal

    AplicarPerfilPublico wbDestino

    If Len(sTimestamp) > 0 Then
        sCarimbo = sTimestamp
    Else
        sCarimbo = Format(Now, "ddmmyyyyhhnnss")
    End If

    ' Nome exato: "Painel de Controle Memoria Petrobras" + Contrato + datahora sem espaços ou símbolos
    sNomeArquivo = PREFIXO_ARQUIVO_CONSOLIDADO & sContratoBase & sCarimbo & ".xlsx"
    sCaminhoCompleto = sPastaDestino & Application.PathSeparator & sNomeArquivo

    wbDestino.SaveAs Filename:=sCaminhoCompleto, FileFormat:=xlOpenXMLWorkbook
    wbDestino.Close SaveChanges:=False

    GerarArquivoConsolidadoContrato = sCaminhoCompleto
    Exit Function

TratarErroLocal:
    On Error Resume Next
    If Not wbDestino Is Nothing Then wbDestino.Close SaveChanges:=False
    GerarArquivoConsolidadoContrato = ""
End Function

'----------------------------------------------------------------------------------------------------
' EXPORTAÇÃO EM MEMÓRIA DE ABA CONSOLIDADA FILTRADA PELO CONTRATO-BASE (ATÉ O 2º HÍFEN)
'----------------------------------------------------------------------------------------------------
Private Sub ExportarAbaFiltradaPorContratoBase(ByVal wsOrigem As Worksheet, _
                                               ByVal wsDestino As Worksheet, _
                                               ByVal sNomeTabelaDest As String, _
                                               ByVal sContratoBaseAlvo As String, _
                                               ByVal lMaxColunasExportar As Long)
    Dim tblOrig             As ListObject
    Dim tblDest             As ListObject
    Dim rngTabelaOrig       As Range
    Dim rngHeaderOrig       As Range
    Dim rngFormatoLinhaOrig As Range
    Dim rngTabelaDest       As Range
    Dim vOrig               As Variant
    Dim vOut()              As Variant

    Dim lTotalLinhasOrig    As Long
    Dim lTotalColsOrig      As Long
    Dim lColsExport         As Long
    Dim colFiltroContrato   As Long
    Dim lQtdFiltradas       As Long
    Dim lLinhasTabelaDest   As Long

    Dim r As Long
    Dim k As Long
    Dim c As Long
    Dim sAlvoUpper As String
    Dim vCel As Variant

    On Error GoTo TratarErroLocal

    If wsOrigem.ListObjects.Count > 0 Then
        Set tblOrig = wsOrigem.ListObjects(1)
        Set rngTabelaOrig = tblOrig.Range
    Else
        Set rngTabelaOrig = ObterAreaRealDados(wsOrigem)
    End If

    If rngTabelaOrig Is Nothing Then Exit Sub

    lTotalLinhasOrig = rngTabelaOrig.Rows.Count
    lTotalColsOrig = rngTabelaOrig.Columns.Count
    If lTotalColsOrig < 1 Then Exit Sub

    If lMaxColunasExportar > 0 And lMaxColunasExportar < lTotalColsOrig Then
        lColsExport = lMaxColunasExportar
    Else
        lColsExport = lTotalColsOrig
    End If

    ' Identifica a coluna oficial de contrato na tabela de origem (Contrato.1 > Contrato-Reg. > Contrato)
    colFiltroContrato = LocalizarColunaContratoConsolidado(rngTabelaOrig)
    If colFiltroContrato < 1 Then colFiltroContrato = 1

    ' 1. Copia cabeçalho (valores, formatos e larguras de colunas) até lColsExport
    Set rngHeaderOrig = rngTabelaOrig.Rows(1).Resize(1, lColsExport)
    rngHeaderOrig.Copy
    wsDestino.Range("A1").PasteSpecial Paste:=xlPasteValues
    wsDestino.Range("A1").PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False
    ReplicarLargurasColunas rngHeaderOrig, wsDestino

    If lTotalLinhasOrig >= 2 Then
        Set rngFormatoLinhaOrig = rngTabelaOrig.Rows(2).Resize(1, lColsExport)
    End If

    ' 2. Filtra linhas em memória RAM comparando o contrato truncado antes do 2º '-'
    sAlvoUpper = UCase$(Trim$(sContratoBaseAlvo))
    lQtdFiltradas = 0

    If lTotalLinhasOrig >= 2 Then
        vOrig = rngTabelaOrig.Value2

        For r = 2 To lTotalLinhasOrig
            If ExtrairContratoAteSegundoHifen(CStr(vOrig(r, colFiltroContrato) & "")) = sAlvoUpper Then
                lQtdFiltradas = lQtdFiltradas + 1
            End If
        Next r

        If lQtdFiltradas > 0 Then
            ReDim vOut(1 To lQtdFiltradas, 1 To lColsExport)
            k = 0
            For r = 2 To lTotalLinhasOrig
                If ExtrairContratoAteSegundoHifen(CStr(vOrig(r, colFiltroContrato) & "")) = sAlvoUpper Then
                    k = k + 1
                    For c = 1 To lColsExport
                        vCel = vOrig(r, c)
                        If VarType(vCel) = vbString Then
                            If Len(vCel) > 1 And Left$(vCel, 1) = "0" And IsNumeric(vCel) Then
                                vOut(k, c) = "'" & vCel
                            Else
                                vOut(k, c) = vCel
                            End If
                        Else
                            vOut(k, c) = vCel
                        End If
                    Next c
                End If
            Next r

            ' Grava dados filtrados de uma só vez e replica formatos numéricos/datas das colunas
            wsDestino.Range("A2").Resize(lQtdFiltradas, lColsExport).Value2 = vOut

            If Not rngFormatoLinhaOrig Is Nothing Then
                rngFormatoLinhaOrig.Copy
                wsDestino.Range("A2").Resize(lQtdFiltradas, lColsExport).PasteSpecial Paste:=xlPasteFormats
                Application.CutCopyMode = False
            End If
        ElseIf Not rngFormatoLinhaOrig Is Nothing Then
            rngFormatoLinhaOrig.Copy
            wsDestino.Range("A2").Resize(1, lColsExport).PasteSpecial Paste:=xlPasteFormats
            Application.CutCopyMode = False
        End If
    End If

    ' 3. Recria a Tabela Estruturada (ListObject) limpa no arquivo de destino
    On Error Resume Next
    If lQtdFiltradas > 0 Then
        lLinhasTabelaDest = lQtdFiltradas + 1
    Else
        lLinhasTabelaDest = 2
    End If

    Set rngTabelaDest = wsDestino.Range("A1").Resize(lLinhasTabelaDest, lColsExport)
    Set tblDest = wsDestino.ListObjects.Add(xlSrcRange, rngTabelaDest, , xlYes)
    If Not tblDest Is Nothing Then
        tblDest.Name = sNomeTabelaDest
        If Not tblOrig Is Nothing Then
            tblDest.TableStyle = tblOrig.TableStyle
        End If
    End If
    On Error GoTo 0
    Exit Sub

TratarErroLocal:
    Application.CutCopyMode = False
End Sub

'----------------------------------------------------------------------------------------------------
' EXTRAI O PREFIXO DO CONTRATO DESCONSIDERANDO TUDO APÓS O SEGUNDO '-'
' Exemplos: "IRON-LT1-RJ" -> "IRON-LT1" | "PA-LT2-BA" -> "PA-LT2" | "IRON-LT1" -> "IRON-LT1"
'----------------------------------------------------------------------------------------------------
Private Function ExtrairContratoAteSegundoHifen(ByVal sContratoBruto As String) As String
    Dim sLimpo As String
    Dim p1 As Long
    Dim p2 As Long

    sLimpo = UCase$(Trim$(Replace(sContratoBruto, Chr$(160), " ")))
    If Len(sLimpo) = 0 Then
        ExtrairContratoAteSegundoHifen = ""
        Exit Function
    End If

    p1 = InStr(1, sLimpo, "-")
    If p1 > 0 Then
        p2 = InStr(p1 + 1, sLimpo, "-")
        If p2 > 0 Then
            ExtrairContratoAteSegundoHifen = Left$(sLimpo, p2 - 1)
            Exit Function
        End If
    End If

    ExtrairContratoAteSegundoHifen = sLimpo
End Function

'----------------------------------------------------------------------------------------------------
' LOCALIZA A COLUNA DE CONTRATO CONSOLIDADO ('Contrato.1' > 'Contrato-Reg.' > 'Contrato')
'----------------------------------------------------------------------------------------------------
Private Function LocalizarColunaContratoConsolidado(ByVal rngTabela As Range) As Long
    Dim c As Long
    Dim lCols As Long
    Dim sHdr As String

    lCols = rngTabela.Columns.Count

    ' Prioridade 1: 'Contrato.1' ou 'Contrato-Reg.' (contrato oficial calculado no Power Query)
    For c = 1 To lCols
        sHdr = NormalizarCabecalho(CStr(rngTabela.Cells(1, c).Value & ""))
        If sHdr = "CONTRATO.1" Or sHdr = "CONTRATO-REG." Or InStr(1, sHdr, "CONTRATO-REG", vbTextCompare) > 0 Then
            LocalizarColunaContratoConsolidado = c
            Exit Function
        End If
    Next c

    ' Prioridade 2: 'Contrato'
    For c = 1 To lCols
        sHdr = NormalizarCabecalho(CStr(rngTabela.Cells(1, c).Value & ""))
        If sHdr = "CONTRATO" Then
            LocalizarColunaContratoConsolidado = c
            Exit Function
        End If
    Next c

    LocalizarColunaContratoConsolidado = 1
End Function

'----------------------------------------------------------------------------------------------------
' LOCALIZA A COLUNA 'Observações' (COLUNA U = 21) NA ABA 'Painel T2M'
'----------------------------------------------------------------------------------------------------
Private Function ObterColunaLimiteObservacoesT2M(ByVal wsT2M As Worksheet) As Long
    Dim rngRef As Range
    Dim c As Long
    Dim lMaxBusca As Long
    Dim sHdr As String

    If wsT2M.ListObjects.Count > 0 Then
        Set rngRef = wsT2M.ListObjects(1).HeaderRowRange
    Else
        Set rngRef = wsT2M.Rows(1)
    End If

    lMaxBusca = Application.WorksheetFunction.Min(rngRef.Columns.Count, 30)
    For c = 1 To lMaxBusca
        sHdr = NormalizarCabecalho(CStr(rngRef.Cells(1, c).Value & ""))
        If sHdr = "OBSERVACOES" Or sHdr = "OBSERVAÇÕES" Or Left$(sHdr, 9) = "OBSERVACO" Or Left$(sHdr, 9) = "OBSERVAÇÕ" Then
            ObterColunaLimiteObservacoesT2M = c
            Exit Function
        End If
    Next c

    ' Fallback seguro: Coluna U (índice 21)
    ObterColunaLimiteObservacoesT2M = 21
End Function

Private Function NormalizarCabecalho(ByVal sTexto As String) As String
    Dim s As String
    s = Replace(sTexto, "_x000a_", " ")
    s = Replace(s, Chr$(160), " ")
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    NormalizarCabecalho = UCase$(Trim$(s))
End Function

'----------------------------------------------------------------------------------------------------
' FUNÇÃO DE GERAÇÃO INDIVIDUAL DE ARQUIVO REGIONAL POR CONTRATO (MANTIDA INTEGRALMENTE)
'----------------------------------------------------------------------------------------------------
Private Function GerarArquivoDoContrato(ByVal wbOrigem As Workbook, ByVal sContrato As String, _
                                        ByVal sNumero As String, ByVal sSufixo As String, _
                                        ByVal vTipos As Variant, ByVal sPastaDestino As String) As String
    Dim wbDestino        As Workbook
    Dim wsOrigem         As Worksheet
    Dim wsDestino        As Worksheet
    Dim iTipo            As Long
    Dim sTipo            As String
    Dim sNomeAbaOrig     As String
    Dim iCopiadas        As Long
    Dim sFaltantes       As String
    Dim sNomeArquivo     As String
    Dim sCaminhoCompleto As String
    Dim sSufixoComHifen  As String
    Dim rngArea          As Range

    On Error GoTo TratarErroLocal
    GerarArquivoDoContrato = ""

    Set wbDestino = Workbooks.Add(xlWBATWorksheet)
    Do While wbDestino.Worksheets.Count > 1
        wbDestino.Worksheets(wbDestino.Worksheets.Count).Delete
    Loop

    iCopiadas = 0
    sFaltantes = ""

    For iTipo = LBound(vTipos) To UBound(vTipos)
        sTipo = CStr(vTipos(iTipo))
        sNomeAbaOrig = sTipo & "_" & sContrato

        Set wsOrigem = ObterPlanilha(wbOrigem, sNomeAbaOrig)

        If iTipo = LBound(vTipos) Then
            Set wsDestino = wbDestino.Worksheets(1)
        Else
            Set wsDestino = wbDestino.Worksheets.Add(After:=wbDestino.Worksheets(wbDestino.Worksheets.Count))
        End If

        wsDestino.Name = sTipo

        If wsOrigem Is Nothing Then
            sFaltantes = sFaltantes & " - Aba '" & sNomeAbaOrig & "' não encontrada." & vbNewLine
            wsDestino.Range("A1").Value = "Aba de origem '" & sNomeAbaOrig & "' não encontrada."
        Else
            If sTipo = "MC" Then
                CopiarComFormatacaoCompleta wsOrigem.UsedRange, wsDestino
            Else
                Set rngArea = ObterAreaRealDados(wsOrigem)
                CopiarValoresEFormatos rngArea, wsDestino
            End If
            iCopiadas = iCopiadas + 1
        End If
    Next iTipo

    If iCopiadas = 0 Then
        wbDestino.Close SaveChanges:=False
        Exit Function
    End If

    AjustarFormulasAbaMC wbDestino, wbOrigem, sContrato
    AplicarPerfilPublico wbDestino

    If Len(sSufixo) > 0 Then
        sSufixoComHifen = "-" & sSufixo
    Else
        sSufixoComHifen = ""
    End If

    sNomeArquivo = sNumero & "-PLA-MemoriaPetrobras-" & sContrato & sSufixoComHifen & "_" & _
                   Format(Now, "ddmmyyyy_hhnnss") & ".xlsx"

    sCaminhoCompleto = sPastaDestino & Application.PathSeparator & sNomeArquivo

    wbDestino.SaveAs Filename:=sCaminhoCompleto, FileFormat:=xlOpenXMLWorkbook
    wbDestino.Close SaveChanges:=False

    GerarArquivoDoContrato = sCaminhoCompleto
    Exit Function

TratarErroLocal:
    On Error Resume Next
    If Not wbDestino Is Nothing Then wbDestino.Close SaveChanges:=False
    GerarArquivoDoContrato = ""
End Function

'----------------------------------------------------------------------------------------------------
' ROTINAS DE CÓPIA OTIMIZADA E FORMATAÇÃO (CADERNOS REGIONAIS)
'----------------------------------------------------------------------------------------------------
Private Sub CopiarValoresEFormatos(ByVal rngOrigem As Range, ByVal wsDestino As Worksheet)
    If rngOrigem Is Nothing Then Exit Sub
    On Error GoTo TratarErroLocal

    rngOrigem.Copy
    wsDestino.Range("A1").PasteSpecial Paste:=xlPasteValues
    wsDestino.Range("A1").PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False

    ReplicarLargurasColunas rngOrigem, wsDestino

    On Error Resume Next
    If rngOrigem.Worksheet.ListObjects.Count > 0 Then
        Dim tblOrig As ListObject
        Dim tblDest As ListObject
        Dim rngTabelaDest As Range

        Set tblOrig = rngOrigem.Worksheet.ListObjects(1)
        If rngOrigem.Rows.Count >= 2 Then
            Set rngTabelaDest = wsDestino.Range("A1").Resize(rngOrigem.Rows.Count, rngOrigem.Columns.Count)
            Set tblDest = wsDestino.ListObjects.Add(xlSrcRange, rngTabelaDest, , xlYes)
            If Not tblDest Is Nothing Then
                tblDest.Name = tblOrig.Name
                tblDest.TableStyle = tblOrig.TableStyle
            End If
        End If
    End If
    On Error GoTo 0
    Exit Sub

TratarErroLocal:
    Application.CutCopyMode = False
End Sub

Private Sub CopiarComFormatacaoCompleta(ByVal rngOrigem As Range, ByVal wsDestino As Worksheet)
    If rngOrigem Is Nothing Then Exit Sub
    On Error GoTo TratarErroLocal

    rngOrigem.Copy
    wsDestino.Range("A1").PasteSpecial Paste:=xlPasteAll
    Application.CutCopyMode = False

    ReplicarLargurasColunas rngOrigem, wsDestino
    ReplicarAlturasLinhas rngOrigem, wsDestino
    Exit Sub

TratarErroLocal:
    Application.CutCopyMode = False
    On Error Resume Next
    rngOrigem.Copy
    wsDestino.Range("A1").PasteSpecial Paste:=xlPasteAll
    Application.CutCopyMode = False
    On Error GoTo 0
End Sub

'----------------------------------------------------------------------------------------------------
' AJUSTE DE FÓRMULAS DA ABA MC PARA REFERENCIAR AS ABAS DO NOVO ARQUIVO
'----------------------------------------------------------------------------------------------------
Private Sub AjustarFormulasAbaMC(ByVal wbDestino As Workbook, ByVal wbOrigem As Workbook, ByVal sContrato As String)
    Dim wsMC                As Worksheet
    Dim rngFormulas         As Range
    Dim cel                 As Range
    Dim sFormula            As String
    Dim sFormulaOriginal    As String
    Dim sNomeWbOrigem       As String
    Dim sCaminhoWbOrigem    As String
    Dim sAbaOrigARM         As String
    Dim sAbaOrigDADOS       As String
    Dim sAbaOrigFRETE       As String
    Dim sAbaOrigMC          As String
    Dim vLinks              As Variant
    Dim iLink               As Long

    On Error Resume Next
    Set wsMC = wbDestino.Worksheets("MC")
    On Error GoTo 0
    If wsMC Is Nothing Then Exit Sub

    On Error Resume Next
    wbDestino.Calculate
    On Error GoTo 0

    sNomeWbOrigem = wbOrigem.Name
    sCaminhoWbOrigem = wbOrigem.FullName
    sAbaOrigARM = "ARM_" & sContrato
    sAbaOrigDADOS = "DADOS_" & sContrato
    sAbaOrigFRETE = "FRETE_" & sContrato
    sAbaOrigMC = "MC_" & sContrato

    On Error Resume Next
    Set rngFormulas = wsMC.UsedRange.SpecialCells(xlCellTypeFormulas)
    On Error GoTo 0

    If Not rngFormulas Is Nothing Then
        For Each cel In rngFormulas
            sFormula = cel.Formula
            sFormulaOriginal = sFormula

            If InStr(1, sFormula, sCaminhoWbOrigem, vbTextCompare) > 0 Then
                sFormula = Replace(sFormula, "'[" & sCaminhoWbOrigem & "]'", "")
                sFormula = Replace(sFormula, "[" & sCaminhoWbOrigem & "]", "")
                sFormula = Replace(sFormula, "'" & sCaminhoWbOrigem & "'!", "")
                sFormula = Replace(sFormula, sCaminhoWbOrigem & "!", "")
            End If

            If InStr(1, sFormula, sNomeWbOrigem, vbTextCompare) > 0 Then
                sFormula = Replace(sFormula, "'[" & sNomeWbOrigem & "]'", "")
                sFormula = Replace(sFormula, "[" & sNomeWbOrigem & "]", "")
                sFormula = Replace(sFormula, "'" & sNomeWbOrigem & "'!", "")
                sFormula = Replace(sFormula, sNomeWbOrigem & "!", "")
            End If

            If InStr(1, sFormula, sAbaOrigARM, vbTextCompare) > 0 Then
                sFormula = Replace(sFormula, "'" & sAbaOrigARM & "'!", "ARM!")
                sFormula = Replace(sFormula, sAbaOrigARM & "!", "ARM!")
            End If

            If InStr(1, sFormula, sAbaOrigDADOS, vbTextCompare) > 0 Then
                sFormula = Replace(sFormula, "'" & sAbaOrigDADOS & "'!", "DADOS!")
                sFormula = Replace(sFormula, sAbaOrigDADOS & "!", "DADOS!")
            End If

            If InStr(1, sFormula, sAbaOrigFRETE, vbTextCompare) > 0 Then
                sFormula = Replace(sFormula, "'" & sAbaOrigFRETE & "'!", "FRETE!")
                sFormula = Replace(sFormula, sAbaOrigFRETE & "!", "FRETE!")
            End If

            If InStr(1, sFormula, sAbaOrigMC, vbTextCompare) > 0 Then
                sFormula = Replace(sFormula, "'" & sAbaOrigMC & "'!", "MC!")
                sFormula = Replace(sFormula, sAbaOrigMC & "!", "MC!")
            End If

            If sFormula <> sFormulaOriginal Then
                On Error Resume Next
                cel.Formula = sFormula
                On Error GoTo 0
            End If
        Next cel
    End If

    On Error Resume Next
    vLinks = wbDestino.LinkSources(xlExcelLinks)
    If Not IsEmpty(vLinks) Then
        For iLink = LBound(vLinks) To UBound(vLinks)
            wbDestino.BreakLink Name:=vLinks(iLink), Type:=xlLinkTypeExcelLinks
        Next iLink
    End If

    wbDestino.Calculate
    On Error GoTo 0
End Sub

Private Sub ReplicarLargurasColunas(ByVal rngOrigem As Range, ByVal wsDestino As Worksheet)
    On Error Resume Next
    rngOrigem.Rows(1).Copy
    wsDestino.Range("A1").PasteSpecial Paste:=xlPasteColumnWidths
    Application.CutCopyMode = False

    If Err.Number <> 0 Then
        Err.Clear
        Dim iCol As Long
        For iCol = 1 To rngOrigem.Columns.Count
            wsDestino.Columns(iCol).ColumnWidth = rngOrigem.Worksheet.Columns( _
                rngOrigem.Column + iCol - 1).ColumnWidth
        Next iCol
    End If
    On Error GoTo 0
End Sub

Private Sub ReplicarAlturasLinhas(ByVal rngOrigem As Range, ByVal wsDestino As Worksheet)
    Dim iLin As Long
    For iLin = 1 To rngOrigem.Rows.Count
        wsDestino.Rows(iLin).RowHeight = rngOrigem.Worksheet.Rows( _
            rngOrigem.Row + iLin - 1).RowHeight
    Next iLin
End Sub

Private Function ObterAreaRealDados(ByVal ws As Worksheet) As Range
    Dim ultimaColuna As Long
    Dim ultimaLinha As Long

    If ws.Cells(1, ws.Columns.Count).Value <> "" Then
        ultimaColuna = ws.Columns.Count
    Else
        ultimaColuna = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    End If

    If ws.Cells(ws.Rows.Count, 1).Value <> "" Then
        ultimaLinha = ws.Rows.Count
    Else
        ultimaLinha = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    End If

    If ultimaColuna < 1 Then ultimaColuna = 1
    If ultimaLinha < 1 Then ultimaLinha = 1

    Set ObterAreaRealDados = ws.Range(ws.Cells(1, 1), ws.Cells(ultimaLinha, ultimaColuna))
End Function

'----------------------------------------------------------------------------------------------------
' GERENCIAMENTO DE LINKS E STATUS NA ABA PAINEL (F14:F17 REGIONAIS E F19:F20 CONSOLIDADOS)
'----------------------------------------------------------------------------------------------------
Private Function ObterLinhaCelulaLink(ByVal iIndiceContrato As Long) As Long
    ObterLinhaCelulaLink = LINHA_INICIAL_LINK + iIndiceContrato
End Function

Private Function ObterLinhaCelulaLinkConsolidado(ByVal iIndiceCons As Long) As Long
    ObterLinhaCelulaLinkConsolidado = LINHA_INICIAL_LINK_CONSOLIDADO + iIndiceCons
End Function

Private Sub CopiarEstiloCelulaModelo(ByVal ws As Worksheet, ByVal lLinhaModelo As Long, ByVal lLinhaAlvo As Long)
    On Error Resume Next
    Dim rngMod As Range
    Dim rngAlv As Range
    Set rngMod = ws.Range(COLUNA_CELULA_LINK & lLinhaModelo)
    Set rngAlv = ws.Range(COLUNA_CELULA_LINK & lLinhaAlvo)
    rngAlv.Font.Name = rngMod.Font.Name
    rngAlv.Font.Size = rngMod.Font.Size
    rngAlv.HorizontalAlignment = rngMod.HorizontalAlignment
    rngAlv.VerticalAlignment = rngMod.VerticalAlignment
    ws.Rows(lLinhaAlvo).RowHeight = ws.Rows(lLinhaModelo).RowHeight
    On Error GoTo 0
End Sub

Private Sub MarcarCelulaEmAndamento(ByVal ws As Worksheet, ByVal lLinha As Long, ByVal sRotulo As String)
    Dim rngIni As Range
    Set rngIni = ws.Range(COLUNA_CELULA_LINK & lLinha)
    rngIni.Hyperlinks.Delete
    rngIni.Value = "Gerando: " & sRotulo & "..."
    If rngIni.MergeCells Then
        rngIni.MergeArea.Interior.Color = COR_STATUS_AMARELO
        rngIni.MergeArea.Font.Color = vbBlack
        rngIni.MergeArea.Font.Bold = False
    Else
        rngIni.Interior.Color = COR_STATUS_AMARELO
        rngIni.Font.Color = vbBlack
        rngIni.Font.Bold = False
    End If
End Sub

Private Sub EscreverLinkArquivoEmLinha(ByVal ws As Worksheet, ByVal iLinha As Long, _
                                       ByVal sCaminhoCompleto As String)
    Dim rngCelula As Range
    Dim sNomeArq As String

    On Error GoTo TratarErroLocal
    Set rngCelula = ws.Range(COLUNA_CELULA_LINK & iLinha)
    sNomeArq = ExtrairNomeArquivo(sCaminhoCompleto)

    rngCelula.Hyperlinks.Delete
    rngCelula.ClearContents

    If rngCelula.MergeCells Then
        rngCelula.MergeArea.Interior.Color = COR_STATUS_VERDE_CLARO
        rngCelula.MergeArea.Font.Bold = False
    Else
        rngCelula.Interior.Color = COR_STATUS_VERDE_CLARO
        rngCelula.Font.Bold = False
    End If

    ws.Hyperlinks.Add Anchor:=rngCelula, Address:=sCaminhoCompleto, TextToDisplay:=sNomeArq
    Exit Sub

TratarErroLocal:
    On Error Resume Next
    rngCelula.Value = sCaminhoCompleto
    If rngCelula.MergeCells Then
        rngCelula.MergeArea.Interior.Color = COR_STATUS_VERMELHO
        rngCelula.MergeArea.Font.Color = vbWhite
        rngCelula.MergeArea.Font.Bold = True
    Else
        rngCelula.Interior.Color = COR_STATUS_VERMELHO
        rngCelula.Font.Color = vbWhite
        rngCelula.Font.Bold = True
    End If
    On Error GoTo 0
End Sub

Private Sub EscreverErroLinkEmLinha(ByVal ws As Worksheet, ByVal iLinha As Long, ByVal sContrato As String)
    Dim rngCelula As Range
    On Error Resume Next

    Set rngCelula = ws.Range(COLUNA_CELULA_LINK & iLinha)
    rngCelula.Hyperlinks.Delete
    rngCelula.Value = "Falha ao gerar arquivo de " & sContrato

    If rngCelula.MergeCells Then
        rngCelula.MergeArea.Interior.Color = COR_STATUS_VERMELHO
        rngCelula.MergeArea.Font.Color = vbWhite
        rngCelula.MergeArea.Font.Bold = True
    Else
        rngCelula.Interior.Color = COR_STATUS_VERMELHO
        rngCelula.Font.Color = vbWhite
        rngCelula.Font.Bold = True
    End If
    On Error GoTo 0
End Sub

Private Sub SinalizarVermelhoSeSemLink(ByVal ws As Worksheet, ByVal iLinha As Long)
    Dim rngFim As Range
    Set rngFim = ws.Range(COLUNA_CELULA_LINK & iLinha)
    If rngFim.Hyperlinks.Count = 0 Then
        If rngFim.MergeCells Then
            rngFim.MergeArea.Interior.Color = COR_STATUS_VERMELHO
            rngFim.MergeArea.Font.Color = vbWhite
            rngFim.MergeArea.Font.Bold = True
        Else
            rngFim.Interior.Color = COR_STATUS_VERMELHO
            rngFim.Font.Color = vbWhite
            rngFim.Font.Bold = True
        End If
    End If
End Sub

Private Sub SinalizarErroInesperadoCelula(ByVal ws As Worksheet, ByVal iLinha As Long, ByVal sRotulo As String)
    Dim rngErr As Range
    Set rngErr = ws.Range(COLUNA_CELULA_LINK & iLinha)
    If rngErr.Hyperlinks.Count = 0 Then
        rngErr.Hyperlinks.Delete
        If rngErr.MergeCells Then
            rngErr.MergeArea.Interior.Color = COR_STATUS_VERMELHO
            rngErr.MergeArea.Font.Color = vbWhite
            rngErr.MergeArea.Font.Bold = True
        Else
            rngErr.Interior.Color = COR_STATUS_VERMELHO
            rngErr.Font.Color = vbWhite
            rngErr.Font.Bold = True
        End If
        If Len(rngErr.Value) = 0 Or InStr(rngErr.Value, "Gerando") > 0 Then
            rngErr.Value = "Erro ao gerar: " & sRotulo
        End If
    End If
End Sub

Private Function ExtrairNomeArquivo(ByVal sCaminhoCompleto As String) As String
    ExtrairNomeArquivo = Mid$(sCaminhoCompleto, InStrRev(sCaminhoCompleto, Application.PathSeparator) + 1)
End Function

'----------------------------------------------------------------------------------------------------
' UTILITÁRIOS DE SISTEMA DE ARQUIVOS E PLANILHAS
'----------------------------------------------------------------------------------------------------
Private Function ObterOuCriarSubpasta(ByVal sPastaBase As String, ByVal sNomeSubpasta As String) As String
    Dim sCaminhoCompleto As String
    On Error GoTo TratarErroLocal

    sCaminhoCompleto = sPastaBase & Application.PathSeparator & sNomeSubpasta
    If Len(Dir(sCaminhoCompleto, vbDirectory)) = 0 Then
        MkDir sCaminhoCompleto
    End If

    ObterOuCriarSubpasta = sCaminhoCompleto
    Exit Function

TratarErroLocal:
    ObterOuCriarSubpasta = ""
End Function

Private Function ObterOuCriarPlanilha(ByVal wb As Workbook, ByVal nomeAba As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(nomeAba)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add
        ws.Name = nomeAba
    End If

    Set ObterOuCriarPlanilha = ws
End Function

Private Function ObterPlanilha(ByVal wb As Workbook, ByVal nomeAba As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(nomeAba)
    On Error GoTo 0
    Set ObterPlanilha = ws
End Function

Private Sub AplicarPerfilPublico(ByVal wb As Workbook)
    If Len(PUBLICO_LABEL_ID) = 0 Then Exit Sub
    On Error GoTo TratarErroLocal

    Dim lbl As Object
    Set lbl = wb.SensitivityLabel.CreateLabelInfo
    lbl.LabelId = PUBLICO_LABEL_ID
    lbl.AssignmentMethod = MSO_ASSIGNMENT_METHOD_STANDARD
    lbl.Justification = "Rótulo aplicado automaticamente na geração do arquivo."

    wb.SensitivityLabel.SetLabel lbl, lbl
    Exit Sub

TratarErroLocal:
    ' Falha não-bloqueante caso o ambiente Office não possua AIP/MIP habilitado
End Sub
