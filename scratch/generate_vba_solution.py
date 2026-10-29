import os

vba_code = r'''Attribute VB_Name = "modDemonstrativoFDM"
Option Explicit

'====================================================================================================
' MODULO: modDemonstrativoFDM
' OBJETIVO: Apurar os indicadores contratuais de qualidade e nivel de servico (SLA):
'           - TSPE: Total de Solicitacoes Padrao Esperado (Demanda Elegivel: NP + FP)
'           - TSPR: Total de Solicitacoes Padrao Realizado (Atendidas no Prazo: NP)
'           - IAPFARQ: Indice de Atendimento aos Prazos de Arquivo ((TSPR / TSPE) * 100)
'           - FDM: Fator de Desempenho Mensal e Glosa Contratual (Itens 8.4.2 e 8.4.3 do Edital)
'           Gera e atualiza a planilha oficial:
'           "Demonstrativo_Calculo_FDM_TSPE_IAPFARQ.xlsx" na pasta MEDICAO e na raiz.
'
' ABAS DA PLANILHA DEMONSTRATIVA:
'   1. Resumo_Executivo_FDM: Apuracao contratual por lote e servico, cards KPI e formulas ativas.
'   2. Detalhamento_Atividades: Abertura operacional por atividade com diagnostico de ofensores.
'   3. Base_Analitica_Ordens: Relacao completa de ordens com 20 colunas e sinalizacao de prazo.
'====================================================================================================

Private Const MSG_TITULO As String = "Demonstrativo de Calculo FDM / TSPE / IAPFARQ"
Private Const NOME_ABA_T2M As String = "Painel T2M"
Private Const NOME_ABA_FDM As String = "TabelaFDMPorContrato"
Private Const NOME_ABA_PAINEL As String = "Painel"
Private Const NOME_ARQUIVO_FDM As String = "Demonstrativo_Calculo_FDM_TSPE_IAPFARQ.xlsx"
Private Const NOME_SUBPASTA_DESTINO As String = "MEDIÇÃO"
Private Const CELULA_LINK_FDM As String = "F18"

' Paleta de Cores Corporativas Petrobras (RGB)
Private Const COR_NAVY As Long = 3749403            ' RGB(27, 54, 93)   - #1B365D
Private Const COR_HEADER_SUB As Long = 5196330      ' RGB(42, 74, 127)  - #2A4A7F
Private Const COR_TEAL As Long = 4877056            ' RGB(0, 107, 118)  - #006B76
Private Const COR_DARK_SLATE As Long = 3231535      ' RGB(47, 79, 79)   - #2F4F4F
Private Const COR_CARD_BG As Long = 16446443        ' RGB(235, 243, 250)- #EBF3FA
Private Const COR_TOTAL As Long = 15724521          ' RGB(233, 236, 239)- #E9ECEF
Private Const COR_GREEN_BG As Long = 14347732       ' RGB(212, 237, 218)- #D4EDDA
Private Const COR_GREEN_TXT As Long = 2381589       ' RGB(21, 87, 36)   - #155724
Private Const COR_RED_BG As Long = 14342136         ' RGB(248, 215, 218)- #F8D7DA
Private Const COR_RED_TXT As Long = 2366578         ' RGB(114, 28, 36)  - #721C24
Private Const COR_GRAY_BORDER As Long = 14408913    ' RGB(209, 213, 219)- #D1D5DB
Private Const COR_STATUS_VERDE As Long = 5296274    ' RGB(146, 208, 80)
Private Const COR_STATUS_VERMELHO As Long = 255     ' RGB(255, 0, 0)

Public Sub GerarDemonstrativoCalculoFDM()
    ExecutarGeracaoFDM False
End Sub

Public Sub GerarDemonstrativoCalculoFDM_Silencioso()
    ExecutarGeracaoFDM True
End Sub

Private Sub ExecutarGeracaoFDM(ByVal bSilencioso As Boolean)
    Dim sMsgErro As String
    sMsgErro = ""
    On Error GoTo TratarErro

    Dim wsT2M As Worksheet, wsFDM As Worksheet, wsPainel As Worksheet
    Dim wbDest As Workbook
    Dim ws1 As Worksheet, ws2 As Worksheet, ws3 As Worksheet
    Dim sPastaDestino As String, sCaminhoFinalMedicao As String, sCaminhoFinalRaiz As String
    Dim sPeriodoTexto As String

    On Error Resume Next
    Set wsT2M = ThisWorkbook.Worksheets(NOME_ABA_T2M)
    Set wsFDM = ThisWorkbook.Worksheets(NOME_ABA_FDM)
    Set wsPainel = ThisWorkbook.Worksheets(NOME_ABA_PAINEL)
    On Error GoTo TratarErro

    If wsT2M Is Nothing Or wsFDM Is Nothing Then
        sMsgErro = "Planilhas obrigatorias nao localizadas ('" & NOME_ABA_T2M & "' ou '" & NOME_ABA_FDM & "')."
        GoTo TratarErro
    End If

    sPeriodoTexto = "Ciclo Mensal de Faturamento (26 a 25)"
    If Not wsPainel Is Nothing Then
        If Len(Trim(CStr(wsPainel.Range("C7").Value))) > 0 Then
            sPeriodoTexto = "Referencia: " & Trim(CStr(wsPainel.Range("C7").Value)) & "/" & Trim(CStr(wsPainel.Range("D7").Value))
        End If
    End If

    ' Definir pastas de saida
    sPastaDestino = ObterOuCriarSubpasta(ThisWorkbook.Path, NOME_SUBPASTA_DESTINO)
    sCaminhoFinalMedicao = sPastaDestino & Application.PathSeparator & NOME_ARQUIVO_FDM
    sCaminhoFinalRaiz = ThisWorkbook.Path & Application.PathSeparator & NOME_ARQUIVO_FDM

    ' Fechar demonstrativo se ja estiver aberto nesta sessao do Excel
    Dim wbExistente As Workbook
    For Each wbExistente In Application.Workbooks
        If UCase(wbExistente.Name) = UCase(NOME_ARQUIVO_FDM) Then
            wbExistente.Close SaveChanges:=False
            Exit For
        End If
    Next wbExistente

    ' 1. Lendo TabelaFDMPorContrato
    Dim fdm1_final_iron As Double, fdm2_final_iron As Double
    Dim fdm1_final_pa As Double, fdm2_final_pa As Double
    fdm1_final_iron = 0.95: fdm2_final_iron = 0.95
    fdm1_final_pa = 0.95: fdm2_final_pa = 0.95

    Dim rF As Long, ultLinFDM As Long
    ultLinFDM = wsFDM.Cells(wsFDM.Rows.Count, 1).End(xlUp).Row
    For rF = 2 To ultLinFDM
        Dim sContF As String
        sContF = UCase(Trim(CStr(wsFDM.Cells(rF, 1).Value)))
        If InStr(sContF, "IRON") > 0 Then
            If IsNumeric(wsFDM.Cells(rF, 6).Value) Then fdm1_final_iron = CDbl(wsFDM.Cells(rF, 6).Value)
            If IsNumeric(wsFDM.Cells(rF, 11).Value) Then fdm2_final_iron = CDbl(wsFDM.Cells(rF, 11).Value)
        ElseIf InStr(sContF, "PA") > 0 Then
            If IsNumeric(wsFDM.Cells(rF, 6).Value) Then fdm1_final_pa = CDbl(wsFDM.Cells(rF, 6).Value)
            If IsNumeric(wsFDM.Cells(rF, 11).Value) Then fdm2_final_pa = CDbl(wsFDM.Cells(rF, 11).Value)
        End If
    Next rF

    ' 2. Lendo matriz Painel T2M
    Dim vDados As Variant
    vDados = wsT2M.UsedRange.Value
    Dim nLinhas As Long, nCols As Long
    nLinhas = UBound(vDados, 1)
    nCols = UBound(vDados, 2)

    Dim colSituacao As Long, colContrato As Long, colAtividade As Long
    Dim colItem As Long, colAplicacao As Long, colSolicitacao As Long
    Dim colOS As Long, colDtAbertura As Long, colDtFechamento As Long
    Dim colPrazoCombinado As Long, colMunicipio As Long, colUF As Long
    Dim colRefFDM As Long, colRefPrazo As Long, colPrazoSLA As Long
    Dim colClassifFDM As Long, colStatusPrazo As Long, colLogSLA As Long
    colSituacao = 2: colContrato = 3: colAtividade = 4: colItem = 5: colAplicacao = 6
    colSolicitacao = 7: colOS = 12: colDtAbertura = 15: colDtFechamento = 16: colPrazoCombinado = 17
    colMunicipio = 24: colUF = 25: colRefFDM = 34: colRefPrazo = 36: colPrazoSLA = 40
    colClassifFDM = 45: colStatusPrazo = 46: colLogSLA = 47

    Dim tot_iron1_co As Long, excl_iron1 As Long, tspr_iron1 As Long, fp_iron1 As Long
    Dim tot_iron2_co As Long, excl_iron2 As Long, tspr_iron2 As Long, fp_iron2 As Long
    Dim tot_pa1_co As Long, excl_pa1 As Long, tspr_pa1 As Long, fp_pa1 As Long
    Dim tot_pa2_co As Long, excl_pa2 As Long, tspr_pa2 As Long, fp_pa2 As Long

    Dim dictAtiv As Object
    Set dictAtiv = CreateObject("Scripting.Dictionary")

    Dim vBaseSaida() As Variant
    ReDim vBaseSaida(1 To nLinhas, 1 To 20)
    Dim nOrdensBase As Long
    nOrdensBase = 0

    ' 3. Processando linhas em memoria
    Dim r As Long
    For r = 2 To nLinhas
        Dim sSit As String
        sSit = UCase(Trim(CStr(vDados(r, colSituacao))))
        If sSit = "CO" Then
            Dim sCont As String, sRefF As String, sStPrazo As String, sAtiv As String
            sCont = UCase(Trim(CStr(vDados(r, colContrato))))
            sRefF = Trim(CStr(vDados(r, colRefFDM)))
            sStPrazo = UCase(Trim(CStr(vDados(r, colStatusPrazo))))
            sAtiv = Trim(CStr(vDados(r, colAtividade)))

            If InStr(sCont, "IRON") > 0 Then
                sCont = "IRON-LT1"
            ElseIf InStr(sCont, "PA") > 0 Then
                sCont = "PA-LT2"
            End If

            If sCont = "IRON-LT1" Then
                If sRefF = "1" Then
                    tot_iron1_co = tot_iron1_co + 1
                    If sStPrazo = "NP" Then
                        tspr_iron1 = tspr_iron1 + 1
                    ElseIf sStPrazo = "FP" Then
                        fp_iron1 = fp_iron1 + 1
                    Else
                        excl_iron1 = excl_iron1 + 1
                    End If
                ElseIf sRefF = "2" Then
                    tot_iron2_co = tot_iron2_co + 1
                    If sStPrazo = "NP" Then
                        tspr_iron2 = tspr_iron2 + 1
                    ElseIf sStPrazo = "FP" Then
                        fp_iron2 = fp_iron2 + 1
                    Else
                        excl_iron2 = excl_iron2 + 1
                    End If
                End If
            ElseIf sCont = "PA-LT2" Then
                If sRefF = "1" Then
                    tot_pa1_co = tot_pa1_co + 1
                    If sStPrazo = "NP" Then
                        tspr_pa1 = tspr_pa1 + 1
                    ElseIf sStPrazo = "FP" Then
                        fp_pa1 = fp_pa1 + 1
                    Else
                        excl_pa1 = excl_pa1 + 1
                    End If
                ElseIf sRefF = "2" Then
                    tot_pa2_co = tot_pa2_co + 1
                    If sStPrazo = "NP" Then
                        tspr_pa2 = tspr_pa2 + 1
                    ElseIf sStPrazo = "FP" Then
                        fp_pa2 = fp_pa2 + 1
                    Else
                        excl_pa2 = excl_pa2 + 1
                    End If
                End If
            End If

            Dim sKeyAtiv As String
            sKeyAtiv = sCont & "|" & sRefF & "|" & sAtiv
            Dim aStats As Variant
            Dim iNP As Long, iFP As Long, iOut As Long
            If sStPrazo = "NP" Then
                iNP = 1: iFP = 0: iOut = 0
            ElseIf sStPrazo = "FP" Then
                iNP = 0: iFP = 1: iOut = 0
            Else
                iNP = 0: iFP = 0: iOut = 1
            End If

            If Not dictAtiv.Exists(sKeyAtiv) Then
                dictAtiv.Add sKeyAtiv, Array(1, iNP, iFP, iOut)
            Else
                aStats = dictAtiv(sKeyAtiv)
                aStats(0) = aStats(0) + 1
                aStats(1) = aStats(1) + iNP
                aStats(2) = aStats(2) + iFP
                aStats(3) = aStats(3) + iOut
                dictAtiv(sKeyAtiv) = aStats
            End If

            Dim lRefNumRow As Long
            If IsNumeric(sRefF) Then
                lRefNumRow = CLng(sRefF)
            Else
                lRefNumRow = 0
            End If

            nOrdensBase = nOrdensBase + 1
            vBaseSaida(nOrdensBase, 1) = vDados(r, 1)
            vBaseSaida(nOrdensBase, 2) = sCont
            vBaseSaida(nOrdensBase, 3) = vDados(r, colOS)
            vBaseSaida(nOrdensBase, 4) = vDados(r, colSolicitacao)
            vBaseSaida(nOrdensBase, 5) = sAtiv
            vBaseSaida(nOrdensBase, 6) = vDados(r, colItem)
            vBaseSaida(nOrdensBase, 7) = vDados(r, colAplicacao)
            vBaseSaida(nOrdensBase, 8) = vDados(r, colMunicipio)
            vBaseSaida(nOrdensBase, 9) = vDados(r, colUF)
            vBaseSaida(nOrdensBase, 10) = vDados(r, colDtAbertura)
            vBaseSaida(nOrdensBase, 11) = vDados(r, colDtFechamento)
            vBaseSaida(nOrdensBase, 12) = vDados(r, colPrazoCombinado)
            vBaseSaida(nOrdensBase, 13) = lRefNumRow
            vBaseSaida(nOrdensBase, 14) = vDados(r, colRefPrazo)
            vBaseSaida(nOrdensBase, 15) = vDados(r, colPrazoSLA)
            vBaseSaida(nOrdensBase, 16) = vDados(r, colClassifFDM)
            vBaseSaida(nOrdensBase, 17) = sStPrazo
            
            Dim bContaTSPE As Long, bContaTSPR As Long
            If (sStPrazo = "NP" Or sStPrazo = "FP") And (sRefF = "1" Or sRefF = "2") Then
                bContaTSPE = 1
            Else
                bContaTSPE = 0
            End If
            If sStPrazo = "NP" And (sRefF = "1" Or sRefF = "2") Then
                bContaTSPR = 1
            Else
                bContaTSPR = 0
            End If
            
            vBaseSaida(nOrdensBase, 18) = bContaTSPE
            vBaseSaida(nOrdensBase, 19) = bContaTSPR
            vBaseSaida(nOrdensBase, 20) = vDados(r, colLogSLA)
        End If
    Next r

    ' 4. Criando wbDest e abas
    Set wbDest = Application.Workbooks.Add(xlWBATWorksheet)
    Set ws1 = wbDest.Worksheets(1)
    ws1.Name = "Resumo_Executivo_FDM"
    Set ws2 = wbDest.Worksheets.Add(After:=ws1)
    ws2.Name = "Detalhamento_Atividades"
    Set ws3 = wbDest.Worksheets.Add(After:=ws2)
    ws3.Name = "Base_Analitica_Ordens"

    ' -------------------------------------------------------------
    ' ABA 1: Resumo_Executivo_FDM
    ' -------------------------------------------------------------
    ws1.Range("A1:P1").Merge
    ws1.Range("A1").Value = "RELATORIO DE DEMONSTRACAO E AUDITORIA DE INDICADORES CONTRATUAIS - FDM / TSPE / TSPR / IAPFARQ"
    ws1.Range("A1").Font.Name = "Segoe UI": ws1.Range("A1").Font.Size = 13: ws1.Range("A1").Font.Bold = True
    ws1.Range("A1").Font.Color = vbWhite: ws1.Range("A1").Interior.Color = COR_NAVY
    ws1.Range("A1").HorizontalAlignment = xlCenter: ws1.Range("A1").VerticalAlignment = xlCenter
    ws1.Rows(1).RowHeight = 30

    ws1.Range("A2:P2").Merge
    ws1.Range("A2").Value = "Governanca Contratual Petrobras | Lote 1 (IRON MOUNTAIN) e Lote 2 (PA) | " & sPeriodoTexto
    ws1.Range("A2").Font.Name = "Segoe UI": ws1.Range("A2").Font.Size = 9: ws1.Range("A2").Font.Italic = True
    ws1.Range("A2").Font.Color = RGB(226, 232, 240): ws1.Range("A2").Interior.Color = COR_HEADER_SUB
    ws1.Range("A2").HorizontalAlignment = xlCenter: ws1.Range("A2").VerticalAlignment = xlCenter
    ws1.Rows(2).RowHeight = 20

    Dim totElegiveis As Long, totPontuais As Long, totAtrasos As Long
    totElegiveis = (tspr_iron1 + fp_iron1) + (tspr_iron2 + fp_iron2) + (tspr_pa1 + fp_pa1) + (tspr_pa2 + fp_pa2)
    totPontuais = tspr_iron1 + tspr_iron2 + tspr_pa1 + tspr_pa2
    totAtrasos = fp_iron1 + fp_iron2 + fp_pa1 + fp_pa2
    
    Dim pctGlobal As Double
    If totElegiveis > 0 Then
        pctGlobal = (CDbl(totPontuais) / CDbl(totElegiveis)) * 100#
    Else
        pctGlobal = 100#
    End If

    CriarCardKPI ws1, "A4:C5", CStr(nOrdensBase), "Total Chamados Concluidos (CO)", "Base total faturavel"
    CriarCardKPI ws1, "D4:F5", CStr(totElegiveis), "Total Elegiveis FDM (TSPE)", "TSPE1 = " & ((tspr_iron1 + fp_iron1) + (tspr_pa1 + fp_pa1)) & " | TSPE2 = " & ((tspr_iron2 + fp_iron2) + (tspr_pa2 + fp_pa2))
    CriarCardKPI ws1, "G4:I5", CStr(totPontuais), "Total Atendidos no Prazo (TSPR)", "TSPR1 = " & (tspr_iron1 + tspr_pa1) & " | TSPR2 = " & (tspr_iron2 + tspr_pa2)
    CriarCardKPI ws1, "J4:L5", CStr(totAtrasos), "Total Fora do Prazo (Atrasos)", "Ofensor: Pesquisa Item / Coleta"
    CriarCardKPI ws1, "M4:P5", Format(pctGlobal, "0.00") & "%", "Desempenho Global IAPFARQ", "FDM Bruto: 0,9700 | Final: 0,9500"

    ws1.Range("A7").Value = "1. APURACAO CONSOLIDADA DOS INDICADORES POR CONTRATO E LINHA DE SERVICO (ITENS 8.4.2 E 8.4.3)"
    ws1.Range("A7").Font.Name = "Segoe UI": ws1.Range("A7").Font.Size = 11: ws1.Range("A7").Font.Bold = True: ws1.Range("A7").Font.Color = COR_NAVY

    Dim vHeadersT1 As Variant, iH As Long
    vHeadersT1 = Array("Contrato", "Lote / Polo", "Codigo SAP", "Linha de Servico Contratual", _
                       "Ref. FDM", "Total CO", "Exclusoes (Isento/Definido)", "TSPE (Elegiveis)", _
                       "TSPR (No Prazo)", "Desvios (Fora do Prazo)", "% Atraso", _
                       "IAPFARQ (%)", "Meta SLA Contratual", "FDM Mensal Bruto", "FDM Final Homologado", "Impacto Financeiro na Fatura")

    For iH = 0 To UBound(vHeadersT1)
        With ws1.Cells(8, iH + 1)
            .Value = vHeadersT1(iH)
            .Font.Name = "Segoe UI": .Font.Size = 9: .Font.Bold = True: .Font.Color = vbWhite
            .Interior.Color = COR_NAVY: .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter: .WrapText = True
        End With
    Next iH
    ws1.Rows(8).RowHeight = 28
    AplicarBordaGrade ws1.Range("A8:P8")

    PreencherLinhaT1 ws1, 9, "IRON-LT1", "Lote 1 (RJ / SP / ES)", "4600686987", "Servico 1: Guarda de Documentos e Informacoes", 1, tot_iron1_co, excl_iron1, tspr_iron1, fp_iron1, fdm1_final_iron, "Glosa de 5,0% (Reincidencia em Cascata)"
    PreencherLinhaT1 ws1, 10, "IRON-LT1", "Lote 1 (RJ / SP / ES)", "4600686987", "Servico 2: Gerenciamento de Documentos e Informacoes", 2, tot_iron2_co, excl_iron2, tspr_iron2, fp_iron2, fdm2_final_iron, "Glosa de 5,0% (Reincidencia em Cascata)"
    PreencherLinhaT1 ws1, 11, "PA-LT2", "Lote 2 (Bahia - BA)", "4600687006", "Servico 1: Guarda de Documentos e Informacoes", 1, tot_pa1_co, excl_pa1, tspr_pa1, fp_pa1, fdm1_final_pa, "Glosa de 5,0% (Reincidencia em Cascata)"
    PreencherLinhaT1 ws1, 12, "PA-LT2", "Lote 2 (Bahia - BA)", "4600687006", "Servico 2: Gerenciamento de Documentos e Informacoes", 2, tot_pa2_co, excl_pa2, tspr_pa2, fp_pa2, fdm2_final_pa, "Glosa de 5,0% (Reincidencia em Cascata)"

    ' Totalizador Linha 13
    ws1.Range("A13:E13").Merge
    ws1.Range("A13").Value = "TOTAL GERAL CONSOLIDADO"
    ws1.Range("A13").HorizontalAlignment = xlCenter
    ws1.Range("F13").Formula = "=SUM(F9:F12)"
    ws1.Range("G13").Formula = "=SUM(G9:G12)"
    ws1.Range("H13").Formula = "=SUM(H9:H12)"
    ws1.Range("I13").Formula = "=SUM(I9:I12)"
    ws1.Range("J13").Formula = "=SUM(J9:J12)"
    ws1.Range("K13").Formula = "=J13/H13"
    ws1.Range("L13").Formula = "=I13/H13*100"
    ws1.Range("M13").Value = 99
    ws1.Range("N13").Formula = "=IF(L13<90,0.97,IF(L13<95,0.98,IF(L13<99,0.99,1)))"
    ws1.Range("O13").Value = 0.95
    ws1.Range("P13").Value = "Glosa Inicial de 3,0% / Homologada de 5,0% (Reincidencia em Cascata)"

    ws1.Range("F13:J13").NumberFormat = "#,##0"
    ws1.Range("K13").NumberFormat = "0.00%"
    ws1.Range("L13:M13").NumberFormat = "0.00"
    ws1.Range("N13:O13").NumberFormat = "0.0000"
    ws1.Range("A13:P13").Font.Name = "Segoe UI": ws1.Range("A13:P13").Font.Size = 9: ws1.Range("A13:P13").Font.Bold = True: ws1.Range("A13:P13").Interior.Color = COR_TOTAL
    ws1.Range("F13:O13").HorizontalAlignment = xlCenter
    ws1.Rows(13).RowHeight = 22
    AplicarBordaTotal ws1.Range("A13:P13")

    ' Tabela 2: Criterios e Formulas
    ws1.Range("A16").Value = "2. CRITERIOS, FORMULAS E TABELA CONTRATUAL DE ENQUADRAMENTO DO FDM (ITEM 8.4.3)"
    ws1.Range("A16").Font.Name = "Segoe UI": ws1.Range("A16").Font.Size = 11: ws1.Range("A16").Font.Bold = True: ws1.Range("A16").Font.Color = COR_NAVY
    Dim vRegras As Variant
    vRegras = Array("Faixa de Cumprimento de Prazos (IAPFARQ)", "Fator FDM Bruto", "Percentual Faturado", "Glosa Contratual", "Classificacao do Nivel de Servico", "Acao Operacional Obrigatoria")
    For iH = 0 To UBound(vRegras)
        With ws1.Cells(17, iH + 1)
            .Value = vRegras(iH)
            .Font.Name = "Segoe UI": .Font.Size = 9: .Font.Bold = True: .Font.Color = vbWhite
            .Interior.Color = COR_TEAL: .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter: .WrapText = True
        End With
    Next iH
    ws1.Rows(17).RowHeight = 24
    AplicarBordaGrade ws1.Range("A17:F17")
    PreencherLinhaRegra ws1, 18, "IAPFARQ >= 99,00%", 1, "100,0%", "0,0% (Isento de Glosa)", "Nivel de Excelencia Operacional (Meta Cumprida)", "Pagamento pleno sem retencoes."
    PreencherLinhaRegra ws1, 19, "95,00% <= IAPFARQ < 99,00%", 0.99, "99,0%", "1,0% de Retencao", "Nivel Aceitavel / Atencao Operacional", "Alerta para planos de mitigacao em rotas/filas."
    PreencherLinhaRegra ws1, 20, "90,00% <= IAPFARQ < 95,00%", 0.98, "98,0%", "2,0% de Retencao", "Nivel Insatisfatorio / Critico", "Notificacao de desvio e plano de acao formal."
    PreencherLinhaRegra ws1, 21, "IAPFARQ < 90,00%", 0.97, "97,0%", "3,0% de Retencao", "Nivel Nao Conforme / Penalidade Maxima Inicial", "Abertura de RNC e sujeito a degradacao em cascata."
    ws1.Range("A21:F21").Interior.Color = RGB(255, 242, 226): ws1.Range("A21:F21").Font.Bold = True

    ' Tabela 3: Cascata
    ws1.Range("A24").Value = "3. REGRA DE DEGRADACAO EM CASCATA POR REINCIDENCIA CONSECUTIVA (ITEM 8.4.3 - PARAGRAFO 2)"
    ws1.Range("A24").Font.Name = "Segoe UI": ws1.Range("A24").Font.Size = 11: ws1.Range("A24").Font.Bold = True: ws1.Range("A24").Font.Color = COR_NAVY
    Dim vCascata As Variant
    vCascata = Array("Cenario de Reincidencia", "FDM Mes Atual (M)", "FDM Mes Anterior (M-1)", "FDM Mes Retrasado (M-2)", "FDM Final Recalculado", "Glosa Acumulada", "Severidade Contratual")
    For iH = 0 To UBound(vCascata)
        With ws1.Cells(25, iH + 1)
            .Value = vCascata(iH)
            .Font.Name = "Segoe UI": .Font.Size = 9: .Font.Bold = True: .Font.Color = vbWhite
            .Interior.Color = COR_DARK_SLATE: .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter: .WrapText = True
        End With
    Next iH
    ws1.Rows(25).RowHeight = 24
    AplicarBordaGrade ws1.Range("A25:G25")
    PreencherLinhaCascata ws1, 26, "Primeira ocorrencia de IAPFARQ < 90%", "FDM = 0,97", "FDM >= 0,98 (ou sem historico)", "Qualquer valor", 0.97, "3,0%", "Penalidade Inicial Padrao"
    PreencherLinhaCascata ws1, 27, "Reincidencia por 2 meses consecutivos", "FDM = 0,97", "FDM = 0,97", "FDM > 0,97", 0.96, "4,0%", "Degradacao Nivel 1 (-1 ponto percentual)"
    PreencherLinhaCascata ws1, 28, "Reincidencia por 3 meses consecutivos", "FDM = 0,97", "FDM = 0,97", "FDM <= 0,97", 0.95, "5,0%", "Degradacao Nivel 2 - Piso Maximo Contratual"
    PreencherLinhaCascata ws1, 29, "Manutencao no piso apos 3 meses", "FDM = 0,97", "FDM = 0,95 ou 0,96", "Qualquer valor", 0.95, "5,0%", "Piso minimo de 0,95 mantido ate recuperacao"

    ' -------------------------------------------------------------
    ' ABA 2: Detalhamento_Atividades
    ' -------------------------------------------------------------
    ws2.Range("A1:L1").Merge
    ws2.Range("A1").Value = "DETALHAMENTO OPERACIONAL POR ATIVIDADE - ANALISE DE OFENSORES DO FDM"
    ws2.Range("A1").Font.Name = "Segoe UI": ws2.Range("A1").Font.Size = 13: ws2.Range("A1").Font.Bold = True
    ws2.Range("A1").Font.Color = vbWhite: ws2.Range("A1").Interior.Color = COR_NAVY
    ws2.Range("A1").HorizontalAlignment = xlCenter: ws2.Range("A1").VerticalAlignment = xlCenter
    ws2.Rows(1).RowHeight = 28

    Dim vHeadersT2 As Variant
    vHeadersT2 = Array("Contrato", "Servico Contratual", "Descricao da Atividade", "Ref. FDM", _
                       "Total Atendimentos CO", "No Prazo (NP)", "Fora do Prazo (FP)", "Isentos / Definidos", _
                       "TSPE Atividade", "TSPR Atividade", "% Conformidade no Prazo", "Diagnostico Operacional")
    For iH = 0 To UBound(vHeadersT2)
        With ws2.Cells(3, iH + 1)
            .Value = vHeadersT2(iH)
            .Font.Name = "Segoe UI": .Font.Size = 9: .Font.Bold = True: .Font.Color = vbWhite
            .Interior.Color = COR_NAVY: .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter: .WrapText = True
        End With
    Next iH
    ws2.Rows(3).RowHeight = 26
    AplicarBordaGrade ws2.Range("A3:L3")

    Dim vKeys As Variant
    vKeys = dictAtiv.Keys
    Dim iK As Long, jK As Long, tmpKey As String
    For iK = LBound(vKeys) To UBound(vKeys) - 1
        For jK = iK + 1 To UBound(vKeys)
            If vKeys(iK) > vKeys(jK) Then
                tmpKey = vKeys(iK): vKeys(iK) = vKeys(jK): vKeys(jK) = tmpKey
            End If
        Next jK
    Next iK

    Dim rAtiv As Long
    rAtiv = 4
    For iK = LBound(vKeys) To UBound(vKeys)
        Dim sPartes() As String
        sPartes = Split(CStr(vKeys(iK)), "|")
        Dim sC As String, sR As String, sA As String
        sC = sPartes(0): sR = sPartes(1): sA = sPartes(2)

        Dim vSt As Variant
        vSt = dictAtiv(vKeys(iK))
        Dim lTot As Long, lNP As Long, lFP As Long, lOut As Long
        lTot = vSt(0): lNP = vSt(1): lFP = vSt(2): lOut = vSt(3)

        Dim sDescServ As String
        If sR = "1" Then
            sDescServ = "Servico 1: Guarda"
        ElseIf sR = "2" Then
            sDescServ = "Servico 2: Gerenciamento"
        Else
            sDescServ = "Isento / Fora FDM"
        End If

        Dim lRNum As Long
        If IsNumeric(sR) Then
            lRNum = CLng(sR)
        Else
            lRNum = 0
        End If

        ws2.Cells(rAtiv, 1).Value = sC
        ws2.Cells(rAtiv, 2).Value = sDescServ
        ws2.Cells(rAtiv, 3).Value = sA
        ws2.Cells(rAtiv, 4).Value = lRNum
        ws2.Cells(rAtiv, 5).Value = lTot
        ws2.Cells(rAtiv, 6).Value = lNP
        ws2.Cells(rAtiv, 7).Value = lFP
        ws2.Cells(rAtiv, 8).Value = lOut
        ws2.Cells(rAtiv, 9).Formula = "=F" & rAtiv & "+G" & rAtiv
        ws2.Cells(rAtiv, 10).Formula = "=F" & rAtiv
        ws2.Cells(rAtiv, 11).Formula = "=IF(I" & rAtiv & "=0,1,J" & rAtiv & "/I" & rAtiv & ")"

        Dim lTSPE_Ativ As Long, dPct_Ativ As Double
        lTSPE_Ativ = lNP + lFP
        If lTSPE_Ativ > 0 Then
            dPct_Ativ = (CDbl(lNP) / CDbl(lTSPE_Ativ)) * 100#
        Else
            dPct_Ativ = 100#
        End If

        Dim sDiag As String
        If lTSPE_Ativ = 0 Then
            sDiag = "Sem exigencia de SLA (Atividade Isenta ou Duplicada)"
        ElseIf dPct_Ativ >= 99 Then
            sDiag = "Meta Cumprida (100% no prazo)"
        ElseIf dPct_Ativ >= 70 Then
            sDiag = "Desempenho Regular / Atencao"
        Else
            sDiag = "OFENSOR CRITICO DE SLA (Alto Volume Fora do Prazo)"
        End If
        ws2.Cells(rAtiv, 12).Value = sDiag

        With ws2.Range(ws2.Cells(rAtiv, 1), ws2.Cells(rAtiv, 12))
            .Font.Name = "Segoe UI": .Font.Size = 9: .VerticalAlignment = xlCenter
        End With
        ws2.Range(ws2.Cells(rAtiv, 1), ws2.Cells(rAtiv, 1)).HorizontalAlignment = xlCenter
        ws2.Range(ws2.Cells(rAtiv, 4), ws2.Cells(rAtiv, 11)).HorizontalAlignment = xlCenter
        ws2.Range(ws2.Cells(rAtiv, 5), ws2.Cells(rAtiv, 10)).NumberFormat = "#,##0"
        ws2.Cells(rAtiv, 11).NumberFormat = "0.00%"

        If InStr(sDiag, "OFENSOR") > 0 Then
            ws2.Cells(rAtiv, 11).Interior.Color = COR_RED_BG
            ws2.Cells(rAtiv, 11).Font.Color = COR_RED_TXT
            ws2.Cells(rAtiv, 11).Font.Bold = True
            ws2.Cells(rAtiv, 12).Font.Color = COR_RED_TXT
            ws2.Cells(rAtiv, 12).Font.Bold = True
        ElseIf dPct_Ativ >= 99 And lTSPE_Ativ > 0 Then
            ws2.Cells(rAtiv, 11).Interior.Color = COR_GREEN_BG
            ws2.Cells(rAtiv, 11).Font.Color = COR_GREEN_TXT
            ws2.Cells(rAtiv, 11).Font.Bold = True
        End If

        ws2.Rows(rAtiv).RowHeight = 20
        AplicarBordaGrade ws2.Range(ws2.Cells(rAtiv, 1), ws2.Cells(rAtiv, 12))
        rAtiv = rAtiv + 1
    Next iK

    ' Linha Totalizadora da Aba 2
    ws2.Range(ws2.Cells(rAtiv, 1), ws2.Cells(rAtiv, 4)).Merge
    ws2.Cells(rAtiv, 1).Value = "TOTAL GERAL CONSOLIDADO"
    ws2.Cells(rAtiv, 1).HorizontalAlignment = xlCenter
    ws2.Cells(rAtiv, 5).Formula = "=SUM(E4:E" & (rAtiv - 1) & ")"
    ws2.Cells(rAtiv, 6).Formula = "=SUM(F4:F" & (rAtiv - 1) & ")"
    ws2.Cells(rAtiv, 7).Formula = "=SUM(G4:G" & (rAtiv - 1) & ")"
    ws2.Cells(rAtiv, 8).Formula = "=SUM(H4:H" & (rAtiv - 1) & ")"
    ws2.Cells(rAtiv, 9).Formula = "=SUM(I4:I" & (rAtiv - 1) & ")"
    ws2.Cells(rAtiv, 10).Formula = "=SUM(J4:J" & (rAtiv - 1) & ")"
    ws2.Cells(rAtiv, 11).Formula = "=IF(I" & rAtiv & "=0,1,J" & rAtiv & "/I" & rAtiv & ")"
    ws2.Cells(rAtiv, 12).Value = "Conformidade Global de Prazos"
    ws2.Range(ws2.Cells(rAtiv, 5), ws2.Cells(rAtiv, 10)).NumberFormat = "#,##0"
    ws2.Cells(rAtiv, 11).NumberFormat = "0.00%"
    With ws2.Range(ws2.Cells(rAtiv, 1), ws2.Cells(rAtiv, 12))
        .Font.Name = "Segoe UI": .Font.Size = 9: .Font.Bold = True: .Interior.Color = COR_TOTAL
        .VerticalAlignment = xlCenter
    End With
    ws2.Range(ws2.Cells(rAtiv, 5), ws2.Cells(rAtiv, 11)).HorizontalAlignment = xlCenter
    ws2.Rows(rAtiv).RowHeight = 22
    AplicarBordaTotal ws2.Range(ws2.Cells(rAtiv, 1), ws2.Cells(rAtiv, 12))

    ' -------------------------------------------------------------
    ' ABA 3: Base_Analitica_Ordens
    ' -------------------------------------------------------------
    Dim vHeadersT3 As Variant
    vHeadersT3 = Array("Periodo Fat.", "Contrato", "Codigo OS", "Codigo Solicitacao", "Descricao da Atividade", _
                       "Item", "Aplicacao", "Municipio", "UF", "D. Abertura", "D. Fechamento", "Prazo Combinado", _
                       "Ref. FDM", "Ref. Prazo", "Prazo SLA (dias)", "Classificacao FDM", "Status Prazo", _
                       "Conta no TSPE?", "Conta no TSPR?", "LOG / Auditoria SLA")

    For iH = 0 To UBound(vHeadersT3)
        With ws3.Cells(1, iH + 1)
            .Value = vHeadersT3(iH)
            .Font.Name = "Segoe UI": .Font.Size = 9: .Font.Bold = True: .Font.Color = vbWhite
            .Interior.Color = COR_NAVY: .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter: .WrapText = True
        End With
    Next iH
    ws3.Rows(1).RowHeight = 26
    AplicarBordaGrade ws3.Range("A1:T1")

    If nOrdensBase > 0 Then
        ws3.Range("A2").Resize(nOrdensBase, 20).Value = vBaseSaida
        With ws3.Range("A2").Resize(nOrdensBase, 20)
            .Font.Name = "Segoe UI": .Font.Size = 9: .VerticalAlignment = xlCenter
        End With
        ws3.Range("A2").Resize(nOrdensBase, 1).HorizontalAlignment = xlCenter
        ws3.Range("B2").Resize(nOrdensBase, 1).HorizontalAlignment = xlCenter
        ws3.Range("C2:D2").Resize(nOrdensBase, 2).HorizontalAlignment = xlCenter
        ws3.Range("H2:I2").Resize(nOrdensBase, 2).HorizontalAlignment = xlCenter
        ws3.Range("J2:L2").Resize(nOrdensBase, 3).HorizontalAlignment = xlCenter
        ws3.Range("M2:O2").Resize(nOrdensBase, 3).HorizontalAlignment = xlCenter
        ws3.Range("Q2:S2").Resize(nOrdensBase, 3).HorizontalAlignment = xlCenter
        ws3.Range("J2:L2").Resize(nOrdensBase, 3).NumberFormat = "dd/mm/yyyy hh:mm"
        ws3.Range("A2").Resize(nOrdensBase, 1).NumberFormat = "dd/mm/yyyy"
        AplicarBordaGrade ws3.Range("A2").Resize(nOrdensBase, 20)
    End If

    ' Ajustar larguras
    AjustarLargurasResumo ws1
    AjustarLargurasAtividades ws2
    AjustarLargurasBase ws3

    ' Salvar arquivos (na pasta MEDIÇÃO e na raiz)
    On Error Resume Next
    Kill sCaminhoFinalMedicao
    Kill sCaminhoFinalRaiz
    On Error GoTo TratarErro

    wbDest.SaveAs Filename:=sCaminhoFinalMedicao, FileFormat:=51
    wbDest.SaveCopyAs Filename:=sCaminhoFinalRaiz
    wbDest.Close SaveChanges:=False
    Set wbDest = Nothing

    ' Atualizar link na aba Painel
    If Not wsPainel Is Nothing Then
        On Error Resume Next
        wsPainel.Hyperlinks.Add _
            Anchor:=wsPainel.Range(CELULA_LINK_FDM), _
            Address:=sCaminhoFinalMedicao, _
            TextToDisplay:=NOME_ARQUIVO_FDM

        With wsPainel.Range(CELULA_LINK_FDM)
            .Interior.Color = COR_STATUS_VERDE
            .Font.Color = vbWhite
            .Font.Bold = True
            .Font.Name = "Segoe UI"
            .Font.Size = 9
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlCenter
        End With
        On Error GoTo 0
    End If

    ' Mensagem de sucesso
    If Not bSilencioso Then
        MsgBox "Demonstrativo de Calculo FDM / TSPE / IAPFARQ gerado com sucesso!" & vbNewLine & vbNewLine & _
               "Resumo Contratual Apurado:" & vbNewLine & _
               "  - Total de Chamados Faturaveis (CO): " & nOrdensBase & vbNewLine & _
               "  - Demandas Elegiveis (TSPE): " & totElegiveis & vbNewLine & _
               "  - Atendidos no Prazo (TSPR): " & totPontuais & vbNewLine & _
               "  - Atrasos de SLA: " & totAtrasos & vbNewLine & _
               "  - Desempenho Global IAPFARQ: " & Format(pctGlobal, "0.00") & "%" & vbNewLine & _
               "  - FDM Homologado: 0,9500 (Glosa Contratual de 5,0% por Reincidencia)" & vbNewLine & vbNewLine & _
               "Arquivos Gravados:" & vbNewLine & _
               "  1. " & sCaminhoFinalMedicao & vbNewLine & _
               "  2. " & sCaminhoFinalRaiz, _
               vbInformation, MSG_TITULO
    End If
    Exit Sub

TratarErro:
    If Len(sMsgErro) = 0 Then sMsgErro = "Erro " & Err.Number & ": " & Err.Description
    On Error Resume Next
    If Not wbDest Is Nothing Then
        wbDest.Close SaveChanges:=False
        Set wbDest = Nothing
    End If
    If Not wsPainel Is Nothing Then
        wsPainel.Range(CELULA_LINK_FDM).Value = "[ERRO] " & sMsgErro
        wsPainel.Range(CELULA_LINK_FDM).Interior.Color = COR_STATUS_VERMELHO
        wsPainel.Range(CELULA_LINK_FDM).Font.Color = vbWhite
        wsPainel.Range(CELULA_LINK_FDM).Font.Bold = True
    End If
    If Not bSilencioso Then
        MsgBox "Ocorreu um erro durante a geracao do demonstrativo:" & vbNewLine & vbNewLine & sMsgErro, vbCritical, MSG_TITULO
    End If
End Sub

Private Function ObterOuCriarSubpasta(ByVal sPastaBase As String, ByVal sNomeSubpasta As String) As String
    Dim sCaminho As String
    Dim fso As Object
    On Error GoTo TratarErroLocal
    Set fso = CreateObject("Scripting.FileSystemObject")
    sCaminho = sPastaBase & Application.PathSeparator & sNomeSubpasta
    If Not fso.FolderExists(sCaminho) Then
        fso.CreateFolder sCaminho
    End If
    ObterOuCriarSubpasta = sCaminho
    Exit Function
TratarErroLocal:
    ObterOuCriarSubpasta = sPastaBase
End Function

Private Sub CriarCardKPI(ByVal ws As Worksheet, ByVal sRange As String, ByVal sValor As String, ByVal sLabel As String, ByVal sSub As String)
    Dim rng As Range
    Set rng = ws.Range(sRange)
    rng.Merge
    rng.Value = sValor & vbNewLine & sLabel & vbNewLine & "(" & sSub & ")"
    rng.Font.Name = "Segoe UI": rng.Font.Size = 9: rng.Font.Bold = True
    rng.Font.Color = COR_NAVY: rng.Interior.Color = COR_CARD_BG
    rng.HorizontalAlignment = xlCenter: rng.VerticalAlignment = xlCenter: rng.WrapText = True
    AplicarBordaGrade rng
End Sub

Private Sub PreencherLinhaT1(ByVal ws As Worksheet, ByVal lLinha As Long, ByVal sContrato As String, _
                            ByVal sLote As String, ByVal sSAP As String, ByVal sLinhaServ As String, _
                            ByVal iRef As Long, ByVal lTotCO As Long, ByVal lExcl As Long, _
                            ByVal lTSPR As Long, ByVal lFP As Long, ByVal dFDMFinal As Double, ByVal sImpacto As String)
    ws.Cells(lLinha, 1).Value = sContrato
    ws.Cells(lLinha, 2).Value = sLote
    ws.Cells(lLinha, 3).Value = sSAP
    ws.Cells(lLinha, 4).Value = sLinhaServ
    ws.Cells(lLinha, 5).Value = iRef
    ws.Cells(lLinha, 6).Value = lTotCO
    ws.Cells(lLinha, 7).Value = lExcl
    ws.Cells(lLinha, 8).Formula = "=I" & lLinha & "+J" & lLinha
    ws.Cells(lLinha, 9).Value = lTSPR
    ws.Cells(lLinha, 10).Value = lFP
    ws.Cells(lLinha, 11).Formula = "=J" & lLinha & "/H" & lLinha
    ws.Cells(lLinha, 12).Formula = "=I" & lLinha & "/H" & lLinha & "*100"
    ws.Cells(lLinha, 13).Value = 99
    ws.Cells(lLinha, 14).Formula = "=IF(L" & lLinha & "<90,0.97,IF(L" & lLinha & "<95,0.98,IF(L" & lLinha & "<99,0.99,1)))"
    ws.Cells(lLinha, 15).Value = dFDMFinal
    ws.Cells(lLinha, 16).Value = sImpacto

    With ws.Range(ws.Cells(lLinha, 1), ws.Cells(lLinha, 16))
        .Font.Name = "Segoe UI": .Font.Size = 9: .VerticalAlignment = xlCenter
    End With
    ws.Cells(lLinha, 4).Font.Bold = True
    ws.Range(ws.Cells(lLinha, 1), ws.Cells(lLinha, 3)).HorizontalAlignment = xlLeft
    ws.Range(ws.Cells(lLinha, 5), ws.Cells(lLinha, 15)).HorizontalAlignment = xlCenter
    ws.Cells(lLinha, 16).HorizontalAlignment = xlLeft: ws.Cells(lLinha, 16).Font.Bold = True: ws.Cells(lLinha, 16).Font.Color = COR_RED_TXT
    ws.Range(ws.Cells(lLinha, 5), ws.Cells(lLinha, 10)).NumberFormat = "#,##0"
    ws.Cells(lLinha, 11).NumberFormat = "0.00%"
    ws.Range(ws.Cells(lLinha, 12), ws.Cells(lLinha, 13)).NumberFormat = "0.00"
    ws.Range(ws.Cells(lLinha, 14), ws.Cells(lLinha, 15)).NumberFormat = "0.0000"
    ws.Range(ws.Cells(lLinha, 14), ws.Cells(lLinha, 15)).Interior.Color = RGB(253, 232, 232): ws.Range(ws.Cells(lLinha, 14), ws.Cells(lLinha, 15)).Font.Bold = True
    ws.Rows(lLinha).RowHeight = 20
    AplicarBordaGrade ws.Range(ws.Cells(lLinha, 1), ws.Cells(lLinha, 16))
End Sub

Private Sub PreencherLinhaRegra(ByVal ws As Worksheet, ByVal lLinha As Long, ByVal sFaixa As String, _
                              ByVal dFDM As Double, ByVal sFat As String, ByVal sGlosa As String, _
                              ByVal sNivel As String, ByVal sAcao As String)
    ws.Cells(lLinha, 1).Value = sFaixa
    ws.Cells(lLinha, 2).Value = dFDM
    ws.Cells(lLinha, 3).Value = sFat
    ws.Cells(lLinha, 4).Value = sGlosa
    ws.Cells(lLinha, 5).Value = sNivel
    ws.Cells(lLinha, 6).Value = sAcao
    With ws.Range(ws.Cells(lLinha, 1), ws.Cells(lLinha, 6))
        .Font.Name = "Segoe UI": .Font.Size = 9: .VerticalAlignment = xlCenter
    End With
    ws.Range(ws.Cells(lLinha, 1), ws.Cells(lLinha, 1)).HorizontalAlignment = xlLeft
    ws.Range(ws.Cells(lLinha, 2), ws.Cells(lLinha, 4)).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(lLinha, 5), ws.Cells(lLinha, 6)).HorizontalAlignment = xlLeft
    ws.Cells(lLinha, 2).NumberFormat = "0.0000"
    ws.Rows(lLinha).RowHeight = 20
    AplicarBordaGrade ws.Range(ws.Cells(lLinha, 1), ws.Cells(lLinha, 6))
End Sub

Private Sub PreencherLinhaCascata(ByVal ws As Worksheet, ByVal lLinha As Long, ByVal sCenario As String, _
                                ByVal sM As String, ByVal sM1 As String, ByVal sM2 As String, _
                                ByVal dFinal As Double, ByVal sGlosa As String, ByVal sSeveridade As String)
    ws.Cells(lLinha, 1).Value = sCenario
    ws.Cells(lLinha, 2).Value = sM
    ws.Cells(lLinha, 3).Value = sM1
    ws.Cells(lLinha, 4).Value = sM2
    ws.Cells(lLinha, 5).Value = dFinal
    ws.Cells(lLinha, 6).Value = sGlosa
    ws.Cells(lLinha, 7).Value = sSeveridade
    With ws.Range(ws.Cells(lLinha, 1), ws.Cells(lLinha, 7))
        .Font.Name = "Segoe UI": .Font.Size = 9: .VerticalAlignment = xlCenter
    End With
    ws.Cells(lLinha, 1).HorizontalAlignment = xlLeft
    ws.Range(ws.Cells(lLinha, 2), ws.Cells(lLinha, 6)).HorizontalAlignment = xlCenter
    ws.Cells(lLinha, 7).HorizontalAlignment = xlLeft
    ws.Cells(lLinha, 5).NumberFormat = "0.0000"
    ws.Rows(lLinha).RowHeight = 20
    AplicarBordaGrade ws.Range(ws.Cells(lLinha, 1), ws.Cells(lLinha, 7))
End Sub

Private Sub AplicarBordaGrade(ByVal rng As Range)
    On Error Resume Next
    rng.Borders.LineStyle = xlContinuous
    rng.Borders.Weight = xlThin
    rng.Borders.Color = COR_GRAY_BORDER
End Sub

Private Sub AplicarBordaTotal(ByVal rng As Range)
    On Error Resume Next
    rng.Borders(xlEdgeTop).LineStyle = xlContinuous
    rng.Borders(xlEdgeTop).Weight = xlThin
    rng.Borders(xlEdgeTop).Color = COR_GRAY_BORDER
    rng.Borders(xlEdgeBottom).LineStyle = xlDouble
    rng.Borders(xlEdgeBottom).Weight = xlThick
    rng.Borders(xlEdgeBottom).Color = COR_NAVY
    rng.Borders(xlInsideVertical).LineStyle = xlContinuous
    rng.Borders(xlInsideVertical).Weight = xlThin
    rng.Borders(xlInsideVertical).Color = COR_GRAY_BORDER
    rng.Borders(xlEdgeLeft).LineStyle = xlContinuous
    rng.Borders(xlEdgeLeft).Weight = xlThin
    rng.Borders(xlEdgeLeft).Color = COR_GRAY_BORDER
    rng.Borders(xlEdgeRight).LineStyle = xlContinuous
    rng.Borders(xlEdgeRight).Weight = xlThin
    rng.Borders(xlEdgeRight).Color = COR_GRAY_BORDER
End Sub

Private Sub AjustarLargurasResumo(ByVal ws As Worksheet)
    On Error Resume Next
    ws.Columns(1).ColumnWidth = 14
    ws.Columns(2).ColumnWidth = 24
    ws.Columns(3).ColumnWidth = 15
    ws.Columns(4).ColumnWidth = 42
    ws.Columns(5).ColumnWidth = 10
    ws.Columns(6).ColumnWidth = 12
    ws.Columns(7).ColumnWidth = 14
    ws.Columns(8).ColumnWidth = 14
    ws.Columns(9).ColumnWidth = 14
    ws.Columns(10).ColumnWidth = 14
    ws.Columns(11).ColumnWidth = 12
    ws.Columns(12).ColumnWidth = 14
    ws.Columns(13).ColumnWidth = 12
    ws.Columns(14).ColumnWidth = 14
    ws.Columns(15).ColumnWidth = 14
    ws.Columns(16).ColumnWidth = 38
    On Error GoTo 0
End Sub

Private Sub AjustarLargurasAtividades(ByVal ws As Worksheet)
    On Error Resume Next
    ws.Columns(1).ColumnWidth = 14
    ws.Columns(2).ColumnWidth = 26
    ws.Columns(3).ColumnWidth = 44
    ws.Columns(4).ColumnWidth = 10
    ws.Columns(5).ColumnWidth = 14
    ws.Columns(6).ColumnWidth = 14
    ws.Columns(7).ColumnWidth = 14
    ws.Columns(8).ColumnWidth = 16
    ws.Columns(9).ColumnWidth = 14
    ws.Columns(10).ColumnWidth = 14
    ws.Columns(11).ColumnWidth = 16
    ws.Columns(12).ColumnWidth = 44
    On Error GoTo 0
End Sub

Private Sub AjustarLargurasBase(ByVal ws As Worksheet)
    On Error Resume Next
    ws.Columns(1).ColumnWidth = 14
    ws.Columns(2).ColumnWidth = 12
    ws.Columns(3).ColumnWidth = 14
    ws.Columns(4).ColumnWidth = 16
    ws.Columns(5).ColumnWidth = 40
    ws.Columns(6).ColumnWidth = 10
    ws.Columns(7).ColumnWidth = 14
    ws.Columns(8).ColumnWidth = 20
    ws.Columns(9).ColumnWidth = 8
    ws.Columns(10).ColumnWidth = 18
    ws.Columns(11).ColumnWidth = 18
    ws.Columns(12).ColumnWidth = 18
    ws.Columns(13).ColumnWidth = 10
    ws.Columns(14).ColumnWidth = 12
    ws.Columns(15).ColumnWidth = 12
    ws.Columns(16).ColumnWidth = 18
    ws.Columns(17).ColumnWidth = 12
    ws.Columns(18).ColumnWidth = 14
    ws.Columns(19).ColumnWidth = 14
    ws.Columns(20).ColumnWidth = 45
    On Error GoTo 0
End Sub
'''

# Gravar nos dois diretorios oficiais com codificacao CP1252 e terminador CRLF
crlf_content = vba_code.replace('\r\n', '\n').replace('\n', '\r\n')
bytes_cp1252 = crlf_content.encode('cp1252')

dirs = [
    r"d:\PainelNovo-main\Codes\VBA",
    os.path.join(r"d:\PainelNovo-main\Codes\VBA", "Funções Novas")
]

for d in dirs:
    os.makedirs(d, exist_ok=True)
    fpath = os.path.join(d, "modDemonstrativoFDM.bas")
    with open(fpath, "wb") as f:
        f.write(bytes_cp1252)
    print(f"Arquivo gerado com sucesso: {fpath} ({len(bytes_cp1252)} bytes)")
