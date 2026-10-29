import os

code = r'''Attribute VB_Name = "modDemonstrativoFDM"
Option Explicit

Private Sub LogDebug(ByVal sMsg As String)
    On Error Resume Next
    Dim nF As Integer
    nF = FreeFile
    Open "d:\PainelNovo-main\scratch\vba_debug.log" For Append As #nF
    Print #nF, Format(Now, "hh:mm:ss") & " - " & sMsg
    Close #nF
End Sub

Public Sub TestePasso2()
    LogDebug "TestePasso2 inicio"
    Dim wsT2M As Worksheet, wsFDM As Worksheet
    Set wsT2M = ThisWorkbook.Worksheets("Painel T2M")
    Set wsFDM = ThisWorkbook.Worksheets("TabelaFDMPorContrato")
    
    Dim vDados As Variant
    vDados = wsT2M.UsedRange.Value
    Dim nLinhas As Long, nCols As Long
    nLinhas = UBound(vDados, 1)
    nCols = UBound(vDados, 2)
    
    Dim colSituacao As Long, colContrato As Long, colRefFDM As Long, colStatusPrazo As Long, colAtividade As Long
    colSituacao = 2: colContrato = 3: colAtividade = 4: colRefFDM = 34: colStatusPrazo = 46
    
    Dim tot_iron1_co As Long, excl_iron1 As Long, tspr_iron1 As Long, fp_iron1 As Long
    Dim tot_iron2_co As Long, excl_iron2 As Long, tspr_iron2 As Long, fp_iron2 As Long
    Dim tot_pa1_co As Long, excl_pa1 As Long, tspr_pa1 As Long, fp_pa1 As Long
    Dim tot_pa2_co As Long, excl_pa2 As Long, tspr_pa2 As Long, fp_pa2 As Long
    
    Dim dictAtiv As Object
    Set dictAtiv = CreateObject("Scripting.Dictionary")
    
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
            
            If InStr(sCont, "IRON") > 0 Then sCont = "IRON-LT1" Else If InStr(sCont, "PA") > 0 Then sCont = "PA-LT2"
            
            If sCont = "IRON-LT1" Then
                If sRefF = "1" Then
                    tot_iron1_co = tot_iron1_co + 1
                    If sStPrazo = "NP" Then tspr_iron1 = tspr_iron1 + 1 Else If sStPrazo = "FP" Then fp_iron1 = fp_iron1 + 1 Else excl_iron1 = excl_iron1 + 1
                ElseIf sRefF = "2" Then
                    tot_iron2_co = tot_iron2_co + 1
                    If sStPrazo = "NP" Then tspr_iron2 = tspr_iron2 + 1 Else If sStPrazo = "FP" Then fp_iron2 = fp_iron2 + 1 Else excl_iron2 = excl_iron2 + 1
                End If
            ElseIf sCont = "PA-LT2" Then
                If sRefF = "1" Then
                    tot_pa1_co = tot_pa1_co + 1
                    If sStPrazo = "NP" Then tspr_pa1 = tspr_pa1 + 1 Else If sStPrazo = "FP" Then fp_pa1 = fp_pa1 + 1 Else excl_pa1 = excl_pa1 + 1
                ElseIf sRefF = "2" Then
                    tot_pa2_co = tot_pa2_co + 1
                    If sStPrazo = "NP" Then tspr_pa2 = tspr_pa2 + 1 Else If sStPrazo = "FP" Then fp_pa2 = fp_pa2 + 1 Else excl_pa2 = excl_pa2 + 1
                End If
            End If
            
            Dim sKeyAtiv As String
            sKeyAtiv = sCont & "|" & sRefF & "|" & sAtiv
            Dim aStats As Variant
            If Not dictAtiv.Exists(sKeyAtiv) Then
                dictAtiv.Add sKeyAtiv, Array(1, IIf(sStPrazo = "NP", 1, 0), IIf(sStPrazo = "FP", 1, 0), IIf(sStPrazo <> "NP" And sStPrazo <> "FP", 1, 0))
            Else
                aStats = dictAtiv(sKeyAtiv)
                aStats(0) = aStats(0) + 1
                If sStPrazo = "NP" Then aStats(1) = aStats(1) + 1
                If sStPrazo = "FP" Then aStats(2) = aStats(2) + 1
                If sStPrazo <> "NP" And sStPrazo <> "FP" Then aStats(3) = aStats(3) + 1
                dictAtiv(sKeyAtiv) = aStats
            End If
        End If
    Next r
    
    LogDebug "TestePasso2 fim. IRON1=" & tot_iron1_co & " TSPR=" & tspr_iron1 & " FP=" & fp_iron1
    LogDebug "Chaves em dictAtiv: " & dictAtiv.Count
End Sub
'''

with open(r'd:\PainelNovo-main\Codes\VBA\modDemonstrativoFDM.bas', 'wb') as f:
    f.write(code.replace('\n', '\r\n').encode('cp1252'))

print("modDemonstrativoFDM.bas updated with TestePasso2")

