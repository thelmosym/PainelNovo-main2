import os

# Let's test a very clean, modular implementation of modDemonstrativoFDM.bas
# where everything is super simple and directly testable.

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

Public Sub TestePasso1()
    LogDebug "TestePasso1 inicio"
    Dim wsT2M As Worksheet, wsFDM As Worksheet
    Set wsT2M = ThisWorkbook.Worksheets("Painel T2M")
    Set wsFDM = ThisWorkbook.Worksheets("TabelaFDMPorContrato")
    LogDebug "T2M rows: " & wsT2M.UsedRange.Rows.Count & " | FDM rows: " & wsFDM.UsedRange.Rows.Count
    
    Dim vDados As Variant
    vDados = wsT2M.UsedRange.Value
    LogDebug "vDados lido com sucesso. Linhas: " & UBound(vDados, 1) & " Cols: " & UBound(vDados, 2)
End Sub
'''

with open(r'd:\PainelNovo-main\Codes\VBA\modDemonstrativoFDM.bas', 'wb') as f:
    f.write(code.replace('\n', '\r\n').encode('cp1252'))

print("modDemonstrativoFDM.bas updated with TestePasso1")

