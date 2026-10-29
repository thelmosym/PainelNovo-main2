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

Public Sub TestePasso3()
    LogDebug "TestePasso3 inicio"
    Dim wbDest As Workbook
    Set wbDest = Application.Workbooks.Add(xlWBATWorksheet)
    LogDebug "wbDest criado: " & wbDest.Name
    
    Dim ws1 As Worksheet
    Set ws1 = wbDest.Worksheets(1)
    ws1.Name = "Resumo_Executivo_FDM"
    ws1.Range("A1").Value = "Teste Titulo"
    
    Dim sPathDest As String
    sPathDest = "d:\PainelNovo-main\scratch\test_dest.xlsx"
    On Error Resume Next
    Kill sPathDest
    On Error GoTo 0
    
    LogDebug "Salvando test_dest.xlsx..."
    wbDest.SaveAs Filename:=sPathDest, FileFormat:=51
    LogDebug "Salvo! Fechando wbDest..."
    wbDest.Close SaveChanges:=False
    LogDebug "TestePasso3 fim com sucesso!"
End Sub
'''

with open(r'd:\PainelNovo-main\Codes\VBA\modDemonstrativoFDM.bas', 'wb') as f:
    f.write(code.replace('\n', '\r\n').encode('cp1252'))

print("modDemonstrativoFDM.bas updated with TestePasso3")

