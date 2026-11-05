import subprocess
import os
import re

# Read clean baseline from commit f72cb3c
base_text = subprocess.check_output(['git', 'show', 'f72cb3c:Codes/VBA/modDemonstrativoFDM.bas'], encoding='latin1')
print("Baseline length:", len(base_text))

# 1. Start of ExecutarGeracaoFDM: add DisplayAlerts, ScreenUpdating, EnableEvents
old_start = """Private Sub ExecutarGeracaoFDM(ByVal bSilencioso As Boolean)
    Dim sMsgErro As String
    sMsgErro = ""
    On Error GoTo TratarErro"""

new_start = """Private Sub ExecutarGeracaoFDM(ByVal bSilencioso As Boolean)
    Dim sMsgErro As String
    sMsgErro = ""
    On Error GoTo TratarErro

    On Error Resume Next
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    On Error GoTo TratarErro"""

assert old_start in base_text, "old_start not found"
code = base_text.replace(old_start, new_start, 1)

# 2. Inside If sSit = "CO" Then: add resolution logic right after sAtiv = ...
old_co_head = """        If sSit = "CO" Then
            Dim sCont As String, sRefF As String, sStPrazo As String, sAtiv As String
            sCont = UCase(Trim(CStr(vDados(r, colContrato))))
            sRefF = Trim(CStr(vDados(r, colRefFDM)))
            sStPrazo = UCase(Trim(CStr(vDados(r, colStatusPrazo))))
            sAtiv = Trim(CStr(vDados(r, colAtividade)))"""

new_co_head = """        If sSit = "CO" Then
            Dim sCont As String, sRefF As String, sStPrazo As String, sAtiv As String
            sCont = UCase(Trim(CStr(vDados(r, colContrato))))
            sRefF = Trim(CStr(vDados(r, colRefFDM)))
            sStPrazo = UCase(Trim(CStr(vDados(r, colStatusPrazo))))
            sAtiv = Trim(CStr(vDados(r, colAtividade)))

            ' Resolucao automatica de SLA e Diagnostico inteligente
            Dim sLogOrig As String, sLogFinal As String
            Dim sClassifResolv As String, sStatusResolv As String, sLogResolv As String
            sLogOrig = Trim(CStr(vDados(r, colLogSLA)))
            sClassifResolv = Trim(CStr(vDados(r, colClassifFDM)))
            sStatusResolv = sStPrazo

            If InStr(1, sLogOrig, "Erro no calculo", vbTextCompare) > 0 Or _
               (sStPrazo = "" And (sRefF = "1" Or sRefF = "2") And UCase(Trim(CStr(vDados(r, colRefPrazo)))) = "TABELA" And UCase(Trim(CStr(vDados(r, colClassifFDM)))) <> "FDM DEFINIDO") Then
                
                ResolverSLAContingencia vDados(r, colDtAbertura), vDados(r, colDtFechamento), _
                                        vDados(r, colPrazoSLA), vDados(r, colUF), _
                                        vDados(r, colAtividade), vDados(r, colAplicacao), _
                                        sClassifResolv, sStatusResolv, sLogResolv
                
                If sStatusResolv <> "" Then
                    sStPrazo = sStatusResolv
                    sLogFinal = ""
                    ' Atualiza matriz em memoria
                    vDados(r, colClassifFDM) = sClassifResolv
                    vDados(r, colStatusPrazo) = sStatusResolv
                    vDados(r, colLogSLA) = ""
                    ' Sincroniza celulas na aba Painel T2M
                    On Error Resume Next
                    wsT2M.Cells(r, colClassifFDM).Value = sClassifResolv
                    wsT2M.Cells(r, colStatusPrazo).Value = sStatusResolv
                    wsT2M.Cells(r, colLogSLA).Value = ""
                    On Error GoTo TratarErro
                Else
                    sLogFinal = DiagnosticarMotivoErroSLA(sLogOrig, _
                                                         vDados(r, colContrato), _
                                                         vDados(r, colDtAbertura), _
                                                         vDados(r, colDtFechamento), _
                                                         vDados(r, colPrazoCombinado), _
                                                         vDados(r, colPrazoSLA), _
                                                         vDados(r, colRefPrazo), _
                                                         vDados(r, colMunicipio), _
                                                         vDados(r, colUF), _
                                                         vDados(r, colAtividade))
                End If
            Else
                sLogFinal = sLogOrig
            End If"""

assert old_co_head in code, "old_co_head not found"
code = code.replace(old_co_head, new_co_head, 1)

# 3. In vBaseSaida: replace the old redundant diagnostic block
old_base_tail = """            vBaseSaida(nOrdensBase, 16) = vDados(r, colClassifFDM)
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
            ' Diagnostico inteligente de erros de calculo SLA
            Dim sLogOrig As String, sLogFinal As String
            sLogOrig = Trim(CStr(vDados(r, colLogSLA)))

            If InStr(1, sLogOrig, "Erro no calculo", vbTextCompare) > 0 Or _
               (sStPrazo = "" And (sRefF = "1" Or sRefF = "2") And UCase(Trim(CStr(vDados(r, colRefPrazo)))) = "TABELA") Then
                sLogFinal = DiagnosticarMotivoErroSLA(sLogOrig, _
                                                     vDados(r, colContrato), _
                                                     vDados(r, colDtAbertura), _
                                                     vDados(r, colDtFechamento), _
                                                     vDados(r, colPrazoCombinado), _
                                                     vDados(r, colPrazoSLA), _
                                                     vDados(r, colRefPrazo), _
                                                     vDados(r, colMunicipio), _
                                                     vDados(r, colUF), _
                                                     vDados(r, colAtividade))
            Else
                sLogFinal = sLogOrig
            End If

            vBaseSaida(nOrdensBase, 18) = bContaTSPE
            vBaseSaida(nOrdensBase, 19) = bContaTSPR
            vBaseSaida(nOrdensBase, 20) = sLogFinal"""

new_base_tail = """            vBaseSaida(nOrdensBase, 16) = sClassifResolv
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
            vBaseSaida(nOrdensBase, 20) = sLogFinal"""

assert old_base_tail in code, "old_base_tail not found"
code = code.replace(old_base_tail, new_base_tail, 1)

# 4. Save file block: ensure alerts are suppressed and existing files deleted
old_save = """    ' Salvar arquivos (na pasta MEDIÇÃO e na raiz)
    On Error Resume Next
    Kill sCaminhoFinalMedicao
    Kill sCaminhoFinalRaiz
    On Error GoTo TratarErro

    wbDest.SaveAs Filename:=sCaminhoFinalMedicao, FileFormat:=51
    wbDest.SaveCopyAs Filename:=sCaminhoFinalRaiz
    wbDest.Close SaveChanges:=False
    Set wbDest = Nothing"""

new_save = """    ' Salvar arquivos (na pasta MEDICAO e na raiz)
    On Error Resume Next
    Application.DisplayAlerts = False
    If Dir(sCaminhoFinalMedicao) <> "" Then Kill sCaminhoFinalMedicao
    If Dir(sCaminhoFinalRaiz) <> "" Then Kill sCaminhoFinalRaiz
    On Error GoTo TratarErro

    wbDest.SaveAs Filename:=sCaminhoFinalMedicao, FileFormat:=51
    wbDest.SaveCopyAs Filename:=sCaminhoFinalRaiz
    wbDest.Close SaveChanges:=False
    Set wbDest = Nothing
    Application.DisplayAlerts = True"""

assert old_save in code, "old_save not found"
code = code.replace(old_save, new_save, 1)

# 5. Restore screenupdating/alerts at success exit
old_exit = """    If Not bSilencioso Then
        MsgBox "Demonstrativo de Calculo FDM / TSPE / IAPFARQ gerado com sucesso!\""""

new_exit = """    On Error Resume Next
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True

    If Not bSilencioso Then
        MsgBox "Demonstrativo de Calculo FDM / TSPE / IAPFARQ gerado com sucesso!\""""

assert old_exit in code, "old_exit not found"
code = code.replace(old_exit, new_exit, 1)

# 6. Restore screenupdating/alerts at TratarErro
old_err = """TratarErro:
    If Len(sMsgErro) = 0 Then sMsgErro = "Erro " & Err.Number & ": " & Err.Description"""

new_err = """TratarErro:
    On Error Resume Next
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    If Len(sMsgErro) = 0 Then sMsgErro = "Erro " & Err.Number & ": " & Err.Description"""

assert old_err in code, "old_err not found"
code = code.replace(old_err, new_err, 1)

# 7. Append helper functions at the end of the file
contingency_helpers = """

' ==============================================================================
' MOTOR DE CALCULO DE CONTINGENCIA SLA / FERIADOS CONTRATUAIS PETROBRAS
' ==============================================================================
Private Function EhDiaUtilContingencia(ByVal dData As Date, ByVal sUF As String) As Boolean
    Dim iDow As Integer
    iDow = Weekday(dData, vbMonday)
    If iDow >= 6 Then
        EhDiaUtilContingencia = False
        Exit Function
    End If

    Dim dOnly As Date
    dOnly = DateSerial(Year(dData), Month(dData), Day(dData))

    ' Feriados Nacionais homologados
    Select Case dOnly
        Case DateSerial(Year(dData), 1, 1): EhDiaUtilContingencia = False: Exit Function
        Case DateSerial(Year(dData), 4, 21): EhDiaUtilContingencia = False: Exit Function
        Case DateSerial(Year(dData), 5, 1): EhDiaUtilContingencia = False: Exit Function
        Case DateSerial(Year(dData), 9, 7): EhDiaUtilContingencia = False: Exit Function
        Case DateSerial(Year(dData), 10, 12): EhDiaUtilContingencia = False: Exit Function
        Case DateSerial(Year(dData), 11, 2): EhDiaUtilContingencia = False: Exit Function
        Case DateSerial(Year(dData), 11, 15): EhDiaUtilContingencia = False: Exit Function
        Case DateSerial(Year(dData), 12, 25): EhDiaUtilContingencia = False: Exit Function
    End Select

    ' Feriados Estaduais
    If sUF = "RJ" And dOnly = DateSerial(Year(dData), 4, 23) Then EhDiaUtilContingencia = False: Exit Function
    If sUF = "BA" And dOnly = DateSerial(Year(dData), 7, 2) Then EhDiaUtilContingencia = False: Exit Function
    If sUF = "SP" And dOnly = DateSerial(Year(dData), 7, 9) Then EhDiaUtilContingencia = False: Exit Function

    EhDiaUtilContingencia = True
End Function

Private Function CalcularLimiteSLAContingencia(ByVal dtAber As Date, ByVal dPrazoDias As Double, ByVal sUF As String) As Date
    Dim dMinRestantes As Double
    dMinRestantes = dPrazoDias * 8# * 60#

    Dim curDate As Date, h As Double
    curDate = DateSerial(Year(dtAber), Month(dtAber), Day(dtAber))
    h = CDbl(Hour(dtAber)) + CDbl(Minute(dtAber)) / 60# + CDbl(Second(dtAber)) / 3600#

    Do While Not EhDiaUtilContingencia(curDate, sUF)
        curDate = curDate + 1
        h = 8#
    Loop

    If h < 8# Then
        h = 8#
    ElseIf h >= 17# Then
        curDate = curDate + 1
        Do While Not EhDiaUtilContingencia(curDate, sUF)
            curDate = curDate + 1
        Loop
        h = 8#
    ElseIf h >= 12# And h < 13# Then
        h = 13#
    End If

    Dim dispManha As Double, dispTarde As Double
    Do While dMinRestantes > 0
        If Not EhDiaUtilContingencia(curDate, sUF) Then
            curDate = curDate + 1
            h = 8#
        Else
            If h < 12# Then
                dispManha = (12# - h) * 60#
                If dMinRestantes <= dispManha Then
                    h = h + dMinRestantes / 60#
                    dMinRestantes = 0
                    Exit Do
                Else
                    dMinRestantes = dMinRestantes - dispManha
                    h = 13#
                End If
            End If

            If h >= 13# And h < 17# Then
                dispTarde = (17# - h) * 60#
                If dMinRestantes <= dispTarde Then
                    h = h + dMinRestantes / 60#
                    dMinRestantes = 0
                    Exit Do
                Else
                    dMinRestantes = dMinRestantes - dispTarde
                    curDate = curDate + 1
                    h = 8#
                    Do While Not EhDiaUtilContingencia(curDate, sUF)
                        curDate = curDate + 1
                    Loop
                End If
            End If
        End If
    Loop

    Dim iHour As Long, iMin As Long
    iHour = Int(h)
    iMin = CLng(Round((h - CDbl(iHour)) * 60#, 0))
    If iMin >= 60 Then
        iHour = iHour + 1
        iMin = 0
    End If

    CalcularLimiteSLAContingencia = DateSerial(Year(curDate), Month(curDate), Day(curDate)) + TimeSerial(iHour, iMin, 0)
End Function

Private Sub ResolverSLAContingencia(ByVal vDtAber As Variant, ByVal vDtFech As Variant, _
                                   ByVal vPrazoSLA As Variant, ByVal vUF As Variant, _
                                   ByVal vAtiv As Variant, ByVal vAplic As Variant, _
                                   ByRef sOutClassif As String, ByRef sOutStatus As String, _
                                   ByRef sOutLog As String)
    On Error GoTo TratarErroLocal

    Dim dtAber As Date, dtFech As Date, dPrazo As Double
    Dim sUF As String, sAtiv As String, sAplic As String

    If IsDate(vDtAber) Then
        dtAber = CDate(vDtAber)
    ElseIf IsNumeric(vDtAber) Then
        dtAber = CDate(CDbl(vDtAber))
    Else
        sOutClassif = "": sOutStatus = "": sOutLog = "Erro: Data de abertura invalida": Exit Sub
    End If

    If IsDate(vDtFech) Then
        dtFech = CDate(vDtFech)
    ElseIf IsNumeric(vDtFech) Then
        dtFech = CDate(CDbl(vDtFech))
    Else
        sOutClassif = "": sOutStatus = "": sOutLog = "Erro: Data de fechamento invalida": Exit Sub
    End If

    If IsNumeric(vPrazoSLA) Then
        dPrazo = CDbl(vPrazoSLA)
    Else
        dPrazo = 0
    End If

    If dPrazo <= 0 Then
        sOutClassif = "": sOutStatus = "": sOutLog = "Erro: Prazo SLA nulo ou zero": Exit Sub
    End If

    sUF = UCase(Trim(CStr(vUF)))
    sAtiv = UCase(Trim(CStr(vAtiv)))
    sAplic = UCase(Trim(CStr(vAplic)))

    Dim dtLimite As Date, bDentro As Boolean
    dtLimite = CalcularLimiteSLAContingencia(dtAber, dPrazo, sUF)
    bDentro = (dtFech <= dtLimite)

    Dim bExpresso As Boolean
    bExpresso = (InStr(sAtiv, "EXPRESSO") > 0)

    If bDentro Then
        If bExpresso Then
            If sAplic = "IPF" Then
                sOutClassif = "Normal": sOutStatus = "FP"
            Else
                sOutClassif = "Expresso": sOutStatus = "NP"
            End If
        Else
            sOutClassif = "Normal": sOutStatus = "NP"
        End If
    Else
        If sAplic = "ABF" Or sAplic = "RFT" Then
            sOutClassif = IIf(bExpresso, "Expresso", "Normal"): sOutStatus = "NP"
        Else
            sOutClassif = "Normal": sOutStatus = "FP"
        End If
    End If

    sOutLog = ""
    Exit Sub

TratarErroLocal:
    sOutClassif = "": sOutStatus = "": sOutLog = "Erro no calculo: " & Err.Description
End Sub
"""

final_code = code.rstrip() + contingency_helpers + "\n"

# Verify no duplicate procedures
proc_pattern = re.compile(r'^(?:Public |Private )?(?:Sub|Function) (\w+)', re.MULTILINE)
matches = proc_pattern.findall(final_code)
seen = set()
duplicates = []
for m in matches:
    if m in seen:
        duplicates.append(m)
    seen.add(m)

print("Procedures found:", len(matches))
print("Distinct:", len(seen))
if duplicates:
    print("WARNING: Duplicate procedures found:", duplicates)
    exit(1)
else:
    print("NO duplicate procedures! Verification passed.")

# Ensure CRLF line endings
final_crlf = final_code.replace('\r\n', '\n').replace('\n', '\r\n')

# Write to both target paths with windows-1252
paths = [
    r'Codes\VBA\modDemonstrativoFDM.bas',
    r'Codes\VBA\Funções Novas\modDemonstrativoFDM.bas'
]

for p in paths:
    os.makedirs(os.path.dirname(p), exist_ok=True)
    with open(p, 'wb') as f:
        f.write(final_crlf.encode('windows-1252'))
    print(f"Successfully wrote {p} ({len(final_crlf)} chars, Windows-1252, CRLF)")

