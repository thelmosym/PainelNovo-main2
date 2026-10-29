# -*- coding: utf-8 -*-
import os

with open('Codes/VBA/modDemonstrativoFDM.bas', 'r', encoding='latin1') as f:
    text = f.read()

# Replace line 268
old_line = '            vBaseSaida(nOrdensBase, 20) = vDados(r, colLogSLA)'
new_line = '''            ' Diagnostico inteligente de erros de calculo SLA
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
            vBaseSaida(nOrdensBase, 20) = sLogFinal'''

assert old_line in text, 'old_line not found'
text = text.replace(old_line, new_line, 1)

# Update column width in AjustarLargurasBase
old_width = '    ws.Columns(20).ColumnWidth = 45'
new_width = '''    ws.Columns(20).ColumnWidth = 65
    ws.Columns(20).WrapText = False'''
assert old_width in text, 'old_width not found'
text = text.replace(old_width, new_width, 1)

diagnostic_func = '''
' ==============================================================================
' FUNCAO DE AUDITORIA E DIAGNOSTICO DETALHADO DE ERROS DE CALCULO SLA / FDM
' ==============================================================================
Private Function DiagnosticarMotivoErroSLA(ByVal sLogAtual As String, _
                                         ByVal vContrato As Variant, _
                                         ByVal vDtAber As Variant, _
                                         ByVal vDtFech As Variant, _
                                         ByVal vPrazoComb As Variant, _
                                         ByVal vPrazoSLA As Variant, _
                                         ByVal vRefPrazo As Variant, _
                                         ByVal vMun As Variant, _
                                         ByVal vUF As Variant, _
                                         ByVal vAtiv As Variant) As String
    On Error GoTo TratarErro

    Dim sMotivos As String
    sMotivos = ""

    ' 1. Validacao do Contrato
    Dim sCont As String
    sCont = UCase(Trim(CStr(vContrato)))
    If sCont = "" Then
        sMotivos = sMotivos & "Contrato ausente; "
    ElseIf InStr(sCont, "IRON") = 0 And InStr(sCont, "PA") = 0 Then
        sMotivos = sMotivos & "Contrato fora dos lotes homologados (" & sCont & "); "
    End If

    ' 2. Validacao de Integridade de Datas
    Dim bTemAber As Boolean, bTemFech As Boolean
    Dim dtAber As Date, dtFech As Date
    bTemAber = False: bTemFech = False

    If IsEmpty(vDtAber) Or IsNull(vDtAber) Or Trim(CStr(vDtAber)) = "" Then
        sMotivos = sMotivos & "Data de abertura ausente; "
    ElseIf IsDate(vDtAber) Then
        bTemAber = True
        dtAber = CDate(vDtAber)
    ElseIf IsNumeric(vDtAber) Then
        bTemAber = True
        dtAber = CDate(CDbl(vDtAber))
    Else
        sMotivos = sMotivos & "Data de abertura em formato invalido (" & CStr(vDtAber) & "); "
    End If

    If IsEmpty(vDtFech) Or IsNull(vDtFech) Or Trim(CStr(vDtFech)) = "" Then
        sMotivos = sMotivos & "Data de fechamento ausente; "
    ElseIf IsDate(vDtFech) Then
        bTemFech = True
        dtFech = CDate(vDtFech)
    ElseIf IsNumeric(vDtFech) Then
        bTemFech = True
        dtFech = CDate(CDbl(vDtFech))
    Else
        sMotivos = sMotivos & "Data de fechamento em formato invalido (" & CStr(vDtFech) & "); "
    End If

    If bTemAber And bTemFech Then
        If dtFech < dtAber Then
            sMotivos = sMotivos & "Inversao cronologica: Fechamento (" & Format(dtFech, "dd/mm/yyyy hh:nn") & _
                                  ") anterior a Abertura (" & Format(dtAber, "dd/mm/yyyy hh:nn") & "); "
        End If
    End If

    ' 3. Validacao de Regra Territorial e Parametrizacao SLA
    Dim sRefP As String, sMun As String, sUF As String
    sRefP = UCase(Trim(CStr(vRefPrazo)))
    sMun = Trim(CStr(vMun))
    sUF = UCase(Trim(CStr(vUF)))

    If sRefP = "TABELA" Then
        If sMun = "" Then sMotivos = sMotivos & "Municipio ausente para apuracao regional; "
        If sUF = "" Then sMotivos = sMotivos & "UF ausente para apuracao regional; "

        If IsEmpty(vPrazoSLA) Or IsNull(vPrazoSLA) Or Trim(CStr(vPrazoSLA)) = "" Then
            sMotivos = sMotivos & "Prazo SLA contratual nao parametrizado (nulo); "
        ElseIf IsNumeric(vPrazoSLA) Then
            If CDbl(vPrazoSLA) <= 0 Then
                sMotivos = sMotivos & "Prazo SLA contratual zerado para a atividade; "
            End If
        Else
            sMotivos = sMotivos & "Prazo SLA nao numerico (" & CStr(vPrazoSLA) & "); "
        End If
    End If

    ' 4. Deteccao de Travessia de Feriados e Calendario
    If bTemAber And bTemFech Then
        Dim sFeriadosAtravessados As String
        sFeriadosAtravessados = ""

        Dim dIni As Date, dFim As Date
        dIni = DateSerial(Year(dtAber), Month(dtAber), Day(dtAber))
        dFim = DateSerial(Year(dtFech), Month(dtFech), Day(dtFech))

        ' Estende horizonte caso o prazo SLA se projete alem da data de fechamento
        Dim dLimiteEstendido As Date
        dLimiteEstendido = dFim
        If IsNumeric(vPrazoSLA) Then
            Dim dPrazoDias As Long
            dPrazoDias = CLng(vPrazoSLA)
            If dPrazoDias > 0 Then
                Dim dEstimada As Date
                dEstimada = dIni + dPrazoDias + 4 ' margem para fins de semana
                If dEstimada > dLimiteEstendido Then dLimiteEstendido = dEstimada
            End If
        End If

        ' Checagem dos feriados homologados
        Dim vFeriados As Variant, iF As Long
        vFeriados = Array( _
            Array(DateSerial(2026, 1, 1), "01/01/2026 (Confraternizacao Universal)", "TODOS"), _
            Array(DateSerial(2026, 2, 16), "16/02/2026 (Carnaval - Segunda)", "TODOS"), _
            Array(DateSerial(2026, 2, 17), "17/02/2026 (Carnaval - Terca)", "TODOS"), _
            Array(DateSerial(2026, 4, 3), "03/04/2026 (Sexta-feira Santa)", "TODOS"), _
            Array(DateSerial(2026, 4, 21), "21/04/2026 (Tiradentes)", "TODOS"), _
            Array(DateSerial(2026, 4, 23), "23/04/2026 (Dia de Sao Jorge)", "RJ"), _
            Array(DateSerial(2026, 5, 1), "01/05/2026 (Dia do Trabalho)", "TODOS"), _
            Array(DateSerial(2026, 6, 4), "04/06/2026 (Corpus Christi)", "TODOS"), _
            Array(DateSerial(2026, 7, 2), "02/07/2026 (Independencia da Bahia)", "BA"), _
            Array(DateSerial(2026, 7, 9), "09/07/2026 (Revolucao Constitucionalista)", "SP"), _
            Array(DateSerial(2026, 9, 7), "07/09/2026 (Independencia do Brasil)", "TODOS"), _
            Array(DateSerial(2026, 10, 12), "12/10/2026 (Nossa Senhora Aparecida)", "TODOS"), _
            Array(DateSerial(2026, 11, 2), "02/11/2026 (Finados)", "TODOS"), _
            Array(DateSerial(2026, 11, 15), "15/11/2026 (Proclamacao da Republica)", "TODOS"), _
            Array(DateSerial(2026, 12, 25), "25/12/2026 (Natal)", "TODOS") _
        )

        For iF = 0 To UBound(vFeriados)
            Dim dtFeriado As Date, sNomeFeriado As String, sUFFeriado As String
            dtFeriado = CDate(vFeriados(iF)(0))
            sNomeFeriado = CStr(vFeriados(iF)(1))
            sUFFeriado = CStr(vFeriados(iF)(2))

            If sUFFeriado = "TODOS" Or sUFFeriado = sUF Then
                If dtFeriado >= dIni And dtFeriado <= dLimiteEstendido Then
                    If sFeriadosAtravessados <> "" Then sFeriadosAtravessados = sFeriadosAtravessados & ", "
                    sFeriadosAtravessados = sFeriadosAtravessados & sNomeFeriado
                End If
            End If
        Next iF

        If sFeriadosAtravessados <> "" Then
            sMotivos = sMotivos & "Travessia de feriado " & sFeriadosAtravessados & _
                       " com incompatibilidade de esquema na tabela Calendario (coluna Municipio ausente na Tabela B); "
        End If
    End If

    ' 5. Se nenhum motivo especifico foi mapeado, relatar dados de contexto
    If sMotivos = "" Then
        If InStr(1, sLogAtual, "Erro no calculo", vbTextCompare) > 0 Then
            Dim sDetalheOrig As String
            sDetalheOrig = Replace(sLogAtual, "Erro no calculo - verificar dados da linha;", "")
            sDetalheOrig = Replace(sDetalheOrig, "Erro no calculo - verificar dados da linha", "")
            sDetalheOrig = Replace(sDetalheOrig, "Erro no calculo;", "")
            sDetalheOrig = Replace(sDetalheOrig, "Erro no calculo", "")
            sDetalheOrig = Trim(sDetalheOrig)
            If sDetalheOrig <> "" Then
                sMotivos = sDetalheOrig & "; "
            Else
                sMotivos = "Inconsistencia no processamento do motor de SLA para os parametros da linha; "
            End If
        Else
            sMotivos = "Status Prazo nao preenchido na base analitica; "
        End If
    End If

    ' Monta retorno formatado com parametros da OS
    Dim sParametros As String
    sParametros = " | Parametros: "
    If bTemAber Then sParametros = sParametros & "Abertura: " & Format(dtAber, "dd/mm/yyyy hh:nn") Else sParametros = sParametros & "Abertura: " & CStr(vDtAber)
    If bTemFech Then sParametros = sParametros & " | Fechamento: " & Format(dtFech, "dd/mm/yyyy hh:nn") Else sParametros = sParametros & " | Fechamento: " & CStr(vDtFech)
    sParametros = sParametros & " | Prazo SLA: " & CStr(vPrazoSLA) & " dia(s)"
    If sMun <> "" Or sUF <> "" Then sParametros = sParametros & " | Local: " & sMun & "-" & sUF

    DiagnosticarMotivoErroSLA = "Erro no calculo: " & sMotivos & sParametros & "; "
    Exit Function

TratarErro:
    DiagnosticarMotivoErroSLA = "Erro no calculo: falha no diagnostico interno (" & Err.Description & "); "
End Function
'''

text = text + '\n' + diagnostic_func

# Normalize to CRLF
lines_out = [line.rstrip('\r\n') for line in text.split('\n')]
crlf_text = '\r\n'.join(lines_out) + '\r\n'

with open('Codes/VBA/modDemonstrativoFDM.bas', 'wb') as f:
    f.write(crlf_text.encode('latin1'))

funcoes_novas_path = os.path.join('Codes', 'VBA', 'Funções Novas', 'modDemonstrativoFDM.bas')
if os.path.exists(os.path.dirname(funcoes_novas_path)):
    with open(funcoes_novas_path, 'wb') as f:
        f.write(crlf_text.encode('latin1'))

print('Successfully updated modDemonstrativoFDM.bas with Windows-1252 and CRLF!')

