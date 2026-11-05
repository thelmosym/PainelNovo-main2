# -*- coding: cp1252 -*-
import os

with open(r'Codes\VBA\modAnaliseMonitoramento.bas', 'rb') as f:
    raw = f.read()

text = raw.decode('cp1252').replace('\r\n', '\n')
print(f"Loaded modAnaliseMonitoramento.bas: {len(text)} chars, {text.count(chr(10))} lines")

# 1. Update entry point to support silent mode
target_entry = """Public Sub AnalisarErrosMonitoramentoEGerarLog()
    Dim wsMonit As Worksheet"""

repl_entry = """Public Sub AnalisarErrosMonitoramentoEGerarLog()
    ExecutarAnaliseMonitoramento False
End Sub

Public Sub AnalisarErrosMonitoramento_Silencioso()
    ExecutarAnaliseMonitoramento True
End Sub

Public Sub analiseMonitoramento()
    ExecutarAnaliseMonitoramento False
End Sub

Private Sub ExecutarAnaliseMonitoramento(ByVal bSilencioso As Boolean)
    Dim wsMonit As Worksheet"""

assert target_entry in text, "target_entry not found"
text = text.replace(target_entry, repl_entry, 1)
print("1. Entry point updated.")

# 2. Wrap MsgBox confirmation
target_msg_start = 'If MsgBox("Será executada a análise de conformidade nos registros da aba Monitoramento."'
pos_m_start = text.find(target_msg_start)
assert pos_m_start != -1, "target_msg_start not found"
target_msg_end = 'vbYesNo + vbQuestion, MSG_TITULO) = vbNo Then\n        Exit Sub\n    End If'
pos_m_end = text.find(target_msg_end, pos_m_start)
assert pos_m_end != -1, "target_msg_end not found"
full_msg_end = pos_m_end + len(target_msg_end)

original_msg_block = text[pos_m_start:full_msg_end]
wrapped_msg_block = """If Not bSilencioso Then
        If MsgBox("Será executada a análise de conformidade nos registros da aba Monitoramento." & vbNewLine & vbNewLine & _
                  "  Regras de Auditoria:" & vbNewLine & _
                  "   - Linhas com Situação 'CO' (Concluído) ou 'AT' (Atendido);" & vbNewLine & _
                  "   - Linhas com 'OS disponibilizada = Sim' e/ou data de fechamento informada." & vbNewLine & vbNewLine & _
                  "  Destino: Os erros encontrados serão gravados na Coluna Z (LOG) da respectiva linha." & vbNewLine & vbNewLine & _
                  "Total de registros a varrer: " & Format(iTotalLinhas, "#,##0") & vbNewLine & vbNewLine & _
                  "Deseja iniciar a auditoria?", _
                  vbYesNo + vbQuestion, MSG_TITULO) = vbNo Then
            Exit Sub
        End If
    End If"""

text = text[:pos_m_start] + wrapped_msg_block + text[full_msg_end:]
print("2. MsgBox confirmation wrapped.")

# 3. Wrap ExibirResumoAuditoria
target_resumo = "    ExibirResumoAuditoria iTotalLinhas, iLinhasAuditadas, iLinhasComErro, (tFim - tInicio)\n    Exit Sub"
repl_resumo = """    If Not bSilencioso Then
        ExibirResumoAuditoria iTotalLinhas, iLinhasAuditadas, iLinhasComErro, (tFim - tInicio)
    End If
    Exit Sub"""

assert target_resumo in text, "target_resumo not found"
text = text.replace(target_resumo, repl_resumo, 1)
print("3. ExibirResumoAuditoria wrapped.")

# 4. Remove lower duplicate analiseMonitoramento alias
target_lower_alias = """' Alias para manter compatibilidade com macros já associadas a botões
Public Sub analiseMonitoramento()
    AnalisarErrosMonitoramentoEGerarLog
End Sub

"""
if target_lower_alias in text:
    text = text.replace(target_lower_alias, "", 1)
    print("4. Lower duplicate alias removed.")

# 5. Wrap TratarErro MsgBox
target_err = """TratarErro:
    MsgBox "Erro durante a auditoria de dados do monitoramento:" & vbNewLine & _
           Err.Number & " - " & Err.Description, vbCritical, MSG_TITULO
    Resume SairRotina"""

repl_err = """TratarErro:
    If Not bSilencioso Then
        MsgBox "Erro durante a auditoria de dados do monitoramento:" & vbNewLine & _
               Err.Number & " - " & Err.Description, vbCritical, MSG_TITULO
    End If
    Resume SairRotina"""

assert target_err in text, "target_err not found"
text = text.replace(target_err, repl_err, 1)
print("5. TratarErro wrapped.")

# 6. Update the check block in AuditarLinhaMonitoramento (Situação da OS rule)
target_check_start = "    Dim sSituacao    As String\n    \n    ' Captura inicial da Situação (Coluna C)"
pos_c_start = text.find(target_check_start)
assert pos_c_start != -1, "target_check_start not found"

target_check_end = "    Dim dDtAbert     As Date\n    Dim dDtFech      As Date\n    Dim bTemAbert    As Boolean\n    Dim bTemFech     As Boolean"
pos_c_end = text.find(target_check_end, pos_c_start)
assert pos_c_end != -1, "target_check_end not found"
full_c_end = pos_c_end + len(target_check_end)

new_check_block = """    Dim sSituacao    As String
    Dim sOSDisp      As String
    Dim vDtFech      As Variant
    Dim bTemFechVal  As Boolean
    Dim bDispSim     As Boolean
    
    ' Captura inicial da Situação (Coluna C), OS Disponibilizada (Coluna Y) e Fechamento (Coluna R)
    sSituacao = UCase(Trim(CStr(vDados(r, 3) & "")))
    sOSDisp = UCase(Trim(CStr(vDados(r, 25) & "")))
    vDtFech = vDados(r, 18)
    
    bTemFechVal = (Not IsEmpty(vDtFech) And Len(Trim(CStr(vDtFech & ""))) > 0)
    bDispSim = (sOSDisp = "SIM")
    
    '===============================================================================================================
    ' REGRA DE AUDITORIA DE SITUAÇÃO DA OS:
    ' 1. Linhas com Situação preenchida com "CO" (Concluído) ou "AT" (Atendido) são auditadas integralmente.
    ' 2. Se o campo "OS disponibilizada" estiver preenchido como "SIM" e/ou houver data de fechamento preenchida,
    '    mas a Situação estiver diferente de "CO" ou "AT" (ou vazia), a linha DEVE ser auditada e receber o log
    '    apontando a inconsistência na Coluna C (Situação incompatível com fechamento/disponibilização).
    ' 3. Registros com outra situação (ex: CA, PE) e que NÃO tenham OS disponibilizada como SIM nem fechamento
    '    continuam sumariamente ignorados.
    '===============================================================================================================
    If sSituacao = "CO" Or sSituacao = "AT" Then
        bAuditado = True
    ElseIf bDispSim Or bTemFechVal Then
        bAuditado = True
        
        Dim sMotivoSit As String
        If bDispSim And bTemFechVal Then
            sMotivoSit = "'OS disponibilizada' preenchido como 'Sim' e data de fechamento preenchida"
        ElseIf bDispSim Then
            sMotivoSit = "'OS disponibilizada' preenchido como 'Sim'"
        Else
            sMotivoSit = "data de fechamento preenchida"
        End If
        
        If Len(sSituacao) = 0 Then
            sLog = sLog & "(Coluna C): Situação da OS vazia, porém com " & sMotivoSit & " (exige situação 'CO' ou 'AT'); "
        Else
            sLog = sLog & "(Coluna C): Situação da OS '" & sSituacao & "' incompatível: " & sMotivoSit & " exige situação 'CO' ou 'AT'; "
        End If
    Else
        bAuditado = False
        Exit Sub
    End If

    Dim sAtendente   As String
    Dim sPerFat      As String
    Dim sContrato    As String
    Dim sUF          As String
    Dim sAtividade   As String
    Dim sItem        As String
    Dim sAplicacao   As String
    Dim sCodSol      As String
    Dim sChaveSol    As String
    Dim sNomeSol     As String
    Dim sGerencia    As String
    Dim sCodOS       As String
    Dim sSiglaObs    As String
    Dim sLocalidade  As String
    
    Dim vQtdSol      As Variant
    Dim vQtdAtend    As Variant
    Dim vDtSol       As Variant
    Dim vDtAbert     As Variant
    Dim vPrazoApp    As Variant
    Dim vPrazoComb   As Variant
    
    Dim dDtAbert     As Date
    Dim dDtFech      As Date
    Dim bTemAbert    As Boolean
    Dim bTemFech     As Boolean"""

text = text[:pos_c_start] + new_check_block + text[full_c_end:]
print("6. Check block in AuditarLinhaMonitoramento updated.")

# 7. Update header comments
p_h_start = text.find('REGRAS PRINCIPAIS ATENDIDAS:')
assert p_h_start != -1, "p_h_start not found"
p_h_end = text.find("'   2. GRAVAÇÃO DO LOG NA COLUNA Z:", p_h_start)
assert p_h_end != -1, "p_h_end not found"

new_header_part = """REGRAS PRINCIPAIS ATENDIDAS:
'   1. REGRA DE FILTRAGEM E AUDITORIA DA SITUAÇÃO DA OS:
'      São auditadas linhas com Situação "CO" (Concluído) ou "AT" (Atendido), BEM COMO linhas onde o campo
'      "OS disponibilizada" estiver "Sim" e/ou houver data de fechamento preenchida, mas a Situação estiver
'      diferente de CO ou AT (ou vazia), gerando o log de inconformidade correspondente na Coluna C.
'      Linhas com outras situações sem fechamento nem disponibilização continuam desconsideradas.
'
"""
text = text[:p_h_start] + new_header_part + text[p_h_end:]
print("7. Header comment updated.")

# Ensure CRLF line endings
text_crlf = text.replace('\n', '\r\n')

# Output paths
output_paths = [
    r'Codes\VBA\modAnaliseMonitoramento.bas',
    r'Codes\VBA\Funções Novas\modAnaliseMonitoramento.bas',
    r'Codes\VBA\vba_extracted_A4UU\modAnaliseMonitoramento.bas'
]

for p in output_paths:
    os.makedirs(os.path.dirname(p), exist_ok=True)
    with open(p, 'wb') as f:
        f.write(text_crlf.encode('cp1252'))
    print(f"Written: {p} ({len(text_crlf)} chars, {text_crlf.count(chr(10))} lines)")

print("ALL CHANGES APPLIED SUCCESSFULLY!")

