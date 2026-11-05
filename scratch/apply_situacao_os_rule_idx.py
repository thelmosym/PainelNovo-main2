# -*- coding: cp1252 -*-
import os

with open(r'Codes\VBA\modAnaliseMonitoramento.bas', 'rb') as f:
    raw = f.read()

text = raw.decode('cp1252').replace('\r\n', '\n')
print(f"Loaded modAnaliseMonitoramento.bas: {len(text)} chars, {text.count(chr(10))} lines")

# 1. Update header comment
p_h_start = text.find('REGRAS PRINCIPAIS ATENDIDAS:')
assert p_h_start != -1, "p_h_start not found"
p_h_end = text.find("2. ", p_h_start)
assert p_h_end != -1, "p_h_end not found"
line_end_pos = text.rfind('\n', 0, p_h_end) + 1

new_header_part = """REGRAS PRINCIPAIS ATENDIDAS:
'   1. REGRA DE FILTRAGEM E AUDITORIA DA SITUAÇÃO DA OS:
'      São auditadas linhas com Situação "CO" (Concluído) ou "AT" (Atendido), BEM COMO linhas onde o campo
'      "OS disponibilizada" estiver "Sim" e/ou houver data de fechamento preenchida, mas a Situação estiver
'      diferente de CO ou AT (ou vazia), gerando o log de inconformidade correspondente na Coluna C.
'      Linhas com outras situações sem fechamento nem disponibilização continuam desconsideradas.
'   """
text = text[:p_h_start] + new_header_part + text[p_h_end:]
print("1. Header updated.")

# 2. Update initial check block in AuditarLinhaMonitoramento
p1 = text.find('Dim sSituacao    As String')
assert p1 != -1, "p1 not found"
p2 = text.find('Dim bTemFech     As Boolean', p1)
assert p2 != -1, "p2 not found"
p2_end = text.find('\n', p2) + 1

new_check_block = """Dim sSituacao    As String
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
            sMotivoSit = "'OS disponibilizada' preenchido como 'Sim' e data de fechamento informada"
        ElseIf bDispSim Then
            sMotivoSit = "'OS disponibilizada' preenchido como 'Sim'"
        Else
            sMotivoSit = "data de fechamento informada"
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
    Dim bTemFech     As Boolean
"""
text = text[:p1] + new_check_block + text[p2_end:]
print("2. Check block in AuditarLinhaMonitoramento updated.")

# 3. Update prompt MsgBox text in ExecutarAnaliseMonitoramento
p_msg = text.find('Deseja iniciar a auditoria')
assert p_msg != -1, "p_msg not found"
p_msg_line_start = text.rfind('If MsgBox(', 0, p_msg)
p_msg_line_end = text.find('Exit Sub\n    End If', p_msg) + len('Exit Sub\n    End If')

new_prompt_block = """If MsgBox("Será executada a análise de conformidade nos registros da aba Monitoramento." & vbNewLine & vbNewLine & _
              "  Regras de Auditoria:" & vbNewLine & _
              "   - Linhas com Situação 'CO' (Concluído) ou 'AT' (Atendido);" & vbNewLine & _
              "   - Linhas com 'OS disponibilizada = Sim' e/ou data de fechamento informada." & vbNewLine & vbNewLine & _
              "  Destino: Os erros encontrados serão gravados na Coluna Z (LOG) da respectiva linha." & vbNewLine & vbNewLine & _
              "Total de registros a varrer: " & Format(iTotalLinhas, "#,##0") & vbNewLine & vbNewLine & _
              "Deseja iniciar a auditoria?", _
              vbYesNo + vbQuestion, MSG_TITULO) = vbNo Then
        Exit Sub
    End If"""

text = text[:p_msg_line_start] + new_prompt_block + text[p_msg_line_end:]
print("3. Prompt MsgBox updated.")

# 4. Update ExibirResumoAuditoria comment text
p_res = text.find('Regra inicial aplicada:')
if p_res != -1:
    p_res_end = text.find('\n', p_res)
    text = text[:p_res] + "Regra aplicada: Situação CO/AT ou OS disponibilizada (Sim) / Fechamento informado\" & vbNewLine & _" + text[p_res_end:]
    print("4. Resumo text updated.")

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

print("SUCCESS: Rule 'Situação da OS' implemented in all source files!")
