# -*- coding: utf-8 -*-
import os

with open(r'Codes\VBA\modAnaliseMonitoramento.bas', 'rb') as f:
    raw = f.read()

text = raw.decode('latin1').replace('\r\n', '\n')
print(f"Loaded: {len(text)} chars, {text.count(chr(10))} lines")

# 1. Update header comment
old_h1 = """' REGRAS PRINCIPAIS ATENDIDAS:
'   1. REGRA INICIAL DE FILTRAGEM:
'      Apenas linhas onde a coluna 'Situa\xe7\xe3o' (Coluna C) estiver preenchida com "CO" (Conclu\xeddo) ou "AT" (Atendido)
'      s\xe3o analisadas para procura de erros. Linhas com outras situa\xe7\xf5es ou em branco s\xe3o desconsideradas."""

new_h1 = """' REGRAS PRINCIPAIS ATENDIDAS:
'   1. REGRA DE FILTRAGEM E AUDITORIA DA SITUA\xc7\xc3O DA OS:
'      S\xe3o auditadas linhas com Situa\xe7\xe3o "CO" (Conclu\xeddo) ou "AT" (Atendido), BEM COMO linhas onde o campo
'      "OS disponibilizada" estiver "Sim" e/ou houver data de fechamento preenchida, mas a Situa\xe7\xe3o estiver
'      diferente de CO ou AT (ou vazia), gerando o log de inconformidade correspondente na Coluna C.
'      Linhas com outras situa\xe7\xf5es sem fechamento nem disponibiliza\xe7\xe3o continuam desconsideradas."""

assert old_h1 in text, "old_h1 not found"
text = text.replace(old_h1, new_h1, 1)
print("1. Header comment updated.")

# 2. Update initial check block in AuditarLinhaMonitoramento
old_check_block = """    Dim sSituacao    As String
    
    ' Captura inicial da Situa\xe7\xe3o (Coluna C)
    sSituacao = UCase(Trim(CStr(vDados(r, 3) & "")))
    
    '===============================================================================================================
    ' REGRA INICIAL MANDAT\xc1RIA:
    ' Apenas procurar erros se a coluna Situa\xe7\xe3o estiver preenchida estritamente com "CO" ou "AT".
    ' Registros com qualquer outra situa\xe7\xe3o (ex: CA, PE, etc.) ou vazias s\xe3o sumariamente ignorados.
    '===============================================================================================================
    If sSituacao <> "CO" And sSituacao <> "AT" Then
        bAuditado = False
        Exit Sub
    End If
    
    ' Marca que a linha atende \xe0 regra inicial e deve ser auditada
    bAuditado = True

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
    Dim sOSDisp      As String
    
    Dim vQtdSol      As Variant
    Dim vQtdAtend    As Variant
    Dim vDtSol       As Variant
    Dim vDtAbert     As Variant
    Dim vDtFech      As Variant
    Dim vPrazoApp    As Variant
    Dim vPrazoComb   As Variant
    
    Dim dDtAbert     As Date
    Dim dDtFech      As Date
    Dim bTemAbert    As Boolean
    Dim bTemFech     As Boolean"""

new_check_block = """    Dim sSituacao    As String
    Dim sOSDisp      As String
    Dim vDtFech      As Variant
    Dim bTemFechVal  As Boolean
    Dim bDispSim     As Boolean
    
    ' Captura inicial da Situa\xe7\xe3o (Coluna C), OS Disponibilizada (Coluna Y) e Fechamento (Coluna R)
    sSituacao = UCase(Trim(CStr(vDados(r, 3) & "")))
    sOSDisp = UCase(Trim(CStr(vDados(r, 25) & "")))
    vDtFech = vDados(r, 18)
    
    bTemFechVal = (Not IsEmpty(vDtFech) And Len(Trim(CStr(vDtFech & ""))) > 0)
    bDispSim = (sOSDisp = "SIM")
    
    '===============================================================================================================
    ' REGRA DE AUDITORIA DE SITUA\xc7\xc3O DA OS:
    ' 1. Linhas com Situa\xe7\xe3o preenchida com "CO" (Conclu\xeddo) ou "AT" (Atendido) s\xe3o auditadas integralmente.
    ' 2. Se o campo "OS disponibilizada" estiver preenchido como "SIM" e/ou houver data de fechamento preenchida,
    '    mas a Situa\xe7\xe3o estiver diferente de "CO" ou "AT" (ou vazia), a linha DEVE ser auditada e receber o log
    '    apontando a inconsist\xeancia na Coluna C (Situa\xe7\xe3o incompat\xedvel com fechamento/disponibiliza\xe7\xe3o).
    ' 3. Registros com outra situa\xe7\xe3o (ex: CA, PE) e que N\xc3O tenham OS disponibilizada como SIM nem fechamento
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
            sLog = sLog & "(Coluna C): Situa\xe7\xe3o da OS vazia, por\xe9m com " & sMotivoSit & " (exige situa\xe7\xe3o 'CO' ou 'AT'); "
        Else
            sLog = sLog & "(Coluna C): Situa\xe7\xe3o da OS '" & sSituacao & "' incompat\xedvel: " & sMotivoSit & " exige situa\xe7\xe3o 'CO' ou 'AT'; "
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

assert old_check_block in text, "old_check_block not found"
text = text.replace(old_check_block, new_check_block, 1)
print("2. Check block in AuditarLinhaMonitoramento updated.")

# 3. Update prompt MsgBox text in ExecutarAnaliseMonitoramento
old_prompt = """    If Not bSilencioso Then
        If MsgBox("Ser\xe1 executada a an\xe1lise de conformidade nos registros da aba Monitoramento." & vbNewLine & vbNewLine & _
                  "  Regra Inicial: Somente linhas com Situa\xe7\xe3o 'CO' (Conclu\xeddo) ou 'AT' (Atendido) ser\xe1o auditadas." & vbNewLine & _
                  "  Destino: Os erros encontrados ser\xe1o gravados na Coluna Z (LOG) da respectiva linha." & vbNewLine & vbNewLine & _
                  "Total de registros a varrer: " & Format(iTotalLinhas, "#,##0") & vbNewLine & vbNewLine & _
                  "Deseja iniciar a auditoria?", _
                  vbYesNo + vbQuestion, MSG_TITULO) = vbNo Then
            Exit Sub
        End If
    End If"""

new_prompt = """    If Not bSilencioso Then
        If MsgBox("Ser\xe1 executada a an\xe1lise de conformidade nos registros da aba Monitoramento." & vbNewLine & vbNewLine & _
                  "  Regras de Auditoria:" & vbNewLine & _
                  "   - Linhas com Situa\xe7\xe3o 'CO' (Conclu\xeddo) ou 'AT' (Atendido);" & vbNewLine & _
                  "   - Linhas com 'OS disponibilizada = Sim' e/ou data de fechamento informada." & vbNewLine & vbNewLine & _
                  "  Destino: Os erros encontrados ser\xe1o gravados na Coluna Z (LOG) da respectiva linha." & vbNewLine & vbNewLine & _
                  "Total de registros a varrer: " & Format(iTotalLinhas, "#,##0") & vbNewLine & vbNewLine & _
                  "Deseja iniciar a auditoria?", _
                  vbYesNo + vbQuestion, MSG_TITULO) = vbNo Then
            Exit Sub
        End If
    End If"""

assert old_prompt in text, "old_prompt not found"
text = text.replace(old_prompt, new_prompt, 1)
print("3. Prompt MsgBox updated.")

# 4. Update ExibirResumoAuditoria comment text
old_resumo_text = '           "  Regra inicial aplicada: Situa\xe7\xe3o = CO (Conclu\xeddo) ou AT (Atendido)" & vbNewLine & _'
new_resumo_text = '           "  Regra aplicada: Situa\xe7\xe3o CO/AT ou OS disponibilizada (Sim) / Fechamento informado" & vbNewLine & _'

assert old_resumo_text in text, "old_resumo_text not found"
text = text.replace(old_resumo_text, new_resumo_text, 1)
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
        f.write(text_crlf.encode('latin1'))
    print(f"Written: {p} ({len(text_crlf)} chars, {text_crlf.count(chr(10))} lines)")

print("SUCCESS: Rule 'Situação da OS' implemented in all source files!")
