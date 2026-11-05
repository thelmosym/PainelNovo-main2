# -*- coding: cp1252 -*-
import os

with open(r'Codes\VBA\modAnaliseMonitoramento.bas', 'rb') as f:
    raw = f.read()

text = raw.decode('cp1252').replace('\r\n', '\n')

# 1. Replace entry points
old_entry = """Public Sub AnalisarErrosMonitoramentoEGerarLog()
    Dim wsMonit As Worksheet"""

new_entry = """Public Sub AnalisarErrosMonitoramentoEGerarLog()
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

assert old_entry in text, "old_entry not found"
text = text.replace(old_entry, new_entry, 1)

# 2. Wrap MsgBox confirmation
old_msg_conf = """    If MsgBox("Será executada a análise de conformidade nos registros da aba Monitoramento." & vbNewLine & vbNewLine & _
              "  Regra Inicial: Somente linhas com Situação 'CO' (Concluído) ou 'AT' (Atendido) serão auditadas." & vbNewLine & _
              "  Destino: Os erros encontrados serão gravados na Coluna Z (LOG) da respectiva linha." & vbNewLine & vbNewLine & _
              "Total de registros a varrer: " & Format(iTotalLinhas, "#,##0") & vbNewLine & vbNewLine & _
              "Deseja iniciar a auditoria?", _
              vbYesNo + vbQuestion, MSG_TITULO) = vbNo Then
        Exit Sub
    End If"""

new_msg_conf = """    If Not bSilencioso Then
        If MsgBox("Será executada a análise de conformidade nos registros da aba Monitoramento." & vbNewLine & vbNewLine & _
                  "  Regra Inicial: Somente linhas com Situação 'CO' (Concluído) ou 'AT' (Atendido) serão auditadas." & vbNewLine & _
                  "  Destino: Os erros encontrados serão gravados na Coluna Z (LOG) da respectiva linha." & vbNewLine & vbNewLine & _
                  "Total de registros a varrer: " & Format(iTotalLinhas, "#,##0") & vbNewLine & vbNewLine & _
                  "Deseja iniciar a auditoria?", _
                  vbYesNo + vbQuestion, MSG_TITULO) = vbNo Then
            Exit Sub
        End If
    End If"""

assert old_msg_conf in text, "old_msg_conf not found"
text = text.replace(old_msg_conf, new_msg_conf, 1)

# 3. Wrap ExibirResumoAuditoria
old_resumo = "    ExibirResumoAuditoria iTotalLinhas, iLinhasAuditadas, iLinhasComErro, (tFim - tInicio)\n    Exit Sub"
new_resumo = """    If Not bSilencioso Then
        ExibirResumoAuditoria iTotalLinhas, iLinhasAuditadas, iLinhasComErro, (tFim - tInicio)
    End If
    Exit Sub"""

assert old_resumo in text, "old_resumo not found"
text = text.replace(old_resumo, new_resumo, 1)

# 4. Remove duplicate analiseMonitoramento if it exists lower in the file
dup_alias = """' Alias para manter compatibilidade com macros já associadas a botões
Public Sub analiseMonitoramento()
    AnalisarErrosMonitoramentoEGerarLog
End Sub

"""
if dup_alias in text:
    text = text.replace(dup_alias, "", 1)
    print("Removed lower duplicate alias.")

# 5. Wrap TratarErro MsgBox
old_err = """TratarErro:
    MsgBox "Erro durante a auditoria de dados do monitoramento:" & vbNewLine & _
           Err.Number & " - " & Err.Description, vbCritical, MSG_TITULO
    Resume SairRotina"""

new_err = """TratarErro:
    If Not bSilencioso Then
        MsgBox "Erro durante a auditoria de dados do monitoramento:" & vbNewLine & _
               Err.Number & " - " & Err.Description, vbCritical, MSG_TITULO
    End If
    Resume SairRotina"""

assert old_err in text, "old_err not found"
text = text.replace(old_err, new_err, 1)

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

print("SUCCESS: Silent mode added!")

