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
print("1. Entry points replaced.")

# 2. Wrap MsgBox confirmation (search by non-accented string)
pos_total = text.find('iTotalLinhas = ul - LINHA_INICIAL_DADOS + 1')
assert pos_total != -1, "pos_total not found"
idx_start = text.find('If MsgBox(', pos_total)
assert idx_start != -1, "MsgBox start not found"
idx_end = text.find('Exit Sub\n    End If', idx_start) + len('Exit Sub\n    End If')
old_block = text[idx_start:idx_end]
new_block = 'If Not bSilencioso Then\n    ' + old_block.replace('\n', '\n    ') + '\n    End If'
text = text[:idx_start] + new_block + text[idx_end:]
print("2. Confirmation wrapped.")

# 3. Wrap ExibirResumoAuditoria
idx_resumo = text.find('ExibirResumoAuditoria iTotalLinhas, iLinhasAuditadas')
assert idx_resumo != -1, "ExibirResumoAuditoria not found"
line_start = text.rfind('\n', 0, idx_resumo) + 1
line_end = text.find('\n', idx_resumo)
resumo_line = text[line_start:line_end]
wrapped_resumo = f"""    If Not bSilencioso Then
    {resumo_line}
    End If"""
text = text[:line_start] + wrapped_resumo + text[line_end:]
print("3. Resumo wrapped.")

# 4. Remove duplicate analiseMonitoramento if it exists lower in the file
dup_anchor = 'Public Sub analiseMonitoramento()\n    AnalisarErrosMonitoramentoEGerarLog\nEnd Sub'
if dup_anchor in text:
    text = text.replace(dup_anchor, "", 1)
    print("4. Removed lower duplicate alias.")

# 5. Wrap TratarErro MsgBox
idx_err = text.find('TratarErro:')
assert idx_err != -1, "TratarErro not found"
msg_start = text.find('MsgBox', idx_err)
msg_end = text.find('MSG_TITULO', msg_start) + len('MSG_TITULO')
err_msg_block = text[msg_start:msg_end]
wrapped_err = f"""If Not bSilencioso Then
        {err_msg_block}
    End If"""
text = text[:msg_start] + wrapped_err + text[msg_end:]
print("5. TratarErro wrapped.")

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

print("SUCCESS: Silent mode added to all files!")
