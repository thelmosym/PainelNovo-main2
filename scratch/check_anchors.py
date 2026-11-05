# -*- coding: cp1252 -*-
with open(r'Codes\VBA\modAnaliseMonitoramento.bas', 'rb') as f:
    text = f.read().decode('cp1252').replace('\r\n', '\n')

print(f"Loaded: {len(text)} chars, {text.count(chr(10))} lines")

p_entry = text.find('Public Sub AnalisarErrosMonitoramentoEGerarLog()\n    Dim wsMonit As Worksheet')
print("p_entry:", p_entry)

p_msg = text.find('Deseja iniciar a auditoria')
print("p_msg:", p_msg)

p_check = text.find('Dim sSituacao    As String\n    \n    \' Captura inicial da Situa')
print("p_check:", p_check)

p_resumo = text.find('ExibirResumoAuditoria iTotalLinhas, iLinhasAuditadas')
print("p_resumo:", p_resumo)

