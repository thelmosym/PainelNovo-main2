# -*- coding: cp1252 -*-
import os

with open(r'Codes\VBA\modAnaliseMonitoramento.bas', 'r', encoding='cp1252') as f:
    text = f.read().replace('\r\n', '\n')

print(f"Original text: {len(text)} chars")

# Check targets
targets = [
    ('decl_target', '    Dim dictPrazoAtiv   As Object'),
    ('init_target', '    Set dictPrazoAtiv = CreateObject("Scripting.Dictionary")'),
    ('comp_target', '    dictPrazoAtiv.CompareMode = vbTextCompare'),
    ('call_carregar', 'dictSiglasObs, dictLocalidades, dictPrazoAtiv\n'),
    ('call_auditar', 'dictSiglasObs, dictLocalidades, dictPrazoAtiv\n'),
    ('sig_auditar', 'ByRef dictPrazoAtiv As Object)\n\n    \' Mapeamento de colunas'),
    ('sig_carregar', 'ByRef dictPrazoAtiv As Object)\n\n    Dim ulRef As Long'),
    ('val_loc', '    \'--- 23. COLUNA Y: OS DISPONIBILIZADA? ---'),
    ('end_carregar', 'dictContratoUF("4600687006|BA") = True\n\nEnd Sub')
]

for name, t in targets:
    found = t in text
    print(f"Target '{name}': {'FOUND' if found else 'NOT FOUND'}")
    if not found:
        print(f"   Value: {repr(t)}")

