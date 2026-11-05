# -*- coding: cp1252 -*-
import shutil

with open(r'Codes\VBA\modAnaliseMonitoramento.bas', 'rb') as f:
    text = f.read().decode('cp1252').replace('\r\n', '\n')

old_block = """        If IsNumeric(vQtdSol) Then
            If Round(CDbl(vQtdSol), 4) <> Round(CDbl(vQtdAtend), 4) Then
                sLog = sLog & "(Coluna O e T): Quantidade solicitada diferente da quantidade atendida; "
            End If
        End If
        End If"""

new_block = """        If IsNumeric(vQtdSol) Then
            If Round(CDbl(vQtdSol), 4) <> Round(CDbl(vQtdAtend), 4) Then
                sLog = sLog & "(Coluna O e T): Quantidade solicitada diferente da quantidade atendida; "
            End If
        End If"""

assert old_block in text, "old_block not found"
text = text.replace(old_block, new_block, 1)

target_paths = [
    r'Codes\VBA\modAnaliseMonitoramento.bas',
    r'Codes\VBA\Funções Novas\modAnaliseMonitoramento.bas',
    r'Codes\VBA\vba_extracted_A4UU\modAnaliseMonitoramento.bas'
]

for p in target_paths:
    with open(p, 'wb') as f:
        f.write(text.replace('\n', '\r\n').encode('cp1252'))
    print("Fixed extra End If in:", p)

