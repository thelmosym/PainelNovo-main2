# -*- coding: cp1252 -*-
import os

target_paths = [
    r'Codes\VBA\modAnaliseMonitoramento.bas',
    r'Codes\VBA\Funções Novas\modAnaliseMonitoramento.bas',
    r'Codes\VBA\vba_extracted_A4UU\modAnaliseMonitoramento.bas'
]

dup = """' Alias para manter compatibilidade com macros já associadas a botões
Public Sub analiseMonitoramento()
    AnalisarErrosMonitoramentoEGerarLog
End Sub
"""

for p in target_paths:
    if os.path.exists(p):
        with open(p, 'rb') as f:
            raw = f.read()
        text = raw.decode('cp1252').replace('\r\n', '\n')
        
        dup_clean = dup.replace('\r\n', '\n')
        if dup_clean in text:
            text = text.replace(dup_clean, '')
            print('Removed duplicate alias in:', p)
        else:
            print('dup_clean not found in:', p)
            
        crlf_text = text.replace('\n', '\r\n')
        with open(p, 'wb') as f:
            f.write(crlf_text.encode('cp1252'))
        print('Saved clean CP1252 to:', p)

