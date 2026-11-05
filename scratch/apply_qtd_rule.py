# -*- coding: cp1252 -*-
import os

target_paths = [
    r'Codes\VBA\modAnaliseMonitoramento.bas',
    r'Codes\VBA\Funções Novas\modAnaliseMonitoramento.bas',
    r'Codes\VBA\vba_extracted_A4UU\modAnaliseMonitoramento.bas'
]

new_rule = """        If IsNumeric(vQtdSol) Then
            If Round(CDbl(vQtdSol), 4) <> Round(CDbl(vQtdAtend), 4) Then
                sLog = sLog & "(Coluna O e T): Quantidade solicitada diferente da quantidade atendida; "
            End If
        End If"""

for p in target_paths:
    if not os.path.exists(p):
        print("Path not found:", p)
        continue
    with open(p, 'rb') as f:
        text = f.read().decode('cp1252').replace('\r\n', '\n')
    
    idx = text.find('Quantidade atendida superior')
    assert idx != -1, f"'Quantidade atendida superior' not found in {p}"
    
    p_start = text.rfind('        If IsNumeric(vQtdSol)', 0, idx)
    assert p_start != -1, f"'If IsNumeric(vQtdSol)' not found before idx in {p}"
    
    p_end = text.find('        End If', idx) + len('        End If')
    
    old_slice = text[p_start:p_end]
    print(f"In {p}, replacing slice of length {len(old_slice)}")
    text = text[:p_start] + new_rule + text[p_end:]
    
    # Update header comment if present
    target_hdr = "'   3. CONSISTÊNCIA CONTRATO x LOCALIDADE (TABELAB):\n'      Valida se o contrato informado na Coluna D condiz com a localidade da Coluna V,\n'      resolvendo a cadeia Localidade -> Centro -> UF -> Contrato homologado na TabelaB."
    new_hdr = target_hdr + "\n'   4. CONSISTÊNCIA QUANTIDADE SOLICITADA x ATENDIDA (COLUNAS O e T):\n'      Gera log quando a quantidade solicitada (Coluna O) for diferente da quantidade atendida (Coluna T)."
    if target_hdr in text and new_hdr not in text:
        text = text.replace(target_hdr, new_hdr, 1)
        
    crlf_text = text.replace('\n', '\r\n')
    with open(p, 'wb') as f:
        f.write(crlf_text.encode('cp1252'))
    print(f"Updated {p} successfully.")
