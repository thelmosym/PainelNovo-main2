# -*- coding: utf-8 -*-
import openpyxl
import glob
import os

files = sorted([f for f in glob.glob('Monitoramento/*.xlsm') if not os.path.basename(f).startswith('~$')])

print("=" * 85)
print("AUDITORIA CONSOLIDADA DE RESULTADOS (INCLUINDO QUANTIDADE SOLICITADA x ATENDIDA)")
print("=" * 85)

for f in files:
    fname = os.path.basename(f)
    wb = openpyxl.load_workbook(f, data_only=True)
    ws = wb['Monitoramento']
    max_r = ws.max_row
    
    total_logs = 0
    logs_situacao = []
    logs_contrato_loc = []
    logs_qtd = []
    outros_logs = []
    
    for r in range(2, max_r + 1):
        num_os = ws.cell(r, 1).value
        sit = ws.cell(r, 3).value
        contrato = ws.cell(r, 4).value
        loc = ws.cell(r, 22).value
        q_sol = ws.cell(r, 15).value
        q_atend = ws.cell(r, 20).value
        disp = ws.cell(r, 25).value
        fech = ws.cell(r, 18).value
        log_val = str(ws.cell(r, 26).value or '').strip()
        
        if log_val:
            total_logs += 1
            if '(Coluna C):' in log_val or 'Situação da OS' in log_val:
                logs_situacao.append((r, num_os, sit, disp, fech, log_val))
            if '(Coluna D e V):' in log_val:
                logs_contrato_loc.append((r, num_os, contrato, loc, log_val))
            if '(Coluna O e T):' in log_val or 'Quantidade solicitada diferente' in log_val:
                logs_qtd.append((r, num_os, q_sol, q_atend, log_val))
            if not any(k in log_val for k in ['(Coluna C):', '(Coluna D e V):', '(Coluna O e T):']):
                outros_logs.append((r, num_os, log_val))
                
    print(f"\n>>> Arquivo: {fname} (Total de Linhas: {max_r})")
    print(f"    - Total de registros com LOG: {total_logs}")
    print(f"    - Inconsistências de Situação da OS (Coluna C): {len(logs_situacao)}")
    print(f"    - Inconsistências de Contrato x Localidade (Coluna D e V): {len(logs_contrato_loc)}")
    print(f"    - Inconsistências de Quantidade Solicitada x Atendida (Coluna O e T): {len(logs_qtd)}")
    print(f"    - Outros desvios cadastrais: {len(outros_logs)}")
    
    if logs_qtd:
        print("    [DETALHE - REGRA QUANTIDADE SOLICITADA x ATENDIDA]:")
        for item in logs_qtd:
            print(f"      * Linha {item[0]} (OS {item[1]}): Qtd. Solicitada={item[2]}, Qtd. Atendida={item[3]}")
            print(f"        Log: {item[4][:110]}...")

print("\n" + "=" * 85)
print("AUDITORIA FINAL CONCLUÍDA COM SUCESSO!")
print("=" * 85)
