# -*- coding: utf-8 -*-
import openpyxl
import glob
import os

files = glob.glob('Monitoramento/*.xlsm')
for f in sorted(files):
    wb = openpyxl.load_workbook(f, read_only=True, data_only=True)
    ws = wb['Monitoramento']
    print(f"=== {os.path.basename(f)} ===")
    print(f"Max rows: {ws.max_row}")
    
    headers = [cell for cell in next(ws.iter_rows(min_row=1, max_row=1, values_only=True))]
    print(f"Col 3 (C): {headers[2]} | Col 4 (D): {headers[3]} | Col 22 (V): {headers[21]} | Col 25 (Y): {headers[24]} | Col 26 (Z): {headers[25]}")
    
    total_logs = 0
    col_c_errors = 0
    col_dv_errors = 0
    
    for r, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        log_val = str(row[25] or '')
        if log_val.strip():
            total_logs += 1
            if '(Coluna C):' in log_val or 'Situa' in log_val:
                col_c_errors += 1
                if col_c_errors <= 3:
                    print(f"  Row {r}: [OS {row[0]}] Situacao='{row[2]}' Disp='{row[24]}' Fech='{row[17]}' -> Log: {log_val[:120]}")
            if 'Localidade' in log_val:
                col_dv_errors += 1
                if col_dv_errors <= 3:
                    print(f"  Row {r}: [OS {row[0]}] Contrato='{row[3]}' Loc='{row[21]}' -> Log: {log_val[:120]}")
                    
    print(f"Summary {os.path.basename(f)}: Total rows with LOG={total_logs}, Col C errors={col_c_errors}, Col D/V errors={col_dv_errors}\n")

