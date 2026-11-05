import openpyxl
import os

print("=== VERIFYING DEMONSTRATIVO FILES ===")
files = [
    r"MEDIÇÃO\Demonstrativo_Calculo_FDM_TSPE_IAPFARQ.xlsx",
    r"Demonstrativo_Calculo_FDM_TSPE_IAPFARQ.xlsx"
]

for fpath in files:
    if not os.path.exists(fpath):
        print(f"ERROR: {fpath} not found!")
        continue
    
    mtime = os.path.getmtime(fpath)
    size = os.path.getsize(fpath)
    print(f"File: {fpath} | Size: {size} bytes | Mtime: {mtime}")
    
    wb = openpyxl.load_workbook(fpath, data_only=True)
    print("Sheets:", wb.sheetnames)
    
    # 1. Base_Analitica_Ordens
    ws_base = wb["Base_Analitica_Ordens"]
    max_r = ws_base.max_row
    print(f"  Base_Analitica_Ordens rows: {max_r}")
    
    # Headers
    headers = [ws_base.cell(1, c).value for c in range(1, 21)]
    print(f"  Headers: {headers}")
    
    # Count errors and statuses
    np_count = 0
    fp_count = 0
    err_count = 0
    tspe_sum = 0
    tspr_sum = 0
    
    # Check specific rows: 10, 130, 278, 356, 444, 447, 448, 453, 615, 645, 660, 713, 714, 715, 716
    sample_rows = [10, 130, 278, 356, 444, 447, 448, 453, 615, 645, 660, 713, 714, 715, 716]
    sample_details = []
    
    for r in range(2, max_r + 1):
        os_num = ws_base.cell(r, 3).value
        status = ws_base.cell(r, 17).value
        tspe = ws_base.cell(r, 18).value or 0
        tspr = ws_base.cell(r, 19).value or 0
        log_val = ws_base.cell(r, 20).value
        
        if status == "NP": np_count += 1
        elif status == "FP": fp_count += 1
        
        tspe_sum += tspe
        tspr_sum += tspr
        
        if log_val and "Erro no calculo" in str(log_val):
            err_count += 1
            if err_count <= 3:
                print(f"  Calculation Error at row {r}: {log_val}")
                
        if r in sample_rows:
            sample_details.append(f"Row {r} (OS {os_num}): Classif={ws_base.cell(r, 16).value}, Status={status}, TSPE={tspe}, TSPR={tspr}, Log='{log_val}'")
            
    print(f"  Total NP: {np_count}, Total FP: {fp_count}, Total TSPE: {tspe_sum}, Total TSPR: {tspr_sum}")
    print(f"  Total rows with Erro no calculo: {err_count}")
    print("  Sample rows audit:")
    for s in sample_details[:6]:
        print("    ", s)
        
    # 2. Resumo_Executivo_FDM
    ws_resumo = wb["Resumo_Executivo_FDM"]
    print("  Resumo_Executivo_FDM sample values:")
    for r in range(4, 10):
        row_vals = [ws_resumo.cell(r, c).value for c in range(1, 17)]
        if any(row_vals):
            print(f"    Row {r}: {row_vals}")

print("\n=== VERIFYING PAINEL T2M IN MAIN WORKBOOK ===")
wb_main = openpyxl.load_workbook("Painel de controle MemoriaPetrobras V.8.5-Outubro.xlsm", data_only=True)
ws_t2m = wb_main["Painel T2M"]
print(f"Painel T2M max row: {ws_t2m.max_row}")

t2m_errs = 0
for r in range(2, ws_t2m.max_row + 1):
    log_val = str(ws_t2m.cell(r, 47).value or "")
    if "Erro no calculo" in log_val:
        t2m_errs += 1
print(f"Painel T2M total rows with 'Erro no calculo' in Col 47: {t2m_errs}")

# Check link in Painel cell F18
ws_pnl = wb_main["Painel"]
print(f"Painel F18 value: {ws_pnl['F18'].value}")

