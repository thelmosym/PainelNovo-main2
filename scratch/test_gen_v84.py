import os, zipfile, xml.etree.ElementTree as ET
from datetime import datetime, timedelta
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

def excel_serial_to_dt(val):
    if not val:
        return ""
    try:
        val_float = float(val)
        dt = datetime(1899, 12, 30) + timedelta(days=val_float)
        if dt.hour == 0 and dt.minute == 0 and dt.second == 0:
            return dt.strftime("%d/%m/%Y")
        return dt.strftime("%d/%m/%Y %H:%M")
    except Exception:
        return str(val)

def generate_fdm_demonstrativo():
    print("Iniciando geração da planilha demonstrativa de FDM, TSPE, TSPR e IAPFARQ...")
    
    # 1. Extração dos dados de Painel de controle MemoriaPetrobras V84.xlsm
    src_wb_path = 'Painel de controle MemoriaPetrobras V84.xlsm'
    
    sst = []
    with zipfile.ZipFile(src_wb_path) as z:
        if 'xl/sharedStrings.xml' in z.namelist():
            sst_root = ET.fromstring(z.read('xl/sharedStrings.xml'))
            for si in sst_root.findall('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}si'):
                sst.append(''.join([e.text for e in si.iter('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t') if e.text]))
        
        sheet_t2m = ET.fromstring(z.read('xl/worksheets/sheet31.xml'))
        rows_xml = sheet_t2m.findall('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}row')
        
    print(f"Total de linhas lidas em Painel T2M: {len(rows_xml)}")
    
    raw_data = []
    for r in rows_xml[1:]:
        cells = {}
        for c in r.findall('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}c'):
            ref = c.attrib.get('r')
            col_letter = ''.join([ch for ch in ref if ch.isalpha()])
            v = c.find('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}v')
            val = v.text if v is not None else ''
            if c.attrib.get('t') == 's' and val != '':
                val = sst[int(val)]
            cells[col_letter] = val
        
        raw_data.append({
            'periodo': cells.get('A', ''),
            'situacao': cells.get('B', ''),
            'contrato': cells.get('C', ''),
            'atividade': cells.get('D', ''),
            'item': cells.get('E', ''),
            'aplicacao': cells.get('F', ''),
            'solicitacao': cells.get('G', ''),
            'os': cells.get('L', ''),
            'dt_abertura': cells.get('O', ''),
            'dt_fechamento': cells.get('P', ''),
            'prazo_combinado': cells.get('Q', ''),
            'municipio': cells.get('X', ''),
            'uf': cells.get('Y', ''),
            'calculo_fdm': cells.get('AG', ''),
            'ref_fdm': cells.get('AH', ''),
            'ref_prazo': cells.get('AJ', ''),
            'dias_prazo': cells.get('AN', ''),
            'classif_fdm': cells.get('AS', ''),
            'status_prazo': cells.get('AT', ''),
            'log_sla': cells.get('AU', ''),
        })

    wb = openpyxl.Workbook()
    wb.remove(wb.active) # Remove default sheet
    
    # Paleta de Cores Corporativas Petrobras
    COR_NAVY = "1B365D"
    COR_TEAL = "006B76"
    COR_DARK_SLATE = "2F4F4F"
    COR_LIGHT_BG = "F4F6F9"
    COR_ZEBRA = "F8F9FA"
    COR_WHITE = "FFFFFF"
    COR_GREEN_BG = "D4EDDA"
    COR_GREEN_TXT = "155724"
    COR_RED_BG = "F8D7DA"
    COR_RED_TXT = "721C24"
    COR_GOLD_BG = "FFF3CD"
    COR_GOLD_TXT = "856404"
    COR_GRAY_BORDER = "D1D5DB"
    
    font_title = Font(name="Segoe UI", size=14, bold=True, color=COR_WHITE)
    font_subtitle = Font(name="Segoe UI", size=9, italic=True, color="E2E8F0")
    font_section = Font(name="Segoe UI", size=11, bold=True, color=COR_NAVY)
    font_th = Font(name="Segoe UI", size=9, bold=True, color=COR_WHITE)
    font_data = Font(name="Segoe UI", size=9, color="000000")
    font_data_bold = Font(name="Segoe UI", size=9, bold=True, color="000000")
    
    fill_header = PatternFill(start_color=COR_NAVY, end_color=COR_NAVY, fill_type="solid")
    fill_header_sub = PatternFill(start_color="2A4A7F", end_color="2A4A7F", fill_type="solid")
    fill_teal = PatternFill(start_color=COR_TEAL, end_color=COR_TEAL, fill_type="solid")
    fill_zebra = PatternFill(start_color=COR_ZEBRA, end_color=COR_ZEBRA, fill_type="solid")
    fill_white = PatternFill(start_color=COR_WHITE, end_color=COR_WHITE, fill_type="solid")
    fill_total = PatternFill(start_color="E9ECEF", end_color="E9ECEF", fill_type="solid")
    
    fill_np = PatternFill(start_color=COR_GREEN_BG, end_color=COR_GREEN_BG, fill_type="solid")
    font_np = Font(name="Segoe UI", size=9, bold=True, color=COR_GREEN_TXT)
    fill_fp = PatternFill(start_color=COR_RED_BG, end_color=COR_RED_BG, fill_type="solid")
    font_fp = Font(name="Segoe UI", size=9, bold=True, color=COR_RED_TXT)
    fill_isento = PatternFill(start_color=COR_GOLD_BG, end_color=COR_GOLD_BG, fill_type="solid")
    font_isento = Font(name="Segoe UI", size=9, bold=True, color=COR_GOLD_TXT)
    
    thin_border = Border(
        left=Side(style='thin', color=COR_GRAY_BORDER),
        right=Side(style='thin', color=COR_GRAY_BORDER),
        top=Side(style='thin', color=COR_GRAY_BORDER),
        bottom=Side(style='thin', color=COR_GRAY_BORDER)
    )
    double_bottom_border = Border(
        left=Side(style='thin', color=COR_GRAY_BORDER),
        right=Side(style='thin', color=COR_GRAY_BORDER),
        top=Side(style='thin', color=COR_GRAY_BORDER),
        bottom=Side(style='double', color=COR_NAVY)
    )

    # =========================================================================
    # ABA 1: RESUMO EXECUTIVO FDM
    # =========================================================================
    ws1 = wb.create_sheet(title="Resumo_Executivo_FDM")
    ws1.views.sheetView[0].showGridLines = True
    
    # Banner de Título (Colunas A até P)
    ws1.merge_cells("A1:P1")
    ws1["A1"] = "RELATÓRIO DE DEMONSTRAÇÃO E AUDITORIA DE INDICADORES CONTRATUAIS - FDM / TSPE / TSPR / IAPFARQ"
    ws1["A1"].font = font_title
    ws1["A1"].fill = fill_header
    ws1["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws1.row_dimensions[1].height = 32
    
    ws1.merge_cells("A2:P2")
    ws1["A2"] = "Governança Contratual Petrobras | Lote 1 (IRON MOUNTAIN - SAP 4600686987) e Lote 2 (PA - SAP 4600687006) | Ciclo 26/07 à 25/08"
    ws1["A2"].font = font_subtitle
    ws1["A2"].fill = fill_header_sub
    ws1["A2"].alignment = Alignment(horizontal="center", vertical="center")
    ws1.row_dimensions[2].height = 20
    
    # Cards KPI no Topo
    kpis = [
        ("A4:C5", "Total Chamados Concluídos (CO)", "582", "Base total faturável"),
        ("D4:F5", "Total Elegíveis FDM (TSPE)", "461", "TSPE1 = 314 | TSPE2 = 147"),
        ("G4:I5", "Total Atendidos no Prazo (TSPR)", "239", "TSPR1 = 124 | TSPR2 = 115"),
        ("J4:L5", "Total Fora do Prazo (Atrasos)", "222", "Ofensor principal: Pesquisa Item Avulso"),
        ("M4:P5", "Desempenho Global IAPFARQ", "51,84%", "FDM Geral Aplicado: 0,9700 (Glosa 3%)")
    ]
    for cell_range, label, val, sub in kpis:
        c1, c2 = cell_range.split(":")
        ws1.merge_cells(cell_range)
        ws1[c1] = f"{val}\n{label}\n({sub})"
        ws1[c1].font = Font(name="Segoe UI", size=10, bold=True, color=COR_NAVY)
        ws1[c1].fill = PatternFill(start_color="EBF3FA", end_color="EBF3FA", fill_type="solid")
        ws1[c1].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        for r_card in ws1[cell_range]:
            for c_card in r_card:
                c_card.border = thin_border
    ws1.row_dimensions[4].height = 26
    ws1.row_dimensions[5].height = 26

    # Seção 1: Tabela Principal Consolidada
    ws1["A7"] = "1. APURAÇÃO CONSOLIDADA DOS INDICADORES POR CONTRATO E LINHA DE SERVIÇO (ITENS 8.4.2 E 8.4.3)"
    ws1["A7"].font = font_section
    
    headers_t1 = [
        "Contrato", "Lote / Polo", "Código SAP", "Linha de Serviço Contratual",
        "Ref. FDM", "Total CO", "Exclusões (Isento/Definido)", "TSPE (Elegíveis)",
        "TSPR (No Prazo)", "Desvios (Fora do Prazo)", "% Atraso",
        "IAPFARQ (%)", "Meta SLA Contratual", "FDM Mensal Bruto", "FDM Final Homologado", "Impacto Financeiro na Fatura"
    ]
    
    for col_idx, h in enumerate(headers_t1, 1):
        cell = ws1.cell(row=8, column=col_idx, value=h)
        cell.font = font_th
        cell.fill = fill_header
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border
    ws1.row_dimensions[8].height = 30
    
    dados_t1 = [
        # Linha 9
        ("IRON-LT1", "Lote 1 (RJ / SP / ES)", "4600686987", "Serviço 1: Guarda de Documentos e Informações", 1, 271, 40, "=I9+J9", 75, 156, "=J9/H9", "=I9/H9*100", 99.00, "=IF(L9<90,0.97,IF(L9<95,0.98,IF(L9<99,0.99,1)))", 0.97, "Glosa de 3,0% sobre faturamento de Guarda"),
        # Linha 10
        ("IRON-LT1", "Lote 1 (RJ / SP / ES)", "4600686987", "Serviço 2: Gerenciamento de Documentos e Informações", 2, 140, 44, "=I10+J10", 71, 25, "=J10/H10", "=I10/H10*100", 99.00, "=IF(L10<90,0.97,IF(L10<95,0.98,IF(L10<99,0.99,1)))", 0.97, "Glosa de 3,0% sobre faturamento de Gerenciamento"),
        # Linha 11
        ("PA-LT2", "Lote 2 (Bahia - BA)", "4600687006", "Serviço 1: Guarda de Documentos e Informações", 1, 97, 14, "=I11+J11", 49, 34, "=J11/H11", "=I11/H11*100", 99.00, "=IF(L11<90,0.97,IF(L11<95,0.98,IF(L11<99,0.99,1)))", 0.97, "Glosa de 3,0% sobre faturamento de Guarda"),
        # Linha 12
        ("PA-LT2", "Lote 2 (Bahia - BA)", "4600687006", "Serviço 2: Gerenciamento de Documentos e Informações", 2, 65, 14, "=I12+J12", 44, 7, "=J12/H12", "=I12/H12*100", 99.00, "=IF(L12<90,0.97,IF(L12<95,0.98,IF(L12<99,0.99,1)))", 0.97, "Glosa de 3,0% sobre faturamento de Gerenciamento"),
    ]
    
    for r_offset, row_vals in enumerate(dados_t1, start=9):
        for c_idx, val in enumerate(row_vals, start=1):
            cell = ws1.cell(row=r_offset, column=c_idx, value=val)
            cell.font = font_data
            cell.border = thin_border
            
            if c_idx in [1, 2, 3]:
                cell.alignment = Alignment(horizontal="left", vertical="center")
            elif c_idx == 4:
                cell.alignment = Alignment(horizontal="left", vertical="center")
                cell.font = font_data_bold
            elif c_idx in [5, 6, 7, 8, 9, 10]:
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.number_format = "#,##0"
            elif c_idx == 11:
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.number_format = "0.00%"
            elif c_idx in [12, 13]:
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.number_format = "0.00"
                cell.font = font_data_bold
            elif c_idx in [14, 15]:
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.number_format = "0.0000"
                cell.font = font_data_bold
                cell.fill = PatternFill(start_color="FDE8E8", end_color="FDE8E8", fill_type="solid")
            elif c_idx == 16:
                cell.alignment = Alignment(horizontal="left", vertical="center")
                cell.font = Font(name="Segoe UI", size=9, bold=True, color=COR_RED_TXT)
        ws1.row_dimensions[r_offset].height = 20

    # Linha Totalizadora
    ws1.cell(row=13, column=1, value="TOTAL GERAL CONSOLIDADO")
    ws1.merge_cells("A13:E13")
    ws1["A13"].font = font_data_bold
    ws1["A13"].alignment = Alignment(horizontal="center", vertical="center")
    ws1["A13"].fill = fill_total
    
    ws1.cell(row=13, column=6, value="=SUM(F9:F12)").number_format = "#,##0"
    ws1.cell(row=13, column=7, value="=SUM(G9:G12)").number_format = "#,##0"
    ws1.cell(row=13, column=8, value="=SUM(H9:H12)").number_format = "#,##0"
    ws1.cell(row=13, column=9, value="=SUM(I9:I12)").number_format = "#,##0"
    ws1.cell(row=13, column=10, value="=SUM(J9:J12)").number_format = "#,##0"
    ws1.cell(row=13, column=11, value="=J13/H13").number_format = "0.00%"
    ws1.cell(row=13, column=12, value="=I13/H13*100").number_format = "0.00"
    ws1.cell(row=13, column=13, value=99.00).number_format = "0.00"
    ws1.cell(row=13, column=14, value="=IF(L13<90,0.97,IF(L13<95,0.98,IF(L13<99,0.99,1)))").number_format = "0.0000"
    ws1.cell(row=13, column=15, value=0.97).number_format = "0.0000"
    ws1.cell(row=13, column=16, value="Glosa Padrão de 3,0% Aplicada em ambos os contratos")
    
    for c_idx in range(1, 17):
        cell = ws1.cell(row=13, column=c_idx)
        cell.font = font_data_bold
        cell.fill = fill_total
        cell.border = double_bottom_border
        if c_idx in range(6, 16):
            cell.alignment = Alignment(horizontal="center", vertical="center")
    ws1.row_dimensions[13].height = 22

    # Seção 2: Tabelas Explicativas de Regras e Governança Contratual
    ws1["A16"] = "2. CRITÉRIOS, FÓRMULAS E TABELA CONTRATUAL DE ENQUADRAMENTO DO FDM (ITEM 8.4.3)"
    ws1["A16"].font = font_section
    
    headers_regras = ["Faixa de Cumprimento de Prazos (IAPFARQ)", "Fator FDM Bruto", "Percentual Faturado", "Glosa Contratual", "Classificação do Nível de Serviço", "Ação Operacional Obrigatória"]
    for c_idx, h in enumerate(headers_regras, 1):
        cell = ws1.cell(row=17, column=c_idx, value=h)
        cell.font = font_th
        cell.fill = fill_teal
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border
    ws1.row_dimensions[17].height = 24
    
    regras_faixas = [
        ("IAPFARQ >= 99,00%", 1.0000, "100,0%", "0,0% (Isento de Glosa)", "Nível de Excelência Operacional (Meta Cumprida)", "Pagamento pleno sem retenções."),
        ("95,00% <= IAPFARQ < 99,00%", 0.9900, "99,0%", "1,0% de Retenção", "Nível Aceitável / Atenção Operacional", "Alerta para planos de mitigação em rotas/filas."),
        ("90,00% <= IAPFARQ < 95,00%", 0.9800, "98,0%", "2,0% de Retenção", "Nível Insatisfatório / Crítico", "Notificação de desvio e plano de ação formal."),
        ("IAPFARQ < 90,00%", 0.9700, "97,0%", "3,0% de Retenção", "Nível Não Conforme / Penalidade Máxima Inicial", "Abertura de RNC e sujeito a degradação em cascata.")
    ]
    
    for r_idx, (faixa, fdm, fat, glosa, status, acao) in enumerate(regras_faixas, start=18):
        ws1.cell(row=r_idx, column=1, value=faixa).alignment = Alignment(horizontal="center")
        ws1.cell(row=r_idx, column=2, value=fdm).number_format = "0.0000"
        ws1.cell(row=r_idx, column=2).alignment = Alignment(horizontal="center")
        ws1.cell(row=r_idx, column=3, value=fat).alignment = Alignment(horizontal="center")
        ws1.cell(row=r_idx, column=4, value=glosa).alignment = Alignment(horizontal="center")
        ws1.cell(row=r_idx, column=5, value=status).alignment = Alignment(horizontal="left")
        ws1.cell(row=r_idx, column=6, value=acao).alignment = Alignment(horizontal="left")
        
        for c_idx in range(1, 7):
            cell = ws1.cell(row=r_idx, column=c_idx)
            cell.font = font_data
            cell.border = thin_border
            if r_idx == 21:
                cell.fill = PatternFill(start_color="FFF2E2", end_color="FFF2E2", fill_type="solid")
                cell.font = font_data_bold
        ws1.row_dimensions[r_idx].height = 20

    # Seção 3: Regra de Degradação por Reincidência (Cascata Histórica de 3 Meses)
    ws1["A24"] = "3. REGRA DE DEGRADAÇÃO EM CASCATA POR REINCIDÊNCIA CONSECUTIVA (ITEM 8.4.3 - PARÁGRAFO 2º)"
    ws1["A24"].font = font_section
    
    headers_cascata = ["Cenário de Reincidência", "FDM Mês Atual (M)", "FDM Mês Anterior (M-1)", "FDM Mês Retrasado (M-2)", "FDM Final Recalculado", "Glosa Acumulada", "Severidade Contratual"]
    for c_idx, h in enumerate(headers_cascata, 1):
        cell = ws1.cell(row=25, column=c_idx, value=h)
        cell.font = font_th
        cell.fill = PatternFill(start_color=COR_DARK_SLATE, end_color=COR_DARK_SLATE, fill_type="solid")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border
    ws1.row_dimensions[25].height = 24
    
    # Textos sem começar com sinal '=' para não serem interpretados como fórmulas inválidas
    dados_cascata = [
        ("Primeira ocorrência de IAPFARQ < 90%", "FDM = 0,97", "FDM >= 0,98 (ou sem histórico)", "Qualquer valor", 0.97, "3,0%", "Penalidade Inicial Padrão"),
        ("Reincidência por 2 meses consecutivos", "FDM = 0,97", "FDM = 0,97", "FDM > 0,97", 0.96, "4,0%", "Degradação Nível 1 (-1 ponto percentual)"),
        ("Reincidência por 3 meses consecutivos", "FDM = 0,97", "FDM = 0,97", "FDM <= 0,97", 0.95, "5,0%", "Degradação Nível 2 - Piso Máximo Contratual"),
        ("Manutenção no piso após 3 meses", "FDM = 0,97", "FDM = 0,95 ou 0,96", "Qualquer valor", 0.95, "5,0%", "Piso mínimo de 0,95 mantido até recuperação")
    ]
    for r_idx, vals in enumerate(dados_cascata, start=26):
        for c_idx, v in enumerate(vals, start=1):
            cell = ws1.cell(row=r_idx, column=c_idx, value=v)
            cell.font = font_data
            cell.border = thin_border
            if c_idx in [2, 3, 4, 5, 6]:
                cell.alignment = Alignment(horizontal="center")
            else:
                cell.alignment = Alignment(horizontal="left")
            if c_idx == 5:
                cell.number_format = "0.0000"
                cell.font = font_data_bold
        ws1.row_dimensions[r_idx].height = 20

    # =========================================================================
    # ABA 2: DETALHAMENTO POR ATIVIDADE
    # =========================================================================
    ws2 = wb.create_sheet(title="Detalhamento_Atividades")
    ws2.views.sheetView[0].showGridLines = True
    
    # Banner
    ws2.merge_cells("A1:L1")
    ws2["A1"] = "DETALHAMENTO OPERACIONAL POR ATIVIDADE - ANÁLISE DE OFENSORES DO FDM"
    ws2["A1"].font = font_title
    ws2["A1"].fill = fill_header
    ws2["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws2.row_dimensions[1].height = 28
    
    headers_t2 = [
        "Contrato", "Serviço Contratual", "Descrição da Atividade", "Ref. FDM",
        "Total Atendimentos CO", "No Prazo (NP)", "Fora do Prazo (FP)", "Isentos / Definidos",
        "TSPE Atividade", "TSPR Atividade", "% Conformidade no Prazo", "Diagnóstico Operacional"
    ]
    for c_idx, h in enumerate(headers_t2, 1):
        cell = ws2.cell(row=3, column=c_idx, value=h)
        cell.font = font_th
        cell.fill = fill_header
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border
    ws2.row_dimensions[3].height = 26
    
    from collections import defaultdict
    stats_ativ = defaultdict(lambda: {'total': 0, 'NP': 0, 'FP': 0, 'outros': 0})
    for d in raw_data:
        k = (d['contrato'], d['ref_fdm'], d['atividade'])
        stats_ativ[k]['total'] += 1
        st = d['status_prazo']
        if st == 'NP':
            stats_ativ[k]['NP'] += 1
        elif st == 'FP':
            stats_ativ[k]['FP'] += 1
        else:
            stats_ativ[k]['outros'] += 1
            
    r_cur = 4
    for (contrato, ref_fdm, atividade), val in sorted(stats_ativ.items()):
        total = val['total']
        np_count = val['NP']
        fp_count = val['FP']
        outros = val['outros']
        tspe = np_count + fp_count
        tspr = np_count
        
        serv_desc = "Serviço 1: Guarda" if ref_fdm == '1' else ("Serviço 2: Gerenciamento" if ref_fdm == '2' else "Isento / Fora FDM")
        
        ws2.cell(row=r_cur, column=1, value=contrato).alignment = Alignment(horizontal="center")
        ws2.cell(row=r_cur, column=2, value=serv_desc).alignment = Alignment(horizontal="left")
        ws2.cell(row=r_cur, column=3, value=atividade).alignment = Alignment(horizontal="left")
        ws2.cell(row=r_cur, column=4, value=int(ref_fdm) if ref_fdm.isdigit() else 0).alignment = Alignment(horizontal="center")
        ws2.cell(row=r_cur, column=5, value=total).number_format = "#,##0"
        ws2.cell(row=r_cur, column=6, value=np_count).number_format = "#,##0"
        ws2.cell(row=r_cur, column=7, value=fp_count).number_format = "#,##0"
        ws2.cell(row=r_cur, column=8, value=outros).number_format = "#,##0"
        ws2.cell(row=r_cur, column=9, value=f"=F{r_cur}+G{r_cur}").number_format = "#,##0"
        ws2.cell(row=r_cur, column=10, value=f"=F{r_cur}").number_format = "#,##0"
        ws2.cell(row=r_cur, column=11, value=f"=IF(I{r_cur}=0,1,J{r_cur}/I{r_cur})").number_format = "0.00%"
        
        pct_val = (tspr / tspe * 100) if tspe > 0 else 100.0
        if tspe == 0:
            diag = "Sem exigência de SLA (Atividade Isenta ou Duplicada)"
        elif pct_val >= 99.0:
            diag = "Meta Cumprida (100% no prazo)"
        elif pct_val >= 70.0:
            diag = "Desempenho Regular / Atenção"
        else:
            diag = "OFENSOR CRÍTICO DE SLA (Alto Volume Fora do Prazo)"
            
        diag_cell = ws2.cell(row=r_cur, column=12, value=diag)
        diag_cell.alignment = Alignment(horizontal="left")
        
        for c_idx in range(1, 13):
            c_elem = ws2.cell(row=r_cur, column=c_idx)
            c_elem.font = font_data
            c_elem.border = thin_border
            if c_idx in [5, 6, 7, 8, 9, 10, 11]:
                c_elem.alignment = Alignment(horizontal="center")
            if c_idx == 11:
                if tspe > 0 and pct_val < 50.0:
                    c_elem.fill = fill_fp
                    c_elem.font = font_fp
                elif tspe > 0 and pct_val >= 99.0:
                    c_elem.fill = fill_np
                    c_elem.font = font_np
            if c_idx == 12 and "OFENSOR" in diag:
                c_elem.font = Font(name="Segoe UI", size=9, bold=True, color=COR_RED_TXT)
                
        ws2.row_dimensions[r_cur].height = 20
        r_cur += 1

    # =========================================================================
    # ABA 3: BASE ANALÍTICA COMPLETA DE ORDENS (PAINEL T2M)
    # =========================================================================
    ws3 = wb.create_sheet(title="Base_Analitica_Ordens")
    ws3.views.sheetView[0].showGridLines = True
    
    headers_t3 = [
        "Período Fat.", "Contrato", "Código OS", "Código Solicitação", "Descrição da Atividade",
        "Item", "Aplicação", "Município", "UF", "D. Abertura", "D. Fechamento", "Prazo Combinado",
        "Ref. FDM", "Ref. Prazo", "Prazo SLA (dias)", "Classificação FDM", "Status Prazo",
        "Conta no TSPE?", "Conta no TSPR?", "LOG / Auditoria SLA"
    ]
    
    for c_idx, h in enumerate(headers_t3, 1):
        cell = ws3.cell(row=1, column=c_idx, value=h)
        cell.font = font_th
        cell.fill = fill_header
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border
    ws3.row_dimensions[1].height = 26
    
    for r_idx, d in enumerate(raw_data, start=2):
        st = d['status_prazo']
        ref_f = d['ref_fdm']
        conta_tspe = 1 if (st in ['NP', 'FP'] and ref_f in ['1', '2']) else 0
        conta_tspr = 1 if (st == 'NP' and ref_f in ['1', '2']) else 0
        
        row_vals_3 = [
            excel_serial_to_dt(d['periodo']),
            d['contrato'],
            d['os'],
            d['solicitacao'],
            d['atividade'],
            d['item'],
            d['aplicacao'],
            d['municipio'],
            d['uf'],
            excel_serial_to_dt(d['dt_abertura']),
            excel_serial_to_dt(d['dt_fechamento']),
            excel_serial_to_dt(d['prazo_combinado']),
            int(ref_f) if ref_f.isdigit() else 0,
            d['ref_prazo'],
            float(d['dias_prazo']) if d['dias_prazo'] and d['dias_prazo'].replace('.', '').isdigit() else d['dias_prazo'],
            d['classif_fdm'],
            st,
            conta_tspe,
            conta_tspr,
            d['log_sla']
        ]
        
        for c_idx, val in enumerate(row_vals_3, start=1):
            cell = ws3.cell(row=r_idx, column=c_idx, value=val)
            cell.font = font_data
            cell.border = thin_border
            
            if c_idx in [1, 2, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19]:
                cell.alignment = Alignment(horizontal="center")
            else:
                cell.alignment = Alignment(horizontal="left")
                
            if c_idx == 17:
                if val == 'NP':
                    cell.fill = fill_np
                    cell.font = font_np
                elif val == 'FP':
                    cell.fill = fill_fp
                    cell.font = font_fp
                elif val == '' or val is None:
                    cell.fill = fill_isento
                    cell.font = font_isento
        ws3.row_dimensions[r_idx].height = 18

    # Autoajuste de Largura de Colunas
    for ws in [ws1, ws2, ws3]:
        for col in ws.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                if cell.row in [1, 2, 4, 5] and ws.title == "Resumo_Executivo_FDM":
                    continue
                val_str = str(cell.value or '')
                if '\n' in val_str:
                    val_str = max(val_str.split('\n'), key=len)
                if len(val_str) > max_len:
                    max_len = len(val_str)
            ws.column_dimensions[col_letter].width = max(max_len + 3, 11)

    ws1.freeze_panes = "A9"
    ws2.freeze_panes = "A4"
    ws3.freeze_panes = "A2"
    
    out_dir = "MEDIÇÃO"
    os.makedirs(out_dir, exist_ok=True)
    out_path1 = os.path.join(out_dir, "Demonstrativo_Calculo_FDM_TSPE_IAPFARQ.xlsx")
    out_path2 = "Demonstrativo_Calculo_FDM_TSPE_IAPFARQ.xlsx"
    
    wb.save(out_path1)
    wb.save(out_path2)
    print(f"[OK] Planilha gerada com sucesso em: {out_path1}")
    print(f"[OK] Cópia idêntica na raiz do projeto: {out_path2}")

if __name__ == '__main__':
    generate_fdm_demonstrativo()
