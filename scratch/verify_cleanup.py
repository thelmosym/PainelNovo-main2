import glob, os, zipfile, xml.etree.ElementTree as ET

def verify_file(f):
    name = os.path.basename(f)
    print("=" * 70)
    print(f"AUDITORIA: {name}")
    print("=" * 70)
    
    with zipfile.ZipFile(f) as z:
        # 1. VBA Verification
        bin_data = z.read('xl/vbaProject.bin')
        has_mod = b'modLimparMonitoramento' in bin_data
        has_manter_nc = b'LimparAbaMonitoramentoManterNC' in bin_data
        has_entrypoint = b'LimparMonitoramento' in bin_data
        print(f"  [VBA] modLimparMonitoramento: {has_mod}")
        print(f"  [VBA] Sub LimparAbaMonitoramentoManterNC: {has_manter_nc}")
        print(f"  [VBA] Sub LimparMonitoramento: {has_entrypoint}")
        
        # 2. Shared strings
        sst = []
        if 'xl/sharedStrings.xml' in z.namelist():
            sst_root = ET.fromstring(z.read('xl/sharedStrings.xml'))
            for si in sst_root.findall('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}si'):
                sst.append(''.join([e.text for e in si.iter('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t') if e.text]))
                
        # 3. Table verification
        t_name = ""
        t_ref = ""
        for t in z.namelist():
            if 'tables/table' in t:
                t_root = ET.fromstring(z.read(t))
                if t_root.attrib.get('name') == 'Monitoramento' or 'table1' in t:
                    t_name = t_root.attrib.get('name')
                    t_ref = t_root.attrib.get('ref')
                    break
        print(f"  [TABELA] Nome: '{t_name}' | Intervalo: {t_ref}")
        
        # 4. Sheet rows verification
        sheet1 = ET.fromstring(z.read('xl/worksheets/sheet1.xml'))
        rows = sheet1.findall('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}row')
        total_rows = len(rows)
        data_rows = max(0, total_rows - 2)
        print(f"  [LINHAS] Total na aba: {total_rows} (2 cabeçalho + {data_rows} dados)")
        
        contracts = []
        situations = set()
        for r in rows:
            r_idx = int(r.attrib.get('r'))
            if r_idx < 3:
                continue
            sit = ""
            ct = ""
            for c in r.findall('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}c'):
                ref = c.attrib.get('r')
                cell_type = c.attrib.get('t')
                v = c.find('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}v')
                val = v.text if v is not None else ""
                if cell_type == 's' and val != "":
                    val = sst[int(val)]
                if ref == f'C{r_idx}':
                    sit = val
                elif ref == f'D{r_idx}':
                    ct = val
            situations.add(sit)
            contracts.append(ct)
            
        print(f"  [SITUAÇÃO] Valores únicos encontrados: {situations}")
        is_only_nc = situations == {'NC'}
        print(f"  [SITUAÇÃO] Apenas registros 'NC' mantidos? {is_only_nc}")
        
        is_sorted = contracts == sorted(contracts)
        print(f"  [ORDENAÇÃO] Contratos ordenados crescentemente? {is_sorted}")
        contract_summary = {}
        for c in contracts:
            contract_summary[c] = contract_summary.get(c, 0) + 1
        print(f"  [CONTRATOS] Distribuição: {contract_summary}")
        
        # 5. Data Validations
        validations = sheet1.findall('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}dataValidation')
        print(f"  [VALIDAÇÃO] Regras de validação ativas na aba: {len(validations)}")

def main():
    files = sorted(glob.glob('Monitoramento/novo/*.xlsm'))
    for f in files:
        if 'backup' not in f:
            verify_file(f)

if __name__ == '__main__':
    main()
