import zipfile, xml.etree.ElementTree as ET
from collections import defaultdict

with zipfile.ZipFile('Painel de controle MemoriaPetrobras V8.xlsm') as z:
    sst = []
    if 'xl/sharedStrings.xml' in z.namelist():
        sst_root = ET.fromstring(z.read('xl/sharedStrings.xml'))
        for si in sst_root.findall('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}si'):
            sst.append(''.join([e.text for e in si.iter('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t') if e.text]))
            
    sheet_t2m = ET.fromstring(z.read('xl/worksheets/sheet31.xml'))
    rows = sheet_t2m.findall('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}row')
    
    stats = defaultdict(lambda: {'total': 0, 'NP': 0, 'FP': 0, 'outros': 0})
    for r in rows[1:]:
        cells = {}
        for c in r.findall('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}c'):
            ref = c.attrib.get('r')
            col_letter = ''.join([ch for ch in ref if ch.isalpha()])
            v = c.find('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}v')
            val = v.text if v is not None else ''
            if c.attrib.get('t') == 's' and val != '':
                val = sst[int(val)]
            cells[col_letter] = val
        
        contrato = cells.get('C', '')
        ref_fdm = cells.get('AH', '')
        atividade = cells.get('D', '')
        status_prazo = cells.get('AT', '')
        
        key = (contrato, ref_fdm, atividade)
        stats[key]['total'] += 1
        if status_prazo == 'NP':
            stats[key]['NP'] += 1
        elif status_prazo == 'FP':
            stats[key]['FP'] += 1
        else:
            stats[key]['outros'] += 1

print(f"{'Contrato':<10} | {'Ref':<4} | {'Atividade':<45} | {'Total':<5} | {'NP':<4} | {'FP':<4} | {'Outros':<6} | {'TSPE':<5} | {'% Prazo':<8}")
print('-'*110)
for (contrato, ref_fdm, atividade), val in sorted(stats.items()):
    tspe = val['NP'] + val['FP']
    pct = (val['NP'] / tspe * 100) if tspe > 0 else 0
    print(f"{contrato:<10} | {ref_fdm:<4} | {atividade[:45]:<45} | {val['total']:>5} | {val['NP']:>4} | {val['FP']:>4} | {val['outros']:>6} | {tspe:>5} | {pct:>7.2f}%")
