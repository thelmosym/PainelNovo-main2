import zipfile, xml.etree.ElementTree as ET

with zipfile.ZipFile('Monitoramento/Monitoramento individual_A4UU_v4.xlsm') as z:
    shared_strings = []
    if 'xl/sharedStrings.xml' in z.namelist():
        sst_root = ET.fromstring(z.read('xl/sharedStrings.xml'))
        for si in sst_root.findall('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}si'):
            text_parts = [elem.text for elem in si.iter('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t') if elem.text]
            shared_strings.append(''.join(text_parts))
    sheet_root = ET.fromstring(z.read('xl/worksheets/sheet1.xml'))
    rows = sheet_root.findall('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}row')
    print(f'Total rows in sheet1 (Monitoramento): {len(rows)}')
    for r in rows:
        r_idx = int(r.attrib.get('r'))
        cells = {}
        for c in r.findall('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}c'):
            ref = c.attrib.get('r')
            cell_type = c.attrib.get('t')
            v = c.find('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}v')
            val = v.text if v is not None else None
            if cell_type == 's' and val is not None:
                val = shared_strings[int(val)]
            cells[ref] = val
        if r_idx == 2:
            print(f'Header Row {r_idx}: Col A={cells.get("A2")}, Col C={cells.get("C2")}, Col D={cells.get("D2")}, Col F={cells.get("F2")}')
        elif r_idx <= 25:
            print(f'Data Row {r_idx}: Situação={cells.get(f"C{r_idx}")}, Contrato={cells.get(f"D{r_idx}")}, Desc={cells.get(f"F{r_idx}")}')
