import os, re

codigos_m_dir = r'd:\PainelNovo-main\Codes\CodigosM'
with open(r'd:\PainelNovo-main\scratch\Section1_agosto.m', 'r', encoding='utf-8') as f:
    agosto_m = f.read()

parts = re.split(r'\n(?=shared\s+)', agosto_m)
agosto_queries = {}
for p in parts:
    m = re.match(r'shared\s+(#?"?[^\s=]+"?)\s*=\s*(.*)', p, re.DOTALL)
    if m:
        name = m.group(1).strip('"#')
        agosto_queries[name] = m.group(2).strip()

print(f'Total queries in agosto: {len(agosto_queries)}')

for f in sorted(os.listdir(codigos_m_dir)):
    if f.endswith('.m'):
        qname = f[:-2].strip('"#')
        with open(os.path.join(codigos_m_dir, f), 'r', encoding='utf-8') as mf:
            code = mf.read().strip()
        if qname in agosto_queries:
            c1 = re.sub(r'\s+', ' ', code)
            c2 = re.sub(r'\s+', ' ', agosto_queries[qname])
            if c1 != c2:
                print(f'MISMATCH in {qname}:')
                print(f'   Codes/CodigosM len: {len(code)}, Agosto len: {len(agosto_queries[qname])}')
        else:
            print(f'Query {qname} NOT found in agosto workbook')
