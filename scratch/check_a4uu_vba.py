import zipfile

with zipfile.ZipFile(r'Monitoramento\Monitoramento individual_A4UU_v4.xlsm', 'r') as z:
    vba = z.read('xl/vbaProject.bin')
    print('Planilha1 in A4UU vbaProject:', b'Planilha1' in vba)
    print('analiseMonitoramento in A4UU vbaProject:', b'analiseMonitoramento' in vba)

