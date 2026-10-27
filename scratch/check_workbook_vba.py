import zipfile

with zipfile.ZipFile('Monitoramento/Monitoramento individual_A4UU_v4.xlsm') as z:
    names = z.namelist()
    print("Files in xlsm:", [n for n in names if 'vba' in n.lower() or 'bin' in n.lower()])
    if 'xl/vbaProject.bin' in names:
        data = z.read('xl/vbaProject.bin')
        print("vbaProject.bin size:", len(data))
        # Let's check for module names
        for m in [b'modLimparMonitoramento', b'LimparMonitoramento', b'LimparAbaMonitoramento', b'modAnaliseMonitoramento']:
            print(f"Contains {m.decode()}: {m in data}")
