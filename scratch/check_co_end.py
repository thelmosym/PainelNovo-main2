import subprocess

base_text = subprocess.check_output(['git', 'show', 'f72cb3c:Codes/VBA/modDemonstrativoFDM.bas'], encoding='latin1')

idx_base16 = base_text.find('vBaseSaida(nOrdensBase, 16) =')
idx_next = base_text.find('Next r', idx_base16)
print(repr(base_text[idx_base16:idx_next+10]))

