import subprocess

base_text = subprocess.check_output(['git', 'show', 'f72cb3c:Codes/VBA/modDemonstrativoFDM.bas'], encoding='latin1')

idx_co = base_text.find('If sSit = "CO" Then')
print('Found sSit == CO at:', idx_co)
print(repr(base_text[idx_co:idx_co+300]))

idx_base16 = base_text.find('vBaseSaida(nOrdensBase, 16) =')
print('Found vBaseSaida 16 at:', idx_base16)
print(repr(base_text[idx_base16-50:idx_base16+500]))

