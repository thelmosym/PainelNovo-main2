import subprocess
import os

base_text = subprocess.check_output(['git', 'show', 'f72cb3c:Codes/VBA/modDemonstrativoFDM.bas'], encoding='latin1')
print('Baseline length:', len(base_text))

# Check lines around ExecutarGeracaoFDM
idx_start = base_text.find('Private Sub ExecutarGeracaoFDM(ByVal bSilencioso As Boolean)')
print('Found ExecutarGeracaoFDM at:', idx_start)

# Check lines around Kill
idx_kill = base_text.find('Kill sCaminhoFinalMedicao')
print('Found Kill at:', idx_kill)
print(repr(base_text[idx_kill-60:idx_kill+250]))

# Check lines around TratarErro:
idx_err = base_text.find('TratarErro:')
print('Found TratarErro at:', idx_err)
print(repr(base_text[idx_err:idx_err+150]))

# Check lines around DiagnosticarMotivoErroSLA
idx_diag = base_text.find('Private Function DiagnosticarMotivoErroSLA')
print('Found DiagnosticarMotivoErroSLA at:', idx_diag)

