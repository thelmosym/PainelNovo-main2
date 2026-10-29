# -*- coding: utf-8 -*-
import os

target_old = '''    try
        fnClassificarFDM(
            [CalculoFDM], [#"Ref. FDM"], [#"Ref. prazo"],
            [#"Descrição da atividade"], [Contrato],
            [Data Abertura], [Hora Abertura],
            [D. fechamento], [Prazo combinado],
            [Município], [UF], [Grupo.1],
            [#"tabelaPrazoContrato.Prazo"], [Aplicação], RefCalendario
        )
    otherwise
        [ Classificacao = null, StatusPrazo = null, Log = "Erro no calculo - verificar dados da linha; " ]'''

replacement = '''    let
        res = try fnClassificarFDM(
            [CalculoFDM], [#"Ref. FDM"], [#"Ref. prazo"],
            [#"Descrição da atividade"], [Contrato],
            [Data Abertura], [Hora Abertura],
            [D. fechamento], [Prazo combinado],
            [Município], [UF], [Grupo.1],
            [#"tabelaPrazoContrato.Prazo"], [Aplicação], RefCalendario
        )
    in
        if res[HasError] then
            [ Classificacao = null, StatusPrazo = null, Log = "Erro no calculo: " & (res[Error][Message] ?? "falha no processamento M") & "; " ]
        else
            res[Value]'''

for fname in ['Codes/CodigosM/Painel T2M.m', 'Codes/CodigosM/Painel otimizado.m']:
    if os.path.exists(fname):
        with open(fname, 'r', encoding='utf-8', errors='ignore') as f:
            content = f.read()
        
        # Look for the try fnClassificarFDM block
        idx_start = content.find('try\n        fnClassificarFDM(')
        if idx_start == -1:
            idx_start = content.find('try\r\n        fnClassificarFDM(')
        
        if idx_start != -1:
            idx_end = content.find('type record', idx_start)
            if idx_end != -1:
                # Find otherwise line
                block = content[idx_start:idx_end]
                # Replace with let res = try ... in if ...
                # extract parameters
                fn_call_start = block.find('fnClassificarFDM(')
                fn_call_end = block.find(')', fn_call_start) + 1
                fn_call = block[fn_call_start:fn_call_end]
                
                new_block = f'''let
        res = try {fn_call}
    in
        if res[HasError] then
            [ Classificacao = null, StatusPrazo = null, Log = "Erro no calculo: " & (res[Error][Message] ?? "falha no processamento M") & "; " ]
        else
            res[Value],
    '''
                content = content[:idx_start] + new_block + content[idx_end:]
                with open(fname, 'w', encoding='utf-8') as f:
                    f.write(content)
                print(f'Updated {fname}')
            else:
                print(f'Could not find type record in {fname}')
        else:
            print(f'Could not find try fnClassificarFDM in {fname}')

