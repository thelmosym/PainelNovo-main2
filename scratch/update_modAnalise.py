import subprocess
import re

# Read baseline
raw = subprocess.check_output(
    ['powershell', '-ExecutionPolicy', 'Bypass', '-Command',
     "Get-Content 'D:\\PainelNovo-main\\Codes\\VBA\\modAnaliseMonitoramento.bas' -Encoding Default -Raw"],
    encoding='latin1'
)

print(f"Original length: {len(raw)} chars, {raw.count(chr(10))} lines")

# ============================================================
# CHANGE 1: Add dictContratoLocalidade declaration
# ============================================================
# After: Dim dictPrazoAtiv   As Object ' "Fiscalização" ou "Tabela"
old_dict_decl = """    Dim dictPrazoAtiv   As Object ' "Fiscaliza\xe7\xe3o" ou "Tabela"
"""

new_dict_decl = """    Dim dictPrazoAtiv   As Object ' "Fiscaliza\xe7\xe3o" ou "Tabela"
    Dim dictContratoLoc As Object ' Pares Contrato|Localidade v\xe1lidos (da TabelaB)
"""

assert old_dict_decl in raw, "CHANGE 1: dict declaration target not found"
code = raw.replace(old_dict_decl, new_dict_decl, 1)

# ============================================================
# CHANGE 2: Initialize the dictionary (Set + CompareMode)
# ============================================================
old_init = """    Set dictPrazoAtiv = CreateObject("Scripting.Dictionary")
"""

new_init = """    Set dictPrazoAtiv = CreateObject("Scripting.Dictionary")
    Set dictContratoLoc = CreateObject("Scripting.Dictionary")
"""

assert old_init in code, "CHANGE 2a: dict init target not found"
code = code.replace(old_init, new_init, 1)

old_compare = """    dictPrazoAtiv.CompareMode = vbTextCompare
"""

new_compare = """    dictPrazoAtiv.CompareMode = vbTextCompare
    dictContratoLoc.CompareMode = vbTextCompare
"""

assert old_compare in code, "CHANGE 2b: compare mode target not found"
code = code.replace(old_compare, new_compare, 1)

# ============================================================
# CHANGE 3: Pass dictContratoLocalidade to CarregarDicionariosReferencia
# ============================================================
old_carregar_call = """    CarregarDicionariosReferencia wsRef, dictSituacoes, dictContratos, dictUFs, dictContratoUF, _
                                  dictAtividades, dictItens, dictAtivItem, dictAplicacoes, _
                                  dictSiglasObs, dictLocalidades, dictPrazoAtiv"""

new_carregar_call = """    CarregarDicionariosReferencia wsRef, dictSituacoes, dictContratos, dictUFs, dictContratoUF, _
                                  dictAtividades, dictItens, dictAtivItem, dictAplicacoes, _
                                  dictSiglasObs, dictLocalidades, dictPrazoAtiv, dictContratoLoc"""

assert old_carregar_call in code, "CHANGE 3: carregar call target not found"
code = code.replace(old_carregar_call, new_carregar_call, 1)

# ============================================================
# CHANGE 4: Pass dictContratoLoc to AuditarLinhaMonitoramento
# ============================================================
old_auditar_call = """        AuditarLinhaMonitoramento vDados, r, sLogLinha, bLinhaAuditada, _
                                  dictSituacoes, dictContratos, dictUFs, dictContratoUF, _
                                  dictAtividades, dictItens, dictAtivItem, dictAplicacoes, _
                                  dictSiglasObs, dictLocalidades, dictPrazoAtiv"""

new_auditar_call = """        AuditarLinhaMonitoramento vDados, r, sLogLinha, bLinhaAuditada, _
                                  dictSituacoes, dictContratos, dictUFs, dictContratoUF, _
                                  dictAtividades, dictItens, dictAtivItem, dictAplicacoes, _
                                  dictSiglasObs, dictLocalidades, dictPrazoAtiv, dictContratoLoc"""

assert old_auditar_call in code, "CHANGE 4: auditar call target not found"
code = code.replace(old_auditar_call, new_auditar_call, 1)

# ============================================================
# CHANGE 5: Update AuditarLinhaMonitoramento signature
# ============================================================
old_auditar_sig = """                                     ByRef dictSiglasObs As Object, ByRef dictLocalidades As Object, _
                                     ByRef dictPrazoAtiv As Object)"""

new_auditar_sig = """                                     ByRef dictSiglasObs As Object, ByRef dictLocalidades As Object, _
                                     ByRef dictPrazoAtiv As Object, ByRef dictContratoLoc As Object)"""

assert old_auditar_sig in code, "CHANGE 5: auditar signature target not found"
code = code.replace(old_auditar_sig, new_auditar_sig, 1)

# ============================================================
# CHANGE 6: Add validation Contrato x Localidade (after section 22 - Localidade)
# ============================================================
# Find the end of the localidade section (after "End If" of section 22, before section 23)
old_after_localidade = """    '--- 23. COLUNA Y: OS DISPONIBILIZADA? ---"""

new_validation_contrato_loc = """    '--- 22b. COLUNAS D e V: CONSIST\xCANCIA CONTRATO x LOCALIDADE (TabelaB) ---
    If Len(sContrato) > 0 And Len(sLocalidade) > 0 Then
        Dim sLocLimpo As String
        sLocLimpo = Replace(sLocalidade, "  ", " ")
        If dictContratoLoc.Count > 0 Then
            ' Verifica se o par Contrato|Localidade existe na TabelaB
            If Not dictContratoLoc.Exists(sContrato & "|" & sLocLimpo) Then
                ' Busca o contrato esperado para esta localidade
                Dim sContratoEsperado As String
                sContratoEsperado = ""
                Dim vChaves As Variant
                Dim iK As Long
                vChaves = dictContratoLoc.Keys
                For iK = 0 To UBound(vChaves)
                    If Right(CStr(vChaves(iK)), Len(sLocLimpo) + 1) = "|" & sLocLimpo Then
                        Dim iPipe As Long
                        iPipe = InStr(CStr(vChaves(iK)), "|")
                        If iPipe > 0 Then
                            sContratoEsperado = Left(CStr(vChaves(iK)), iPipe - 1)
                        End If
                        Exit For
                    End If
                Next iK
                If Len(sContratoEsperado) > 0 Then
                    sLog = sLog & "(Coluna D e V): Localidade '" & sLocLimpo & "' pertence ao contrato '" & sContratoEsperado & "', mas contrato informado e '" & sContrato & "'; "
                Else
                    sLog = sLog & "(Coluna D e V): Localidade '" & sLocLimpo & "' nao encontrada na TabelaB para validacao de contrato; "
                End If
            End If
        End If
    End If
    
    '--- 23. COLUNA Y: OS DISPONIBILIZADA? ---"""

assert old_after_localidade in code, "CHANGE 6: after localidade target not found"
code = code.replace(old_after_localidade, new_validation_contrato_loc, 1)

# ============================================================
# CHANGE 7: Update CarregarDicionariosReferencia signature
# ============================================================
old_carregar_sig = """                                         ByRef dictSiglasObs As Object, ByRef dictLocalidades As Object, _
                                         ByRef dictPrazoAtiv As Object)"""

new_carregar_sig = """                                         ByRef dictSiglasObs As Object, ByRef dictLocalidades As Object, _
                                         ByRef dictPrazoAtiv As Object, ByRef dictContratoLoc As Object)"""

assert old_carregar_sig in code, "CHANGE 7: carregar signature target not found"
code = code.replace(old_carregar_sig, new_carregar_sig, 1)

# ============================================================
# CHANGE 8: Add TabelaB loading logic inside CarregarDicionariosReferencia
# After the End Sub of CarregarDicionariosReferencia, we need to add the 
# TabelaB loading BEFORE End Sub. Find the right anchor point.
# ============================================================
old_end_carregar = """    dictContratoUF("4600686987|RJ") = True
    dictContratoUF("4600686987|SP") = True
    dictContratoUF("4600686987|ES") = True
    dictContratoUF("4600687006|BA") = True

End Sub"""

new_end_carregar = """    dictContratoUF("4600686987|RJ") = True
    dictContratoUF("4600686987|SP") = True
    dictContratoUF("4600686987|ES") = True
    dictContratoUF("4600687006|BA") = True
    
    ' =====================================================================================
    ' CARREGAMENTO DA TABELAB: Mapeamento Localidade -> Centro -> UF -> Contrato
    ' Objetivo: Validar se o contrato informado e compativel com a localidade via TabelaB
    ' Cadeia: Localidade (col 36) -> Centro (col 37) -> UF do Centro (col 45)
    '         Grupo: UF (col 56) + Contrato (col 60) -> par valido
    ' =====================================================================================
    On Error Resume Next
    Dim wsTabelaB As Worksheet
    Set wsTabelaB = wsRef.Parent.Sheets("TabelaB")
    On Error GoTo 0
    
    If Not wsTabelaB Is Nothing Then
        Dim ulB As Long
        ulB = wsTabelaB.UsedRange.Rows.Count
        If ulB > 1 Then
            Dim vTB As Variant
            ' Le as colunas relevantes da TabelaB (ate col 65)
            vTB = wsTabelaB.Range(wsTabelaB.Cells(1, 1), wsTabelaB.Cells(ulB, 65)).Value
            
            ' Passo 1: Montar dicionario Centro -> UF (cols 38 e 45)
            Dim dictCentroUF As Object
            Set dictCentroUF = CreateObject("Scripting.Dictionary")
            dictCentroUF.CompareMode = vbTextCompare
            
            Dim rB As Long
            Dim sCentro As String, sUFCentro As String
            For rB = 2 To UBound(vTB, 1)
                sCentro = Trim(CStr(vTB(rB, 38) & ""))
                sUFCentro = UCase(Trim(CStr(vTB(rB, 45) & "")))
                If Len(sCentro) > 0 And Len(sUFCentro) > 0 Then
                    If Not dictCentroUF.Exists(sCentro) Then
                        dictCentroUF.Add sCentro, sUFCentro
                    End If
                End If
            Next rB
            
            ' Passo 2: Montar dicionario UF|Contrato validos (Grupo: cols 56 e 60)
            Dim dictGrupoContratoUF As Object
            Set dictGrupoContratoUF = CreateObject("Scripting.Dictionary")
            dictGrupoContratoUF.CompareMode = vbTextCompare
            
            Dim sUFGrupo As String, sContratoGrupo As String
            For rB = 2 To UBound(vTB, 1)
                sUFGrupo = UCase(Trim(CStr(vTB(rB, 56) & "")))
                sContratoGrupo = Trim(CStr(vTB(rB, 60) & ""))
                If Len(sUFGrupo) > 0 And Len(sContratoGrupo) > 0 Then
                    dictGrupoContratoUF(sContratoGrupo & "|" & sUFGrupo) = True
                End If
            Next rB
            
            ' Passo 3: Para cada Localidade, resolver Centro -> UF -> Contratos validos
            Dim sLoc As String, sCentroLoc As String, sUFLoc As String
            For rB = 2 To UBound(vTB, 1)
                sLoc = Trim(CStr(vTB(rB, 36) & ""))
                sCentroLoc = Trim(CStr(vTB(rB, 37) & ""))
                If Len(sLoc) > 0 And Len(sCentroLoc) > 0 Then
                    ' Resolve UF do Centro
                    If dictCentroUF.Exists(sCentroLoc) Then
                        sUFLoc = dictCentroUF(sCentroLoc)
                        ' Para cada contrato valido nessa UF, registra par Contrato|Localidade
                        Dim vGrupoKeys As Variant
                        Dim iG As Long
                        vGrupoKeys = dictGrupoContratoUF.Keys
                        For iG = 0 To UBound(vGrupoKeys)
                            If Right(CStr(vGrupoKeys(iG)), Len(sUFLoc) + 1) = "|" & sUFLoc Then
                                Dim iPipeG As Long
                                iPipeG = InStr(CStr(vGrupoKeys(iG)), "|")
                                If iPipeG > 0 Then
                                    Dim sContratoValido As String
                                    sContratoValido = Left(CStr(vGrupoKeys(iG)), iPipeG - 1)
                                    dictContratoLoc(sContratoValido & "|" & sLoc) = True
                                End If
                            End If
                        Next iG
                    End If
                End If
            Next rB
            
            Set dictCentroUF = Nothing
            Set dictGrupoContratoUF = Nothing
        End If
    End If

End Sub"""

assert old_end_carregar in code, "CHANGE 8: end carregar target not found"
code = code.replace(old_end_carregar, new_end_carregar, 1)

# Verify no duplicate procedures
proc_pattern = re.compile(r'^(?:Public |Private )?(?:Sub|Function) (\w+)', re.MULTILINE)
matches = proc_pattern.findall(code)
seen = set()
duplicates = []
for m in matches:
    if m in seen:
        duplicates.append(m)
    seen.add(m)

print(f"Procedures found: {len(matches)}")
print(f"Distinct: {len(seen)}")
if duplicates:
    print(f"WARNING: Duplicate procedures: {duplicates}")
    exit(1)
else:
    print("NO duplicate procedures!")

# Ensure CRLF
final = code.replace('\r\n', '\n').replace('\n', '\r\n')

# Write to both locations
import os
paths = [
    r'Codes\VBA\modAnaliseMonitoramento.bas',
    r'Codes\VBA\Funções Novas\modAnaliseMonitoramento.bas',
]
for p in paths:
    os.makedirs(os.path.dirname(p), exist_ok=True)
    with open(p, 'wb') as f:
        f.write(final.encode('windows-1252'))
    lines_count = final.count('\n')
    print(f"Wrote {p} ({len(final)} chars, {lines_count} lines, Windows-1252, CRLF)")

