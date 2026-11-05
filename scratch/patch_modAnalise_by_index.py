# -*- coding: cp1252 -*-
import os

with open(r'Codes\VBA\modAnaliseMonitoramento.bas', 'rb') as f:
    raw = f.read()

text = raw.decode('cp1252').replace('\r\n', '\n')
print(f"Loaded: {len(text)} chars, {text.count(chr(10))} lines")

# 1. Header
pos = text.find('exatamente na coluna Z ("LOG") da aba "Monitoramento".\n')
assert pos != -1, "header target not found"
ins_header = """'   3. CONSISTÊNCIA CONTRATO x LOCALIDADE (TABELAB):
'      Valida se o contrato informado na Coluna D condiz com a localidade da Coluna V,
'      resolvendo a cadeia Localidade -> Centro -> UF -> Contrato homologado na TabelaB.
"""
idx_after_header = pos + len('exatamente na coluna Z ("LOG") da aba "Monitoramento".\n')
text = text[:idx_after_header] + ins_header + text[idx_after_header:]
print("1. Header updated.")

# 2. Add dict declarations
anchor = 'Dim dictPrazoAtiv   As Object'
pos = text.find(anchor)
assert pos != -1, "anchor 1 not found"
end_line = text.find('\n', pos) + 1
ins_decl = """    Dim dictLocContrato As Object ' Localidade -> Contratos permitidos (ex: "|IRON-LT1|4600686987|")
    Dim dictLocInfo     As Object ' Localidade -> "UF|Centro|ContratoPrincipal"
"""
text = text[:end_line] + ins_decl + text[end_line:]
print("2. Dict declarations added.")

# 3. Add dict init
anchor = 'Set dictPrazoAtiv = CreateObject("Scripting.Dictionary")\n'
pos = text.find(anchor)
assert pos != -1, "anchor 2 not found"
end_line = pos + len(anchor)
ins_init = """    Set dictLocContrato = CreateObject("Scripting.Dictionary")
    Set dictLocInfo = CreateObject("Scripting.Dictionary")
"""
text = text[:end_line] + ins_init + text[end_line:]

anchor = 'dictPrazoAtiv.CompareMode = vbTextCompare\n'
pos = text.find(anchor)
assert pos != -1, "anchor 3 not found"
end_line = pos + len(anchor)
ins_comp = """    dictLocContrato.CompareMode = vbTextCompare
    dictLocInfo.CompareMode = vbTextCompare
"""
text = text[:end_line] + ins_comp + text[end_line:]
print("3. Dict init & compare added.")

# 4. Pass new dicts to CarregarDicionariosReferencia
anchor = 'dictSiglasObs, dictLocalidades, dictPrazoAtiv\n'
pos = text.find(anchor)
assert pos != -1, "anchor 4 not found"
repl_call = 'dictSiglasObs, dictLocalidades, dictPrazoAtiv, _\n                                  dictLocContrato, dictLocInfo\n'
text = text[:pos] + repl_call + text[pos + len(anchor):]
print("4. Carregar call updated.")

# 5. Pass new dicts to AuditarLinhaMonitoramento
anchor = 'dictSiglasObs, dictLocalidades, dictPrazoAtiv\n'
pos = text.find(anchor)
assert pos != -1, "anchor 5 not found"
repl_call = 'dictSiglasObs, dictLocalidades, dictPrazoAtiv, _\n                                  dictLocContrato, dictLocInfo\n'
text = text[:pos] + repl_call + text[pos + len(anchor):]
print("5. Auditar call updated.")

# 6. Update signature of AuditarLinhaMonitoramento
anchor = "Private Sub AuditarLinhaMonitoramento("
pos = text.find(anchor)
assert pos != -1, "anchor 6 not found"
anchor_end = "ByRef dictPrazoAtiv As Object)\n"
pos_end = text.find(anchor_end, pos)
assert pos_end != -1, "anchor 6 end not found"
repl_sig = "ByRef dictPrazoAtiv As Object, _\n                                     ByRef dictLocContrato As Object, ByRef dictLocInfo As Object)\n"
text = text[:pos_end] + repl_sig + text[pos_end + len(anchor_end):]
print("6. Auditar signature updated.")

# 7. Add validation logic for Contrato x Localidade in AuditarLinhaMonitoramento
anchor = "    '--- 23. COLUNA Y: OS DISPONIBILIZADA? ---"
pos = text.find(anchor)
assert pos != -1, "anchor 7 not found"
ins_val = """    '--- 22b. COLUNAS D e V: CONSISTÊNCIA CONTRATO x LOCALIDADE (TABELAB) ---
    If Len(sContrato) > 0 And Len(sLocalidade) > 0 Then
        Dim sLocBusca As String
        sLocBusca = UCase(Trim(Replace(sLocalidade, "  ", " ")))
        
        If dictLocContrato.Count > 0 Then
            If dictLocContrato.Exists(sLocBusca) Then
                Dim sContratosValidos As String
                sContratosValidos = dictLocContrato(sLocBusca)
                
                Dim sContratoTestar As String
                sContratoTestar = "|" & UCase(Trim(sContrato)) & "|"
                
                If InStr(1, sContratosValidos, sContratoTestar, vbTextCompare) = 0 Then
                    Dim sDetalhesLoc As String
                    Dim vPartesInfo As Variant
                    Dim sUFEsperada As String
                    Dim sContratoEsperado As String
                    
                    sDetalhesLoc = dictLocInfo(sLocBusca)
                    vPartesInfo = Split(sDetalhesLoc, "|")
                    If UBound(vPartesInfo) >= 2 Then
                        sUFEsperada = vPartesInfo(0)
                        sContratoEsperado = vPartesInfo(2)
                    Else
                        sUFEsperada = "-"
                        sContratoEsperado = "homologado na TabelaB"
                    End If
                    
                    sLog = sLog & "(Coluna D e V): Localidade '" & sLocalidade & "' pertence ao contrato '" & _
                                  sContratoEsperado & "' (UF: " & sUFEsperada & "), mas o contrato informado é '" & _
                                  sContrato & "'; "
                End If
            Else
                ' Se a localidade não consta na TabelaB (mas informada)
                sLog = sLog & "(Coluna D e V): Localidade '" & sLocalidade & "' não cadastrada na TabelaB para validação de contrato; "
            End If
        End If
    End If
    
"""
text = text[:pos] + ins_val + text[pos:]
print("7. Contrato x Localidade validation added.")

# 8. Update signature of CarregarDicionariosReferencia
anchor = "Private Sub CarregarDicionariosReferencia("
pos = text.find(anchor)
assert pos != -1, "anchor 8 not found"
anchor_end = "ByRef dictPrazoAtiv As Object)\n"
pos_end = text.find(anchor_end, pos)
assert pos_end != -1, "anchor 8 end not found"
repl_sig = "ByRef dictPrazoAtiv As Object, _\n                                         ByRef dictLocContrato As Object, ByRef dictLocInfo As Object)\n"
text = text[:pos_end] + repl_sig + text[pos_end + len(anchor_end):]
print("8. Carregar signature updated.")

# 9. Add TabelaB loading logic in CarregarDicionariosReferencia
anchor = 'dictContratoUF("4600687006|BA") = True\n'
pos = text.find(anchor)
assert pos != -1, "anchor 9 not found"
end_line = pos + len(anchor)
ins_tabelab = """
    ' =====================================================================================
    ' CARREGAMENTO DA TABELAB: Mapeamento Localidade -> Centro -> UF -> Contrato
    ' Valida se o contrato informado condiz com a localidade através da cadeia:
    ' 1. Localidade (Col 36) -> Centro (Col 37)
    ' 2. Centro (Col 38) -> UF (Col 45) e Município (Col 44)
    ' 3. Grupo (Cols 56 e 60) e Contrato (Cols 1 e 4): UF -> Contrato(s) homologado(s)
    ' =====================================================================================
    Dim wsTabelaB As Worksheet
    On Error Resume Next
    Set wsTabelaB = wsRef.Parent.Sheets("TabelaB")
    If wsTabelaB Is Nothing Then Set wsTabelaB = wsRef.Parent.Sheets("tabelaB")
    On Error GoTo 0
    
    If Not wsTabelaB Is Nothing Then
        Dim ulB As Long
        ulB = wsTabelaB.UsedRange.Rows.Count
        If ulB >= 2 Then
            Dim vTB As Variant
            Dim numColsB As Long
            numColsB = wsTabelaB.UsedRange.Columns.Count
            If numColsB < 60 Then numColsB = 65
            
            vTB = wsTabelaB.Range(wsTabelaB.Cells(1, 1), wsTabelaB.Cells(ulB, numColsB)).Value
            
            ' Dicionários temporários para resolução de relacionamentos
            Dim dictCentroToUF As Object
            Dim dictLocToCentro As Object
            Dim dictUFToContratos As Object
            Dim dictSAPToContrato As Object
            
            Set dictCentroToUF = CreateObject("Scripting.Dictionary")
            Set dictLocToCentro = CreateObject("Scripting.Dictionary")
            Set dictUFToContratos = CreateObject("Scripting.Dictionary")
            Set dictSAPToContrato = CreateObject("Scripting.Dictionary")
            
            dictCentroToUF.CompareMode = vbTextCompare
            dictLocToCentro.CompareMode = vbTextCompare
            dictUFToContratos.CompareMode = vbTextCompare
            dictSAPToContrato.CompareMode = vbTextCompare
            
            ' Equivalências contratuais SAP <-> Sigla
            dictSAPToContrato("4600686987") = "IRON-LT1"
            dictSAPToContrato("4600687006") = "PA-LT2"
            dictSAPToContrato("4600617111") = "T2M"
            dictSAPToContrato("IRON-LT1") = "4600686987"
            dictSAPToContrato("PA-LT2") = "4600687006"
            dictSAPToContrato("T2M") = "4600617111"
            
            Dim rB As Long
            Dim sCen As String, sUFCen As String
            Dim sLocNome As String, sCenLoc As String
            Dim sUFGrupo As String, sContratoGrupo As String
            
            ' Passo 1: Carrega Centro -> UF (Cols 38 e 45)
            For rB = 2 To UBound(vTB, 1)
                sCen = UCase(Trim(Replace(CStr(vTB(rB, 38) & ""), "  ", " ")))
                sUFCen = UCase(Trim(CStr(vTB(rB, 45) & "")))
                If Len(sCen) > 0 And Len(sUFCen) > 0 Then
                    If Not dictCentroToUF.Exists(sCen) Then
                        dictCentroToUF(sCen) = sUFCen
                    End If
                End If
            Next rB
            
            ' Passo 2: Carrega Grupo: UF -> Contrato (Cols 56 e 60)
            For rB = 2 To UBound(vTB, 1)
                sUFGrupo = UCase(Trim(CStr(vTB(rB, 56) & "")))
                sContratoGrupo = UCase(Trim(CStr(vTB(rB, 60) & "")))
                If Len(sUFGrupo) > 0 And Len(sContratoGrupo) > 0 Then
                    If Not dictUFToContratos.Exists(sUFGrupo) Then
                        dictUFToContratos(sUFGrupo) = sContratoGrupo
                    Else
                        If InStr(1, dictUFToContratos(sUFGrupo), sContratoGrupo, vbTextCompare) = 0 Then
                            dictUFToContratos(sUFGrupo) = dictUFToContratos(sUFGrupo) & "|" & sContratoGrupo
                        End If
                    End If
                End If
            Next rB
            
            ' Fallbacks regionais contratuais homologados Petrobras
            If Not dictUFToContratos.Exists("RJ") Then dictUFToContratos("RJ") = "IRON-LT1"
            If Not dictUFToContratos.Exists("SP") Then dictUFToContratos("SP") = "IRON-LT1"
            If Not dictUFToContratos.Exists("ES") Then dictUFToContratos("ES") = "IRON-LT1"
            If Not dictUFToContratos.Exists("BA") Then dictUFToContratos("BA") = "PA-LT2"
            If Not dictUFToContratos.Exists("SE") Then dictUFToContratos("SE") = "PA-LT2"
            If Not dictUFToContratos.Exists("RN") Then dictUFToContratos("RN") = "PA-LT2"
            If Not dictUFToContratos.Exists("AM") Then dictUFToContratos("AM") = "PA-LT2"
            
            ' Passo 3: Carrega Localidade -> Centro (Cols 36 e 37)
            For rB = 2 To UBound(vTB, 1)
                sLocNome = UCase(Trim(Replace(CStr(vTB(rB, 36) & ""), "  ", " ")))
                sCenLoc = UCase(Trim(Replace(CStr(vTB(rB, 37) & ""), "  ", " ")))
                If Len(sLocNome) > 0 Then
                    If Len(sCenLoc) > 0 Then
                        dictLocToCentro(sLocNome) = sCenLoc
                    Else
                        dictLocToCentro(sLocNome) = sLocNome
                    End If
                End If
            Next rB
            
            ' Passo 4: Mapeia todas as Localidades cadastradas
            Dim vLocKeys As Variant
            Dim iL As Long
            Dim sTargetLoc As String, sTargetCen As String, sTargetUF As String
            Dim sContratosPerm As String, sContratoPrin As String
            
            vLocKeys = dictLocToCentro.Keys
            For iL = 0 To UBound(vLocKeys)
                sTargetLoc = CStr(vLocKeys(iL))
                sTargetCen = dictLocToCentro(sTargetLoc)
                sTargetUF = ""
                
                ' Tenta achar UF pelo Centro mapeado
                If dictCentroToUF.Exists(sTargetCen) Then
                    sTargetUF = dictCentroToUF(sTargetCen)
                ElseIf dictCentroToUF.Exists(sTargetLoc) Then
                    ' A própria localidade é o nome do Centro (ex: PORTO ENCONTRO DAS AGUAS)
                    sTargetUF = dictCentroToUF(sTargetLoc)
                    sTargetCen = sTargetLoc
                End If
                
                If Len(sTargetUF) > 0 Then
                    If dictUFToContratos.Exists(sTargetUF) Then
                        sContratoPrin = dictUFToContratos(sTargetUF)
                        ' Monta string de contratos permitidos para checagem rápida (|IRON-LT1|4600686987|IRON|)
                        sContratosPerm = "|" & sContratoPrin & "|"
                        If sContratoPrin = "IRON-LT1" Then
                            sContratosPerm = sContratosPerm & "4600686987|IRON|"
                        ElseIf sContratoPrin = "PA-LT2" Then
                            sContratosPerm = sContratosPerm & "4600687006|PA|"
                        ElseIf sContratoPrin = "T2M" Then
                            sContratosPerm = sContratosPerm & "4600617111|"
                        End If
                        
                        dictLocContrato(sTargetLoc) = sContratosPerm
                        dictLocInfo(sTargetLoc) = sTargetUF & "|" & sTargetCen & "|" & sContratoPrin
                    End If
                End If
            Next iL
            
            ' Passo 5: Registra também todos os Centros diretamente como Localidades
            Dim vCenKeys As Variant
            Dim iC As Long
            vCenKeys = dictCentroToUF.Keys
            For iC = 0 To UBound(vCenKeys)
                sTargetCen = CStr(vCenKeys(iC))
                If Not dictLocContrato.Exists(sTargetCen) Then
                    sTargetUF = dictCentroToUF(sTargetCen)
                    If dictUFToContratos.Exists(sTargetUF) Then
                        sContratoPrin = dictUFToContratos(sTargetUF)
                        sContratosPerm = "|" & sContratoPrin & "|"
                        If sContratoPrin = "IRON-LT1" Then
                            sContratosPerm = sContratosPerm & "4600686987|IRON|"
                        ElseIf sContratoPrin = "PA-LT2" Then
                            sContratosPerm = sContratosPerm & "4600687006|PA|"
                        End If
                        dictLocContrato(sTargetCen) = sContratosPerm
                        dictLocInfo(sTargetCen) = sTargetUF & "|" & sTargetCen & "|" & sContratoPrin
                    End If
                End If
            Next iC
            
            Set dictCentroToUF = Nothing
            Set dictLocToCentro = Nothing
            Set dictUFToContratos = Nothing
            Set dictSAPToContrato = Nothing
        End If
    End If
"""
text = text[:end_line] + ins_tabelab + text[end_line:]
print("9. TabelaB block added.")

# Ensure CRLF line endings
text_crlf = text.replace('\n', '\r\n')

# Output paths
output_paths = [
    r'Codes\VBA\modAnaliseMonitoramento.bas',
    r'Codes\VBA\Funções Novas\modAnaliseMonitoramento.bas',
    r'Codes\VBA\vba_extracted_A4UU\modAnaliseMonitoramento.bas'
]

for p in output_paths:
    os.makedirs(os.path.dirname(p), exist_ok=True)
    with open(p, 'wb') as f:
        f.write(text_crlf.encode('cp1252'))
    print(f"Written: {p} ({len(text_crlf)} chars, {text_crlf.count(chr(10))} lines)")

print("ALL .BAS FILES UPDATED SUCCESSFULLY!")

