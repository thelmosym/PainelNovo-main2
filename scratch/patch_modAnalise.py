# -*- coding: cp1252 -*-
import os

with open(r'Codes\VBA\modAnaliseMonitoramento.bas', 'r', encoding='cp1252') as f:
    text = f.read().replace('\r\n', '\n')

print(f"Loaded original text: {len(text)} chars")

# 1. Update header comments
header_target = "exatamente na coluna Z (\"LOG\") da aba \"Monitoramento\".\n"
header_repl = """exatamente na coluna Z ("LOG") da aba "Monitoramento".
'   3. CONSISTÊNCIA CONTRATO x LOCALIDADE (TABELAB):
'      Valida se o contrato informado na Coluna D condiz com a localidade da Coluna V,
'      resolvendo a cadeia Localidade -> Centro -> UF -> Contrato homologado na TabelaB.
"""
assert header_target in text, "header_target failed"
text = text.replace(header_target, header_repl, 1)

# 2. Add dict declarations
decl_target = "    Dim dictPrazoAtiv   As Object ' \"Fiscalização\" ou \"Tabela\"\n"
decl_repl = """    Dim dictPrazoAtiv   As Object ' "Fiscalização" ou "Tabela"
    Dim dictLocContrato As Object ' Localidade -> Contratos permitidos (ex: "|IRON-LT1|4600686987|")
    Dim dictLocInfo     As Object ' Localidade -> "UF|Centro|ContratoPrincipal"
"""
assert decl_target in text, "decl_target failed"
text = text.replace(decl_target, decl_repl, 1)

# 3. Add dict init
init_target = "    Set dictPrazoAtiv = CreateObject(\"Scripting.Dictionary\")\n"
init_repl = """    Set dictPrazoAtiv = CreateObject("Scripting.Dictionary")
    Set dictLocContrato = CreateObject("Scripting.Dictionary")
    Set dictLocInfo = CreateObject("Scripting.Dictionary")
"""
assert init_target in text, "init_target failed"
text = text.replace(init_target, init_repl, 1)

comp_target = "    dictPrazoAtiv.CompareMode = vbTextCompare\n"
comp_repl = """    dictPrazoAtiv.CompareMode = vbTextCompare
    dictLocContrato.CompareMode = vbTextCompare
    dictLocInfo.CompareMode = vbTextCompare
"""
assert comp_target in text, "comp_target failed"
text = text.replace(comp_target, comp_repl, 1)

# 4. Pass new dicts to CarregarDicionariosReferencia
call_carregar_target = "                                  dictSiglasObs, dictLocalidades, dictPrazoAtiv\n"
call_carregar_repl = """                                  dictSiglasObs, dictLocalidades, dictPrazoAtiv, _
                                  dictLocContrato, dictLocInfo
"""
assert call_carregar_target in text, "call_carregar_target failed"
text = text.replace(call_carregar_target, call_carregar_repl, 1)

# 5. Pass new dicts to AuditarLinhaMonitoramento
call_auditar_target = "                                  dictSiglasObs, dictLocalidades, dictPrazoAtiv\n"
call_auditar_repl = """                                  dictSiglasObs, dictLocalidades, dictPrazoAtiv, _
                                  dictLocContrato, dictLocInfo
"""
assert call_auditar_target in text, "call_auditar_target failed"
text = text.replace(call_auditar_target, call_auditar_repl, 1)

# 6. Update signature of AuditarLinhaMonitoramento
sig_auditar_target = "                                     ByRef dictSiglasObs As Object, ByRef dictLocalidades As Object, _\n                                     ByRef dictPrazoAtiv As Object)\n"
sig_auditar_repl = """                                     ByRef dictSiglasObs As Object, ByRef dictLocalidades As Object, _
                                     ByRef dictPrazoAtiv As Object, _
                                     ByRef dictLocContrato As Object, ByRef dictLocInfo As Object)
"""
assert sig_auditar_target in text, "sig_auditar_target failed"
text = text.replace(sig_auditar_target, sig_auditar_repl, 1)

# 7. Add validation logic for Contrato x Localidade in AuditarLinhaMonitoramento
val_loc_target = "    '--- 23. COLUNA Y: OS DISPONIBILIZADA? ---"
val_loc_repl = """    '--- 22b. COLUNAS D e V: CONSISTÊNCIA CONTRATO x LOCALIDADE (TABELAB) ---
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
    
    '--- 23. COLUNA Y: OS DISPONIBILIZADA? ---"""
assert val_loc_target in text, "val_loc_target failed"
text = text.replace(val_loc_target, val_loc_repl, 1)

# 8. Update signature of CarregarDicionariosReferencia
sig_carregar_target = "                                         ByRef dictSiglasObs As Object, ByRef dictLocalidades As Object, _\n                                         ByRef dictPrazoAtiv As Object)\n"
sig_carregar_repl = """                                         ByRef dictSiglasObs As Object, ByRef dictLocalidades As Object, _
                                         ByRef dictPrazoAtiv As Object, _
                                         ByRef dictLocContrato As Object, ByRef dictLocInfo As Object)
"""
assert sig_carregar_target in text, "sig_carregar_target failed"
text = text.replace(sig_carregar_target, sig_carregar_repl, 1)

# 9. Add TabelaB loading logic in CarregarDicionariosReferencia
tabelab_block_target = "dictContratoUF(\"4600687006|BA\") = True\n\nEnd Sub"
tabelab_block_repl = """dictContratoUF("4600687006|BA") = True

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

End Sub"""
assert tabelab_block_target in text, "tabelab_block_target failed"
text = text.replace(tabelab_block_target, tabelab_block_repl, 1)

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

print("SUCCESS: All .bas files updated cleanly!")

