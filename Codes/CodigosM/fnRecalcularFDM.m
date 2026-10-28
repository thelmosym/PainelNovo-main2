let
    // -------------------------------------------------------------------
    // fnRecalcularFDM
    // -------------------------------------------------------------------
    // Equivalente M da função VBA recalcular_FDM (módulo fncTabelaB)
    // - item nº 8.4.3 do contrato de guarda externa Petrobras.
    //
    // O QUE FAZ:
    //   Aplica a regra de DEGRADAÇÃO EM CASCATA sobre o FDM mensal "bruto"
    //   (calculado por fnCalcularFDMMensal), penalizando a REINCIDÊNCIA
    //   de baixo desempenho ao longo de até 3 meses consecutivos:
    //
    //     FDM > 0,97                                  -> mantém o FDM
    //     FDM = 0,97 e FDM(mês-1) >= 0,98              -> mantém 0,97
    //     FDM = 0,97 e FDM(mês-1) = 0,97 e
    //                  FDM(mês-2) > 0,97               -> degrada a 0,96
    //     FDM = 0,97 e FDM(mês-1) = 0,97 e
    //                  FDM(mês-2) <= 0,97              -> degrada a 0,95
    //     FDM = 0,97 e FDM(mês-1) = 0,97 e
    //                  FDM(mês-2) = null (só 1 anterior)-> degrada a 0,96
    //     FDM = 0,97 e FDM(mês-1) = 0,96               -> degrada a 0,95
    //     FDM = 0,97 e FDM(mês-1) = 0,95               -> mantém 0,95 (piso)
    // -------------------------------------------------------------------
    fnRecalcularFDM = (
        sPerVig as text,
        sPerAnt1 as text,
        sPerAnt2 as text,
        sContrato as text,
        dFDM as nullable number,
        iItem as number
    ) as nullable number =>
        if dFDM = null then
            null
        else if dFDM > 0.97 then
            dFDM
        else if dFDM = 0.97 then
            let
                fdmAnt1 = fnBuscarFDMAnterior(sPerAnt1, sContrato, iItem),
                fdmAnt2 = fnBuscarFDMAnterior(sPerAnt2, sContrato, iItem),
                // padroniza ausencia de histórico para 0.0 para manter a regra
                // consistente com o VBA legado e evitar quebra por valores nulos
                vFdmAnt1 = if fdmAnt1 = null then 0 else fdmAnt1,
                vFdmAnt2 = if fdmAnt2 = null then 0 else fdmAnt2
            in
                if vFdmAnt1 >= 0.98 then
                    dFDM
                else if vFdmAnt1 = 0.97 and vFdmAnt2 > 0.97 then
                    0.96
                else if vFdmAnt1 = 0.97 and vFdmAnt2 <= 0.97 then
                    0.95
                else if vFdmAnt1 = 0.96 then
                    0.95
                else if vFdmAnt1 = 0.95 then
                    0.95
                else
                    dFDM
        else
            dFDM
in
    fnRecalcularFDM
