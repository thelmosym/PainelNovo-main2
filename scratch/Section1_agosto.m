section Section1;

shared ARM = let
    Combinado = Table.Combine({ConsolidadoES, ConsolidadoRJ, ConsolidadoSP, ConsolidadoPA}),

    // Bufferiza o resultado final — evita reabrir os 4 arquivos físicos
    // a cada referência subsequente desta consulta em outras etapas.
    ResultadoFinal = Table.Buffer(Combinado),
    #"Colunas Renomeadas" = Table.RenameColumns(ResultadoFinal,{{"Contrato", "Linha de serviço PPU"}}),
    #"Colunas Removidas" = Table.RemoveColumns(#"Colunas Renomeadas",{"Ø", "A", "L", "P", "cm³ = V1", "cm³ = V2"}),
    #"Tipo Alterado" = Table.TransformColumnTypes(#"Colunas Removidas",{{"FC = V1/V2", type number}}),
    #"Colunas Removidas1" = Table.RemoveColumns(#"Tipo Alterado",{"$ Unitário", "$ Total", "Tabela", "UA"}),
    #"Tipo Alterado1" = Table.TransformColumnTypes(#"Colunas Removidas1",{{"FC = V1/V2", type number}, {"QUAm-1", type number}, {"QBm", type number}, {"QDm", type number}, {"QNm", type number}, {"QUAm", type number}}),
    #"Valor Substituído" = Table.ReplaceValue(#"Tipo Alterado1",null,0,Replacer.ReplaceValue,{"FC = V1/V2", "QUAm-1", "QBm", "QDm", "QNm", "QUAm"}),
    #"Colunas Renomeadas1" = Table.RenameColumns(#"Valor Substituído",{{"FC = V1/V2", "FC"}}),
    #"Colunas Reordenadas" = Table.ReorderColumns(#"Colunas Renomeadas1",{"Contrato-Reg.", "MEDIDAS CONTRATO", "Padrão", "QUAm-1", "QBm", "QDm", "QNm", "QUAm", "FC", "Linha de serviço PPU"}),
    // QUAm-1 - (QBm + QDm) + QNm
    // 
    QUA = Table.AddColumn(#"Colunas Reordenadas", "QUA", each [#"QUAm-1"]-([QBm]+[QDm]) + [QNm]),
    #"Colunas Reordenadas1" = Table.ReorderColumns(QUA,{"Contrato-Reg.", "MEDIDAS CONTRATO", "Padrão", "QUAm-1", "QBm", "QDm", "QNm", "QUAm", "FC", "QUA", "Linha de serviço PPU"}),
    #"Colunas Renomeadas2" = Table.RenameColumns(#"Colunas Reordenadas1",{{"Padrão", "Item"}})
in
    #"Colunas Renomeadas2"
;

shared localTabelaB2 = #"Arquivo de Amostra" meta [IsParameterQuery=true, BinaryIdentifier=#"Arquivo de Amostra", Type="Binary", IsParameterQueryRequired=true];

shared #"Transformar o Arquivo de Exemplo" = let
    Fonte = Excel.Workbook(localTabelaB2, null, true),
    Monitoramento_Table = Fonte{[Item="Monitoramento",Kind="Table"]}[Data]
in
    Monitoramento_Table;

shared #"Arquivo de Amostra" = let
    Fonte = Folder.Files(localMonitIndividualT2M),
    ArquivosValidos = Table.SelectRows(Fonte, each [Attributes]?[Hidden]? <> true and not Text.StartsWith([Name], "~$") and ([Extension] = ".xlsm" or [Extension] = ".xlsx")),
    Navegação1 = ArquivosValidos{0}[Content]
in
    Navegação1
;

shared #"Transformar Arquivo" = let
    Fonte = (Parâmetro1) => let
        Fonte = Excel.Workbook(Parâmetro1, null, true),
        Monitoramento_Table = Fonte{[Item="Monitoramento",Kind="Table"]}[Data]
    in
        Monitoramento_Table
in
    Fonte;

shared Contrato = let
    ParametroTabela = Excel.CurrentWorkbook(){[Name="localTabelaB"]}[Content],
    CaminhoTabelaB = Text.From(ParametroTabela{0}[Column1]),
    ArquivoTabelaB = Binary.Buffer(File.Contents(CaminhoTabelaB)),

    Fonte = Excel.Workbook(ArquivoTabelaB, null, true),
    Contrato_Table = Fonte{[Item="Contrato", Kind="Table"]}[Data],

    #"Tipo Alterado" = Table.TransformColumnTypes(
        Contrato_Table,
        {
            {"Contrato", type text},
            {"Nome Fornecedor", type text},
            {"CNPJ", type text},
            {"Contrato SAP", Int64.Type},
            {"Instrumento Contratual Jurídico", Int64.Type},
            {"Data Início", type date},
            {"Data Término", type date},
            {"Região", type text},
            {"Endereço", type text},
            {"Município", type text},
            {"UF", type text},
            {"Filial/Matriz", type text}
        }
    )
in
    #"Tipo Alterado";

shared Grupo = let
    Fonte = Excel.Workbook(File.Contents(localTabelaB), null, true),
    Grupo_Table = Fonte{[Item="Grupo",Kind="Table"]}[Data],
    #"Tipo Alterado" = Table.TransformColumnTypes(Grupo_Table,{{"Município", type text}, {"UF", type text}, {"RMT", type text}, {"Região", type text}, {"Grupo", type text}, {"Contrato", type text}}),
    #"Colunas Reordenadas" = Table.ReorderColumns(#"Tipo Alterado",{"Contrato", "Município", "UF", "RMT", "Região", "Grupo"}),
    #"Personalização Adicionada" = Table.AddColumn(#"Colunas Reordenadas", "codGrupo", each [Município]&"\"&[UF]),
    #"Colunas Reordenadas1" = Table.ReorderColumns(#"Personalização Adicionada",{"codGrupo", "Contrato", "Município", "UF", "RMT", "Região", "Grupo"})
in
    #"Colunas Reordenadas1";

shared Localidade = let
    Fonte = Excel.Workbook(File.Contents(localTabelaB), null, true),
    Localidade_Table = Fonte{[Item="Localidade",Kind="Table"]}[Data],
    #"Tipo Alterado" = Table.TransformColumnTypes(Localidade_Table,{{"Localidade", type text}, {"Centro", type text}}),
    #"Texto em Maiúscula" = Table.TransformColumns(#"Tipo Alterado",{{"Localidade", Text.Upper, type text}, {"Centro", Text.Upper, type text}}),
    #"Texto Aparado" = Table.TransformColumns(#"Texto em Maiúscula",{{"Localidade", Text.Trim, type text}, {"Centro", Text.Trim, type text}}),
    #"Texto Limpo" = Table.TransformColumns(#"Texto Aparado",{{"Localidade", Text.Clean, type text}, {"Centro", Text.Clean, type text}})
in
    #"Texto Limpo";

shared Origem_Destino = let
    Fonte = Excel.Workbook(File.Contents(localTabelaB), null, true),
    Origem_Destino_Table = Fonte{[Item="Origem_Destino",Kind="Table"]}[Data],
    #"Tipo Alterado" = Table.TransformColumnTypes(Origem_Destino_Table,{{"DE_LocalCentro", type text}, {"DE_RegMetrop", type text}, {"DE_Tipo", type text}, {"DE_Logradouro", type text}, {"DE_CEP", type text}, {"DE_Bairro", type text}, {"DE_Cidade", type text}, {"DE_UF", type text}, {"PARA_LocalCentro", type text}, {"PARA_RegMetrop", type text}, {"PARA_Tipo", type text}, {"PARA_Logradouro", type text}, {"PARA_CEP", type text}, {"PARA_Bairro", type text}, {"PARA_Cidade", type text}, {"PARA_UF", type text}, {"KM", type number}}),
    #"Personalização Adicionada" = Table.AddColumn(#"Tipo Alterado", "codOrigenDestino", each [DE_LocalCentro]&
[DE_Cidade]&
[PARA_LocalCentro]&
[PARA_Cidade])
in
    #"Personalização Adicionada";

shared Estrutura = let
    Fonte = Excel.Workbook(File.Contents(localTabelaB), null, true),
    Estrutura_Table = Fonte{[Item="Estrutura",Kind="Table"]}[Data],
    #"Tipo Alterado" = Table.TransformColumnTypes(Estrutura_Table,{{"Lotação", type text}, {"Diretoria Executiva", type text}, {"Gerência Executiva", type text}, {"Ativo", type text}})
in
    #"Tipo Alterado";

shared PPU = let
    Fonte = Excel.Workbook(File.Contents(localTabelaB), null, true),
    PPU_Table = Fonte{[Item="PPU",Kind="Table"]}[Data],
    #"Tipo Alterado" = Table.TransformColumnTypes(PPU_Table,{{"Contrato", type text}, {"Item PPU", type text}, {"Linha de Serviço PPU", type text}, {"Descrição da Linha de Serviço", type text}, {"Unidade da Linha de Serviço", type text}, {"P. Unitário (R$)", type number}})
in
    #"Tipo Alterado";

shared #"Localidade Com Endereço" = let
    Fonte = Excel.Workbook(File.Contents(localTabelaB), null, true),
    Localidade_Table = Fonte{[Item="Localidade",Kind="Table"]}[Data],
    #"Tipo Alterado" = Table.TransformColumnTypes(Localidade_Table,{{"Localidade", type text}, {"Centro", type text}}),
    #"Texto em Maiúscula" = Table.TransformColumns(#"Tipo Alterado",{{"Localidade", Text.Upper, type text}, {"Centro", Text.Upper, type text}}),
    #"Texto Aparado" = Table.TransformColumns(#"Texto em Maiúscula",{{"Localidade", Text.Trim, type text}, {"Centro", Text.Trim, type text}}),
    #"Texto Limpo" = Table.TransformColumns(#"Texto Aparado",{{"Localidade", Text.Clean, type text}, {"Centro", Text.Clean, type text}}),
    #"Consultas Mescladas" = Table.NestedJoin(#"Texto Limpo", {"Centro"}, Centro, {"Centro"}, "Centro.1", JoinKind.LeftOuter),
    #"Centro.1 Expandido" = Table.ExpandTableColumn(#"Consultas Mescladas", "Centro.1", {"Código Centro", "CNPJ", "Endereço", "CEP", "Bairro", "Município", "UF"}, {"Código Centro", "CNPJ", "Endereço", "CEP", "Bairro", "Município", "UF"}),
    #"Linhas Filtradas" = Table.SelectRows(#"Centro.1 Expandido", each ([Endereço] <> null))
in
    #"Linhas Filtradas";

shared Galpão = let
    Fonte = Excel.Workbook(File.Contents(localTabelaB), null, true),
    Origem_Destino_Table = Fonte{[Item="Origem_Destino",Kind="Table"]}[Data],
    #"Tipo Alterado" = Table.TransformColumnTypes(Origem_Destino_Table,{{"DE_LocalCentro", type text}, {"DE_RegMetrop", type text}, {"DE_Tipo", type text}, {"DE_Logradouro", type text}, {"DE_CEP", type text}, {"DE_Bairro", type text}, {"DE_Cidade", type text}, {"DE_UF", type text}, {"PARA_LocalCentro", type text}, {"PARA_RegMetrop", type text}, {"PARA_Tipo", type text}, {"PARA_Logradouro", type text}, {"PARA_CEP", type text}, {"PARA_Bairro", type text}, {"PARA_Cidade", type text}, {"PARA_UF", type text}, {"KM", type number}}),
    #"Colunas Removidas" = Table.RemoveColumns(#"Tipo Alterado",{"DE_LocalCentro", "DE_RegMetrop", "DE_Tipo", "DE_Logradouro", "DE_CEP", "DE_Bairro", "DE_Cidade", "DE_UF", "PARA_RegMetrop", "KM"}),
    #"Duplicatas Removidas" = Table.Distinct(#"Colunas Removidas", {"PARA_LocalCentro"}),
    #"Colunas Removidas1" = Table.RemoveColumns(#"Duplicatas Removidas",{"PARA_Tipo"}),
    #"Colunas Renomeadas" = Table.RenameColumns(#"Colunas Removidas1",{{"PARA_LocalCentro", "Galpãp"}, {"PARA_Logradouro", "Galpãp_Logradouro"}, {"PARA_CEP", "Galpãp_CEP"}, {"PARA_Bairro", "Galpãp_Bairro"}, {"PARA_Cidade", "Galpãp_Cidade"}, {"PARA_UF", "Galpãp_UF"}})
in
    #"Colunas Renomeadas";

shared FDM_IAPFARQ = let
    Fonte = Excel.Workbook(File.Contents(localTabelaB), null, true),
    FDM_IAPFARQ_Table = Fonte{[Item="FDM_IAPFARQ",Kind="Table"]}[Data],
    #"Tipo Alterado" = Table.TransformColumnTypes(FDM_IAPFARQ_Table,{{"Medição", type text}, {"Contrato", type text}, {"Item contrato", Int64.Type}, {"IAPFARQ (%)", type number}, {"FDM", type number}})
in
    #"Tipo Alterado";

shared Centro = let
    Fonte = Excel.Workbook(File.Contents(localTabelaB), null, true),
    
    Centro_Table = Fonte{[Item="Centro", Kind="Table"]}[Data],

    // Tipos e limpeza da coluna Centro combinados em um único passo
    #"Tipos e Transformações" = Table.TransformColumnTypes(
        Centro_Table,
        {
            {"Código Centro", Int64.Type},
            {"CNPJ",          type text},
            {"Endereço",      type text},
            {"CEP",           type text},
            {"Bairro",        type text},
            {"Município",     type text},
            {"UF",            type text}
        }
    ),

    // Upper + Trim + Clean aplicados juntos na coluna Centro
    #"Centro Padronizado" = Table.TransformColumns(
        #"Tipos e Transformações",
        {{"Centro", each Text.Clean(Text.Trim(Text.Upper(_))), type text}}
    ),

    #"Linhas Classificadas" = Table.Sort(
        #"Centro Padronizado",
        {{"Centro", Order.Ascending}}
    )

in
    #"Linhas Classificadas";

shared ValidaGalpao = let
    Fonte = Excel.CurrentWorkbook(){[Name="ValidaGalpao"]}[Content],
    #"Tipo Alterado" = Table.TransformColumnTypes(Fonte,{{"Validação Descrição da atividade para GALPAO", type text}})
in
    #"Tipo Alterado";

shared validUFGalpao = let
    Fonte = Excel.CurrentWorkbook(){[Name="validUFGalpao"]}[Content],
    #"Tipo Alterado" = Table.TransformColumnTypes(Fonte,{{"UF", type text}, {"Contrato", type text}})
in
    #"Tipo Alterado";

shared valEnderecoGalpao = let
    Fonte = Excel.CurrentWorkbook(){[Name="valEnderecoGalpao"]}[Content],
    #"Tipo Alterado" = Table.TransformColumnTypes(Fonte,{{"Galpãp", type text}, {"Galpãp_Logradouro", type text}, {"Galpãp_CEP", type text}, {"Galpãp_Bairro", type text}, {"Galpãp_Cidade", type text}, {"Galpãp_UF", type text}})
in
    #"Tipo Alterado";

shared #"Localidade Sem endereço" = let
    Fonte = Excel.Workbook(File.Contents(localTabelaB), null, true),
    Localidade_Table = Fonte{[Item="Localidade",Kind="Table"]}[Data],
    #"Tipo Alterado" = Table.TransformColumnTypes(Localidade_Table,{{"Localidade", type text}, {"Centro", type text}}),
    #"Consultas Mescladas" = Table.NestedJoin(#"Tipo Alterado", {"Centro"}, Centro, {"Centro"}, "Centro.1", JoinKind.LeftOuter),
    #"Centro.1 Expandido" = Table.ExpandTableColumn(#"Consultas Mescladas", "Centro.1", {"Código Centro", "CNPJ", "Endereço", "CEP", "Bairro", "Município", "UF"}, {"Código Centro", "CNPJ", "Endereço", "CEP", "Bairro", "Município", "UF"}),
    #"Linhas Filtradas" = Table.SelectRows(#"Centro.1 Expandido", each ([Endereço] = null))
in
    #"Linhas Filtradas";

shared #"Qtd Registros por Doc Monitoramento" = let
    Fonte = Folder.Files(localMonitIndividualT2M),
    #"Arquivos Ocultos Filtrados1" = Table.SelectRows(Fonte, each [Attributes]?[Hidden]? <> true and not Text.StartsWith([Name], "~$") and ([Extension] = ".xlsm" or [Extension] = ".xlsx")),
    #"Invocar Função Personalizada1" = Table.AddColumn(#"Arquivos Ocultos Filtrados1", "Transformar Arquivo", each #"Transformar Arquivo"([Content])),
    #"Colunas Renomeadas1" = Table.RenameColumns(#"Invocar Função Personalizada1", {"Name", "Nome da Origem"}),
    #"Outras Colunas Removidas1" = Table.SelectColumns(#"Colunas Renomeadas1", {"Nome da Origem", "Transformar Arquivo"}),
    #"Coluna de Tabela Expandida1" = Table.ExpandTableColumn(#"Outras Colunas Removidas1", "Transformar Arquivo", Table.ColumnNames(#"Transformar Arquivo"(#"Arquivo de Amostra"))),
    #"Tipo Alterado" = Table.TransformColumnTypes(#"Coluna de Tabela Expandida1",{{"Nome da Origem", type text}, {"Atendente", type text}, {"Período#(lf)faturamento", type date}, {"Situação", type text}, {"Contrato", type text}, {"UF", type text}, {"Descrição da atividade", type text}, {"Item", type text}, {"Aplicação", type text}, {"Código da#(lf)solicitação", type text}, {"Chave#(lf)solicitante", type text}, {"Nome Solicitante", type text}, {"D. solicitação", type datetime}, {"Prazo#(lf)aplicação", type datetime}, {"Gerência solicitante", type text}, {"Qtd.#(lf)Solicitada", Int64.Type}, {"Código OS", type any}, {"D. abertura", type datetime}, {"D. fechamento", type datetime}, {"Prazo combinado", type datetime}, {"Qtd.#(lf)Atendida", Int64.Type}, {"Obs. Sigla", type text}, {"Localidade", type text}, {"Comentários", type text}, {"Observações", type text}, {"OS#(lf)disponibilizada?", type text}, {"LOG", type any}, {"Caixa#(lf)(20kg)", Int64.Type}, {"Caixa#(lf)(Mídia)", Int64.Type}, {"Caixa Tubo#(lf)(perfil de poço)", type any}, {"Caixa Tubo#(lf)(Engenharia)", Int64.Type}, {"Lacre", Int64.Type}, {"Etiqueta (20kg)", Int64.Type}, {"Etiqueta (Mídia)", Int64.Type}, {"Etiqueta Tubo#(lf)(perfil de poço)", type any}, {"Etiqueta Tubo#(lf)(Engenharia)", Int64.Type}, {"Lançamento de etiqueta/Registro OS", type text}, {"Observação", type text}}),
    #"Linhas Filtradas" = Table.SelectRows(#"Tipo Alterado", each ([#"Período#(lf)faturamento"] <> null) and ([Situação] = "AT" or [Situação] = "CO")),
    #"Linhas Agrupadas" = Table.Group(#"Linhas Filtradas", {"Nome da Origem"}, {{"Contagem", each Table.RowCount(_), Int64.Type}})
in
    #"Linhas Agrupadas"
;

shared PeriodoMes = let
    Fonte = Excel.CurrentWorkbook(){[Name="PeriodoMes"]}[Content],
    Column1 = Fonte{0}[Column1]
in
    Column1;

shared PeriodoAno = let
    Fonte = Excel.CurrentWorkbook(){[Name="PeriodoAno"]}[Content],
    Column1 = Fonte{0}[Column1]
in
    Column1;

shared fnGerarPeriodo = let
    fn = (iPerMes as text, iPerAno as number) as text =>
    let
        // Tabela de conversão: nome do mês → número
        meses = [
            Janeiro   = 1,  Fevereiro = 2,  Março    = 3,
            Abril     = 4,  Maio      = 5,  Junho    = 6,
            Julho     = 7,  Agosto    = 8,  Setembro = 9,
            Outubro   = 10, Novembro  = 11, Dezembro = 12
        ],

        // Converte o nome para número
        iPerMesNum = Record.Field(meses, iPerMes),

        // Lógica original
        iMesAnt      = iPerMesNum - 1,
        iMesAntFinal = if iMesAnt = 0 then 12 else iMesAnt,
        iAnoAnt      = if iMesAnt = 0 then iPerAno - 1 else iPerAno,
        sMesAntFmt   = Text.PadStart(Text.From(iMesAntFinal), 2, "0"),
        sMesFmt      = Text.PadStart(Text.From(iPerMesNum), 2, "0"),
        sPerAnt      = "26/" & sMesAntFmt & "/" & Text.From(iAnoAnt),
        sPerAtual    = "25/" & sMesFmt    & "/" & Text.From(iPerAno),
        resultado    = sPerAnt & " à " & sPerAtual
    in
        resultado
in
    fn
;

shared localTabelaB = let
    Fonte = Excel.CurrentWorkbook(){[Name="localTabelaB"]}[Content],
    Column1 = Fonte{0}[Column1]
in
    Column1;

shared localMonitIndividualT2M = let
    Fonte = Excel.CurrentWorkbook(){[Name="localMonitIndividualT2M"]}[Content],
    Column1 = Fonte{0}[Column1]
in
    Column1;

shared localLote1 = let
    Fonte = Excel.CurrentWorkbook(){[Name="localLote1"]}[Content],
    Column1 = Fonte{0}[Column1]
in
    Column1;

shared localES = let
    Fonte = Excel.CurrentWorkbook(){[Name="localES"]}[Content],
    Column1 = Fonte{0}[Column1]
in
    Column1;

shared localRJ = let
    Fonte = Excel.CurrentWorkbook(){[Name="localRJ"]}[Content],
    Column1 = Fonte{0}[Column1]
in
    Column1;

shared localSP = let
    Fonte = Excel.CurrentWorkbook(){[Name="localSP"]}[Content],
    Column1 = Fonte{0}[Column1]
in
    Column1;

shared localLote2 = let
    Fonte = Excel.CurrentWorkbook(){[Name="localLote2"]}[Content],
    Column1 = Fonte{0}[Column1]
in
    Column1;

shared localBA = let
    Fonte = Excel.CurrentWorkbook(){[Name="localBA"]}[Content],
    Column1 = Fonte{0}[Column1]
in
    Column1;

shared valKmAdicional = let
    Fonte = Excel.CurrentWorkbook(){[Name="valKmAdicional"]}[Content],
    #"Tipo Alterado" = Table.TransformColumnTypes(Fonte,{{"Km a sere usados no calclo de KM adicional", Int64.Type}}),
    #"Km a sere usados no calclo de KM adicional" = #"Tipo Alterado"{0}[Km a sere usados no calclo de KM adicional]
in
    #"Km a sere usados no calclo de KM adicional";

shared tabRefPrazoFDM = let
    Fonte = Excel.CurrentWorkbook(){[Name="tabRefPrazoFDM"]}[Content],
    #"Tipo Alterado" = Table.TransformColumnTypes(Fonte,{{"Serviço", type text}, {"Nº contrato", Int64.Type}, {"Descrição da atividade", type text}, {"Atividade do painel", type text}, {"Ref. FDM", Int64.Type}, {"Ref. QExec", type text}, {"Ref. prazo", type text}, {"Obs", type text}})
in
    #"Tipo Alterado";

shared tabelaAtividadeItemPPU = let
    Fonte = Excel.CurrentWorkbook(){[Name="tabelaAtividadeItemPPU"]}[Content],
    #"Tipo Alterado" = Table.TransformColumnTypes(Fonte,{{"Item PPU", type text}, {"Atividade do painel", type text}, {"Item", type text}, {"Unidade", type text}, {"Linha de serviço PPU", type text}}),
    #"Personalização Adicionada" = Table.AddColumn(#"Tipo Alterado", "CodLinhaPPU", each [Atividade do painel]&[Item])
in
    #"Personalização Adicionada";

shared tabela_FC_arm = let
    Fonte = Excel.CurrentWorkbook(){[Name="tabela_FC_arm"]}[Content],
    #"Tipo Alterado" = Table.TransformColumnTypes(Fonte,{{"Atividade do painel", type text}, {"Tab. Contrato", type text}, {"Descrição da tabela", type text}, {"Item", type text}, {"Diâmetro", Int64.Type}, {"A", type number}, {"L", type number}, {"P", type number}, {"FC", type number}})
in
    #"Tipo Alterado";

shared tabela_FC_migracao = let
    Fonte = Excel.CurrentWorkbook(){[Name="tabela_FC_migracao"]}[Content],
    #"Tipo Alterado" = Table.TransformColumnTypes(Fonte,{{"Atividade do painel", type text}, {"Tab. Contrato", type text}, {"Descrição da tabela", type text}, {"Item", type text}, {"Unidade", type text}, {"FC", type number}})
in
    #"Tipo Alterado";

shared tabela_FC_digitalizacao = let
    Fonte = Excel.CurrentWorkbook(){[Name="tabela_FC_digitalizacao"]}[Content],
    #"Tipo Alterado" = Table.TransformColumnTypes(Fonte,{{"Atividade do painel", type text}, {"Tab. Contrato", type text}, {"Descrição da tabela", type text}, {"Item", type text}, {"Unidade", type text}, {"FC", type number}})
in
    #"Tipo Alterado";

shared tabelaPrazoContrato = let
    Fonte = Excel.CurrentWorkbook(){[Name="tabelaPrazoContrato"]}[Content],
    #"Tipo Alterado" = Table.TransformColumnTypes(Fonte,{{"Contrato", type text}, {"Atividade do painel", type text}, {"Tab. Contrato", type text}, {"Grupo de município", type text}, {"Classificação", type text}, {"Prazo", type number}}),
    #"Personalização Adicionada" = Table.AddColumn(#"Tipo Alterado", "codPrazoContrato", each [Contrato]
&"\"&
[Atividade do painel] 
&"\"&
[Grupo de município]),
    #"Colunas Reordenadas" = Table.ReorderColumns(#"Personalização Adicionada",{"codPrazoContrato", "Contrato", "Atividade do painel", "Tab. Contrato", "Grupo de município", "Classificação", "Prazo"})
in
    #"Colunas Reordenadas";

shared fnCalcularDataSLA = let
    fnCalcularDataSLA = (dataInicial as date, horaInicial as time, tempoResposta as duration, optional municipio as text, optional uf as text, optional tbFeriados as table) as datetime =>
        let
            HoraEntrada = #time(8, 0, 0),
            HoraSaida = #time(17, 0, 0),
            HoraIniAlmoco = #time(12, 0, 0),
            HoraFimAlmoco = #time(13, 0, 0),

            TemExpediente = (data as date) as logical =>
                let
                    dow = Date.DayOfWeek(data, Day.Monday),
                    ehDiaUtil = dow <= 4,
                    ehFeriado = if tbFeriados = null then false else
                        let
                            munStr = if municipio = null then "" else Text.Upper(municipio),
                            ufStr = if uf = null then "" else Text.Upper(uf),
                            feriadosData = Table.SelectRows(tbFeriados, each [Data] = data),
                            matchFeriado = Table.RowCount(
                                Table.SelectRows(feriadosData, each 
                                    ([UF] = "Todos" or [UF] = null or Text.Upper([UF]) = ufStr) and
                                    ([Município] = "Todos" or [Município] = null or Text.Upper([Município]) = munStr)
                                )
                            ) > 0
                        in
                            matchFeriado
                in
                    ehDiaUtil and not ehFeriado,

            ProcessarDia = (estado as record) as record =>
                let
                    data0 = estado[Data],
                    hora0 = estado[Hora],
                    restante0 = estado[Restante],
                    Resultado =
                        if not TemExpediente(data0) then
                            [ Data = Date.AddDays(data0, 1), Hora = #time(0,0,0), Restante = restante0 ]
                        else
                            let
                                HoraAjustada =
                                    if hora0 > HoraSaida then HoraSaida
                                    else if hora0 < HoraEntrada then HoraEntrada
                                    else if hora0 > HoraIniAlmoco and hora0 < HoraFimAlmoco then HoraFimAlmoco
                                    else hora0,
                                DisponivelManha = if HoraAjustada < HoraIniAlmoco then HoraIniAlmoco - HoraAjustada else #duration(0,0,0,0),
                                CabeNaManha = HoraAjustada < HoraIniAlmoco and restante0 <= DisponivelManha,
                                HoraInicioTarde = if HoraAjustada < HoraIniAlmoco then HoraFimAlmoco else HoraAjustada,
                                RestanteNaTarde = if HoraAjustada < HoraIniAlmoco then restante0 - DisponivelManha else restante0,
                                DisponivelTarde = HoraSaida - HoraInicioTarde,
                                CabeNaTarde = RestanteNaTarde <= DisponivelTarde,
                                ResultadoDia =
                                    if CabeNaManha then
                                        [ Data = data0, Hora = HoraAjustada + restante0, Restante = #duration(0,0,0,0) ]
                                    else if CabeNaTarde then
                                        [ Data = data0, Hora = HoraInicioTarde + RestanteNaTarde, Restante = #duration(0,0,0,0) ]
                                    else
                                        [ Data = Date.AddDays(data0, 1), Hora = #time(0,0,0), Restante = RestanteNaTarde - DisponivelTarde ]
                            in
                                ResultadoDia
                in
                    Resultado,

            EstadoInicial = [ Data = dataInicial, Hora = horaInicial, Restante = tempoResposta ],

            Historico = List.Generate(
                () => EstadoInicial,
                (e) => e[Restante] > #duration(0,0,0,0),
                (e) => ProcessarDia(e)
            ),

            UltimoPendente = if List.Count(Historico) = 0 then EstadoInicial else List.Last(Historico),
            EstadoFinal = if tempoResposta <= #duration(0,0,0,0) then EstadoInicial else ProcessarDia(UltimoPendente)
        in
            // ? CORREÇÃO: time - time = duration (conversão válida)
            DateTime.From(EstadoFinal[Data]) + (EstadoFinal[Hora] - #time(0,0,0))
in
    fnCalcularDataSLA
;

shared fnCalcularPrazoSLA = let
    fnCalcularPrazoSLA = (dPrazo as number, dataAbertura as date, horaAbertura as time, optional municipio as text, optional uf as text, optional tbFeriados as table) as nullable datetime =>
        if dPrazo = 0 then null
        else if dPrazo = 0.5 then
            let
                Base = fnCalcularDataSLA(dataAbertura, horaAbertura, #duration(0,4,0,0), municipio, uf, tbFeriados),
                DataBase = Date.From(Base),
                Resultado =
                    if horaAbertura >= #time(8,0,0) and horaAbertura <= #time(11,0,0) then
                        DateTime.From(DataBase) + #duration(0,17,0,0)
                    else if horaAbertura > #time(11,0,0) and horaAbertura <= #time(16,0,0) then
                        fnCalcularDataSLA(DataBase, #time(12,0,0), #duration(0,8,0,0), municipio, uf, tbFeriados)
                    else
                        DateTime.From(DataBase) + #duration(0,17,0,0)
            in
                Resultado
        else if dPrazo = 1 then fnCalcularDataSLA(dataAbertura, horaAbertura, #duration(0,8,0,0), municipio, uf, tbFeriados)
        else if dPrazo = 2 then fnCalcularDataSLA(dataAbertura, horaAbertura, #duration(0,16,0,0), municipio, uf, tbFeriados)
        else if dPrazo = 3 then fnCalcularDataSLA(dataAbertura, horaAbertura, #duration(1,0,0,0), municipio, uf, tbFeriados)
        else if dPrazo = 4 then fnCalcularDataSLA(dataAbertura, horaAbertura, #duration(1,8,0,0), municipio, uf, tbFeriados)
        else if dPrazo = 5 then fnCalcularDataSLA(dataAbertura, horaAbertura, #duration(1,16,0,0), municipio, uf, tbFeriados)
        else null
in
    fnCalcularPrazoSLA
;

shared fnClassificarFDM = let
    // -------------------------------------------------------------------
    // fnClassificarFDM
    // -------------------------------------------------------------------
    // Equivalente M dos procedimentos VBA analiseDeDados()/
    // preenchimentoPrazoFDM() (modulo Planilha7) - classifica cada
    // linha do painel em uma categoria de FDM (Normal/Expresso/Isento/
    // FDM definido) e um Status de Prazo (NP/FP), com base no SLA
    // calculado e na referencia de prazo da atividade.
    //
    // O QUE FAZ:
    //   - calculoFDM = "S" -> linha ja e duplicata de FDM (mesma
    //     atividade/solicitacao/OS ja vista antes) -> "FDM definido".
    //   - refFDM = null/0 -> atividade fora do escopo de medicao de
    //     FDM -> "Isento".
    //   - refPrazo = "Tabela" -> valida campos obrigatorios (contrato,
    //     datas, municipio, UF, grupo); calcula o SLA via
    //     fnCalcularPrazoSLA e classifica em Normal/Expresso e NP/FP,
    //     com regras especificas para IPF/ABF/RFT.
    //   - refPrazo = "Fiscalizacao" -> compara prazo combinado x data
    //     de fechamento; se prazo combinado estiver vazio, assume
    //     Normal/NP (ajuste sem log de pendencia).
    //   - Qualquer outro valor de refPrazo -> log de inconsistencia.
    //
    // AJUSTES APLICADOS:
    //   - refPrazo = "Fiscalizacao" e prazoCombinado = null resulta em
    //     Classificacao = "Normal", StatusPrazo = "NP", Log = "".
    //   - Mensagens de observacao/log padronizadas sem acentos para
    //     evitar problemas de incompatibilidade de codificacao.
    //   - Suporte opcional a tabela de feriados para o calculo de SLA.
    //
    // PARAMETROS:
    //   calculoFDM, refFDM, refPrazo, descricaoAtividade, contrato,
    //   dataAbertura, horaAbertura, dataFechamento, prazoCombinado,
    //   municipio, uf, grupo, dPrazo, aplicacao, optional tbFeriados
    //
    // RETORNO:
    //   record [ Classificacao, StatusPrazo, Log ]
    //
    // DEPENDENCIA:
    //   fnCalcularPrazoSLA (usada dentro do ramo refPrazo="Tabela")
    // -------------------------------------------------------------------
    fnClassificarFDM = (
        calculoFDM as text,
        refFDM as any,
        refPrazo as nullable text,
        descricaoAtividade as text,
        contrato as nullable text,
        dataAbertura as nullable date,
        horaAbertura as nullable time,
        dataFechamento as nullable datetime,
        prazoCombinado as nullable datetime,
        municipio as nullable text,
        uf as nullable text,
        grupo as nullable text,
        dPrazo as nullable number,
        aplicacao as nullable text,
        optional tbFeriados as table
    ) as record =>
    let
        Resultado =
            if calculoFDM = "S" then
                [ Classificacao = "FDM definido", StatusPrazo = null, Log = "" ]

            else if refFDM = null or refFDM = 0 then
                [ Classificacao = "Isento", StatusPrazo = null, Log = "" ]

            else if refPrazo = "Tabela" then
                let
                    ValidacaoCampoVazio =
                        if contrato = null or contrato = "" then "Contrato vazio; "
                        else if dataAbertura = null then "D. abertura vazio; "
                        else if dataFechamento = null then "D. fechamento vazio; "
                        else if municipio = null or municipio = "" then "Municipio vazio; "
                        else if uf = null or uf = "" then "UF vazio; "
                        else if grupo = null or grupo = "Nao encontrado" then "Municipio sem grupo valido para a UF; "
                        else "",

                    ResultadoTabela =
                        if ValidacaoCampoVazio <> "" then
                            [ Classificacao = null, StatusPrazo = null, Log = ValidacaoCampoVazio ]
                        else if dPrazo = 0 or dPrazo = null then
                            [ Classificacao = null, StatusPrazo = null, Log = "FDM nao permitido para este grupo/atividade; " ]
                        else
                            let
                                dtReSLA = fnCalcularPrazoSLA(dPrazo, dataAbertura, horaAbertura, municipio, uf, tbFeriados),
                                DentroDoPrazo = dtReSLA >= dataFechamento,
                                EhExpresso = Text.Contains(Text.Upper(descricaoAtividade), "EXPRESSO"),
                                ResultadoFinal =
                                    if DentroDoPrazo then
                                        if EhExpresso then
                                            if aplicacao = "IPF" then
                                                [ Classificacao = "Normal", StatusPrazo = "FP", Log = "" ]
                                            else
                                                [ Classificacao = "Expresso", StatusPrazo = "NP", Log = "" ]
                                        else
                                            [ Classificacao = "Normal", StatusPrazo = "NP", Log = "" ]
                                    else
                                        if aplicacao = "ABF" or aplicacao = "RFT" then
                                            [ Classificacao = if EhExpresso then "Expresso" else "Normal", StatusPrazo = "NP", Log = "" ]
                                        else
                                            [ Classificacao = "Normal", StatusPrazo = "FP", Log = "" ]
                            in
                                ResultadoFinal
                in
                    ResultadoTabela

            else if refPrazo = "Fiscalizacao" or refPrazo = "Fiscalização" or (refPrazo <> null and Text.StartsWith(Text.Upper(refPrazo), "FISCALIZA")) then
                if dataFechamento = null then
                    [ Classificacao = null, StatusPrazo = null, Log = "D. fechamento vazio; " ]
                else if prazoCombinado = null then
                    // Ajuste aplicado nesta conversa: assume Normal/NP,
                    // sem registrar pendencia de log
                    [ Classificacao = "Normal", StatusPrazo = "NP", Log = "" ]
                else if prazoCombinado >= dataFechamento then
                    [ Classificacao = "Normal", StatusPrazo = "NP", Log = "" ]
                else
                    [ Classificacao = "Normal", StatusPrazo = "FP", Log = "" ]

            else
                [ Classificacao = null, StatusPrazo = null, Log = "Ref. prazo nao reconhecida; " ]
    in
        Resultado
in
    fnClassificarFDM

;

shared Consulta1 = let
    Fonte = tabelaPrazoContrato,
    Amostra = Table.FirstN(Fonte, 10)
in
    Amostra
;

shared tabela_FC = let
    Fonte = Excel.CurrentWorkbook(){[Name="tabela_FC_arm"]}[Content],
    #"Tipo Alterado" = Table.TransformColumnTypes(Fonte,{{"Atividade do painel", type text}, {"Tab. Contrato", type text}, {"Descrição da tabela", type text}, {"Item", type text}, {"Diâmetro", Int64.Type}, {"A", type number}, {"L", type number}, {"P", type number}, {"FC", type number}}),
    #"Consulta Acrescentada" = Table.Combine({#"Tipo Alterado", tabela_FC_migracao, tabela_FC_digitalizacao}),
    #"Personalização Adicionada" = Table.AddColumn(#"Consulta Acrescentada", "cod_Tabela_FC", each [Atividade do painel]
&
[Item])
in
    #"Personalização Adicionada";

shared ConsolidadoES = let
    Fonte = Excel.Workbook(File.Contents(localLote1&localES), null, true),
    ConsolidadoES_Table = Fonte{[Item="ConsolidadoES",Kind="Table"]}[Data],
    #"Linhas Filtradas" = Table.SelectRows(ConsolidadoES_Table, each ([MEDIDAS CONTRATO] <> null)),
    #"Personalização Adicionada" = Table.AddColumn(#"Linhas Filtradas", "Contrato-Reg.", each "IRON-LT1-ES"),
    #"Colunas Reordenadas" = Table.ReorderColumns(#"Personalização Adicionada",{"Contrato-Reg.", "MEDIDAS CONTRATO", "Padrão", "Ø", "A", "L", "P", "cm³ = V1", "cm³ = V2", "FC = V1/V2", "QUAm-1", "QBm", "QDm", "QNm", "QUAm", "UA", "$ Unitário", "$ Total", "Tabela", "Contrato"})
in
    #"Colunas Reordenadas";

shared ConsolidadoRJ = let
    Fonte = Excel.Workbook(File.Contents(localLote1&localRJ), null, true),
    ConsolidadoRJ_Table = Fonte{[Item="ConsolidadoRJ",Kind="Table"]}[Data],
    #"Tipo Alterado" = Table.TransformColumnTypes(ConsolidadoRJ_Table,{{"MEDIDAS CONTRATO", type text}, {"Padrão", type text}, {"Ø", Int64.Type}, {"A", type number}, {"L", type number}, {"P", type number}, {"cm³ = V1", type number}, {"cm³ = V2", type number}, {"FC = V1/V2", type number}, {"QUAm-1", Int64.Type}, {"QBm", Int64.Type}, {"QDm", Int64.Type}, {"QNm", Int64.Type}, {"QUAm", Int64.Type}, {"UA", type number}, {"$ Unitário", type number}, {"$ Total", type number}, {"Tabela", type text}, {"Contrato", type text}}),
    #"Linhas Filtradas" = Table.SelectRows(#"Tipo Alterado", each ([MEDIDAS CONTRATO] <> null)),
    #"Personalização Adicionada" = Table.AddColumn(#"Linhas Filtradas", "Contrato-Reg.", each "IRON-LT1-RJ"),
    #"Colunas Reordenadas" = Table.ReorderColumns(#"Personalização Adicionada",{"Contrato-Reg.", "MEDIDAS CONTRATO", "Padrão", "Ø", "A", "L", "P", "cm³ = V1", "cm³ = V2", "FC = V1/V2", "QUAm-1", "QBm", "QDm", "QNm", "QUAm", "UA", "$ Unitário", "$ Total", "Tabela", "Contrato"})
in
    #"Colunas Reordenadas";

shared ConsolidadoSP = let
    Fonte = Excel.Workbook(File.Contents(localLote1&localSP), null, true),
    ConsolidadoSP_Table = Fonte{[Item="ConsolidadoSP",Kind="Table"]}[Data],
    #"Tipo Alterado" = Table.TransformColumnTypes(ConsolidadoSP_Table,{{"MEDIDAS CONTRATO", type text}, {"Padrão", type text}, {"Ø", Int64.Type}, {"A", type number}, {"L", type number}, {"P", type number}, {"cm³ = V1", type number}, {"cm³ = V2", type number}, {"FC = V1/V2", type number}, {"QUAm-1", Int64.Type}, {"QBm", Int64.Type}, {"QDm", Int64.Type}, {"QNm", Int64.Type}, {"QUAm", Int64.Type}, {"UA", type number}, {"$ Unitário", type number}, {"$ Total", type number}, {"Tabela", type text}, {"Contrato", type text}}),
    #"Linhas Filtradas" = Table.SelectRows(#"Tipo Alterado", each ([MEDIDAS CONTRATO] <> null)),
    #"Personalização Adicionada" = Table.AddColumn(#"Linhas Filtradas", "Contrato-Reg.", each "IRON-LT1-SP"),
    #"Colunas Reordenadas" = Table.ReorderColumns(#"Personalização Adicionada",{"Contrato-Reg.", "MEDIDAS CONTRATO", "Padrão", "Ø", "A", "L", "P", "cm³ = V1", "cm³ = V2", "FC = V1/V2", "QUAm-1", "QBm", "QDm", "QNm", "QUAm", "UA", "$ Unitário", "$ Total", "Tabela", "Contrato"})
in
    #"Colunas Reordenadas";

shared ConsolidadoPA = let
    Fonte = Excel.Workbook(File.Contents(localLote2&localBA), null, true),
    ConsolidadoPA_Table = Fonte{[Item="ConsolidadoPA",Kind="Table"]}[Data],
    #"Tipo Alterado" = Table.TransformColumnTypes(ConsolidadoPA_Table,{{"MEDIDAS CONTRATO", type text}, {"Padrão", type text}, {"Ø", type any}, {"A", type any}, {"L", type any}, {"P", type any}, {"cm³ = V1", type number}, {"cm³ = V2", Int64.Type}, {"FC = V1/V2", type number}, {"QUAm-1", Int64.Type}, {"QBm", type any}, {"QDm", type any}, {"QNm", Int64.Type}, {"QUAm", Int64.Type}, {"UA", type number}, {"$ Unitário", type number}, {"$ Total", type any}, {"Tabela", type text}, {"Contrato", type text}}),
    #"Linhas Filtradas" = Table.SelectRows(#"Tipo Alterado", each ([MEDIDAS CONTRATO] <> null)),
    #"Personalização Adicionada" = Table.AddColumn(#"Linhas Filtradas", "Contrato-Reg.", each "PA-LT2-BA"),
    #"Colunas Reordenadas" = Table.ReorderColumns(#"Personalização Adicionada",{"Contrato-Reg.", "MEDIDAS CONTRATO", "Padrão", "Ø", "A", "L", "P", "cm³ = V1", "cm³ = V2", "FC = V1/V2", "QUAm-1", "QBm", "QDm", "QNm", "QUAm", "UA", "$ Unitário", "$ Total", "Tabela", "Contrato"})
in
    #"Colunas Reordenadas";

shared fnCalcularIAPFARQ = let
    // -------------------------------------------------------------------
    // fnCalcularIAPFARQ
    // -------------------------------------------------------------------
    // Equivalente M da função VBA calcular_IAPFARQ (módulo fncEditalGuarda)
    // - item nº 8.4.2 do contrato de guarda externa Petrobras.
    //
    // O QUE FAZ:
    //   Calcula o índice IAPFARQ (percentual de solicitações atendidas
    //   DENTRO DO PRAZO em relação ao total de solicitações elegíveis
    //   atendidas no período), pela fórmula:
    //
    //       IAPFARQ = (TSPR / TSPE) * 100
    //
    // CASOS ESPECIAIS & REGRAS DE ROBUSTEZ:
    //   - Se TSPE = 0: retorna 100.0 (desempenho pleno por ausência de demanda).
    //   - Se TSPR = 0 e TSPE > 0: retorna 0.0 (conformidade nula real).
    //   - Clamping defensivo: garante que TSPR não exceda TSPE.
    //   - Tratamento seguro de nulos com condicionais nativas M.
    //   - Arredondamento padrão a 2 casas decimais via Number.Round.
    // -------------------------------------------------------------------
    fnCalcularIAPFARQ = (
        TSPE as nullable number,
        TSPR as nullable number,
        optional casasDecimais as nullable number
    ) as nullable number =>
        let
            vTSPE = if TSPE = null or TSPE < 0 then 0 else TSPE,
            vTSPR = if TSPR = null or TSPR < 0 then 0 else TSPR,
            precisao = if casasDecimais = null then 2 else casasDecimais,

            ResultadoCalculado =
                if vTSPE = 0 then
                    100.0
                else if vTSPR = 0 then
                    0.0
                else
                    let
                        vTSPRValido = if vTSPR > vTSPE then vTSPE else vTSPR,
                        Indice = (vTSPRValido / vTSPE) * 100.0
                    in
                        Indice,

            ResultadoFinal =
                if ResultadoCalculado = null then
                    null
                else
                    Number.Round(ResultadoCalculado, precisao)
        in
            ResultadoFinal
in
    fnCalcularIAPFARQ
;

shared fnCalcularQExec = let
    // ═══════════════════════════════════════════════════════════════════
    // fnCalcularQExec
    // ─────────────────────────────────────────────────────────────────
    // Equivalente M da função VBA calcular_QExec (módulo
    // fncEditalGuarda) — usada para medição/faturamento das atividades
    // de guarda externa (ver resultado no template de Memória de
    // Cálculo, aba DADOS, coluna AA).
    //
    // O QUE FAZ:
    //   Converte a quantidade física solicitada (qtd) na quantidade
    //   padronizada de medição (QExec), de acordo com a atividade e o
    //   item, seguindo diferentes regras de conversão por padrão de
    //   atividade:
    //
    //   PADRÃO A — Atividades tipo "Embalagem" (COLETA/ENTREGA de
    //   embalagem): se qtd <= 10 e sem agrupamento, aplica piso mínimo
    //   de 1 unidade; caso contrário, qtd * 0,1 (fator FC_EMB).
    //
    //   PADRÃO B — Atividades tipo "Item avulso" (COLETA/ENTREGA de
    //   item avulso): se qtd <= 60 e sem agrupamento, piso mínimo de 1
    //   unidade; caso contrário, qtd * 0,017 (fator FC_ITEM).
    //
    //   PADRÃO C — Passagem direta (sem conversão): QExec = qtd, para
    //   atividades como TRANSFERENCIA DE ACERVO, DESTRUICAO SEGURA DE
    //   DOCUMENTO, HIGIENIZACAO DE DOCUMENTO, entre outras.
    //
    //   PADRÃO D — Organização de documento, com fator fixo:
    //     ANALITICA → fator 1 | SIMPLES → fator 0,5
    //
    //   PADRÃO E — Migração/Digitalização: o fator de conversão (FC)
    //   é lido diretamente do parâmetro "fc" (valor já pré-calculado
    //   em uma coluna "FC" na planilha principal, via lookup/merge
    //   externo — decisão tomada nesta conversa para simplificar a
    //   função, em vez de recriar internamente as funções VBA
    //   pesquisar_FC_migracao/pesquisar_FC_digitalizacao).
    //
    //   PADRÃO F — INDEXACAO DE DOCUMENTO: regra especial que NÃO
    //   verifica agrupamento (diferente dos Padrões A/B) — replica
    //   fielmente o comportamento do VBA original, ainda pendente de
    //   confirmação se é regra intencional ou lacuna do código legado.
    //
    //   ATIVIDADES SEM CÁLCULO — "MATERIAL PARA ARQUIVAMENTO" e
    //   "BAIXA PERMANENTE" retornam NULL diretamente: ambas entram na
    //   Memória de Cálculo como registro informativo, mas sem nenhum
    //   valor de QExec associado (replica o "GoTo proximaLinha" do
    //   VBA original, que pulava todo o bloco de cálculo para essas
    //   atividades específicas).
    //
    // ✅ CORREÇÕES APLICADAS EM RELAÇÃO AO VBA ORIGINAL:
    //   1. Removidos branches ElseIf/Else redundantes presentes em 6
    //      Case do VBA original (calculavam o mesmo valor duas vezes).
    //   2. "ENTREGA DE ITEM NORMAL" agora arredonda a 3 casas
    //      decimais, igual a "ENTREGA DE ITEM EXPRESSO" (o VBA
    //      original tinha essa inconsistência de precisão entre as
    //      duas atividades análogas).
    //   3. Fallback explícito com "error" para qualquer atividade não
    //      mapeada nesta função — em vez de retornar 0 silenciosamente
    //      (bug real identificado em produção: "BAIXA PERMANENTE" e
    //      "MATERIAL PARA ARQUIVAMENTO" geravam esse erro até serem
    //      tratadas explicitamente como exceções sem cálculo).
    //   4. "DEVOLUCAO DE EMPRESTIMO" com valor de Item não reconhecido
    //      (diferente de "Embalagem"/"Item avulso") também lança erro
    //      explícito, em vez de zerar silenciosamente.
    //   5. Uso de Number.Round em vez de Format(...) — evita
    //      dependência do separador decimal do locale do sistema, e
    //      mantém o tipo numérico do início ao fim do cálculo.
    //
    // PARÂMETROS:
    //   descricaoAtividade - texto com o nome da atividade (ex.:
    //                         "COLETA DE EMBALAGEM")
    //   item                - texto com o item (ex.: "Embalagem",
    //                         "Item avulso") — usado apenas por
    //                         DEVOLUCAO DE EMPRESTIMO
    //   qtd                 - quantidade física solicitada/atendida
    //   agrupamento         - texto do código de agrupamento da linha
    //                         (vazio/null se a linha não pertence a
    //                         um agrupamento)
    //   fc                  - fator de conversão pré-calculado (coluna
    //                         "FC" da planilha principal), usado
    //                         apenas pelas atividades de
    //                         Migração/Digitalização
    //
    // RETORNO:
    //   nullable number — a quantidade padronizada de medição (QExec),
    //   ou null para atividades sem cálculo aplicável
    // ═══════════════════════════════════════════════════════════════════

    fnCalcularQExec = (
        descricaoAtividade as text,
        item as nullable text,
        qtd as nullable number,
        agrupamento as nullable text,
        fc as nullable number
    ) as nullable number =>
        let
            FC_EMB = 0.1,
            FC_ITEM = 0.017,

            SemAgrupamento = agrupamento = null or agrupamento = "",

            // ✅ CORREÇÃO DEFINITIVA: qtd nula é tratada ANTES de
            // qualquer comparação "qtd <= X" ou multiplicação, evitando
            // que null se propague para dentro de um "if" e cause o
            // erro "não conseguimos converter null em Logical/Number".
            // Decisão: qtd nula é tratada como 0 (equivalente a "nenhum
            // item atendido" naquela linha).
            qtdValidada = if qtd = null then 0 else qtd,

            AtividadesEmbalagem = {
                "COLETA DE EMBALAGEM",
                "ENTREGA DE EMBALAGEM EXPRESSO",
                "ENTREGA DE EMBALAGEM NORMAL"
            },

            AtividadesItemAvulso = {
                "COLETA DE ITEM AVULSO",
                "ENTREGA DE ITEM EXPRESSO",
                "ENTREGA DE ITEM NORMAL"
            },

            AtividadesPassagemDireta = {
                "INSERCAO DE ITEM AVULSO",
                "PESQUISA DE ITEM AVULSO EXPRESSO",
                "PESQUISA DE ITEM AVULSO NORMAL",
                "DESTRUICAO SEGURA DE DOCUMENTO",
                "TRANSFERENCIA DE ACERVO",
                "CATALOGACAO DE EMBALAGEM",
                "CATALOGACAO DE ITEM AVULSO",
                "HIGIENIZACAO DE DOCUMENTO",
                "CONVERSAO DE MIDIA",
                "GRAVACAO DE MIDIA",
                "COPIA DE MIDIA",
                "COPIA DE VIDEO"
            },

            AtividadesFCExterno = {
                "MIGRACAO DE ACERVO DOCUMENTAL",
                "DIGITALIZACAO DE MICROFILME",
                "DIGITALIZACAO DE MICROFICHA",
                "DIGITALIZACAO DE DOCUMENTO CONTRATUAL",
                "DIGITALIZACAO DE DOCUMENTO TECNICO",
                "DIGITALIZACAO DE DOCUMENTO BIBLIOGRAFICO",
                "DIGITALIZACAO DE DOCUMENTO CONTABIL E FINANCEIRO",
                "DIGITALIZACAO DE DOCUMENTO ADMINISTRATIVO"
            },

            AtividadesSemCalculo = {
                "MATERIAL PARA ARQUIVAMENTO",
                "BAIXA PERMANENTE"
            },

            Resultado =

                if List.Contains(AtividadesSemCalculo, descricaoAtividade) then
                    null

                else if descricaoAtividade = "ORGANIZACAO DE DOCUMENTO ANALITICA" then
                    Number.Round(1 * qtdValidada, 3)

                else if descricaoAtividade = "ORGANIZACAO DE DOCUMENTO SIMPLES" then
                    Number.Round(0.5 * qtdValidada, 3)

                else if List.Contains(AtividadesFCExterno, descricaoAtividade) then
                    if fc = null then
                        null   // Opção C — sem FC cadastrado, sem interromper a consulta
                    else
                        Number.Round(fc * qtdValidada, 3)

                else if descricaoAtividade = "INDEXACAO DE DOCUMENTO" then
                    if qtdValidada <= 10 then 1
                    else Number.Round(FC_EMB * qtdValidada, 3)

                else if List.Contains(AtividadesPassagemDireta, descricaoAtividade) then
                    Number.Round(qtdValidada, 3)

                else if List.Contains(AtividadesEmbalagem, descricaoAtividade) then
                    if qtdValidada <= 10 and SemAgrupamento then 1
                    else Number.Round(FC_EMB * qtdValidada, 3)

                else if List.Contains(AtividadesItemAvulso, descricaoAtividade) then
                    if qtdValidada <= 60 and SemAgrupamento then 1
                    else Number.Round(FC_ITEM * qtdValidada, 3)

                else if descricaoAtividade = "DEVOLUCAO DE EMPRESTIMO" then
                    if item = "Embalagem" then
                        if qtdValidada <= 10 and SemAgrupamento then 1
                        else Number.Round(FC_EMB * qtdValidada, 3)
                    else if item = "Item avulso" then
                        if qtdValidada <= 60 and SemAgrupamento then 1
                        else Number.Round(FC_ITEM * qtdValidada, 3)
                    else
                        error Error.Record(
                            "ItemNaoReconhecido",
                            "DEVOLUCAO DE EMPRESTIMO com Item não reconhecido: '" & Text.From(item) & "'. Esperado 'Embalagem' ou 'Item avulso'."
                        )

                else
                    error Error.Record(
                        "AtividadeNaoMapeada",
                        "Atividade '" & descricaoAtividade & "' não possui regra de cálculo de QExec definida em fnCalcularQExec."
                    )
        in
            Resultado
in
    fnCalcularQExec
;

shared fnValidarLocalizacao = let
    // ═══════════════════════════════════════════════════════════════════
    // fnValidarLocalizacao
    // Verifica se a atividade está entre as que permitem validação de
    // dados entre localidades (usado para gerar o campo Localidade do
    // galpão de guarda / cálculo de KM adicional).
    // ═══════════════════════════════════════════════════════════════════
    fnValidarLocalizacao = (sAtividade as nullable text) as logical =>
        List.Contains({
            "DEVOLUCAO DE EMPRESTIMO",
            "COLETA DE EMBALAGEM",
            "COLETA DE ITEM AVULSO",
            "ENTREGA DE EMBALAGEM EXPRESSO",
            "ENTREGA DE EMBALAGEM NORMAL",
            "ENTREGA DE ITEM EXPRESSO",
            "ENTREGA DE ITEM NORMAL"
        }, sAtividade)
in
    fnValidarLocalizacao
;

shared fnCalcularQExecAgrupado = let
    // -------------------------------------------------------------------
    // fnCalcularQExecAgrupado
    // -------------------------------------------------------------------
    // Calculo ADICIONAL/PARALELO ao QExec por linha (fnCalcularQExec).
    // Cria um indice de agrupamento a partir de (Data Abertura + Local
    // PETROBRAS + Local da empresa de guarda), soma o QExec de todas
    // as linhas com o mesmo indice, e atribui o resultado da formula
    // de frete/agrupamento (soma/10, piso 1) APENAS na linha de
    // primeira ocorrencia do indice - mesmo padrao de deduplicacao
    // usado em concatenar_dados_FDM (VBA) e CalcularQuilometragemAdicional
    // (Sistema 2, via Collection).
    //
    // O QUE FAZ:
    //   resultado = 1,          se (soma do grupo / 10) <= 1
    //   resultado = soma / 10,  caso contrario
    //   As demais linhas do mesmo grupo (nao a 1a ocorrencia) recebem
    //   null nesta coluna.
    //
    // ? CORRECAO APLICADA (durante os testes em producao):
    //   Se TODAS as linhas de um grupo tiverem QExec = null (ex.:
    //   grupo formado inteiramente por "MATERIAL PARA ARQUIVAMENTO"/
    //   "BAIXA PERMANENTE"), List.Sum de uma lista vazia retornaria
    //   null, causando o erro "Nao conseguimos converter o valor null
    //   em tipo Logical" ao comparar null <= 1. Corrigido verificando
    //   explicitamente List.IsEmpty antes de somar, e propagando null
    //   como resultado final nesse caso, em vez de tentar a divisao/
    //   comparacao.
    //
    // PARAMETROS:
    //   tabela            - tabela de entrada (ja com QExec calculado
    //                       por fnCalcularQExec)
    //   colDataFechamento - nome da coluna de Data de Fechamento
    //   colLocalPetrobras - nome da coluna de Localidade PETROBRAS
    //   colLocalGuarda    - nome da coluna de Localidade da empresa de
    //                       guarda (galpao)
    //   colLinhaServico   - nome da coluna contendo a Linha de Servico PPU
    //                       (ex.: "Linha de serviço PPU" - separa FRE-EXP de FRE-NRM)
    //   colQExec          - nome da coluna ja contendo o QExec por
    //                       linha
    //
    // RETORNO:
    //   table - a mesma tabela de entrada, com a coluna adicional
    //   "QExecAgrupado" preenchida apenas na 1a ocorrencia de cada
    //   grupo
    // -------------------------------------------------------------------
    fnCalcularQExecAgrupado = (
        tabela as table,
        colDataFechamento as text,
        colLocalPetrobras as text,
        colLocalGuarda as text,
        colLinhaServico as text,
        colQExec as text
    ) as table =>
        let
            // 1) Monta a chave de agrupamento (Data + Local PETROBRAS +
            //    Local Guarda + Linha de Servico PPU). A inclusao de
            //    colLinhaServico garante que fretes com SLAs e precos
            //    diferentes (FRE-NRM vs FRE-EXP) nao sejam agrupados
            //    indevidamente na mesma viagem (conforme A_Calc_Medicao.bas linha 492).
            ComChave = Table.AddColumn(tabela, "_ChaveAgrupamento", each
                let
                    dt = Record.Field(_, colDataFechamento),
                    locP = Record.Field(_, colLocalPetrobras),
                    locG = Record.Field(_, colLocalGuarda),
                    serv = Record.Field(_, colLinhaServico)
                in
                    (if dt = null then "" else Text.From(dt, "pt-BR"))
                    & "|" & (if locP = null then "" else Text.From(locP))
                    & "|" & (if locG = null then "" else Text.From(locG))
                    & "|" & (if serv = null then "" else Text.From(serv))
            ),
            // 2) Indice sequencial estavel - necessario para localizar
            //    a primeira ocorrencia de cada chave
            ComIndice = Table.AddIndexColumn(ComChave, "_Indice", 0, 1, Int64.Type),
            // 3) Agrupa por chave: acha o indice minimo (1a ocorrencia)
            //    e soma o QExec de todas as linhas do grupo
            Agrupado = Table.Group(ComIndice, {"_ChaveAgrupamento"}, {
                {"_IndiceMinimo", each List.Min([_Indice]), Int64.Type},
                {"_SomaQExec", each
                    let
                        valores = List.RemoveNulls(Table.Column(_, colQExec)),
                        // ? Se nao houver nenhum valor valido no
                        // grupo, retorna null em vez de deixar
                        // List.Sum({}) produzir null "por acidente"
                        soma = if List.IsEmpty(valores) then null else List.Sum(valores)
                    in
                        soma,
                    type nullable number
                }
            }),
            // 4) Traz de volta o indice minimo e a soma do grupo para
            //    cada linha original
            Mesclado = Table.NestedJoin(ComIndice, {"_ChaveAgrupamento"},
                Agrupado, {"_ChaveAgrupamento"}, "_Grupo", JoinKind.LeftOuter),
            Expandido = Table.ExpandTableColumn(Mesclado, "_Grupo",
                {"_IndiceMinimo", "_SomaQExec"}, {"_IndiceMinimo", "_SomaQExec"}),
            // 5) Aplica o calculo APENAS na linha de 1a ocorrencia
            //    (Indice = IndiceMinimo); demais linhas = null.
            //    Tambem protege contra soma nula (grupo sem QExec
            //    valido nenhum).
            ComResultado = Table.AddColumn(Expandido, "QExecAgrupado", each
                if [_Indice] = [_IndiceMinimo] then
                    if [_SomaQExec] = 0 or [_SomaQExec] = null then
                        0
                    else
                        let divisao = [_SomaQExec] / 10 in
                            if divisao <= 1 then 1 else divisao
                else
                    null
            ),
            // 6) Remove colunas auxiliares, mantendo so o resultado final
            Limpeza = Table.RemoveColumns(ComResultado,
                {"_ChaveAgrupamento", "_Indice", "_IndiceMinimo", "_SomaQExec"}
            )
        in
            Limpeza
in
    fnCalcularQExecAgrupado
;

shared DADOS = let
    Fonte = #"Painel T2M",

    // Table.SelectColumns seleciona E reordena em uma única operação —
    // substitui os 3 passos de Table.ReorderColumns + os 2 passos de
    // Table.RemoveColumns do código original. "Obs. Sigla" já é
    // excluída aqui diretamente — no original ela era renomeada para
    // "Obs." e só depois removida, desperdiçando processamento.
    ColunasSelecionadas = Table.SelectColumns(Fonte, {
        "Contrato.1",
        "Descrição da atividade", "Item", "Aplicação",
        "Código da#(lf)solicitação", "Data solicitação", "Gerência solicitante",
        "Qtd.#(lf)Solicitada", "Código OS", "D. abertura", "D. fechamento",
        "Prazo combinado", "Qtd.#(lf)Atendida",
        "Classificação FDM", "Status Prazo",
        "Centro", "Município", "UF",
        "Galpão", "Galpãp_Cidade", "Galpãp_UF",
        "KM Adicional", "FC", "QExecAgrupado","QExec",
        "Unidade", "Localidade", "Código Centro", "Linha de serviço PPU"
    }),

    // As 10 renomeações combinadas em UMA única chamada (o original
    // fazia isso em 6 chamadas separadas de RenameColumns, cada uma
    // intercalada com uma reordenação completa desnecessária)
    ColunasRenomeadas = Table.RenameColumns(ColunasSelecionadas, {
        {"Centro", "Localidade - Origem"},
        {"Município", "Município - Origem"},
        {"UF", "UF - Origem"},
        {"Galpão", "Localidade - Destino"},
        {"Galpãp_Cidade", "Município - Destino"},
        {"Galpãp_UF", "UF - Destino"},
        {"Localidade", "Centro"},
        {"Classificação FDM", "Classificação"},
        {"Status Prazo", "FDM"}
    }),

    // Colunas de preenchimento manual pelo analista (Agrupamento, FA,
    // Obs. Isenção) — usando null com tipo explícito, em vez de ""
    // (texto vazio) do original. Isso é semanticamente mais correto
    // para "ainda não preenchido" e evita erro de tipagem quando a
    // coluna FA (numérica) for usada em cálculos futuros.
    AdicionarAgrupamento = Table.AddColumn(ColunasRenomeadas, "Agrupamento", each null, type text),
    AdicionarFA = Table.AddColumn(AdicionarAgrupamento, "FA", each null, type number),
    AdicionarObsIsencao = Table.AddColumn(AdicionarFA, "Obs. Isenção", each null, type text),

    // Reordenação final ÚNICA — intercala as 3 colunas manuais nas
    // posições corretas em relação às colunas já selecionadas.
    // ✅ "Contrato" agora posicionada como PRIMEIRA coluna da tabela,
    // conforme solicitado.
    ColunasFinal = Table.ReorderColumns(AdicionarObsIsencao, {
        "Contrato.1",
        "Descrição da atividade", "Item", "Aplicação",
        "Código da#(lf)solicitação", "Data solicitação", "Gerência solicitante",
        "Qtd.#(lf)Solicitada", "Código OS", "D. abertura", "D. fechamento",
        "Prazo combinado", "Qtd.#(lf)Atendida",
        "Agrupamento", "Classificação", "FDM", "Obs. Isenção", "FA",
        "Localidade - Origem", "Município - Origem", "UF - Origem",
        "Localidade - Destino", "Município - Destino", "UF - Destino",
        "KM Adicional", "FC", "QExec",
        "Unidade", "Centro", "Código Centro", "Linha de serviço PPU"
    }),

    // ✅ Table.Buffer no resultado final — materializa a tabela em
    // memória uma única vez, evitando reprocessar toda a cadeia acima
    // (que já depende do buffer final de #"Painel T2M") caso esta
    // consulta "DADOS" seja referenciada por outras consultas (ex.:
    // SaidaMEDICAO) mais de uma vez.
    ResultadoFinalBuffer = Table.Buffer(ColunasFinal),
    #"Colunas Reordenadas" = Table.ReorderColumns(ResultadoFinalBuffer,{"Contrato.1", "Descrição da atividade", "Item", "Aplicação", "Código da#(lf)solicitação", "Data solicitação", "Gerência solicitante", "Qtd.#(lf)Solicitada", "Código OS", "D. abertura", "D. fechamento", "Prazo combinado", "Qtd.#(lf)Atendida", "Agrupamento", "Classificação", "FDM", "Obs. Isenção", "FA", "Localidade - Origem", "Município - Origem", "UF - Origem", "Localidade - Destino", "Município - Destino", "UF - Destino", "KM Adicional", "FC", "Unidade", "Centro", "Código Centro", "Linha de serviço PPU", "QExec", "QExecAgrupado"})
in
    #"Colunas Reordenadas";

shared FRETE = let
    Fonte = DADOS,
    #"Colunas Removidas" = Table.RemoveColumns(Fonte,{"Aplicação", "Código da#(lf)solicitação", "Código OS", "D. abertura", "Prazo combinado", "Qtd.#(lf)Atendida", "Agrupamento", "Classificação", "FDM", "Obs. Isenção", "FA", "FC", "QExec", "Código Centro"}),
    #"Linhas Filtradas" = Table.SelectRows(#"Colunas Removidas", each ([Linha de serviço PPU] = "FRE-EXP" or [Linha de serviço PPU] = "FRE-NRM"))
in
    #"Linhas Filtradas";

shared fnRecalcularFDM = let
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
                fdmAnt2 = fnBuscarFDMAnterior(sPerAnt2, sContrato, iItem)
            in
                if fdmAnt1 = null then
                    error Error.Record(
                        "FDMAnteriorNaoEncontrado",
                        "FDM do período anterior (" & sPerAnt1 & ") não localizado para o contrato '" & sContrato & "', item " & Text.From(iItem) & "."
                    )
                else if fdmAnt1 >= 0.98 then
                    dFDM
                else if fdmAnt1 = 0.97 and fdmAnt2 <> null and fdmAnt2 > 0.97 then
                    0.96
                else if fdmAnt1 = 0.97 and fdmAnt2 <> null and fdmAnt2 <= 0.97 then
                    0.95
                else if fdmAnt1 = 0.97 and fdmAnt2 = null then
                    0.96
                else if fdmAnt1 = 0.96 then
                    0.95
                else if fdmAnt1 = 0.95 then
                    0.95
                else
                    error Error.Record(
                        "RegraFDMNaoTratada",
                        "Combinação de FDM anterior não prevista na regra 8.4.3: fdmAnt1=" & Text.From(fdmAnt1) & ", fdmAnt2=" & Text.From(fdmAnt2)
                    )
        else
            dFDM
in
    fnRecalcularFDM
;

shared RefFDMHistorico = let
    // ═══════════════════════════════════════════════════════════════════
    // RefFDMHistorico
    // ─────────────────────────────────────────────────────────────────
    // Consulta de referência (não é uma função — é uma tabela) que
    // carrega e bufferiza o histórico de FDM/IAPFARQ por período,
    // contrato e item, a partir da aba "FDM_IAPFARQ" da tabelaB.
    //
    // O QUE FAZ:
    //   1. Abre o arquivo físico da tabelaB (localTabelaB) e extrai a
    //      tabela nomeada "FDM_IAPFARQ".
    //   2. Define explicitamente o tipo de cada coluna:
    //        - "Medição"        → texto (chave de período "AAAA#MM")
    //        - "Contrato"       → texto
    //        - "Item contrato"  → número inteiro (1 ou 2)
    //        - "IAPFARQ (%)"    → número decimal
    //        - "FDM"            → número decimal
    //   3. Aplica Table.Buffer — carrega o resultado em memória UMA
    //      ÚNICA VEZ, evitando que o arquivo físico seja reaberto a
    //      cada chamada de fnValidarPeriodoAnterior ou
    //      fnBuscarFDMAnterior (que são invocadas repetidamente, uma
    //      vez por contrato, dentro de TabelaFDMPorContrato).
    //
    // USADA POR:
    //   - fnValidarPeriodoAnterior (consulta RefFDMHistorico[Medição])
    //   - fnBuscarFDMAnterior (Table.SelectRows sobre RefFDMHistorico)
    //
    // ⚠️ IMPORTANTE:
    //   Esta consulta deve ser referenciada pelo nome exato
    //   "RefFDMHistorico" dentro das duas funções acima. Se você
    //   renomear esta consulta, ajuste também as referências internas
    //   de fnValidarPeriodoAnterior e fnBuscarFDMAnterior.
    // ═══════════════════════════════════════════════════════════════════
    Fonte = Excel.Workbook(File.Contents(localTabelaB), null, true),
    FDM_IAPFARQ_Table = Fonte{[Item = "FDM_IAPFARQ", Kind = "Table"]}[Data],
    #"Tipo Alterado" = Table.TransformColumnTypes(FDM_IAPFARQ_Table, {
        {"Medição", type text},
        {"Contrato", type text},
        {"Item contrato", Int64.Type},
        {"IAPFARQ (%)", type number},
        {"FDM", type number}
    }),

    // Bufferiza o resultado final — evita reabrir o arquivo físico
    // a cada linha/contrato processado nas funções que dependem
    // desta consulta
    RefFDMHistorico = Table.Buffer(#"Tipo Alterado")
in
    RefFDMHistorico
;

shared fnBuscarFDMAnterior = let
    // -------------------------------------------------------------------
    // fnBuscarFDMAnterior
    // -------------------------------------------------------------------
    // Equivalente M da função VBA validar_FDMAnterior (módulo fncTabelaB).
    // -------------------------------------------------------------------
    fnBuscarFDMAnterior = (periodo as text, contrato as text, item as number) as nullable number =>
        let
            Filtrado = Table.SelectRows(RefFDMHistorico, each
                [Medição] = periodo and [Contrato] = contrato and [#"Item contrato"] = item
            ),
            Resultado = if not Table.IsEmpty(Filtrado) then Filtrado{0}[FDM] else null
        in
            Resultado
in
    fnBuscarFDMAnterior
;

shared fnValidarPeriodoAnterior = let
    // -------------------------------------------------------------------
    // fnValidarPeriodoAnterior
    // -------------------------------------------------------------------
    // Equivalente M da função VBA validar_periodo_anterior (módulo fncTabelaB).
    // -------------------------------------------------------------------
    fnValidarPeriodoAnterior = (periodo as nullable text) as logical =>
        if periodo = null or periodo = "" then
            false
        else
            List.Contains(RefFDMHistorico[Medição], periodo)
in
    fnValidarPeriodoAnterior
;

shared fnCalcularFDMMensal = let
    // -------------------------------------------------------------------
    // fnCalcularFDMMensal
    // -------------------------------------------------------------------
    // Equivalente M da função VBA calcular_FDM (módulo fncEditalGuarda)
    // - item nº 8.4.3 do contrato de guarda externa Petrobras.
    //
    // O QUE FAZ:
    //   Classifica o índice IAPFARQ em uma das 4 faixas de FDM (Fator
    //   de Desempenho Mensal), conforme regra contratual:
    //
    //       IAPFARQ < 90        ->  FDM = 0,97
    //       90 <= IAPFARQ < 95  ->  FDM = 0,98
    //       95 <= IAPFARQ < 99  ->  FDM = 0,99
    //       IAPFARQ >= 99       ->  FDM = 1,00
    //
    // MELHORIAS DE ROBUSTEZ:
    //   - Aceita parâmetros nulos (retorna null defensivamente).
    //   - Estabiliza o valor de entrada com arredondamento a 2 casas
    //     decimais antes da comparação, garantindo fidelidade às faixas
    //     e eliminando distorções de ponto flutuante.
    // -------------------------------------------------------------------
    fnCalcularFDMMensal = (IAPFARQ as nullable number) as nullable number =>
        if IAPFARQ = null then
            null
        else
            let
                vIAPFARQ = Number.Round(IAPFARQ, 2)
            in
                if vIAPFARQ < 90.0 then 0.97
                else if vIAPFARQ < 95.0 then 0.98
                else if vIAPFARQ < 99.0 then 0.99
                else 1.00
in
    fnCalcularFDMMensal
;

shared fnCalcularPeriodosAnteriores = let
    // ═══════════════════════════════════════════════════════════════════
    // fnCalcularPeriodosAnteriores
    // ─────────────────────────────────────────────────────────────────
    // Equivalente M do trecho de cálculo de período anterior presente
    // dentro de btMCGuardaExterna/btMC_GE_individual (módulo
    // subMCGuardaExterna) — usado para localizar o FDM histórico dos
    // 2 meses anteriores ao período vigente (item 8.4.3 do contrato).
    //
    // O QUE FAZ:
    //   Calcula as chaves de período no formato "AAAA#MM" para o mês
    //   anterior (PerAnt1) e para o mês anterior a esse (PerAnt2),
    //   tratando corretamente a VIRADA DE ANO:
    //
    //   - PerAnt1 (1 mês atrás):
    //       Se o mês vigente for Janeiro (mes=1), o mês anterior é
    //       Dezembro do ANO ANTERIOR.
    //       Caso contrário, é o mês vigente - 1, no mesmo ano.
    //
    //   - PerAnt2 (2 meses atrás):
    //       Se o mês vigente for Janeiro (mes=1), 2 meses atrás é
    //       Novembro do ANO ANTERIOR.
    //       Se o mês vigente for Fevereiro (mes=2), 2 meses atrás é
    //       Dezembro do ANO ANTERIOR.
    //       Caso contrário, é o mês vigente - 2, no mesmo ano.
    //
    // FORMATO DE SAÍDA:
    //   Cada período é formatado como "AAAA#MM" (ex.: "2026#07"), com
    //   o mês sempre em 2 dígitos (zero à esquerda quando necessário),
    //   compatível com a chave usada na tabela histórica FDM_IAPFARQ
    //   (coluna "Medição").
    //
    // PARÂMETROS:
    //   mes - número do mês vigente (1 a 12)
    //   ano - número do ano vigente (ex.: 2026)
    //
    // RETORNO:
    //   Um record com 2 campos: [PerAnt1, PerAnt2]
    // ═══════════════════════════════════════════════════════════════════
    fnCalcularPeriodosAnteriores = (mes as number, ano as number) as record =>
        let
            // Calcula o período de 1 mês atrás, tratando virada de ano
            PerAnt1 = if mes - 1 = 0
                then Text.From(ano - 1) & "#12"
                else Text.From(ano) & "#" & Text.PadStart(Text.From(mes - 1), 2, "0"),

            // Calcula o período de 2 meses atrás, tratando virada de ano
            // (2 casos possíveis: mês vigente = Janeiro ou Fevereiro)
            PerAnt2 = if mes - 2 = -1
                then Text.From(ano - 1) & "#11"
                else if mes - 2 = 0
                then Text.From(ano - 1) & "#12"
                else Text.From(ano) & "#" & Text.PadStart(Text.From(mes - 2), 2, "0")
        in
            [ PerAnt1 = PerAnt1, PerAnt2 = PerAnt2 ]
in
    fnCalcularPeriodosAnteriores
;

shared fnCalcularTSPE_TSPR = let
    // -------------------------------------------------------------------
    // fnCalcularTSPE_TSPR
    // -------------------------------------------------------------------
    // Equivalente M da função VBA calcular_TSPE_TSPR (módulo fncEditalGuarda)
    // - item nº 8.4.2 do contrato de guarda externa Petrobras.
    //
    // O QUE FAZ:
    //   Calcula, para um contrato específico, os quantitativos de
    //   solicitações atendidas (TSPE) e atendidas no prazo (TSPR),
    //   segregados por linha de serviço (Serviço 1 e Serviço 2,
    //   determinados pela coluna "Ref. FDM"):
    //     - TSPE1 / TSPE2 = Total de solicitações elegíveis no período
    //       (linhas com Situação = "CO" e Status Prazo = "NP" ou "FP")
    //     - TSPR1 / TSPR2 = Total de solicitações atendidas no prazo
    //       (linhas com Situação = "CO" e Status Prazo = "NP")
    // -------------------------------------------------------------------
    fnCalcularTSPE_TSPR = (tabela as table, contratoA as nullable text) as record =>
        let
            vContrato = if contratoA = null then "" else contratoA,

            // Filtra o contrato informado e apenas atendimentos concluídos ("CO")
            Filtrado = Table.SelectRows(tabela, each
                [Contrato] = vContrato and [Situação] = "CO"
            ),

            // Agrupa por Serviço ("Ref. FDM") com agregação vetorizada via lista
            Agrupado = Table.Group(Filtrado, {"Ref. FDM"}, {
                {"TSPE", each
                    let st = [Status Prazo]
                    in List.Count(List.Select(st, each _ = "NP" or _ = "FP")),
                    Int64.Type
                },
                {"TSPR", each
                    let st = [Status Prazo]
                    in List.Count(List.Select(st, each _ = "NP")),
                    Int64.Type
                }
            }),

            Servico1 = Table.SelectRows(Agrupado, each [#"Ref. FDM"] = 1),
            Servico2 = Table.SelectRows(Agrupado, each [#"Ref. FDM"] = 2),

            rTSPE1 = if not Table.IsEmpty(Servico1) then Servico1{0}[TSPE] else 0,
            rTSPR1 = if not Table.IsEmpty(Servico1) then Servico1{0}[TSPR] else 0,
            rTSPE2 = if not Table.IsEmpty(Servico2) then Servico2{0}[TSPE] else 0,
            rTSPR2 = if not Table.IsEmpty(Servico2) then Servico2{0}[TSPR] else 0
        in
            [ TSPE1 = rTSPE1, TSPR1 = rTSPR1, TSPE2 = rTSPE2, TSPR2 = rTSPR2 ]
in
    fnCalcularTSPE_TSPR
;

shared Armazenamento = let
    Combinado = Table.Combine({ConsolidadoES, ConsolidadoRJ, ConsolidadoSP, ConsolidadoPA}),

    // Bufferiza o resultado final — evita reabrir os 4 arquivos físicos
    // a cada referência subsequente desta consulta em outras etapas.
    ResultadoFinal = Table.Buffer(Combinado),
    #"Colunas Renomeadas" = Table.RenameColumns(ResultadoFinal,{{"Contrato", "Linha de serviço PPU"}}),
    #"Colunas Removidas" = Table.RemoveColumns(#"Colunas Renomeadas",{"Ø", "A", "L", "P", "cm³ = V1", "cm³ = V2"}),
    #"Tipo Alterado" = Table.TransformColumnTypes(#"Colunas Removidas",{{"FC = V1/V2", type number}}),
    #"Colunas Removidas1" = Table.RemoveColumns(#"Tipo Alterado",{"$ Unitário", "$ Total", "Tabela", "UA"}),
    #"Tipo Alterado1" = Table.TransformColumnTypes(#"Colunas Removidas1",{{"FC = V1/V2", type number}, {"QUAm-1", type number}, {"QBm", type number}, {"QDm", type number}, {"QNm", type number}, {"QUAm", type number}}),
    #"Valor Substituído" = Table.ReplaceValue(#"Tipo Alterado1",null,0,Replacer.ReplaceValue,{"FC = V1/V2", "QUAm-1", "QBm", "QDm", "QNm", "QUAm"}),
    #"Colunas Renomeadas1" = Table.RenameColumns(#"Valor Substituído",{{"FC = V1/V2", "FC"}}),
    #"Colunas Reordenadas" = Table.ReorderColumns(#"Colunas Renomeadas1",{"Contrato-Reg.", "MEDIDAS CONTRATO", "Padrão", "QUAm-1", "QBm", "QDm", "QNm", "QUAm", "FC", "Linha de serviço PPU"}),
    // QUAm-1 - (QBm + QDm) + QNm
    // 
    QUA = Table.AddColumn(#"Colunas Reordenadas", "QUA", each [#"QUAm-1"]-([QBm]+[QDm]) + [QNm]),
    #"Colunas Reordenadas1" = Table.ReorderColumns(QUA,{"Contrato-Reg.", "MEDIDAS CONTRATO", "Padrão", "QUAm-1", "QBm", "QDm", "QNm", "QUAm", "FC", "QUA", "Linha de serviço PPU"}),
    #"Colunas Renomeadas2" = Table.RenameColumns(#"Colunas Reordenadas1",{{"Padrão", "Item"}})
in
    #"Colunas Renomeadas2";

shared #"Painel T2M" = let
// -------------------------------------------------------------------
// SECAO 0 - TABELAS DE REFERENCIA (BUFFER AQUI, NAO NA TABELA PRINCIPAL)
// -------------------------------------------------------------------
// Cada tabela de lookup usada em Table.NestedJoin mais adiante e
// carregada e "bufferizada" (materializada em memoria) UMA UNICA VEZ
// aqui. Isso evita reabrir os arquivos de origem repetidamente durante
// os merges e garante estabilidade caso a fonte externa seja alterada
// no meio da execucao da consulta.
RefLocalidadeComEndereco = Table.Buffer(#"Localidade Com Endereço"), // Localidade -> Centro/Codigo Centro/Municipio/UF
RefValidUFGalpao = Table.Buffer(validUFGalpao), // Lista de UFs validas por contrato de galpao
RefValEnderecoGalpao = Table.Buffer(valEnderecoGalpao), // Codigo do galpao -> Cidade/UF do galpao
RefTabRefPrazoFDM = Table.Buffer(tabRefPrazoFDM), // Atividade -> Ref. FDM / Ref. QExec / Ref. prazo
RefGrupo = Table.Buffer(Grupo),
RefCalendario = Table.Buffer(Calendario), // Municipio+UF -> Grupo (referencia de prazo)
RefTabelaPrazoContrato = Table.Buffer(
    Table.TransformColumnTypes(tabelaPrazoContrato, {{"Prazo", type number}})
), // Contrato+Atividade+Grupo -> Prazo contratual
// -------------------------------------------------------------------
// SECAO 1 - CARREGAMENTO E EXPANSAO DOS ARQUIVOS DE MONITORAMENTO
// -------------------------------------------------------------------
// Le todos os arquivos da pasta de monitoramento individual (T2M),
// aplica a transformacao padrao de cada arquivo (funcao "Transformar
// Arquivo") e expande o conteudo de todos eles em uma unica tabela.
Fonte = Folder.Files(localMonitIndividualT2M),
// Ignora arquivos ocultos/temporarios e garante que apenas pastas de trabalho Excel (.xlsm / .xlsx) sejam processadas
ArquivosVisiveis = Table.SelectRows(Fonte, each [Attributes]?[Hidden]? <> true and not Text.StartsWith([Name], "~$") and ([Extension] = ".xlsm" or [Extension] = ".xlsx")),
// Aplica a funcao de transformacao (parsing) a cada arquivo
ArquivosTransformados = Table.AddColumn(ArquivosVisiveis, "Transformar Arquivo", each #"Transformar Arquivo"([Content])),
RenomearOrigem = Table.RenameColumns(ArquivosTransformados, {"Name", "Nome da Origem"}),
SelecionarColunas = Table.SelectColumns(RenomearOrigem, {"Nome da Origem", "Transformar Arquivo"}),
// Table.ColumnNames sobre "Arquivo de Amostra" apenas obtem o nome
// das colunas resultantes da transformacao, para poder expandir a
// coluna de tabelas aninhadas - padrao comum em consultas de
// "combinar arquivos de uma pasta"
ExpandirTabela = Table.ExpandTableColumn(
    SelecionarColunas,
    "Transformar Arquivo",
    Table.ColumnNames(#"Transformar Arquivo"(#"Arquivo de Amostra"))
),
// BUFFER 1 - evita reler e reprocessar os arquivos fisicos a cada
// passo seguinte (carregamento de arquivo e a operacao mais custosa
// de toda a consulta)
ExpandirTabelaBuffer = Table.Buffer(ExpandirTabela),
// -------------------------------------------------------------------
// SECAO 2 - LIMPEZA, TIPAGEM E FILTRO DE SITUACAO + CONTRATO
// -------------------------------------------------------------------
RemoverColunas = Table.RemoveColumns(ExpandirTabelaBuffer, {
    "Nome da Origem", "Atendente", "UF", "Nome Solicitante"
}),
DefinirTiposIniciais = Table.TransformColumnTypes(RemoverColunas, {
    {"Situação", type text},
    {"Contrato", type text},
    {"Período#(lf)faturamento", type date}
}),
//  OTIMIZACAO APLICADA: as duas filtragens originais (Situacao="CO"
// e Contrato IRON-LT1/PA-LT2 + periodo nao nulo) faziam DUAS
// passagens completas pela tabela. Combinadas em uma unica chamada
// Table.SelectRows com "and", o motor do Power Query avalia todas as
// condicoes em UMA unica passagem por linha.
FiltrarSituacaoEContrato = Table.SelectRows(DefinirTiposIniciais, each
    [Situação] = "CO"
    and [#"Período#(lf)faturamento"] <> null
    and ([Contrato] = "IRON-LT1" or [Contrato] = "PA-LT2")
),
// BUFFER 2 - cacheia o subconjunto ja bem reduzido (Situacao + Contrato + periodo preenchido)
FiltrarContratosBuffer = Table.Buffer(FiltrarSituacaoEContrato),
// -------------------------------------------------------------------
// SECAO 3 - VALIDACAO DE PERIODO
// -------------------------------------------------------------------
//  OTIMIZACAO CRITICA APLICADA:
// No codigo original, fnGerarPeriodo(PeriodoMes, PeriodoAno),
// Text.Split(...) e Date.FromText(...) eram recalculados A CADA LINHA
// da tabela (dentro do "each" de Table.AddColumn) - mas PeriodoMes e
// PeriodoAno sao PARAMETROS FIXOS para toda a execucao da consulta.
// Ou seja, o mesmo calculo de texto era refeito milhares de vezes sem
// necessidade. Aqui, a data de inicio e fim do periodo de referencia
// e calculada UMA UNICA VEZ, fora do "each", e reutilizada no filtro.
//
// Alem disso, isso elimina a necessidade de: (1) criar uma coluna
// auxiliar "ValidarPeriodo", (2) filtrar por ela, e (3) remove-la
// depois - reduzindo 3 passos do pipeline original para apenas 1.
sPeriodoDefinido = fnGerarPeriodo(PeriodoMes, PeriodoAno),
PartesPeriodo = Text.Split(sPeriodoDefinido, " à "),
DataInicioPeriodo = Date.FromText(PartesPeriodo{0}, [Format = "dd/MM/yyyy"]),
DataFimPeriodo = Date.FromText(PartesPeriodo{1}, [Format = "dd/MM/yyyy"]),
// Filtra diretamente pelo intervalo ja calculado (apenas a conversao
// de data da linha em si - Date.From - precisa ocorrer por linha,
// ja que depende do valor especifico daquela linha)
FiltrarPeriodoValido = Table.SelectRows(FiltrarContratosBuffer, each
    Date.From([#"Período#(lf)faturamento"]) >= DataInicioPeriodo
    and Date.From([#"Período#(lf)faturamento"]) <= DataFimPeriodo
),
// BUFFER 3 - cacheia antes das transformacoes de texto e reordenacao
FiltrarPeriodoValidoBuffer = Table.Buffer(FiltrarPeriodoValido),
// -------------------------------------------------------------------
// SECAO 4 - RENOMEAR, REORDENAR, REMOVER COLUNAS FINAIS E TIPAR O RESTO
// -------------------------------------------------------------------
RenomearColunas = Table.RenameColumns(FiltrarPeriodoValidoBuffer, {{"D. solicitação", "Data solicitação"}}),
// Remove colunas que nao sao utilizadas em nenhuma etapa posterior
// do calculo de FDM/SLA/QExec (campos de embalagem/etiqueta fisica,
// LOG antigo, observacoes livres, etc.)
RemoverColunasFinais = Table.RemoveColumns(RenomearColunas, {
    "OS#(lf)disponibilizada?", "LOG",
    "Caixa#(lf)(20kg)", "Caixa#(lf)(Mídia)",
    "Caixa Tubo#(lf)(perfil de poço)", "Caixa Tubo#(lf)(Engenharia)",
    "Lacre", "Etiqueta (20kg)", "Etiqueta (Mídia)",
    "Etiqueta Tubo#(lf)(perfil de poço)", "Etiqueta Tubo#(lf)(Engenharia)",
    "Lançamento de etiqueta/Registro OS", "Observação"
}),
// Padroniza a ordem das colunas para facilitar leitura/depuracao
ReordenarColunas = Table.ReorderColumns(RemoverColunasFinais, {
    "Período#(lf)faturamento", "Situação", "Contrato",
    "Descrição da atividade", "Item", "Aplicação",
    "Código da#(lf)solicitação", "Chave#(lf)solicitante",
    "Gerência solicitante", "Qtd.#(lf)Solicitada", "Qtd.#(lf)Atendida",
    "Código OS", "Data solicitação", "Prazo#(lf)aplicação",
    "D. abertura", "D. fechamento", "Prazo combinado",
    "Obs. Sigla", "Localidade", "Comentários", "Observações"
}),
DefinirTiposRestantes = Table.TransformColumnTypes(ReordenarColunas, {
    {"Descrição da atividade", type text},
    {"Item", type text},
    {"Aplicação", type text},
    {"Código da#(lf)solicitação", type text},
    {"Chave#(lf)solicitante", type text},
    {"Gerência solicitante", type text},
    {"Código OS", type text},
    {"Qtd.#(lf)Atendida", Int64.Type},
    {"Qtd.#(lf)Solicitada", Int64.Type},
    {"Data solicitação", type datetime},
    {"Prazo#(lf)aplicação", type datetime},
    {"D. abertura", type datetime},
    {"D. fechamento", type datetime},
    {"Prazo combinado", type datetime},
    {"Obs. Sigla", type text},
    {"Localidade", type text},
    {"Comentários", type text},
    {"Observações", type text}
}),
// -------------------------------------------------------------------
// SECAO 5 - TRANSFORMACOES DE TEXTO E ORDENACAO
// -------------------------------------------------------------------
// Padroniza a Descricao da atividade e a Localidade em CAIXA ALTA,
// removendo espacos/caracteres nao-imprimiveis (Text.Clean), para
// garantir correspondencia exata nos merges das Secoes 6-8
TransformarTexto = Table.TransformColumns(DefinirTiposRestantes, {
    {"Descrição da atividade", Text.Upper, type text},
    {"Localidade", each Text.Trim(Text.Clean(Text.Upper(_))), type text}
}),
OrdenarContrato = Table.Sort(TransformarTexto, {{"Contrato", Order.Ascending}}),
// BUFFER 4 - conjunto final antes das mesclagens
OrdenarContratoBuffer = Table.Buffer(OrdenarContrato),
// -------------------------------------------------------------------
// SECAO 6 - ENRIQUECER COM LOCALIDADE E GALPAO
// -------------------------------------------------------------------
// Traz Centro/Codigo Centro/Municipio/UF a partir da Localidade
// PETROBRAS informada na solicitacao
MesclarLocalidade = Table.NestedJoin(
    OrdenarContratoBuffer, {"Localidade"},
    RefLocalidadeComEndereco, {"Localidade"},
    "Localidade Com Endereço", JoinKind.LeftOuter
),
ExpandirLocalidade = Table.ExpandTableColumn(MesclarLocalidade, "Localidade Com Endereço",
    {"Centro", "Código Centro", "Município", "UF"}, {"Centro", "Código Centro", "Município", "UF"}),
// Marca se a atividade e uma das 7 que permitem calculo de KM
// adicional/validacao de galpao (mesma lista usada em
// fnValidarLocalizacao)
AdicionarValidarGalpao = Table.AddColumn(ExpandirLocalidade, "ValidarGalpao", each
    if List.Contains({
        "DEVOLUCAO DE EMPRESTIMO",
        "COLETA DE EMBALAGEM",
        "COLETA DE ITEM AVULSO",
        "ENTREGA DE EMBALAGEM EXPRESSO",
        "ENTREGA DE EMBALAGEM NORMAL",
        "ENTREGA DE ITEM EXPRESSO",
        "ENTREGA DE ITEM NORMAL"
    }, Text.Upper(Text.Trim([Descrição da atividade])))
    then "Verdadeiro"
    else "Falso"
),
// Converte a sigla informada (Obs. Sigla) no codigo de galpao
// correspondente (ex.: "GRJ" -> "IRON-RJ")
AdicionarObsSigla = Table.AddColumn(AdicionarValidarGalpao, "ObsSiglaGalpao", each
    if [Obs. Sigla] = "GRJ" then "IRON-RJ"
    else if [Obs. Sigla] = "GES" then "IRON-ES"
    else if [Obs. Sigla] = "GSP" then "IRON-SP"
    else if [Obs. Sigla] = "GBA" then "PA-BA"
    else null
),
// Valida se a UF de destino e atendida por algum galpao cadastrado
MesclarUFGalpao = Table.NestedJoin(
    AdicionarObsSigla, {"UF"},
    RefValidUFGalpao, {"UF"},
    "validUFGalpao", JoinKind.LeftOuter
),
ExpandirUFGalpao = Table.ExpandTableColumn(
    MesclarUFGalpao, "validUFGalpao",
    {"Contrato"}, {"Contrato.1"}
),
// Determina o codigo do Galpao final: "N/A" se a atividade nao
// permite (ValidarGalpao="Falso"), ou o galpao da UF de destino
// (ValidarGalpao="Verdadeiro")
AdicionarGalpao = Table.AddColumn(ExpandirUFGalpao, "Galpão", each
    if [ValidarGalpao] = "Falso" then "N/A"
    else if [ValidarGalpao] = "Verdadeiro" then [Contrato.1]
    else null
),
// Traz Cidade/UF do galpao a partir do codigo do Galpao
MesclarEndGalpao = Table.NestedJoin(
    AdicionarGalpao, {"Galpão"},
    RefValEnderecoGalpao, {"Galpãp"},
    "valEnderecoGalpao", JoinKind.LeftOuter
),
ExpandirEndGalpao = Table.ExpandTableColumn(
    MesclarEndGalpao, "valEnderecoGalpao",
    {"Galpãp_Cidade", "Galpãp_UF"},
    {"Galpãp_Cidade", "Galpãp_UF"}
),
// -------------------------------------------------------------------
// SECAO 7 - CALCULO DE DUPLICATAS PARA FDM
// -------------------------------------------------------------------
// A mesma solicitacao (Atividade + Codigo da solicitacao + Codigo OS)
// pode gerar multiplas linhas no monitoramento (ex.: atendimentos de
// digitalizacao separados por item A4/A3/A2). Para o calculo de FDM,
// apenas a PRIMEIRA ocorrencia de cada combinacao deve ser
// considerada "nova" (N); as demais sao tratadas como duplicata (S).
//
// TECNICA: em vez de um loop percorrendo linha a linha com CountIf
// (como no VBA legado), usa-se Table.Group para encontrar o indice
// minimo de cada chave, e depois um merge para trazer esse indice de
// volta a cada linha - abordagem muito mais performatica, pois evita
// N2 comparacoes.
AdicionarColunaConcatenada = Table.AddColumn(ExpandirEndGalpao, "ColunaConcatenadaCalculoFDN", each
    [Descrição da atividade] & "/" & [#"Código da#(lf)solicitação"] & "/" & [Código OS]
),
AdicionarIndice = Table.AddIndexColumn(AdicionarColunaConcatenada, "Indice", 0, 1, Int64.Type),
// BUFFER 5 - garante indice estavel antes do Group By / Merge
AdicionarIndiceBuffer = Table.Buffer(AdicionarIndice),
// Para cada chave concatenada, encontra o menor indice (= primeira
// ocorrencia daquela combinacao na tabela)
PrimeiraOcorrencia = Table.Group(
    AdicionarIndiceBuffer,
    {"ColunaConcatenadaCalculoFDN"},
    {{"IndiceMinimo", each List.Min([Indice]), Int64.Type}}
),
// BUFFER 6 - cacheia o resultado do agrupamento antes do merge
PrimeiraOcorrenciaBuffer = Table.Buffer(PrimeiraOcorrencia),
MesclarPrimeiraOcorrencia = Table.NestedJoin(
    AdicionarIndiceBuffer, {"ColunaConcatenadaCalculoFDN"},
    PrimeiraOcorrenciaBuffer, {"ColunaConcatenadaCalculoFDN"},
    "PrimeiraOc", JoinKind.LeftOuter
),
ExpandirIndiceMinimo = Table.ExpandTableColumn(
    MesclarPrimeiraOcorrencia, "PrimeiraOc",
    {"IndiceMinimo"}, {"IndiceMinimo"}
),
// S = duplicata (ocorrencia seguinte) | N = primeira ocorrencia
AdicionarCalculoFDM = Table.AddColumn(ExpandirIndiceMinimo, "CalculoFDM", each
    if [Indice] = [IndiceMinimo] then "N" else "S"
),
RemoverAuxiliares = Table.RemoveColumns(AdicionarCalculoFDM, {"Indice", "IndiceMinimo"}),
// -------------------------------------------------------------------
// SECAO 8 - REFERENCIAS DE PRAZO, GRUPO E CONTRATO
// -------------------------------------------------------------------
// Traz da tabelaA (via RefTabRefPrazoFDM) a classificacao da
// atividade: se entra no escopo de FDM (Ref. FDM), se conta para
// QExec (Ref. QExec) e qual e a referencia de prazo (Tabela x
// Fiscalizacao)
MesclarRefPrazoFDM = Table.NestedJoin(
    RemoverAuxiliares, {"Descrição da atividade"},
    RefTabRefPrazoFDM, {"Atividade do painel"},
    "tabRefPrazoFDM", JoinKind.LeftOuter
),
ExpandirRefPrazoFDM = Table.ExpandTableColumn(
    MesclarRefPrazoFDM, "tabRefPrazoFDM",
    {"Ref. FDM", "Ref. QExec", "Ref. prazo"},
    {"Ref. FDM", "Ref. QExec", "Ref. prazo"}
),
// Monta a chave Municipio\UF para descobrir o Grupo de referencia
// de prazo (usado no calculo do SLA)
AdicionarChaveMunicipioUF = Table.AddColumn(ExpandirRefPrazoFDM, "codContDataMuniUF", each
    [Município] & "\" & [UF]
),
MesclarGrupoMunicipio = Table.NestedJoin(
    AdicionarChaveMunicipioUF, {"codContDataMuniUF"},
    RefGrupo, {"codGrupo"},
    "Grupo", JoinKind.LeftOuter
),
ExpandirGrupoMunicipio = Table.ExpandTableColumn(
    MesclarGrupoMunicipio, "Grupo",
    {"Grupo"}, {"Grupo.1"}
),
// Monta a chave Contrato\Atividade\Grupo para descobrir o Prazo
// contratual especifico daquela combinacao
AdicionarChavePrazoContrato = Table.AddColumn(ExpandirGrupoMunicipio, "codPrazoContrato", each
    [Contrato] & "\" & [Descrição da atividade] & "\" & [Grupo.1]
),
MesclarPrazoContrato = Table.NestedJoin(
    AdicionarChavePrazoContrato, {"codPrazoContrato"},
    RefTabelaPrazoContrato, {"codPrazoContrato"},
    "tabelaPrazoContrato", JoinKind.LeftOuter
),
ExpandirPrazoContrato = Table.ExpandTableColumn(
    MesclarPrazoContrato, "tabelaPrazoContrato",
    {"Prazo"}, {"tabelaPrazoContrato.Prazo"}
),
// -------------------------------------------------------------------
// SECAO 9 - SEPARACAO DE DATA/HORA
// -------------------------------------------------------------------
// Separa Data e Hora de Abertura/Fechamento em colunas independentes
// (necessario para o calculo de SLA, que trata data e hora
// separadamente - expediente 08h-17h, almoco 12h-13h)
AdicionarDataHoraSeparadas = Table.AddColumn(ExpandirPrazoContrato, "DataHoraAux", each
    [
        #"Data Abertura" = DateTime.Date([D. abertura]),
        #"Hora Abertura" = DateTime.Time([D. abertura]),
        #"Data Fechamento" = DateTime.Date([D. fechamento]),
        #"Hora Fechamento" = DateTime.Time([D. fechamento])
    ]
),
ExpandirDataHoraSeparadas = Table.ExpandRecordColumn(
    AdicionarDataHoraSeparadas, "DataHoraAux",
    {"Data Abertura", "Hora Abertura", "Data Fechamento", "Hora Fechamento"},
    {"Data Abertura", "Hora Abertura", "Data Fechamento", "Hora Fechamento"}
),
// -------------------------------------------------------------------
// SECAO 10 - CLASSIFICACAO FDM/SLA, QEXEC, KM ADICIONAL
// -------------------------------------------------------------------
// Classifica cada linha em Normal/Expresso/Isento/FDM definido e
// NP/FP, calculando o SLA internamente (fnClassificarFDM). Envolvida
// em try/otherwise para que uma linha com dado inconsistente nao
// interrompa o processamento de toda a tabela - o erro e registrado
// na propria coluna de LOG.
AdicionarClassificacaoFDM = Table.AddColumn(ExpandirDataHoraSeparadas, "ResultadoFDM", each
    try
        fnClassificarFDM(
            [CalculoFDM], [#"Ref. FDM"], [#"Ref. prazo"],
            [#"Descrição da atividade"], [Contrato],
            [Data Abertura], [Hora Abertura],
            [D. fechamento], [Prazo combinado],
            [Município], [UF], [Grupo.1],
            [#"tabelaPrazoContrato.Prazo"], [Aplicação], RefCalendario
        )
    otherwise
        [ Classificacao = null, StatusPrazo = null, Log = "Erro no calculo - verificar dados da linha; " ],
    type record
),
ExpandirClassificacaoFDM = Table.ExpandRecordColumn(
    AdicionarClassificacaoFDM, "ResultadoFDM",
    {"Classificacao", "StatusPrazo", "Log"},
    {"Classificação FDM", "Status Prazo", "LOG SLA"}
),
//  CORRECAO DE ROBUSTEZ APLICADA: a concatenacao original com "&"
// falha com erro se [Item] for null. Usando o operador "??"
// (coalescencia de nulo), qualquer campo nulo e tratado como texto
// vazio, preservando o resultado quando ambos os campos existem e
// evitando que a consulta inteira quebre por uma unica linha
// incompleta.
AdicionarCodTabelaFC = Table.AddColumn(ExpandirClassificacaoFDM, "cod_Tabela_FC", each
    ([Descrição da atividade] ?? "") & ([Item] ?? "")
),
// Traz o Fator de Conversao (FC) especifico daquela Atividade+Item,
// usado por fnCalcularQExec nas atividades de Migracao/Digitalizacao
MesclarTabelaFC = Table.NestedJoin(AdicionarCodTabelaFC, {"cod_Tabela_FC"}, tabela_FC, {"cod_Tabela_FC"}, "tabela_FC", JoinKind.LeftOuter),
ExpandirTabelaFC = Table.ExpandTableColumn(MesclarTabelaFC, "tabela_FC", {"FC"}, {"FC"}),
// Calcula a quantidade padronizada de medicao (QExec) por linha,
// conforme a atividade/item/quantidade/fator de conversao
AdicionarQExec = Table.AddColumn(ExpandirTabelaFC, "QExec", each
    fnCalcularQExec([Descrição da atividade], [Item], [#"Qtd.#(lf)Atendida"], null, [FC])
),
// Traz Item PPU, Unidade e Linha de servico PPU associados a
// combinacao Atividade+Item (mesma chave "cod_Tabela_FC" ja criada)
MesclarAtividadeItemPPU = Table.NestedJoin(AdicionarQExec, {"cod_Tabela_FC"}, tabelaAtividadeItemPPU, {"CodLinhaPPU"}, "tabelaAtividadeItemPPU", JoinKind.LeftOuter),
ExpandirAtividadeItemPPU = Table.ExpandTableColumn(MesclarAtividadeItemPPU, "tabelaAtividadeItemPPU", {"Item PPU", "Unidade", "Linha de serviço PPU"}, {"Item PPU", "Unidade", "Linha de serviço PPU"}),
//  CORRECAO DE ROBUSTEZ APLICADA: mesma protecao contra nulos
// (operador "??") na chave de origem/destino, evitando erro se
// Centro/Municipio/Galpao/Cidade do Galpao estiverem vazios
AdicionarCodOrigemDestino = Table.AddColumn(ExpandirAtividadeItemPPU, "codOrigemDestino", each
    ([Centro] ?? "") & ([Município] ?? "") & ([Galpão] ?? "") & ([Galpãp_Cidade] ?? "")
),
// Traz a distancia (KM) cadastrada para a rota Centro<->Galpao
MesclarOrigemDestino = Table.NestedJoin(AdicionarCodOrigemDestino, {"codOrigemDestino"}, Origem_Destino, {"codOrigenDestino"}, "Origem_Destino", JoinKind.LeftOuter),
ExpandirOrigemDestinoKM = Table.ExpandTableColumn(MesclarOrigemDestino, "Origem_Destino", {"KM"}, {"Origem_Destino.KM"}),
// -------------------------------------------------------------------
// AdicionarKMAdicional - [OK] CORRIGIDA (deduplicacao por rota+data+tipo)
// -----------------------------------------------------------------
// Equivalente M de CalcularQuilometragemAdicional (Sistema 2, modulo
// A_Calc_Medicao) - que, diferente do Sistema 1 legado
// (subMCGuardaExterna/btMCGuardaExterna, confirmado sem essa
// deduplicacao), aplica deduplicacao por chave rota+data+tipo de
// servico via Collection, evitando cobrar o mesmo KM adicional
// multiplas vezes para atendimentos da mesma viagem.
//
// A chave de deduplicacao combina:
//   - codOrigemDestino (Centro+Municipio+Galpao+Cidade Galpao)
//   - Data Abertura
//   - Linha de servico PPU (equivalente ao "linhaServico" do VBA,
//     que verifica FRE-NRM/FRE-EXP)
//
// Apenas a PRIMEIRA ocorrencia de cada combinacao recebe o KM
// Adicional calculado; as demais linhas do mesmo grupo ficam null -
// replicando fielmente o comportamento do Collection VBA
// (chaveConcatenadaProcessada.Add ... If Err.Number = 0).
// -------------------------------------------------------------------
// Monta a chave de deduplicacao (rota + data + tipo de servico)
AdicionarChaveKM = Table.AddColumn(ExpandirOrigemDestinoKM, "_ChaveKM", each
    Text.From([codOrigemDestino] ?? "") & "|"
    & Text.From([Data Abertura], "pt-BR") & "|"
    & Text.From([#"Linha de serviço PPU"] ?? "")
),
// Indice sequencial estavel, para localizar a 1a ocorrencia de cada chave
AdicionarIndiceKM = Table.AddIndexColumn(AdicionarChaveKM, "_IndiceKM", 0, 1, Int64.Type),
// Agrupa por chave: acha o indice minimo (1a ocorrencia) de cada grupo
AgrupadoKM = Table.Group(AdicionarIndiceKM, {"_ChaveKM"}, {
    {"_IndiceMinimoKM", each List.Min([_IndiceKM]), Int64.Type}
}),
// Traz de volta o indice minimo do grupo para cada linha original
MesclarIndiceKM = Table.NestedJoin(AdicionarIndiceKM, {"_ChaveKM"},
    AgrupadoKM, {"_ChaveKM"}, "_GrupoKM", JoinKind.LeftOuter),
ExpandirIndiceKM = Table.ExpandTableColumn(MesclarIndiceKM, "_GrupoKM", {"_IndiceMinimoKM"}, {"_IndiceMinimoKM"}),
// Calcula o KM Adicional SOMENTE na 1a ocorrencia do grupo (mesma
// rota+data+tipo); as demais linhas do grupo ficam null - evitando
// cobranca duplicada do mesmo frete adicional em multiplos
// atendimentos da mesma viagem. Mantem a regra original de raio de
// tolerancia (valKmAdicional): null se rota nao encontrada; 0 se
// dentro do raio; excedente acima do raio, caso contrario.
AdicionarKMAdicional = Table.AddColumn(ExpandirIndiceKM, "KM Adicional", each
    if [_IndiceKM] <> [_IndiceMinimoKM] then null
    else if [Origem_Destino.KM] = null then null
    else if [Origem_Destino.KM] <= valKmAdicional then 0
    else [Origem_Destino.KM] - valKmAdicional
),
// Remove as colunas auxiliares de deduplicacao, mantendo apenas o
// resultado final "KM Adicional"
LimpezaColunasKM = Table.RemoveColumns(AdicionarKMAdicional, {"_ChaveKM", "_IndiceKM", "_IndiceMinimoKM"}),
// -------------------------------------------------------------------
// SECAO 11 - QEXEC AGRUPADO (calculo adicional/paralelo)
// -------------------------------------------------------------------
// Cria um indice de agrupamento por Data Fechamento + Localidade
// PETROBRAS + Cidade do Galpao + Linha de Servico PPU (FRE-NRM vs FRE-EXP);
// soma o QExec de todas as linhas do grupo e atribui o resultado da formula
// de frete (soma/10, piso 1) somente na linha de primeira ocorrencia do grupo
// (as demais linhas do mesmo grupo ficam com null nesta coluna).
AdicionarQExecAgrupado = fnCalcularQExecAgrupado(
    LimpezaColunasKM,
    "Data Fechamento",
    "Localidade",
    "Galpãp_Cidade",
    "Linha de serviço PPU",
    "QExec"
),
// -------------------------------------------------------------------
// BUFFER FINAL - RECOMENDADO
// -------------------------------------------------------------------
// Esta consulta e referenciada por diversas outras consultas
// independentes (ex.: SaidaMC, SaidaDADOS, SaidaMEDICAO,
// TabelaFDMPorContrato, SaidaFretes). Sem este buffer, CADA UMA
// dessas consultas reexecuta TODO o pipeline acima (todos os merges,
// chamadas de funcao customizada, calculo de FDM/SLA/QExec) do zero,
// de forma independente. Bufferizar o resultado final materializa a
// tabela em memoria uma unica vez, reduzindo o reprocessamento
// repetido nas consultas dependentes - o custo e maior uso de memoria
// durante a atualizacao.
ResultadoFinalBuffer = Table.Buffer(AdicionarQExecAgrupado)
in
    ResultadoFinalBuffer
;

shared fnConverterMesTextoParaNumero = let
    // ═══════════════════════════════════════════════════════════════════
    // fnConverterMesTextoParaNumero
    // ─────────────────────────────────────────────────────────────────
    // Converte o nome do mês em português (ex.: "Julho") para seu
    // número correspondente (1 a 12). Necessária porque o parâmetro
    // PeriodoMes do pipeline está definido como TEXTO (nome do mês),
    // enquanto fnCalcularPeriodosAnteriores e fnRecalcularFDM operam
    // com o mês em formato NUMÉRICO (equivalente ao "iPerMes" do VBA,
    // que já vinha convertido via Left(Planilha6.Range("C9").Value,2)
    // na origem — aqui replicamos essa conversão de forma explícita).
    //
    // PARÂMETROS:
    //   mesTexto - nome do mês em português, com a mesma grafia usada
    //              no parâmetro PeriodoMes (ex.: "Julho")
    //
    // RETORNO:
    //   number — o número do mês (1 a 12)
    //
    // Lança erro explícito se o texto não corresponder a nenhum dos
    // 12 meses cadastrados (evita falha silenciosa/confusa mais
    // adiante no cálculo de períodos anteriores).
    // ═══════════════════════════════════════════════════════════════════
    fnConverterMesTextoParaNumero = (mesTexto as text) as number =>
        let
            Meses = {
                "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
                "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"
            },
            Indice = List.PositionOf(Meses, mesTexto)
        in
            if Indice = -1 then
                error Error.Record(
                    "MesInvalido",
                    "Mês não reconhecido: '" & mesTexto & "'. Esperado um dos 12 nomes de mês em português (ex.: 'Julho')."
                )
            else
                Indice + 1
in
    fnConverterMesTextoParaNumero
;

shared #"Painel otimizado" = let
// -------------------------------------------------------------------
// SECAO 0 - TABELAS DE REFERENCIA (BUFFER AQUI, NAO NA TABELA PRINCIPAL)
// -------------------------------------------------------------------
RefLocalidadeComEndereco = Table.Buffer(#"Localidade Com Endereço"),
RefValidUFGalpao = Table.Buffer(validUFGalpao),
RefValEnderecoGalpao = Table.Buffer(valEnderecoGalpao),
RefTabRefPrazoFDM = Table.Buffer(tabRefPrazoFDM),
RefGrupo = Table.Buffer(Grupo),
RefCalendario = Table.Buffer(Calendario),
RefTabelaPrazoContrato = Table.Buffer(
    Table.TransformColumnTypes(tabelaPrazoContrato, {{"Prazo", type number}})
),
// -------------------------------------------------------------------
// SECAO 1 - CARREGAMENTO E EXPANSAO
// -------------------------------------------------------------------
Fonte = Folder.Files(localMonitIndividualT2M),
ArquivosVisiveis = Table.SelectRows(Fonte, each [Attributes]?[Hidden]? <> true and not Text.StartsWith([Name], "~$") and ([Extension] = ".xlsm" or [Extension] = ".xlsx")),
ArquivosTransformados = Table.AddColumn(ArquivosVisiveis, "Transformar Arquivo", each #"Transformar Arquivo"([Content])),
RenomearOrigem = Table.RenameColumns(ArquivosTransformados, {"Name", "Nome da Origem"}),
SelecionarColunas = Table.SelectColumns(RenomearOrigem, {"Nome da Origem", "Transformar Arquivo"}),
ExpandirTabela = Table.ExpandTableColumn(
    SelecionarColunas,
    "Transformar Arquivo",
    Table.ColumnNames(#"Transformar Arquivo"(#"Arquivo de Amostra"))
),
// BUFFER 1 - evita reler e reprocessar os arquivos fisicos
ExpandirTabelaBuffer = Table.Buffer(ExpandirTabela),
// -------------------------------------------------------------------
// SECAO 2 - LIMPEZA, TIPAGEM E FILTRO DE SITUACAO
// -------------------------------------------------------------------
RemoverColunas = Table.RemoveColumns(ExpandirTabelaBuffer, {
    "Nome da Origem", "Atendente", "UF", "Nome Solicitante"
}),
DefinirTiposIniciais = Table.TransformColumnTypes(RemoverColunas, {
    {"Situação", type text},
    {"Contrato", type text},
    {"Período#(lf)faturamento", type date}
}),
FiltrarSituacao = Table.SelectRows(DefinirTiposIniciais, each [Situação] = "CO"),
FiltrarContratos = Table.SelectRows(FiltrarSituacao, each
    ([#"Período#(lf)faturamento"] <> null) and
    ([Contrato] = "IRON-LT1" or [Contrato] = "PA-LT2")
),
// BUFFER 2 - cacheia o subconjunto ja bem reduzido (Situacao + Contrato)
FiltrarContratosBuffer = Table.Buffer(FiltrarContratos),
// -------------------------------------------------------------------
// SECAO 3 - VALIDACAO DE PERIODO
// -------------------------------------------------------------------
ValidarPeriodo = Table.AddColumn(FiltrarContratosBuffer, "ValidarPeriodo", each
    let
        sPeriodo = fnGerarPeriodo(PeriodoMes, PeriodoAno),
        partes = Text.Split(sPeriodo, " à "),
        dataInicio = Date.FromText(partes{0}, [Format="dd/MM/yyyy"]),
        dataFim = Date.FromText(partes{1}, [Format="dd/MM/yyyy"]),
        dataFat = Date.From([#"Período#(lf)faturamento"])
    in
        if dataFat >= dataInicio and dataFat <= dataFim then "S" else "N"
),
FiltrarPeriodoValido = Table.SelectRows(ValidarPeriodo, each [ValidarPeriodo] = "S"),
RemoverColunaValidacao = Table.RemoveColumns(FiltrarPeriodoValido, {"ValidarPeriodo"}),
// BUFFER 3 - cacheia antes das transformacoes de texto e reordenacao
FiltrarPeriodoValidoBuffer = Table.Buffer(RemoverColunaValidacao),
// -------------------------------------------------------------------
// SECAO 4 - RENOMEAR, REORDENAR, REMOVER COLUNAS FINAIS E TIPAR O RESTO
// -------------------------------------------------------------------
RenomearColunas = Table.RenameColumns(FiltrarPeriodoValidoBuffer, {{"D. solicitação", "Data solicitação"}}),
RemoverColunasFinais = Table.RemoveColumns(RenomearColunas, {
    "OS#(lf)disponibilizada?", "LOG",
    "Caixa#(lf)(20kg)", "Caixa#(lf)(Mídia)",
    "Caixa Tubo#(lf)(perfil de poço)", "Caixa Tubo#(lf)(Engenharia)",
    "Lacre", "Etiqueta (20kg)", "Etiqueta (Mídia)",
    "Etiqueta Tubo#(lf)(perfil de poço)", "Etiqueta Tubo#(lf)(Engenharia)",
    "Lançamento de etiqueta/Registro OS", "Observação"
}),
ReordenarColunas = Table.ReorderColumns(RemoverColunasFinais, {
    "Período#(lf)faturamento", "Situação", "Contrato",
    "Descrição da atividade", "Item", "Aplicação",
    "Código da#(lf)solicitação", "Chave#(lf)solicitante",
    "Gerência solicitante", "Qtd.#(lf)Solicitada", "Qtd.#(lf)Atendida",
    "Código OS", "Data solicitação", "Prazo#(lf)aplicação",
    "D. abertura", "D. fechamento", "Prazo combinado",
    "Obs. Sigla", "Localidade", "Comentários", "Observações"
}),
DefinirTiposRestantes = Table.TransformColumnTypes(ReordenarColunas, {
    {"Descrição da atividade", type text},
    {"Item", type text},
    {"Aplicação", type text},
    {"Código da#(lf)solicitação", type text},
    {"Chave#(lf)solicitante", type text},
    {"Gerência solicitante", type text},
    {"Código OS", type text},
    {"Qtd.#(lf)Atendida", Int64.Type},
    {"Qtd.#(lf)Solicitada", Int64.Type},
    {"Data solicitação", type datetime},
    {"Prazo#(lf)aplicação", type datetime},
    {"D. abertura", type datetime},
    {"D. fechamento", type datetime},
    {"Prazo combinado", type datetime},
    {"Obs. Sigla", type text},
    {"Localidade", type text},
    {"Comentários", type text},
    {"Observações", type text}
}),
// -------------------------------------------------------------------
// SECAO 5 - TRANSFORMACOES DE TEXTO E ORDENACAO
// -------------------------------------------------------------------
TransformarTexto = Table.TransformColumns(DefinirTiposRestantes, {
    {"Descrição da atividade", Text.Upper, type text},
    {"Localidade", each Text.Trim(Text.Clean(Text.Upper(_))), type text}
}),
OrdenarContrato = Table.Sort(TransformarTexto, {{"Contrato", Order.Ascending}}),
// BUFFER 4 - conjunto final antes das mesclagens
OrdenarContratoBuffer = Table.Buffer(OrdenarContrato),
// -------------------------------------------------------------------
// SECAO 6 - ENRIQUECER COM LOCALIDADE E GALPAO
// -------------------------------------------------------------------
MesclarLocalidade = Table.NestedJoin(
    OrdenarContratoBuffer, {"Localidade"},
    RefLocalidadeComEndereco, {"Localidade"},
    "Localidade Com Endereço", JoinKind.LeftOuter
),
ExpandirLocalidade = Table.ExpandTableColumn(
    MesclarLocalidade, "Localidade Com Endereço",
    {"Centro", "Código Centro", "Município", "UF"},
    {"Centro", "Código Centro", "Município", "UF"}
),
AdicionarValidarGalpao = Table.AddColumn(ExpandirLocalidade, "ValidarGalpao", each
    if List.Contains({
        "DEVOLUCAO DE EMPRESTIMO",
        "COLETA DE EMBALAGEM",
        "COLETA DE ITEM AVULSO",
        "ENTREGA DE EMBALAGEM EXPRESSO",
        "ENTREGA DE EMBALAGEM NORMAL",
        "ENTREGA DE ITEM EXPRESSO",
        "ENTREGA DE ITEM NORMAL"
    }, Text.Upper(Text.Trim([Descrição da atividade])))
    then "Verdadeiro"
    else "Falso"
),
AdicionarObsSigla = Table.AddColumn(AdicionarValidarGalpao, "ObsSiglaGalpao", each
    if [Obs. Sigla] = "GRJ" then "IRON-RJ"
    else if [Obs. Sigla] = "GES" then "IRON-ES"
    else if [Obs. Sigla] = "GSP" then "IRON-SP"
    else if [Obs. Sigla] = "GBA" then "PA-BA"
    else null
),
MesclarUFGalpao = Table.NestedJoin(
    AdicionarObsSigla, {"UF"},
    RefValidUFGalpao, {"UF"},
    "validUFGalpao", JoinKind.LeftOuter
),
ExpandirUFGalpao = Table.ExpandTableColumn(
    MesclarUFGalpao, "validUFGalpao",
    {"Contrato"}, {"Contrato.1"}
),
AdicionarGalpao = Table.AddColumn(ExpandirUFGalpao, "Galpão", each
    if [ValidarGalpao] = "Falso" then "N/A"
    else if [ValidarGalpao] = "Verdadeiro" then [Contrato.1]
    else null
),
MesclarEndGalpao = Table.NestedJoin(
    AdicionarGalpao, {"Galpão"},
    RefValEnderecoGalpao, {"Galpãp"},
    "valEnderecoGalpao", JoinKind.LeftOuter
),
ExpandirEndGalpao = Table.ExpandTableColumn(
    MesclarEndGalpao, "valEnderecoGalpao",
    {"Galpãp_Cidade", "Galpãp_UF"},
    {"Galpãp_Cidade", "Galpãp_UF"}
),
// -------------------------------------------------------------------
// SECAO 7 - CALCULO DE DUPLICATAS (FDM)
// -------------------------------------------------------------------
AdicionarColunaConcatenada = Table.AddColumn(ExpandirEndGalpao, "ColunaConcatenadaCalculoFDN", each
    [Descrição da atividade] & "/" & [#"Código da#(lf)solicitação"] & "/" & [Código OS]
),
AdicionarIndice = Table.AddIndexColumn(AdicionarColunaConcatenada, "Indice", 0, 1, Int64.Type),
// BUFFER 5 - garante indice estavel antes do Group By / Merge
AdicionarIndiceBuffer = Table.Buffer(AdicionarIndice),
PrimeiraOcorrencia = Table.Group(
    AdicionarIndiceBuffer,
    {"ColunaConcatenadaCalculoFDN"},
    {{"IndiceMinimo", each List.Min([Indice]), Int64.Type}}
),
// BUFFER 6 - cacheia o resultado do agrupamento antes do merge
PrimeiraOcorrenciaBuffer = Table.Buffer(PrimeiraOcorrencia),
MesclarPrimeiraOcorrencia = Table.NestedJoin(
    AdicionarIndiceBuffer, {"ColunaConcatenadaCalculoFDN"},
    PrimeiraOcorrenciaBuffer, {"ColunaConcatenadaCalculoFDN"},
    "PrimeiraOc", JoinKind.LeftOuter
),
ExpandirIndiceMinimo = Table.ExpandTableColumn(
    MesclarPrimeiraOcorrencia, "PrimeiraOc",
    {"IndiceMinimo"}, {"IndiceMinimo"}
),
// S = duplicata (ocorrencia seguinte) | N = primeira ocorrencia
AdicionarCalculoFDM = Table.AddColumn(ExpandirIndiceMinimo, "CalculoFDM", each
    if [Indice] = [IndiceMinimo] then "N" else "S"
),
RemoverAuxiliares = Table.RemoveColumns(AdicionarCalculoFDM, {"Indice", "IndiceMinimo"}),
// -------------------------------------------------------------------
// SECAO 8 - REFERENCIAS DE PRAZO, GRUPO E CONTRATO
// -------------------------------------------------------------------
MesclarRefPrazoFDM = Table.NestedJoin(
    RemoverAuxiliares, {"Descrição da atividade"},
    RefTabRefPrazoFDM, {"Atividade do painel"},
    "tabRefPrazoFDM", JoinKind.LeftOuter
),
ExpandirRefPrazoFDM = Table.ExpandTableColumn(
    MesclarRefPrazoFDM, "tabRefPrazoFDM",
    {"Ref. FDM", "Ref. QExec", "Ref. prazo"},
    {"Ref. FDM", "Ref. QExec", "Ref. prazo"}
),
AdicionarChaveMunicipioUF = Table.AddColumn(ExpandirRefPrazoFDM, "codContDataMuniUF", each
    [Município] & "\" & [UF]
),
MesclarGrupoMunicipio = Table.NestedJoin(
    AdicionarChaveMunicipioUF, {"codContDataMuniUF"},
    RefGrupo, {"codGrupo"},
    "Grupo", JoinKind.LeftOuter
),
ExpandirGrupoMunicipio = Table.ExpandTableColumn(
    MesclarGrupoMunicipio, "Grupo",
    {"Grupo"}, {"Grupo.1"}
),
AdicionarChavePrazoContrato = Table.AddColumn(ExpandirGrupoMunicipio, "codPrazoContrato", each
    [Contrato] & "\" & [Descrição da atividade] & "\" & [Grupo.1]
),
MesclarPrazoContrato = Table.NestedJoin(
    AdicionarChavePrazoContrato, {"codPrazoContrato"},
    RefTabelaPrazoContrato, {"codPrazoContrato"},
    "tabelaPrazoContrato", JoinKind.LeftOuter
),
ExpandirPrazoContrato = Table.ExpandTableColumn(
    MesclarPrazoContrato, "tabelaPrazoContrato",
    {"Prazo"}, {"tabelaPrazoContrato.Prazo"}
),
// -------------------------------------------------------------------
// SECAO 9 - SEPARACAO DE DATA/HORA
// -------------------------------------------------------------------
AdicionarDataHoraSeparadas = Table.AddColumn(ExpandirPrazoContrato, "DataHoraAux", each
    [
        #"Data Abertura" = DateTime.Date([D. abertura]),
        #"Hora Abertura" = DateTime.Time([D. abertura]),
        #"Data Fechamento" = DateTime.Date([D. fechamento]),
        #"Hora Fechamento" = DateTime.Time([D. fechamento])
    ]
),
ExpandirDataHoraSeparadas = Table.ExpandRecordColumn(
    AdicionarDataHoraSeparadas, "DataHoraAux",
    {"Data Abertura", "Hora Abertura", "Data Fechamento", "Hora Fechamento"},
    {"Data Abertura", "Hora Abertura", "Data Fechamento", "Hora Fechamento"}
),
// -------------------------------------------------------------------
// SECAO 10 - CLASSIFICACAO FDM E SLA (com tratamento de erro)
// -------------------------------------------------------------------
AdicionarClassificacaoFDM = Table.AddColumn(ExpandirDataHoraSeparadas, "ResultadoFDM", each
    try
        fnClassificarFDM(
            [CalculoFDM], [#"Ref. FDM"], [#"Ref. prazo"],
            [#"Descrição da atividade"], [Contrato],
            [Data Abertura], [Hora Abertura],
            [D. fechamento], [Prazo combinado],
            [Município], [UF], [Grupo.1],
            [#"tabelaPrazoContrato.Prazo"], [#"Aplicação"], RefCalendario
        )
    otherwise
        [ Classificacao = null, StatusPrazo = null, Log = "Erro no calculo - verificar dados da linha; " ],
    type record
),
ExpandirClassificacaoFDM = Table.ExpandRecordColumn(
    AdicionarClassificacaoFDM, "ResultadoFDM",
    {"Classificacao", "StatusPrazo", "Log"},
    {"Classificação FDM", "Status Prazo", "LOG SLA"}
),
cod_Tabela_FC = Table.AddColumn(ExpandirClassificacaoFDM, "cod_Tabela_FC", each [Descrição da atividade] & [Item]),
#"Consultas Mescladas" = Table.NestedJoin(cod_Tabela_FC, {"cod_Tabela_FC"}, tabela_FC, {"cod_Tabela_FC"}, "tabela_FC", JoinKind.LeftOuter),
#"tabela_FC Expandido" = Table.ExpandTableColumn(#"Consultas Mescladas", "tabela_FC", {"FC"}, {"FC"}),
#"Valor Substituído" = Table.ReplaceValue(#"tabela_FC Expandido", null, 0, Replacer.ReplaceValue, {"FC"}),
#"Função Personalizada Invocada" = Table.AddColumn(#"Valor Substituído", "QExec", each fnCalcularQExec([Descrição da atividade], [Item], [#"Qtd.#(lf)Atendida"], null, [FC])),
#"Consultas Mescladas1" = Table.NestedJoin(#"Função Personalizada Invocada", {"cod_Tabela_FC"}, tabelaAtividadeItemPPU, {"CodLinhaPPU"}, "tabelaAtividadeItemPPU", JoinKind.LeftOuter),
#"tabelaAtividadeItemPPU Expandido" = Table.ExpandTableColumn(#"Consultas Mescladas1", "tabelaAtividadeItemPPU", {"Item PPU", "Unidade", "Linha de serviço PPU"}, {"Item PPU", "Unidade", "Linha de serviço PPU"}),
#"Personalização Adicionada" = Table.AddColumn(#"tabelaAtividadeItemPPU Expandido", "codOrigemDestino", each [Centro] & [Município] & [Galpão] & [Galpãp_Cidade]),
#"Consultas Mescladas2" = Table.NestedJoin(#"Personalização Adicionada", {"codOrigemDestino"}, Origem_Destino, {"codOrigenDestino"}, "Origem_Destino", JoinKind.LeftOuter),
#"Origem_Destino Expandido" = Table.ExpandTableColumn(#"Consultas Mescladas2", "Origem_Destino", {"KM"}, {"Origem_Destino.KM"}),
#"Personalização Adicionada1" = Table.AddColumn(#"Origem_Destino Expandido", "KM Adicional", each
    if [Origem_Destino.KM] = null then null
    else if [Origem_Destino.KM] <= valKmAdicional then 0
    else [Origem_Destino.KM] - valKmAdicional
),
#"QExec Agrupado Calculado" = fnCalcularQExecAgrupado(
    #"Personalização Adicionada1",
    "Data Fechamento",
    "Localidade",
    "Galpãp_Cidade",
    "Linha de serviço PPU",
    "QExec"
)
in
    #"QExec Agrupado Calculado"
;

shared TabelaFDMPorContrato = let
    // -------------------------------------------------------------------
    // TabelaFDMPorContrato
    // -------------------------------------------------------------------
    // Consolida os indicadores de conformidade de SLA e FDM por contrato.
    //
    // OTIMIZAÇÕES & MELHORIAS APLICADAS:
    //   1. Projeção antecipada de colunas: seleciona apenas as 4 colunas
    //      necessárias de 'Painel T2M', reduzindo drasticamente o consumo
    //      de memória RAM durante as agregações de TSPE/TSPR.
    //   2. Pré-cálculo de parâmetros estáticos fora do pipeline linha a linha:
    //      mês numérico, períodos anteriores e validação histórica são
    //      computados uma única vez no bloco 'let'.
    //   3. Fallback seguro de FDM: caso os períodos anteriores não estejam
    //      disponíveis na base histórica, o FDM final assume o FDM Bruto
    //      calculado para o próprio mês (em vez de resultar em null),
    //      preservando o comportamento fiel do VBA e a continuidade da fatura.
    // -------------------------------------------------------------------

    // 1) Constantes e parâmetros estáticos calculados uma única vez
    MesNumero = fnConverterMesTextoParaNumero(PeriodoMes),
    PeriodoVigenteConst = Text.From(PeriodoAno) & "#" & Text.PadStart(Text.From(MesNumero), 2, "0"),
    PeriodosAnterioresRec = fnCalcularPeriodosAnteriores(MesNumero, PeriodoAno),
    PerAnt1 = PeriodosAnterioresRec[PerAnt1],
    PerAnt2 = PeriodosAnterioresRec[PerAnt2],
    PeriodosValidosConst = fnValidarPeriodoAnterior(PerAnt1) and fnValidarPeriodoAnterior(PerAnt2),

    // 2) Projeção otimizada da base analítica (somente colunas essenciais para FDM)
    BaseParaCalculo = Table.SelectColumns(
        #"Painel T2M",
        {"Contrato", "Situação", "Ref. FDM", "Status Prazo"}
    ),

    // 3) Lista de contratos distintos presentes na base
    ContratosDistintos = Table.Distinct(Table.SelectColumns(BaseParaCalculo, {"Contrato"})),

    // 4) Para cada contrato, calcula TSPE1/TSPR1/TSPE2/TSPR2
    ComTSPE_TSPR = Table.AddColumn(ContratosDistintos, "_TSPE_TSPR", each
        fnCalcularTSPE_TSPR(BaseParaCalculo, [Contrato]),
        type record
    ),
    ExpandirTSPE_TSPR = Table.ExpandRecordColumn(
        ComTSPE_TSPR, "_TSPE_TSPR",
        {"TSPE1", "TSPR1", "TSPE2", "TSPR2"},
        {"TSPE1", "TSPR1", "TSPE2", "TSPR2"}
    ),

    // 5) Calcula os índices IAPFARQ1 e IAPFARQ2 com precisão decimal
    ComIAPFARQ1 = Table.AddColumn(ExpandirTSPE_TSPR, "IAPFARQ1", each
        fnCalcularIAPFARQ([TSPE1], [TSPR1]),
        type number
    ),
    ComIAPFARQ2 = Table.AddColumn(ComIAPFARQ1, "IAPFARQ2", each
        fnCalcularIAPFARQ([TSPE2], [TSPR2]),
        type number
    ),

    // 6) FDM mensal "bruto" (classificação direta do IAPFARQ)
    ComFDM1Bruto = Table.AddColumn(ComIAPFARQ2, "FDM1_Bruto", each
        fnCalcularFDMMensal([IAPFARQ1]),
        type number
    ),
    ComFDM2Bruto = Table.AddColumn(ComFDM1Bruto, "FDM2_Bruto", each
        fnCalcularFDMMensal([IAPFARQ2]),
        type number
    ),

    // 7) Agrega o flag de auditoria de períodos válidos
    ComValidacaoPeriodos = Table.AddColumn(ComFDM2Bruto, "PeriodosValidos", each
        PeriodosValidosConst,
        type logical
    ),

    // 8) Recálculo final do FDM com degradação em cascata (item 8.4.3)
    //    Se não houver histórico válido, assume o FDM Bruto mensal como fallback
    ComFDM1Final = Table.AddColumn(ComValidacaoPeriodos, "FDM1", each
        if not PeriodosValidosConst then
            [FDM1_Bruto]
        else
            try fnRecalcularFDM(PeriodoVigenteConst, PerAnt1, PerAnt2, [Contrato], [FDM1_Bruto], 1)
            otherwise [FDM1_Bruto],
        type number
    ),
    ComFDM2Final = Table.AddColumn(ComFDM1Final, "FDM2", each
        if not PeriodosValidosConst then
            [FDM2_Bruto]
        else
            try fnRecalcularFDM(PeriodoVigenteConst, PerAnt1, PerAnt2, [Contrato], [FDM2_Bruto], 2)
            otherwise [FDM2_Bruto],
        type number
    ),

    // 9) Seleção final e ordenação das colunas
    TabelaFinal = Table.SelectColumns(ComFDM2Final, {
        "Contrato",
        "TSPE1", "TSPR1", "IAPFARQ1", "FDM1_Bruto", "FDM1",
        "TSPE2", "TSPR2", "IAPFARQ2", "FDM2_Bruto", "FDM2",
        "PeriodosValidos"
    })
in
    TabelaFinal
;

shared #"ARM_PA-LT2" = let
    Combinado = Table.Combine({ConsolidadoES, ConsolidadoRJ, ConsolidadoSP, ConsolidadoPA}),

    // Bufferiza o resultado final — evita reabrir os 4 arquivos físicos
    // a cada referência subsequente desta consulta em outras etapas.
    ResultadoFinal = Table.Buffer(Combinado),
    #"Colunas Renomeadas" = Table.RenameColumns(ResultadoFinal,{{"Contrato", "Linha de serviço PPU"}}),
    #"Colunas Removidas" = Table.RemoveColumns(#"Colunas Renomeadas",{"Ø", "A", "L", "P", "cm³ = V1", "cm³ = V2"}),
    #"Tipo Alterado" = Table.TransformColumnTypes(#"Colunas Removidas",{{"FC = V1/V2", type number}}),
    #"Colunas Removidas1" = Table.RemoveColumns(#"Tipo Alterado",{"$ Unitário", "$ Total", "Tabela", "UA"}),
    #"Tipo Alterado1" = Table.TransformColumnTypes(#"Colunas Removidas1",{{"FC = V1/V2", type number}, {"QUAm-1", type number}, {"QBm", type number}, {"QDm", type number}, {"QNm", type number}, {"QUAm", type number}}),
    #"Valor Substituído" = Table.ReplaceValue(#"Tipo Alterado1",null,0,Replacer.ReplaceValue,{"FC = V1/V2", "QUAm-1", "QBm", "QDm", "QNm", "QUAm"}),
    #"Colunas Renomeadas1" = Table.RenameColumns(#"Valor Substituído",{{"FC = V1/V2", "FC"}}),
    #"Colunas Reordenadas" = Table.ReorderColumns(#"Colunas Renomeadas1",{"Contrato-Reg.", "MEDIDAS CONTRATO", "Padrão", "QUAm-1", "QBm", "QDm", "QNm", "QUAm", "FC", "Linha de serviço PPU"}),
    // QUAm-1 - (QBm + QDm) + QNm
    // 
    QUA = Table.AddColumn(#"Colunas Reordenadas", "QUA", each [#"QUAm-1"]-([QBm]+[QDm]) + [QNm]),
    #"Colunas Reordenadas1" = Table.ReorderColumns(QUA,{"Contrato-Reg.", "MEDIDAS CONTRATO", "Padrão", "QUAm-1", "QBm", "QDm", "QNm", "QUAm", "FC", "QUA", "Linha de serviço PPU"}),
    #"Colunas Renomeadas2" = Table.RenameColumns(#"Colunas Reordenadas1",{{"Padrão", "Item"}}),
    #"Linhas Filtradas" = Table.SelectRows(#"Colunas Renomeadas2", each ([#"Contrato-Reg."] = "PA-LT2-BA")),
    #"Colunas Removidas2" = Table.RemoveColumns(#"Linhas Filtradas",{"Contrato-Reg."})
in
    #"Colunas Removidas2";

shared #"DADOS_PA-LT2" = let
    Fonte = DADOS,
    #"Linhas Filtradas" = Table.SelectRows(Fonte, each ([Contrato.1] = "PA-LT2-BA"))
in
    #"Linhas Filtradas";

shared #"FRETE_PA-LT2" = let
    Fonte = DADOS,
    #"Colunas Removidas" = Table.RemoveColumns(Fonte,{"Aplicação", "Código da#(lf)solicitação", "Código OS", "D. abertura", "Prazo combinado", "Qtd.#(lf)Atendida", "Agrupamento", "Classificação", "FDM", "Obs. Isenção", "FA", "FC", "QExec", "Código Centro"}),
    #"Linhas Filtradas" = Table.SelectRows(#"Colunas Removidas", each ([Linha de serviço PPU] = "FRE-EXP" or [Linha de serviço PPU] = "FRE-NRM") and ([#"Localidade - Destino"] = "PA-LT2-BA"))
in
    #"Linhas Filtradas";

shared #"ARM_IRON-LT1" = let
    Combinado = Table.Combine({ConsolidadoES, ConsolidadoRJ, ConsolidadoSP, ConsolidadoPA}),

    // Bufferiza o resultado final — evita reabrir os 4 arquivos físicos
    // a cada referência subsequente desta consulta em outras etapas.
    ResultadoFinal = Table.Buffer(Combinado),
    #"Colunas Renomeadas" = Table.RenameColumns(ResultadoFinal,{{"Contrato", "Linha de serviço PPU"}}),
    #"Colunas Removidas" = Table.RemoveColumns(#"Colunas Renomeadas",{"Ø", "A", "L", "P", "cm³ = V1", "cm³ = V2"}),
    #"Tipo Alterado" = Table.TransformColumnTypes(#"Colunas Removidas",{{"FC = V1/V2", type number}}),
    #"Colunas Removidas1" = Table.RemoveColumns(#"Tipo Alterado",{"$ Unitário", "$ Total", "Tabela", "UA"}),
    #"Tipo Alterado1" = Table.TransformColumnTypes(#"Colunas Removidas1",{{"FC = V1/V2", type number}, {"QUAm-1", type number}, {"QBm", type number}, {"QDm", type number}, {"QNm", type number}, {"QUAm", type number}}),
    #"Valor Substituído" = Table.ReplaceValue(#"Tipo Alterado1",null,0,Replacer.ReplaceValue,{"FC = V1/V2", "QUAm-1", "QBm", "QDm", "QNm", "QUAm"}),
    #"Colunas Renomeadas1" = Table.RenameColumns(#"Valor Substituído",{{"FC = V1/V2", "FC"}}),
    #"Colunas Reordenadas" = Table.ReorderColumns(#"Colunas Renomeadas1",{"Contrato-Reg.", "MEDIDAS CONTRATO", "Padrão", "QUAm-1", "QBm", "QDm", "QNm", "QUAm", "FC", "Linha de serviço PPU"}),
    // QUAm-1 - (QBm + QDm) + QNm
    // 
    QUA = Table.AddColumn(#"Colunas Reordenadas", "QUA", each [#"QUAm-1"]-([QBm]+[QDm]) + [QNm]),
    #"Colunas Reordenadas1" = Table.ReorderColumns(QUA,{"Contrato-Reg.", "MEDIDAS CONTRATO", "Padrão", "QUAm-1", "QBm", "QDm", "QNm", "QUAm", "FC", "QUA", "Linha de serviço PPU"}),
    #"Colunas Renomeadas2" = Table.RenameColumns(#"Colunas Reordenadas1",{{"Padrão", "Item"}}),
    #"Linhas Filtradas" = Table.SelectRows(#"Colunas Renomeadas2", each ([#"Contrato-Reg."] <> "PA-LT2-BA"))
in
    #"Linhas Filtradas";

shared #"DADOS_IRON-LT1" = let
    Fonte = DADOS,
    #"Linhas Filtradas" = Table.SelectRows(Fonte, each ([Contrato.1] <> "PA-LT2-BA"))
in
    #"Linhas Filtradas";

shared #"FRETE_IRON-LT1" = let
    Fonte = DADOS,
    #"Colunas Removidas" = Table.RemoveColumns(Fonte,{"Aplicação", "Código da#(lf)solicitação", "Código OS", "D. abertura", "Prazo combinado", "Qtd.#(lf)Atendida", "Agrupamento", "Classificação", "FDM", "Obs. Isenção", "FA", "FC", "QExec", "Código Centro"}),
    #"Linhas Filtradas" = Table.SelectRows(#"Colunas Removidas", each ([Linha de serviço PPU] = "FRE-EXP" or [Linha de serviço PPU] = "FRE-NRM") and ([Contrato.1] <> "PA-LT2-BA"))
in
    #"Linhas Filtradas";

shared #"PPU-PA-LT2" = let
    Fonte = Excel.Workbook(File.Contents(localTabelaB), null, true),
    PPU_Table = Fonte{[Item="PPU",Kind="Table"]}[Data],
    #"Tipo Alterado" = Table.TransformColumnTypes(PPU_Table,{{"Contrato", type text}, {"Item PPU", type text}, {"Linha de Serviço PPU", type text}, {"Descrição da Linha de Serviço", type text}, {"Unidade da Linha de Serviço", type text}, {"P. Unitário (R$)", type number}}),
    #"Linhas Filtradas" = Table.SelectRows(#"Tipo Alterado", each ([Contrato] = "PA-LT2")),
    #"Tipo Alterado1" = Table.TransformColumnTypes(#"Linhas Filtradas",{{"P. Unitário (R$)", Currency.Type}}),
    #"Colunas Renomeadas" = Table.RenameColumns(#"Tipo Alterado1",{{"Unidade da Linha de Serviço", "Unidade"}, {"Linha de Serviço PPU", "Linha de Serviço"}})
in
    #"Colunas Renomeadas";

shared #"PPU-IRON-LT1" = let
    Fonte = Excel.Workbook(File.Contents(localTabelaB), null, true),
    PPU_Table = Fonte{[Item="PPU",Kind="Table"]}[Data],
    #"Tipo Alterado" = Table.TransformColumnTypes(PPU_Table,{{"Contrato", type text}, {"Item PPU", type text}, {"Linha de Serviço PPU", type text}, {"Descrição da Linha de Serviço", type text}, {"Unidade da Linha de Serviço", type text}, {"P. Unitário (R$)", type number}}),
    #"Linhas Filtradas" = Table.SelectRows(#"Tipo Alterado", each ([Contrato] = "IRON-LT1")),
    #"Tipo Alterado1" = Table.TransformColumnTypes(#"Linhas Filtradas",{{"P. Unitário (R$)", Currency.Type}}),
    #"Colunas Renomeadas" = Table.RenameColumns(#"Tipo Alterado1",{{"Linha de Serviço PPU", "Linha de Serviço"}, {"Unidade da Linha de Serviço", "Unidade"}})
in
    #"Colunas Renomeadas";

shared #"ARM_IRON-LT1-RJ" = let
    Fonte = ARM,
    #"Linhas Filtradas" = Table.SelectRows(Fonte, each ([#"Contrato-Reg."] = "IRON-LT1-RJ")),
    #"Colunas Removidas" = Table.RemoveColumns(#"Linhas Filtradas",{"Contrato-Reg."})
in
    #"Colunas Removidas";

shared #"DADOS_IRON-LT1-RJ" = let
    Fonte = DADOS,
    #"Linhas Filtradas" = Table.SelectRows(Fonte, each ([Contrato.1] = "IRON-LT1-RJ"))
in
    #"Linhas Filtradas";

shared #"FRETE_IRON-LT1-RJ" = let
    Fonte = DADOS,
    #"Colunas Removidas" = Table.RemoveColumns(Fonte,{"Aplicação", "Código da#(lf)solicitação", "Código OS", "D. abertura", "Prazo combinado", "Qtd.#(lf)Atendida", "Agrupamento", "Classificação", "FDM", "Obs. Isenção", "FA", "FC", "QExec", "Código Centro"}),
    #"Linhas Filtradas" = Table.SelectRows(#"Colunas Removidas", each ([Linha de serviço PPU] = "FRE-EXP" or [Linha de serviço PPU] = "FRE-NRM") and ([Contrato.1] = "IRON-LT1-RJ"))
in
    #"Linhas Filtradas";

shared #"ARM_IRON-LT1-SP" = let
    Fonte = ARM,
    #"Linhas Filtradas" = Table.SelectRows(Fonte, each ([#"Contrato-Reg."] = "IRON-LT1-SP"))
in
    #"Linhas Filtradas";

shared #"DADOS_IRON-LT1-SP" = let
    Fonte = DADOS,
    #"Linhas Filtradas" = Table.SelectRows(Fonte, each ([Contrato.1] = "IRON-LT1-SP"))
in
    #"Linhas Filtradas";

shared #"FRETE_IRON-LT1-SP" = let
    Fonte = FRETE,
    #"Linhas Filtradas" = Table.SelectRows(Fonte, each ([Contrato.1] = "IRON-LT1-SP"))
in
    #"Linhas Filtradas";

shared #"ARM_IRON-LT1-ES" = let
    Fonte = ARM,
    #"Linhas Filtradas" = Table.SelectRows(Fonte, each ([#"Contrato-Reg."] = "IRON-LT1-ES"))
in
    #"Linhas Filtradas";

shared #"DADOS_IRON-LT1-ES" = let
    Fonte = DADOS,
    #"Linhas Filtradas" = Table.SelectRows(Fonte, each ([Contrato.1] = "IRON-LT1-ES"))
in
    #"Linhas Filtradas";

shared #"FRETE_IRON-LT1-ES" = let
    Fonte = FRETE,
    #"Linhas Filtradas" = Table.SelectRows(Fonte, each ([Contrato.1] = "IRON-LT1-ES"))
in
    #"Linhas Filtradas";

shared Calendario = let
       Fonte = Excel.Workbook(File.Contents(localTabelaB), null, true),
    CalendarioFeriado_Table = Fonte{[Item="CalendarioFeriado",Kind="Table"]}[Data],
    #"Tipo Alterado" = Table.TransformColumnTypes(CalendarioFeriado_Table,{{"Data", type date}, {"Feriado", type text}, {"Abrangência", type text}, {"Tipo", type text}, {"UF", type text}})
in
    #"Tipo Alterado";