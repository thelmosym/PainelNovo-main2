let
    Fonte = Excel.Workbook(File.Contents(localTabelaB), null, true),
    Calendario_Table = Fonte{[Item="Calendario",Kind="Table"]}[Data],
    ColunasAtuais = Table.ColumnNames(Calendario_Table),
    
    // Assegura existencia de Descricao do feriado
    ComDescricao = if List.Contains(ColunasAtuais, "Descrição do feriado") then Calendario_Table
                   else if List.Contains(ColunasAtuais, "Feriado") then Table.AddColumn(Calendario_Table, "Descrição do feriado", each [Feriado], type text)
                   else Table.AddColumn(Calendario_Table, "Descrição do feriado", each "", type text),

    // Assegura existencia de Tipo de feriado
    ComTipo = if List.Contains(Table.ColumnNames(ComDescricao), "Tipo de feriado") then ComDescricao
              else if List.Contains(Table.ColumnNames(ComDescricao), "Tipo") then Table.AddColumn(ComDescricao, "Tipo de feriado", each [Tipo], type text)
              else Table.AddColumn(ComDescricao, "Tipo de feriado", each "", type text),

    // Assegura existencia de Municipio
    ComMunicipio = if List.Contains(Table.ColumnNames(ComTipo), "Município") then ComTipo
                   else if List.Contains(Table.ColumnNames(ComTipo), "Municipio") then Table.RenameColumns(ComTipo, {{"Municipio", "Município"}})
                   else Table.AddColumn(ComTipo, "Município", each "Todos", type text),

    // Converte Data para type date (assegura compatibilidade sem hora)
    DataDate = Table.TransformColumns(ComMunicipio, {{"Data", each Date.From(_), type date}}),
    #"Tipo Alterado" = Table.TransformColumnTypes(DataDate, {
        {"Data", type date},
        {"Descrição do feriado", type text},
        {"Tipo de feriado", type text},
        {"Município", type text},
        {"UF", type text}
    })
in
    #"Tipo Alterado"
