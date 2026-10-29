let
    Fonte = Excel.Workbook(File.Contents(localTabelaB), null, true),
    Calendario_Table = Fonte{[Item="Calendario",Kind="Table"]}[Data],
    ColunasAtuais = Table.ColumnNames(Calendario_Table),
    // Assegura existencia de Municipio
    ComMunicipio = if List.Contains(ColunasAtuais, "Município") then Calendario_Table
                   else if List.Contains(ColunasAtuais, "Municipio") then Table.RenameColumns(Calendario_Table, {{"Municipio", "Município"}})
                   else Table.AddColumn(Calendario_Table, "Município", each "Todos", type text),
    // Assegura existencia de Descricao do feriado
    ComDescricao = if List.Contains(Table.ColumnNames(ComMunicipio), "Descrição do feriado") then ComMunicipio
                   else if List.Contains(Table.ColumnNames(ComMunicipio), "Feriado") then Table.DuplicateColumn(ComMunicipio, "Feriado", "Descrição do feriado")
                   else Table.AddColumn(ComMunicipio, "Descrição do feriado", each "", type text),
    // Assegura existencia de Tipo de feriado
    ComTipo = if List.Contains(Table.ColumnNames(ComDescricao), "Tipo de feriado") then ComDescricao
              else if List.Contains(Table.ColumnNames(ComDescricao), "Tipo") then Table.DuplicateColumn(ComDescricao, "Tipo", "Tipo de feriado")
              else Table.AddColumn(ComDescricao, "Tipo de feriado", each "", type text),
    #"Tipo Alterado" = Table.TransformColumnTypes(ComTipo,{{"Data", type date}, {"Descrição do feriado", type text}, {"Tipo de feriado", type text}, {"Município", type text}, {"UF", type text}})
in
    #"Tipo Alterado"
