# Vlookup-function-in-Power-Query

The VLOOKUP function is widely recognized in Excel for searching values in a table using both approximate and exact match criteria. In Power Query, the "Merge" function can replicate VLOOKUP's exact match capability. However, there isn't a direct equivalent for approximate matches. Typically, users combine several functions depending on their specific needs. Below, I've provided a custom function to emulate the exact functionality of VLOOKUP in Power Query.



```powerquery-m
(lookup_value, table_array as table, col_index_num as number, optional range_lookup) =>
  let
    Column_names = Table.ColumnNames(table_array), 
    Serch_column = Table.Column(table_array, Column_names{0}), 
    Serch_column2 = 
      if lookup_value is text then
        List.Transform(Serch_column, Text.Lower)
      else
        Serch_column, 
    search_value = if lookup_value is text then Text.Lower(lookup_value) else lookup_value, 
    search_value2 = 
      if List.ContainsAny({range_lookup}, {null, "false", 0}) then
        search_value
      else
        List.Last(List.FirstN(Serch_column2, each _ <= search_value)), 
    Result_column = Table.Column(table_array, Column_names{col_index_num - 1}), 
    Resultindex = List.PositionOf(Serch_column2, search_value2), 
    Result = try Result_column{Resultindex} otherwise "#N/A"
  in
    Result

```
