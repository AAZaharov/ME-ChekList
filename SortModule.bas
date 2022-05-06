Attribute VB_Name = "SortModule"
' сортировка по двум столбцам - вначале по RelRecNr,
' затем по Rework
Sub sortDataBase()

    ' перва€ и последн€€ строка с данными
    Dim firstDataRow, lastDataRow As Integer
        firstDataRow = Sheet_DataBase.ListObjects("DataTable").DataBodyRange.Row
        lastDataRow = Sheet_DataBase.Cells(Rows.Count, "B").End(xlUp).Row
    
    Sheet_DataBase.Sort.SortFields.clear
    ' добавл€ем сортировку по столбцу RelRecNr
    Sheet_DataBase.Sort.SortFields.Add Key:=Range("B" & firstDataRow & ":B" & lastDataRow), _
        SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortTextAsNumbers  ' сортируем как числа
    
    ' добавл€ем сортировку по столбцу Rework
    Sheet_DataBase.Sort.SortFields.Add Key:=Range("F" & firstDataRow & ":F" & lastDataRow) _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With Sheet_DataBase.Sort
        .SetRange Sheet_DataBase.ListObjects("DataTable").Range
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
End Sub


