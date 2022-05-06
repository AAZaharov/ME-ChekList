Attribute VB_Name = "SortModule"
' ���������� �� ���� �������� - ������� �� RelRecNr,
' ����� �� Rework
Sub sortDataBase()

    ' ������ � ��������� ������ � �������
    Dim firstDataRow, lastDataRow As Integer
        firstDataRow = Sheet_DataBase.ListObjects("DataTable").DataBodyRange.Row
        lastDataRow = Sheet_DataBase.Cells(Rows.Count, "B").End(xlUp).Row
    
    Sheet_DataBase.Sort.SortFields.clear
    ' ��������� ���������� �� ������� RelRecNr
    Sheet_DataBase.Sort.SortFields.Add Key:=Range("B" & firstDataRow & ":B" & lastDataRow), _
        SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortTextAsNumbers  ' ��������� ��� �����
    
    ' ��������� ���������� �� ������� Rework
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


