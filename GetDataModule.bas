Attribute VB_Name = "GetDataModule"
' поиск RelRecNr в базе данных
Function getRelRecNr() As Integer

    ' если в базе ещё нет записей, возвращаем 0 и выходим из функции
    If Sheet_DataBase.Cells(3, "B") = "" Then
        getRelRecNr = 0
        Exit Function
    End If
    
    ' массив для копирования столбца RelRecNr с листа DataBase
    Dim rrNumbers() As Variant
    ' последняя заполненная строка
    Dim lastRrnRow As Integer
    lastRrnRow = Sheet_DataBase.Cells(Rows.Count, "B").End(xlUp).Row
    
    Dim rrnR As Range
    Set rrnR = Range(Sheet_DataBase.Cells(3, "A"), Sheet_DataBase.Cells(lastRrnRow, "B"))
    rrNumbers = rrnR.Value
    
    ' RelRecNr
    Dim rrn As String
    rrn = Sheet_IP_Check.Cells(2, "F").Value
    
    ' ищем указанный RelRecNr в массиве и возвращаем номер строки
    For i = 3 To lastRrnRow
        If rrNumbers(i - 2, 2) = rrn Then
            getRelRecNr = i
        End If
    Next i
    
End Function

Function getFinishedRRN()

    ' массив для копирования столбца RelRecNr с листа DataBase
    Dim rrNumbers() As Variant
    ' последняя заполненная строка
    Dim lastRrnRow As Integer
    lastRrnRow = Sheet_DataBase.Cells(Rows.Count, "B").End(xlUp).Row
    
    Dim rrnR As Range
    Set rrnR = Range(Sheet_DataBase.Cells(3, "B"), Sheet_DataBase.Cells(lastRrnRow, "F"))
    rrNumbers = rrnR.Value
    
    ' RelRecNr
    Dim rrn As String
    rrn = Sheet_IP_Check.Cells(2, "F").Value
    ' Rework
    Dim rework As String
    rework = "FINISHED"
    
    ' ищем указанный RelRecNr в массиве и возвращаем номер строки
    For i = 3 To lastRrnRow
        If rrNumbers(i - 2, 1) = rrn And rrNumbers(i - 2, 5) = rework Then
            getFinishedRRN = i
        End If
    Next i
    
End Function

Function getEqualRework()

    ' массив для копирования столбца RelRecNr с листа DataBase
    Dim rrNumbers() As Variant
    ' последняя заполненная строка
    Dim lastRrnRow As Integer
    lastRrnRow = Sheet_DataBase.Cells(Rows.Count, "B").End(xlUp).Row
    
    Dim rrnR As Range
    Set rrnR = Range(Sheet_DataBase.Cells(3, "B"), Sheet_DataBase.Cells(lastRrnRow, "F"))
    rrNumbers = rrnR.Value
    
    ' RelRecNr
    Dim rrn As String
    rrn = Sheet_IP_Check.Cells(2, "F").Value
    ' Rework
    Dim rework As String
    rework = Sheet_IP_Check.reworkComboBox.Value
    
    ' ищем указанный RelRecNr в массиве и возвращаем номер строки
    For i = 3 To lastRrnRow
        If rrNumbers(i - 2, 1) = rrn And rrNumbers(i - 2, 5) = rework Then
            getEqualRework = i
        End If
    Next i
    
End Function

Function getLastRework()

    ' массив для копирования столбца RelRecNr с листа DataBase
    Dim rrNumbers() As Variant
    ' последняя заполненная строка
    Dim lastRrnRow As Integer
    lastRrnRow = Sheet_DataBase.Cells(Rows.Count, "B").End(xlUp).Row
    
    Dim rrnR As Range
    Set rrnR = Range(Sheet_DataBase.Cells(3, "B"), Sheet_DataBase.Cells(lastRrnRow, "F"))
    rrNumbers = rrnR.Value
    
    ' RelRecNr
    Dim rrn As String
    rrn = Sheet_IP_Check.Cells(2, "F").Value
    ' IP Number
    Dim ipn As String
    ipn = Sheet_IP_Check.Cells(4, "F").Value
    ' Rework
    Dim lastRework As Integer
    lastRework = Sheet_IP_Check.reworkComboBox.Value
    
    getLastRework = lastRework
    
    ' проверяем записи с одинаковыми RelRecNr и IP Number,
    ' находим и возвращаем последний номер Rework
    For i = 3 To lastRrnRow
    
        If rrNumbers(i - 2, 1) = rrn And rrNumbers(i - 2, 3) = ipn _
            And rrNumbers(i - 2, 5) > lastRework Then
            
            lastRework = rrNumbers(i - 2, 5)
            
        End If
        
    Next i
    
    getLastRework = lastRework
    
End Function

Function getAllReworks() As Collection

    ' массив для копирования столбца RelRecNr с листа DataBase
    Dim rrNumbers() As Variant
    ' последняя заполненная строка
    Dim lastRrnRow As Integer
    lastRrnRow = Sheet_DataBase.Cells(Rows.Count, "B").End(xlUp).Row
    
    Dim rrnR As Range
    Set rrnR = Range(Sheet_DataBase.Cells(3, "B"), Sheet_DataBase.Cells(lastRrnRow, "F"))
    rrNumbers = rrnR.Value
    
    ' коллекция для хранения номеров Rework для выбранных RelRecNr и IP Number
    Dim reworkColl As Collection
    Set reworkColl = New Collection
    
    ' RelRecNr
    Dim rrn As String
    rrn = Sheet_IP_Check.Cells(2, "F").Value
    ' IP Number
    Dim ipn As String
    ipn = Sheet_IP_Check.Cells(4, "F").Value
    ' Rework
    Dim lastRework As String
    lastRework = Sheet_IP_Check.reworkComboBox.Value
    
    ' проверяем записи с одинаковыми RelRecNr и IP Number,
    ' записываем имеющиеся номера Rework в коллекцию
    For i = 3 To lastRrnRow
    
        If rrNumbers(i - 2, 1) = rrn And rrNumbers(i - 2, 3) = ipn Then
        
            reworkColl.Add (rrNumbers(i - 2, 5))
            
        End If
        
    Next i
    
    Set getAllReworks = reworkColl
    
End Function

Function getPerformer() As Integer

    ' поиск исполнителя в базе данных
    ' массив для копирования исполнителей с листа Send_Email
    Dim performers() As Variant
    ' последняя заполненная строка в списке исполнителей
    Dim lastPerfRow As Integer
    lastPerfRow = Sheet_SendEmail.Cells(Rows.Count, "A").End(xlUp).Row
    
    Dim pr As Range
    Set pr = Range(Sheet_SendEmail.Cells(1, "A"), Sheet_SendEmail.Cells(lastPerfRow, "B"))
    performers = pr.Value
    
    ' исполнитель
    Dim performer As String
    performer = Sheet_IP_Check.performerComboBox.Value
    
    ' ищем указанного исполнителя в массиве и запоминаем номер строки
    For i = 1 To lastPerfRow
        If performers(i, 1) = performer Then
            getPerformer = i
        End If
    Next i

End Function

Function getUpdatedRow() As Integer
    
    ' массив для копирования столбца RelRecNr с листа DataBase
    Dim rrNumbers() As Variant
    ' последняя заполненная строка
    Dim lastRrnRow As Integer
    lastRrnRow = Sheet_DataBase.Cells(Rows.Count, "B").End(xlUp).Row
    
    Dim rrnR As Range
    Set rrnR = Range(Sheet_DataBase.Cells(3, "B"), Sheet_DataBase.Cells(lastRrnRow, "F"))
    rrNumbers = rrnR.Value
    
    ' RelRecNr
    Dim rrn As String
    rrn = Sheet_IP_Check.Cells(2, "F").Value
    ' Rework
    Dim rework As String
    rework = Sheet_IP_Check.reworkComboBox.Value
    ' IP Number
    Dim ipNum As String
    ipNum = Sheet_IP_Check.Cells(4, "F").Value
    
    
    ' ищем подходящую строку в массиве и возвращаем номер строки
    For i = 3 To lastRrnRow
        If rrNumbers(i - 2, 1) = rrn _
            And rrNumbers(i - 2, 3) = ipNum _
            And rrNumbers(i - 2, 5) = rework Then
            
            getUpdatedRow = i
            
        End If
    Next i
    
End Function

Function getUpdatedDescrRows() As Collection
    
    ' коллекция для записи номеров строк с описанием ошибок
    Dim errDescrColl As Collection
    Set errDescrColl = New Collection
    
    ' массив для копирования столбца RelRecNr с листа PERFORMER
    Dim rrNumbers() As Variant
    ' последняя заполненная строка
    Dim lastRrnRow As Integer
    lastRrnRow = Sheet_ErrDescr.Cells(Rows.Count, "B").End(xlUp).Row
    
    Dim rrnR As Range
    Set rrnR = Range(Sheet_ErrDescr.Cells(3, "B"), Sheet_ErrDescr.Cells(lastRrnRow, "F"))
    rrNumbers = rrnR.Value
    
    ' RelRecNr
    Dim rrn As String
    rrn = Sheet_IP_Check.Cells(2, "F").Value
    ' Rework
    Dim rework As String
    rework = Sheet_IP_Check.reworkComboBox.Value
    ' IP Number
    Dim ipNum As String
    ipNum = Sheet_IP_Check.Cells(4, "F").Value
    
    
    ' ищем подходящую строку в массиве и записываем в коллекцию
    For i = 3 To lastRrnRow
        If rrNumbers(i - 2, 1) = rrn _
            And rrNumbers(i - 2, 3) = ipNum _
            And rrNumbers(i - 2, 5) = rework Then
            
            errDescrColl.Add (i)
            
        End If
    Next i
    
    Set getUpdatedDescrRows = errDescrColl
    
End Function

Function getRow()
    
    ' последняя заполненная строка в базе, определяем по столбцу "B" (RelRecNr) на листе DataBase
    Dim lastDbRow As Integer
    lastDbRow = Sheet_DataBase.Cells(Rows.Count, "B").End(xlUp).Row
    
    ' если в базе ещё нет записей, возвращаем 3,
    ' иначе возвращаем первую пустую строку
    If Sheet_DataBase.Cells(3, "B") = "" Then
        getRow = 3
    Else
        getRow = lastDbRow + 1
    End If
    
End Function

Function getDescrRow()
    
    ' последняя заполненная строка в базе описания ошибок, определяем по столбцу "B" (RelRecNr) на листе PERFORMER
    Dim lastDescrDbRow As Integer
    lastDescrDbRow = Sheet_ErrDescr.Cells(Rows.Count, "B").End(xlUp).Row
    
    ' если в базе ещё нет записей, возвращаем 3,
    ' иначе возвращаем первую пустую строку
    If Sheet_ErrDescr.Cells(3, "B") = "" Then
        getDescrRow = 3
    Else
        getDescrRow = lastDescrDbRow + 1
    End If
    
End Function

Function getSumIpErrors(rowNum As Integer)
    getSumIpErrors = Sheet_DataBase.Cells(rowNum, "BS")
End Function

Function getSumPdmErrors(rowNum As Integer)
    getSumPdmErrors = Sheet_DataBase.Cells(rowNum, "BT")
End Function

Sub initIpComboBoxes()

    ' fill performer ComboBox
    Sheet_IP_Check.performerComboBox.ColumnCount = 1
    lr = Sheet_SendEmail.Cells(Rows.Count, 1).End(xlUp).Row
    Sheet_IP_Check.performerComboBox.List = Sheet_SendEmail.Range("A1:A" & lr).Value
    
    ' fill rework ComboBox
    Sheet_IP_Check.reworkComboBox.ColumnCount = 1
    lr = Sheet_SendEmail.Cells(Rows.Count, 3).End(xlUp).Row
    Sheet_IP_Check.reworkComboBox.List = Sheet_SendEmail.Range("C1:C" & lr).Value
    
    ' fill MESA status ComboBox
    Sheet_IP_Check.mesaStatusComboBox.ColumnCount = 1
    lr = Sheet_SendEmail.Cells(Rows.Count, 4).End(xlUp).Row
    Sheet_IP_Check.mesaStatusComboBox.List = Sheet_SendEmail.Range("D1:D" & lr).Value

End Sub

Sub initPdmComboBoxes()

    ' fill performer ComboBox
    Sheet_PDM_Check.performerComboBox.ColumnCount = 1
    lr = Sheet_SendEmail.Cells(Rows.Count, 1).End(xlUp).Row
    Sheet_PDM_Check.performerComboBox.List = Sheet_SendEmail.Range("A1:A" & lr).Value
    
    ' fill rework ComboBox
    Sheet_PDM_Check.reworkComboBox.ColumnCount = 1
    lr = Sheet_SendEmail.Cells(Rows.Count, 3).End(xlUp).Row
    Sheet_PDM_Check.reworkComboBox.List = Sheet_SendEmail.Range("C1:C" & lr).Value
    
    ' fill MESA status ComboBox
    Sheet_PDM_Check.mesaStatusComboBox.ColumnCount = 1
    lr = Sheet_SendEmail.Cells(Rows.Count, 4).End(xlUp).Row
    Sheet_PDM_Check.mesaStatusComboBox.List = Sheet_SendEmail.Range("D1:D" & lr).Value

End Sub

