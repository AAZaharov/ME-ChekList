Attribute VB_Name = "SaveModule"
' номер строки для записи на лист DataBase
Dim dbSaveRowNumber As Integer

' номер обновляемой строки для записи на лист DataBase
Dim dbUpdatedRowNumber As Integer

Sub validateAndSave()
    
    'отключаем все обработчики событий
    Application.EnableEvents = False
    
    Dim saveResult As Boolean
    
    ' выполняем проверку
    ' и записываем результат в логическую переменную
    Dim isValidRecord As Boolean
    isValidRecord = ValidationModule.beforeSaveValidation
    
    Dim mailReason As String
    mailReason = "Статус проверки: "
    
    ' номер строки с новой записью в DataBase
    Dim checkRowNum As Integer
    
    ' если проверка пройдена, выполняем запись/исправление и отправляем письмо
    If isValidRecord Then
        If Sheet_IP_Check.saveRecordToggleButton.Value = True Then
            checkRowNum = updateCheck()
            mailReason = mailReason & "исправление"
        Else
            checkRowNum = saveCheck()
            mailReason = mailReason & "новая проверка"
        End If
        
        ' записываем в поле EMAIL STATUS значение "No", если письмо будет отправлено, значение будет изменено
        Sheet_DataBase.Cells(checkRowNum, "BQ").Value = "No"
        
        ' если выбраны опции, делаем проверку на наличие ошибок
        ' и отправляем соответствующее письмо
        
        ' если статус проверки Task Status = "Completed"
        If Sheet_DataBase.Cells(checkRowNum, "BR") = "Completed" Then
            ' если выбрана опция "отправлять письмо об отсутствии ошибок"
            If Sheet_IP_Check.sendFinishedStateCheckBox.Value Then
            
                Call sendFinishedMail(checkRowNum)
                
            End If
        Else
            ' если выбрана опция "отправлять письмо с описанием ошибок"
            If Sheet_IP_Check.sendErrDescrCheckBox.Value Then
                
                Call sendMail(checkRowNum)
                
            End If
        End If
        
        ' сортируем базу
        Call SortModule.sortDataBase
        
        Sheet_IP_Check.Activate
        
        ' сохраняем файл
        ActiveWorkbook.Save
        
        ' показываем форму с информацией о последнем отчёте
        ReportInfoForm.Show
        
    End If
    
    
    'включаем все обработчики событий
    Application.EnableEvents = True
    '---------------------------------------------
    
End Sub

Function saveCheck()

    ' определяем номер строки для сохранения
    dbSaveRowNumber = getRow()
        
    ' записываем атрибуты проверки
    Call saveAttributes(dbSaveRowNumber)
    
    ' сохраняем ошибки в секции IP
    Call saveIpQuestions(dbSaveRowNumber)
    
    ' сохраняем ошибки в секции PDM
    Call savePdmQuestions(dbSaveRowNumber)
    
    ' сохраняем поле TASK STATUS
    Call saveTaskStatus(dbSaveRowNumber)
    
    ' сохр описание ошибок
    Call saveDescr
    
    ' возвращаем номер строки с записью
    saveCheck = dbSaveRowNumber
    
End Function

Function updateCheck()
    
    ' определяем номер строки записи, которую меняем
    dbUpdatedRowNumber = getUpdatedRow()
    
    ' удаляем старую запись
    Call SaveModule.deleteExistingCheck(dbUpdatedRowNumber)
    
    ' определяем номер строки для сохранения
    dbSaveRowNumber = getRow()

    ' записываем атрибуты проверки
    Call saveAttributes(dbSaveRowNumber)

    ' сохраняем ошибки в секции IP
    Call saveIpQuestions(dbSaveRowNumber)

    ' сохраняем ошибки в секции PDM
    Call savePdmQuestions(dbSaveRowNumber)

    ' сохраняем поле TASK STATUS
    Call saveTaskStatus(dbSaveRowNumber)
    
    ' находим номера строк с описанием ошибок,
    ' которые нужно заменить
    Dim errDescrRowsColl As Collection
    Set errDescrRowsColl = getUpdatedDescrRows()
    
    ' удаляем строки с описанием ошибок
    ' удаление идёт от нижней строки к верхней
    For i = errDescrRowsColl.Count To 1 Step -1
        deleteExistingDescription (errDescrRowsColl(i))
    Next i
    
    ' сохраняем новые строки с описанием ошибок
    Call saveDescr
    
    ' возвращаем номер строки с записью
    updateCheck = dbSaveRowNumber
    
End Function

Sub saveAttributes(rowNumber As Integer)

    ' сохраняем дату
    Sheet_DataBase.Cells(rowNumber, "A").Value = Sheet_IP_Check.Cells(1, "F")
    ' сохраняем RelRecNr
    Sheet_DataBase.Cells(rowNumber, "B").Value = Sheet_IP_Check.Cells(2, "F")
    ' сохраняем исполнителя
    Sheet_DataBase.Cells(rowNumber, "C").Value = Sheet_IP_Check.performerComboBox.Value
    ' сохраняем IP Number
    Sheet_DataBase.Cells(rowNumber, "D").Value = Sheet_IP_Check.Cells(4, "F")
    ' сохраняем номер модуля
    Sheet_DataBase.Cells(rowNumber, "E").Value = Sheet_IP_Check.Cells(5, "F")
    ' сохраняем Rework
    Sheet_DataBase.Cells(rowNumber, "F").Value = Sheet_IP_Check.reworkComboBox.Value
    ' сохраняем MESA status
    Sheet_DataBase.Cells(rowNumber, "G").Value = Sheet_IP_Check.mesaStatusComboBox.Value
    
End Sub

Sub saveIpQuestions(rowNumber As Integer)
    
    ' массив для копирования секции вопросов на листе Checklist
    Dim questions() As Variant
    ' последняя заполненная строка, определяем по столбцу "B" с текстом вопроса
    Dim lastQRow As Integer
    lastQRow = Sheet_IP_Check.Cells(Rows.Count, "B").End(xlUp).Row
    
    Dim qR As Range
    Set qR = Range(Sheet_IP_Check.Cells(3, "A"), Sheet_IP_Check.Cells(lastQRow, "C"))
    questions = qR.Value
    
    ' коллекция для копирования строки заголовков вопросов на листе DataBase
    Dim qHeader As Collection
    Set qHeader = New Collection
    ' последний заполненный столбец заголовка - вторая строка листа DataBase
    Dim lastHeadCol As Integer
    lastHeadCol = Sheet_DataBase.Cells(2, Columns.Count).End(xlToLeft).Column
    Dim itemOfColl(1 To 2) As Variant
    ' заполняем коллекцию массивами - в первом элементе каждого массива будет значение ячейки,
    ' во втором - номер столбца этой ячейки
    ' ключ - текст в ячейке
    For i = 8 To lastHeadCol
        itemOfColl(1) = Sheet_DataBase.Cells(2, i).Value
        itemOfColl(2) = Sheet_DataBase.Cells(2, i).Column
        qHeader.Add Item:=itemOfColl, Key:=CStr(Sheet_DataBase.Cells(2, i))
    Next i
    
    ' суммарное количество ошибок в секции IP
    Dim sumIpErrors As Integer
    sumIpErrors = 0
    
    For k = 1 To UBound(questions, 1)
        ' если в третьем столбце единица, это отмеченный вопрос,
        ' поэтому записываем его в базу
        If questions(k, 3) = 1 Then
            ' прибавляем к сумме ошибок
            sumIpErrors = sumIpErrors + 1
            ' определяем номер столбца нужного вопроса
            colNumber = qHeader.Item(CStr(Sheet_IP_Check.Cells(k + 2, 1)))(2)
            ' записываем ошибку в базу
            Sheet_DataBase.Cells(rowNumber, colNumber).Value = 1
        End If
        
    Next k
    
    ' записываем суммарное количество ошибок в секции IP
    Sheet_DataBase.Cells(rowNumber, qHeader.Item("IP_SUMM")(2)).Value = sumIpErrors
    
End Sub

Sub savePdmQuestions(rowNumber As Integer)
    
    ' массив для копирования секции вопросов на листе PDM_Checklist
    Dim pdmQuestions() As Variant
    ' последняя заполненная строка, определяем по столбцу "B" с номером вопроса
    Dim lastPdmpdmQRow As Integer
    lastPdmpdmQRow = Sheet_PDM_Check.Cells(Rows.Count, "B").End(xlUp).Row
    
    Dim pdmQR As Range
    Set pdmQR = Range(Sheet_PDM_Check.Cells(2, "B"), Sheet_PDM_Check.Cells(lastPdmpdmQRow, "D"))
    pdmQuestions = pdmQR.Value
    
    ' коллекция для копирования строки заголовков вопросов на листе DataBase
    Dim qHeader As Collection
    Set qHeader = New Collection
    ' последний заполненный столбец заголовка - вторая строка листа DataBase
    Dim lastHeadCol As Integer
    lastHeadCol = Sheet_DataBase.Cells(2, Columns.Count).End(xlToLeft).Column
    Dim itemOfColl(1 To 2) As Variant
    ' заполняем коллекцию массивами - в первом элементе каждого массива будет значение ячейки,
    ' во втором - номер столбца этой ячейки
    ' ключ - текст в ячейке
    For i = 8 To lastHeadCol
        itemOfColl(1) = Sheet_DataBase.Cells(2, i).Value
        itemOfColl(2) = Sheet_DataBase.Cells(2, i).Column
        qHeader.Add Item:=itemOfColl, Key:=CStr(Sheet_DataBase.Cells(2, i))
    Next i
    
    ' суммарное количество ошибок в секции PDM
    Dim sumPdmErrors As Integer
    sumPdmErrors = 0
    
    For k = 1 To UBound(pdmQuestions, 1)
        ' если в четвёртом столбце листа (третий столбец массива) единица,
        ' это отмеченный вопрос, поэтому записываем его в базу
        If pdmQuestions(k, 3) = 1 Then
            ' прибавляем к сумме ошибок
            sumPdmErrors = sumPdmErrors + 1
            ' определяем номер столбца нужного вопроса
            colNumber = qHeader.Item(Sheet_PDM_Check.Cells(k + 1, 2))(2)
            ' записываем ошибку в базу
            Sheet_DataBase.Cells(rowNumber, colNumber).Value = 1
        End If
    Next k
    
    ' записываем суммарное количество ошибок в секции PDM
    Sheet_DataBase.Cells(rowNumber, qHeader.Item("PDM_SUMM")(2)).Value = sumPdmErrors
    
End Sub

Sub saveTaskStatus(rowNumber As Integer)
    
    ' если Rework = "FINISHED", записываем в поле TASK STATUS
    ' значение "Completed", иначе - "Incompleted"
    If Sheet_IP_Check.reworkComboBox.Value = "FINISHED" Then
        Sheet_DataBase.Cells(rowNumber, "BR").Value = "Completed"
    Else
        Sheet_DataBase.Cells(rowNumber, "BR").Value = "Incompleted"
    End If
    
End Sub

Sub deleteExistingCheck(rowNumber As Integer)
    
    Sheet_DataBase.Rows(rowNumber).EntireRow.Delete

'    ' тестовый вариант - закрашиваем строку, которая удаляется, красным цветом
'    Sheet_DataBase.Rows(rowNumber).Interior.Color = vbRed
    
End Sub

Sub deleteExistingDescription(rowNumber As Integer)
    
    Sheet_ErrDescr.Rows(rowNumber).EntireRow.Delete

'    ' тестовый вариант - закрашиваем строку, которая удаляется, красным цветом
'    Sheet_ErrDescr.Rows(rowNumber).Interior.Color = vbRed
    
End Sub

Sub saveDescr()
    
    ' -----------------------------------------------
    '       СОХРАНЯЕМ ОПИСАНИЕ ОШИБОК В СЕКЦИИ IP
    ' -----------------------------------------------
    ' диапазон таблицы с описанием ошибок
    Dim descrTableRange As Range
    
    ' количество строк с описанием
    Dim descrRows As Integer
    
    ' диапазон динамической таблицы с описанием ошибок
    Set descrTableRange = Sheet_IP_Check.ListObjects("IpDescrTable").Range
    
    ' количество строк с ошибками
    descrRows = descrTableRange.Rows.Count
    
    ' номер последней заполненной строки на листе PERFORMER
    Dim dbDescrRowNumber As Integer
    dbDescrRowNumber = getDescrRow()
    
    ' если первая строка в таблице не пустая
    ' (для предотвращения записи пустых строк на лист PERFORMER)
    If descrTableRange(2, 1).Value <> "" Then
    
        For i = 2 To descrRows
            
            ' записываем атрибуты
            Call saveDescrAttributes(dbDescrRowNumber + i - 2)
            
            ' записываем код ошибки и описание
            Sheet_ErrDescr.Cells(dbDescrRowNumber + i - 2, "H").Value = descrTableRange(i, 1)
            Sheet_ErrDescr.Cells(dbDescrRowNumber + i - 2, "I").Value = descrTableRange(i, 2)
            
        Next i
    
    End If
    
    ' -----------------------------------------------
    '       СОХРАНЯЕМ ОПИСАНИЕ ОШИБОК В СЕКЦИИ PDM
    ' -----------------------------------------------
    
    ' диапазон таблицы с описанием ошибок
    Dim pdmDescrTableRange As Range
    
    ' количество строк с описанием
    Dim pdmDescrRows As Integer
    
    ' диапазон динамической таблицы с описанием ошибок
    Set pdmDescrTableRange = Sheet_PDM_Check.ListObjects("PdmDescrTable").Range
    
    ' количество строк с ошибками
    pdmDescrRows = pdmDescrTableRange.Rows.Count
    
    ' номер последней заполненной строки на листе PERFORMER
    dbDescrRowNumber = getDescrRow()
    
    If pdmDescrTableRange(2, 1).Value <> "" Then
    
        For i = 2 To pdmDescrRows
            
            ' записываем атрибуты
            Call saveDescrAttributes(dbDescrRowNumber + i - 2)
            
            ' записываем код ошибки и описание
            Sheet_ErrDescr.Cells(dbDescrRowNumber + i - 2, "H").Value = pdmDescrTableRange(i, 1)
            Sheet_ErrDescr.Cells(dbDescrRowNumber + i - 2, "I").Value = pdmDescrTableRange(i, 2)
            
        Next i
        
    End If
    
End Sub

Sub saveDescrAttributes(rowNumber As Integer)

    ' сохраняем дату
    Sheet_ErrDescr.Cells(rowNumber, "A").Value = Sheet_IP_Check.Cells(1, "F")
    ' сохраняем RelRecNr
    Sheet_ErrDescr.Cells(rowNumber, "B").Value = Sheet_IP_Check.Cells(2, "F")
    ' сохраняем исполнителя
    Sheet_ErrDescr.Cells(rowNumber, "C").Value = Sheet_IP_Check.performerComboBox.Value
    ' сохраняем IP Number
    Sheet_ErrDescr.Cells(rowNumber, "D").Value = Sheet_IP_Check.Cells(4, "F")
    ' сохраняем номер модуля
    Sheet_ErrDescr.Cells(rowNumber, "E").Value = Sheet_IP_Check.Cells(5, "F")
    ' сохраняем Rework
    Sheet_ErrDescr.Cells(rowNumber, "F").Value = Sheet_IP_Check.reworkComboBox.Value
    ' сохраняем MESA status
    Sheet_ErrDescr.Cells(rowNumber, "G").Value = Sheet_IP_Check.mesaStatusComboBox.Value
    
End Sub

