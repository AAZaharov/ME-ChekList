Attribute VB_Name = "ValidationModule"
' флаг проверки наличия исполнителя в списке (лист SendEmail)
Dim isPerfNotExist As Boolean
' флаг отправки письма
Dim isSendEmail As Boolean

' флаг для проверки существования пустых полей описания ошибок
' используется, если не выбрана опция "сохранять без описания ошибок"
Dim isEmptyDescrExist As Boolean

' флаг для проверки, заполнены ли обязательные поля -
' RelRecNr, Performer и IP Number
Dim isWrongAttrExist As Boolean

' флаг для проверки правильности заполнения поля RelRecNr
Dim isRelRecNrValid As Boolean

' сообщение об ошибке заполнения поля даты
Dim dateErrMsg As String
' сообщение об ошибке заполнения поля RelRecNr
Dim rrnErrMsg As String
' сообщение об ошибке заполнения поля Performer
Dim perfErrMsg As String
' сообщение об ошибке заполнения поля IP Number
Dim ipNumErrMsg As String
' все ошибки
Dim wrongAttrMsg As String

' флаг правильности записи
Dim isRecordValid As Boolean
    

Function beforeSaveValidation() As Boolean
    ' проверяем, нажата или нет кнопка режима записи,
    ' и вызываем соответствующую функцию для проверки данных
    If Sheet_IP_Check.saveRecordToggleButton.Value Then ' если кнопка нажата, делаем проверку для редактирования записи
        isRecordValid = updSaveRecValidation
    Else ' иначе делаем проверку для добавления записи
        isRecordValid = addRecSaveValidation()
    End If
    
    ' если в данных найдены ошибки, завершаем проверку
    If isRecordValid = False Then
        beforeSaveValidation = False
        Exit Function
    End If
    
    ' если не выбрана опция "сохранять без описания ошибок",
    ' проверяем наличие описания у каждого поля
    If Sheet_IP_Check.saveWithoutDescrCheckBox.Value = False Then
        isEmptyDescrExist = emptyDescrValidation()
    Else
        isEmptyDescrExist = False
    End If
    
    ' Если описание отсутствует, показываем сообщение
    ' и выходим из процедуры проверки
    If isEmptyDescrExist Then
        MsgBox Prompt:="Отсутствует описание ошибок. Сохранение прервано", _
               Title:="Ошибка проверки", _
               Buttons:=vbExclamation
        beforeSaveValidation = False
        Exit Function
    End If
    
    ' если включена опция отправки письма, проверяем
    ' наличие исполнителя в списке с почтовыми адресами
    If Sheet_IP_Check.sendErrDescrCheckBox.Value _
        Or Sheet_IP_Check.sendFinishedStateCheckBox.Value Then
        isPerfNotExist = perfNotExistValidation()
        isSendEmail = True
    Else
        isSendEmail = False
    End If
    
    ' если исполнителя нет в списке, а опция отправки письма выбрана,
    ' показываем диалоговое окно с выбором варианта действий
    If isPerfNotExist And isSendEmail Then
        ' показываем диалоговое окно и запоминаем ответ
        noPerformerAction = MsgBox(Prompt:="Невозможно отправить письмо - " & vbNewLine _
                                         & "указанного исполнителя нет в списке" & vbNewLine _
                                         & "Сохранить без отправки письма?", _
                                    Title:="Ошибка проверки", _
                                    Buttons:=vbYesNo + vbExclamation)
        
        If noPerformerAction = vbNo Then
            beforeSaveValidation = False
            Exit Function
        Else
'            Sheet_IP_Check.sendErrDescrCheckBox.Value = False
'            Sheet_IP_Check.sendFinishedStateCheckBox.Value = False
        End If
        
'        beforeSaveValidation = False
'        Exit Function
    End If
    
    ' если функция доработала до конца, значит, ошибок нет
    beforeSaveValidation = True
    
End Function

Function emptyDescrValidation()

    ' fill description zone
    ' find first and last rows with errors description
    Dim firstDescrRow As Long, lastDescrRow As Long
    ' первая строка таблицы IpDescrTable
    firstDescrRow = Sheet_IP_Check.ListObjects("IpDescrTable").DataBodyRange.Row
    ' последняя заполненная строка (с номером вопроса) в таблице описания
    lastDescrRow = Sheet_IP_Check.Cells(Rows.Count, "J").End(xlUp).Row
    
    ' предполагаем, что ошибок нет
    emptyDescrValidation = False
    
    ' если есть ошибки (первая строка таблицы не пустая)
    ' то проверяем наличие описания
    If Sheet_IP_Check.Cells(firstDescrRow, "J").Value <> "" Then
        
        ' проверяем наличие описания
        For i = firstDescrRow To lastDescrRow
            If Sheet_IP_Check.Cells(i, "K").Value = "" Then
                ' если найдены пустые поля, выходим из функции
                emptyDescrValidation = True
                Exit Function
            End If
        Next i

    End If
    
    ' то же самое для ошибок PDM
    ' первая строка таблицы PdmDescrTable
    firstDescrRow = Sheet_PDM_Check.ListObjects("PdmDescrTable").DataBodyRange.Row
    ' последняя заполненная строка (с номером вопроса) в таблице описания
    lastDescrRow = Sheet_PDM_Check.Cells(Rows.Count, "J").End(xlUp).Row
    
    If Sheet_PDM_Check.Cells(firstDescrRow, "J").Value <> "" Then
        
        For i = firstDescrRow To lastDescrRow
            If Sheet_PDM_Check.Cells(i, "K").Value = "" Then
                ' если найдены пустые поля, выходим из функции
                emptyDescrValidation = True
                Exit Function
            End If
        Next i
    
    End If

End Function

Function perfNotExistValidation() As Boolean

    ' номер строки с фамилией исполнителя
    Dim selectedPerfRow As Integer
    
    selectedPerfRow = getPerformer()
    
    If selectedPerfRow = 0 Then
        perfNotExistValidation = True
    Else
        perfNotExistValidation = False
    End If
    
End Function

Function updSaveRecValidation() As Boolean
    
    ' проверка атрибутов на соответствие записи в базе данных
    ' по номеру RelRecNr находим запись в базе данных
    Dim updatedRow As Integer
    updatedRow = getUpdatedRow()
    
    ' если такая строка существует, возвращаем значение True, иначе False
    If updatedRow > 0 Then
        updSaveRecValidation = True
    Else
        MsgBox ("В базе нет строки с такими RelRecNr, Rework и IP Number" & vbNewLine _
              & "Сохранение прервано.")
        updSaveRecValidation = False
    End If
    
End Function

Function addRecSaveValidation() As Boolean
    ' проверка атрибутов
    isWrongAttrExist = wrongAttrCheck()
    
    If isWrongAttrExist Then
        addRecSaveValidation = False
        Exit Function
    End If
    
    ' проверка полей RelRecNr и Rework
    isRelRecNrValid = relRecNrCheck()
    
    If isRelRecNrValid = False Then
        addRecSaveValidation = False
        Exit Function
    End If
    
    addRecSaveValidation = True
    
End Function

Function wrongAttrCheck() As Boolean
    
    wrongAttrMsg = ""
    
    ' проверка даты
    Call dateFieldCheck
    ' проверка поля RelRecNr
    Call rrnFieldCheck
    ' проверка поля Performer
    Call perfFieldCheck
    ' проверка поля Performer
    Call ipNumFieldCheck
    
    Prompt = "Найдены ошибки:" & vbNewLine _
                                & vbNewLine
    
    wrongAttrMsg = dateErrMsg _
                    & rrnErrMsg _
                    & perfErrMsg _
                    & ipNumErrMsg
    
    If wrongAttrMsg = "" Then
        wrongAttrCheck = False
    Else
        wrongAttrCheck = True
        MsgBox Prompt:=Prompt & wrongAttrMsg, Title:="Проверка ошибок", Buttons:=vbExclamation
    End If
    
End Function

Sub dateFieldCheck()
    dateErrMsg = ""
    ' проверка, заполнено поле даты
    If Sheet_IP_Check.Cells(1, "F") = "" Then
        dateErrMsg = "Не заполнено поле даты." & vbNewLine
    End If
    ' для создания новой записи дата должна быть не старше текущей
    If Sheet_IP_Check.Cells(1, "F") <> "" And CDate(Sheet_IP_Check.Cells(1, "F")) < Date Then
        dateErrMsg = "Указанная дата уже прошла. Укажите актуальную дату." & vbNewLine
    End If
End Sub

Sub rrnFieldCheck()
    rrnErrMsg = ""
    ' проверка, заполнено ли поле RelRecNr
    If Sheet_IP_Check.Cells(2, "F") = "" Then
        rrnErrMsg = "Не заполнено поле RelRecNr." & vbNewLine
    End If
End Sub

Sub perfFieldCheck()
    perfErrMsg = ""
    ' проверка, заполнено ли поле Performer
    If Sheet_IP_Check.performerComboBox.Value = "" Then
        perfErrMsg = "Не заполнено поле Performer." & vbNewLine
    End If
End Sub

Sub ipNumFieldCheck()
    ipNumErrMsg = ""
    ' проверка, заполнено ли поле IP Number
    If Sheet_IP_Check.Cells(4, "F") = "" Then
        ipNumErrMsg = "Не заполнено поле IP Number." & vbNewLine
    End If
End Sub

Function relRecNrCheck() As Boolean
    ' если RelRecNr новый, то проверить поле Rework, оно должно быть равно 0
    If getRelRecNr() = 0 Then
        
        If Sheet_IP_Check.reworkComboBox.Value <> "0" _
        And Sheet_IP_Check.reworkComboBox.Value <> "FINISHED" Then
            
            If MsgBox(Prompt:="Для нового RelRecNr поле Rework" & vbNewLine _
                    & "должно быть равно 0." & vbNewLine _
                    & "Установить Rework в 0 и продолжить?", Buttons:=vbOKCancel) = vbOK Then
                Sheet_IP_Check.reworkComboBox.Value = 0
            Else
                relRecNrCheck = False
                Exit Function
            End If
        
        End If
    
    End If
    
    ' если запись с таким RelRecNr существует, то
    If getRelRecNr() > 0 Then
    
        ' проверить поле Rework на значение "FINISHED"
        If getFinishedRRN() > 0 Then
        
            MsgBox ("Работа с таким RelRecNr уже закончена." & vbNewLine _
                 & "(Rework = FINISHED)" & vbNewLine _
                 & "Сохранение прервано.")
                 
            relRecNrCheck = False
            Exit Function
            
        End If
        
        ' проверить наличие записи с таким же полем Rework
        ' и предложить увеличить номер
        If getEqualRework() > 0 Then
            
            ' находим в базе номера Rework для выбранных RelRecNr и IP Number
            Dim allRework As Collection
            Set allRework = getAllReworks
            
            Dim reworksMsg As String
            reworksMsg = allRework.Item(1)
            
            For i = 2 To allRework.Count
                reworksMsg = reworksMsg & ", " & allRework.Item(i)
            Next i
            
            If allRework.Count = 1 Then
                
                If MsgBox(Prompt:="Запись с Rework = " & getLastRework() & " уже существует в базе." & vbNewLine _
                        & vbNewLine _
                        & "Установить Rework в " & getLastRework() + 1 & " и продолжить?", Buttons:=vbOKCancel) = vbOK Then
                        
                    Sheet_IP_Check.reworkComboBox.Value = getLastRework() + 1
                Else
                    relRecNrCheck = False
                    Exit Function
                End If
            
            Else
                
                If MsgBox(Prompt:="Записи с Rework = " & reworksMsg & " уже существуют в базе." & vbNewLine _
                        & vbNewLine _
                        & "Установить Rework в " & getLastRework() + 1 & " и продолжить?", Buttons:=vbOKCancel) = vbOK Then
                        
                    Sheet_IP_Check.reworkComboBox.Value = getLastRework() + 1
                Else
                    relRecNrCheck = False
                    Exit Function
                End If
            
            End If
            
        End If
        
    End If
    
    relRecNrCheck = True
    
End Function
