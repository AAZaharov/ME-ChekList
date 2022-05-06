Attribute VB_Name = "SendMailModule"
Sub sendMail(rowNum As Integer)
    
    ' адрес получателя
    Dim mailTo As String
    ' адрес получателя копии
    Dim Email_CC As String
    
    ' атрибуты плана
    Dim inputValue As String
    
    ' описание ошибок IP
    Dim ipErrDescr As String
    
    ' описание ошибок PDM
    Dim pdmErrDescr As String
    
    On Error Resume Next
    
    Dim performerRow, lastPerformer As Integer
    lastPerformer = Sheet_SendEmail.Cells(Rows.Count, 1).End(xlUp).Row
    
    For counter = 1 To lastPerformer
        If Sheet_SendEmail.Cells(counter, 1).Value = Sheet_IP_Check.performerComboBox.Value Then
            mailTo = Sheet_SendEmail.Cells(counter, 2)
        End If
    Next counter
    
    If mailTo = "" Then
        MsgBox "Указанного исполнителя нет в списке." & vbNewLine _
        & vbNewLine _
        & "Письмо не будет отправлено", vbExclamation
        Exit Sub
    End If
    
    ' если выбрано не отправлять письмо, выходим из функции
    If MsgBox("Редактировать письмо: " & mailTo, vbYesNo, "Подтверждение отправки") = vbNo Then
        Exit Sub
    End If
    
    Dim mailSubject As String
    mailSubject = Sheet_IP_Check.Cells(2, "F").Value & "," & " " & Sheet_IP_Check.Cells(4, "F").Value
    
    'Email Text
    
    Dim ipNotifString As String, J As Integer, pdmNotifString As String
    If getSumIpErrors(rowNum) = 0 Then
        ipNotifString = ""
    End If

    If getSumPdmErrors(rowNum) = 0 Then
        pdmNotifString = ""
    End If
    
    For i = 1 To 7
        Input_value = Input_value & " " & Sheet_IP_Check.Cells(i, "E").Value & " : " & Sheet_IP_Check.Cells(i, "F").Value
    Next i
    
    ipErrDescr = ipErrMailText
    
    pdmErrDescr = pdmErrMailText
    
    Dim mailText As String
    mailText = Input_value & vbNewLine _
                & vbNewLine _
                & ipNotifString & vbNewLine _
                & vbNewLine _
                & ipErrDescr & vbNewLine _
                & pdmNotifString & vbNewLine _
                & vbNewLine _
                & pdmErrDescr
        
    RES = SendEmailUsingOutlook(mailTo, mailText, mailSubject)
    
    If RES Then
        ' записываем поле EMAIL STATUS
        Sheet_DataBase.Cells(rowNum, "BQ").Value = "Yes"
        
        ' показываем сообщение, что письмо отправлено
        MsgBox ("Письмо создано успешно")
    Else
        MsgBox ("Ошибка")
    End If

End Sub

Sub sendFinishedMail(rowNum As Integer)

    ' атрибуты плана
    Dim inputValue As String
    
    ' адрес получателя
    Dim mailTo As String
    
    ' описание ошибок IP
    Dim ipErrDescr As String
    
    ' описание ошибок PDM
    Dim pdmErrDescr As String
    
    On Error Resume Next
    
    ' проверка, есть ли указанный исполнитель в списке с адресами эл. почты
    ' массив для копирования исполнителей
    Dim performers() As Variant
    ' последняя заполненная строка
    Dim lastPerfRow As Integer
    lastPerfRow = Sheet_SendEmail.Cells(Rows.Count, "A").End(xlUp).Row
    
    Dim mailR As Range
    Set mailR = Range(Sheet_SendEmail.Cells(1, "A"), Sheet_SendEmail.Cells(lastPerfRow, "B"))
    performers = mailR.Value
    
    ' выбранный исполнитель
    Dim p As String
    p = Sheet_IP_Check.performerComboBox.Value
    
    ' ищем в массиве нужный адрес эл. почты
    For i = 1 To lastPerfRow
        If performers(i, 1) = p Then
            mailTo = performers(i, 2)
        End If
    Next i
    
    ' если адрес не найден, показываем сообщение и завершаем процедуру
    If mailTo = "" Then
        MsgBox "Указанного исполнителя нет в списке." & vbNewLine _
        & vbNewLine _
        & "Письмо не будет отправлено", vbExclamation
        Exit Sub
    End If
    
    ' атрибуты плана
    For i = 1 To 7
        inputValue = inputValue & " " & Sheet_IP_Check.Cells(i, "E").Value & " : " & Sheet_IP_Check.Cells(i, "F").Value & vbNewLine
    Next i

    ' текст письма
    Dim mailText As String
    
    If getSumIpErrors(rowNum) = 0 And getSumPdmErrors(rowNum) = 0 Then
        mailText = inputValue & vbNewLine _
                 & "Ошибок не найдено. Работа сохранена в базу со статусом Completed."
    Else
        mailText = inputValue & vbNewLine _
                 & "Ошибки исправлены чекером. Работа сохранена в базу со статусом Completed."
    End If
    
    ipErrDescr = ipErrMailText
    
    pdmErrDescr = pdmErrMailText
    
    mailText = mailText & vbNewLine _
                & vbNewLine _
                & ipErrDescr & vbNewLine _
                & vbNewLine _
                & pdmErrDescr
    
    ' тема письма
    Dim mailSubject As String
    mailSubject = "Checklist for " & Sheet_IP_Check.Cells(2, "F").Value & "," & " " & Sheet_IP_Check.Cells(4, "F").Value
            
    ' вызываем функцию отправки письма с помощью Outlook
    RES = SendEmailUsingOutlook(mailTo, mailText, mailSubject)
    
    If RES Then
        ' записываем поле EMAIL STATUS
        Sheet_DataBase.Cells(rowNum, "BQ").Value = "Yes"
        
        ' показываем сообщение, что письмо отправлено
        MsgBox ("Письмо отправлено успешно")
    Else
        MsgBox ("Ошибка отправки")
    End If

End Sub

Function SendEmailUsingOutlook(ByVal Email_To$, ByVal mailText$, Optional ByVal Subject$ = "", _
                               Optional ByVal AttachFilename As Variant) As Boolean
'ByVal Email_CC$
    On Error Resume Next: Err.clear
    Dim OA As Object: Set OA = CreateObject("Outlook.Application")
    Set objOutlookMsg = Outapp.CreateItem(olMailItem)
    If OA Is Nothing Then MsgBox "The Application was not found", vbCritical: Exit Function
    st = 1
    With OA.CreateItem(0)
        .To = Email_To$:
       ' .CC = Email_CC$:
        .Subject = Subject$:
        .Body = mailText$
        If VarType(AttachFilename) = vbString Then .Attachments.Add AttachFilename
        If VarType(AttachFilename) = vbObject Then    ' AttachFilename as Collection
            For Each file In AttachFilename: .Attachments.Add file: Next
            objOutlookMsg.Display
            Set Outapp = Nothing
        End If
        
        ' строка для задержки (без задержки не отправляются письма с вложениями)
        ' у нас нет вложений, поэтому отключаем этот код
        ' For i = 1 To 100000: DoEvents: Next

        Err.clear: .Display
        SendEmailUsingOutlook = Err = 0
    End With
    Set Outapp = Nothing
    
End Function

' -----------------------------------------------
'        ОПИСАНИЕ ОШИБОК В СЕКЦИИ IP
' -----------------------------------------------
Function ipErrMailText() As String

    ' строка для описания ошибок в секции IP
    Dim ipErrMailString As String
    ' диапазон таблицы с описанием ошибок
    Dim descrTableRange As Range
    
    ' количество строк с описанием
    Dim descrRows As Integer
    
    ' диапазон динамической таблицы с описанием ошибок
    Set descrTableRange = Sheet_IP_Check.ListObjects("IpDescrTable").Range
    
    ' количество строк с ошибками
    descrRows = descrTableRange.Rows.Count
    
    ' если в таблице есть вопросы
    If descrTableRange(2, 1) <> "" Then
    
        ipErrMailString = "Ошибки в секции ADPP" & vbNewLine _
                      & "----------------------------" & vbNewLine _
                      & vbNewLine
        
        For i = 2 To descrRows
            
            ' записываем код ошибки и описание
            ipErrMailString = ipErrMailString & "Вопрос " & descrTableRange(i, 1) & ": " & descrTableRange(i, 2) & vbNewLine _
                                & vbNewLine
            
        Next i
        
    End If
    
    ipErrMailText = ipErrMailString
    
End Function

' -----------------------------------------------
'        ОПИСАНИЕ ОШИБОК В СЕКЦИИ PDM
' -----------------------------------------------
Function pdmErrMailText() As String

    ' строка для описания ошибок в секции PDM
    Dim pdmErrMailString As String
    
    ' диапазон таблицы с описанием ошибок
    Dim pdmDescrTableRange As Range
    
    ' количество строк с описанием
    Dim pdmDescrRows As Integer
    
    ' диапазон динамической таблицы с описанием ошибок PDM
    Set pdmDescrTableRange = Sheet_PDM_Check.ListObjects("PdmDescrTable").Range
    
    ' количество строк с ошибками
    pdmDescrRows = pdmDescrTableRange.Rows.Count
    
    ' если в таблице есть вопросы
    If pdmDescrTableRange(2, 1) <> "" Then
    
       pdmErrMailString = "Ошибки в секции PDM" & vbNewLine _
                      & "------------------------------" & vbNewLine _
                      & vbNewLine
       
       For i = 2 To pdmDescrRows
            
            ' записываем код ошибки и описание
            pdmErrMailString = pdmErrMailString & "Вопрос " & pdmDescrTableRange(i, 1) & ": " & pdmDescrTableRange(i, 2) & vbNewLine _
                                & vbNewLine
            
        Next i
    
    End If
    
    pdmErrMailText = pdmErrMailString
    
End Function


