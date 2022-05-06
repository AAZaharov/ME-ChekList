VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SearchForm 
   Caption         =   "Найти запись"
   ClientHeight    =   5115
   ClientLeft      =   50
   ClientTop       =   380
   ClientWidth     =   13900
   OleObjectBlob   =   "SearchForm.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SearchForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' number of column "RelRecNr"
Const RelRecNrColumn As Integer = 2
    
' header row number
Const headRowNum As Integer = 2
    
Private Sub searchButton_Click()
    Dim li As ListItem
    Dim lastFilledRow As Integer
    Dim dateCond, rrnCond, perfCond, ipNumCond, moduleCond, reworkCond, mesaCond, summCond As Boolean
    
    SearchResultListView.ListItems.clear
    
    lastFilledRow = Sheet_DataBase.Cells(Rows.Count, RelRecNrColumn).End(xlUp).Row
    
    ' column numbers
    Dim dateCol, rrnCol, perfCol, ipNumCol, moduleCol, reworkCol, mesaCol As Integer
    dateCol = 1
    rrnCol = 2
    perfCol = 3
    ipNumCol = 4
    moduleCol = 5
    reworkCol = 6
    mesaCol = 7
    
    Dim dateText, rrnText, perfText, ipNumText, moduleText, reworkText, mesaText As String
    
    For i = headRowNum + 1 To lastFilledRow
        ' check date condition
        If dateFromTextBox.Value = "" And dateToTextBox.Value = "" Then
            dateCond = True
        ElseIf dateFromTextBox.Value <> "" And dateToTextBox.Value = "" Then
            dc1 = CDate(Sheet_DataBase.Cells(i, dateCol).Value) >= CDate(dateFromTextBox.Text)
            dateCond = dc1
        ElseIf dateFromTextBox.Value = "" And dateToTextBox.Value <> "" Then
            dc2 = CDate(Sheet_DataBase.Cells(i, dateCol).Value) <= CDate(dateToTextBox.Text)
            dateCond = dc2
        Else
            dc1 = CDate(Sheet_DataBase.Cells(i, dateCol).Value) >= CDate(dateFromTextBox.Text)
            dc2 = CDate(Sheet_DataBase.Cells(i, dateCol).Value) <= CDate(dateToTextBox.Text)
            dateCond = dc1 * dc2
        End If
        
        ' check RelRecNumber condition
        If relRecNrTextBox.Text = "" Then
            rrnCond = True
        Else
            rrnCond = (InStr(1, Sheet_DataBase.Cells(i, rrnCol).Value, relRecNrTextBox.Text) > 0)
        End If
        
        ' check Performer condition
        If performerComboBox.Value = "" Then
            perfCond = True
        Else
            perfCond = (InStr(1, Sheet_DataBase.Cells(i, perfCol).Value, performerComboBox.Value) > 0)
        End If
        
        ' check IP Number condition
        If ipNumTextBox.Text = "" Then
            ipNumCond = True
        Else
            ipNumCond = (InStr(1, Sheet_DataBase.Cells(i, ipNumCol).Value, ipNumTextBox.Text) > 0)
        End If
        
        ' check Module condition
        If moduleNumTextBox.Text = "" Then
            moduleCond = True
        Else
            moduleCond = (InStr(1, Sheet_DataBase.Cells(i, moduleCol).Value, moduleNumTextBox.Text) > 0)
        End If
        
        ' check Rework condition
        If reworkComboBox.Value = "" Then
            reworkCond = True
        ElseIf reworkComboBox.Value = "In work" Then
            If Sheet_DataBase.Cells(i, reworkCol).Value <> "FINISHED" Then
                reworkCond = True
            Else
                reworkCond = False
            End If
        Else
            reworkCond = (InStr(1, UCase(Sheet_DataBase.Cells(i, reworkCol).Value), UCase(reworkComboBox.Value)) > 0)
        End If
        
        ' check MESA condition
        If mesaStatusComboBox.Value = "" Then
            mesaCond = True
        Else
            mesaCond = (InStr(1, UCase(Sheet_DataBase.Cells(i, mesaCol).Value), UCase(mesaStatusComboBox.Value)) > 0)
        End If
        
        ' summary condition
        summCond = rrnCond * dateCond * perfCond * ipNumCond * moduleCond * reworkCond * mesaCond
        
        If summCond Then
            
            ' add new row in listview control
            Set li = SearchResultListView.ListItems.Add()
                       
            li.Text = Sheet_DataBase.Cells(i, rrnCol)
            
            ' add date subitem
            dateText = CStr(Sheet_DataBase.Cells(i, dateCol))
            li.ListSubItems.Add , , dateText
            
            ' add IP Number subitem
            ipNumText = Sheet_DataBase.Cells(i, ipNumCol)
            li.ListSubItems.Add , , ipNumText
            
            ' add Rework subitem
            reworkText = CStr(Sheet_DataBase.Cells(i, reworkCol))
            li.ListSubItems.Add , , reworkText
            
            ' add Performer subitem
            perfText = Sheet_DataBase.Cells(i, perfCol)
            li.ListSubItems.Add , , perfText
            
            ' add MESA Status subitem
            mesaText = Sheet_DataBase.Cells(i, mesaCol)
            li.ListSubItems.Add , , mesaText
            
            ' add Module subitem
            moduleText = Sheet_DataBase.Cells(i, moduleCol)
            li.ListSubItems.Add , , moduleText
            
            ' add Row subitem
            li.ListSubItems.Add , , i
        End If
    Next i
    
    rowCountLabel.Caption = SearchResultListView.ListItems.Count

End Sub

' обработка двойного клика по элементу списка
Private Sub searchResultListView_DblClick()
    
    ' загружаем запись
    Call loadRecord
    
End Sub

' обработка клика по кнопке "next Rework"
Private Sub nextReworkButton_Click()
    
    ' загружаем запись
    Call loadRecord
    
    ' устанавливаем текущую дату
    Sheet_IP_Check.Cells(1, "F").Value = Date
    
    ' делаем следующий Rework
    Sheet_IP_Check.reworkComboBox.Value = getLastRework + 1
    
    ' переключаем режим сохранения на добавление записи
    Sheet_IP_Check.saveRecordToggleButton.Value = False
    
End Sub

' обработка клика по кнопке "Изменить"
Private Sub updateButton_Click()
    
    ' загружаем запись
    Call loadRecord
    
    ' переключаем режим сохранения на изменение записи
    Sheet_IP_Check.saveRecordToggleButton.Value = True
    
End Sub

' обработка клика по кнопке "Загрузить"
Private Sub loadButton_Click()
    
    ' загружаем запись
    Call loadRecord
    
End Sub

' обработка клика по списку
' если дата выбранной записи не равна сегодняшней,
' отключаем кнопку "Изменить"
Private Sub searchResultListView_Click()
    
    If SearchResultListView.SelectedItem.ListSubItems(1) <> Date Then
        updateButton.Enabled = False
    Else
        updateButton.Enabled = True
    End If
    
    If SearchResultListView.SelectedItem.ListSubItems(3) = "FINISHED" Then
        nextReworkButton.Enabled = False
    Else
        nextReworkButton.Enabled = True
    End If
    
End Sub

Private Sub searchResultListView_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With SearchResultListView
        
        .SortKey = ColumnHeader.Index - 1
        .Sorted = True
        
        If .SortOrder = lvwAscending Then
            .SortOrder = lvwDescending
        Else
            .SortOrder = lvwAscending
        End If
        
    End With
End Sub

Private Sub UserForm_Initialize()
    
    'fill performer combobox
    i = 1
    Do While Sheet_SendEmail.Cells(i, 1).Value <> ""
        performerComboBox.AddItem (Sheet_SendEmail.Cells(i, 1).Value)
        i = i + 1
    Loop
    
    'fill rework combobox
    reworkComboBox.AddItem ("In work")
    reworkComboBox.AddItem ("0")
    reworkComboBox.AddItem ("1")
    reworkComboBox.AddItem ("2")
    reworkComboBox.AddItem ("3")
    reworkComboBox.AddItem ("4")
    reworkComboBox.AddItem ("5")
    reworkComboBox.AddItem ("Finished")
    
    'fill MESA status combobox
    mesaStatusComboBox.AddItem ("no MESA")
    mesaStatusComboBox.AddItem ("In work")
    mesaStatusComboBox.AddItem ("Complete")
    
    ' look for last filled RelRecNr field
    Dim lastFilledRow As Integer
    lastFilledRow = Sheet_DataBase.Cells(Rows.Count, RelRecNrColumn).End(xlUp).Row
    
    ' fill listview
    With SearchResultListView
        .View = lvwReport
        .FullRowSelect = True
        .Gridlines = True
        With .ColumnHeaders
            .clear
            ' add RelRecNr header
            .Add , Sheet_DataBase.Cells(headRowNum, 2).Value, Sheet_DataBase.Cells(headRowNum, 2).Value, 55
            ' add Date header
            .Add , Sheet_DataBase.Cells(headRowNum, 1).Value, Sheet_DataBase.Cells(headRowNum, 1).Value, 50
            ' add IP Number header
            .Add , Sheet_DataBase.Cells(headRowNum, 4).Value, Sheet_DataBase.Cells(headRowNum, 4).Value, 70
            ' add Rework header
            .Add , Sheet_DataBase.Cells(headRowNum, 6).Value, Sheet_DataBase.Cells(headRowNum, 6).Value, 50
            ' add Performer header
            .Add , Sheet_DataBase.Cells(headRowNum, 3).Value, Sheet_DataBase.Cells(headRowNum, 3).Value, 115
            ' add MESA status header
            .Add , Sheet_DataBase.Cells(headRowNum, 7).Value, Sheet_DataBase.Cells(headRowNum, 7).Value, 60
            ' add Module header
            .Add , Sheet_DataBase.Cells(headRowNum, 5).Value, Sheet_DataBase.Cells(headRowNum, 5).Value
            ' add Row Number header
            .Add , "Row", "Row", 30
        End With
    End With
    
    StartUpPosition = 0
    Left = Application.Left + (0.5 * Application.Width) - 0.5 * Width
    Top = Application.Top + (0.5 * Application.Height) - 0.5 * Height
    
End Sub

Private Sub loadAttrFields()
    
        ' write Date field
        Sheet_IP_Check.Cells(1, "F").Value = SearchResultListView.SelectedItem.SubItems(1)
        
        ' write RelRecNr field
        Sheet_IP_Check.Cells(2, "F").Value = SearchResultListView.SelectedItem.Text
        
        ' write performer ComboBox
        Sheet_IP_Check.performerComboBox.Value = SearchResultListView.SelectedItem.SubItems(4)
        
        ' write IP Number field
        Sheet_IP_Check.Cells(4, "F").Value = SearchResultListView.SelectedItem.SubItems(2)
        
        ' write Module field
        Sheet_IP_Check.Cells(5, "F").Value = SearchResultListView.SelectedItem.SubItems(6)
        
        ' write Rework ComoBox
        Sheet_IP_Check.reworkComboBox.Value = SearchResultListView.SelectedItem.SubItems(3)
        
        ' write MESA Status ComoBox
        Sheet_IP_Check.mesaStatusComboBox.Value = SearchResultListView.SelectedItem.SubItems(5)
        
End Sub

Private Sub loadErrQuestions()

    ' row number in DataBase for selected record
    Dim dbRowNumber As Integer
    dbRowNumber = SearchResultListView.SelectedItem.SubItems(7)
    
    ' clear IP questions zone
    For i = 3 To 39
        Sheet_IP_Check.Cells(i, "C").Value = ""
    Next i
    
    Dim errRecord As Integer
    
    ' fill IP questions zone
'    For i = 27 To 56
    For i = 8 To 68
        errRecord = Sheet_DataBase.Cells(dbRowNumber, i).Value
        If errRecord = 1 Then
            questionCode = Sheet_DataBase.Cells(headRowNum, i).Value
            For k = 3 To 39
                If questionCode = CStr(Sheet_IP_Check.Cells(k, 1).Value) Then
                    Sheet_IP_Check.Cells(k, 3).Value = 1
                End If
            Next k
        End If
    Next i
    
    ' clear PDM questions zone
    For i = 2 To 19
        Sheet_PDM_Check.Cells(i, "D").Value = ""
    Next i
    
    ' fill PDM questions zone
    For i = 51 To 68
        errRecord = Sheet_DataBase.Cells(dbRowNumber, i).Value
        If errRecord = 1 Then
            questionCode = Sheet_DataBase.Cells(headRowNum, i).Value
            For k = 2 To 19
                If questionCode = CStr(Sheet_PDM_Check.Cells(k, 2).Value) Then
                    Sheet_PDM_Check.Cells(k, 4).Value = 1
                End If
            Next k
        End If
    Next i
    
End Sub

Function isEqualDescr(rowNum As Variant) As Boolean
    
    ' condition variables
    Dim dateCond, rrnCond, perfCond, ipNumCond, moduleCond, reworkCond, mesaCond, errCodeCond As Boolean
    
    ' column numbers
    Dim dateCol, rrnCol, perfCol, ipNumCol, moduleCol, reworkCol, mesaCol, errCodeCol, descrCol As Integer
    dateCol = 1
    rrnCol = 2
    perfCol = 3
    ipNumCol = 4
    moduleCol = 5
    reworkCol = 6
    mesaCol = 7
    errCodeCol = 8
    descrCol = 9
    
    'all conditions to false
    dateCond = False
    rrnCond = False
    perfCond = False
    ipNumCond = False
    moduleCond = False
    reworkCond = False
    mesaCond = False
    
    ' check date condition
    If Sheet_ErrDescr.Cells(rowNum, dateCol).Value _
            = Sheet_IP_Check.Cells(1, "F").Value Then
        dateCond = True
    End If
    
    ' check RelRecNumber condition
    If CStr(Sheet_ErrDescr.Cells(rowNum, rrnCol).Value) _
            = Sheet_IP_Check.Cells(2, "F").Value Then
        rrnCond = True
    End If
    
    ' check Performer condition
    If Sheet_ErrDescr.Cells(rowNum, perfCol).Value = Sheet_IP_Check.performerComboBox.Value Then
        perfCond = True
    End If
    
    ' check IP Number condition
    If Sheet_ErrDescr.Cells(rowNum, ipNumCol).Value _
            = Sheet_IP_Check.Cells(4, "F").Value Then
        ipNumCond = True
    End If
    
    ' check Module condition
    If Sheet_ErrDescr.Cells(rowNum, moduleCol).Value _
            = Sheet_IP_Check.Cells(5, "F").Value Then
        moduleCond = True
    End If
    
    ' check Rework condition
'    If Sheet_ErrDescr.Cells(rowNum, reworkCol).Value = CInt(Sheet_IP_Check.reworkComboBox.Value) Then
    If CStr(Sheet_ErrDescr.Cells(rowNum, reworkCol).Value) = Sheet_IP_Check.reworkComboBox.Value Then
        reworkCond = True
    End If
    
    ' check MESA condition
    If Sheet_ErrDescr.Cells(rowNum, mesaCol).Value = Sheet_IP_Check.mesaStatusComboBox.Value Then
        mesaCond = True
    End If
    
    ' возвращаем общее условие
    isEqualDescr = rrnCond * dateCond * perfCond * ipNumCond * moduleCond * reworkCond * mesaCond
    
End Function

Sub loadRecord()

    If SearchResultListView.ListItems.Count > 0 Then
        If loadedDescrRowNums Is Nothing Then
            Call initLoadedDescrRowNumbers
        End If
        
        ' загружаем атрибуты записи
        Call loadAttrFields
        
        ' загружаем отметки ошибок (столбец "C" на листе Checklist)
        Call loadErrQuestions
        
        ' загружаем описание ошибок
        Call loadErrDescr
        
        ' закрываем форму
        Unload SearchForm
    Else
        MsgBox ("Не выбрано ни одной записи")
    End If

End Sub

Sub loadErrDescr()
    
    Dim summCond As Boolean
    Dim errCodeCol As Integer
    errCodeCol = 8
    Dim descrCol As Integer
    descrCol = 9

    Dim firstIpDescrRow As Integer, lastIpDescrRow As Integer
    Dim firstPdmDescrRow As Integer, lastPdmDescrRow As Integer

    ' last filled row on sheet PERFORMER
    Dim lastDescrRecordRow As Integer
    lastDescrRecordRow = Sheet_ErrDescr.Cells(Rows.Count, "B").End(xlUp).Row
    
    ' row number in DataBase for selected record
    Dim dbRowNumber As Integer
    dbRowNumber = SearchResultListView.SelectedItem.SubItems(7)
    
    ' количество ошибок в секции IP
    Dim ipErrCount As Integer
    ipErrCount = Sheet_DataBase.Cells(dbRowNumber, "BS").Value
    
    ' количество ошибок в секции PDM
    Dim pdmErrCount As Integer
    pdmErrCount = Sheet_DataBase.Cells(dbRowNumber, "BT").Value
    
    If ipErrCount > 0 Or pdmErrCount > 0 Then
    
        For i = 3 To lastDescrRecordRow
        
            summCond = isEqualDescr(i)
            
            ' если атрибуты совпадают,
            ' находим нужный номер вопроса и загружаем описание
            If summCond Then
                
                ' если в секции IP есть ошибки
                If ipErrCount > 0 Then
                
                    ' находим первую и последнюю строку с описанием ошибок в секции IP
                    firstIpDescrRow = Sheet_IP_Check.ListObjects("IpDescrTable").DataBodyRange.Row
                    lastIpDescrRow = Sheet_IP_Check.Cells(Rows.Count, "J").End(xlUp).Row
                    
                    ' загружаем описание в секцию IP
                    For k = firstIpDescrRow To lastIpDescrRow
                        
                        If Sheet_IP_Check.Cells(k, "J").Value = _
                                Sheet_ErrDescr.Cells(i, errCodeCol).Value Then
                            errCodeCond = True
                        Else
                            errCodeCond = False
                        End If
                        
                        If errCodeCond Then
                            ' записываем описание ошибки на лист Checklist
                            Sheet_IP_Check.Cells(k, "K").Value = Sheet_ErrDescr.Cells(i, descrCol).Value
                            ' добавляем номер строки в коллекцию
                            loadedDescrRowNums.Add (i)
                        End If
                        
                    Next k
                    
                End If
                
                ' если в секции PDM есть ошибки
                If pdmErrCount > 0 Then
                
                    ' находим первую и последнюю строку с описанием ошибок в секции PDM
                    firstPdmDescrRow = Sheet_PDM_Check.ListObjects("PdmDescrTable").DataBodyRange.Row
                    lastPdmDescrRow = Sheet_PDM_Check.Cells(Rows.Count, "J").End(xlUp).Row
                    
                    ' загружаем описание в секцию PDM
                    For k = firstPdmDescrRow To lastPdmDescrRow
                        
                        If Sheet_PDM_Check.Cells(k, "J").Value = _
                                Sheet_ErrDescr.Cells(i, errCodeCol).Value Then
                            errCodeCond = True
                        Else
                            errCodeCond = False
                        End If
                        
                        If errCodeCond Then
                            ' записываем описание ошибки на лист Checklist
                            Sheet_PDM_Check.Cells(k, "K").Value = Sheet_ErrDescr.Cells(i, descrCol).Value
                            ' добавляем номер строки в коллекцию
                            loadedDescrRowNums.Add (i)
                        End If
                        
                    Next k
                
                End If
                
            End If
            
        Next i
    
    End If

End Sub
