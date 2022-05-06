Attribute VB_Name = "SaveModule"
' ����� ������ ��� ������ �� ���� DataBase
Dim dbSaveRowNumber As Integer

' ����� ����������� ������ ��� ������ �� ���� DataBase
Dim dbUpdatedRowNumber As Integer

Sub validateAndSave()
    
    '��������� ��� ����������� �������
    Application.EnableEvents = False
    
    Dim saveResult As Boolean
    
    ' ��������� ��������
    ' � ���������� ��������� � ���������� ����������
    Dim isValidRecord As Boolean
    isValidRecord = ValidationModule.beforeSaveValidation
    
    Dim mailReason As String
    mailReason = "������ ��������: "
    
    ' ����� ������ � ����� ������� � DataBase
    Dim checkRowNum As Integer
    
    ' ���� �������� ��������, ��������� ������/����������� � ���������� ������
    If isValidRecord Then
        If Sheet_IP_Check.saveRecordToggleButton.Value = True Then
            checkRowNum = updateCheck()
            mailReason = mailReason & "�����������"
        Else
            checkRowNum = saveCheck()
            mailReason = mailReason & "����� ��������"
        End If
        
        ' ���������� � ���� EMAIL STATUS �������� "No", ���� ������ ����� ����������, �������� ����� ��������
        Sheet_DataBase.Cells(checkRowNum, "BQ").Value = "No"
        
        ' ���� ������� �����, ������ �������� �� ������� ������
        ' � ���������� ��������������� ������
        
        ' ���� ������ �������� Task Status = "Completed"
        If Sheet_DataBase.Cells(checkRowNum, "BR") = "Completed" Then
            ' ���� ������� ����� "���������� ������ �� ���������� ������"
            If Sheet_IP_Check.sendFinishedStateCheckBox.Value Then
            
                Call sendFinishedMail(checkRowNum)
                
            End If
        Else
            ' ���� ������� ����� "���������� ������ � ��������� ������"
            If Sheet_IP_Check.sendErrDescrCheckBox.Value Then
                
                Call sendMail(checkRowNum)
                
            End If
        End If
        
        ' ��������� ����
        Call SortModule.sortDataBase
        
        Sheet_IP_Check.Activate
        
        ' ��������� ����
        ActiveWorkbook.Save
        
        ' ���������� ����� � ����������� � ��������� ������
        ReportInfoForm.Show
        
    End If
    
    
    '�������� ��� ����������� �������
    Application.EnableEvents = True
    '---------------------------------------------
    
End Sub

Function saveCheck()

    ' ���������� ����� ������ ��� ����������
    dbSaveRowNumber = getRow()
        
    ' ���������� �������� ��������
    Call saveAttributes(dbSaveRowNumber)
    
    ' ��������� ������ � ������ IP
    Call saveIpQuestions(dbSaveRowNumber)
    
    ' ��������� ������ � ������ PDM
    Call savePdmQuestions(dbSaveRowNumber)
    
    ' ��������� ���� TASK STATUS
    Call saveTaskStatus(dbSaveRowNumber)
    
    ' ���� �������� ������
    Call saveDescr
    
    ' ���������� ����� ������ � �������
    saveCheck = dbSaveRowNumber
    
End Function

Function updateCheck()
    
    ' ���������� ����� ������ ������, ������� ������
    dbUpdatedRowNumber = getUpdatedRow()
    
    ' ������� ������ ������
    Call SaveModule.deleteExistingCheck(dbUpdatedRowNumber)
    
    ' ���������� ����� ������ ��� ����������
    dbSaveRowNumber = getRow()

    ' ���������� �������� ��������
    Call saveAttributes(dbSaveRowNumber)

    ' ��������� ������ � ������ IP
    Call saveIpQuestions(dbSaveRowNumber)

    ' ��������� ������ � ������ PDM
    Call savePdmQuestions(dbSaveRowNumber)

    ' ��������� ���� TASK STATUS
    Call saveTaskStatus(dbSaveRowNumber)
    
    ' ������� ������ ����� � ��������� ������,
    ' ������� ����� ��������
    Dim errDescrRowsColl As Collection
    Set errDescrRowsColl = getUpdatedDescrRows()
    
    ' ������� ������ � ��������� ������
    ' �������� ��� �� ������ ������ � �������
    For i = errDescrRowsColl.Count To 1 Step -1
        deleteExistingDescription (errDescrRowsColl(i))
    Next i
    
    ' ��������� ����� ������ � ��������� ������
    Call saveDescr
    
    ' ���������� ����� ������ � �������
    updateCheck = dbSaveRowNumber
    
End Function

Sub saveAttributes(rowNumber As Integer)

    ' ��������� ����
    Sheet_DataBase.Cells(rowNumber, "A").Value = Sheet_IP_Check.Cells(1, "F")
    ' ��������� RelRecNr
    Sheet_DataBase.Cells(rowNumber, "B").Value = Sheet_IP_Check.Cells(2, "F")
    ' ��������� �����������
    Sheet_DataBase.Cells(rowNumber, "C").Value = Sheet_IP_Check.performerComboBox.Value
    ' ��������� IP Number
    Sheet_DataBase.Cells(rowNumber, "D").Value = Sheet_IP_Check.Cells(4, "F")
    ' ��������� ����� ������
    Sheet_DataBase.Cells(rowNumber, "E").Value = Sheet_IP_Check.Cells(5, "F")
    ' ��������� Rework
    Sheet_DataBase.Cells(rowNumber, "F").Value = Sheet_IP_Check.reworkComboBox.Value
    ' ��������� MESA status
    Sheet_DataBase.Cells(rowNumber, "G").Value = Sheet_IP_Check.mesaStatusComboBox.Value
    
End Sub

Sub saveIpQuestions(rowNumber As Integer)
    
    ' ������ ��� ����������� ������ �������� �� ����� Checklist
    Dim questions() As Variant
    ' ��������� ����������� ������, ���������� �� ������� "B" � ������� �������
    Dim lastQRow As Integer
    lastQRow = Sheet_IP_Check.Cells(Rows.Count, "B").End(xlUp).Row
    
    Dim qR As Range
    Set qR = Range(Sheet_IP_Check.Cells(3, "A"), Sheet_IP_Check.Cells(lastQRow, "C"))
    questions = qR.Value
    
    ' ��������� ��� ����������� ������ ���������� �������� �� ����� DataBase
    Dim qHeader As Collection
    Set qHeader = New Collection
    ' ��������� ����������� ������� ��������� - ������ ������ ����� DataBase
    Dim lastHeadCol As Integer
    lastHeadCol = Sheet_DataBase.Cells(2, Columns.Count).End(xlToLeft).Column
    Dim itemOfColl(1 To 2) As Variant
    ' ��������� ��������� ��������� - � ������ �������� ������� ������� ����� �������� ������,
    ' �� ������ - ����� ������� ���� ������
    ' ���� - ����� � ������
    For i = 8 To lastHeadCol
        itemOfColl(1) = Sheet_DataBase.Cells(2, i).Value
        itemOfColl(2) = Sheet_DataBase.Cells(2, i).Column
        qHeader.Add Item:=itemOfColl, Key:=CStr(Sheet_DataBase.Cells(2, i))
    Next i
    
    ' ��������� ���������� ������ � ������ IP
    Dim sumIpErrors As Integer
    sumIpErrors = 0
    
    For k = 1 To UBound(questions, 1)
        ' ���� � ������� ������� �������, ��� ���������� ������,
        ' ������� ���������� ��� � ����
        If questions(k, 3) = 1 Then
            ' ���������� � ����� ������
            sumIpErrors = sumIpErrors + 1
            ' ���������� ����� ������� ������� �������
            colNumber = qHeader.Item(CStr(Sheet_IP_Check.Cells(k + 2, 1)))(2)
            ' ���������� ������ � ����
            Sheet_DataBase.Cells(rowNumber, colNumber).Value = 1
        End If
        
    Next k
    
    ' ���������� ��������� ���������� ������ � ������ IP
    Sheet_DataBase.Cells(rowNumber, qHeader.Item("IP_SUMM")(2)).Value = sumIpErrors
    
End Sub

Sub savePdmQuestions(rowNumber As Integer)
    
    ' ������ ��� ����������� ������ �������� �� ����� PDM_Checklist
    Dim pdmQuestions() As Variant
    ' ��������� ����������� ������, ���������� �� ������� "B" � ������� �������
    Dim lastPdmpdmQRow As Integer
    lastPdmpdmQRow = Sheet_PDM_Check.Cells(Rows.Count, "B").End(xlUp).Row
    
    Dim pdmQR As Range
    Set pdmQR = Range(Sheet_PDM_Check.Cells(2, "B"), Sheet_PDM_Check.Cells(lastPdmpdmQRow, "D"))
    pdmQuestions = pdmQR.Value
    
    ' ��������� ��� ����������� ������ ���������� �������� �� ����� DataBase
    Dim qHeader As Collection
    Set qHeader = New Collection
    ' ��������� ����������� ������� ��������� - ������ ������ ����� DataBase
    Dim lastHeadCol As Integer
    lastHeadCol = Sheet_DataBase.Cells(2, Columns.Count).End(xlToLeft).Column
    Dim itemOfColl(1 To 2) As Variant
    ' ��������� ��������� ��������� - � ������ �������� ������� ������� ����� �������� ������,
    ' �� ������ - ����� ������� ���� ������
    ' ���� - ����� � ������
    For i = 8 To lastHeadCol
        itemOfColl(1) = Sheet_DataBase.Cells(2, i).Value
        itemOfColl(2) = Sheet_DataBase.Cells(2, i).Column
        qHeader.Add Item:=itemOfColl, Key:=CStr(Sheet_DataBase.Cells(2, i))
    Next i
    
    ' ��������� ���������� ������ � ������ PDM
    Dim sumPdmErrors As Integer
    sumPdmErrors = 0
    
    For k = 1 To UBound(pdmQuestions, 1)
        ' ���� � �������� ������� ����� (������ ������� �������) �������,
        ' ��� ���������� ������, ������� ���������� ��� � ����
        If pdmQuestions(k, 3) = 1 Then
            ' ���������� � ����� ������
            sumPdmErrors = sumPdmErrors + 1
            ' ���������� ����� ������� ������� �������
            colNumber = qHeader.Item(Sheet_PDM_Check.Cells(k + 1, 2))(2)
            ' ���������� ������ � ����
            Sheet_DataBase.Cells(rowNumber, colNumber).Value = 1
        End If
    Next k
    
    ' ���������� ��������� ���������� ������ � ������ PDM
    Sheet_DataBase.Cells(rowNumber, qHeader.Item("PDM_SUMM")(2)).Value = sumPdmErrors
    
End Sub

Sub saveTaskStatus(rowNumber As Integer)
    
    ' ���� Rework = "FINISHED", ���������� � ���� TASK STATUS
    ' �������� "Completed", ����� - "Incompleted"
    If Sheet_IP_Check.reworkComboBox.Value = "FINISHED" Then
        Sheet_DataBase.Cells(rowNumber, "BR").Value = "Completed"
    Else
        Sheet_DataBase.Cells(rowNumber, "BR").Value = "Incompleted"
    End If
    
End Sub

Sub deleteExistingCheck(rowNumber As Integer)
    
    Sheet_DataBase.Rows(rowNumber).EntireRow.Delete

'    ' �������� ������� - ����������� ������, ������� ���������, ������� ������
'    Sheet_DataBase.Rows(rowNumber).Interior.Color = vbRed
    
End Sub

Sub deleteExistingDescription(rowNumber As Integer)
    
    Sheet_ErrDescr.Rows(rowNumber).EntireRow.Delete

'    ' �������� ������� - ����������� ������, ������� ���������, ������� ������
'    Sheet_ErrDescr.Rows(rowNumber).Interior.Color = vbRed
    
End Sub

Sub saveDescr()
    
    ' -----------------------------------------------
    '       ��������� �������� ������ � ������ IP
    ' -----------------------------------------------
    ' �������� ������� � ��������� ������
    Dim descrTableRange As Range
    
    ' ���������� ����� � ���������
    Dim descrRows As Integer
    
    ' �������� ������������ ������� � ��������� ������
    Set descrTableRange = Sheet_IP_Check.ListObjects("IpDescrTable").Range
    
    ' ���������� ����� � ��������
    descrRows = descrTableRange.Rows.Count
    
    ' ����� ��������� ����������� ������ �� ����� PERFORMER
    Dim dbDescrRowNumber As Integer
    dbDescrRowNumber = getDescrRow()
    
    ' ���� ������ ������ � ������� �� ������
    ' (��� �������������� ������ ������ ����� �� ���� PERFORMER)
    If descrTableRange(2, 1).Value <> "" Then
    
        For i = 2 To descrRows
            
            ' ���������� ��������
            Call saveDescrAttributes(dbDescrRowNumber + i - 2)
            
            ' ���������� ��� ������ � ��������
            Sheet_ErrDescr.Cells(dbDescrRowNumber + i - 2, "H").Value = descrTableRange(i, 1)
            Sheet_ErrDescr.Cells(dbDescrRowNumber + i - 2, "I").Value = descrTableRange(i, 2)
            
        Next i
    
    End If
    
    ' -----------------------------------------------
    '       ��������� �������� ������ � ������ PDM
    ' -----------------------------------------------
    
    ' �������� ������� � ��������� ������
    Dim pdmDescrTableRange As Range
    
    ' ���������� ����� � ���������
    Dim pdmDescrRows As Integer
    
    ' �������� ������������ ������� � ��������� ������
    Set pdmDescrTableRange = Sheet_PDM_Check.ListObjects("PdmDescrTable").Range
    
    ' ���������� ����� � ��������
    pdmDescrRows = pdmDescrTableRange.Rows.Count
    
    ' ����� ��������� ����������� ������ �� ����� PERFORMER
    dbDescrRowNumber = getDescrRow()
    
    If pdmDescrTableRange(2, 1).Value <> "" Then
    
        For i = 2 To pdmDescrRows
            
            ' ���������� ��������
            Call saveDescrAttributes(dbDescrRowNumber + i - 2)
            
            ' ���������� ��� ������ � ��������
            Sheet_ErrDescr.Cells(dbDescrRowNumber + i - 2, "H").Value = pdmDescrTableRange(i, 1)
            Sheet_ErrDescr.Cells(dbDescrRowNumber + i - 2, "I").Value = pdmDescrTableRange(i, 2)
            
        Next i
        
    End If
    
End Sub

Sub saveDescrAttributes(rowNumber As Integer)

    ' ��������� ����
    Sheet_ErrDescr.Cells(rowNumber, "A").Value = Sheet_IP_Check.Cells(1, "F")
    ' ��������� RelRecNr
    Sheet_ErrDescr.Cells(rowNumber, "B").Value = Sheet_IP_Check.Cells(2, "F")
    ' ��������� �����������
    Sheet_ErrDescr.Cells(rowNumber, "C").Value = Sheet_IP_Check.performerComboBox.Value
    ' ��������� IP Number
    Sheet_ErrDescr.Cells(rowNumber, "D").Value = Sheet_IP_Check.Cells(4, "F")
    ' ��������� ����� ������
    Sheet_ErrDescr.Cells(rowNumber, "E").Value = Sheet_IP_Check.Cells(5, "F")
    ' ��������� Rework
    Sheet_ErrDescr.Cells(rowNumber, "F").Value = Sheet_IP_Check.reworkComboBox.Value
    ' ��������� MESA status
    Sheet_ErrDescr.Cells(rowNumber, "G").Value = Sheet_IP_Check.mesaStatusComboBox.Value
    
End Sub

