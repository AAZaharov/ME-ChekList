Attribute VB_Name = "GetDataModule"
' ����� RelRecNr � ���� ������
Function getRelRecNr() As Integer

    ' ���� � ���� ��� ��� �������, ���������� 0 � ������� �� �������
    If Sheet_DataBase.Cells(3, "B") = "" Then
        getRelRecNr = 0
        Exit Function
    End If
    
    ' ������ ��� ����������� ������� RelRecNr � ����� DataBase
    Dim rrNumbers() As Variant
    ' ��������� ����������� ������
    Dim lastRrnRow As Integer
    lastRrnRow = Sheet_DataBase.Cells(Rows.Count, "B").End(xlUp).Row
    
    Dim rrnR As Range
    Set rrnR = Range(Sheet_DataBase.Cells(3, "A"), Sheet_DataBase.Cells(lastRrnRow, "B"))
    rrNumbers = rrnR.Value
    
    ' RelRecNr
    Dim rrn As String
    rrn = Sheet_IP_Check.Cells(2, "F").Value
    
    ' ���� ��������� RelRecNr � ������� � ���������� ����� ������
    For i = 3 To lastRrnRow
        If rrNumbers(i - 2, 2) = rrn Then
            getRelRecNr = i
        End If
    Next i
    
End Function

Function getFinishedRRN()

    ' ������ ��� ����������� ������� RelRecNr � ����� DataBase
    Dim rrNumbers() As Variant
    ' ��������� ����������� ������
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
    
    ' ���� ��������� RelRecNr � ������� � ���������� ����� ������
    For i = 3 To lastRrnRow
        If rrNumbers(i - 2, 1) = rrn And rrNumbers(i - 2, 5) = rework Then
            getFinishedRRN = i
        End If
    Next i
    
End Function

Function getEqualRework()

    ' ������ ��� ����������� ������� RelRecNr � ����� DataBase
    Dim rrNumbers() As Variant
    ' ��������� ����������� ������
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
    
    ' ���� ��������� RelRecNr � ������� � ���������� ����� ������
    For i = 3 To lastRrnRow
        If rrNumbers(i - 2, 1) = rrn And rrNumbers(i - 2, 5) = rework Then
            getEqualRework = i
        End If
    Next i
    
End Function

Function getLastRework()

    ' ������ ��� ����������� ������� RelRecNr � ����� DataBase
    Dim rrNumbers() As Variant
    ' ��������� ����������� ������
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
    
    ' ��������� ������ � ����������� RelRecNr � IP Number,
    ' ������� � ���������� ��������� ����� Rework
    For i = 3 To lastRrnRow
    
        If rrNumbers(i - 2, 1) = rrn And rrNumbers(i - 2, 3) = ipn _
            And rrNumbers(i - 2, 5) > lastRework Then
            
            lastRework = rrNumbers(i - 2, 5)
            
        End If
        
    Next i
    
    getLastRework = lastRework
    
End Function

Function getAllReworks() As Collection

    ' ������ ��� ����������� ������� RelRecNr � ����� DataBase
    Dim rrNumbers() As Variant
    ' ��������� ����������� ������
    Dim lastRrnRow As Integer
    lastRrnRow = Sheet_DataBase.Cells(Rows.Count, "B").End(xlUp).Row
    
    Dim rrnR As Range
    Set rrnR = Range(Sheet_DataBase.Cells(3, "B"), Sheet_DataBase.Cells(lastRrnRow, "F"))
    rrNumbers = rrnR.Value
    
    ' ��������� ��� �������� ������� Rework ��� ��������� RelRecNr � IP Number
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
    
    ' ��������� ������ � ����������� RelRecNr � IP Number,
    ' ���������� ��������� ������ Rework � ���������
    For i = 3 To lastRrnRow
    
        If rrNumbers(i - 2, 1) = rrn And rrNumbers(i - 2, 3) = ipn Then
        
            reworkColl.Add (rrNumbers(i - 2, 5))
            
        End If
        
    Next i
    
    Set getAllReworks = reworkColl
    
End Function

Function getPerformer() As Integer

    ' ����� ����������� � ���� ������
    ' ������ ��� ����������� ������������ � ����� Send_Email
    Dim performers() As Variant
    ' ��������� ����������� ������ � ������ ������������
    Dim lastPerfRow As Integer
    lastPerfRow = Sheet_SendEmail.Cells(Rows.Count, "A").End(xlUp).Row
    
    Dim pr As Range
    Set pr = Range(Sheet_SendEmail.Cells(1, "A"), Sheet_SendEmail.Cells(lastPerfRow, "B"))
    performers = pr.Value
    
    ' �����������
    Dim performer As String
    performer = Sheet_IP_Check.performerComboBox.Value
    
    ' ���� ���������� ����������� � ������� � ���������� ����� ������
    For i = 1 To lastPerfRow
        If performers(i, 1) = performer Then
            getPerformer = i
        End If
    Next i

End Function

Function getUpdatedRow() As Integer
    
    ' ������ ��� ����������� ������� RelRecNr � ����� DataBase
    Dim rrNumbers() As Variant
    ' ��������� ����������� ������
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
    
    
    ' ���� ���������� ������ � ������� � ���������� ����� ������
    For i = 3 To lastRrnRow
        If rrNumbers(i - 2, 1) = rrn _
            And rrNumbers(i - 2, 3) = ipNum _
            And rrNumbers(i - 2, 5) = rework Then
            
            getUpdatedRow = i
            
        End If
    Next i
    
End Function

Function getUpdatedDescrRows() As Collection
    
    ' ��������� ��� ������ ������� ����� � ��������� ������
    Dim errDescrColl As Collection
    Set errDescrColl = New Collection
    
    ' ������ ��� ����������� ������� RelRecNr � ����� PERFORMER
    Dim rrNumbers() As Variant
    ' ��������� ����������� ������
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
    
    
    ' ���� ���������� ������ � ������� � ���������� � ���������
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
    
    ' ��������� ����������� ������ � ����, ���������� �� ������� "B" (RelRecNr) �� ����� DataBase
    Dim lastDbRow As Integer
    lastDbRow = Sheet_DataBase.Cells(Rows.Count, "B").End(xlUp).Row
    
    ' ���� � ���� ��� ��� �������, ���������� 3,
    ' ����� ���������� ������ ������ ������
    If Sheet_DataBase.Cells(3, "B") = "" Then
        getRow = 3
    Else
        getRow = lastDbRow + 1
    End If
    
End Function

Function getDescrRow()
    
    ' ��������� ����������� ������ � ���� �������� ������, ���������� �� ������� "B" (RelRecNr) �� ����� PERFORMER
    Dim lastDescrDbRow As Integer
    lastDescrDbRow = Sheet_ErrDescr.Cells(Rows.Count, "B").End(xlUp).Row
    
    ' ���� � ���� ��� ��� �������, ���������� 3,
    ' ����� ���������� ������ ������ ������
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

