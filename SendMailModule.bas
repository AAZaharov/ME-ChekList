Attribute VB_Name = "SendMailModule"
Sub sendMail(rowNum As Integer)
    
    ' ����� ����������
    Dim mailTo As String
    ' ����� ���������� �����
    Dim Email_CC As String
    
    ' �������� �����
    Dim inputValue As String
    
    ' �������� ������ IP
    Dim ipErrDescr As String
    
    ' �������� ������ PDM
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
        MsgBox "���������� ����������� ��� � ������." & vbNewLine _
        & vbNewLine _
        & "������ �� ����� ����������", vbExclamation
        Exit Sub
    End If
    
    ' ���� ������� �� ���������� ������, ������� �� �������
    If MsgBox("������������� ������: " & mailTo, vbYesNo, "������������� ��������") = vbNo Then
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
        ' ���������� ���� EMAIL STATUS
        Sheet_DataBase.Cells(rowNum, "BQ").Value = "Yes"
        
        ' ���������� ���������, ��� ������ ����������
        MsgBox ("������ ������� �������")
    Else
        MsgBox ("������")
    End If

End Sub

Sub sendFinishedMail(rowNum As Integer)

    ' �������� �����
    Dim inputValue As String
    
    ' ����� ����������
    Dim mailTo As String
    
    ' �������� ������ IP
    Dim ipErrDescr As String
    
    ' �������� ������ PDM
    Dim pdmErrDescr As String
    
    On Error Resume Next
    
    ' ��������, ���� �� ��������� ����������� � ������ � �������� ��. �����
    ' ������ ��� ����������� ������������
    Dim performers() As Variant
    ' ��������� ����������� ������
    Dim lastPerfRow As Integer
    lastPerfRow = Sheet_SendEmail.Cells(Rows.Count, "A").End(xlUp).Row
    
    Dim mailR As Range
    Set mailR = Range(Sheet_SendEmail.Cells(1, "A"), Sheet_SendEmail.Cells(lastPerfRow, "B"))
    performers = mailR.Value
    
    ' ��������� �����������
    Dim p As String
    p = Sheet_IP_Check.performerComboBox.Value
    
    ' ���� � ������� ������ ����� ��. �����
    For i = 1 To lastPerfRow
        If performers(i, 1) = p Then
            mailTo = performers(i, 2)
        End If
    Next i
    
    ' ���� ����� �� ������, ���������� ��������� � ��������� ���������
    If mailTo = "" Then
        MsgBox "���������� ����������� ��� � ������." & vbNewLine _
        & vbNewLine _
        & "������ �� ����� ����������", vbExclamation
        Exit Sub
    End If
    
    ' �������� �����
    For i = 1 To 7
        inputValue = inputValue & " " & Sheet_IP_Check.Cells(i, "E").Value & " : " & Sheet_IP_Check.Cells(i, "F").Value & vbNewLine
    Next i

    ' ����� ������
    Dim mailText As String
    
    If getSumIpErrors(rowNum) = 0 And getSumPdmErrors(rowNum) = 0 Then
        mailText = inputValue & vbNewLine _
                 & "������ �� �������. ������ ��������� � ���� �� �������� Completed."
    Else
        mailText = inputValue & vbNewLine _
                 & "������ ���������� �������. ������ ��������� � ���� �� �������� Completed."
    End If
    
    ipErrDescr = ipErrMailText
    
    pdmErrDescr = pdmErrMailText
    
    mailText = mailText & vbNewLine _
                & vbNewLine _
                & ipErrDescr & vbNewLine _
                & vbNewLine _
                & pdmErrDescr
    
    ' ���� ������
    Dim mailSubject As String
    mailSubject = "Checklist for " & Sheet_IP_Check.Cells(2, "F").Value & "," & " " & Sheet_IP_Check.Cells(4, "F").Value
            
    ' �������� ������� �������� ������ � ������� Outlook
    RES = SendEmailUsingOutlook(mailTo, mailText, mailSubject)
    
    If RES Then
        ' ���������� ���� EMAIL STATUS
        Sheet_DataBase.Cells(rowNum, "BQ").Value = "Yes"
        
        ' ���������� ���������, ��� ������ ����������
        MsgBox ("������ ���������� �������")
    Else
        MsgBox ("������ ��������")
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
        
        ' ������ ��� �������� (��� �������� �� ������������ ������ � ����������)
        ' � ��� ��� ��������, ������� ��������� ���� ���
        ' For i = 1 To 100000: DoEvents: Next

        Err.clear: .Display
        SendEmailUsingOutlook = Err = 0
    End With
    Set Outapp = Nothing
    
End Function

' -----------------------------------------------
'        �������� ������ � ������ IP
' -----------------------------------------------
Function ipErrMailText() As String

    ' ������ ��� �������� ������ � ������ IP
    Dim ipErrMailString As String
    ' �������� ������� � ��������� ������
    Dim descrTableRange As Range
    
    ' ���������� ����� � ���������
    Dim descrRows As Integer
    
    ' �������� ������������ ������� � ��������� ������
    Set descrTableRange = Sheet_IP_Check.ListObjects("IpDescrTable").Range
    
    ' ���������� ����� � ��������
    descrRows = descrTableRange.Rows.Count
    
    ' ���� � ������� ���� �������
    If descrTableRange(2, 1) <> "" Then
    
        ipErrMailString = "������ � ������ ADPP" & vbNewLine _
                      & "----------------------------" & vbNewLine _
                      & vbNewLine
        
        For i = 2 To descrRows
            
            ' ���������� ��� ������ � ��������
            ipErrMailString = ipErrMailString & "������ " & descrTableRange(i, 1) & ": " & descrTableRange(i, 2) & vbNewLine _
                                & vbNewLine
            
        Next i
        
    End If
    
    ipErrMailText = ipErrMailString
    
End Function

' -----------------------------------------------
'        �������� ������ � ������ PDM
' -----------------------------------------------
Function pdmErrMailText() As String

    ' ������ ��� �������� ������ � ������ PDM
    Dim pdmErrMailString As String
    
    ' �������� ������� � ��������� ������
    Dim pdmDescrTableRange As Range
    
    ' ���������� ����� � ���������
    Dim pdmDescrRows As Integer
    
    ' �������� ������������ ������� � ��������� ������ PDM
    Set pdmDescrTableRange = Sheet_PDM_Check.ListObjects("PdmDescrTable").Range
    
    ' ���������� ����� � ��������
    pdmDescrRows = pdmDescrTableRange.Rows.Count
    
    ' ���� � ������� ���� �������
    If pdmDescrTableRange(2, 1) <> "" Then
    
       pdmErrMailString = "������ � ������ PDM" & vbNewLine _
                      & "------------------------------" & vbNewLine _
                      & vbNewLine
       
       For i = 2 To pdmDescrRows
            
            ' ���������� ��� ������ � ��������
            pdmErrMailString = pdmErrMailString & "������ " & pdmDescrTableRange(i, 1) & ": " & pdmDescrTableRange(i, 2) & vbNewLine _
                                & vbNewLine
            
        Next i
    
    End If
    
    pdmErrMailText = pdmErrMailString
    
End Function


