Attribute VB_Name = "ValidationModule"
' ���� �������� ������� ����������� � ������ (���� SendEmail)
Dim isPerfNotExist As Boolean
' ���� �������� ������
Dim isSendEmail As Boolean

' ���� ��� �������� ������������� ������ ����� �������� ������
' ������������, ���� �� ������� ����� "��������� ��� �������� ������"
Dim isEmptyDescrExist As Boolean

' ���� ��� ��������, ��������� �� ������������ ���� -
' RelRecNr, Performer � IP Number
Dim isWrongAttrExist As Boolean

' ���� ��� �������� ������������ ���������� ���� RelRecNr
Dim isRelRecNrValid As Boolean

' ��������� �� ������ ���������� ���� ����
Dim dateErrMsg As String
' ��������� �� ������ ���������� ���� RelRecNr
Dim rrnErrMsg As String
' ��������� �� ������ ���������� ���� Performer
Dim perfErrMsg As String
' ��������� �� ������ ���������� ���� IP Number
Dim ipNumErrMsg As String
' ��� ������
Dim wrongAttrMsg As String

' ���� ������������ ������
Dim isRecordValid As Boolean
    

Function beforeSaveValidation() As Boolean
    ' ���������, ������ ��� ��� ������ ������ ������,
    ' � �������� ��������������� ������� ��� �������� ������
    If Sheet_IP_Check.saveRecordToggleButton.Value Then ' ���� ������ ������, ������ �������� ��� �������������� ������
        isRecordValid = updSaveRecValidation
    Else ' ����� ������ �������� ��� ���������� ������
        isRecordValid = addRecSaveValidation()
    End If
    
    ' ���� � ������ ������� ������, ��������� ��������
    If isRecordValid = False Then
        beforeSaveValidation = False
        Exit Function
    End If
    
    ' ���� �� ������� ����� "��������� ��� �������� ������",
    ' ��������� ������� �������� � ������� ����
    If Sheet_IP_Check.saveWithoutDescrCheckBox.Value = False Then
        isEmptyDescrExist = emptyDescrValidation()
    Else
        isEmptyDescrExist = False
    End If
    
    ' ���� �������� �����������, ���������� ���������
    ' � ������� �� ��������� ��������
    If isEmptyDescrExist Then
        MsgBox Prompt:="����������� �������� ������. ���������� ��������", _
               Title:="������ ��������", _
               Buttons:=vbExclamation
        beforeSaveValidation = False
        Exit Function
    End If
    
    ' ���� �������� ����� �������� ������, ���������
    ' ������� ����������� � ������ � ��������� ��������
    If Sheet_IP_Check.sendErrDescrCheckBox.Value _
        Or Sheet_IP_Check.sendFinishedStateCheckBox.Value Then
        isPerfNotExist = perfNotExistValidation()
        isSendEmail = True
    Else
        isSendEmail = False
    End If
    
    ' ���� ����������� ��� � ������, � ����� �������� ������ �������,
    ' ���������� ���������� ���� � ������� �������� ��������
    If isPerfNotExist And isSendEmail Then
        ' ���������� ���������� ���� � ���������� �����
        noPerformerAction = MsgBox(Prompt:="���������� ��������� ������ - " & vbNewLine _
                                         & "���������� ����������� ��� � ������" & vbNewLine _
                                         & "��������� ��� �������� ������?", _
                                    Title:="������ ��������", _
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
    
    ' ���� ������� ���������� �� �����, ������, ������ ���
    beforeSaveValidation = True
    
End Function

Function emptyDescrValidation()

    ' fill description zone
    ' find first and last rows with errors description
    Dim firstDescrRow As Long, lastDescrRow As Long
    ' ������ ������ ������� IpDescrTable
    firstDescrRow = Sheet_IP_Check.ListObjects("IpDescrTable").DataBodyRange.Row
    ' ��������� ����������� ������ (� ������� �������) � ������� ��������
    lastDescrRow = Sheet_IP_Check.Cells(Rows.Count, "J").End(xlUp).Row
    
    ' ������������, ��� ������ ���
    emptyDescrValidation = False
    
    ' ���� ���� ������ (������ ������ ������� �� ������)
    ' �� ��������� ������� ��������
    If Sheet_IP_Check.Cells(firstDescrRow, "J").Value <> "" Then
        
        ' ��������� ������� ��������
        For i = firstDescrRow To lastDescrRow
            If Sheet_IP_Check.Cells(i, "K").Value = "" Then
                ' ���� ������� ������ ����, ������� �� �������
                emptyDescrValidation = True
                Exit Function
            End If
        Next i

    End If
    
    ' �� �� ����� ��� ������ PDM
    ' ������ ������ ������� PdmDescrTable
    firstDescrRow = Sheet_PDM_Check.ListObjects("PdmDescrTable").DataBodyRange.Row
    ' ��������� ����������� ������ (� ������� �������) � ������� ��������
    lastDescrRow = Sheet_PDM_Check.Cells(Rows.Count, "J").End(xlUp).Row
    
    If Sheet_PDM_Check.Cells(firstDescrRow, "J").Value <> "" Then
        
        For i = firstDescrRow To lastDescrRow
            If Sheet_PDM_Check.Cells(i, "K").Value = "" Then
                ' ���� ������� ������ ����, ������� �� �������
                emptyDescrValidation = True
                Exit Function
            End If
        Next i
    
    End If

End Function

Function perfNotExistValidation() As Boolean

    ' ����� ������ � �������� �����������
    Dim selectedPerfRow As Integer
    
    selectedPerfRow = getPerformer()
    
    If selectedPerfRow = 0 Then
        perfNotExistValidation = True
    Else
        perfNotExistValidation = False
    End If
    
End Function

Function updSaveRecValidation() As Boolean
    
    ' �������� ��������� �� ������������ ������ � ���� ������
    ' �� ������ RelRecNr ������� ������ � ���� ������
    Dim updatedRow As Integer
    updatedRow = getUpdatedRow()
    
    ' ���� ����� ������ ����������, ���������� �������� True, ����� False
    If updatedRow > 0 Then
        updSaveRecValidation = True
    Else
        MsgBox ("� ���� ��� ������ � ������ RelRecNr, Rework � IP Number" & vbNewLine _
              & "���������� ��������.")
        updSaveRecValidation = False
    End If
    
End Function

Function addRecSaveValidation() As Boolean
    ' �������� ���������
    isWrongAttrExist = wrongAttrCheck()
    
    If isWrongAttrExist Then
        addRecSaveValidation = False
        Exit Function
    End If
    
    ' �������� ����� RelRecNr � Rework
    isRelRecNrValid = relRecNrCheck()
    
    If isRelRecNrValid = False Then
        addRecSaveValidation = False
        Exit Function
    End If
    
    addRecSaveValidation = True
    
End Function

Function wrongAttrCheck() As Boolean
    
    wrongAttrMsg = ""
    
    ' �������� ����
    Call dateFieldCheck
    ' �������� ���� RelRecNr
    Call rrnFieldCheck
    ' �������� ���� Performer
    Call perfFieldCheck
    ' �������� ���� Performer
    Call ipNumFieldCheck
    
    Prompt = "������� ������:" & vbNewLine _
                                & vbNewLine
    
    wrongAttrMsg = dateErrMsg _
                    & rrnErrMsg _
                    & perfErrMsg _
                    & ipNumErrMsg
    
    If wrongAttrMsg = "" Then
        wrongAttrCheck = False
    Else
        wrongAttrCheck = True
        MsgBox Prompt:=Prompt & wrongAttrMsg, Title:="�������� ������", Buttons:=vbExclamation
    End If
    
End Function

Sub dateFieldCheck()
    dateErrMsg = ""
    ' ��������, ��������� ���� ����
    If Sheet_IP_Check.Cells(1, "F") = "" Then
        dateErrMsg = "�� ��������� ���� ����." & vbNewLine
    End If
    ' ��� �������� ����� ������ ���� ������ ���� �� ������ �������
    If Sheet_IP_Check.Cells(1, "F") <> "" And CDate(Sheet_IP_Check.Cells(1, "F")) < Date Then
        dateErrMsg = "��������� ���� ��� ������. ������� ���������� ����." & vbNewLine
    End If
End Sub

Sub rrnFieldCheck()
    rrnErrMsg = ""
    ' ��������, ��������� �� ���� RelRecNr
    If Sheet_IP_Check.Cells(2, "F") = "" Then
        rrnErrMsg = "�� ��������� ���� RelRecNr." & vbNewLine
    End If
End Sub

Sub perfFieldCheck()
    perfErrMsg = ""
    ' ��������, ��������� �� ���� Performer
    If Sheet_IP_Check.performerComboBox.Value = "" Then
        perfErrMsg = "�� ��������� ���� Performer." & vbNewLine
    End If
End Sub

Sub ipNumFieldCheck()
    ipNumErrMsg = ""
    ' ��������, ��������� �� ���� IP Number
    If Sheet_IP_Check.Cells(4, "F") = "" Then
        ipNumErrMsg = "�� ��������� ���� IP Number." & vbNewLine
    End If
End Sub

Function relRecNrCheck() As Boolean
    ' ���� RelRecNr �����, �� ��������� ���� Rework, ��� ������ ���� ����� 0
    If getRelRecNr() = 0 Then
        
        If Sheet_IP_Check.reworkComboBox.Value <> "0" _
        And Sheet_IP_Check.reworkComboBox.Value <> "FINISHED" Then
            
            If MsgBox(Prompt:="��� ������ RelRecNr ���� Rework" & vbNewLine _
                    & "������ ���� ����� 0." & vbNewLine _
                    & "���������� Rework � 0 � ����������?", Buttons:=vbOKCancel) = vbOK Then
                Sheet_IP_Check.reworkComboBox.Value = 0
            Else
                relRecNrCheck = False
                Exit Function
            End If
        
        End If
    
    End If
    
    ' ���� ������ � ����� RelRecNr ����������, ��
    If getRelRecNr() > 0 Then
    
        ' ��������� ���� Rework �� �������� "FINISHED"
        If getFinishedRRN() > 0 Then
        
            MsgBox ("������ � ����� RelRecNr ��� ���������." & vbNewLine _
                 & "(Rework = FINISHED)" & vbNewLine _
                 & "���������� ��������.")
                 
            relRecNrCheck = False
            Exit Function
            
        End If
        
        ' ��������� ������� ������ � ����� �� ����� Rework
        ' � ���������� ��������� �����
        If getEqualRework() > 0 Then
            
            ' ������� � ���� ������ Rework ��� ��������� RelRecNr � IP Number
            Dim allRework As Collection
            Set allRework = getAllReworks
            
            Dim reworksMsg As String
            reworksMsg = allRework.Item(1)
            
            For i = 2 To allRework.Count
                reworksMsg = reworksMsg & ", " & allRework.Item(i)
            Next i
            
            If allRework.Count = 1 Then
                
                If MsgBox(Prompt:="������ � Rework = " & getLastRework() & " ��� ���������� � ����." & vbNewLine _
                        & vbNewLine _
                        & "���������� Rework � " & getLastRework() + 1 & " � ����������?", Buttons:=vbOKCancel) = vbOK Then
                        
                    Sheet_IP_Check.reworkComboBox.Value = getLastRework() + 1
                Else
                    relRecNrCheck = False
                    Exit Function
                End If
            
            Else
                
                If MsgBox(Prompt:="������ � Rework = " & reworksMsg & " ��� ���������� � ����." & vbNewLine _
                        & vbNewLine _
                        & "���������� Rework � " & getLastRework() + 1 & " � ����������?", Buttons:=vbOKCancel) = vbOK Then
                        
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
