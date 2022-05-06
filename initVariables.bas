Attribute VB_Name = "initVariables"
Sub InitFileName()
    
    If CLFileName = "" Then
        CLFileName = "CHECKLIST_v0-6.xlsm"
    End If

End Sub

Sub initLoadedCheckRowNumber()
    loadedCheckRowNum = 0
End Sub

Sub initLoadedDescrRowNumbers()
    Set loadedDescrRowNums = New Collection
End Sub
