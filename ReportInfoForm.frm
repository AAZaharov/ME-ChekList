VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ReportInfoForm 
   Caption         =   "Saved Report Info"
   ClientHeight    =   1965
   ClientLeft      =   50
   ClientTop       =   380
   ClientWidth     =   3660
   OleObjectBlob   =   "ReportInfoForm.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ReportInfoForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub okButton_Click()
    
    Sheet_IP_Check.Activate
    
    Unload ReportInfoForm
    
End Sub

Private Sub showDescrRowsButton_Click()
    
    Sheet_ErrDescr.Activate
    
    Dim fRow As Integer
    Dim lRow As Integer
    
    pos = InStr(1, descrRowsLabel.Caption, " - ")
    
    fRow = Left(descrRowsLabel.Caption, pos - 1)
    lRow = Mid(descrRowsLabel.Caption, pos + 3)
    
    Sheet_ErrDescr.Range("A" & fRow & ":I" & lRow).Select
    
End Sub

Private Sub showRepRowButton_Click()
    
    Sheet_DataBase.Activate
    
    Dim repRow As Integer
    repRow = CInt(repRowLabel.Caption)
    
    Sheet_DataBase.Range("A" & repRow & ":BT" & repRow).Select
    
End Sub

Private Sub UserForm_Initialize()
    
    Call setRepRowLabel
    
    Call setDescrRowsLabel
    
    StartUpPosition = 0
    Left = Application.Left + (0.5 * Application.Width) - 0.5 * Width
    Top = Application.Top + (0.5 * Application.Height) - 0.5 * Height
    
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    
    Sheet_IP_Check.Activate
    
End Sub

Sub setRepRowLabel()
    
    repRowLabel.Caption = GetDataModule.getEqualRework()
    
End Sub

Sub setDescrRowsLabel()
    
    Dim repRow As Integer
    repRow = GetDataModule.getEqualRework()
    
    Dim allErr As Integer
    allErr = GetDataModule.getSumIpErrors(repRow) + GetDataModule.getSumPdmErrors(repRow)
    
    If allErr = 0 Then
    
        descrRowsLabel.Caption = "No errors"
        showDescrRowsButton.Enabled = False
        Exit Sub
        
    End If
    
    Dim descrRows As Collection
    Set descrRows = GetDataModule.getUpdatedDescrRows
    
    Dim fRow  As Integer
    fRow = descrRows.Item(1)
    
    Dim lRow As Integer
    lRow = descrRows(descrRows.Count)
    
    descrRowsLabel.Caption = fRow & " - " & lRow
    showDescrRowsButton.Enabled = True
    
End Sub
