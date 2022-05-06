Attribute VB_Name = "globalVariables"
' file name
Public CLFileName As String

' flag for saving without attribute (RelRecNr must be filled)
Public allowSaveWithoutAttr As Boolean

' flag for saving without error descriptions
Public allowSaveWithoutDescr As Boolean

' flag for sending e-mail with error descriptions
Public sendErrDescrEmail As Boolean

' flag for sending e-mail that plan is finished
Public sendFinishedStateEmail As Boolean

' row number of loaded record on Sheet "DataBase"
Public loadedCheckRowNum As Integer

' collection of row numbers for loaded descriptions on Sheet "PERFORMER"
Public loadedDescrRowNums As Collection


