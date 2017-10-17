Option Explicit

Function GetAppVersion() As Long
    ' This is to get current version of the application
    ' Modify this line when a new version is comming
    GetAppVersion = 0
    Patch
End Function

Private Sub Patch()
    Dim tSheet As Worksheet
    Set tSheet = ThisWorkbook.Sheets("household")
    With tSheet
        .Range("W:W").NumberFormat = "General"
        .Range("EI:EJ").NumberFormat = "General"
    End With
    Set tSheet = Nothing
End Sub
