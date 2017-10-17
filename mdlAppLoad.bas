Option Explicit

Function GetAppVersion() As Long
    ' This is to get current version of the application
    ' Modify this line when a new version is comming
    GetAppVersion = 1
    'Call Patch("15_10_2014")
End Function

Sub Patch(PatchNumber As String)
'
' Setting format
'
'
    Select Case PatchNumber
    Case "15_10_2014"
        'UnProtectSheet Sheet10
        'Range("G9:G59").NumberFormat = "#,##0.0"
        'ProtectSheet Sheet10
    End Select
End Sub


