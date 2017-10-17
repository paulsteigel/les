Attribute VB_Name = "mdl_App"
'gistThat@mcpher.com :do not modify this line - see ramblings.mcpher.com for details: updated on 10/17/2017 1:04:59 PM : from manifest: gist https://raw.githubusercontent.com/paulsteigel/NMPRP/master/mdl_App.bas
Option Explicit

Function GetAppVersion() As Long
    ' This is to get current version of the application
    ' Modify this line when a new version is comming
    GetAppVersion = 0
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


