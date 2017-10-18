'gistThat@mcpher.com :do not modify this line - see ramblings.mcpher.com for details: updated on 10/18/2017 1:26:29 AM : from manifest: gist https://raw.githubusercontent.com/paulsteigel/les/master/mdlAppLoad.bas
Option Explicit

Function GetAppVersion() As Long
    ' This is to get current version of the application
    ' Modify this line when a new version is comming
    GetAppVersion = 2
End Function

Sub Patch()
    MsgBox "Hi there1... this is a test for auto update...", vbInformation
    ShowOff
    Dim tSheet As Worksheet
    Set tSheet = ThisWorkbook.Sheets("household")
    With tSheet
        .Range("W:W").NumberFormat = "General"
        .Range("EI:EJ").NumberFormat = "General"
    End With
    Set tSheet = Nothing
    SetValidation
    ShowOff True
End Sub

Private Sub SetValidation()
    Dim Ptr As Range, W1 As String, W1txt As String, E1 As String, E1txt As String
    W1 = MSG("MSG_WARNING"): W1txt = MSG("MSG_NUMBER_ONLY"): E1 = MSG("MSG_ERROR"): E1txt = MSG("MSG_ERROR_WARN")
    Set Ptr = Range("tblFormInfor").Offset(0, 1)
    ProtectSheet False
    While Ptr <> ""
        If Ptr.Offset(-1) = "INTEGER" Or Ptr.Offset(-1) = "SINGLE" Then
            If IsRangeValid(CStr(Ptr.Value)) Then
                If Not Range(Ptr).Locked Then
                    ApplyValidation Range(Ptr), W1, W1txt, E1, E1txt
                End If
            End If
        End If
        Set Ptr = Ptr.Offset(0, 1)
    Wend
    ProtectSheet True
End Sub

Private Sub ProtectSheet(Prm As Boolean)
    Dim tsh As Worksheet
    On Error Resume Next
    For Each tsh In ThisWorkbook.Sheets
        If tsh.Name Like "Part*" Or tsh.Name = "General" Or tsh.Name = "Ranking" Then
            If Prm Then
                tsh.Protect
            Else
                tsh.Unprotect
            End If
        End If
    Next
    Set tsh = Nothing
End Sub

Private Sub ApplyValidation(tCell As Range, Vl1 As String, vl2 As String, vl3 As String, vl4 As String)
    With tCell.Validation
        .Delete
        .Add Type:=xlValidateDecimal, AlertStyle:=xlValidAlertStop, Operator _
        :=xlBetween, Formula1:="0", Formula2:="10000000"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = Vl1
        .ErrorTitle = vl3
        .InputMessage = vl2
        .ErrorMessage = vl4
        .ShowInput = True
        .ShowError = True
    End With
End Sub
