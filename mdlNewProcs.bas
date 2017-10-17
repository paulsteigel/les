Option Explicit
Global AppCalculation As Boolean

Sub LocateFirstCell()
    Dim lngRowNumber As Long, lngColNumber As Long
    Dim strColLetter As String
    Dim rngFreezeWindow As Range
    
    With ActiveWindow
        If .SplitRow > 0 And .SplitColumn > 0 Then
            lngRowNumber = .SplitRow + 1
            lngColNumber = .SplitColumn + 1
        Else
            Exit Sub
        End If
    End With
    
    'Code to convert a Column Number to a Column String has been adapted from: _
    http://www.freevbcode.com/ShowCode.asp?ID=4303
    If lngColNumber > 26 Then
        strColLetter = Chr(Int((lngColNumber - 1) / 26) + 64) & Chr(((lngColNumber - 1) Mod 26) + 65)
    Else
        'Columns A-Z
        strColLetter = Chr(lngColNumber + 64)
    End If
    
    Set rngFreezeWindow = Range(strColLetter & lngRowNumber)
    rngFreezeWindow.Select
    rngFreezeWindow.Activate
    Set rngFreezeWindow = Nothing
End Sub

Property Let SheetChanged(SheetName As String, NewValue As Boolean)
    ' Record Sheet change actions and it will be reset at sheet closing
    'CONF_SHEET_CHANGE
    Dim SrcRng As Range, CellFound As Boolean
    Set SrcRng = Range("CONF_SHEET_CHANGE").Offset(1)
    While SrcRng <> "" And Not CellFound
        If SrcRng = SheetName Then
            CellFound = True
        Else
            Set SrcRng = SrcRng.Offset(1)
        End If
    Wend
    ' Just write the change here
    ShowOff False
    SrcRng = SheetName
    ShowOff True
    Set SrcRng = Nothing
End Property

Property Get SheetChanged(SheetName As String) As Boolean
    ' Record Sheet change actions and it will be reset at sheet closing
    Dim SrcRng As Range
    Set SrcRng = Range("CONF_SHEET_CHANGE").Offset(1)
    While SrcRng <> "" And Not SheetChanged
        If SrcRng = SheetName Then
            SheetChanged = True
        Else
            Set SrcRng = SrcRng.Offset(1)
        End If
    Wend
    Set SrcRng = Nothing
End Property

Sub ResetChanges()
    ' Reset all changes
    Dim SrcRng As Range, CellFound As Boolean
    Set SrcRng = Range("CONF_SHEET_CHANGE").Offset(1)
    ShowOff False
    While SrcRng <> ""
        SrcRng = ""
        Set SrcRng = SrcRng.Offset(1)
    Wend
    Set SrcRng = Nothing
    ShowOff True
End Sub

Sub ForceCalculate(Optional IndirectCall As Boolean = True)
    AppCalculation = True
    Application.Calculate ' force calculation first
    AppCalculation = False
End Sub

Sub CleanName()
    Dim tName As Name, xPos As Long
    For Each tName In ThisWorkbook.Names
        xPos = InStr(tName.RefersTo, "]")
        If xPos > 0 Then
            tName.RefersTo = Mid(tName.RefersTo, xPos + 1)
        End If
    Next
End Sub




