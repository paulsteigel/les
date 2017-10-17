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

Sub SaveDataCurrentRow(ByVal Target As Range)
    ' Only save change if activesheet is II.5.B.1
    ShowOff False
    Dim dbRange As Range, iRow As Long, iCol As Long, i As Long
    
    Select Case ActiveSheet.Name
    Case "II.5.B.1":
        'push data to II.5.B
        If Val(Range("SEL_PLN_YEAR_CUR")) = 0 Then GoTo ExitSub
        'SEL_PLN_YEAR_COL
        iCol = Range("COL_YEARS_DIV_5B1").Column
        iRow = Target.Row
        
        ' Find the approriate column to save data
        Set dbRange = Sheets("II.5.B").Cells(Target.Row, Range("COL_YEARS_FUNDS").Column + 5 * (Range("SEL_PLN_YEAR_COL") - 1))
        For i = 0 To 4
            dbRange.Offset(0, i) = Target.Parent.Cells(iRow, iCol + i)
        Next
    End Select
    
ExitSub:
    Set dbRange = Nothing
    ShowOff True
End Sub

Sub DetectChange(ByVal Target As Range)
    Dim theSheet As Worksheet, XTerm As Long, curCol As Long, recCol As Long, baseRange As Range
    Set theSheet = ActiveSheet
    Select Case theSheet.Name
    Case "II.5.B":
        ' If a change in this table is recorded, the total colum of later years shall be also recorded

        curCol = Range("COL_YEARS_DIV").Column
        If Target.Column <= curCol Or Target.Column > curCol + Range("SEL_PLN_YEAR_OFFSET") + 1 Then Exit Sub

        recCol = Range("COL_YEARS_FUNDS").Column + 5 * (Target.Column - curCol - 1)
        Set baseRange = Sheets("II.5.B").Cells(Target.Row, recCol)
        
        If baseRange = "" Then baseRange = Target
    Case "II.5.B.1":
        ' If a change in this table was recorded, we will have to update in II.5.B
        curCol = Range("COL_YEARS_DIV_5B1").Column
        If Target.Column >= curCol And _
            Target.Column <= curCol + 5 Then
            If Target.Row = 2 Then
                ChangeDisplayYear
            ElseIf Target.Row > 6 And Target.Row < 556 Then
                'SheetChanged(ActiveSheet.Name) = True
                SaveDataCurrentRow Target
            End If
        End If
    Case "II.5.F":
        SheetChanged("II.5.F") = True
    Case Else:
    End Select
End Sub

Private Sub UpdateChange(ColID As Long, SheetName As String)
    ' This will do the update using rowID
    Dim StCol As Long, PrcCell As Range
    Set PrcCell = ActiveCell
    Select Case SheetName
    Case "II.5.B":
        ' The plan colum will be changed - other still kept unchanged
        ' Just inform user about changed they has made for this record and ask them to review change if neccessary
        ' COL_YEARS_FUNDS= start column
        StCol = Range("COL_YEARS_FUNDS").Column + 5 * ColID
        If PrcCell <> PrcCell.Offset(0, StCol) Then
            If MsgBox("MSG_DATA_CHANGED_REVISE", vbYesNo + vbQuestion) = vbYes Then
            Else
                ' if no, such change will be marked with red
            End If
        End If
    Case "II.5.B.1":
        
    End Select
End Sub

Sub ForceCalculate(Optional IndirectCall As Boolean = True)
    AppCalculation = True
    Application.Calculate ' force calculation first
    AppCalculation = False
End Sub

Private Sub ChangeDisplayYear()
    Dim OldYear As Long
    OldYear = Range("SEL_PLN_YEAR_CUR")
    Range("SEL_PLN_YEAR_CUR") = Range("SEL_PLN_YEAR")
    If Range("SEL_PLN_YEAR_CUR") = OldYear Then Exit Sub
    ' Now charge this table with data from the II.5.B table
    'COL_YEARS_FUNDS
    Dim StartCol As Long
    'COL_YEARS_DIV_5B1 - datapool for selected year
    ' Just identify which year to start extracting data
    ' There is following case...
    '1. At saving II.5.B, figures shall be updated in the details column and notice user on such change
    ' We have a Planed Column and ReviseTotal colum
    
    ForceCalculate
    ShowOff
    
    Dim idRow As Range
    Dim dbRange As Range, iRow As Long, iCol As Long, i As Long
    
    On Error GoTo ExitSub
    
    ' Now copy data to II.5.B.1
    If Val(Range("SEL_PLN_YEAR_CUR")) = 0 Then GoTo ExitSub
    ' Get last non blank cell

    ' Find the last blank activity cell in II.5.B
    Set idRow = Sheets("II.5.B").Cells(556, 3).End(xlUp)
    ' Then start on II.5.B.1
    Set idRow = Sheets("II.5.B.1").Cells(idRow.Row, idRow.Column)
    ' Now clear all data first in II.5.B.1
    Set dbRange = Range(Sheets("II.5.B.1").Cells(7, Range("COL_YEARS_DIV_5B1").Column), Sheets("II.5.B.1").Cells(555, Range("COL_YEARS_DIV_5B1").Column + 4))
    dbRange = ""
    While idRow.Row >= 7
        If idRow <> "" Then
            'only copy  nonblank cell
            'push data to II.5.B
            If Val(Range("SEL_PLN_YEAR_CUR")) = 0 Then GoTo ExitSub
            
            iCol = Range("COL_YEARS_DIV_5B1").Column
            iRow = idRow.Row
            
            ' Find the approriate column to save data
            ' can be quicker with copying all columns>>>
            Set dbRange = Sheets("II.5.B").Cells(idRow.Row, Range("COL_YEARS_FUNDS").Column + 5 * (Range("SEL_PLN_YEAR_COL") - 1))
            
            If idRow.Row = 9 Then
                Debug.Print dbRange.Address
            End If
            For i = 0 To 4
                idRow.Parent.Cells(iRow, iCol + i) = dbRange.Offset(0, i)
            Next
        End If
        Set idRow = idRow.Offset(-1)
    Wend
ExitSub:
    If Err.Number <> 0 Then
        Debug.Print Err.description
    End If
    Set idRow = Nothing
    Set dbRange = Nothing
    ShowOff True
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


