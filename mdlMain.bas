Option Explicit

' For storing item attribute
Public Type ItemAttributes
    ItemDetails As String
    ItemHeading As String
    ItemEmphasizeHeading As String
    DataSource As String
    Name As String
End Type

' Counter vaiables for field access
Const AppPassWord = "d1ndh1sk" ' global password for protection

Private Sub ActivateSheet(SheetName As String)
    ThisWorkbook.Sheets(SheetName).Activate
End Sub

Function Nz(obj As Range, Optional NullValue As Variant) As Variant
    ' Like Access nz function
    On Error GoTo ErrHandler
    If obj = "" Then
        If NullValue = "" Then GoTo ErrHandler
        Nz = NullValue
    Else
        Nz = obj
    End If
    Exit Function
ErrHandler:
    Nz = ""
End Function

Private Sub SaveFile(FileName, DocObj As Object)
    On Error GoTo ErrHandler
    DocObj.Paragraphs(1).Range.Delete
    If Dir(FileName) <> "" Then Kill FileName
    DocObj.SaveAs FileName
ErrHandler:
    If Err.Number <> 0 Then
        MsgBox MSG("MSG_SAVE_FALSE"), vbCritical
    End If
End Sub

Private Sub InsertPara(DocObj As Object, ItemStyle As ItemAttributes, ItemText As String, Optional OverideAdd As Boolean = False)
    'On Error Resume Next
    Dim prCount As Long, tmpText As String, tmpItem As ItemAttributes
    tmpItem = ItemStyle
    With DocObj
        If ItemStyle.ItemHeading = "" Or ItemText = "" Then Exit Sub
        .Paragraphs.Add
        prCount = .Paragraphs.Count
        .Paragraphs(prCount).Range.Style = .Styles(ItemStyle.ItemHeading)
        .Paragraphs(prCount).Range.Text = ItemText
        
        If ItemStyle.ItemDetails <> "" And Not OverideAdd Then
            ' Add new introduction line if neccessary
            tmpItem.ItemHeading = tmpItem.ItemEmphasizeHeading
            tmpText = tmpItem.ItemDetails
            tmpItem.ItemDetails = ""
            InsertPara DocObj, tmpItem, tmpText
        End If
    End With
End Sub

Private Function GetFilteredData(iFilter As String, iColumn As String) As String
    'Base on the defined filter, try to get somedata from this - Don't care data range for
    Dim SrcArr As Variant, SrcRange As Range, i As Long, lRetStr  As String
    Dim OldText As String
    
    If Val(iColumn) <= 0 Then Exit Function
    Set SrcRange = Range("tblUnicode_1")
    i = 1
    While SrcRange.Cells(i, 1) <> ""
        If InStr(iFilter & "/", "/" & SrcRange.Cells(i, 1) & "/") <> 0 Then
            ' I found first ocurrence of the text
            OldText = SrcRange.Cells(i, 1)
            While SrcRange.Cells(i, 1) = OldText
                lRetStr = lRetStr & "//" & SrcRange.Cells(i, Val(iColumn))
                i = i + 1
            Wend
        Else
            i = i + 1
        End If
    Wend
    If lRetStr <> "" Then
        lRetStr = Replace(Mid(lRetStr, 3), "//", vbLf)
        GetFilteredData = lRetStr
    End If
End Function

Sub ActivateData()
    Sheets("Data").Activate
End Sub

Private Function GetLastCell(CellObj As Range) As Range
    While Len(Trim(CellObj)) > 0
        Set CellObj = CellObj.Offset(1)
    Wend
    Set GetLastCell = CellObj.Offset(-1)
End Function

Private Function FindColHeader(shtObj As Worksheet, FindRow As Long, FindTxt As String) As Long
    ' This function will return number of column with data specified in the Find text
    Dim FoundCell As Boolean, CellObj As Range, i As Long
    Set CellObj = shtObj.Cells(FindRow, 1)
    While i < 10 And Not FoundCell
        If Len(Trim(CellObj)) = 0 Then
            i = i + 1
        ElseIf CellObj = FindTxt Then
            FoundCell = True
        End If
        Set CellObj = CellObj.Offset(, 1)
    Wend
    If FoundCell Then FindColHeader = CellObj.Column - 1
End Function

Private Function ShrinkRange(rngIn As Range) As Range
    Dim LastCell As Range, tmpRange As Range
    Set LastCell = rngIn.Cells(rngIn.Rows.Count, 1)
    While Len(Trim(LastCell)) = 0
        Set LastCell = LastCell.Offset(-1)
    Wend
    Set tmpRange = rngIn.Range(rngIn.Cells(1, 1), LastCell)
    Set ShrinkRange = tmpRange
End Function

Private Function RangeValid(RangeName As String, shtObj As Worksheet) As Boolean
    Dim txtRange As Range
    On Error GoTo ErrHandler
    Set txtRange = shtObj.Range(RangeName)
    RangeValid = True
ErrHandler:
End Function

Function SheetValid(SheetName As String, WrbObj As Workbook) As Boolean
    Dim txtRange As Worksheet
    On Error GoTo ErrHandler
    Set txtRange = WrbObj.Sheets(SheetName)
    SheetValid = True
ErrHandler:
End Function

Private Sub CreateOrReizeName(OldName As String, NewNameCol As Long, Optional MoveRange As Boolean = True)
    ' This will create names base on the oldname offset new column
    Dim NameObj As Range
    Set NameObj = Range(OldName)
    With NameObj
        If MoveRange Then
            ThisWorkbook.Names.Add OldName, "=" & .Parent.Name & "!" & .Offset(0, NewNameCol).Address
        Else
            ' Expand the range
            ThisWorkbook.Names(OldName).RefersTo = ThisWorkbook.Names(OldName).RefersToRange.Resize(.Rows.Count, .Columns.Count + NewNameCol)
        End If
    End With
    Set NameObj = Nothing
End Sub

Function GetOpenWorkbook(FilePath As String) As Workbook
    'Open a workbook
    On Error GoTo ErrHandler
    Dim WrkBook As Workbook
    Set WrkBook = Application.Workbooks.Open(FilePath, False, True)
    Set GetOpenWorkbook = WrkBook
ErrHandler:
    Set WrkBook = Nothing
End Function

Function IsCollection(inCol As Object) As Boolean
    ' Check whether an object is a collection or not
    On Error GoTo ErrHandler
    IsCollection = IIf(inCol.Count > 0, True, False)
ErrHandler:
End Function

Property Let ProtectWorkbook(NewValue As Boolean)
    On Error Resume Next
    If NewValue Then
        ThisWorkbook.Protect AppPassWord
    Else
        ThisWorkbook.Unprotect AppPassWord
    End If
End Property

Property Let HideWorkSheet(SheetName As String, NewValue As Boolean)
    ' This will hide worksheet
    Dim tSheet As Worksheet
    If SheetValid(SheetName, ThisWorkbook) Then
        Set tSheet = Sheets(SheetName)
    Else
        Exit Property
    End If
    If NewValue Then
        tSheet.Visible = xlSheetVeryHidden
    Else
        tSheet.Visible = xlSheetVisible
    End If
End Property

