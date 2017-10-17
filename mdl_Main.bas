Attribute VB_Name = "mdl_Main"
Option Explicit

' For storing item attribute
Public Type ItemAttributes
    ItemDetails As String
    ItemHeading As String
    ItemEmphasizeHeading As String
    DataSource As String
    Name As String
End Type

Private Type TextObject
    TextCount As Long
    TextValue As String
End Type

Private Type ObjectEquation
    VariableName As String
    VariableFomular As String
End Type

Private OldSheet As Boolean

' Cached variable for keeping some temporary stuff
Private CachedListDistinct As Collection
Private ColListing() As New Collection
Private CurrentPointer As Long
Private OldTableName As String

' Counter vaiables for field access
Private proc_UseListOnly As Boolean
Private proc_ColumData As Long
Private NonCalculatedField As Boolean
Const AppPassWord = "d1ndh1sk" ' global password for protection

Sub Back2Main()
    ' For returning to Main Screen
    ActivateSheet "Manhinhchinh"
End Sub

Private Sub ActivateSheet(SheetName As String)
    ThisWorkbook.Sheets(SheetName).Activate
End Sub

Sub Act_II_2_A()
    If OldSheet Then
        ActivateSheet "II.2"
    Else
        ActivateSheet "II.2.A"
    End If
    OldSheet = Not OldSheet
End Sub

Sub Act_II_2_B()
    ActivateSheet "II.2.B"
End Sub

Sub Act_II_5_A()
    ActivateSheet "II.5.A"
End Sub

Sub Act_II_5_B()
    ActivateSheet "II.5.B"
End Sub

Sub Act_II_6_E()
    If Range("CONF_SCORE") <> 1 Then Exit Sub
    ActivateSheet "II.6.E"
    Range("COND_FLOOR").Activate
End Sub

Sub CriteriaEditor()
    ' Activate form for creating criteria
End Sub

Sub EvaluateActivity()
    ' Show form to conduct evaluation...
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

Sub SortTable(WrbObj As Workbook, WksObjName As String, RngName As String, SortKey1 As String, Optional SortKey2 As String)
    ' This procedure will sort the selected table using sortkey
    Dim theSheet As Worksheet
    Set theSheet = WrbObj.Sheets(WksObjName)
    ' unprotect the sheet first
    ProtectWorkSheet(theSheet) = False
    'Activate the sheet
    WrbObj.Worksheets(WksObjName).Activate
    WrbObj.Worksheets(WksObjName).Range(RngName).Sort Key1:=Range(SortKey1), Order1:=xlAscending, Key2:=Range(SortKey2) _
        , Order2:=xlAscending, Header:=xlNo, OrderCustom:=1, MatchCase:= _
        False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal, DataOption2:=xlSortNormal
    
    ' ReProtect the sheet
    ProtectWorkSheet(theSheet) = True
    Set theSheet = Nothing
End Sub

Sub ActivateData()
    Sheets("Data").Activate
End Sub

Sub ActivateMain()
    Sheets("Main").Activate
End Sub

Sub UpdateII2B()
    'Update this sheet
    ProtectWorkSheet(Sheets("II.2.B")) = False
    
    Dim theRange As Range, CellFirst As Range, CellLast As Range
    Dim fltStr As String, SkipFormat As String
    fltStr = Range("CAP_FLTR_STR")
    ' For list of columns to be skipped with formatting
    SkipFormat = Range("SKIP_FMT_II2B")
    
    ' Assign Cell Marker
    Set CellFirst = Sheets("II.2.B").Range("II2BFIRST")
    Set CellLast = Sheets("II.2.B").Range("II2BLAST")
    Set theRange = Sheets("II.2.B").Range(CellFirst.Offset(1), CellLast.Offset(-1))
    While CellFirst <> "" And CellFirst <> fltStr
        ' there are 3 things we have to copy...
        ' Format/ Validation and fomular
        CellLast.Copy
        
        ' Copy validation for all
        theRange.PasteSpecial xlPasteValidation
        ' Conidtionally copy fomular
        If CellLast.HasFormula Then theRange.PasteSpecial xlPasteFormulas
        ' Conidtionally copy Format
        If InStr(SkipFormat, CellFirst) = 0 Then theRange.PasteSpecial xlPasteFormats
        Application.CutCopyMode = False
        
        ' unlock the range
        theRange.Locked = False
        
        ' Increment 1 column
        Set CellFirst = CellFirst.Offset(0, 1)
        Set CellLast = CellLast.Offset(0, 1)
        Set theRange = theRange.Offset(0, 1)
    Wend
    Set CellFirst = Nothing
    Set CellLast = Nothing
    Set theRange = Nothing
    ProtectWorkSheet(Sheets("II.2.B")) = True
End Sub

Private Sub Repair_II5A(Optional SheetName As String)
    ' Unprotect sheets
    ProtectWorkSheet(ThisWorkbook.Sheets(SheetName)) = False
    With ThisWorkbook.Sheets(SheetName)
        .Range("A385:G385").Copy
        .Range("A6:G384").PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone
        .Range("A6:G384").Locked = False
        .Activate
        .Range("A6").Select
    End With
    Application.CutCopyMode = False
    ' Reprotect sheet
    ProtectWorkSheet(ThisWorkbook.Sheets(SheetName)) = True
End Sub

Private Sub Repair_II5B1()
    ' Repair II.5.B
    
    ProtectWorkSheet(ThisWorkbook.Sheets("II.5.B.1")) = False
    Dim SeedRow As Range
    Set SeedRow = Range("tblUnicode_2_1").Offset(Range("tblUnicode_2_1").Rows.Count).Resize(1)
    With ThisWorkbook.Sheets("II.5.B.1")
        .Range("J7").FormulaR1C1 = "=SUM(RC[1]:RC[4])"
        .Range("J7").Copy
        ' paste formular
        .Range("tblDataSumCol_1").PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone
        Application.CutCopyMode = False
        SeedRow.Copy
        ' paste format
        .Range("tblUnicode_2_1").PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone
        .Range("tblUnicode_2_1").PasteSpecial xlPasteValidation, Operation:=xlNone
        Application.CutCopyMode = False
        
        .Activate
        .Range("C7").Select
    End With
    Set SeedRow = Nothing
    ' Unlock some areas
    ThisWorkbook.Sheets("II.5.B.1").Range("tblUnicode_2_1").Locked = False
    ' Reprotect the sheet
    ProtectWorkSheet(ThisWorkbook.Sheets("II.5.B.1")) = True
End Sub

Private Sub Repair_II5B()
    ' Repair II.5.B
    
    ProtectWorkSheet(ThisWorkbook.Sheets("II.5.B")) = False
    Dim SeedRow As Range
    Set SeedRow = Range("tblUnicode_2").Offset(Range("tblUnicode_2").Rows.Count).Resize(1)
    
    With ThisWorkbook.Sheets("II.5.B")
        .Range("J7").FormulaR1C1 = "=SUM(RC[1]:RC[" & (Range("FIG_END_YEAR") - Range("FIG_STR_YEAR") + 1) & "])"
        .Range("J7").Copy
        ' paste formular
        .Range("tblDataSumCol").PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone
        Application.CutCopyMode = False
        .Range("B556:S556").Copy
        ' paste format
        .Range("tblUnicode_2").PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone
        .Range("tblUnicode_2").PasteSpecial xlPasteValidation, Operation:=xlNone
        Application.CutCopyMode = False
        
        .Activate
        .Range("C7").Select
    End With
    Set SeedRow = Nothing
    
    ' Unlock some areas
    ThisWorkbook.Sheets("II.5.B").Range("tblUnicode_2").Locked = False
    ' Reprotect the sheet
    ProtectWorkSheet(ThisWorkbook.Sheets("II.5.B")) = True
End Sub

Sub RepairSheet(Optional SheetObj As String = "")
    ' This procedure shall repare all sheet.
    ShowOff
    If SheetObj = "" Then
        Repair_II5A "II.5.A"
        Repair_II5A "II.5.C"
        Repair_II5B
        ' update II.2.B
        'UpdateII2B
    Else
        Select Case SheetObj
        Case "II.5.A", "II.5.C":
            Repair_II5A SheetObj
        Case "II.5.B":
            Repair_II5B
        Case "II.5.B.1":
            Repair_II5B1
        Case "II.2.B":
            'UpdateII2B
        Case Else
        End Select
    End If
    ' Make all sheet printable in one single page
    ThisWorkbook.Sheets(SheetObj).PageSetup.FitToPagesWide = 1
    'Sheet11.Activate
    ShowOff True
End Sub

Sub ApplySheetFilter()
    'Activate filter on selected sheets
    ApplyFilter ThisWorkbook.Sheets("II.6.A"), "A7", 3, "<>"
    ApplyFilter ThisWorkbook.Sheets("II.6.B"), "A7", 3, "<>"
    'ApplyFilter ThisWorkbook.Sheets("II.6.C"), "B5", 1, "<>"
    ApplyFilter ThisWorkbook.Sheets("II.2.B"), "I4", 1, "Có"
End Sub

Private Sub ApplyFilter(SheetObj As Worksheet, AppliedRange As String, FieldNum As Long, Criteria1 As String)
    ProtectWorkSheet(SheetObj) = False
    SheetObj.Range(AppliedRange).AutoFilter field:=FieldNum, Criteria1:=Criteria1
    ProtectWorkSheet(SheetObj) = True
End Sub

Sub QuickFilter()
    Dim FldCriteria As String, FldNum As Long
    ProtectWorkSheet(ActiveSheet) = False
    FldCriteria = "<>"
    Select Case ActiveSheet.Name
    Case "II.2.A":
        FldNum = 1
        FldCriteria = "Có"
    Case "II.2.B":
        FldNum = 1
        FldCriteria = "Có"
    Case "II.6.A", "II.6.B", "II.6.A.1", "II.6.B.1":
        FldNum = 3
    Case "II.6.C.1", "II.6.D.1", "II.6.F.1":
        FldNum = 1
    Case "II.5.A", "II.5.C", "II.5.D":
        FldNum = 1
    Case "II.5.B":
        FldNum = Range("II5BSTATUS").Column
        ActiveSheet.Range(ActiveSheet.Name & "!_FilterDatabase").AutoFilter field:=FldNum, _
            Criteria1:="=" & MSG("MSG_ST_NOTOK"), Operator:=xlOr, Criteria2:="=" & MSG("MSG_ST_VERIFY")
    Case "II.5.B.1":
        ' Filter just nonacceptable stuff
        FldNum = Range("II5B1STATUS").Column
        ActiveSheet.Range(ActiveSheet.Name & "!_FilterDatabase").AutoFilter field:=FldNum, _
            Criteria1:="=" & MSG("MSG_ST_NOTOK"), Operator:=xlOr, Criteria2:="=" & MSG("MSG_ST_VERIFY")
        GoTo ExitSub
    Case Else
        GoTo ExitSub
    End Select
    ActiveSheet.Range(ActiveSheet.Name & "!_FilterDatabase").AutoFilter field:=FldNum, Criteria1:=FldCriteria
ExitSub:
    ProtectWorkSheet(ActiveSheet) = True
End Sub

Sub ReleaseSheetFilter()
    ShowAll ThisWorkbook.Sheets("II.5.A")
    ShowAll ThisWorkbook.Sheets("II.5.B")
    ShowAll ThisWorkbook.Sheets("II.6.A")
    ShowAll ThisWorkbook.Sheets("II.6.B")
    ShowAll ThisWorkbook.Sheets("II.6.D")
    ShowAll ThisWorkbook.Sheets("II.6.C")
    ShowAll ThisWorkbook.Sheets("II.2.B")
End Sub

Sub ShowOff(Optional TurnEventOn As Boolean = False)
    ' Turn off everything, toggle
    ' avoid double calculation...
    If TurnEventOn And Application.ScreenUpdating = True Then Exit Sub
    Application.StatusBar = ""
    Application.ScreenUpdating = TurnEventOn
    Application.EnableEvents = TurnEventOn
    Application.CutCopyMode = False
    If TurnEventOn Then
        Application.Calculation = xlCalculationAutomatic
    Else
        Application.Calculation = xlCalculationManual
    End If
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

Private Sub ParseRange(frBook As Workbook, toBook As Workbook, shtName As String, RngName As String, Optional NeedUnprotect As Boolean = False)
    Dim RngArr As Variant, i As Long
    ' Revised by Ngoc on May 7 2014
    If NeedUnprotect Then ProtectWorkSheet(toBook.Sheets(shtName)) = False
    RngArr = Split(RngName, ",")
    For i = 0 To UBound(RngArr)
        toBook.Sheets(shtName).Range(RngArr(i)) = frBook.Sheets(shtName).Range(RngArr(i))
    Next
    If NeedUnprotect Then ProtectWorkSheet(toBook.Sheets(shtName)) = True
End Sub

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

Sub ModifyColumns(Optional NumberOfCols As Long = 1)
    'This is a hack to help people add/remove column for a new village
    ' First - unprotect the sheet
    ProtectWorkSheet(Sheet4) = False
    
    Dim rngEnd As Range, rngStart As Range, i As Long
        
    ' First cell of the table on the left, address for this cell will remain unchange
    Set rngStart = Range("RNG_IIAST").Offset(0, 4)
    ' top cell of the last column
    
    Set rngEnd = Range("RNG_II2A")
    ' modified on 11 Sep 2015 - simplifying the work
    
    ' to change a name address, just do this
    ' ThisWorkbook.Names.Add "Nothing", "=Sheet2!$B$11"
    
    Application.StatusBar = MSG("MSG_CREATE_II2A")
 
    ' move this range back 1 column
    CreateOrReizeName "RNG_II2A", NumberOfCols
    ' Revise position of range dta_bsc_vil
    CreateOrReizeName "dta_bsc_vil", NumberOfCols, False
    
    ' Now Resize, display columns.
    ResizeAndHide  ' resize stuff
    ' Create fomular and format header
    CreateFomular
    FormatHeaderCell
    
ExitCode:
    'Clean up
    Set rngStart = Nothing
    Set rngEnd = Nothing
    
    ProtectWorkSheet(Sheet4) = True
End Sub

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

Private Sub ResizeAndHide()
    Dim rng As Range, FmtRange As Range
    Set rng = Range("dta_bsc_vil")
    ' unlock first
    rng.Locked = False
    ' Now resize
    Set rng = rng.Offset(-1).Resize(rng.Rows.Count + 2)
    ' get the first range for later copying
    Set FmtRange = rng.Resize(, 1)
        
    FmtRange.Copy
    rng.PasteSpecial xlPasteFormats
    rng.PasteSpecial xlPasteValidation
    Application.CutCopyMode = False
    
    With rng
        With .Interior
            .ColorIndex = 2
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
        End With
        With .Borders(xlInsideVertical)
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .Weight = xlThin
        End With
        With .Borders(xlEdgeRight)
            .LineStyle = xlDouble
            .ColorIndex = xlAutomatic
            .Weight = xlThick
        End With
    End With
 
    ' resize column width
    rng.ColumnWidth = FmtRange.ColumnWidth
    ' resize last column
    Set FmtRange = FmtRange.Offset(, rng.Columns.Count)
    With FmtRange
        .ClearFormats
        .ClearContents
        .ColumnWidth = 2.25
    End With
    ' Now all other colum will be hidden
    Set FmtRange = FmtRange.Offset(0, 1).Resize(, 5)
    With FmtRange
        .EntireColumn.Hidden = True
        .ClearFormats
        .ClearContents
    End With
    ' Now change the print area...for this sheet
    Set rng = Sheets("II.2.A").Range(Sheets("II.2.A").PageSetup.PrintArea)
    Set rng = rng.Resize(, rng.Columns.Count + Range("dta_bsc_vil").Columns.Count)
    ' now set up printarea again
    Sheets("II.2.A").PageSetup.PrintArea = rng.Address
    
    Set FmtRange = Nothing
    Set rng = Nothing
End Sub

Private Sub CreateFomular()
    'This will help reformatting newly created table
    ' Begining column shall be total - 1 (coz RNG_IIA always stays at the end
    ' Range("RNG_IIAST").Offset(1, 4) offset 4 will always be the column for total
    ' Range("RNG_II2A_CELL_LAST").Offset(-1) will alway be the last cell at total column
    Dim rngStart As Range, rngEnd As Range, rngLastCell As Range
    Dim rngTotal As Range
    Dim MyCell As Range, i As Long
    
    ' reassign current worksheet
    Dim CurrentWorksheet As Worksheet
    Set CurrentWorksheet = ThisWorkbook.Sheets("II.2.A")
    
    With CurrentWorksheet
        Set rngStart = .Range("RNG_IIAST").Offset(0, 4)
        Set rngEnd = .Range("RNG_II2A")
        ' already inserted columns... so this failed
        Set rngLastCell = .Range("dta_bsc_vil").Cells(.Range("dta_bsc_vil").Rows.Count, 1).Offset(0, -1)
        Set rngTotal = .Range(rngStart.Offset(1), rngLastCell)
        
        ' Now that create total fomular
        ' + Create total column
        rngTotal.Formula = "=SUM(INDIRECT(""RC[1]" & ":RC[" & rngEnd.Column - rngStart.Column & "]"",FALSE))"
        
        ' + Create header link to data
        Set MyCell = Range("tblVillageStart")
        i = 0
        While Len(Trim(MyCell)) > 0
            i = i + 1
            .Range("RNG_IIAST").Offset(, 4 + i).Formula = "=INDIRECT(""Data!" & MyCell.Address & """)"
            Set MyCell = MyCell.Offset(1)
        Wend
        
        ' we have to unlock all data cells in this tables
        .Range("dta_bsc_vil").Locked = False
    End With
    Set rngStart = Nothing
    Set rngEnd = Nothing
    Set rngLastCell = Nothing
    Set rngTotal = Nothing
    Set MyCell = Nothing
    Set CurrentWorksheet = Nothing
End Sub

Private Sub FormatHeaderCell()
    ' Just for formatting the header
    Dim rngStart As Range, rngEnd As Range, MyCell As Range
    Set rngStart = Range("RNG_IIAST").Offset(0, 4)
    Set rngEnd = Range("RNG_II2A")
    Set MyCell = Sheet4.Range(rngStart.Offset(0, 1).Address & ":" & rngEnd.Address)
    
    ' Now format the header
    With MyCell
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 90
    End With
    With MyCell.Font
        .Name = "Times New Roman"
        .FontStyle = "Bold"
        .Size = 10
    End With
    With MyCell.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    Set rngStart = Nothing
    Set rngEnd = Nothing
    Set MyCell = Nothing
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

Sub ListName()
    Dim sh As Worksheet, wrk As Workbook, theName As Name
    Set wrk = ThisWorkbook
    For Each theName In wrk.Names
        If theName.RefersToLocal Like "*II.2*" Then
            Debug.Print theName.Name
            'dta_bsc_vil for II.2.A - then refer back to II.2
            'TBLMAJORINDS II.2.B Key indicators
            'II2BFIRST and II2BLAST for II.2.B

        End If
    Next
    Set wrk = Nothing
End Sub

Sub GetDistintiveList(TableName As String, KeyColumn As Long, ColumData As Long, Optional UseListOnly As Boolean = True)
    ' Check whether a temporary variable is valid
    On Error GoTo ErrHandler
    If IsCollection(CachedListDistinct) And OldTableName = TableName Then GoTo SetFuncValue
    OldTableName = TableName
    ' Now built the list, the sortable has been done before so we don't care
    Dim theRange As Range, theCell As Range, StrCount As Long, SpStr As String
    
    Dim ColDistinctive As New Collection
    Set theRange = ThisWorkbook.Names(TableName).RefersToRange
    SpStr = "[||]"
    
    Dim txtDistinct() As String, txtListing() As String, ColCount As Long
    Dim i As Long, MaxPos As Long, MaxStr As String, xPos As Long
    Dim FoundMatch As Boolean
    
    With theRange
        ColCount = .Columns.Count
        ReDim txtDistinct(ColCount - 1)
        ReDim txtListing(ColCount - 1)
        ReDim ColListing(xPos)
        
        'We keep each stuff in one collection item and the very first shall alway be the type, frequency
        ' next hack - convert range to array for quicker access
        Set theCell = .Cells(1, 1)
        While theCell <> ""
            ' move through all column
            If ", " & SpStr & theCell.Offset(0, KeyColumn - 1) & SpStr <> txtListing(KeyColumn - 1) Then
                If i > 0 Then
                    ' if already in process, so flush current array to variable and startnew array
                    xPos = xPos + 1
                    ReDim Preserve ColListing(xPos)
                    For i = LBound(txtListing) To UBound(txtListing)
                        ColListing(xPos).Add Mid(Replace(txtListing(i), SpStr, ""), 3)
                    Next
                    ReDim txtListing(ColCount - 1)
                End If
            End If
            For i = 1 To ColCount
                If InStr(txtListing(i - 1), SpStr & theCell.Offset(0, i - 1) & SpStr) = 0 Then
                    ' only add the new thing
                    txtListing(i - 1) = txtListing(i - 1) & ", " & SpStr & theCell.Offset(0, i - 1) & SpStr
                End If
                If InStr(txtDistinct(i - 1), SpStr & theCell.Offset(0, i - 1) & SpStr) = 0 Then
                    txtDistinct(i - 1) = txtDistinct(i - 1) & ", " & SpStr & theCell.Offset(0, i - 1) & SpStr
                    If i = KeyColumn Then
                        StrCount = 1
                        If MaxPos < StrCount Then
                            MaxPos = StrCount
                            MaxStr = theCell.Offset(0, KeyColumn - 1)
                        End If
                    End If
                Else
                    ' Find Max freq for key column
                    If i = KeyColumn Then
                        StrCount = StrCount + 1
                        If MaxPos < StrCount Then
                            MaxPos = StrCount
                            MaxStr = theCell.Offset(0, KeyColumn - 1)
                        End If
                    End If
                End If
            Next
            ' Okie - get along all columns already, now we need to see whether the next would be different
            Set theCell = theCell.Offset(1)
        Wend
        ' Add the last stuff
        xPos = xPos + 1
        ReDim Preserve ColListing(xPos)
        
        ' Now pass the array to the collection and cached them
        For i = LBound(txtDistinct) To UBound(txtDistinct)
            ColDistinctive.Add Mid(Replace(txtDistinct(i), SpStr, ""), 3)
            ColListing(xPos).Add Mid(Replace(txtListing(i), SpStr, ""), 3)
        Next
        Set CachedListDistinct = ColDistinctive
    End With
    ' Now we have to find the most appeared object and put it on top
    For i = 1 To UBound(ColListing)
        If ColListing(i).Item(KeyColumn) = MaxStr Then
            ' set the first item with this one
            Set ColListing(0) = ColListing(i)
            FoundMatch = True
        End If
        If FoundMatch And i < UBound(ColListing) Then Set ColListing(i) = ColListing(i + 1)
    Next
    ' resize the array
    ReDim Preserve ColListing(i - 2)
SetFuncValue:
    ' set the overall variable for later accessing
    proc_UseListOnly = UseListOnly
    proc_ColumData = ColumData
    NonCalculatedField = True
ErrHandler:
    If Err.Number <> 0 Then Debug.Print Err.description & "CurrentPointer=[" & CurrentPointer & "]"
End Sub

Sub TestAccessFormD()
    Set CachedListDistinct = Nothing
    ReDim ColListing(0)
    'SortTable ThisWorkbook, "II.5.D", "tblUnicode_4", "C6", "A6"
    Call GetDistintiveList("tblUnicode_3", 3, 3)
    For CurrentPointer = 0 To UBound(ColListing)
        Debug.Print ContextData()
    Next
End Sub

Function GetOption(TxtIn As String) As ObjectEquation()
    ' This will read the parametter and convert into an array for later processing
    Dim MyObj() As ObjectEquation, i As Long, ArrItem As Variant
    Dim myArr As Variant
    myArr = Split(TxtIn, "/")
    ReDim MyObj(UBound(myArr))
    
    For i = LBound(myArr) To UBound(myArr)
        ArrItem = Split(myArr(i), "=")
        With MyObj(i)
            .VariableName = ArrItem(0)
            ' call a sub to evaluate first
            Application.Evaluate (ArrItem(1))
            ' now get the data back to avoi 255 characters problem in Excel
            If NonCalculatedField Then
                .VariableFomular = ContextData()
            Else
                .VariableFomular = Application.Evaluate(ArrItem(1))
            End If
            NonCalculatedField = False
        End With
    Next
    GetOption = MyObj
End Function

Private Property Get ContextData() As String
    If proc_UseListOnly Then
        ContextData = ColListing(CurrentPointer).Item(proc_ColumData)
    Else
        ContextData = CachedListDistinct(proc_ColumData)
    End If
End Property

Function CountMaxRepetition(RangeName As String, CountColumn As Long, _
    Optional ReferColumn As Long = 0, Optional CountOnly As Long = 1, Optional InsertLineBreak As Boolean = False) As Variant
    'This function will count and get the maximum number of object repetition
    Dim theRange As Range, theCell As Range, RetObj As TextObject
    Dim StrTxt As String, StrCount As Long, MaxPos As Long, MaxStr As String, MaxRefText As String
    Dim RefStrTxt As String
 
    Set theRange = ThisWorkbook.Names(RangeName).RefersToRange
    ' Turn the range to an array for quick access
    Set theCell = theRange.Cells(1, CountColumn)
    While theCell <> ""
        If StrTxt <> theCell Then
            StrTxt = theCell
            If ReferColumn <> 0 Then RefStrTxt = theCell.Offset(, ReferColumn - CountColumn)
            StrCount = 1
        Else
            StrCount = StrCount + 1
            If ReferColumn <> 0 Then RefStrTxt = RefStrTxt & "[SEP]" & theCell.Offset(, ReferColumn - CountColumn)
            If MaxPos < StrCount Then
                MaxPos = StrCount
                MaxStr = StrTxt
                MaxRefText = RefStrTxt
            End If
        End If
        Set theCell = theCell.Offset(1)
    Wend
    
    'On Error Resume Next
    Select Case CountOnly
    Case 1:
        CountMaxRepetition = MaxStr
    Case 2:
        CountMaxRepetition = IIf(InsertLineBreak, Replace(MaxRefText, "[SEP]", vbCrLf), Replace(MaxRefText, "[SEP]", ", "))
    Case 3:
        CountMaxRepetition = Replace(MaxRefText, "[SEP]", ",")
    End Select
End Function

Private Function GetAverage(inputText As String) As String
    On Error GoTo ErrHandler
    Dim i As Long, theText As String, myArr As Variant, theTotal As Double
    theText = Replace(Replace(Replace(inputText, "(", ""), ")", ""), " ", "")
    myArr = Split(theText, ",")
    For i = LBound(myArr) To UBound(myArr)
        theTotal = theTotal + CDbl(myArr(i))
    Next
    theTotal = theTotal / i
    GetAverage = theTotal
ErrHandler:
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

Property Let ProtectWorkSheet(s As Worksheet, NewValue As Boolean)
    If NewValue Then
        If s.Name = "II.2.B" Then
            s.Protect AppPassWord, Contents:=True, AllowFormattingCells:=True, AllowFormattingColumns:=True, DrawingObjects:=True, Scenarios:=True, _
            AllowFormattingRows:=True, AllowSorting:=True, AllowFiltering:=True, UserInterfaceOnly:=True, AllowInsertingRows:=True, AllowDeletingRows:=True
        Else
            s.Protect AppPassWord, Contents:=True, AllowFormattingCells:=True, AllowFormattingColumns:=True, DrawingObjects:=True, Scenarios:=True, _
            AllowFormattingRows:=True, AllowSorting:=True, AllowFiltering:=True, UserInterfaceOnly:=True
        End If
    Else
        s.Unprotect AppPassWord
    End If
End Property

Sub RemoveProtection()
    Dim sh As Worksheet
    For Each sh In ThisWorkbook.Sheets
        ProtectWorkSheet(sh) = False
    Next
    ProtectWorkbook = False
End Sub

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

Sub SetPrintOption()
    ' Set print option for all
    If Not ActiveSheet.Name Like "II.*" Then Exit Sub
    With ActiveSheet.PageSetup
        .LeftMargin = Application.InchesToPoints(0.25)
        .RightMargin = Application.InchesToPoints(0.25)
        .TopMargin = Application.InchesToPoints(0.75)
        .BottomMargin = Application.InchesToPoints(0.75)
        .HeaderMargin = Application.InchesToPoints(0.3)
        .FooterMargin = Application.InchesToPoints(0.3)
        .PrintQuality = 600
        .CenterHorizontally = True
        .CenterVertically = False
        .Draft = False
        .PaperSize = xlPaperA4
        .Order = xlDownThenOver
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .PrintErrors = xlPrintErrorsDisplayed
        .OddAndEvenPagesHeaderFooter = False
        .DifferentFirstPageHeaderFooter = False
        .ScaleWithDocHeaderFooter = True
        .AlignMarginsHeaderFooter = False
    End With
End Sub

