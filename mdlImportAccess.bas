Option Explicit
Private DefaultTable As String

Function SheetExists(sh As String, wb As Workbook) As Boolean
    On Error GoTo ErrHandler
    Dim sht As Worksheet
    Set sht = wb.Sheets(sh)
    SheetExists = True
ErrHandler:
    Set sht = Nothing
End Function

Function IsRangeValid(RangeName As String) As Boolean
    ' For checking local range
    Dim rng As Range
    On Error GoTo ErrHandler
    Set rng = ThisWorkbook.Names(RangeName).RefersToRange
    IsRangeValid = True
ErrHandler:
End Function

Function GetDateFromString(Optional txtString = "") As Date
    Dim DayTxt As Long, MonthTxt As Long, YearTxt As Long
    If txtString = "" Then Exit Function
    Dim DtArr As Variant
    DtArr = Split(txtString, "/")
    DayTxt = DtArr(0)
    MonthTxt = DtArr(1)
    YearTxt = DtArr(2)
    GetDateFromString = DateSerial(YearTxt, MonthTxt, DayTxt)
End Function

Function DayDiffRange(InDate As Date) As Long
    Dim NoOfDay As Long
    NoOfDay = DateDiff("d", #6/1/2014#, InDate)
    DayDiffRange = NoOfDay
End Function

Function MonthDiffRange(InDate As Date) As Long
    Dim NoOfDay As Long
    NoOfDay = DateDiff("m", #5/30/2014#, InDate)
    MonthDiffRange = NoOfDay
End Function

Sub PushDataToTable()
    ' This will send data to a temporary table for quick things
    Dim db As New clsDbConnection, Sql As String, ptrCell As Range, HdrCell As Range
    Dim FldTxt As String, FldVal As String, i As Long, ptrRec As Long, OldFieldName As String
    Dim FldTxtArr() As String, FldValArr() As String
    
    db.CreateDb (AppDatabase)
    Set ptrCell = Range("tblFormInfor").Offset(1)
    While ptrCell <> ""
        Set HdrCell = Range("tblFormInfor")
        ' Now push all detail into these created table
        FldTxt = FldTxt & HdrCell & ","
        If HdrCell.Offset(-1) = "TEXT" Or HdrCell.Offset(-1) = "MEMO" Then
            FldVal = FldVal & "'" & StrQuoteReplace(ptrCell) & "',"
        Else
            FldVal = FldVal & IIf(ptrCell = "", 0, ptrCell) & ","
        End If
        Set ptrCell = ptrCell.Offset(0, 1)
        Set HdrCell = HdrCell.Offset(0, 1)
    
        Sql = "INSERT INTO tblFormInfor(" & Left(FldTxt, Len(FldTxt) - 1) & ") VALUES(" & Left(FldVal, Len(FldVal) - 1) & ");"
        
        db.ExecuteSQL Sql
        FldTxt = ""
        FldVal = ""
    Wend
    
    Set db = Nothing
End Sub

Private Sub ParseRange(theRange As String, dbObj As clsDbConnection, RefField As String)
    ' This is just a single structure table so we dont need anything too important
    
    Dim tblName As String, cellPtr As Range
    
    If dbObj.TableExist(theRange) Then dbObj.DropTable (theRange)
    dbObj.ExecuteSQL "Create table " & theRange & "(ID AUTOINCREMENT, " & Replace(RefField, ",", " LONG,") & " LONG, FldName Text(100));"
    Set cellPtr = Range(theRange).Cells(1)
    While cellPtr <> ""
        dbObj.ExecuteSQL "INSERT INTO " & theRange & "(" & RefField & ",FldName) VALUES(" & Replace(cellPtr.Offset(0, -1), ".", ",") & ",'" & cellPtr & "');"
        Set cellPtr = cellPtr.Offset(1)
    Wend
End Sub

Function QueryValue(ByVal SqlText As String, Optional FieldNumber As String = "0", _
    Optional RecordNumber As Long = 0, Optional LeadingText As Range = Nothing) As String
    ' This will calculate response base on query string
    Dim db As New clsDbConnection, Sql As String, dbRs As Object, tmpStr As String
    Dim FldArr As Variant, txtSpr As String, i As Long
    FldArr = Split(FieldNumber, ",")
    If UBound(FldArr) = 0 Then txtSpr = "" Else txtSpr = "||"
    db.ConnectDatabase AppDatabase
    Set dbRs = db.GetRecordSet(SqlText)
    While Not dbRs.EOF
        For i = 0 To UBound(FldArr)
            If i = 1 Then tmpStr = tmpStr & LeadingText & txtSpr
            tmpStr = tmpStr & dbRs.Fields(Val(FldArr(i)) - 1) & txtSpr
        Next
        tmpStr = tmpStr & vbLf
        dbRs.MoveNext
        If RecordNumber <> 0 Then GoTo ExitLoop
    Wend
ExitLoop:
    If tmpStr <> "" Then tmpStr = Left(tmpStr, Len(tmpStr) - Len(txtSpr) - 1)
    QueryValue = tmpStr
    dbRs.Close
    db.CloseConnection
End Function

