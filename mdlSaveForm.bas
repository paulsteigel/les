Option Explicit

Sub GenerateFields()
    'This will generate field code for all
    Dim theName As Name, tCell As Range
    Set tCell = Range("tblFormInfor")
    For Each theName In ThisWorkbook.Names
        If theName.Name Like "txt_*" Then
            tCell = theName.Name
            Set tCell = tCell.Offset(0, 1)
        End If
    Next
End Sub

Sub ListLinks()
    'Updateby20140529
    Dim xIndex As Long, link As Object
    Dim wb As Workbook
    Set wb = Application.ActiveWorkbook
    If Not IsEmpty(wb.LinkSources(xlExcelLinks)) Then
        wb.Sheets.Add
        xIndex = 1
        For Each link In wb.LinkSources(xlExcelLinks)
            Application.ActiveSheet.Cells(xIndex, 1).Value = link
            xIndex = xIndex + 1
        Next link
    End If
End Sub

Sub TextName()
    Dim rn As Range
    Set rn = Range("tblFormInfor").Offset(0, 1)
    While rn <> ""
        If Not IsRangeValid(rn.Value) Then
            rn.Offset(1) = "False"
        End If
        Set rn = rn.Offset(0, 1)
    Wend
End Sub

Sub SearchName()
    Dim tObj As Range, tSheet As Worksheet
    Dim tSrc As Range
    Set tSrc = Range("tmpPtr").Offset(1)
    For Each tSheet In ThisWorkbook.Sheets
        For Each tObj In tSheet.UsedRange
            If Not tObj.Locked And tObj <> "" Then
                tSrc = tSheet.Name
                tSrc.Offset(0, 1) = tObj.Address
                tSrc.Offset(0, 2) = tObj.Value
                Set tSrc = tSrc.Offset(1)
            End If
        Next
    Next
End Sub

Sub text_down()
    Dim tName As Name
    Dim tSrc As Range
    Set tSrc = Range("tmpPtr").Offset(1)
    For Each tName In ThisWorkbook.Names
        If tName.Name Like "txt_*" Then
            tSrc = tName.Name
            tSrc.Offset(0, 1) = tName.RefersToRange.Address
            Set tSrc = tSrc.Offset(1)
        End If
    Next
End Sub

Sub SaveFormData()
    ' This sub will do the saving of form data to access table..
    ' What to check:
    ' HHLD code, Week Num, Visit num
    Dim oFile As Variant, oPath As String
    oPath = GetBrowseObject(, True, "xls")
    If oPath = "" Then GoTo CleanUp
    oFile = Split(oPath, "|")
    
    Dim i As Long, FileCount As Long
    Dim db As New clsDbConnection
    
    Dim StatusMsg As String, oldStatusBar As String
    StatusMsg = MSG("MSG_PROCESS_FILE")
    
    FileCount = UBound(oFile)
    db.ConnectDatabase AppDatabase
    
    oldStatusBar = Application.StatusBar
    Application.DisplayStatusBar = True
    ShowOff
    On Error GoTo NextStep
    DoEvents
    For i = 0 To FileCount
        If Not FileOrDirExists(CStr(oFile(i)), True) Then GoTo NextStep
        Application.StatusBar = Replace(StatusMsg, "%%", oFile(i)) & " " & Format(100 * (i + 1) / (FileCount + 1), "##") & "%..."
        ' Just copy all reange from this worbook to our current one
        CopyRangeValue CStr(oFile(i))
        
        SaveThisForm CStr(oFile(i)), db
NextStep:
        If Err.Number <> 0 Then
            Err.Clear
            WriteLog "Failed importing file: [" & oFile(i) & "]", "ImportError.txt"
        End If
    Next
CleanUp:
    
    Application.StatusBar = oldStatusBar
    Set db = Nothing
    ShowOff True
End Sub

Sub SaveCurrentForm()
    Dim dbs As New clsDbConnection
    dbs.ConnectDatabase AppDatabase
    SaveThisForm , dbs
    Set dbs = Nothing
    ThisWorkbook.Save
End Sub

Private Sub SaveThisForm(Optional FileObject As String, Optional db As clsDbConnection)
    Dim Sql As String, HhldCode As Long
    ' Now we have to verify whether this record is existing and delete them first
    Sql = "Delete * from tblFormInfor WHERE txt_IMS_ID ='" & Range("txt_IMS_ID") & "' AND txt_visit_date =#" & CDate(Range("txt_visit_date")) & "#;"
    db.ExecuteSQL Sql
    
    ' Now just insert into database
    Sql = GetSqlString()
    If Sql = "" Then
        If FileObject = "" Then FileObject = ThisWorkbook.Name
        WriteLog "Failed importing file: [" & FileObject & "]", "ImportError.txt"
        GoTo ErrHandler
    End If
    
    Sql = "INSERT INTO tblFormInfor " & Sql & ";"
    db.ExecuteSQL Sql
    
    ' For sub-table
    HhldCode = db.DMax("Form_ID", "tblFormInfor")
    ' Now select them to push into varying tables
    
    'Key tblMembersInfor information
    LoadTableToDb db, "tblMembersInfor", "tbl_hhld_members", HhldCode
ErrHandler:
End Sub

Private Sub CopyRangeValue(DstWb As String)
    Dim tName As Name, oBook As Workbook
    
    Set oBook = Workbooks.Open(DstWb, , True)
    
    For Each tName In oBook.Names
        If Not IsError(tName.RefersToRange) Then
            If tName.Name Like "sub_tbl*" Then
                Debug.Print tName.Name
            End If
            If IsRangeValid(tName.Name) Then
                If Not ThisWorkbook.Names(tName.Name).RefersToRange.Locked Then
                    ThisWorkbook.Names(tName.Name).RefersToRange.Value = tName.RefersToRange.Value
                End If
            End If
        End If
    Next
    oBook.Close False
    Set oBook = Nothing
    
    Application.CalculateFull
End Sub

Private Sub LoadTableToDb(dbs As clsDbConnection, tblName As String, tblRange As String, FormID As Long)
    Dim FldHdr As String, FldValue As String
    Dim HdrCell As Range, CrCell As Range
    Dim ColCount As Long, RowCount As Long
    Dim SqlTxt As String
    Dim idv_id As Long ' id of individual
        
    Dim sub_tableName As String
    ' Take the first cell
    Set HdrCell = Range(tblRange).Offset(0, 1)
    ' set running cell to the first record
    Set CrCell = HdrCell.Offset(1)
    While CrCell.Offset(RowCount) <> ""
        FldHdr = ""
        FldValue = ""
        ColCount = 0
        While HdrCell.Offset(0, ColCount) <> "" And HdrCell.Offset(-2, ColCount) <> "link"
            If IsError(CrCell.Offset(RowCount, ColCount)) Then GoTo NextLoop
            If CrCell.Offset(RowCount, ColCount) <> "" Then
                ' For header row
                FldHdr = FldHdr & ", " & HdrCell.Offset(0, ColCount)
                ' for value
                Select Case HdrCell.Offset(-1, ColCount)
                Case "DATETIME":
                    FldValue = FldValue & ", #" & CDate(CrCell.Offset(RowCount, ColCount)) & "#"
                Case "TEXT", "MEMO":
                    FldValue = FldValue & ", '" & StrQuoteReplace(CrCell.Offset(RowCount, ColCount)) & "'"
                Case Else:
                    FldValue = FldValue & ", " & CrCell.Offset(RowCount, ColCount)
                End Select
            End If
NextLoop:
            ColCount = ColCount + 1
        Wend
        ' Now inject into database
        SqlTxt = "INSERT INTO " & tblName & "(form_id" & FldHdr & ") VALUES(" & FormID & FldValue & ");"
        dbs.ExecuteSQL SqlTxt
        
        ' now get individual_id just inserted
        idv_id = dbs.DMax("Id", tblName)
        
        ' For sub-tables
        ' reset header and values
        FldHdr = ""
        FldValue = ""
        sub_tableName = ""
XXX:
        While HdrCell.Offset(-2, ColCount) = "link"
            If sub_tableName <> HdrCell.Offset(-1, ColCount) Then
                If sub_tableName <> "" Then
                    'Commit SQL now with processing of row level using separator
                    FormatAndInject dbs, sub_tableName, idv_id, FldHdr, FldValue
                    
                    ' reset again...
                    FldHdr = ""
                    FldValue = ""
                End If
                sub_tableName = HdrCell.Offset(-1, ColCount)
            End If
            ' since there maybe multiple value, we will have to parse them row by row
            FldHdr = FldHdr & ", " & HdrCell.Offset(0, ColCount)
            FldValue = FldValue & "[|]" & StrQuoteReplace(CrCell.Offset(RowCount, ColCount))
            
            ColCount = ColCount + 1
        Wend
        'GoTo XXX
        ' now deal the last time for previous table
        FormatAndInject dbs, sub_tableName, idv_id, FldHdr, FldValue
        
        RowCount = RowCount + 1
    Wend
    
End Sub

Private Sub FormatAndInject(dbs As clsDbConnection, TableName As String, ForeignKey As Long, HeaderRow As String, ValueRows As String)
    ' Break row value if needed
    Dim SqlTxt As String, i As Long
    Dim FldArr As Variant, VleArr1 As Variant, VleArr2 As Variant
    HeaderRow = Mid(HeaderRow, 3)
    ValueRows = Mid(ValueRows, 4)
    
    FldArr = Split(HeaderRow, ", ")
    
    If UBound(FldArr) = 0 Then
        ' only one field, separator is ";" - set as priority
        If InStr(ValueRows, ";") > 0 Then
            VleArr1 = Split(ValueRows, ";")
        Else
            VleArr1 = Split(ValueRows, ",")
        End If
        For i = 0 To UBound(VleArr1)
            If VleArr1(i) <> "" Then
                SqlTxt = "INSERT INTO " & TableName & "(individual_id, " & HeaderRow & ") VALUES(" & ForeignKey & ", '" & VleArr1(i) & "');"
                dbs.ExecuteSQL SqlTxt
            End If
        Next
    Else
        FldArr = Split(ValueRows, "[|]")
        VleArr1 = Split(IIf(Left(FldArr(0), 1) = vbLf, Mid(FldArr(0), 2), FldArr(0)), "," & vbLf)
        VleArr2 = Split(IIf(Left(FldArr(1), 1) = vbLf, Mid(FldArr(1), 2), FldArr(1)), vbLf)
        For i = 0 To UBound(VleArr1)
            If VleArr1(i) <> "" Then
                If i > UBound(VleArr2) Then
                    SqlTxt = "INSERT INTO " & TableName & "(individual_id, " & HeaderRow & ") VALUES(" & ForeignKey & ", '" & VleArr1(i) & "', '');"
                Else
                    SqlTxt = "INSERT INTO " & TableName & "(individual_id, " & HeaderRow & ") VALUES(" & ForeignKey & ", '" & VleArr1(i) & "', '" & VleArr2(i) & "');"
                End If
                dbs.ExecuteSQL SqlTxt
            End If
        Next
    End If
End Sub

Private Function GetSqlString() As String
    Dim SqlTxt As String
    Dim fldName As String, FldValue As String, HdrPtr As Range
    
    Set HdrPtr = Range("tblFormInfor").Offset(0, 1)
    
    While HdrPtr <> ""
        If Not IsRangeValid(HdrPtr.Value) Then GoTo NextStep
        ' Just whether value is blank or not
        If IsError(Range(HdrPtr)) Then GoTo NextStep
        If Range(HdrPtr).Value <> "" Then
            fldName = fldName & ", " & HdrPtr.Value
        Else
            If HdrPtr.Offset(-3) = 1 Then
                MsgBox MSG("MSG_NO_BLANK"), vbInformation
                GoTo CleanUp
            End If
            GoTo NextStep
        End If
        
        Select Case HdrPtr.Offset(-1)
        Case "DATETIME":
            FldValue = FldValue & ", #" & CDate(Range(HdrPtr).Value) & "#"
        Case "TEXT", "MEMO":
            FldValue = FldValue & ", '" & StrQuoteReplace(Range(HdrPtr)) & "'"
        Case Else
            FldValue = FldValue & ", " & Range(HdrPtr).Value
        End Select
        
NextStep:
        Set HdrPtr = HdrPtr.Offset(0, 1)
    Wend
        SqlTxt = "(" & Mid(fldName, 2) & ") VALUES(" & Mid(FldValue, 2) & ")"
        
CleanUp:
    GetSqlString = SqlTxt
End Function

Sub Export2Excel(FilterStr As String)
    'GetFiles SaveFormData
    ' this will open all for exporting
    Dim SqlTxt As String, rs As ADODB.Recordset, dbs As New clsDbConnection
    Dim wb As Workbook, wsh As Worksheet, ptrCell As Range
    Dim StatusTxt As String
    ShowOff
    DoEvents
    dbs.ConnectDatabase AppDatabase
    SqlTxt = "Select FieldName,FieldCaption from tblFieldMap Where UseInExport=true AND TableName='tblFormInfor' ORDER BY ExcelFieldOrder ASC;"
    Set rs = dbs.GetRecordSet(SqlTxt, True)
    SqlTxt = ""
    StatusTxt = MSG("MSG_SEND_DATA_TO_SHEET")
    
    Set wb = Workbooks.Add
    'copy me to the new workbook
    ThisWorkbook.Sheets("household").Copy Before:=wb.Sheets(1)
    
    Set wsh = wb.Sheets("household")
    
    Set ptrCell = wsh.Cells(1)
    While Not rs.EOF
        ptrCell.Value = rs.Fields("FieldCaption")
        SqlTxt = SqlTxt & "," & rs.Fields("FieldName")
        Set ptrCell = ptrCell.Offset(0, 1)
        rs.MoveNext
    Wend
    rs.Close
    SqlTxt = "SELECT " & Mid(SqlTxt, 2) & " FROM tblFormInfor WHERE " & FilterStr & " AND txt_project <> '' AND txt_visit_date <> null;"
    
    Application.StatusBar = Replace(StatusTxt, "%%", "[" & wsh.Name & "]")
        
    Set rs = dbs.GetRecordSet(SqlTxt)
    With wsh
        .Cells(2, 1).CopyFromRecordset rs
        .Range("W:W").NumberFormat = "General"
        .Range("EI:EJ").NumberFormat = "General"
    End With
    ' Show the sheeet now
    wb.Names("rngFilter_hhld").RefersToRange.AutoFilter
    wsh.UsedRange.WrapText = False
    wsh.Visible = xlSheetVisible
    
    ' Now we have to load all individual data for these people
    ' It's a bit hard then?? quite some crosstab..
    ThisWorkbook.Sheets("individual").Copy Before:=wb.Sheets(1)
    
    Set wsh = wb.Sheets("individual")
    Set ptrCell = wsh.Cells(2, 2)
    rs.MoveFirst
    While Not rs.EOF
        Application.StatusBar = Replace(StatusTxt, "%%", "[" & wsh.Name & "]" & " Household IMS Code [" & rs.Fields("Form_ID") & "]!")
        GetIndividualList rs.Fields("Form_ID"), dbs, ptrCell
        rs.MoveNext
    Wend
    rs.Close
    '===========
    wb.Names("rngFilter_indv").RefersToRange.AutoFilter
    wsh.UsedRange.WrapText = False
    wsh.Visible = xlSheetVisible
    wsh.Activate
    
    ' Delete other sheets
    Application.DisplayAlerts = False
    For Each wsh In wb.Sheets
        If wsh.Name <> "household" And wsh.Name <> "individual" Then
            Debug.Print wsh.Name
            wsh.Delete
        End If
    Next
    Application.DisplayAlerts = True
    
    Set wb = Nothing
    Set dbs = Nothing
    Application.StatusBar = "Finished exporting..."
    ShowOff True
End Sub

Private Sub GetIndividualList(FormID As Long, db As clsDbConnection, RowPtr As Range)
    Dim Sql As String, RowCount As Long
    Dim rs As New ADODB.Recordset
    ' First create this query....
    Sql = "SELECT b.id, a.Form_ID AS Les_id, a.txt_project, a.txt_month_visit, a.txt_week_visit, " & _
    "a.txt_visit_num_les_id, a.txt_IMS_ID, a.txt_IMS_ID_2, a.txt_house_owner, a.txt_village, " & _
    "a.txt_commune, a.txt_staff_name, a.txt_visit_num, b.Member_Name, b.Mem_IMS, b.Mem_id, " & _
    "b.Mem_gender, b.Mem_DOB, Month([Mem_DOB]) AS Mem_DOB_month, Year([Mem_DOB]) AS Mem_DOB_year, " & _
    "DateDiff('yyyy',[Mem_DOB],Now()) AS Mem_age, b.Mem_DOB AS Mem_age_class, b.Mem_tel, " & _
    "b.Mem_rel_hhld, b.Mem_rel_hhld_other, b.Edu, b.Edu_eval, b.Key_job, b.Key_job_other, " & _
    "b.Min_job, b.Min_job_other, b.Job_status, b.Income_avrg, b.Insurance_support, " & _
    "b.is_hhld_member, b.is_reallocate, b.Move_to, b.Move_reason, b.Move_reason_details, " & _
    "b.skill_eval, b.link_type, b.no_link_reason, b.link_demand, b.link_dificulty " & _
    "FROM tblFormInfor AS a INNER JOIN tblMembersInfor AS b ON a.Form_ID = b.form_id " & _
    "WHERE a.form_id = " & FormID & ";"
    ' Load to Excel
    RowCount = db.GetRecordSet("Select Count(*) " & _
    "FROM tblFormInfor AS a INNER JOIN tblMembersInfor AS b ON a.Form_ID = b.form_id " & _
    "WHERE a.form_id = " & FormID & ";", True).Fields(0)
    
    Set rs = db.GetRecordSet(Sql)
    RowPtr.CopyFromRecordset rs
    Set RowPtr = RowPtr.Offset(RowCount)
    rs.Close
End Sub
Sub RetrieveFields()
    ' This will do the cleaning of data before doing things, to avoid blank
    Dim txtValue As Range, retVal As String, DataType As String, i As Long
    Dim HdrCell As Range, rs As New ADODB.Recordset, db As New clsDbConnection
    db.ConnectDatabase AppDatabase
    Set rs = db.GetRecordSet("Select * from tblMembersInfor where form_id=0;", True)
    Set HdrCell = Range("tbl_hhld_member_details").Offset(0, 1)
    For i = 0 To rs.Fields.Count - 1
        HdrCell.Offset(0, i) = rs.Fields(i).Name
        HdrCell.Offset(-2, i) = ""
        Select Case rs.Fields(i).Type
        Case 202:
            HdrCell.Offset(-1, i) = "TEXT"
            HdrCell.Offset(-2, i) = rs.Fields(i).DefinedSize
        Case 7:
            HdrCell.Offset(-1, i) = "DATETIME"
        Case 3:
            HdrCell.Offset(-1, i) = "INTEGER"
        Case 203:
            HdrCell.Offset(-1, i) = "MEMO"
        Case 4:
            HdrCell.Offset(-1, i) = "SINGLE"
        End Select
    Next
    rs.Close
    Set db = Nothing
End Sub

Sub deleteChart()
    Dim theChart As Chart
    For Each theChart In ThisWorkbook.Charts
        theChart.Delete
    Next
End Sub

Sub TextMeNow()
    SetSheetSize "individual"
    SetSheetSize "household"
End Sub

Private Sub SetSheetSize(SheetName As String)
    Dim tCell As Range, oCell As Range
    Set tCell = ThisWorkbook.Sheets(SheetName).Cells(1)
    Set oCell = ActiveWorkbook.Sheets(SheetName).Cells(1)
    While oCell <> ""
        tCell.ColumnWidth = oCell.ColumnWidth
        Set tCell = tCell.Offset(0, 1)
        Set oCell = oCell.Offset(0, 1)
    Wend
End Sub
