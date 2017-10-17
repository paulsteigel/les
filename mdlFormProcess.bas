Option Explicit
Global Const AppDatabase = "m_c_les_project.mdb"
' This is the control module for M&E sheet
Private Type DataPair
    DataHeader As String
    DataBit As String
End Type

Sub MassPrintOpt(Optional DirectCallMode As Boolean = True, Optional DataFilter As String = "")
    ' To select and print all...
    If Not DirectCallMode Then GoTo DirectPrint
    Select Case ActiveSheet.Name
    Case "PrintList":
        If ActiveCell.Column <> 2 And ActiveCell.Row <= 4 Then Exit Sub
        ' Now loop through the option to deal...
        Dim ptrCell As Range, SelRange As Range, tmpRng As Range
        Dim db As New clsDbConnection, rs As New ADODB.Recordset
        Dim PtrCount As Long, i As Long, tmpSheet As Worksheet, wb As Workbook
        Dim NewtmpRng As Range, j As Long, xW As Long

DirectPrint:
        db.ConnectDatabase (AppDatabase)

        'set wrap text
        Set wb = Workbooks.Add
        Set tmpSheet = wb.Sheets("Sheet1")
        Set tmpRng = tmpSheet.Cells(1)
        tmpRng.WrapText = True
        
        ThisWorkbook.Activate
        If DirectCallMode Then GoTo NextStep
        
        Set SelRange = Selection
        Set ptrCell = SelRange.Cells(1)
        For i = 1 To SelRange.Cells.Count
            If ptrCell = "" Then Exit For
NextStep:
            ' Now call the printer, this actually print all data for the selected household
            If Not DirectCallMode Then
                Set rs = db.GetRecordSet("Select Form_ID from tblFormInfor WHERE " & DataFilter & ";", True)
            Else
                Set rs = db.GetRecordSet("Select Form_ID from tblFormInfor WHERE txt_IMS_ID='" & ptrCell.Value & "';", True)
            End If
            DoEvents
            While Not rs.EOF
                If rs.Fields(0) > 0 Then
                    ShowFormAndPrint "Form_ID=" & rs.Fields(0), True
                    ShowOff True
                    ' Now call the printer
                    If PtrCount = 0 Then
                        Sheets(Array("Part_A.1", "Part_A.2", "Part_B.1", "Part_B.2&C.1", "Part_C.2", "Part_D", "Part_E&F", "General", "Ranking")).Select
                        ' Now format row
                        
                        Set NewtmpRng = Range("rate_hhld_summary").MergeArea
                        For j = 1 To NewtmpRng.Columns.Count
                            xW = xW + NewtmpRng.Cells(1, j).ColumnWidth
                        Next
                        ' now we have width
                        tmpRng.ColumnWidth = xW
                    End If
                    ' Needed unprotection to set rowheight
                    Sheet15.Unprotect
                    ' set width
                    SetWraptSize tmpRng, "rate_hhld_summary"
                    SetWraptSize tmpRng, "txt_Assessment_comments"
                    SetWraptSize tmpRng, "rate_hhld_summary_update"
                    SetWraptSize tmpRng, "rate_hhld_summary_update_content"
                    Sheet15.Unprotect
                    
                    ' now print
                    ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True, IgnorePrintAreas:=False
                    PtrCount = PtrCount + 1
                End If
                rs.MoveNext
            Wend
            ' get out of the loop if indirect call
            If Not DirectCallMode Then Exit For
            
            ptrCell.Offset(0, 4).Value = "x"
            Set ptrCell = ptrCell.Offset(1)
        Next
        rs.Close
        ' only close when this is an indirect call
        db.CloseConnection
        ThisWorkbook.Sheets("PrintList").Select
        wb.Close False
    Case Else
        MsgBox MSG("MSG_CHANGE_TO_PRINTSHEET_AND_SELECT_FIRST"), vbInformation
        If SheetValid("PrintList", ThisWorkbook) Then
            ThisWorkbook.Sheets("PrintList").Activate
        End If
    End Select
End Sub

Private Sub SetWraptSize(dstRange As Range, SrcRange As String)
    dstRange.Value = Range(SrcRange)
    dstRange.WrapText = True
    Range(SrcRange).RowHeight = dstRange.RowHeight
End Sub

Sub CreateNewForm()
    ' This will make a new form and will load whatever that has been save before for this village...
    ' It is better to help user select village, insert fixed data then
    Dim db As New clsDbConnection, dbPath As String
    db.ConnectDatabase (AppDatabase)
    
    Dim NewCellID As Long
    ' reset all
    NewCellID = db.GetRecordSet("Select Max(Form_ID) from tblFormInfor;").Fields(0)
    ClearForm NewCellID + 1
    Set db = Nothing
End Sub

Sub SelectForm()
    ' This will help people look up this from in database
    Dim isSelected As Boolean
    With frmObjectParameter
        .SpecialNote = MSG("MSG_BROWSE_FORM")
        .DataSetName = MSG("MSG_SELECT_FORM")
        .DataSource = "lst_hhld_source"
        .DontAssignActiveCell = True
        .ReadOnly = True
        .WrapOutput = True
    End With
    
    If Not isSelected Then frmSelect.Show vbModal
    If frmObjectParameter.SelectedItem <> "" Then LoadFormData Val(frmObjectParameter.SelectedItem)
    'reset form argument value
    Dim lRet As FormArgument
    frmObjectParameter = lRet
End Sub

Sub ShowFormAndPrint(FormID As String, Optional PrintNow As Boolean = False)
    ClearForm , PrintNow
    LoadFormData FormID
End Sub

Sub ClearForm(Optional NewAssignedID As Long = 0, Optional NoConfirm As Boolean = False)
    ' Clear current enter Form
    ' Makesure
    ' This will only clear the detail form
    If Not NoConfirm Then
        If MsgBox(MSG("MSG_CONFIRM_CLEAR_FORM"), vbQuestion + vbYesNo) = vbNo Then Exit Sub
    End If
    
    ShowOff
    Dim theName As Name
    For Each theName In ThisWorkbook.Names
        If IsRangeValid(theName.Name) Then
            If theName.Name Like "txt_*" And theName.RefersToRange.Locked = False Then
                Debug.Print theName.RefersToRange.Address
                If theName.RefersToRange.MergeCells Then
                    theName.RefersToRange.MergeArea.ClearContents
                Else
                    theName.RefersToRange.ClearContents
                End If
                'txt_visit_date_null
                'txt_hhld_cat_null
            End If
        End If
    Next
    ' now reset the two tables
    Range("sub_tbl_1_1").ClearContents
    Range("sub_tbl_1_2").ClearContents
    Range("sub_tbl_2").ClearContents
    ShowOff True
End Sub

Sub SelectFormFromData()
    frmSearch.Show vbModal
End Sub

Private Sub LoadFormData(Form_ID As String)
    Dim db As New clsDbConnection, i As Long, dbPath As String, HhldCode As Long, rsDb As ADODB.Recordset, Sql As String
    db.ConnectDatabase (AppDatabase)
    Dim rs As ADODB.Recordset
    Set rs = db.GetRecordSet("Select * from tblFormInfor WHERE " & Form_ID & ";", True)
    If rs.EOF Then GoTo ExitMe
    ShowOff
    For i = 1 To rs.Fields.Count - 1
        If Not IsRangeValid(rs.Fields(i).Name) Then GoTo NextStep
        If Not Range(rs.Fields(i).Name).Locked Then
            Range(rs.Fields(i).Name).Value = rs.Fields(i)
        Else
            If IsRangeValid(rs.Fields(i).Name & "_null") Then
                'load indirect range
                Debug.Print rs.Fields(i).Name
                If rs.Fields(i).Name Like "*date" Then
                    Range(rs.Fields(i).Name & "_null") = Format(rs.Fields(i), "dd/mm/yyyy")
                Else
                    If Not Range(rs.Fields(i).Name & "_null").Locked Then
                        If rs.Fields(i).Name & "_null" = "txt_commune_null" Then
                            'Look for named range for commune
                            Set rsDb = db.GetRecordSet("Select RangeName from tblCommune where CommuneName='" & rs.Fields(i) & "';")
                            If Not rsDb.EOF Then Range(rs.Fields(i).Name & "_null") = rsDb.Fields(0)
                        Else
                            Range(rs.Fields(i).Name & "_null") = rs.Fields(i)
                        End If
                    End If
                End If
            End If
        End If
NextStep:
    Next
    ' Now for null range
    HhldCode = rs.Fields("Form_ID")
    
    ' Now load other table...
    Sql = "SELECT Member_Name, Mem_IMS, Mem_id, Mem_gender, Mem_DOB, Mem_tel, Mem_rel_hhld, " & _
        "Mem_rel_hhld_other, Edu FROM tblMembersInfor " & _
        "WHERE form_id=" & HhldCode & ";"
    LoadToTable db, Range("sub_tbl_1_1"), Sql
    
    Sql = "SELECT Key_job, Key_job_other, Min_job, Min_job_other, Job_status, Income_avrg, " & _
        "Insurance_support, is_reallocate, Move_to, Move_reason, Move_reason_details, is_hhld_member " & _
        "FROM tblMembersInfor WHERE form_id=" & HhldCode & ";"
    LoadToTable db, Range("sub_tbl_1_2"), Sql
    
    Dim proc_range As Range, tCell As Range
    Set proc_range = Range("sub_tbl_2")
    
    Set rsDb = db.GetRecordSet("SELECT ID, Member_Name,skill_eval,link_type,link_demand,link_dificulty,no_link_reason " & _
        "FROM tblMembersInfor " & _
        "WHERE form_id = " & HhldCode & ";", True)
    i = 0
    While Not rsDb.EOF
        ' Load name
        proc_range.Cells(i + 1, 1).Value = rsDb.Fields("Member_Name")
        ' Load skill evaluation
        proc_range.Cells(i + 1, 7).Value = rsDb.Fields("skill_eval")
        proc_range.Cells(i + 1, 9).Value = rsDb.Fields("link_type")
        proc_range.Cells(i + 1, 12).Value = rsDb.Fields("no_link_reason")
        proc_range.Cells(i + 1, 13).Value = rsDb.Fields("link_demand")
        proc_range.Cells(i + 1, 14).Value = rsDb.Fields("link_dificulty")
        
        ' Now load skill
        Sql = "SELECT a.SkillName, a.SkillSource FROM tbl_skills AS a " & _
        "WHERE a.individual_id=" & rsDb.Fields("id") & ";"
        
        Set tCell = Range(proc_range.Cells(i + 1, 2), proc_range.Cells(i + 1, 3))
        LoadToTable db, tCell, Sql, True
        
        ' Load course joined
        Sql = "SELECT a.ProjectDetails, a.ProjectName FROM tbl_project_joined AS a " & _
        "WHERE a.individual_id=" & rsDb.Fields("id") & ";"
        
        Set tCell = Range(proc_range.Cells(i + 1, 4), proc_range.Cells(i + 1, 5))
        LoadToTable db, tCell, Sql, True
        
        ' Load expected job
        Sql = "SELECT a.expected_job FROM tbl_job_expect AS a " & _
        "WHERE a.individual_id=" & rsDb.Fields("id") & ";"
        
        Set tCell = Range(proc_range.Cells(i + 1, 6), proc_range.Cells(i + 1, 6))
        LoadToTable db, tCell, Sql, True
        
        ' load skill
        Sql = "SELECT a.expected_skill FROM tbl_skill_expect AS a " & _
        "WHERE a.individual_id=" & rsDb.Fields("id") & ";"
        
        Set tCell = Range(proc_range.Cells(i + 1, 8), proc_range.Cells(i + 1, 8))
        LoadToTable db, tCell, Sql, True
                
        ' Load link demand
        Sql = "SELECT a.linkdetails, a.linktype FROM tbl_job_links AS a " & _
        "WHERE a.individual_id=" & rsDb.Fields("id") & ";"
        
        Set tCell = Range(proc_range.Cells(i + 1, 10), proc_range.Cells(i + 1, 11))
        LoadToTable db, tCell, Sql, True
        
        rsDb.MoveNext
        i = i + 1
    Wend
    
ExitMe:
    rsDb.Close
    rs.Close
    Set db = Nothing
    ShowOff True
End Sub

Private Sub LoadToTable(dbs As Object, tblRange As Range, tbSql As String, Optional ConcatenateText As Boolean = False)
    Dim rs As ADODB.Recordset
    Dim rsDest As Range, j As Long, CrCell As Range, i As Long, CellText As String
    Set rsDest = tblRange
    Set rs = dbs.GetRecordSet(tbSql, True)
    While Not rs.EOF
        For i = 1 To rs.Fields.Count
            Set CrCell = rsDest.Cells(j + 1, i)
            If rs.Fields(i - 1).Name Like "*DOB*" Or rs.Fields(i - 1).Name Like "*date*" Then
                CrCell.Value = Format(rs.Fields(i - 1), "dd/mm/yyyy")
            Else
                If ConcatenateText Then
                    If rs.Fields(i - 1) <> "" Then CrCell.Value = CrCell.Value & "," & vbLf & rs.Fields(i - 1)
                Else
                    CrCell.Value = rs.Fields(i - 1)
                End If
            End If
        Next
        If Not ConcatenateText Then j = j + 1
        rs.MoveNext
    Wend
    If ConcatenateText Then
        j = 0
        For i = 1 To rs.Fields.Count
            Set CrCell = rsDest.Cells(j + 1, i)
            If CrCell.Value = "," & vbLf Then
                CrCell.Value = ""
            Else
                CrCell.Value = Mid(CrCell.Value, 3)
            End If
        Next
    End If
    rs.Close
End Sub

Private Sub DeleteRange(theRange As Range)
    On Error GoTo ErrHandler
    theRange.Delete
ErrHandler:
End Sub

Private Sub SenDatatoDb(dbObject As Object, tblName As String, Optional ParentTableId As Long = 0, Optional PullID As Long = 0)
    Dim Sql As String, ptrCell As Range, HdrCell As Range
    Dim FldTxt As String, FldVal As String, ptrRec As Long
    Dim tmpValue As DataPair
    ' First delete all..
AgainNow:
    ' No deletion is needed for this stage
    'If ParentTableId > 0 Then
    '    dbObject.ExecuteSQL "Delete * from " & tblname & " WHERE hhld_ims_code=" & ParentTableId & ";"
    'Else
    '    dbObject.ExecuteSQL ("Delete * from " & tblname & " WHERE Form_ID=" & PullID & ";")
    'End If
BackToWork:
    Set HdrCell = Range(tblName).Offset(0, 1)
    Set ptrCell = Range(tblName).Offset(0, 1)
    While HdrCell <> ""
        ' Now push all detail into these created table
        
        If ParentTableId > 0 Then
            tmpValue = CleanData(HdrCell, HdrCell.Offset(ptrRec + 1))
        Else
            tmpValue = CleanData(HdrCell, Range(HdrCell))
        End If
        If tmpValue.DataHeader <> "" Then
            FldTxt = FldTxt & tmpValue.DataHeader & ","
            FldVal = FldVal & tmpValue.DataBit & ","
        End If
        Set HdrCell = HdrCell.Offset(0, 1)
    Wend
    
    If FldVal = "" Then Exit Sub
    
    If ParentTableId > 0 Then
        Sql = "INSERT INTO " & tblName & "(hhld_ims_code, " & Left(FldTxt, Len(FldTxt) - 1) & ") " & "VALUES(" & ParentTableId & ", " & Left(FldVal, Len(FldVal) - 1) & ");"
        dbObject.ExecuteSQL Sql
        
        FldTxt = ""
        FldVal = ""
        ptrRec = ptrRec + 1
        If ptrCell.Offset(ptrRec + 1) <> 0 Then GoTo BackToWork
    Else
        Sql = "INSERT INTO " & tblName & "(" & Left(FldTxt, Len(FldTxt) - 1) & ") " & _
        "VALUES(" & Left(FldVal, Len(FldVal) - 1) & ");"
        dbObject.CreateQuery "dd", Sql
        'GoTo AgainNow
        dbObject.ExecuteSQL Sql
    End If
End Sub

Sub import2db()
    ' do the import to database
    Dim i As Long, HdrCell As Range, cellPtr As Range
    Set cellPtr = Sheets("household").Cells(3, 1)
    For i = 1 To 138
        If cellPtr <> "" Then
            Set HdrCell = GetCellValue(cellPtr)
            cellPtr.Offset(-1).Value = HdrCell.Offset(-1).Value
            cellPtr.Offset(-2).Value = HdrCell.Offset(-2).Value
        End If
        Set cellPtr = cellPtr.Offset(0, 1)
    Next
End Sub

Sub ImportSenDatatoDb()
    Dim Sql As String, ptrCell As Range, HdrCell As Range
    Dim FldTxt As String, FldVal As String, ptrRec As Long, i As Long
    Dim tmpValue As DataPair, tblName As String
    Dim dbObject As New clsDbConnection, parent_id As Long, dbPath As String, CurPtr As Range
    dbPath = AppDatabase
    
    If Not FileOrDirExists(dbPath, True) Then
        dbObject.CreateDb dbPath
    Else
        dbObject.ConnectDatabase (dbPath)
    End If
         
    tblName = "tblFormInfor"
    ' First delete all..
    dbObject.ExecuteSQL ("Delete * from " & tblName & ";")
    
    Set ptrCell = Sheets("household").Cells(5, 1)
    While ptrCell <> ""
        Set HdrCell = Sheets("household").Cells(3, 1)
        While HdrCell <> ""
            ' Now push all detail into these created table
            
            tmpValue = CleanData(HdrCell, ptrCell.Offset(0, i))
            If tmpValue.DataHeader <> "" Then
                FldTxt = FldTxt & tmpValue.DataHeader & ","
                FldVal = FldVal & tmpValue.DataBit & ","
            End If
            i = i + 1
            Set HdrCell = HdrCell.Offset(0, 1)
        Wend
        
        If FldVal <> "" Then
            Sql = "INSERT INTO " & tblName & "(" & Left(FldTxt, Len(FldTxt) - 1) & ") " & _
            "VALUES(" & Left(FldVal, Len(FldVal) - 1) & ");"
        '        dbObject.CreateQuery "dd", Sql
            i = 0
            FldTxt = ""
            FldVal = ""
            
            'dbObject.CreateQuery "dd", Sql
            dbObject.ExecuteSQL Sql
            
        End If
        Set ptrCell = ptrCell.Offset(1)
    Wend
End Sub

Private Function GetCellValue(CellData As Range) As Range
    Dim IsFound As Boolean, HdrCell As Range
    Set HdrCell = Range("tblFormInfor").Offset(0, 1)
    While Not IsFound
        If HdrCell <> "" Then
            If HdrCell = CellData Then IsFound = True
        Else
            IsFound = True
        End If
        Set HdrCell = HdrCell.Offset(0, 1)
    Wend
    Set GetCellValue = HdrCell.Offset(0, -1)
End Function

Private Function CleanData(DataHeader As Range, DataCellIn As Range) As DataPair
    ' This will do the cleaning of data before doing things, to avoid blank
    Dim txtValue As Range, retVal As String, DataType As String
    If IsError(DataCellIn) Then
        CleanData.DataHeader = ""
        Exit Function
    Else
        If DataCellIn = "" Then
            Exit Function
        Else
            If DataCellIn = 0 And InStr(DataCellIn.Formula, "=") = 1 Then
                Exit Function
            Else
                CleanData.DataHeader = DataHeader
            End If
        End If
    End If
    Select Case DataHeader.Offset(-1)
    Case "TEXT":
        retVal = "'" & Left(StrQuoteReplace(DataCellIn), DataHeader.Offset(-2)) & "'"
    Case "MEMO":
        retVal = "'" & StrQuoteReplace(DataCellIn) & "'"
    Case "DATETIME":
        retVal = "#" & DataCellIn & "#"
    Case Else
        retVal = Val(DataCellIn)
    End Select
    CleanData.DataBit = retVal
End Function

Function MassReplace(ByVal Target As Range, Optional ReplaceTextArray As String = "OPT") As String
    Dim i As Long, xPos As Long, xRange As Range, tmpStr As String
    If AppCalculation Then Exit Function
    tmpStr = Target
    xPos = InStr(tmpStr, ReplaceTextArray & i + 1)
    Set xRange = Target.Offset(0, 2)
    While xPos <> 0
        tmpStr = Replace(tmpStr, ReplaceTextArray & i + 1, xRange.Offset(0, i))
        i = i + 1
        xPos = InStr(Target, ReplaceTextArray & i + 1)
    Wend
    MassReplace = tmpStr
    Set xRange = Nothing
End Function



