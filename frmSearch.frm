VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSearch 
   Caption         =   "Search box"
   ClientHeight    =   6675
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8925
   OleObjectBlob   =   "frmSearch.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim FldSearch As String, fldName As String
Dim FixFieldName As String
Dim InProgress As Boolean
Dim dbs As New clsDbConnection
Dim CurFilter As String

Private Sub cbCommune_Change()
    LoadComboData cbVillage, "Select VillageName from tblVillage Where CommuneID=" & cbCommune.List(cbCommune.ListIndex, 0) & ";"
End Sub

Private Sub LoadComboData(cbObj As ComboBox, SqlData As String, Optional HasIndex As Boolean = False)
    Dim rs As New ADODB.Recordset, i As Long
    Set rs = dbs.GetRecordSet(SqlData)
    cbObj.Clear
    If Not HasIndex Then
        cbObj.ColumnCount = 1
        cbObj.ColumnWidths = "1"
    Else
        cbObj.ColumnCount = 2
        cbObj.ColumnWidths = "0;1"
    End If
    While Not rs.EOF
        If HasIndex Then
            cbObj.AddItem
            cbObj.List(i, 0) = rs.Fields(0)
            cbObj.List(i, 1) = rs.Fields(1)
            i = i + 1
        Else
            cbObj.AddItem rs.Fields(0)
        End If
        rs.MoveNext
    Wend
    rs.Close
End Sub

Private Sub cbVillage_Change()
    LoadSQL "txt_village='" & cbVillage & "'", True
End Sub

Private Sub cmdExport_Click()
    ' this will do the export to excel file
    Export2Excel CurFilter
    Unload Me
End Sub

Private Sub cmdMassPrint_Click()
    ' This will trigger mass printing option
    Dim txtFilterStr As String
    txtFilterStr = GetFilterString()
    If txtFilterStr <> "" Then
        MassPrintOpt False, txtFilterStr
        Unload Me
    Else
        MsgBox MSG("MSG_NONE_SELECTED"), vbInformation
        lstResult.SetFocus
    End If
End Sub

Private Sub cmdPrint_Click()
    ' This will export the selected value to Excel, then print
    ShowFormAndPrint "Form_ID=" & lstResult.List(lstResult.ListIndex, 0)
    Unload Me
End Sub

Private Sub lstOption_Change()
    If InProgress Then Exit Sub
    UpdateSearchDb lstOption.List(lstOption.ListIndex, 0), lstOption.Selected(lstOption.ListIndex)
End Sub

Private Function GetFilterString() As String
    Dim FltString As String, i As Long
    With lstResult
        For i = 0 To .ListCount - 1
            If .Selected(i) Then FltString = FltString & "," & .List(i, 0)
        Next
    End With
    If FltString <> "" Then GetFilterString = "Form_ID IN (" & Mid(FltString, 2) & ")"
End Function

Private Sub BuildSearchFilter()
    Dim i As Long
    FldSearch = ""
    fldName = ""
    For i = 0 To lstOption.ListCount - 1
        If lstOption.Selected(i) Then
            Select Case lstOption.List(i, 2)
            Case "TEXT":
                FldSearch = FldSearch & " OR " & lstOption.List(i, 0) & " LIKE '*[OPT" & i & "]*'"
            Case Else
                FldSearch = FldSearch & " OR " & lstOption.List(i, 0) & " IN (Val('[OPT" & i & "]'))"
            End Select
            fldName = fldName & "," & lstOption.List(i, 0)
        End If
    Next
    FldSearch = Mid(FldSearch, 4)
    fldName = Mid(fldName, 2)
End Sub

Private Sub UpdateSearchDb(FieldName As String, FieldValue As Boolean)
    ' quickly update seacrh parameter using the entered value
    Dim SqlTxt As String
    SqlTxt = "Update tblFieldMap Set UseInSearch = " & FieldValue & " where FieldName='" & FieldName & "' And TableName = 'tblFormInfor';"
    dbs.ExecuteSQL SqlTxt
End Sub

Private Sub LoadSQL(txtValue As String, Optional DirectCall As Boolean = False)
    Dim rs As New ADODB.Recordset, i As Long, j As Long, prm As Variant, SrcTxt As String, SqlTxt As String
    ' Now process parameter first
    If txtValue = "" Then Exit Sub
    
    If Not DirectCall Then
        prm = Split(txtValue, ";")
        SrcTxt = FldSearch
        For i = 0 To UBound(prm)
            SrcTxt = Replace(SrcTxt, "[OPT" & i & "]", prm(i))
        Next
        i = 0
        While InStr(SrcTxt, "[OPT") <> 0
            SrcTxt = Replace(SrcTxt, "[OPT" & i & "]", prm(0))
            i = i + 1
        Wend
        SqlTxt = "Select Form_ID," & FixFieldName & " from tblFormInfor WHERE " & SrcTxt & ";"
        CurFilter = SrcTxt
    Else
        SqlTxt = "Select Form_ID," & FixFieldName & " from tblFormInfor WHERE " & txtValue & ";"
        CurFilter = txtValue
    End If
    
    Set rs = dbs.GetRecordSet(SqlTxt)
    
    lstResult.ColumnCount = rs.Fields.Count
    lstResult.ColumnWidths = "0;70;100;100;0"
    lstResult.ColumnHeads = True
    lstResult.Clear
    While Not rs.EOF
        With lstResult
            .AddItem
            For i = 0 To rs.Fields.Count - 1
                .List(j, i) = IIf(IsNull(rs.Fields(i)), "", rs.Fields(i))
            Next
        End With
        j = j + 1
        rs.MoveNext
    Wend
    rs.Close
End Sub

Private Sub lstResult_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    cmdPrint_Click
End Sub

Private Sub txtParams_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn And txtParams.Value <> "" Then
        ' Load listbox with typing value
        BuildSearchFilter
        LoadSQL txtParams.Value
        txtParams.SetFocus
    End If
End Sub

Private Sub UserForm_Initialize()
    'Load field list
    Dim rs As New ADODB.Recordset, rsTxt As String, i As Long
    dbs.ConnectDatabase AppDatabase
    Set rs = dbs.GetRecordSet("Select FieldName, FieldCaption, Selected, UseInSearch, FieldType FROM tblFieldMap " & _
        "WHERE TableName='tblFormInfor' AND FieldCaption <> '-UNUSED-' " & _
        "ORDER BY UseInSearch ASC, Selected ASC, FieldName ASC;", True)
    lstOption.ColumnCount = 3
    lstOption.ColumnWidths = "0;" & lstOption.Width - 20 & ";0"
    InProgress = True
    DoEvents
    While Not rs.EOF
        With lstOption
            .AddItem
            .List(i, 0) = rs.Fields(0)
            .List(i, 1) = rs.Fields(1)
            .List(i, 2) = IIf(IsNull(rs.Fields(4)), "INTEGER", rs.Fields(4))
            If rs.Fields(3) Then .Selected(i) = True
            i = i + 1
            ' Just for field name to display
            If rs.Fields(2) Then FixFieldName = FixFieldName & "," & rs.Fields(0)
            
        End With
        rs.MoveNext
    Wend
    lstResult.ColumnCount = i
    lstResult.ListStyle = fmListStyleOption
    lstResult.MultiSelect = fmMultiSelectExtended
    lstResult.ColumnHeads = False
        
    InProgress = False
    FixFieldName = Mid(FixFieldName, 2)
    
    rs.Close
    LoadComboData cbCommune, "Select ID, CommuneName from tblCommune;", True
End Sub

Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txtParams.SetFocus
    End If
End Sub

Private Sub UserForm_Terminate()
    Set dbs = Nothing
End Sub
