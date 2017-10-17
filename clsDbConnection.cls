VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDbConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' Duong dan den csdl
Private dbPath As String
' Doi tuong de ket noi voi CSDL
Private dbObject As ADODB.Connection

' Cach thuc ket noi
Private DatabaseName As String

Property Set OpenConnection(ConnObj As ADODB.Connection)
    Set dbObject = ConnObj
End Property

Property Get HeaderString(Id As Long, tblName As String, SearchField As String, ReturnField As String) As String
    ' Return header column
    Dim rs As Object
    Set rs = GetRecordSet("Select " & ReturnField & " from " & tblName & " Where [" & SearchField & "]=" & Id & ";")
NextStep:
    If rs.EOF Then
        HeaderString = "N/A"
    Else
        HeaderString = rs.Fields(0)
    End If
    
    rs.Close
    Set rs = Nothing
End Property

Property Get HeaderCellID(HdrValueObj, SheetName As String) As Long
    ' Check and return HeaderCellID
    Dim rs As Object, HdrValue As String
    HdrValue = StrQuoteReplace(HdrValueObj)
    
    Set rs = GetRecordSet("Select ID from tblFieldMap Where SheetName='" & SheetName & "' And FieldName='" & HdrValue & "';")
    If rs.EOF Then
        Call ExecuteSQL("Insert Into tblFieldMap(SheetName,FieldName) Values('" & SheetName & "','" & HdrValue & "');")
    Else
        GoTo NextStep
    End If
    Set rs = GetRecordSet("Select ID from tblFieldMap Where SheetName='" & SheetName & "' And FieldName='" & HdrValue & "';")
NextStep:
    HeaderCellID = rs.Fields("ID")
    
    rs.Close
    Set rs = Nothing
End Property

Sub CloseConnection()
    dbObject.Close
    Set dbObject = Nothing
End Sub

Sub CreateDb(Optional dbName As String = "Data.mdb")
    ' Khoi tao CSDL Access
    dbPath = ThisWorkbook.Path & "\" & dbName
        
    ' Kiem tra xem CSDL da co khong, neu co thi xoa di
    If FileOrDirExists(dbPath, True) Then Kill dbPath
    
    CreateDbAdodb dbPath
   
    ' Khoi tao cac bang so lieu
    CreateTable "tblFormInfor", False
    CreateTable "tbl_hhld_member_details"
    CreateTable "tbl_hhld_members"
End Sub

Function DMax(fldName As String, tblName As String) As Long
    'On Error Resume Next
    DMax = GetRecordSet("Select Max(" & fldName & ") as MaxRecord from " & tblName & ";").Fields(0)
    If Err.Number <> 0 Then Err.Clear
End Function

'======================================
' Phan chen so lieu vao bang thi giong
' nhau giua DAO va ADODB

'-------------TAO BANG CSDL-------------
' Nhom cac thu tuc lam viec voi ADODB
Private Sub CreateTable(tblName As String, Optional CreateUplink As Boolean = True)
    Dim SqlTxt As String
    ' Create FormInfor data
    Dim i As Long, ptrCell As Range, OldFieldName As String
    Set ptrCell = Range(tblName).Offset(0, 1)
    While ptrCell <> ""
        If ptrCell.Offset(-2) <> "" Then
            SqlTxt = SqlTxt & ptrCell & " " & ptrCell.Offset(-1) & "(" & ptrCell.Offset(-2) & "), "
        Else
            SqlTxt = SqlTxt & ptrCell & " " & ptrCell.Offset(-1) & ", "
        End If
        Set ptrCell = ptrCell.Offset(0, 1)
    Wend
    SqlTxt = "Create Table " & tblName & "(Form_ID AUTOINCREMENT," & IIf(CreateUplink, "hhld_ims_code LONG, ", "") & Left(SqlTxt, Len(SqlTxt) - 2) & ");"
    Debug.Print SqlTxt
    
    Call ExecuteSQL(SqlTxt)
End Sub

Private Sub AlterRelation(tblLeft As String, tblRight As String, tblLeftCol As String, tblRightCol As String, _
    Optional SkipPrimaryKey As Boolean = False, _
    Optional NormalRel As Boolean = False)
    'For group table
    If Not SkipPrimaryKey Then Call ExecuteSQL("ALTER TABLE " & tblLeft & " ADD PRIMARY KEY (" & tblLeftCol & ");")
    Call ExecuteSQL("ALTER TABLE " & tblRight & " ADD CONSTRAINT " & tblLeft & "_" & tblRight & "_Relationship " _
        & "FOREIGN KEY (" & tblRightCol & ") " _
        & "REFERENCES " & tblLeft & " (" & tblLeftCol & ") " & IIf(NormalRel, ";", "ON UPDATE CASCADE ON DELETE CASCADE;"))
End Sub

Function ExecuteSQL(SQLStr As String) As Boolean
    'On Error GoTo ErrHandler
    Debug.Print SQLStr
    dbObject.Execute SQLStr
    ExecuteSQL = True
ErrHandler:
End Function
'================END======================

'======================================
' Nhom cac thu tuc lam viec voi ADODB
Private Sub CreateDbAdodb(FileName As String)
    Dim Catalog As Object
    Set Catalog = CreateObject("ADOX.Catalog")
    Catalog.Create "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & FileName & ";"
    Set Catalog = Nothing
    ConnectDatabase FileName
End Sub
'================END======================

Sub ConnectDatabase(Optional dbsPath As String)
    ' This would create a dbConnection to database
    Set dbObject = New ADODB.Connection
    With dbObject
        .Provider = "Microsoft.ACE.OLEDB.12.0"
        .Mode = adModeReadWrite
        .ConnectionString = "Data Source=" & ThisWorkbook.Path & "\" & dbsPath
        .Open
    End With
End Sub

' Khoi tao ket noi voi recordset -
' tuy thuoc vao kieu ket noi nhe
Property Get GetRecordSet(SqlText As String, Optional DirectConnect As Boolean = False, Optional EditMode As Boolean = False) As Object
    'On Error GoTo ErrHandler
    Debug.Print SqlText
    Set GetRecordSet = CreateObject("ADODB.Recordset")
    If DirectConnect Then
        If EditMode Then
            GetRecordSet.Open SqlText, dbObject, 0, 3
        Else
            GetRecordSet.Open SqlText, dbObject
        End If
    Else
        GetRecordSet.Open Replace(SqlText, "*", "%"), dbObject
    End If

ErrHandler:
End Property

Property Get TableExist(tblName As String) As Boolean
    On Error GoTo ErrHandler
    Dim Rcs As Object
    Set Rcs = Me.GetRecordSet(tblName)
    TableExist = True
    Rcs.Close
ErrHandler:
End Property

Sub DropTable(tblName As String)
    Me.ExecuteSQL "DROP TABLE " & tblName & ";"
End Sub
'================END======================

Private Sub Class_Terminate()
    On Error Resume Next
    dbObject.Close
    Set dbObject = Nothing
End Sub

Sub CreateQuery(QueryName As String, SQLStr As String)
     '  1/2/06 - Carl for : Oz
     '  requies references : MS Active X Data Objects libs and MS Active X Ext
  
    Dim cat As Object
    Dim cmd As Object
    
    If dbObject.State <> 1 Then
        MsgBox ("Problem with connection")
        Exit Sub
    End If
     
    Set cat = CreateObject("ADOX.Catalog")
    Set cmd = CreateObject("ADODB.Command")
     
    cat.ActiveConnection = dbObject
     
    cmd.CommandText = SQLStr
    cat.Views.Append QueryName, cmd
     
    Set cat = Nothing
    Set cmd = Nothing
End Sub

