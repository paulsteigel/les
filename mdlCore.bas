Option Explicit

' Import
#If VBA7 Then
    Private Declare PtrSafe Function GetCurrentThreadId Lib "kernel32" () As LongPtr
    Private Declare PtrSafe Function SetDlgItemText Lib "user32" Alias "SetDlgItemTextW" (ByVal hDlg As LongPtr, ByVal nIDDlgItem As LongPtr, ByVal lpString As String) As LongPtr
    Private Declare PtrSafe Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As LongPtr, ByVal lpfn As LongPtr, ByVal hmod As LongPtr, ByVal dwThreadId As LongPtr) As LongPtr
    Private Declare PtrSafe Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As LongPtr) As LongPtr
#Else
    Private Declare Function GetCurrentThreadId Lib "kernel32" () As Long
    Private Declare Function SetDlgItemText Lib "user32" Alias "SetDlgItemTextW" (ByVal hDlg As Long, ByVal nIDDlgItem As Long, ByVal lpString As String) As Long
    Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
    Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
#End If

' for Unicode caption painting
#If VBA7 Then
    Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As LongPtr, ByVal nIndex As LongPtr, ByVal dwNewLong As LongPtr) As LongPtr
    Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
    Private Declare PtrSafe Function DefWindowProc Lib "user32.dll" Alias "DefWindowProcW" (ByVal hwnd As LongPtr, ByVal wMsg As LongPtr, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
#Else
    Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Private Declare Function DefWindowProc Lib "user32.dll" Alias "DefWindowProcW" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
#End If

Private Const WM_SETTEXT As Long = &HC

' Handle to the Hook procedure
#If VBA7 Then
    Private hHook As LongPtr
#Else
    Private hHook As Long
#End If
' Hook type
Private Const WH_CBT = 5
Private Const HCBT_ACTIVATE = 5
 
' Constants
Private Const IDOK = 1
Private Const IDCANCEL = 2
Private Const IDABORT = 3
Private Const IDRETRY = 4
Private Const IDIGNORE = 5
Private Const IDYES = 6
Private Const IDNO = 7

' Modify this code for English
Private StrYes As String
Private StrNo As String
Private StrOK As String
Private StrCancel As String

'=============================================
Global Const VnDate = "dd/mm/yyyy"
Public Enum KeyinMode   ' Ch?cho ph? c? nh? k?? ki?
    NumberType = 1      ' Ch?cho nh? s?    DateType = 2        ' Nh? ki? ng?y
    FormularType = 3    ' Ch?nh? k??ng th?   NumberOnlyType = 4
    FreeType = 5
End Enum

Public Type LocaleSetting
    DecimalSeparator As String * 1
    GroupNumber As String * 1
    DateLocale As String * 10
End Type

Public Type FormArgument
    AllowMultipleSelection As Boolean
    DataSource As String    ' Name of source range to be saved or loaded data from
    DataSetName As String   ' Name of object to be processed
    ErrorRange As String    ' Name to be used in case of blank
    ReadOnly As Boolean     ' Define whether to lock the list
    SpecialNote As String   ' Special instruction needed
    WrapOutput As Boolean   ' Wrap output in bracket for attention
    NotAllowSelection As String ' Do not allow selection with those contained this string
    DontAssignActiveCell As Boolean     ' Show or not show selected result
    SelectedItem As String  ' Return selected data
    ReturnIndexOnly As Boolean ' to convert return data
    ReturnDataOrder As String
    RowSource As Variant    ' raw range
End Type

' Messages variable
Global App_Title
Global AppLocale As LocaleSetting
Global frmObjectParameter As FormArgument
' for handling user event if there are any...
Global IndirectSetup As Boolean
Global AppStatus As Boolean
' for storing some temporary stuff
Global AppCalculation As Boolean

'=============================================
Function MsgBox(MessageTxt As String, Optional msgStyle As VbMsgBoxStyle) As VbMsgBoxResult
    Beep
    Dim iVal As VbMsgBoxStyle, msgBoxIcon As MsoAlertIconType, msgButton As MsoAlertButtonType
    iVal = msgStyle
    Select Case msgStyle
    Case 20, 19, 17, 16: ' Critical case
        iVal = iVal - 16
        msgBoxIcon = msoAlertIconCritical
    Case 36, 35, 33, 32: ' Question case
        iVal = iVal - 32
        msgBoxIcon = msoAlertIconQuery
    Case 52, 51, 49, 48: ' Exclamation case
        iVal = iVal - 48
        msgBoxIcon = msoAlertIconWarning
    Case 68, 67, 65, 64: ' Information case
        iVal = iVal - 64
        msgBoxIcon = msoAlertIconInfo
    End Select
  
    Select Case iVal
    Case 4:
        msgButton = msoAlertButtonYesNo
    Case 3:
        msgButton = msoAlertButtonYesNoCancel
    Case 1:
        msgButton = msoAlertButtonOKCancel
    Case 0:
        msgButton = msoAlertButtonOK
    End Select
    ' Set Hook
    hHook = SetWindowsHookEx(WH_CBT, AddressOf MsgBoxHookProc, 0, GetCurrentThreadId)
    ' Display the messagebox
    MsgBox = Application.Assistant.DoAlert(App_Title, MessageTxt, msgButton, msgBoxIcon, msoAlertDefaultFirst, msoAlertCancelDefault, True)
End Function
 
Private Function MsgBoxHookProc(ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    If lMsg = HCBT_ACTIVATE Then
        StrYes = "&C" & ChrW(243)
        StrNo = "&Kh" & ChrW(244) & "ng"
        'StrOK = ChrW(272) & ChrW(7891) & "&ng " & ChrW(253)
        StrOK = "Ch" & ChrW(7845) & "p nh" & ChrW(7853) & "&n"
        StrCancel = "&H" & ChrW(7911) & "y"

        SetDlgItemText wParam, IDYES, StrConv(StrYes, vbUnicode)
        SetDlgItemText wParam, IDNO, StrConv(StrNo, vbUnicode)
        SetDlgItemText wParam, IDCANCEL, StrConv(StrCancel, vbUnicode)
        SetDlgItemText wParam, IDOK, StrConv(StrOK, vbUnicode)
        ' Release the Hook
        UnhookWindowsHookEx hHook
    End If
    MsgBoxHookProc = False
End Function

Function MSG(MsgName As String) As String
    ' This function will return expected string for better userinterface
    MSG = "False"
    Dim MyCell As Range, FoundObj As Boolean
    Set MyCell = ThisWorkbook.Sheets("Data").Range("MSG_ID_START").Offset(1)
    While Not FoundObj
        If Len(Trim(MyCell)) <= 0 Then
            FoundObj = True
        Else
            If MyCell = MsgName Then
                FoundObj = True
                MSG = MyCell.Offset(, 1)
            End If
        End If
        Set MyCell = MyCell.Offset(1)
    Wend
End Function

Sub WriteLog(ErrDesc As String, Optional LogFileName As String = "Error.txt", Optional KillIfExist As Boolean = False)
    Dim txtString As String, FileNames As String
    FileNames = LogFileName
    
    txtString = ErrDesc
    Dim UnicodeFile As Boolean
    
    Const ForAppending = 8
    UnicodeFile = True
    
    Dim fso As Object, ts As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    ' check if the file exist
    If KillIfExist Then If FileOrDirExists(LogFileName, True) Then Kill LogFileName
    
    Set ts = fso.OpenTextFile(FileNames, ForAppending, True, UnicodeFile)
    ts.WriteLine txtString

    ts.Close
    
    Set ts = Nothing
    Set fso = Nothing
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

Sub GetObjSource(ObjControl As Control, Optional ParrentID As String = "", _
    Optional ColCount As Long = 2, Optional RowSourceName As String = "", _
    Optional SearchCell As String = "", Optional ResourceText As String, _
    Optional ReturnIndexOnly As Boolean = False)
    'Fill in Commbo or listbox with region table
    Err.Clear
    On Error GoTo err_handler
    Dim arr() As Variant
    If RowSourceName <> "" Then
        ' This will die when there is only one cell...
        If Range(RowSourceName).Cells.Count = 1 Then
            Dim tmpArr(1, 1)
            tmpArr(1, 1) = Range(RowSourceName)
            arr = tmpArr
        Else
            arr = Range(RowSourceName)
        End If
    Else
        arr = Range("tblRegions")
    End If
    Dim R As Long
    With ObjControl
        .ColumnCount = ColCount
        .ColumnWidths = IIf(ColCount = 1, .Width - 10, "0;" & .Width - 10)
        .Clear
        
        For R = 1 To UBound(arr, 1) ' First array dimension is rows.
            If ParrentID = "" And RowSourceName <> "" Then
                If arr(R, 1) <> "" And Not arr(R, 1) Like "<<*" Then
                    .AddItem arr(R, 1)
                    ResourceText = ResourceText & "[" & arr(R, 1) & "]"
                    If ColCount = 2 Then
                        .List(.ListCount - 1, 1) = arr(R, 2)
                    End If
                End If
                
                If IIf(ReturnIndexOnly, Val(arr(R, 1)), arr(R, 1)) = SearchCell And Trim(SearchCell) <> "" Then
                    If Not arr(1, 1) Like "<<*" Then
                        .Selected(R - 1) = True
                    Else
                        .Selected(R - 2) = True
                    End If
                End If
            Else
                If arr(R, 3) = ParrentID Then
                    If ColCount = 2 Then
                        .AddItem arr(R, 1)
                        .List(.ListCount - 1, 1) = arr(R, 4)
                    Else
                        .AddItem arr(R, 4)
                    End If
                End If
            End If
        Next R
    End With
err_handler:
    If Err.Number <> 0 Then
        Debug.Print Err.description
        ObjControl.Clear
        Err.Clear
    End If
End Sub

Function GetAbrFromText(TextString As String) As String
    ' To get just first letter of the text string
    Dim i As Long, rStr As String
    TextString = Trim(TextString)
    rStr = Left(Trim(TextString), 1)
    i = InStr(TextString, " ")
    If i <= 0 Then
        rStr = rStr & "BL"
        GoTo ExitFunc
    End If
    While i > 0
        rStr = rStr & Mid(TextString, i + 1, 1)
        i = InStr(i + 1, TextString, " ")
    Wend
ExitFunc:
    GetAbrFromText = rStr
End Function

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

Sub ValidateKeycode(kCde As MSForms.ReturnInteger, Optional iNum As KeyinMode = 1)
    Select Case iNum
    Case 1:
        If kCde <> vbKeyBack Then
            If InStr("0123456789", ChrW(kCde)) = 0 Then
                If InStr(".", ChrW(kCde)) <> 0 Then
                    kCde = AscW(AppLocale.DecimalSeparator)
                Else
                    kCde = 0
                End If
            End If
        End If
    Case 2:
        If InStr("0123456789/", ChrW(kCde)) = 0 And kCde <> vbKeyBack Then kCde = 0
    Case 3:
        If InStr("()+-*/0123456789", ChrW(kCde)) = 0 And kCde <> vbKeyBack Then kCde = 0
    Case 4:
        If InStr("0123456789", ChrW(kCde)) = 0 And kCde <> vbKeyBack Then kCde = 0
    Case 5:
        ' enable free type
    End Select
End Sub

Sub NoCutAction(kCde As MSForms.ReturnInteger, ByVal ShiftKey As Integer)
    ' preventing user from pasting
    If (ShiftKey And 2) And (kCde = Asc("V")) Then kCde = 0
End Sub

Sub SetUnicodeCaption(ByVal frm As UserForm, ByVal UnicodeString As String)
    #If VBA7 Then
        Dim hwnd As LongPtr
    #Else
        Dim hwnd&
    #End If
    hwnd = FindWindow("ThunderDFrame", frm.Caption)
    DefWindowProc hwnd, WM_SETTEXT, 0, StrPtr(UnicodeString)
End Sub

Function FalseInput(CtrlName As Control) As Boolean
    On Error Resume Next
    Dim tData As String
    If CtrlName = "" Then Exit Function
    If Not IsDate(CtrlName) Then GoTo tCont
    tData = InputDate(CtrlName)
    If Not tData Like "12:00*" Then Exit Function
tCont:
    CtrlName = ""
    CtrlName.SetFocus
    FalseInput = True
End Function

Function InputDate(iDateStr As Variant) As Date
    ' Send data piece from database to console
    ' default the data will from db to console, output shall be formated
    ' input shall be converted back to serial date
    Dim iStr As String, iSpliter As Variant
    
    On Error GoTo ErrHandler
    iSpliter = Split(iDateStr, "/")
    If UBound(iSpliter) < 2 Then GoTo ErrHandler
    ' Now we have to see what locale we are now at
    InputDate = DateSerial(iSpliter(2), iSpliter(0), iSpliter(1))
ErrHandler:
End Function

Property Get SetConfig(ObjName As String, FrmObj As UserForm) As String
    Dim j As Long
    Dim MyCell As Range, tmpCell As Range
    Set MyCell = Range("tblFormConfig").Offset(1)
    '1. Search for area to keep data
    While MyCell <> "" And MyCell <> ObjName
        Set MyCell = MyCell.Offset(1)
    Wend
    While MyCell = ObjName
        Select Case MyCell.Offset(, 1)
        Case 0:
            SetCaption FrmObj.Controls(CStr(MyCell.Offset(, 2))), MyCell.Offset(, 3), MyCell.Offset(, 4)
        Case 2, 4: ' Just set tag value
            FrmObj.Controls(CStr(MyCell.Offset(, 2))).Tag = MyCell.Offset(, 4)
            SetCaption FrmObj.Controls(CStr(MyCell.Offset(, 2))), MyCell.Offset(, 3), MyCell.Offset(, 4)
        Case 3:
            ' for form caption
            SetConfig = MyCell.Offset(, 3)
        Case Else
            Set tmpCell = MyCell
            For j = 0 To FrmObj.Controls(CStr(MyCell.Offset(, 2))).Pages.Count - 1
                FrmObj.Controls(CStr(MyCell.Offset(, 2))).Pages(j).Caption = tmpCell.Offset(, j + 3)
            Next
        End Select
        Set MyCell = MyCell.Offset(1)
    Wend
    Set MyCell = Nothing
    Set tmpCell = Nothing
End Property

Private Sub SetCaption(MyObj As Object, iCaption As String, Optional ControlTipStr As String = "")
    If iCaption <> "" Then MyObj.Caption = iCaption
    If ControlTipStr <> "" Then MyObj.ControlTipText = ControlTipStr
End Sub

Private Function GetCaption(obj As Object) As String
    On Error GoTo ErrHandler
    GetCaption = obj.Caption
ErrHandler:
End Function

Sub ShowAll(SheetObj As Worksheet)
    On Error Resume Next
    SheetObj.ShowAllData
End Sub

Property Let AppSetting(PropertyID As String, PropertySetting As String)
    ' This function will return expected string for better userinterface
        'Set back all needed variable
    App_Title = Range("APP_TITLE")
    
    Dim MyCell As Range, FoundObj As Boolean
    Set MyCell = ThisWorkbook.Sheets("Data").Range("MSG_ID_START").Offset(1)
    While Not FoundObj
        If Len(Trim(MyCell)) <= 0 Then
            MyCell.Offset(, 1) = PropertySetting
            MyCell = PropertyID
            FoundObj = True
        Else
            If MyCell = PropertyID Then
                FoundObj = True
                MyCell.Offset(, 1) = PropertySetting
            End If
        End If
        Set MyCell = MyCell.Offset(1)
    Wend
End Property

Property Get AppSetting(PropertyID As String) As String
    ' This function will return expected string for better userinterface
    'Set back all needed variable
    App_Title = Range("APP_TITLE")
        
    Dim MyCell As Range, FoundObj As Boolean
    Set MyCell = ThisWorkbook.Sheets("Data").Range("MSG_ID_START").Offset(1)
    While Not FoundObj
        If Len(Trim(MyCell)) <= 0 Then
            FoundObj = True
        Else
            If MyCell = PropertyID Then
                FoundObj = True
                AppSetting = MyCell.Offset(, 1)
            End If
        End If
        Set MyCell = MyCell.Offset(1)
    Wend
End Property

Function StrQuoteReplace(strValue) As String
    Dim sTemp As String
    If IsError(strValue.Value) Then
       '~~> Check if it is a 2029 error
       If strValue.Value = CVErr(2029) Then
           '~~> Get the cell contents
           sTemp = Trim(strValue.Formula)
           If InStr(sTemp, "#NAME?") <> 0 Then GoTo CleanUp
           '~~> Remove =/-
           Do While InStr("+=-*/", Left(sTemp, 1)) <> 0
               sTemp = Trim(Mid(sTemp, 2))
           Loop
           '~~> Either put it in back in the cell or do
           '~~> what ever you want with sTemp
           strValue.Formula = sTemp
       End If
    End If
    StrQuoteReplace = Replace(strValue, "'", "''")
CleanUp:
End Function

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

