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
Public Enum KeyinMode   ' ChØ cho phÐp cËp nhËt ký tù ®ång kiÓu
    NumberType = 1      ' ChØ cho nhËp sè
    DateType = 2        ' NhËp kiÓu ngµy
    FormularType = 3    ' ChØ nhËp ký tù c«ng thøc
    NumberOnlyType = 4
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
Global SheetObjName As String
Global App_Title
Global ExternalLoad As Boolean
Global CurrentWorkBook As Workbook

Global AppLocale As LocaleSetting
Global ShapedLoaded As Boolean
Global frmObjectParameter As FormArgument
' for handling user event if there are any...
Global IndirectSetup As Boolean
Global AppStatus As Boolean
' for storing some temporary stuff
Global TempString As String
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

Sub ToggleFilterKey()
    ' This shall help to disable filter
    If Not ActiveSheet.FilterMode Then
        QuickFilter
    Else
        ShowAll ActiveSheet
        ' Repair sheet if neccessary
        RepairSheet ActiveSheet.Name
    End If
End Sub

Sub InsertVillage()
    If ActiveSheet.Name <> "II.2.A" Then Exit Sub
    If MsgBox(MSG("MSG_ADD_VILLAGE"), vbQuestion + vbYesNo) = vbYes Then
        Dim theRange As Range
        Set theRange = AddRevVillage(1)
        ShowOff
        ModifyColumns
        ShowOff True
        ' Get to Data table for putting village name
        Sheets("Data").Activate
        theRange.Activate
    End If
End Sub

Sub RemoveVillage()
    If ActiveSheet.Name <> "II.2.A" Then Exit Sub
    ' if just remain 2 colums - dont allow removal
    If Range("RNG_II2A").Column - Range("RNG_IIAST").Column = 6 Then
        MsgBox MSG("MSG_REMOVE_VILLAGE_DISALLOW"), vbCritical
        Exit Sub
    End If
    If MsgBox(Replace(MSG("MSG_REMOVE_VILLAGE"), "%s%", Sheet4.Range("RNG_II2A").Offset(0, -1)), vbQuestion + vbYesNo) = vbYes Then
        Call AddRevVillage(-1)
        ShowOff
        ModifyColumns -1
        ShowOff True
    End If
End Sub

Private Function AddRevVillage(param As Long) As Range
    Dim rng As Range
    Set rng = ThisWorkbook.Sheets("Data").Range("tblVillageStart")
    While Len(Trim(rng)) > 0
        Set rng = rng.Offset(1)
    Wend
    ' Now I am at the last point
    If param < 0 Then
        rng.Offset(-2) = rng.Offset(-1)
        rng.Offset(-1) = ""
    Else
        rng = MSG("MSG_VIL_NEW")
        Set AddRevVillage = rng
    End If
    Set rng = Nothing
End Function

Sub ShowAll(SheetObj As Worksheet)
    On Error Resume Next
    ProtectWorkSheet(SheetObj) = False
    SheetObj.ShowAllData
    ProtectWorkSheet(SheetObj) = True
    SheetObj.Range("A8").Activate
End Sub

Function GetDate(txtString As String) As Date
    ' This will help converting Vietnamese date to English date
    Dim arr As Variant
    arr = Split(Replace(txtString, "'", ""), "/")
    GetDate = DateSerial(arr(2), arr(1), arr(0))
End Function

Function FormatDate(GivenDate As Date, Optional FormatType = VnDate, Optional DontSurpress As Boolean = False) As String
    ' This will override problematic date formating in Excel
    FormatDate = IIf(DontSurpress, "", "'") & Format(GivenDate, FormatType)
End Function

Sub UpdateChange(ByVal Target As Range, ByVal WrkSheet As Worksheet, Optional ByVal CmdObject As Object = Nothing)
    ' this process is to make the update buton...
    'If isFrmLoaded Then Exit Sub
    ShowOff
    With Target
        ' Check if a shape has been created or not?
        If (.Row > 6 And .Row <= 555) And (.Column = 12 Or .Column = 14) Then
            CmdObject.Top = ActiveCell.Top + (ActiveCell.Height - CmdObject.Height)
            CmdObject.Left = ActiveCell.Left - CmdObject.Width
            CmdObject.Visible = msoTrue
        Else
            CmdObject.Visible = msoFalse
        End If
    End With
    ShowOff True
End Sub

Sub ToggleCutCopyAndPaste(Allow As Boolean)
     'Activate/deactivate cut, copy, paste and pastespecial menu items
    Call EnableMenuItem(21, Allow) ' cut
    Call EnableMenuItem(19, Allow) ' copy
    Call EnableMenuItem(22, Allow) ' paste
    Call EnableMenuItem(755, Allow) ' pastespecial
     
     'Activate/deactivate drag and drop ability
    Application.CellDragAndDrop = Allow
     
     'Activate/deactivate cut, copy, paste and pastespecial shortcut keys
    With Application
        Select Case Allow
        Case Is = False
            .OnKey "^c", "CutCopyPasteDisabled"
            .OnKey "^v", "CutCopyPasteDisabled"
            .OnKey "^x", "CutCopyPasteDisabled"
            .OnKey "+{DEL}", "CutCopyPasteDisabled"
            .OnKey "^{INSERT}", "CutCopyPasteDisabled"
        Case Is = True
            .OnKey "^c"
            .OnKey "^v"
            .OnKey "^x"
            .OnKey "+{DEL}"
            .OnKey "^{INSERT}"
        End Select
    End With
End Sub
 
Sub EnableMenuItem(ctlId As Integer, Enabled As Boolean)
     'Activate/Deactivate specific menu item
    Dim cBar As CommandBar
    Dim cBarCtrl As CommandBarControl
    For Each cBar In Application.CommandBars
        If cBar.Name <> "Clipboard" Then
            Set cBarCtrl = cBar.FindControl(Id:=ctlId, recursive:=True)
            If Not cBarCtrl Is Nothing Then cBarCtrl.Enabled = Enabled
        End If
    Next
End Sub
 
Sub CutCopyPasteDisabled()
     'Inform user that the functions have been disabled
    MsgBox "Sorry!  Cutting, copying and pasting have been disabled in this workbook!"
    'Selection.Copy
End Sub

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

Property Let AppSetting(PropertyID As String, PropertySetting As String)
    ' This function will return expected string for better userinterface
    If CurrentWorkBook Is Nothing Then
        'Set back all needed variable
        App_Title = Range("APP_TITLE")
        Set CurrentWorkBook = ThisWorkbook
    End If
    Dim MyCell As Range, FoundObj As Boolean
    Set MyCell = CurrentWorkBook.Sheets("Data").Range("MSG_ID_START").Offset(1)
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
    If CurrentWorkBook Is Nothing Then
        'Set back all needed variable
        App_Title = "LES data application..."
        Set CurrentWorkBook = ThisWorkbook
    End If
    Dim MyCell As Range, FoundObj As Boolean
    Set MyCell = CurrentWorkBook.Sheets("Data").Range("MSG_ID_START").Offset(1)
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


Sub ShowSelectForm()
    ' This shall display a form for selecting something
    Dim isSelected As Boolean, tmpStr As String
    tmpStr = vbLf & MSG("MSG_ARROW_TO_SELECT")

    Select Case ActiveSheet.Name
    Case "Part_B.1"
        With ActiveCell
            If .Row >= 6 And .Row <= 17 Then
                With frmObjectParameter
                    .SelectedItem = ActiveCell
                    .ReadOnly = False
                    .WrapOutput = False
                    .ReturnIndexOnly = True
                End With
                Select Case .Column
                Case 2:
                    With frmObjectParameter
                        .SpecialNote = MSG("MSG_SEL_HHLD_MEMBER") & tmpStr
                        .DataSetName = MSG("MSG_SEL_HHLD_MEMBER")
                        .DataSource = "lst_hhld_member"
                        .ReturnIndexOnly = False
                        .ReadOnly = True
                    End With
                Case 3, 4:
                    With frmObjectParameter
                        .SpecialNote = MSG("MSG_SEL_SKILL_OWN") & tmpStr
                        .DataSetName = MSG("MSG_SELECT_SKILL_OWN")
                        .DataSource = "lst_skills_owned"
                        .ReturnDataOrder = "3,4"
                        .RowSource = Range(ActiveCell, ActiveCell.Offset(0, 1))
                    End With
                Case 5, 6:
                    With frmObjectParameter
                        .SpecialNote = MSG("MSG_SEL_COURSE_JOINED") & tmpStr
                        .DataSetName = MSG("MSG_SEL_COURSE")
                        .DataSource = "lst_course_joined"
                        .ReturnDataOrder = "5,6"
                    End With
                Case 8:
                    With frmObjectParameter
                        .SpecialNote = MSG("MSG_SKILL_EVAL") & tmpStr
                        .DataSetName = MSG("MSG_SEL_SKILL_EVAL")
                        .DataSource = "lst_skill_eval"
                        .ReadOnly = True
                    End With
                Case 10:
                    With frmObjectParameter
                        .SpecialNote = MSG("MSG_SKILL_EVAL") & tmpStr
                        .DataSetName = MSG("MSG_SEL_SKILL_EVAL")
                        .DataSource = "lst_status_yes_no"
                        .ReadOnly = True
                    End With
                Case 11, 12  ' Location for activity
                    With frmObjectParameter
                        .SpecialNote = MSG("MSG_SEL_SSI_TYPE") & tmpStr
                        .DataSetName = MSG("MSG_SELECT_SSI_TYPE")
                        .DataSource = "lst_SSI_prog"
                        .ReturnDataOrder = "11,12"
                    End With
                Case Else
                    isSelected = True
                End Select
                If Not isSelected Then frmSelect.Show vbModal
            End If
        End With
    
    Case "Part_A.2"
        With ActiveCell
            If .Row >= 4 And .Row <= 19 Then
                With frmObjectParameter
                    .SelectedItem = ActiveCell
                    .ReadOnly = True
                    .WrapOutput = False
                    .ReturnIndexOnly = True
                End With

                Select Case .Column
                Case 5:    ' Location for activity
                    With frmObjectParameter
                        .SpecialNote = MSG("MSG_GENDER") & tmpStr
                        .DataSetName = MSG("MSG_SEL_GENDER")
                        .DataSource = "lst_gender"
                    End With
                Case 8:
                    With frmObjectParameter
                        .SpecialNote = MSG("MSG_REL_TYPE") & tmpStr
                        .DataSetName = MSG("MSG_REL_TYPE")
                        .DataSource = "lst_rel_type"
                    End With
                Case 10:
                    With frmObjectParameter
                        .SpecialNote = MSG("MSG_EDU_LEVEL") & tmpStr
                        .DataSetName = MSG("MSG_EDU_LEVEL")
                        .DataSource = "lst_edu_level"
                    End With
                Case 12:
                    With frmObjectParameter
                        .SpecialNote = MSG("MSG_JOB_TYPE_MAJOR") & tmpStr
                        .DataSetName = MSG("MSG_JOB_TYPE")
                        .DataSource = "lst_jobs_type"
                    End With
                Case 14:
                    With frmObjectParameter
                        .SpecialNote = MSG("MSG_JOB_TYPE_MINOR") & tmpStr
                        .DataSetName = MSG("MSG_JOB_TYPE")
                        .DataSource = "lst_jobs_type"
                    End With
                Case 16:
                    With frmObjectParameter
                        .SpecialNote = MSG("MSG_JOB_EVAL") & tmpStr
                        .DataSetName = MSG("MSG_JOB_EVAL")
                        .DataSource = "lst_job_status"
                    End With
                Case 19, 21:
                    With frmObjectParameter
                        .SpecialNote = MSG("MSG_SKILL_EVAL") & tmpStr
                        .DataSetName = MSG("MSG_SEL_SKILL_EVAL")
                        .DataSource = "lst_reallocate_reason"
                    End With
                Case 23:
                    With frmObjectParameter
                        .SpecialNote = MSG("MSG_SKILL_EVAL") & tmpStr
                        .DataSetName = MSG("MSG_SEL_SKILL_EVAL")
                        .DataSource = "lst_status_yes_no"
                    End With
                Case Else
                    isSelected = True
                End Select
                If Not isSelected Then frmSelect.Show vbModal
            End If
        End With
    Case "Part_D":
        With ActiveCell
            If .Row >= 19 And .Row <= 22 Then
                If .Column >= 4 And .Column <= 9 Then
                    With frmObjectParameter
                        .SelectedItem = ActiveCell
                        .ReadOnly = True
                        .WrapOutput = False
                        .ReturnIndexOnly = True
                    End With
                    With frmObjectParameter
                        .SpecialNote = MSG("MSG_SEL_HELP_TYPE") & tmpStr
                        .DataSetName = MSG("MSG_SEL_HELP_TYPE")
                        .DataSource = "LST_HELP_TYPE"
                    End With
                    If Not isSelected Then frmSelect.Show vbModal
                End If
            End If
        End With
    End Select
    'reset form argument value
    Dim lRet As FormArgument
    frmObjectParameter = lRet
End Sub

Sub WriteLog(MsgToWrite As String, Optional LogFileName As String = "ImportResult.txt")
    Dim txtString As String, FileNames As String
    FileNames = ThisWorkbook.Path & "\" & LogFileName
    Open FileNames For Append As #1
    Print #1, Format(Now(), "DD/MM/YYYY HH:MM:SS") & vbTab & "[" & MsgToWrite & "]"
    Close #1
    Exit Sub
ErrorChk:
    Close #1
End Sub
