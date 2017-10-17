VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSelect 
   Caption         =   "UserForm1"
   ClientHeight    =   5265
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5235
   OleObjectBlob   =   "frmSelect.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private SaveFailed As Boolean
Private AddedString As String
Private DeletedItem As String
Private Loading As Boolean

Dim RefData_index As String
Dim RefData_text As String

Private Sub cbListCategory_Change()
    'If Loading Then Exit Sub
    '' fill the listbox with new data source
    'With frmObjectParameter
    '    .DataSource = cbListCategory.List(cbListCategory.ListIndex, 0)
    '    'AddName .DataSource
    '    ' reset the string
    '    GetObjSource lstObject, , 1, .DataSource & "_LST", ActiveCell, AddedString
    'End With
End Sub

Private Sub AddName(NameTxt As String)
    On Error Resume Next
    ThisWorkbook.Names.Add Name:=NameTxt & "_LST", RefersToR1C1:="=OFFSET(" & NameTxt & ",0,0,COUNTA(" & NameTxt & "),1)"
End Sub

Private Sub cmdAdd_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    lstObject_KeyDown KeyCode, Shift
End Sub

Private Sub GetListData()
    Dim i As Long
    ' This will put data in good order
    RefData_index = ""
    RefData_text = ""
    With lstObject
        For i = 0 To .ListCount - 1
            If lstObject.ColumnCount > 1 Then
                RefData_index = RefData_index & vbLf & .List(i, 0)
                RefData_text = RefData_text & "," & vbLf & .List(i, 1)
            ElseIf .Selected(i) Then
                RefData_index = RefData_index & ", " & Val(.List(i))
            End If
        Next
    End With
    If RefData_text <> "" Then
        RefData_text = Mid(RefData_text, 2)
        RefData_index = Mid(RefData_index, 2)
    Else
        RefData_index = Mid(RefData_index, 3)
    End If
End Sub

Private Sub cmdOK_Click()
    Dim ctlObj As Control, ErrMsg As String, retCol As Variant
    ' This form shall have to return selected value to current selected row...
    With frmObjectParameter
        If lstObject.ListCount = 0 Then Exit Sub
        GetListData
        
        If Not .ReadOnly Then
            ' This is to return stuff in special format
            If .ReturnDataOrder <> "" Then
                retCol = Split(.ReturnDataOrder, ",")
                If ActiveCell.Column = retCol(0) Then
                    ActiveCell.Value = RefData_text
                    ActiveCell.Offset(0, 1).Value = RefData_index
                Else
                    ActiveCell.Value = RefData_index
                    ActiveCell.Offset(0, -1).Value = RefData_text
                End If
                Unload Me
            End If
        ElseIf .AllowMultipleSelection Then
            ActiveCell.Value = RefData_index
            Unload Me
        End If
        
        If lstObject <> "" Then
            If .DontAssignActiveCell Then
                .SelectedItem = lstObject
            Else
                Dim xOldCell As String
                xOldCell = ActiveCell
                If .ReturnIndexOnly Then
                    ActiveCell.Value = Val(lstObject)
                Else
                    If ActiveCell <> lstObject Then
                        ActiveCell.Value = IIf(.WrapOutput, "[" & lstObject & "]", lstObject)
                    End If
                End If
            End If
        End If
        Unload Me
    End With
End Sub

Private Sub cmdAdd_Click()
    ' Add village name to listbox
    Dim AddStr As String
    If cbListCategory.Visible Then
        ' check for the selection of category, only check this if the form is in Activity Group mode
        If cbListCategory.Text = "" Then
            MsgBox MSG("MSG_YOU_HAVE_NOT_SELECT_A_GROUP"), vbInformation
            cbListCategory.SetFocus
            cbListCategory.DropDown
            GoTo ExitSub
        End If
    End If
    
    AddStr = txtObject
    With lstObject
        .AddItem
        .List(.ListCount - 1, 0) = Val(cbListCategory)
        .List(.ListCount - 1, 1) = AddStr
        AddedString = AddedString & "[" & AddStr & "]"
        cmdOK.Enabled = IIf(.ListCount > 0, True, False)
    End With
    
    cmdAdd.Enabled = False
    txtObject = ""
ExitSub:
End Sub

Private Sub cmdOK_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    lstObject_KeyDown KeyCode, Shift
End Sub

Private Sub cmdRemove_Click()
    ' Try to remove current selected Item
    Dim StrDelete As String, DeletedRange As Range
    If lstObject <> "" Then
        StrDelete = lstObject
        
        ' Remove from list
        AddedString = Replace(AddedString, "[" & StrDelete & "]", "")
        lstObject.RemoveItem lstObject.ListIndex
        DeletedItem = DeletedItem & "[" & StrDelete & "]"
        
    End If
    If lstObject.ListCount <= 0 Then
        cmdRemove.Enabled = False
        cmdOK.Enabled = False
    End If
End Sub

Private Sub cmdRemove_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    lstObject_KeyDown KeyCode, Shift
End Sub

Private Sub lstObject_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    If frmObjectParameter.NotAllowSelection <> "" Then
        ' Not allow selection in case of component criteria
        If InStr(lstObject, frmObjectParameter.NotAllowSelection) <> 0 Then
            lstObject.Selected(lstObject.ListIndex) = False
        End If
    End If
    If frmObjectParameter.ReadOnly Then Exit Sub
    cmdRemove.Enabled = IIf(lstObject <> "", True, False)
    If txtObject <> lstObject Then txtObject = lstObject.List(lstObject.ListIndex, 2)
    'Loading = Not Loading
End Sub

Private Sub lstObject_Change()
    If Loading Then Exit Sub
    If Not frmObjectParameter.ReadOnly Then
        cmdOK.Enabled = IIf(lstObject.ListCount > 0, True, False)
    ElseIf frmObjectParameter.AllowMultipleSelection Then
        cmdOK.Enabled = IIf(lstObject.Selected(lstObject.ListIndex), True, False)
    Else
        cmdOK.Enabled = IIf(lstObject <> "", True, False)
    End If
End Sub

Private Sub lstObject_Click()
    Debug.Print lstObject
End Sub

Private Sub lstObject_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    'If user press Enter - handle it
    Select Case KeyCode
    Case vbKeyReturn
        cmdOK_Click
    Case vbKeyEscape
        ' escape form
        Unload Me
    End Select
End Sub

Private Sub txtObject_Change()
    If frmObjectParameter.ReadOnly Then Exit Sub
    If Len(Trim(txtObject)) <= 1 Or InStr(AddedString, "[" & txtObject & "]") <> 0 Then
        cmdAdd.Enabled = False
    Else
        cmdAdd.Enabled = True
    End If
End Sub

Private Sub txtObject_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    lstObject_KeyDown KeyCode, Shift
End Sub

Private Sub UserForm_Initialize()
' Feb 28 2014, added with new features on grouping of activities
    If ExternalLoad Then Exit Sub
    Loading = True
    With frmObjectParameter
        'set form caption
        SetUnicodeCaption Me, Replace(SetConfig("FORM_frmSelect", Me), "RELITEM", .DataSetName)
        ' As this would be the very first time - village will have nothing
        cmdAdd.Enabled = False
        cmdRemove.Enabled = False
        
        ' Now parse data to the list
        If Not .ReadOnly Then
            ' Just allow editing so we make this a new event
            GetObjSource cbListCategory, , 1, .DataSource, , , True
            lstObject.ColumnCount = 2
            lstObject.ColumnWidths = "20;100"
            cbListCategory.Visible = True
            ' Should we load data into the list?
            'lstObject.RowSource = .RowSource
        Else
            ' Set position of objects
            lstObject.Height = lstObject.Height + (lstObject.Top - txtObject.Top)
            lstObject.Top = txtObject.Top
            txtObject.Top = cbListCategory.Top
            
            GetObjSource lstObject, , 1, .DataSource, ActiveCell, AddedString, .ReturnIndexOnly
            If .AllowMultipleSelection Then
                lstObject.ListStyle = fmListStyleOption
                lstObject.MultiSelect = fmMultiSelectMulti
            End If
        End If
        
        ' Modified the label caption
        If .ReadOnly Then
            lblObject.Caption = .SpecialNote
        Else
            lblObject.Caption = Replace(lblObject.Caption, "RELITEM", .DataSetName)
        End If
    End With
    Loading = False
End Sub

Private Sub UserForm_Terminate()
    ' If the form was closed without doing anything, ok - quit
End Sub
