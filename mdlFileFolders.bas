Option Explicit

' This module is for folder and file stuff
#If VBA7 Then
    Private Declare PtrSafe Function SHBrowseForFolderW Lib "shell32.dll" (lpBrowseInfo As BROWSEINFO) As Long
    Private Declare PtrSafe Function SHGetPathFromIDListW Lib "shell32.dll" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
    Private Declare PtrSafe Function SendMessageA Lib "user32" (ByVal hwnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, lParam As Any) As LongPtr
    Private Declare PtrSafe Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
    Private Declare PtrSafe Function SetWindowText Lib "user32" Alias "SetWindowTextW" (ByVal hwnd As Long, ByVal lpString As String) As Long
    
    Public Type BROWSEINFO
        hWndOwner As LongPtr
        pidlRoot As Long
        pszDisplayName As String
        lpszTitle As String
        ulFlags As Long
        lpfnCallback As LongPtr
        lParam As LongPtr
        iImage As Long
    End Type
#Else
    Private Declare Function SHBrowseForFolderW Lib "shell32.dll" (lpBrowseInfo As BROWSEINFO) As Long
    Private Declare Function SHGetPathFromIDListW Lib "shell32.dll" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
    Private Declare Function SendMessageA Lib "user32" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
    Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
    Private Declare Function SetWindowText Lib "user32" Alias "SetWindowTextW" (ByVal hwnd As Long, ByVal lpString As String) As Long
    
    Public Type BROWSEINFO
        hWndOwner As Long
        pidlRoot As Long
        pszDisplayName As String
        lpszTitle As String
        ulFlags As Long
        lpfnCallback As Long
        lParam As Long
        iImage As Long
    End Type
#End If

Private Const BIF_RETURNONLYFSDIRS  As Long = 1
Private Const CSIDL_DRIVES          As Long = &H11
Private Const WM_USER               As Long = &H400
Private Const MAX_PATH              As Long = 260 ' Is it a bad thing that I memorized this value?
Private Const MAX_PATH_UNICODE      As Long = 260 * 2 - 1

'// message from browser
Private Const BFFM_INITIALIZED     As Long = 1
Private Const BFFM_SELCHANGED      As Long = 2
Private Const BFFM_VALIDATEFAILEDA As Long = 3 '// lParam:szPath ret:1(cont),0(EndDialog)
Private Const BFFM_VALIDATEFAILEDW As Long = 4 '// lParam:wzPath ret:1(cont),0(EndDialog)
Private Const BFFM_IUNKNOWN        As Long = 5 '// provides IUnknown to client. lParam: IUnknown*

'// messages to browser
Private Const BFFM_SETSTATUSTEXTA   As Long = WM_USER + 100
Private Const BFFM_ENABLEOK         As Long = WM_USER + 101
Private Const BFFM_SETSELECTIONA    As Long = WM_USER + 102
Private Const BFFM_SETSELECTIONW    As Long = WM_USER + 103
Private Const BFFM_SETSTATUSTEXTW   As Long = WM_USER + 104
Private Const BFFM_SETOKTEXT        As Long = WM_USER + 105 '// Unicode only
Private Const BFFM_SETEXPANDED      As Long = WM_USER + 106 '// Unicode only
        
#If VBA7 Then
Private Function PtrToFunction(ByVal lFcnPtr As LongPtr) As LongPtr
    PtrToFunction = lFcnPtr
End Function
#Else
Private Function PtrToFunction(ByVal lFcnPtr As Long) As Long
    PtrToFunction = lFcnPtr
End Function
#End If

Public Function BFFCallback(ByVal hwnd As Long, ByVal uMsg As Long, ByVal lParam As Long, ByVal sData As String) As Long
    If uMsg = BFFM_INITIALIZED Then
        SendMessageA hwnd, BFFM_SETSELECTIONA, True, ByVal sData
        ' Set caption for the dialog
        Call SetWindowText(hwnd, StrConv(MSG("DLG_FOLDER_BROWSE"), vbUnicode))
    End If
End Function

' ============FILE FOLDER OPERATIONS========================
Function FileOrDirExists(PathName As String, Optional FileObject As Boolean = False) As Boolean
'No need to set a reference if you use Late binding
    Dim fso As Object
    Dim FilePath As String, lRet As Boolean

    Set fso = CreateObject("Scripting.FileSystemObject")
    If PathName = "" Then Exit Function
    If FileObject Then
        FileOrDirExists = fso.FileExists(PathName)
    Else
        FileOrDirExists = fso.FolderExists(PathName)
    End If
    Set fso = Nothing
End Function

Function GetFolderFromFilePath(PathString As String) As String
    ' Return folder for a file
    On Error Resume Next
    GetFolderFromFilePath = StrReverse(Mid(StrReverse(PathString), InStr(StrReverse(PathString), "\")))
End Function

Function GetPathFileName(PathString As String) As String
    ' Return a folder with file name without extension
    On Error Resume Next
    GetPathFileName = StrReverse(Mid(StrReverse(PathString), InStr(StrReverse(PathString), ".") + 1))
End Function

Function GetFileNameFromFilePath(FileName As String, Optional SearchString As String = "\") As String
    ' Return filename with extension from file path
    On Error Resume Next
    GetFileNameFromFilePath = Mid(FileName, InStrRev(FileName, SearchString) + 1)
End Function

Function GetFileExtension(FileName As String) As String
    ' Return filename with extension from file path
    On Error Resume Next
    GetFileExtension = Mid(FileName, InStrRev(FileName, ".") + 1)
End Function

Function GetFolder(strPath As String, _
    Optional FilePicker As Boolean = False, _
    Optional FileExtension As String = "*.xls*", _
    Optional DlgTitle As String = "") As String
    
    Dim fldr As FileDialog
    Dim sItem As String
    If FilePicker Then
        Set fldr = Application.FileDialog(msoFileDialogFilePicker)
    Else
        Set fldr = Application.FileDialog(msoFileDialogFolderPicker)
    End If
StepResumeFolder:
    With fldr
        .Title = DlgTitle
        .AllowMultiSelect = False
        .InitialFileName = strPath
        If FilePicker Then
            .Filters.Clear
            .Filters.Add MSG("MSG_EXCEL_FILE_TYPE"), "*." & FileExtension
        End If
        If .Show <> -1 Then
            'user select cancel
            sItem = ""
        Else
            sItem = .SelectedItems(1)
        End If
    End With
    
    ' Test to make sure that user selected anything
    Dim FileLocation As String, FldBrowser As String
    
    If Not FileOrDirExists(sItem, FilePicker) Then
        If MsgBox(MSG("MSG_SELECT_NO_FILE"), vbInformation + vbOKCancel) = vbOK Then GoTo StepResumeFolder
        ' safe exit
        sItem = ""
        GoTo NextCode
    End If
    If FilePicker Then
        ' User select current file or not?
        If sItem = ThisWorkbook.Path & "\" & ThisWorkbook.Name Then
            MsgBox MSG("MSG_ERROR_THIS_FILE"), vbInformation
            GoTo StepResumeFolder
        End If
    End If
NextCode:
    GetFolder = sItem
    Set fldr = Nothing
End Function

Function GetBrowseObject(Optional InitialDir As String = "", Optional MultipleObject As Boolean = True, Optional FileExt As String = "xls") As String
    
    ' This will help selecting a file for later processing
    ' Can return a list of file or a single file
    
    '1. We have to Open source data location
    
    ' Just assign inital dir with template folder
    If InitialDir = "" Then InitialDir = AppSetting("CONF_LAST_DOCUMENT_PATH")
    InitialDir = GetFolderFromFilePath(InitialDir)
    If Not FileOrDirExists(InitialDir) Then InitialDir = ThisWorkbook.Path
          
    ' User may want to do differently with importing
    If Not MultipleObject Then
        ' Browse to the file
        InitialDir = GetFolder(InitialDir, True, FileExt, MSG("MSG_SELECTDATAFOLDER"))
        
        ' check again then
        If InitialDir = "" Or InitialDir = "\" Then
            ' Exit code
            MsgBox MSG("MSG_SELECT_NO_FILE"), vbInformation
            GoTo CancelEvent
        End If
        If Not FileOrDirExists(InitialDir, True) Then
            ' Exit code
            MsgBox MSG("MSG_FILE_FOLDER_PROBLEM"), vbInformation
            GoTo CancelEvent
        End If
        GoTo SEL_SINGLE_OBJECT
    End If
        
    ' do a selection of folder
    InitialDir = GetFolder(InitialDir, , , MSG("MSG_SELECTDATAFOLDER"))
    
    ' check again then
    If InitialDir = "" Or InitialDir = "\" Then GoTo CancelEvent
    
    ' Save the location for later use
    AppSetting("CONF_LAST_DOCUMENT_PATH") = InitialDir
    
    ' Ok now - we go forward to search for file
    Application.StatusBar = Replace(MSG("MSG_DOSEARCH"), "%s%", "[" & InitialDir & "]")
    DoEvents
    ' Search for files and other...
    InitialDir = ListFiles(InitialDir, , FileExt)
    
SEL_SINGLE_OBJECT:
    '2. Start the export process
    ' turn file list to an array for easier looping
    GetBrowseObject = InitialDir
    
CancelEvent:
End Function

Function ListFiles(initFolderName As String, Optional initSp As String = "|", Optional initFlExt As String = "xls") As String
    Dim fs As Object, fileStr As String
     
    'Creating File System Object
    Set fs = CreateObject("Scripting.FileSystemObject")
     
    'Call the GetFile function to get all files
    fileStr = GetFiles(fs, initFolderName, initSp, initFlExt)
    Set fs = Nothing
    If Len(fileStr) > 1 Then
        fileStr = Mid(fileStr, Len(initSp) + 1)
        ' remember to replace such double stuff
        ListFiles = Replace(fileStr, initSp & initSp, initSp)
    End If
End Function

Private Function GetFiles(fso As Object, FolderName As String, Optional sp As String = "|", Optional flExt As String = "xls") As String
    On Error Resume Next
    Dim objFolder As Object
    Dim ObjSubFolders As Object
    Dim objSubFolder As Object
    Dim ObjFiles As Object
    Dim objFile As Object
    Dim OutString As String
    
    Set objFolder = fso.GetFolder(FolderName)
    Set ObjFiles = objFolder.Files
     
    'Write all files to output files
    For Each objFile In ObjFiles
        If objFile.Name <> "" Then
            If LCase(GetFileExtension(objFile.Name)) Like LCase(flExt) & "*" Then
                OutString = OutString & sp & objFile.Path
            End If
        End If
    Next
    'Getting all subfolders
    Set ObjSubFolders = objFolder.SubFolders
     
    For Each objFolder In ObjSubFolders
        'Getting all Files from subfolder
        OutString = OutString & sp & GetFiles(fso, objFolder.Path, sp, flExt)
    Next
    GetFiles = Replace(OutString, sp & sp, sp)
End Function
