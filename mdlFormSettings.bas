Attribute VB_Name = "mdlFormSettings"
Option Explicit

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

Function FormIsLoaded(UFName As String) As Boolean
  Dim UF As Integer
  For UF = 0 To VBA.UserForms.Count - 1
    FormIsLoaded = UserForms(UF).Name = UFName
    If FormIsLoaded Then Exit Function
  Next UF
End Function

Function HOA(chuoi As String) As String
  chuoi = Application.WorksheetFunction.Trim(chuoi)
  HOA = UCase(chuoi)
End Function

Sub RegisterAction()
    IndirectSetup = True
End Sub

Sub DeRegisterAction()
    IndirectSetup = False
End Sub
