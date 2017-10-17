Attribute VB_Name = "mdl_WB"
Option Explicit

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

Sub TryExt()
    'Fill in Commbo or listbox with region table
    Dim theCell As Range, i As Long
    ShowOff
    For i = 1 To Range("tblRegions").Rows.Count
        Range("tblRegions").Cells(i, 4) = Trim(Range("tblRegions").Cells(i, 4))
    Next
    ShowOff True
End Sub



