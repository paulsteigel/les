Option Explicit

Private Const wdLineSpaceSingle = 0
Private Const wdAlignParagraphJustify = 3
Private Const wdAlignParagraphCenter = 1
Private Const wdAlignPageNumberCenter = 1
Private Const wdAlignParagraphRight = 2
Private Const wdAlignParagraphLeft = 0
Private Const wdOutlineLevel1 = 1
Private Const wdTrailingSpace = 1
Private Const wdListNumberStyleUppercaseRoman = 1
Private Const wdUndefined = &H98967F
Private Const wdPasteRTF = 1

Private Const wdListNumberStyleArabic = 0
Private Const wdListNumberStyleLowercaseLetter = 4
Private Const wdListNumberStyleNumberInCircle = &H12
Private Const wdListLevelAlignLeft = 0
Private Const wdTrailingTab = 0
Private Const wdOutlineNumberGallery = 3
Private Const wdLineSpaceMultiple = 5
Private Const wdPreferredWidthPercent = 2
Private Const wdPreferredWidthPoints = 3
Private Const wdRowHeightAtLeast = 1
Private Const wdReplaceAll = 2
Private Const wdFindContinue = 1
Private Const wdFindStop = 0

Private Const wdOutlineLevelBodyText = 10
Private Const wdListNumberStyleBullet = &H17
Private Const wdStyleListNumber = &HFFFFFFCE
Private Const wdStyleListNumber2 = &HFFFFFFC5
Private Const wdStyleListNumber3 = &HFFFFFFC4
Private Const wdStyleListNumber4 = &HFFFFFFC3
Private Const wdStyleListNumber5 = &HFFFFFFC2
Private Const wdStyleNormal = &HFFFFFFFF
Private Const wdBulletGallery = 1
Private Const wdAlignTabCenter = 1
Private Const wdAlignTabLeft = 0
Private Const wdTabLeaderSpaces = 0
Private Const wdStyleTypeParagraph = 1
Private Const wdAlignTabRight = 2

Private Const wdSectionBreakNextPage = 2
Private Const wdOrientPortrait = 0
Private Const wdOrientLandscape = 1
Private Const wdPasteMetafilePicture = 3
Private Const wdPasteEnhancedMetafile = 9
Private Const wdInLine = 0
Private Const wdLineSpace1pt5 = 1

Private Type FontFormat
    FontBold As Boolean
    FontItalic As Boolean
    FontUnderlined As Boolean
    FontAllCap As Boolean
    FontSize As Long
End Type

Private Sub BulletText(sDoc As Object, LinkObj As String, Optional BulletStyle As String = "", Optional NumberStyle As Long)
    Dim myList As Object

    ' Add a new ListTemplate object
    Set myList = sDoc.ListTemplates.Add
    
    With myList.ListLevels(1)
        .Alignment = wdListLevelAlignLeft
        .ResetOnHigher = 0
        .StartAt = 1
        
        ' The following sets the font attributes of
        ' the "bullet" text
        .LinkedStyle = LinkObj
        If BulletStyle <> "" Then
            .TrailingCharacter = wdTrailingSpace
            .NumberPosition = 0
            .TextPosition = 0
            .NumberStyle = NumberStyle
            .NumberFormat = BulletStyle
        Else
            .TrailingCharacter = wdTrailingTab
            .NumberFormat = ChrW(183)
            .NumberPosition = Excel.Application.CentimetersToPoints(1)
            .TextPosition = Excel.Application.CentimetersToPoints(1.6)
            .TabPosition = Excel.Application.CentimetersToPoints(1.6)
            With .Font
                .Bold = False
                .Name = "Symbol" '"Wingdings"
                .Size = 13
            End With
        End If
    End With
End Sub

Private Function StyleExist(DocObj As Object, StlName As String) As Boolean
    Dim MyStl As Object, StlObjName As String
    On Error GoTo ErrHandler
    Set MyStl = DocObj.Styles(StlName)
    StlObjName = MyStl.NameLocal
    StyleExist = True
ErrHandler:
End Function

Sub InsertSection(WrdDoc As Object, Optional ToLastPage As Boolean = True)
    If ToLastPage Then
        WrdDoc.Paragraphs.Last.Range.InsertBreak Type:=wdSectionBreakNextPage
    Else
        ' Just to current place
    End If
End Sub

Sub SetSectionLayout(myWordDoc As Object, Optional SetLandscape As Boolean = True)
    myWordDoc.Sections.Last.PageSetup.Orientation = IIf(SetLandscape, wdOrientPortrait, wdOrientLandscape)
End Sub

Sub AddTable(WrdDoc As Object, tblRange As Range)
    Dim Tbl As Object
    Dim iCol As Long, iRow As Long, i As Long, j As Long
    iRow = tblRange.Rows.Count
    iCol = tblRange.Columns.Count
    
    Set Tbl = WrdDoc.Tables.Add(WrdDoc.Paragraphs.Last.Range, iRow, iCol)
    ' Now set table with and column width
    With Tbl
        .PreferredWidthType = wdPreferredWidthPercent
        .PreferredWidth = 100
        .Rows.HeightRule = wdRowHeightAtLeast
        
        .Rows.Height = Excel.Application.CentimetersToPoints(0)
        '.Rows.LeftIndent = Excel.Application.CentimetersToPoints(0)
            
        ' Column size
        For i = 1 To iCol
            'On Error Resume Next
            .Columns(i).PreferredWidthType = wdPreferredWidthPercent
            .Columns(i).PreferredWidth = 100 * tblRange.Columns(i).ColumnWidth / tblRange.Width
        Next
        Err.Clear
        For i = 1 To tblRange.Rows.Count
            For j = 1 To tblRange.Columns.Count
                .Cell(i, j) = Trim(tblRange.Cells(i, j))
                ' alignment
                Select Case tblRange.Cells(i, j).HorizontalAlignment
                Case xlLeft:
                    .Cell(i, j).Range.ParagraphFormat.Alignment = wdAlignParagraphLeft
                Case xlRight
                    .Cell(i, j).Range.ParagraphFormat.Alignment = wdAlignParagraphRight
                Case Else
                    .Cell(i, j).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
                End Select
            Next
        Next
    End With
End Sub

Sub RemoveTagAndFormat(DocObj As Object)
    Dim TagArr As Variant, i As Long, TagFormat As FontFormat, DefaultFontSize As Long
    TagArr = Array("bold", "allcap")
    DefaultFontSize = DocObj.Styles("Normal").Font.Size
    With TagFormat
        For i = 0 To UBound(TagArr)
            ' Initial attributes
            .FontAllCap = False
            .FontBold = False
            .FontItalic = False
            .FontUnderlined = False
            .FontSize = DefaultFontSize
            ' other attributes
            Select Case TagArr(i)
            Case "bold": .FontBold = True
            Case "allcap": .FontAllCap = True
            End Select
            ' do the format
            FormatTags DocObj, CStr(TagArr(i)), TagFormat
        Next
    End With
End Sub

Private Sub FormatTags(DocObj As Object, TagStr As String, ObjFormat As FontFormat)
'
' Macro1 Macro
'
'
    Dim DocRange As Object
    Set DocRange = DocObj.Range
    ' First Setting things to be bold
    With DocRange.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        With .Replacement.Font
            .Bold = ObjFormat.FontBold
            .Italic = ObjFormat.FontItalic
            .AllCaps = ObjFormat.FontAllCap
            .Underline = ObjFormat.FontUnderlined
            .Size = ObjFormat.FontSize
        End With
        .Text = "\<" & TagStr & "\>*\</" & TagStr & "\>"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    
        .Execute Replace:=wdReplaceAll
    End With
    ' Now removing stuff
    With DocRange.Find
        .Text = "<" & TagStr & ">"
        .Format = False
        .MatchWildcards = False
        .Execute Replace:=wdReplaceAll
        .Replacement.Text = ""
        .Execute Replace:=wdReplaceAll
        
        .Text = "</" & TagStr & ">"
        .Replacement.Text = ""
        .Execute Replace:=wdReplaceAll
    End With
    Set DocRange = Nothing
End Sub

Sub ReformatWordTable(WrdDoc As Object, Optional Msg1 As String, Optional Msg2 As String, Optional MsgFin As String)
    Dim tmpObj As Object, prg As Object, i As Long, ErrCount As Long
    Dim DefaultFont As String
    DefaultFont = WrdDoc.Styles("Normal").Font.Name
    For Each tmpObj In WrdDoc.Tables
        ShowStatus Msg1 & " " & tmpObj.Id
        'Format the selected table
        With tmpObj.Range.ParagraphFormat
            .SpaceBefore = 0
            .SpaceAfter = 0
            .LineSpacingRule = wdLineSpaceSingle
            .FirstLineIndent = Excel.Application.CentimetersToPoints(0)
        End With
        With tmpObj
            .PreferredWidthType = wdPreferredWidthPercent
            .PreferredWidth = 100
            .Rows.HeightRule = wdRowHeightAtLeast
            .Rows.Height = Excel.Application.CentimetersToPoints(0)
            .Rows.LeftIndent = Excel.Application.CentimetersToPoints(0)
                
            .TopPadding = Excel.Application.CentimetersToPoints(0)
            .BottomPadding = Excel.Application.CentimetersToPoints(0)
            .LeftPadding = Excel.Application.CentimetersToPoints(0.19)
            .RightPadding = Excel.Application.CentimetersToPoints(0.19)
            .Spacing = 0
            .AllowPageBreaks = True
            .AllowAutoFit = True
    
            'set font
            .Range.Font.Name = DefaultFont
        End With
        
        ' Set header row
        SetHeaderRow tmpObj
        
        ' Remove trailing space
        ShowStatus Msg2 & tmpObj.Id
        
        On Error GoTo ErrPattern
        Dim Pattern As String
        Pattern = ","
        
        With tmpObj.Range.Find
            .ClearFormatting
            .Replacement.ClearFormatting
            .Text = "([ ])[ ]{1" & Pattern & "}"
            .Replacement.Text = "\1"
            .MatchWildcards = True
            .Forward = True
            .Wrap = wdFindContinue
            .Execute Replace:=wdReplaceAll
        End With
    Next
    ShowStatus MsgFin
    Set tmpObj = Nothing
    Exit Sub
    
ErrPattern:
    Err.Clear
    
    ErrCount = ErrCount + 1
    Select Case ErrCount
    Case 1:
        Pattern = ";"
    Case 2:
        Pattern = "."
    Case Else
        Exit Sub
    End Select
    Resume 0
End Sub

Sub SetHeaderRow(myTable As Object)
    Dim HeaderRange As Object
    On Error GoTo ErrHandler
    
    Set HeaderRange = myTable.Rows(1).Range
    HeaderRange.Rows.HeadingFormat = True
    Set HeaderRange = Nothing
    Exit Sub
ErrHandler:
    If Err.Number <> 0 Then Err.Clear
    Set HeaderRange = myTable.Cell(1, 1).Range
    HeaderRange.SetRange Start:=myTable.Cell(1, 1).Range.Start, End:=myTable.Cell(1, 1).Range.Start
    Resume Next
End Sub

Private Sub FormatTable(wrDoc As Object, Tbl As Object, Col2Format As Long, StartRow As Long)
'
' FormatTable Macro, will do the setting up of table and then get it updated quickly..
' The key is number of column to be formatted and starting row...
'
    Dim i As Long, ColNums As Long, MyCells As Object
    ' With table format
    With Tbl
        .Rows.LeftIndent = 0
        .PreferredWidthType = wdPreferredWidthPercent
        .PreferredWidth = 100
        .Rows.HeightRule = wdRowHeightAtLeast
        .Rows.Height = 0
        With .Range.ParagraphFormat
            .SpaceBefore = 0
            .SpaceBeforeAuto = False
            .SpaceAfter = 0
            .SpaceAfterAuto = False
            .LineSpacingRule = wdLineSpaceSingle
            .FirstLineIndent = 0
            .CharacterUnitFirstLineIndent = 0
            .LineUnitBefore = 0
            .LineUnitAfter = 0
        End With
        
        ' Remove trailing space
        ColNums = .Columns.Count
        Set MyCells = wrDoc.Range(.Cell(StartRow + 1, ColNums - Col2Format + 1).Range.Start, .Cell(.Rows.Count, ColNums).Range.End)
    End With
    ' Remove trailing space
    With MyCells.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = " "
        .Replacement.Text = ""
        .Forward = False
        .Wrap = wdFindStop 'wdFindContinue
        .Execute Replace:=wdReplaceAll
    End With
    Set MyCells = Nothing
End Sub


