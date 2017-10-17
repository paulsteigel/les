Option Explicit
' This will manipulate a word document
' Declaration for Word style creation
Private Const wdLineSpaceSingle = 0
Private Const wdAlignParagraphJustify = 3
Private Const wdAlignParagraphCenter = 1
Private Const wdAlignPageNumberCenter = 1
Private Const wdOutlineLevel1 = 1
Private Const wdTrailingSpace = 1
Private Const wdListNumberStyleUppercaseRoman = 1
Private Const wdUndefined = &H98967F
Private Const wdFindContinue = 1
Private Const wdReplaceAll = 2
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
Private Const wdTabLeaderSpaces = 0
Private Const wdStyleTypeParagraph = 1
Private Const wdAlignParagraphRight = 2
Private Const wdAlignParagraphLeft = 0
Private Const wdPasteRTF = 1
Private Const wdFindStop = 0
Private Const wdAlignTabLeft = 0
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

' Counter vaiables for field access
Private proc_UseListOnly As Boolean
Private proc_ColumData As Long
Private NonCalculatedField As Boolean
Const AppPassWord = "d1ndh1sk" ' global password for protection

' Cached variable for keeping some temporary stuff
Private CachedListDistinct As Collection
Private ColListing() As New Collection
Private CurrentPointer As Long
Private OldTableName As String

' For Document
Private WordApplication As Object
Private WordDocument As Object
Private CloseOnExit As Boolean
Private DefaultFontName As String

Private Sub Class_Initialize()
    ' By default set close word on exit
    CloseOnExit = True
    ' Now create a word document
    CreateWordDocument
    ' Set default font name
    DefaultFontName = "Times New Roman"
End Sub

Private Sub Class_Terminate()
    ' Closing Word app now
    If CloseOnExit Then WordApplication.Quit
    Set WordApplication = Nothing
End Sub

Property Let WordAppVisible(VisibleState As Boolean)
    WordApplication.Visible = VisibleState
End Property

Property Get ActiveDocument() As Object
    Set ActiveDocument = WordDocument
End Property

Property Get ActiveWordApplication() As Object
    Set ActiveWordApplication = WordApplication
End Property

Property Let SetDefaultFontName(FontName As String)
    DefaultFontName = FontName
End Property

Property Let CloseWordOnExit(CloseState As Boolean)
    CloseOnExit = CloseState
End Property

Property Let EnableWordEvent(EnableEvent As Boolean)
    ' For disable Word stuff
End Property

Sub Activate()
    WordApplication.Activate
End Sub

Function GenerateWordStyle() As Boolean
    ' Sets up built-in numbered list styles and List Template
    ' including restart paragraph style
    ' Run in document template during design
    ' Macro created by Margaret Aldis, Syntagma
    '
    ' Create list starting style and format if it doesn't already exist
    
    Dim strStyleName As String, tmpName  As String
    strStyleName = "Heading 1" ' the style name in this set up
    Dim strListTemplateName As String
    strListTemplateName = "SEDP_List_Template" ' the list template name in this set up
    Dim astyle As Object
        For Each astyle In WordDocument.Styles
            If astyle.NameLocal = strStyleName Then GoTo Define 'already exists
        Next astyle
    ' doesn't exist
    WordDocument.Styles.Add Name:=strStyleName, Type:=wdStyleTypeParagraph
Define:
    With WordDocument.Styles(strStyleName)
        .AutomaticallyUpdate = False
        .BaseStyle = ""
        .NextParagraphStyle = wdStyleListNumber 'for international version compatibility
    End With
    With WordDocument.Styles(strStyleName).ParagraphFormat
        .LineSpacingRule = wdLineSpaceSingle
        .WidowControl = False
        .KeepWithNext = True
        .KeepTogether = True
        .OutlineLevel = wdOutlineLevelBodyText
    End With
    ' Create the list template if it doesn't exist
    Dim aListTemplate As Object
        For Each aListTemplate In WordDocument.ListTemplates
            If aListTemplate.Name = strListTemplateName Then GoTo Format 'already exists
        Next aListTemplate
    ' doesn't exist
        Dim newlisttemplate As Object
        Set newlisttemplate = WordDocument.ListTemplates.Add(OutlineNumbered:=True, Name:="SEDP_List_Template")
Format:
' Set up starter and three list levels - edit/extend from recorded details if required
    'Level 1
    With WordDocument.ListTemplates(strListTemplateName).ListLevels(1)
        .NumberFormat = "Ph" & ChrW(7847) & "n %1:"
        .TrailingCharacter = wdTrailingSpace
        .NumberStyle = wdListNumberStyleUppercaseRoman
        .NumberPosition = Excel.Application.CentimetersToPoints(0)
        .Alignment = wdListLevelAlignLeft
        .TextPosition = Excel.Application.CentimetersToPoints(0.76)
        .TabPosition = Excel.Application.CentimetersToPoints(0)
        .ResetOnHigher = True
        .StartAt = 1
        .LinkedStyle = strStyleName
    End With
    With WordDocument.Styles(strStyleName)
        With .ParagraphFormat
            .LeftIndent = Excel.Application.CentimetersToPoints(0.76)
            .rightIndent = Excel.Application.CentimetersToPoints(0)
            .SpaceBefore = 12
            .SpaceAfter = 3
            .LineSpacingRule = wdLineSpaceSingle
            .Alignment = wdAlignParagraphJustify
            .KeepWithNext = True
            .PageBreakBefore = False
            .FirstLineIndent = Excel.Application.CentimetersToPoints(-0.76)
            .OutlineLevel = wdOutlineLevel1
        End With
        With .Font
            .Name = DefaultFontName
            .Size = 16
            .Bold = True
        End With
    End With

    ' Level 2
    With WordDocument.ListTemplates(strListTemplateName).ListLevels(2)
        .NumberFormat = "%2."
        .TrailingCharacter = wdTrailingSpace
        .NumberStyle = wdListNumberStyleArabic
        .NumberPosition = Excel.Application.CentimetersToPoints(0)
        .Alignment = wdListLevelAlignLeft
        .TextPosition = Excel.Application.CentimetersToPoints(0.5)
        .TabPosition = Excel.Application.CentimetersToPoints(0.5)
        .ResetOnHigher = True
        .StartAt = 1
        .LinkedStyle = "Heading 2" 'WordDocument.Styles(wdStyleListNumber).NameLocal
        tmpName = "Heading 2" 'WordDocument.Styles(wdStyleListNumber).NameLocal
    End With
    With WordDocument.Styles(tmpName)
        With .Font
            .Name = DefaultFontName
            .Size = 14
            .Bold = True
        End With
    End With
    
    With WordDocument.ListTemplates(strListTemplateName).ListLevels(3)
        .NumberFormat = "%2.%3."
        .TrailingCharacter = wdTrailingSpace
        .NumberStyle = wdListNumberStyleArabic
        .NumberPosition = Excel.Application.CentimetersToPoints(0.5)
        .Alignment = wdListLevelAlignLeft
        .TextPosition = Excel.Application.CentimetersToPoints(1)
        .TabPosition = Excel.Application.CentimetersToPoints(1)
        .ResetOnHigher = True
        .StartAt = 1
        .LinkedStyle = "Heading 3" 'WordDocument.Styles(wdStyleListNumber2).NameLocal
        tmpName = "Heading 3" 'WordDocument.Styles(wdStyleListNumber2).NameLocal
    End With
    With WordDocument.Styles(tmpName)
        With .Font
            .Name = DefaultFontName
            .Size = 13
            .Bold = True
        End With
    End With
    
    With WordDocument.ListTemplates(strListTemplateName).ListLevels(4)
        .NumberFormat = "%2.%3.%4."
        .TrailingCharacter = wdTrailingSpace
        .NumberStyle = wdListNumberStyleArabic
        .NumberPosition = Excel.Application.CentimetersToPoints(1)
        .Alignment = wdListLevelAlignLeft
        .TextPosition = Excel.Application.CentimetersToPoints(1.5)
        .TabPosition = Excel.Application.CentimetersToPoints(1.5)
        .ResetOnHigher = True
        .StartAt = 1
        .LinkedStyle = "Heading 4" 'WordDocument.Styles(wdStyleListNumber3).NameLocal
        tmpName = "Heading 4" 'WordDocument.Styles(wdStyleListNumber3).NameLocal
    End With
    With WordDocument.Styles(tmpName)
        With .Font
            .Name = DefaultFontName
            .Size = 12
            .Bold = True
            .Underline = True
        End With
    End With
    With WordDocument.ListTemplates(strListTemplateName).ListLevels(5)
        .NumberFormat = "%5."
        .TrailingCharacter = wdTrailingSpace
        .NumberStyle = wdListNumberStyleLowercaseLetter
        .NumberPosition = Excel.Application.CentimetersToPoints(1)
        .Alignment = wdListLevelAlignLeft
        .TextPosition = Excel.Application.CentimetersToPoints(1.5)
        .TabPosition = Excel.Application.CentimetersToPoints(1.5)
        .ResetOnHigher = True
        .StartAt = 1
        .LinkedStyle = "Heading 5" 'WordDocument.Styles(wdStyleListNumber4).NameLocal
        tmpName = "Heading 5" 'WordDocument.Styles(wdStyleListNumber4).NameLocal
    End With
    With WordDocument.Styles(tmpName)
        With .Font
            .Name = DefaultFontName
            .Italic = True
            .Bold = True
        End With
    End With
    With WordDocument.ListTemplates(strListTemplateName).ListLevels(6)
        .NumberFormat = "%6."
        .TrailingCharacter = wdTrailingSpace
        .NumberStyle = wdListNumberStyleNumberInCircle
        .NumberPosition = Excel.Application.CentimetersToPoints(1)
        .Alignment = wdListLevelAlignLeft
        .TextPosition = Excel.Application.CentimetersToPoints(1.5)
        .TabPosition = Excel.Application.CentimetersToPoints(1.5)
        .ResetOnHigher = True
        .StartAt = 1
        With .Font
            .Bold = True
        End With
        .LinkedStyle = "Heading 6" 'WordDocument.Styles(wdStyleListNumber5).NameLocal
    End With
    With WordDocument.ListTemplates(strListTemplateName).ListLevels(7)
        .NumberFormat = ""
        .LinkedStyle = ""
    End With
    With WordDocument.ListTemplates(strListTemplateName).ListLevels(8)
        .NumberFormat = ""
        .LinkedStyle = ""
    End With
    With WordDocument.ListTemplates(strListTemplateName).ListLevels(9)
        .NumberFormat = ""
        .LinkedStyle = ""
    End With
    
    '===Bullet & Normal
    With WordDocument.Styles("Normal")
        With .ParagraphFormat
            .LeftIndent = Excel.Application.CentimetersToPoints(0)
            .rightIndent = Excel.Application.CentimetersToPoints(0)
            .SpaceBefore = 3
            .SpaceAfter = 3
            .LineSpacingRule = wdLineSpaceMultiple
            .LineSpacing = WordApplication.Application.LinesToPoints(1.1)
            .Alignment = wdAlignParagraphJustify
            .FirstLineIndent = Excel.Application.CentimetersToPoints(1.27)
            .OutlineLevel = wdOutlineLevelBodyText
        End With
        .Font.Name = DefaultFontName
        .Font.Size = 12
        .NoSpaceBetweenParagraphsOfSameStyle = False
        .AutomaticallyUpdate = False
        .BaseStyle = ""
        .NextParagraphStyle = "Normal"
    End With
           
    With WordDocument
        If Not StyleExist(WordDocument, "TieudeVanban") Then .Styles.Add "TieudeVanban"
        With .Styles("TieudeVanban")
            .Font.Name = DefaultFontName
            .Font.Size = 12
            .Font.Bold = True
            .Font.AllCaps = True
            With .ParagraphFormat
                .FirstLineIndent = 0
                .TabStops.Add Position:=Excel.Application.CentimetersToPoints(2.86), Alignment:=wdAlignTabCenter, Leader:=wdTabLeaderSpaces
                .TabStops.Add Position:=Excel.Application.CentimetersToPoints(10.16), Alignment:=wdAlignTabCenter, Leader:=wdTabLeaderSpaces
            End With
        End With
        If Not StyleExist(WordDocument, "TieudeKehoach") Then .Styles.Add "TieudeKehoach"
        With .Styles("TieudeKehoach")
            .Font.Name = DefaultFontName
            .Font.Size = 16
            .Font.Bold = True
            .Font.AllCaps = True
            With .ParagraphFormat
                .Alignment = wdAlignParagraphCenter
                .SpaceBefore = 18
                .SpaceAfter = 12
                .LineSpacingRule = wdLineSpaceMultiple
                .LineSpacing = WordApplication.Application.LinesToPoints(1.1)
                .FirstLineIndent = 0
            End With
        End With
        If Not StyleExist(WordDocument, "Diemnhan") Then .Styles.Add "Diemnhan"
        With .Styles("Diemnhan")
            With .ParagraphFormat
                .LeftIndent = WordApplication.Application.CentimetersToPoints(1.6)
                .FirstLineIndent = WordApplication.Application.CentimetersToPoints(-0.6)
            End With
            .NextParagraphStyle = wdStyleNormal 'for international version compatibility
            .Font.Name = DefaultFontName
            .Font.Size = 12
        End With
        If Not StyleExist(WordDocument, "Bullet_type1") Then .Styles.Add "Bullet_type1"
        With .Styles("Bullet_type1")
            With .ParagraphFormat
                .LeftIndent = WordApplication.Application.CentimetersToPoints(1.6)
                .FirstLineIndent = WordApplication.Application.CentimetersToPoints(-0.6)
            End With
            .NextParagraphStyle = wdStyleNormal 'for international version compatibility
            .Font.Name = DefaultFontName
            .Font.Size = 12
            .Font.Bold = False
        End With
        If Not StyleExist(WordDocument, "HamucKehoach") Then .Styles.Add "HamucKehoach"
        With .Styles("HamucKehoach")
            .Font.Name = DefaultFontName
            .Font.Size = 15
            .Font.Bold = True
            .Font.AllCaps = True
            With .ParagraphFormat
                .Alignment = wdAlignParagraphCenter
                .LeftIndent = Excel.Application.CentimetersToPoints(6.03)
                .SpaceBeforeAuto = False
                .SpaceAfterAuto = False
                .FirstLineIndent = 0
            End With
        End With
        BulletText WordDocument, "Diemnhan"
        ' add page number here
        .Sections(1).Footers(1).PageNumbers.Add PageNumberAlignment:=wdAlignPageNumberCenter, FirstPage:=True
    End With
    Exit Function
ErrHandler:
    GenerateWordStyle = True
End Function

Private Sub BulletText(sDoc As Object, LinkObj As String)
    Dim myList As Object

    ' Add a new ListTemplate object
    Set myList = sDoc.ListTemplates.Add

    With myList.ListLevels(1)
        .NumberFormat = ChrW(254)
        .TrailingCharacter = wdTrailingTab
        .NumberPosition = Excel.Application.CentimetersToPoints(1)
        .Alignment = wdListLevelAlignLeft
        .TextPosition = Excel.Application.CentimetersToPoints(1.6)
        .TabPosition = Excel.Application.CentimetersToPoints(1.6)
        .ResetOnHigher = 0
        .StartAt = 1
        .LinkedStyle = LinkObj
        ' The following sets the font attributes of
        ' the "bullet" text
        With .Font
            .Bold = False
            .Name = "Wingdings"
            .Size = 12
        End With
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

'============================================
' For word table objects and so...
Sub InsertTable(RangeName As String)
    Dim prCount As Long, tmpObj As Object, CopyRange As Range
    Dim RngName As Variant, ColIndex As Variant
    Dim tmpWbk As Workbook, tmpSheet As Worksheet, i As Long
    Dim FilterColumn As Long, FilterObject As String
    Dim UseHeader As Boolean ' sometimes forgot to get the header of the table
    Dim Row2Copy As Long
    
    ' For inputdata
    RngName = Split(RangeName, "/")
    ' For showing column
    ColIndex = Split(RngName(2), ",")
    ' For column to limit
    FilterColumn = RngName(3)
    If RngName(4) <> "" Then FilterObject = Evaluate(RngName(4))
    UseHeader = Evaluate(RngName(5))
    ' Assign Range now
    Set CopyRange = ThisWorkbook.Names(RngName(1)).RefersToRange
    ' Now create a new workbook and format the table
    Set tmpWbk = Workbooks.Add
    Set tmpSheet = tmpWbk.Sheets.Add
    If UseHeader Then Set CopyRange = CopyRange.Resize(CopyRange.Rows.Count + 1).Offset(-1)
    CopyRange.Copy
    tmpSheet.Range("B1").PasteSpecial xlPasteAll
    Application.CutCopyMode = False
    
    ' Now change column size
    For i = 1 To CopyRange.Columns.Count
        tmpSheet.Columns(i + 1).ColumnWidth = CopyRange.Columns(i).ColumnWidth
    Next
    ' Now remove some rows if needed
    If FilterColumn > 0 Then
        Dim tCell As Range, DeletedAlready As Boolean
        
        Set tCell = tmpSheet.Cells(1, FilterColumn + 1)
        While tCell <> ""
            Row2Copy = tCell.Row
            If FilterObject <> "" Then
                If tCell = FilterObject Then
                    tCell.EntireRow.Delete
                    DeletedAlready = True
                End If
            End If
            
            If DeletedAlready Then
                Set tCell = tmpSheet.Cells(Row2Copy, FilterColumn + 1)
            Else
                Set tCell = tCell.Offset(1)
            End If
            DeletedAlready = False
            Row2Copy = tCell.Row
        Wend
        Row2Copy = Row2Copy - 1
    Else
        Row2Copy = CopyRange.Rows.Count
    End If
    ' Now disable some columns
    ' Build a string with column to be removed
    ' Remove some blank line
    
    ' Continue the next
    Dim tmpStr As String, relCol As Variant
    For i = 4 To CopyRange.Columns.Count
        tmpStr = tmpStr & "," & i
    Next
    
    For i = UBound(ColIndex) To LBound(ColIndex) Step -1
        If Val(ColIndex(i)) > 3 Then
            tmpStr = Replace(tmpStr, "," & CStr(ColIndex(i)), "")
        Else
            Exit For
        End If
    Next
    relCol = Split(tmpStr, ",")
    For i = UBound(relCol) To LBound(relCol) Step -1
        If Val(relCol(i)) > 3 Then
            tmpSheet.Columns(Val(relCol(i)) + 1).Delete Shift:=xlToLeft
        Else
            Exit For
        End If
    Next
    ' Now just copy them to word
    Set CopyRange = tmpSheet.Range("B1", tmpSheet.Cells(Row2Copy, UBound(ColIndex) + 2))
    With WordDocument
        CopyRange.Copy
        prCount = .Paragraphs.Count
        .Paragraphs(prCount).Range.PasteExcelTable False, True, True
        Set tmpObj = .Tables(.Tables.Count)
        With tmpObj
            .PreferredWidthType = wdPreferredWidthPercent
            .PreferredWidth = 100
            .Rows.HeightRule = wdRowHeightAtLeast
            .Rows.Height = Excel.Application.CentimetersToPoints(0)
            .Rows.LeftIndent = Excel.Application.CentimetersToPoints(0)
        End With
    End With
    Application.CutCopyMode = False
    Set tmpObj = Nothing
    Set tmpSheet = Nothing
    tmpWbk.Close False
    Set tmpWbk = Nothing
End Sub

Property Get TableCount() As Long
    On Error GoTo ErrHandler
    TableCount = WordDocument.Tables.Count
ErrHandler:
End Property

Private Sub CreateWordDocument()
    'Muc dich: Co gang thiet lap ket noi voi mot phien lam viec cua Word neu duoc. Neu khong thi tao moi
    'Co gang tao ket noi
    On Error Resume Next
    Set WordApplication = GetObject(, "Word.Application")
    If Err.Number <> 0 Then
        'Khong tao duoc ketnoi
        Err.Clear
        Set WordApplication = CreateObject("Word.Application")
        WordApplication.Visible = True
    End If
    Set WordDocument = WordApplication.Documents.Add
End Sub

Sub InsertPara(ItemStyle As ItemAttributes, ItemText As String, Optional OverideAdd As Boolean = False)
    'On Error Resume Next
    Dim prCount As Long, tmpText As String, tmpItem As ItemAttributes
    tmpItem = ItemStyle
    With WordDocument
        If ItemStyle.ItemHeading = "" Or ItemText = "" Then Exit Sub
        .Paragraphs.Add
        prCount = .Paragraphs.Count
        .Paragraphs(prCount).Range.Style = .Styles(ItemStyle.ItemHeading)
        .Paragraphs(prCount).Range.Text = ItemText
        
        If ItemStyle.ItemDetails <> "" And Not OverideAdd Then
            ' Add new introduction line if neccessary
            tmpItem.ItemHeading = tmpItem.ItemEmphasizeHeading
            tmpText = tmpItem.ItemDetails
            tmpItem.ItemDetails = ""
            InsertPara tmpItem, tmpText
        End If
    End With
End Sub

Sub InsertSection(Optional ToLastPage As Boolean = True)
    If ToLastPage Then
        WordDocument.Paragraphs.Last.Range.InsertBreak Type:=wdSectionBreakNextPage
    Else
        ' Just to current place
    End If
End Sub

Sub SetSectionLayout(Optional SetLandscape As Boolean = True)
    WordDocument.Sections.Last.PageSetup.Orientation = IIf(SetLandscape, wdOrientPortrait, wdOrientLandscape)
End Sub

Sub AddTable(WordTable As Object)
    Dim Tbl As Object
    Dim iCol As Long, iRow As Long, i As Long, j As Long
    iRow = WordTable.Rows.Count
    iCol = WordTable.Columns.Count
    
    Set Tbl = WordDocument.Tables.Add(WordDocument.Paragraphs.Last.Range, iRow, iCol)
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
            .Columns(i).PreferredWidth = 100 * WordTable.Columns(i).ColumnWidth / WordTable.Width
        Next
        Err.Clear
        For i = 1 To WordTable.Rows.Count
            For j = 1 To WordTable.Columns.Count
                .Cell(i, j) = Trim(WordTable.Cells(i, j))
                ' alignment
                Select Case WordTable.Cells(i, j).HorizontalAlignment
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

Sub RemoveTagAndFormat()
    Dim TagArr As Variant, i As Long, TagFormat As FontFormat, DefaultFontSize As Long
    TagArr = Array("bold", "allcap")
    DefaultFontSize = WordDocument.Styles("Normal").Font.Size
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
            FormatTags CStr(TagArr(i)), TagFormat
        Next
    End With
End Sub

Private Sub FormatTags(TagStr As String, ObjFormat As FontFormat)
'
' Macro1 Macro
'
'
    Dim DocRange As Object
    Set DocRange = WordDocument.Range
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

Sub ReformatWordTable(Optional Msg1 As String, Optional Msg2 As String, Optional MsgFin As String)
    Dim tmpObj As Object, prg As Object, i As Long, ErrCount As Long
    Dim DefaultFont As String
    DefaultFont = WordDocument.Styles("Normal").Font.Name
    For Each tmpObj In WordDocument.Tables
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

Sub SetHeaderRow(WordTable As Object)
    Dim HeaderRange As Object
    On Error GoTo ErrHandler
    
    Set HeaderRange = WordTable.Rows(1).Range
    HeaderRange.Rows.HeadingFormat = True
    Set HeaderRange = Nothing
    Exit Sub
ErrHandler:
    If Err.Number <> 0 Then Err.Clear
    Set HeaderRange = WordTable.Cell(1, 1).Range
    HeaderRange.SetRange Start:=WordTable.Cell(1, 1).Range.Start, End:=WordTable.Cell(1, 1).Range.Start
    Resume Next
End Sub

Private Sub FormatTable(WordTable As Object, Col2Format As Long, StartRow As Long)
'
' FormatTable Macro, will do the setting up of table and then get it updated quickly..
' The key is number of column to be formatted and starting row...
'
    Dim i As Long, ColNums As Long, MyCells As Object
    ' With table format
    With WordTable
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
        Set MyCells = WordDocument.Range(.Cell(StartRow + 1, ColNums - Col2Format + 1).Range.Start, .Cell(.Rows.Count, ColNums).Range.End)
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
