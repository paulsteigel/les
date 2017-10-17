Option Explicit

Sub AddChartObject(Wrd As Object, DataSource As Object, Series As Range, ChartTitle As String, Optional ChartType As Long = 1)
'
' This will add a chart object to an excel file and will copy it to word...
'

'
    Dim tChart As Shape, tmpSheet As Worksheet, ChartRange As Range, cellPtr As Range
    Dim i As Long, wb As Workbook
    Set wb = Application.Workbooks.Add
    Set tmpSheet = wb.Sheets.Add
    
    Set tChart = tmpSheet.Shapes.AddChart
    ' Add data source
    Set cellPtr = tmpSheet.Range("A1")
    
    ' Now pull  in the series
    For i = 0 To DataSource.Fields.Count - 1
        cellPtr.Offset(0, i) = Series.Offset(0, i)
    Next

    Set cellPtr = cellPtr.Offset(1)
    While Not DataSource.EOF
        For i = 0 To DataSource.Fields.Count - 1
            cellPtr.Offset(0, i) = DataSource.Fields(i)
        Next
        Set cellPtr = cellPtr.Offset(1)
        DataSource.MoveNext
    Wend
        
    Set ChartRange = Range(tmpSheet.Range("A1"), cellPtr.Offset(-1, i - 1))
    ' now set the title
    cellPtr = ChartTitle
    If ChartType = 1 Then tmpSheet.Range("A1") = ""
    ' set up chart data
    With tChart.Chart
        If ChartType = 1 Then
            .SetSourceData ChartRange
            'SetChartTypePie tChart
            .ChartType = xlColumnClustered
            '.SeriesCollection(1).XValues = Range(tmpSheet.Range("A1"), tmpSheet.Range("A1").Offset(0, DataSource.Fields.Count - 1))
        Else
            .ChartType = xlColumnStacked100
            ' Automatic plot chart to the expected format
            .SetSourceData ChartRange, xlRows
        End If
        
        .SetElement (msoElementChartTitleAboveChart)
        .ChartTitle.Caption = cellPtr
        .ChartArea.Copy
    End With
    
    Dim prg As Object
    Wrd.Paragraphs.Add
    Set prg = Wrd.Paragraphs.Add
    prg.Range.PasteSpecial ' Link:=False, DataType:=wdPasteEnhancedMetafile, Placement:=wdInLine, DisplayAsIcon:=False
    Set tChart = Nothing
    wb.Close False
    DataSource.Close
End Sub

Private Sub SetChartTypeCluster(tChart As Shape)
'
' Macro3 Macro
'

'
    With tChart
        .ChartType = xlColumnClustered
        .SeriesCollection(1).XValues = "=Sheet4!$B$1:$E$1"
    End With
End Sub

Private Sub SetChartTypePie(tChart As Shape)
    With tChart.Chart
    .ChartType = xl3DPie
    .SetElement (msoElementChartTitleAboveChart)
    .SetElement (msoElementLegendNone)
    .SetElement (msoElementDataLabelOutSideEnd)
    .ApplyDataLabels
        With .SeriesCollection(1).DataLabels
            .ShowSeriesName = True
            .ShowSeriesName = False
            .ShowCategoryName = True
            .Position = xlLabelPositionOutsideEnd
            .ShowPercentage = True
            .ShowValue = False
        End With
    End With
End Sub

