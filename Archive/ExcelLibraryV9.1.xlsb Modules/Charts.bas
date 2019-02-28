Attribute VB_Name = "Charts"
Option Base 1
Option Explicit
'
'Public Sub CreateChartExample()
'    Dim aShape As Shape
'    Dim aChart As Chart
'
'    Call DumpInSheet([{1,2;10,20;100,200}], Sheet1.Range("A1"))
'
'    Set aShape = Sheet1.Shapes.AddChart(xlLine, 100, 100, 100, 100)
'    Set aChart = aShape.Chart
'    Call aChart.SetSourceData(Sheet1.Range("A1").CurrentRegion)
'End Sub
'
'Public Sub CreateChartExample2()
'    Dim aChart As ChartObject
'
'    Call DumpInSheet([{1,2;10,20;100,200}], TempComputation.Range("A1"))
'
'    Set aChart = TempComputation.ChartObjects.Add(100, 100, 100, 100)
'    Let aChart.Chart.ChartType = xlLine
'    Call aChart.Chart.SetSourceData(TempComputation.Range("A1").CurrentRegion)
'End Sub

Public Function GenerateLineChartTimeSeries(TargetWorksheet As Worksheet, _
                                            SourceRange As Range, _
                                            TheTitle As String, _
                                            XAxisData As Range, _
                                            ToPlotBy As XlRowCol) As ChartObject
    Dim co As ChartObject

    Set co = TargetWorksheet.ChartObjects.Add(Left:=10, Width:=500, Top:=10, height:=300)
    Let co.Chart.ChartArea.Fill.Visible = msoTrue
    Let co.Chart.PlotArea.Format.Fill.Visible = msoFalse
    Let co.Chart.ChartType = xlLine
    Call co.Chart.SetSourceData(Source:=SourceRange, PlotBy:=ToPlotBy)
    Call co.Chart.SetElement(msoElementLegendBottom)
    Call co.Chart.SetElement(msoElementChartTitleAboveChart)
    Let co.Chart.ChartTitle.Text = TheTitle
    Let co.Chart.Axes(xlValue).MajorGridlines.Format.Line.Visible = msoTrue
    Let co.Chart.Axes(xlValue).MajorGridlines.Format.Line.Weight = 0.25
    Let co.Chart.Axes(xlValue).MajorGridlines.Format.Line.DashStyle = msoLineDash
    Let co.Chart.ChartArea.Border.LineStyle = xlNone
    Let co.Chart.SeriesCollection(1).XValues = XAxisData
    
    Set GenerateLineChartTimeSeries = co
End Function

' DESCRIPTION
' Creates a box plot of the given data in the requested worksheet. It
' returns a reference to the chart
'
' EXAMPLE
' Call BoxPlot([{"Set1", "Set2", "Set3"; 1,2,3;4,5,6;7,8,9}], _
'              ThisWorkbook)
'
' PARAMETERS
' 1. TheData - A matrix of data, with each column representing a data set
'    for which a box plot is required
' 2. TargetWorksheet - Worksheet where data is deposited and chart is drawn
'
' RETURNED VALUE
' 1. Returns a reference to the chart object
' 2. As a side effect, the data table underlying the box plot
'    This data starts in range("A1") of the target worksheet
Public Function BoxPlot(TheData As Variant, CallingWorkbook As Workbook, _
                        Optional TargetWorksheet As Variant) As ChartObject
    Dim c As Integer
    Dim DataMatrix As Variant
    Dim TheDates() As Date
    Dim Maximums() As Double
    Dim ThirdQuartiles() As Double
    Dim Medians() As Double
    Dim FirstQuartiles() As Double
    Dim Minimums() As Double
    Dim DataVector As Variant
    Dim co As ChartObject
    Dim PlotDataRange As Range
    Dim PlotWsht As Worksheet
    
    ' Set PlotWsht to either the given worksheet or a new one
    If IsMissing(TargetWorksheet) Then
        Set PlotWsht = CallingWorkbook.Worksheets.Add
    Else
        Set PlotWsht = TargetWorksheet
    End If
    
    ' Pre-allocate arrays to hold the data underlying the boxplot
    ' All of these items will be computed for each of the data columns
    ' passed via parameter TheData
    ReDim TheDates(1 To NumberOfColumns(TheData), 1 To 1)
    ReDim Maximums(1 To NumberOfColumns(TheData), 1 To 1)
    ReDim ThirdQuartiles(1 To NumberOfColumns(TheData), 1 To 1)
    ReDim Medians(1 To NumberOfColumns(TheData), 1 To 1)
    ReDim FirstQuartiles(1 To NumberOfColumns(TheData), 1 To 1)
    ReDim Minimums(1 To NumberOfColumns(TheData), 1 To 1)
    
    ' Insert the headers for the boxplot data
    Call DumpInSheet([{"DataSet","1stQuartile", "High","Low","3rd Quartile","Median"}], _
                     PlotWsht.Range("A1"))
    
    ' Produce the statistics required to generate the boxplot
    For c = 1 To NumberOfColumns(TheData)
        ' Get the current column of data and sort it in ascending order
        Let DataVector = Sort1DArray(Flatten(Part(TheData, Span(2, -1), c)))
        
        ' Compute the required statistics for this data set
        'slet theates(c,1) =
        Let Maximums(c, 1) = Application.Max(DataVector)
        Let ThirdQuartiles(c, 1) = Application.Quartile_Inc(DataVector, 3)
        Let Medians(c, 1) = Application.Median(DataVector)
        Let FirstQuartiles(c, 1) = Application.Quartile_Inc(DataVector, 1)
        Let Minimums(c, 1) = Application.Min(DataVector)
        Let TheDates(c, 1) = Now + c
    Next
    
    ' Store the statistics in the target worksheet
    Call DumpInSheet(TheDates, PlotWsht.Range("A2"))
    Call DumpInSheet(FirstQuartiles, PlotWsht.Range("B2"))
    Call DumpInSheet(Maximums, PlotWsht.Range("C2"))
    Call DumpInSheet(Minimums, PlotWsht.Range("D2"))
    Call DumpInSheet(ThirdQuartiles, PlotWsht.Range("E2"))
    Call DumpInSheet(Medians, PlotWsht.Range("F2"))
        
    ' Format the dates column appropriately. This is required by the stock chart template
    Let PlotWsht.Range("A:A").NumberFormat = "M/D/YYYY"
    
    ' Create chart object
    Set co = PlotWsht.ChartObjects.Add(Left:=10, Width:=500, Top:=10, height:=300)
    
    ' Set the range holding the data. In this case the first four columns only
    Set PlotDataRange = PlotWsht.Range("a1").CurrentRegion
    Set PlotDataRange = PlotDataRange.Resize(PlotDataRange.Rows.Count, _
                                             PlotDataRange.Columns.Count - 1)
    
    Call co.Chart.SetSourceData(Source:=PlotDataRange)
    
    ' Set the chart typpe to OHLC (a type of stock chart)
    Let co.Chart.ChartType = xlStockOHLC
    
    ' Replace values in the dates column by data set labels in row 1
    Call DumpInSheet(TransposeMatrix(First(TheData)), PlotWsht.Range("A2"))
    
    ' Delete the chart's legend
    Call co.Chart.Legend.Delete

    ' Make invisible the inside of the upbars between the 1nd and 3rd quartiles
    Let co.Chart.ChartGroups(1).UpBars.Format.Fill.Visible = msoFalse
    
    ' The borders of the box visible and black
    With co.Chart.ChartGroups(1).UpBars.Format.Line
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorText1
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0
    End With
    
    ' Add medians to the chart and format them properly
    co.Chart.SeriesCollection.NewSeries
    co.Chart.SeriesCollection(5).Name = "=""Median"""
    Set PlotDataRange = PlotWsht.Range("a1").CurrentRegion
    Set PlotDataRange = PlotDataRange.Columns(6)
    Set PlotDataRange = PlotDataRange.Resize(PlotDataRange.Rows.Count - 1, 1).Offset(1, 0)
    Let co.Chart.SeriesCollection(5).Values = "='" & PlotWsht.Name & "'!" & _
                                              PlotDataRange.Address

    ' You have to switch between plotting against primary and secondary axis
    ' and black or the top half of the box becomes invisble. Not sure why, but
    ' this solves the problem
    Let co.Chart.SeriesCollection(5).AxisGroup = 2
    Let co.Chart.SeriesCollection(5).AxisGroup = 1
    
    ' Set the size and style of the median market to a red line
    With co.Chart.SeriesCollection(5)
        .MarkerStyle = -4115
        .MarkerSize = 20
    End With
    
    ' Set the color of the median market to solid red
    With co.Chart.SeriesCollection(5).Format.Fill
        .Visible = msoTrue
        .ForeColor.RGB = RGB(255, 0, 0)
    End With
    
    ' Make the gridlines dashed and black
    With co.Chart.Axes(xlValue).MajorGridlines.Format.Line
        .Visible = msoTrue
        .DashStyle = msoLineDash
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorText1
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0
        .Transparency = 0
    End With
End Function

