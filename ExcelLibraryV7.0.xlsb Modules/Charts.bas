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
