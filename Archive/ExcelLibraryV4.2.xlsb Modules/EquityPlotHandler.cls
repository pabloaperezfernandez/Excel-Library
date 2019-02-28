VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EquityPlotHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Base 1
Option Explicit
'Purpose   :    Create Tornedo Plot
'Inputs    :    cht                [chart object]
'               dataRange          [input data]
'               chartName          [chart name]
'               title              [chart title]
'               leftLegend         [y axis name]
'               xAxisLegend        [x axis name]
'Outputs   :    Table in Variant Format
'Author    :    Jesus Carrero
'Date      :    25 March 2014
'Notes     :
Sub TornedoPlot(chtObj As Object, wrkSheet As Worksheet, dataRange As Range, chartName As String, _
                title As String, leftLegend As String, xAxisLegend As String)
                
    chtObj.Name = chartName
    With chtObj
        .Chart.ChartType = xlBarClustered
        
        .Chart.HasTitle = True
        .Chart.ChartTitle.Characters.Text = title
        
        .Chart.Axes(xlCategory, xlPrimary).HasTitle = True
        .Chart.Axes(xlCategory, xlPrimary).AxisTitle.Characters.Text = leftLegend
        .Chart.Axes(xlValue, xlPrimary).HasTitle = True
        .Chart.Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = xAxisLegend
        
        .Chart.SetSourceData Source:=dataRange
        .Chart.ApplyLayout (1)
        
        ' Set Collection names
        .Chart.SeriesCollection(1).Name = "=""Returns"""
        .Chart.SeriesCollection(2).Name = "=""Tstats"""
    End With

    Call FormatPlots(chtObj, wrkSheet)
End Sub

'Purpose   :    Plots cumLOW, cumHIGH, cumSPREAD, RETURNS, Drawdown
'Inputs    :    cht           [chart object]
'               dataWsht      [worksheet where the data is]
'               row           [row position]
'               col           [col position]
'               title         [chart title]
'Outputs   :    Table in Variant Format
'Author    :    Jesus Carrero
'Date      :    25 March 2014
'Notes     :    Data should be formatted as output by the routine EquityDbHandler.ComputeReturns
Sub PlotMimickingReturns(cht As Object, dataWsht As Worksheet, _
                         row As Long, col As Long, title As String)
    ' Return Charts
    Dim topBoundary As Long
    Dim botBoundary As Long, tsLength As Long
    
    Let botBoundary = dataWsht.Range(dataWsht.Cells(row + 2, col).Address).End(xlDown).row
    Let tsLength = botBoundary - row + 1
    
    With cht
        .Chart.ChartType = xlLine
        .Chart.HasTitle = True
        .Chart.ChartTitle.Characters.Text = title
        .Chart.Axes(xlCategory, xlPrimary).HasTitle = True
        .Chart.Axes(xlCategory, xlPrimary).AxisTitle.Characters.Text = "Dates"
        .Chart.Axes(xlValue, xlPrimary).HasTitle = True
        .Chart.Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = " Percentage (%)"
        
        ' Set Data
        .Chart.SeriesCollection.NewSeries
        .Chart.SeriesCollection(1).Values = dataWsht.Range(dataWsht.Cells(row + 1, col + 1).Address).Resize(tsLength, 1)
        .Chart.SeriesCollection.NewSeries
        .Chart.SeriesCollection(2).Values = dataWsht.Range(dataWsht.Cells(row + 1, col + 2).Address).Resize(tsLength, 1)
        .Chart.SeriesCollection.NewSeries
        .Chart.SeriesCollection(3).Values = dataWsht.Range(dataWsht.Cells(row + 1, col + 3).Address).Resize(tsLength, 1)
        .Chart.SeriesCollection.NewSeries
        .Chart.SeriesCollection(4).Values = dataWsht.Range(dataWsht.Cells(row + 1, col + 4).Address).Resize(tsLength, 1)
        .Chart.SeriesCollection.NewSeries
        .Chart.SeriesCollection(5).Values = dataWsht.Range(dataWsht.Cells(6, col + 4).Address).Resize(tsLength, 1)
 
        ' Set Dates
        .Chart.SeriesCollection(1).XValues = dataWsht.Range(dataWsht.Cells(row + 1, col).Address).Resize(tsLength, 1)
        
        ' Set Collection names
        .Chart.SeriesCollection(1).Name = "=""High (L)"""
        .Chart.SeriesCollection(2).Name = "=""Low (L)"""
        .Chart.SeriesCollection(3).Name = "=""Spread (L)"""
        .Chart.SeriesCollection(4).Name = "=""Drawdown """
        .Chart.SeriesCollection(5).Name = "=""Return (R)"""
        
        ' Change Drawdown chart style
        .Chart.SeriesCollection(4).Select
        With Selection.Format.Fill
            .Visible = msoTrue
            .ForeColor.RGB = RGB(255, 192, 0)
            .Solid
        End With
        .Chart.SeriesCollection(4).ChartType = xlArea
        
        ' Returns on a second axis
        .Chart.SeriesCollection(5).AxisGroup = 2
        .Chart.SeriesCollection(5).ChartType = xlColumnClustered
    
    End With
    Call FormatInternalLines(cht)
End Sub

'Purpose   :    Create Histogram From data in datWsht at position (leftPosi, topPosi) and name as title
'Inputs    :    chtObj          [instance of chart object]
'               dataWsht        [worksheet where chart leaves]
'               row             [left margnin of data]
'               col             [top marging of data]
'               title           [chart title]
'               leftLegend      [y-axis legend]
'Outputs   :    Table in Variant Format
'Author    :    Jesus Carrero
'Date      :    25 March 2014
'Notes     :
Public Sub PlotHistogram(chtObj As Object, dataWsht As Worksheet, _
                         row As Long, col As Long, title As String, _
                         leftLegend As String)

    ' Return Charts
    Dim topBoundary As Long
    Dim botBoundary As Long, numBins As Long
    
    Let botBoundary = dataWsht.Range(dataWsht.Cells(row + 2, col).Address).End(xlDown).row
    Let numBins = botBoundary - row + 1
    
    With chtObj
        .Chart.ChartType = xlColumnClustered
        .Chart.HasTitle = True
        .Chart.ChartTitle.Characters.Text = title
        .Chart.Axes(xlCategory, xlPrimary).HasTitle = True
        .Chart.Axes(xlCategory, xlPrimary).AxisTitle.Characters.Text = leftLegend
         
        .Chart.SetSourceData Source:=dataWsht.Range(dataWsht.Cells(row, col + 1).Address).Resize(numBins, 1)
        .Chart.SeriesCollection(1).XValues = dataWsht.Range(dataWsht.Cells(row, col).Address).Resize(numBins, 1)

        ' Set Collection names
        .Chart.SeriesCollection(1).Name = "=""Count"""

        .Chart.SeriesCollection(1).Select
        With Selection.Format.Fill
            .Visible = msoTrue
            .ForeColor.ObjectThemeColor = msoThemeColorAccent1
            .ForeColor.TintAndShade = 0
            .ForeColor.Brightness = 0
            .Solid
        End With
    
        With Selection.Format.Fill
            .Visible = msoTrue
            .ForeColor.RGB = RGB(143, 231, 37)
            .Transparency = 0
            .Solid
        End With
    End With
    Call FormatInternalLines(chtObj)
End Sub

'Purpose   :    Remove all the charts in the spreadsheet.
'Inputs    :    targetWsht          [worksheet where chart leaves]
'Outputs   :    Table in Variant Format
'Author    :    Jesus Carrero
'Date      :    25 March 2014
'Notes     :
Public Sub RemovePlots(targetWsht As Worksheet)
    Dim chtObj As ChartObject
    For Each chtObj In targetWsht.ChartObjects
        chtObj.Delete
    Next
End Sub

'Purpose   :    Remove chart from sheet
'Inputs    :    chartName           [chart name]
'               targetWsht          [worksheet where chart leaves]
'Outputs   :    Table in Variant Format
'Author    :    Jesus Carrero
'Date      :    25 March 2014
'Notes     :
Public Sub RemoveOneChart(chartName As String, targetWsht As Worksheet)
    Dim chtObj As ChartObject
    For Each chtObj In targetWsht.ChartObjects
        If StrComp(chtObj.Name, chartName, vbTextCompare) = 0 Then
            chtObj.Delete
        End If
    Next
End Sub

'Purpose   :    Format Grid lines in Chart's Plots.
'Inputs    :    cht           [chart object]
'Outputs   :    Table in Variant Format
'Author    :    Jesus Carrero
'Date      :    25 March 2014
'Notes     :
Public Sub FormatInternalLines(cht As Object)
    With cht.Chart.Axes(xlValue).MajorGridlines.Format.Line
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorText1
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0
        .Transparency = 0
        .DashStyle = msoLineDash
    End With
End Sub

'Purpose   :    Format Chart's Frames
'Inputs    :    cht           [chart object]
'               targetWsht    [worksheet where chart leaves]
'Outputs   :    Table in Variant Format
'Author    :    Jesus Carrero
'Date      :    25 March 2014
'Notes     :
Public Sub FormatPlots(cht As Object, targetWsht As Worksheet)
    With targetWsht.Shapes(cht.Name).Fill
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorAccent1
        .ForeColor.TintAndShade = 0.3399999738
        .ForeColor.Brightness = 0
        .BackColor.ObjectThemeColor = msoThemeColorAccent1
        .BackColor.TintAndShade = 0.7649999857
        .BackColor.Brightness = 0
        .TwoColorGradient msoGradientHorizontal, 1
    End With
    targetWsht.Shapes(cht.Name).Line.Visible = msoFalse
    targetWsht.Shapes(cht.Name).Shadow.Type = msoShadow25
End Sub

