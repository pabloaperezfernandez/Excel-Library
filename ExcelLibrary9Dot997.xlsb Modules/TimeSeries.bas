Attribute VB_Name = "TimeSeries"
' PURPOSE OF THIS MODULE
'
' The purpose of this module is to provide faciities to handle time series operations

Option Explicit
Option Base 1

' DESCRIPTION
' Returns a 2D matrix time series representation of the given DB-like table
'
' PARAMETERS
' 1. DataTableArray - A 2D DB-like table with headers, with at least three columns
' 2. DateHeader - String header of column holding the dates
' 3. RowTagsHeader - String header of the column holding the tags to use as items along
'    the left-most column of the times series matrix
' 4. DataHeader - String header of the column holding the data to summarize by dates and
'    data tags
' 5. DateSortOrder (Optional) - xlAscending by default. Dictates the sorting order of the
'    dates.
'
' RETURNED VALUE
' Returns a times series 2D matrix produced from a 2D, DB-like times series table
Public Function DbTableToTimeSeriesTable(DataTableArray As Variant, _
                                         DateHeader As String, _
                                         RowTagsHeader As String, _
                                         DataHeader As String, _
                                         Optional DateSortOrder As XlSortOrder = xlAscending) As Variant
    Dim DatesIndex As Integer
    Dim TagsIndex As Integer
    Dim DataIndex As Integer
    Dim TheDates As Variant
    Dim DataTags As Variant
    Dim TimeSeriesTable() As Variant
    Dim TagsDict As Dictionary
    Dim DatesDict As Dictionary
    Dim r As Long
    
    ' Initialize dictionaries to hold row and column positions in return array
    Set TagsDict = New Dictionary
    Set DatesDict = New Dictionary
    
    ' Get column indices of the dates, tags, and data columns in DB-like table
    Let DatesIndex = Application.Match(DateHeader, First(DataTableArray), 0)
    Let TagsIndex = Application.Match(RowTagsHeader, First(DataTableArray), 0)
    Let DataIndex = Application.Match(DataHeader, First(DataTableArray), 0)
    
    ' Compute and sort the set of unique dates
    Let TheDates = Sort1DArray(UniqueSubset(Part(DataTableArray, Span(2, -1), DatesIndex)), _
                               DateSortOrder)
    ' Compute and sort the sert of unique tags
    Let DataTags = Sort1DArray(UniqueSubset(Part(DataTableArray, Span(2, -1), TagsIndex)), _
                               xlAscending)
                           
    ' Index the unique dates
    For r = 1 To Length(TheDates)
        Call DatesDict.Add(Key:=Part(TheDates, r), Item:=r + 1)
    Next r
    
    ' Index the unique tags
    For r = 1 To Length(DataTags)
        Call TagsDict.Add(Key:=Part(DataTags, r), Item:=r + 1)
    Next
                               
    ' Pre-allocate the times series array to return
    ReDim TimeSeriesTable(1 To Length(DataTags) + 1, 1 To Length(TheDates) + 1)
    
    ' Insert the data tags in return time series table
    For r = 1 To Length(DataTags)
        Let TimeSeriesTable(1 + r, 1) = Part(DataTags, r)
    Next
    
    ' Insert the dates in return time series table
    For r = 1 To Length(TheDates)
        Let TimeSeriesTable(1, 1 + r) = Part(TheDates, r)
    Next
    
    ' Loop through the original data, populating the return time series table
    For r = 2 To Length(DataTableArray)
        Let TimeSeriesTable(TagsDict.Item(Key:=Part(DataTableArray, r, TagsIndex)), _
                            DatesDict.Item(Key:=Part(DataTableArray, r, DatesIndex))) = _
            DataTableArray(r, DataIndex)
    Next
    
    ' Return the times series table
    Let DbTableToTimeSeriesTable = TimeSeriesTable
End Function

' DESCRIPTION
' Returns a DB-like table from the given 2D matrix time series
'
' PARAMETERS
' 1. TimeSeriesTable - A times series 2D matrix, with dates along the top row and
'    data tags along the first column
' 2. DateHeader - String header to use for column holding the dates
' 3. RowTagsHeader - String header to use for column holding the tags
' 4. DataHeader - String header to use for the column holding the data
'
' RETURNED VALUE
' Returns a times series 2D matrix produced from a 2D, DB-like times series table
Public Function TimeSeriesTableToDbTable(TimeSeriesTable As Variant, _
                                         DateHeader As String, _
                                         RowTagsHeader As String, _
                                         DataHeader As String) As Variant
    Dim DbTable() As Variant
    Dim c As Long
    Dim r As Long
    Dim TheDates As Variant
    Dim TheTags As Variant
    Dim TheData As Variant
    Dim NumIdentifers As Long
    Dim NumDates As Long
    
    Let NumIdentifers = NumberOfRows(TimeSeriesTable) - 1
    Let NumDates = NumberOfColumns(TimeSeriesTable) - 1
    
    ReDim DbTable(1 To NumDates * NumIdentifers + 1, 1 To 3)
    
    Let DbTable(1, 1) = DateHeader
    Let DbTable(1, 2) = RowTagsHeader
    Let DbTable(1, 3) = DataHeader

    For r = 1 To NumIdentifers
        For c = 1 To NumDates
            Let DbTable((r - 1) * NumDates + c + 1, 1) = TimeSeriesTable(1, c + 1)
            Let DbTable((r - 1) * NumDates + c + 1, 2) = TimeSeriesTable(r + 1, 1)
            Let DbTable((r - 1) * NumDates + c + 1, 3) = TimeSeriesTable(r + 1, c + 1)
        Next
    Next
    
    Let TimeSeriesTableToDbTable = DbTable
End Function

Public Function TimeSeriesMovingAverage(TimeSeriesArray As Variant, NumberPtsToAvg As Integer) As Variant
    Dim MaArray As Variant
    Dim DateChunk As Variant
    
    Let TimeSeriesMovingAverage = Null
    
    If Not DateArrayQ(Part(TimeSeriesArray, Span(1, -1), 1)) Then Exit Function
    
    Let MaArray = MovingAverage(Part(TimeSeriesArray, Span(1, -1), 2), NumberPtsToAvg)
    
    If NullQ(MaArray) Then Exit Function
    
    Let DateChunk = Part(TimeSeriesArray, Span(CLng(NumberPtsToAvg), -1), 1)
    
    Let TimeSeriesMovingAverage = Transpose(Pack2DArray(Array(DateChunk, MaArray)))
End Function

Public Function TimeSeriesSimpleAveragePerformance(TimeSeriesArray As Variant, NumberPtsToAvg As Integer) As Variant
    Dim PerfArray As Variant
    Dim DateChunk As Variant
    
    Let TimeSeriesSimpleAveragePerformance = Null
    
    If Not DateArrayQ(Part(TimeSeriesArray, Span(1, -1), 1)) Then Exit Function
    
    Let PerfArray = SimpleAveragePerformance(Part(TimeSeriesArray, Span(1, -1), 2), NumberPtsToAvg)
    
    If NullQ(PerfArray) Then Exit Function
    
    Let DateChunk = Part(TimeSeriesArray, Span(CLng(NumberPtsToAvg), -1), 1)
    
    Let TimeSeriesSimpleAveragePerformance = Transpose(Pack2DArray(Array(DateChunk, PerfArray)))
End Function
