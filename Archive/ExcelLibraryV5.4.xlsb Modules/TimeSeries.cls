VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TimeSeries"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Base 1
Option Explicit

Private RangeDb As Range
Private CoordinateDict As Dictionary

' aRangeDb - contains the header row at the top, identifier column on the left,
'            and the data range starting on cell aRangeDb(2,2)
Public Sub InitializeWithRange(ARange As Range)
    Dim NIdentifiers As Long
    Dim NDates As Long
    Dim Headers As Variant
    Dim Identifiers As Variant
    Dim r As Long
    Dim c As Long

    ' Determine the number of time points and identifiers
    Let NIdentifiers = ARange.Rows.Count - 1
    Let NDates = ARange.Columns.Count - 1
    
    ' Extract headers
    Let Headers = ConvertTo1DArray(ARange(1, 2).Resize(1, NDates).Value2)
    Let Identifiers = ConvertTo1DArray(ARange(2, 1).Resize(NIdentifiers, 1).Value2)
        
    ' Set reference to the underlying range holding the data. Exclude headers row.
    Set RangeDb = ARange.Offset(1, 1).Resize(NIdentifiers, NDates)
    
    ' Initialize CoordinateDict
    Set CoordinateDict = New Dictionary
    
    ' Store references to cells corresponding to row r and column c
    For r = 1 To NIdentifiers
        For c = 1 To NDates
            Call CoordinateDict.Add(Key:=Headers(c) & "-" & Identifiers(r), Item:=RangeDb(r, c))
        Next c
    Next r
End Sub

Property Get GetRangeDb() As Range
    Set GetRangeDb = RangeDb
End Property

Property Get IdExistsQ(TheId As Variant) As Boolean
    Dim FirstTimePoint As Variant
    
    Let FirstTimePoint = GetIdentifiers(1).Value2

    Let IdExistsQ = CoordinateDict.Exists(FirstTimePoint & "-" & TheId)
End Property

Property Get TimePointExistsQ(TheTimePoint As Variant) As Boolean
    Dim FirstId As Variant
    
    Let FirstId = GetTimePoints(1).Value2

    Let TimePointExistsQ = CoordinateDict.Exists(TheTimePoint & "-" & FirstId)
End Property

Property Get GetTimePoints() As Range
    Set GetTimePoints = GetRangeDb.Rows(1).Offset(-1, 0).Cells
End Property

Property Get GetIdentifiers() As Range
    Set GetIdentifiers = GetRangeDb.Columns(1).Offset(0, -1).Cells
End Property

Property Get GetNumberTimePoints() As Long
    Let GetNumberTimePoints = GetRangeDb.Columns.Count
End Property

Property Get GetNumberIdentifiers() As Long
    Let GetNumberIdentifiers = GetRangeDb.Rows.Count
End Property

Property Get GetSecurityTimeSeries(TheId As Variant, Optional StartTimePoint As Variant, _
                                   Optional EndTimePoint As Variant) As Range
    Dim RowIndex As Long
    Dim StartTimePointIndex As Long
    Dim EndTimePointIndex As Long
    Dim StartCell As Range
    Dim EndCell As Range
        
    ' If either of start or end time points are unspecified, return the entire time series. Otherwise, return
    ' the security's time series between start and end.
    If IsMissing(StartTimePoint) Or IsMissing(EndTimePoint) Then
        Let RowIndex = GetDataPoint(TheId, GetTimePoints(1).Value2).row - GetRangeDb(1, 1).Offset(-1, 0).row
    
        Set GetSecurityTimeSeries = GetRangeDb.Rows(RowIndex).Cells
    Else
        Let RowIndex = GetDataPoint(TheId, GetTimePoints(1).Value2).row - GetRangeDb(1, 1).Offset(-1, 0).row
    
        Let StartTimePointIndex = _
            GetDataPoint(GetIdentifiers(1).Value2, StartTimePoint).Column - GetRangeDb(1, 1).Offset(0, -1).Column

        Let EndTimePoint = _
            GetDataPoint(GetIdentifiers(1).Value2, EndTimePoint).Column - GetRangeDb(1, 1).Offset(0, -1).Column
            
        Set StartCell = GetRangeDb.Cells(RowIndex, StartTimePointIndex)
        Set EndCell = GetRangeDb.Cells(RowIndex, EndTimePoint)
                    
        Set GetSecurityTimeSeries = GetRangeDb.Range(StartCell.Offset(-1, -1), EndCell.Offset(-1, -1))
    End If
End Property

Property Get GetTimePointSlice(TheTimePoint As Variant) As Range
    Dim ColumnIndex As Long
    
    Let ColumnIndex = GetDataPoint(GetIdentifiers(1).Value2, TheTimePoint).Column - GetRangeDb(1, 1).Offset(0, -1).Column
    
    Set GetTimePointSlice = GetRangeDb.Columns(ColumnIndex).Cells
End Property

' This sub adds a column to the time series
' Parameters:
' - TimePoint is an atomic value
' - DataPts is a 1D array with the same number of elements as the number of rows in RangeDb
Public Sub AddTimePoint(TimePoint As Variant, DataPts As Variant)
    Dim TheId As Variant

    ' Expand RangeDb by one column
    Set RangeDb = RangeDb.Resize(RangeDb.Rows.Count, RangeDb.Columns.Count + 1)
    
    ' Insert the new column
    Let RangeDb.Columns(RangeDb.Columns.Count).Cells.Value2 = Application.Transpose(DataPts)
    
    ' Insert the new TimePoint in the header row
    Let RangeDb(1, RangeDb.Columns.Count).Offset(-1, 0).Value2 = TimePoint
    
    ' Update CoordinateDict with new column
    For r = 1 To RangeDb.Rows.Count
        ' Extract the current id
        Let TheId = RangeDb(1, 1).Offset(r, -1).Value
        Call CoordinateDict.Add(Key:=TimePoint & "-" & TheId, Item:=RangeDb(r, RangeDb.Columns.Count))
    Next r
End Sub

' This sub adds a row to the time series
' Parameters:
' - Identifier is an atomic value
' - DataPts is a 1D array with the same number of elements as the number of columns in RangeDb
Public Sub AddIdentifier(Identifier As Variant, DataPts As Variant)
    Dim TheTimePoint As Variant

    ' Expand RangeDb by one column
    Set RangeDb = RangeDb.Resize(RangeDb.Rows.Count + 1, RangeDb.Columns.Count)
    
    ' Insert the new column
    Let RangeDb.Rows(RangeDb.Columns.Count).Cells.Value2 = DataPts
    
    ' Insert the new identifier in the header row
    Let RangeDb(RangeDb.Rows.Count, 1).Offset(1, -1).Value2 = Identifier
    
    ' Update CoordinateDict with new row
    For c = 1 To RangeDb.Rows.Count
        ' Extract the current id
        Let TheTimePoint = RangeDb(1, 1).Offset(-1, c).Value
        Call CoordinateDict.Add(Key:=TheTimePoint & "-" & Identifier, Item:=RangeDb(RangeDb.Rows.Count, c))
    Next c
End Sub

' Returns a reference to cell corresponding the given identifier and point in time
Property Get GetDataPoint(Identifier As Variant, TimePoint As Variant) As Range
    Set GetDataPoint = CoordinateDict.Item(Key:=TimePoint & "-" & Identifier)
End Property

' Set the value held by cell corresponding to the given identifer and point in time
Property Let SetDataPoint(Identifier As Variant, TimePoint As Variant, TheValue As Variant)
    Let Me.GetDataPoint(Identifier, TimePoint).Value2 = TheValue
End Property
