Attribute VB_Name = "BloombergWrapper"
Option Explicit

Public Function GetHistorialBloombergData(securities() As String, _
                                          fields() As String, _
                                          startDate As Date, _
                                          endDate As Date, _
                                          Optional ByVal adjustmentFollowDPDF As Boolean = True, _
                                          Optional ByVal adjustmentNormal As Boolean = True, _
                                          Optional ByVal adjustmentAbnormal As Boolean = True, _
                                          Optional ByVal adjustmentSplit As Boolean = True, _
                                          Optional OverrideFields As Variant, _
                                          Optional OverrideValues As Variant) As Variant
    Dim b As New BCOM_wrapper
    Dim ResultsArray As Variant
    Dim FieldTimeSeriesDict As Dictionary
    Dim StockCounter As Integer
    Dim DateCounter As Integer
    Dim FieldCounter As Integer
    Dim TimeSeriesDict As Dictionary
    Dim ArrayOfTimeSeries() As Variant
    Dim SortedUniqueDateArray As Variant
    Dim ADate As Date
    
    If IsMissing(OverrideFields) Then
        Let ResultsArray = b.historicalData(securities:=securities, _
                                            fields:=fields, _
                                            startDate:=startDate, _
                                            endDate:=endDate, _
                                            adjustmentFollowDPDF:=adjustmentFollowDPDF, _
                                            adjustmentNormal:=adjustmentNormal, _
                                            adjustmentAbnormal:=adjustmentAbnormal, _
                                            adjustmentSplit:=adjustmentSplit)
    Else
        Let ResultsArray = b.historicalData(securities:=securities, _
                                            fields:=fields, _
                                            startDate:=startDate, _
                                            endDate:=endDate, _
                                            adjustmentFollowDPDF:=adjustmentFollowDPDF, _
                                            adjustmentNormal:=adjustmentNormal, _
                                            adjustmentAbnormal:=adjustmentAbnormal, _
                                            adjustmentSplit:=adjustmentSplit, _
                                            OverrideFields:=OverrideFields, _
                                            OverrideValues:=OverrideValues)
    End If
    
    ReDim ArrayOfTimeSeries(1 To Length(fields))
    Set TimeSeriesDict = New Dictionary
    Set FieldTimeSeriesDict = New Dictionary
    
    For FieldCounter = 1 To Length(fields)
        ' Extract the time series for each stock
        For StockCounter = LBound(ResultsArray, 1) To UBound(ResultsArray, 1)
            Call TimeSeriesDict.Add(Key:=securities(StockCounter), _
                                    Item:=ConvertBbResponseToStockDateDictionary(StockCounter, _
                                                                                 FieldCounter, _
                                                                                 ResultsArray _
                                                                                ) _
                                   )
        Next
    
        Let SortedUniqueDateArray = GetSortedDatesFromBloombergResponse(ResultsArray)
        
        ' Pre-allocate matrix to hold time series (stocks down and dates to the right)
        ReDim TimeSeriesMatrix(1 To Length(SortedUniqueDateArray) + 1, _
                               1 To UBound(ResultsArray, 2) - LBound(ResultsArray, 2) + 2)
                               
        Let TimeSeriesMatrix(1, 1) = fields(FieldCounter - 1)
        
        ' Write dates and data in matrix
        For DateCounter = LBound(ResultsArray, 2) To UBound(ResultsArray, 2)
            ' Insert date
            If DateCounter + 1 <= Length(SortedUniqueDateArray) Then
                Let TimeSeriesMatrix(1, DateCounter + 2) = SortedUniqueDateArray(DateCounter + 1)
            End If
        Next
    
        ' Insert field values
        For StockCounter = LBound(ResultsArray, 1) To UBound(ResultsArray, 1)
            Let TimeSeriesMatrix(StockCounter + 2, 1) = securities(StockCounter)
            
            Set TimeSeriesDict = ConvertBbResponseToStockDateDictionary(StockCounter, FieldCounter, ResultsArray)
            
            For DateCounter = LBound(ResultsArray, 2) To UBound(ResultsArray, 2)
                If DateCounter + 1 <= Length(SortedUniqueDateArray) Then
                    Let ADate = SortedUniqueDateArray(DateCounter + 1)
                    
                    Let TimeSeriesMatrix(StockCounter + 2, DateCounter + 2) = TimeSeriesDict.Item(Key:=ADate)
                End If
            Next
        Next
        
        Call FieldTimeSeriesDict.Add(Key:=fields(FieldCounter - 1), _
                                     Item:=TimeSeriesMatrix)
    Next

    If Length(fields) = 1 Then
        Let GetHistorialBloombergData = First(FieldTimeSeriesDict.Items)
    Else
        Set GetHistorialBloombergData = FieldTimeSeriesDict
    End If
End Function

Public Function GetReferenceBloombergData(securities() As String, _
                                          fields() As String, _
                                          Optional OverrideFields As Variant, _
                                          Optional OverrideValues As Variant) As Variant
    Dim b As New BCOM_wrapper
    Dim r As Variant
    Dim of() As String
    Dim ov() As String
    Dim c As Integer

    If IsMissing(OverrideFields) Then
        Let r = Prepend(b.referenceData(securities, fields), fields)
    Else
        Let of = OverrideFields
        Let ov = OverrideValues
    
        Let r = Prepend(b.referenceData(securities, fields, of, ov), fields)
    End If

    Set b = Nothing
  
    Let GetReferenceBloombergData = Transpose(Prepend(Transpose(r), Prepend(securities, Empty)))
End Function

' Returns a dictionary indexing the field values by date
Public Function ConvertBbResponseToStockDateDictionary(StockCounter As Integer, _
                                                       FieldCounter As Integer, _
                                                       BloombergResponseMatrix As Variant) As Dictionary
    Dim ADict As Dictionary
    Dim DateCounter As Long
    Dim ADate As Date
    Dim AnItem As Variant
    
    Set ADict = New Dictionary
    
    ' Loop along the date dimension
    For DateCounter = LBound(BloombergResponseMatrix, 2) To UBound(BloombergResponseMatrix, 2)
        If Not EmptyQ(BloombergResponseMatrix(StockCounter, DateCounter)) Then
            Let ADate = BloombergResponseMatrix(StockCounter, DateCounter)(0)
            Let AnItem = BloombergResponseMatrix(StockCounter, DateCounter)(FieldCounter)
            Call ADict.Add(Key:=ADate, Item:=AnItem)
        End If
    Next
    
    Set ConvertBbResponseToStockDateDictionary = ADict
End Function

Public Function GetSortedDatesFromBloombergResponse(BloombergResponseMatrix As Variant) As Variant
    Dim DateCounter As Long
    Dim StockCounter As Integer
    Dim ADict As Dictionary
    
    Set ADict = New Dictionary
    
    For StockCounter = LBound(BloombergResponseMatrix, 1) To UBound(BloombergResponseMatrix, 1)
        For DateCounter = LBound(BloombergResponseMatrix, 2) To UBound(BloombergResponseMatrix, 2)
            If Not EmptyQ(BloombergResponseMatrix(StockCounter, DateCounter)) Then
                If Not ADict.Exists(Key:=BloombergResponseMatrix(StockCounter, DateCounter)(0)) Then
                    Call ADict.Add(Key:=BloombergResponseMatrix(StockCounter, DateCounter)(0), Item:=Empty)
                End If
            End If
        Next
    Next
    
    Let GetSortedDatesFromBloombergResponse = Sort1DArray(UniqueSubset(ADict.Keys))
End Function
