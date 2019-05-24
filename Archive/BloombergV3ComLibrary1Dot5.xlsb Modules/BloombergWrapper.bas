Attribute VB_Name = "BloombergWrapper"
Option Explicit

Public Function GetHistorialBloombergData(Securities() As String, _
                                          Fields() As String, _
                                          StartDate As Date, _
                                          EndDate As Date, _
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
    Dim SortedUniqueDateArray As Variant
    Dim ADate As Date
    Dim SecurityArray() As String
    Dim FieldArray() As String
    Dim OverrideFieldArray() As Variant
    Dim OverrideValuesArray() As Variant
    Dim c As Integer
    
    ReDim SecurityArray(0 To Length(Securities) - 1)
    For c = 1 To Length(Securities)
        Let SecurityArray(c - 1) = Securities(c)
    Next
    
    ReDim FieldArray(0 To Length(Fields) - 1)
    For c = 1 To Length(FieldArray)
        Let FieldArray(c - 1) = Fields(c)
    Next
    
    If IsMissing(OverrideFields) Then
        Let ResultsArray = b.historicalData(Securities:=SecurityArray, _
                                            Fields:=FieldArray, _
                                            StartDate:=StartDate, _
                                            EndDate:=EndDate, _
                                            adjustmentFollowDPDF:=adjustmentFollowDPDF, _
                                            adjustmentNormal:=adjustmentNormal, _
                                            adjustmentAbnormal:=adjustmentAbnormal, _
                                            adjustmentSplit:=adjustmentSplit)
    Else
        ReDim OverrideFieldArray(0 To Length(OverrideFields) - 1)
        ReDim OverrideValuesArray(0 To Length(OverrideValues) - 1)
        
        For c = 1 To Length(OverrideFields)
            Let OverrideFieldArray(c - 1) = OverrideFields(c)
            Let OverrideValuesArray(c - 1) = OverrideValues(c)
        Next
    
        Let ResultsArray = b.historicalData(Securities:=SecurityArray, _
                                            Fields:=FieldArray, _
                                            StartDate:=StartDate, _
                                            EndDate:=EndDate, _
                                            adjustmentFollowDPDF:=adjustmentFollowDPDF, _
                                            adjustmentNormal:=adjustmentNormal, _
                                            adjustmentAbnormal:=adjustmentAbnormal, _
                                            adjustmentSplit:=adjustmentSplit, _
                                            OverrideFields:=OverrideFieldArray, _
                                            OverrideValues:=OverrideValuesArray)
    End If
    
    Set TimeSeriesDict = New Dictionary
    Set FieldTimeSeriesDict = New Dictionary
    
    For FieldCounter = 1 To Length(Fields)
        ' Extract the time series for each stock
        For StockCounter = 1 To Length(Securities)
            Call TimeSeriesDict.Add(Key:=Securities(StockCounter), _
                                    Item:=ConvertBbResponseToStockDateDictionary(StockCounter - 1, _
                                                                                 FieldCounter - 1, _
                                                                                 ResultsArray _
                                                                                ) _
                                   )
        Next
    
        Let SortedUniqueDateArray = GetSortedDatesFromBloombergResponse(ResultsArray)
        
        ' Pre-allocate matrix to hold time series (stocks down and dates to the right)
        ReDim TimeSeriesMatrix(1 To Length(SortedUniqueDateArray) + 1, _
                               1 To UBound(ResultsArray, 2) - LBound(ResultsArray, 2) + 2)
                               
        Let TimeSeriesMatrix(1, 1) = Fields(FieldCounter)
        
        ' Write dates and data in matrix
        For DateCounter = LBound(ResultsArray, 2) To UBound(ResultsArray, 2)
            ' Insert date
            If DateCounter + 1 <= Length(SortedUniqueDateArray) Then
                Let TimeSeriesMatrix(1, DateCounter + 2) = SortedUniqueDateArray(DateCounter + 1)
            End If
        Next
    
        ' Insert field values
        For StockCounter = 1 To Length(Securities)
            Let TimeSeriesMatrix(StockCounter + 1, 1) = Securities(StockCounter)
            
            Set TimeSeriesDict = ConvertBbResponseToStockDateDictionary(StockCounter - 1, _
                                                                        FieldCounter, _
                                                                        ResultsArray)
            
            For DateCounter = LBound(ResultsArray, 2) To UBound(ResultsArray, 2)
                If DateCounter + 1 <= Length(SortedUniqueDateArray) Then
                    Let ADate = SortedUniqueDateArray(DateCounter + 1)
                    
                    Let TimeSeriesMatrix(StockCounter + 1, DateCounter + 2) = _
                        TimeSeriesDict.Item(Key:=ADate)
                End If
            Next
        Next
        
        Call FieldTimeSeriesDict.Add(Key:=Fields(FieldCounter), _
                                     Item:=TimeSeriesMatrix)
    Next

    If Length(FieldArray) = 1 Then
        Let GetHistorialBloombergData = First(FieldTimeSeriesDict.Items)
    Else
        Set GetHistorialBloombergData = FieldTimeSeriesDict
    End If
End Function

Public Function GetReferenceBloombergData(Securities() As String, _
                                          Fields() As String, _
                                          Optional OverrideFields As Variant, _
                                          Optional OverrideValues As Variant) As Variant
    Dim b As New BCOM_wrapper
    Dim r As Variant
    Dim of() As String
    Dim ov() As String
    Dim c As Integer

    If IsMissing(OverrideFields) Then
        Let r = Prepend(b.referenceData(Securities, Fields), Fields)
    Else
        Let of = OverrideFields
        Let ov = OverrideValues
    
        Let r = Prepend(b.referenceData(Securities, Fields, of, ov), Fields)
    End If

    Set b = Nothing
  
    Let GetReferenceBloombergData = Transpose(Prepend(Transpose(r), Prepend(Securities, Empty)))
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
