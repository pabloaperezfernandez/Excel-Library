Attribute VB_Name = "BloombergWrapper"
Option Explicit

Public Function GetHistorialBloombergDataData(securities() As String, _
                                              Fields() As String, _
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
    Dim TimeSeriesMatrix() As Variant
    Dim StockCounter As Integer
    Dim DateCounter As Integer
    
    If IsMissing(OverrideFields) Then
        Let ResultsArray = b.historicalData(securities:=securities, _
                                            Fields:=Fields, _
                                            startDate:=startDate, _
                                            endDate:=endDate, _
                                            adjustmentFollowDPDF:=adjustmentFollowDPDF, _
                                            adjustmentNormal:=adjustmentNormal, _
                                            adjustmentAbnormal:=adjustmentAbnormal, _
                                            adjustmentSplit:=adjustmentSplit)
    Else
        Let ResultsArray = b.historicalData(securities:=securities, _
                                            Fields:=Fields, _
                                            startDate:=startDate, _
                                            endDate:=endDate, _
                                            adjustmentFollowDPDF:=adjustmentFollowDPDF, _
                                            adjustmentNormal:=adjustmentNormal, _
                                            adjustmentAbnormal:=adjustmentAbnormal, _
                                            adjustmentSplit:=adjustmentSplit, _
                                            OverrideFields:=OverrideFields, _
                                            OverrideValues:=OverrideValues)
    End If

    ReDim TimeSeriesMatrix(1 To UBound(ResultsArray, 1) - LBound(ResultsArray, 1) + 2, _
                           1 To UBound(ResultsArray, 2) - LBound(ResultsArray, 2) + 2)

    ' Write the stock tickers to the left column of the matrix
    For StockCounter = LBound(ResultsArray, 1) To UBound(ResultsArray, 1)
        Let TimeSeriesMatrix(StockCounter + 2, 1) = securities(StockCounter)
    Next
    
    ' Write dates and data in matrix
    For DateCounter = LBound(ResultsArray, 2) To UBound(ResultsArray, 2)
        ' Insert date
        Let TimeSeriesMatrix(1, DateCounter + 2) = ResultsArray(0, DateCounter)(0)
        
        ' Insert field values
        For StockCounter = LBound(ResultsArray, 1) To UBound(ResultsArray, 1)
            Let TimeSeriesMatrix(StockCounter + 2, DateCounter + 2) = _
                ResultsArray(StockCounter, DateCounter)(1)
        Next
    Next
    
    Let GetHistorialBloombergDataData = TimeSeriesMatrix
End Function

Public Function GetReferenceBloombergData(securities() As String, _
                                          Fields() As String, _
                                          Optional OverrideFields As Variant, _
                                          Optional OverrideValues As Variant) As Variant
    Dim b As New BCOM_wrapper
    Dim r As Variant
    Dim of() As String
    Dim ov() As String
    Dim c As Integer

    If IsMissing(OverrideFields) Then
        Let r = Prepend(b.referenceData(securities, Fields), Fields)
    Else
        Let of = OverrideFields
        Let ov = OverrideValues
    
        Let r = Prepend(b.referenceData(securities, Fields, of, ov), Fields)
    End If

    Set b = Nothing
  
    Let GetReferenceBloombergData = Transpose(Prepend(Transpose(r), Prepend(securities, Empty)))
End Function

