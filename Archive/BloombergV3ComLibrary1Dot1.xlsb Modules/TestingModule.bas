Attribute VB_Name = "TestingModule"
Option Explicit

' Terminal defaults are for adjusting
Public Sub HistoricalDataTestWithAdjustmentsExample()
    Dim s(0 To 1) As String
    Dim f(0 To 1) As String
    
    Let s(0) = "GE US Equity"
    Let s(1) = "IBM US Equity"
    Let f(0) = "PX_LAST"
    Let f(1) = "PX_VOLUME"
    
    Call DumpInSheet(GetHistorialBloombergDataData(securities:=s, _
                                                   Fields:=f, _
                                                   startDate:=CDate("01/01/2000"), _
                                                   endDate:=CDate("12/05/2019"), _
                                                   adjustmentFollowDPDF:=False, _
                                                   adjustmentNormal:=False, _
                                                   adjustmentAbnormal:=False, _
                                                   adjustmentSplit:=False _
                                                  ), _
                     Sheet1.Range("A1") _
                     )
    Let Sheet1.Range("A1").Value2 = "NoAdjustments"
End Sub

' Terminal defaults are for adjusting
Public Sub HistoricalDataTestWithNoAdjustmentsExample()
    Dim s(0 To 1) As String
    Dim f(0 To 0) As String
    
    Let s(0) = "GE US Equity"
    Let s(1) = "IBM US Equity"
    Let f(0) = "PX_LAST"
    
    Call DumpInSheet(GetHistorialBloombergDataData(securities:=s, _
                                                   Fields:=f, _
                                                   startDate:=CDate("01/01/2000"), _
                                                   endDate:=CDate("12/05/2019"), _
                                                   adjustmentFollowDPDF:=False, _
                                                   adjustmentNormal:=True, _
                                                   adjustmentAbnormal:=True, _
                                                   adjustmentSplit:=True _
                                                  ), _
                     Sheet1.Range("A4") _
                     )
    Let Sheet1.Range("A4").Value2 = "Adjustments"
End Sub

Public Sub HistoricalEtfData()
    Dim s() As String
    Dim f(0 To 0) As String
    Dim var As Variant
    Dim i As Integer
    Dim ResultsArray As Variant
    Dim DumpArray As Variant

    Let f(0) = "PX_LAST"
    ReDim s(0 To 1)
    Let i = 0
    For Each var In Array("XLY US Equity", _
                          "XLV US Equity")
        Let s(i) = CStr(var)
        Let i = i + 1
    Next
    
    Let ResultsArray = GetHistorialBloombergDataData(securities:=s, _
                                                     Fields:=f, _
                                                     startDate:=CDate("01/01/2000"), _
                                                     endDate:=CDate("12/05/2019") _
                                                  )
    
    Call DumpInSheet(ResultsArray, _
                     ThisWorkbook.Worksheets("Sheet1").Range("A1") _
                     )
End Sub

