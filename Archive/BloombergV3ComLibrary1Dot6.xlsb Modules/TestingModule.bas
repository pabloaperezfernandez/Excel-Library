Attribute VB_Name = "TestingModule"
Option Explicit
Option Base 1

' Terminal defaults are for adjusting
Public Sub HistoricalDataTestWithAdjustmentsExample()
    Dim s(0 To 1) As String
    Dim f(0 To 1) As String
    
    Let s(0) = "GE US Equity"
    Let s(1) = "IBM US Equity"
    Let f(0) = "PX_LAST"
    Let f(1) = "PX_VOLUME"
    
    Call DumpInSheet(GetHistorialBloombergData(Securities:=s, _
                                               Fields:=f, _
                                               StartDate:=CDate("01/01/2000"), _
                                               EndDate:=CDate("12/05/2019"), _
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
    
    Call DumpInSheet(GetHistorialBloombergData(Securities:=s, _
                                               Fields:=f, _
                                               StartDate:=CDate("01/01/2000"), _
                                               EndDate:=CDate("12/05/2019"), _
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
    Dim Securities(0 To 9) As String
    Dim Fields(0 To 0) As String
    Dim var As Variant
    Dim i As Integer
    
    Let Fields(0) = "PX_LAST"

    Let i = 0
    For Each var In Array("XLY US Equity", _
                          "XLV US Equity", _
                          "XLU US Equity", _
                          "XLP US Equity", _
                          "XLK US Equity", _
                          "XLI US Equity", _
                          "XLF US Equity", _
                          "XLE US Equity", _
                          "XLC US Equity", _
                          "XLB US Equity")
        Let Securities(i) = CStr(var)
        Let i = i + 1
    Next
    
    Call DumpInSheet(GetHistorialBloombergData(Securities:=Securities, _
                                               Fields:=Fields, _
                                               StartDate:=#1/1/2018#, _
                                               EndDate:=#12/5/2019# _
                                              ), _
                     Sheet1.Range("A1") _
                     )
End Sub

Public Sub HistoricalEtfData2()
    Dim Securities() As String
    Dim Fields() As String
    Dim var As Variant
    Dim i As Integer
    
    Call Sheet1.UsedRange.ClearContents
    
    Let Fields = ToStrings(Array("PX_LAST"))
    Let Securities = ToStrings(Array("XLY US Equity", _
                                     "XLV US Equity", _
                                     "XLU US Equity", _
                                     "XLP US Equity", _
                                     "XLK US Equity", _
                                     "XLI US Equity", _
                                     "XLF US Equity", _
                                     "XLE US Equity"))
    
    Call DumpInSheet(GetHistorialBloombergData(Securities:=Securities, _
                                               Fields:=Fields, _
                                               StartDate:=#1/1/2018#, _
                                               EndDate:=#12/5/2019# _
                                              ), _
                     Sheet1.Range("A1") _
                     )
End Sub

Public Sub HistoricalEtfData3()
    Dim Securities() As String
    Dim Fields() As String
    Dim var As Variant
    Dim i As Integer
    Dim DictOfTimeSeries As Dictionary
    
    Call Sheet1.UsedRange.ClearContents
    
    Let Fields = ToStrings(Array("PX_LAST", "PX_VOLUME"))
    Let Securities = ToStrings(Array("XLY US Equity", _
                                     "XLV US Equity", _
                                     "XLU US Equity", _
                                     "XLP US Equity", _
                                     "XLK US Equity", _
                                     "XLI US Equity", _
                                     "XLF US Equity", _
                                     "XLE US Equity", _
                                     "XLC US Equity", _
                                     "XLB US Equity"))

    Set DictOfTimeSeries = GetHistorialBloombergData(Securities:=Securities, _
                                                     Fields:=Fields, _
                                                     StartDate:=#1/1/2018#, _
                                                     EndDate:=#12/5/2019# _
                                                    )
    Call DumpInSheet(DictOfTimeSeries.Item("PX_LAST"), Sheet1.Range("A1"))
    Call DumpInSheet(DictOfTimeSeries.Item("PX_VOLUME"), Sheet1.Range("A15"))
End Sub

Public Sub HistoricalEtfData4()
    Dim Securities() As String
    Dim Fields() As String
    Dim var As Variant
    Dim i As Integer
    Dim DictOfTimeSeries As Dictionary
    
    Call Sheet1.UsedRange.ClearContents
    
    Let Fields = ToStrings(Array("PX_LAST", "PX_VOLUME"))
    Let Securities = ToStrings(Array("XLY US Equity", _
                                     "XLV US Equity", _
                                     "XLU US Equity", _
                                     "XLP US Equity", _
                                     "XLK US Equity", _
                                     "XLI US Equity", _
                                     "XLF US Equity", _
                                     "XLE US Equity", _
                                     "XLC US Equity", _
                                     "XLB US Equity"))

    Set DictOfTimeSeries = GetHistorialBloombergData(Securities:=Securities, _
                                                     Fields:=Fields, _
                                                     StartDate:=#1/31/2018#, _
                                                     EndDate:=#5/31/2019#, _
                                                     periodicityAdjustment:=ACTUAL, _
                                                     periodicitySelection:=MONTHLY _
                                                    )
    Call DumpInSheet(DictOfTimeSeries.Item("PX_LAST"), Sheet1.Range("A1"))
    Call DumpInSheet(DictOfTimeSeries.Item("PX_VOLUME"), Sheet1.Range("A15"))
End Sub
