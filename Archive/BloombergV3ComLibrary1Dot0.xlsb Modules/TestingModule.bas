Attribute VB_Name = "TestingModule"
Option Explicit

' Terminal defaults are for adjusting
Public Sub HistoricalDataTestWithAdjustmentsExample()
    Dim s(0 To 0) As String
    Dim of(0 To 2) As String
    Dim ov(0 To 2) As String
    
    Let s(0) = "GE US Equity"
    Let of(0) = "adjustmentAbnormal": of(1) = "adjustmentNormal": of(2) = "adjustmentFollowDPDF"
    Let ov(0) = "TRUE": ov(1) = "TRUE": ov(2) = "FALSE"
    
    Call DumpInSheet(GetHistorialBloombergDataData(securities:=s, _
                                                   Field:="PX_LAST", _
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
    Dim s(0 To 0) As String
    Dim of(0 To 2) As String
    Dim ov(0 To 2) As String
    
    Let s(0) = "GE US Equity"
    Let of(0) = "adjustmentAbnormal": of(1) = "adjustmentNormal": of(2) = "adjustmentFollowDPDF"
    Let ov(0) = "TRUE": ov(1) = "TRUE": ov(2) = "FALSE"
    
    Call DumpInSheet(GetHistorialBloombergDataData(securities:=s, _
                                                   Field:="PX_LAST", _
                                                   startDate:=CDate("01/01/2000"), _
                                                   endDate:=CDate("12/05/2019"), _
                                                   adjustmentFollowDPDF:=False, _
                                                   adjustmentNormal:=True, _
                                                   adjustmentAbnormal:=True, _
                                                   adjustmentSplit:=True _
                                                  ), _
                     Sheet1.Range("A4") _
                     )
    Let Sheet1.Range("A1").Value2 = "Adjustments"
End Sub
