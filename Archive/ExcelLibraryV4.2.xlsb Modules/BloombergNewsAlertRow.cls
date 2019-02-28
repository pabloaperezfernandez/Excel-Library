VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BloombergNewsAlertRow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Option Base 1

Public RangeDb As Range

Property Get GetRange() As Range
    Set GetRange = RangeDb
End Property

Public Function GetHeaders() As Variant
    Let GetHeaders = Array("Date", "BB Ticker", "Body", "EqyFundTicker", "Alert Type")
End Function

Public Function GetNumberOfColumns() As Integer
    Let GetNumberOfColumns = GetArrayLength(Me.GetHeaders)
End Function

Public Sub InitializeWithRange(ARange As Range)
    Set Me.RangeDb = ARange.Resize(1, Me.GetNumberOfColumns)
End Sub

Property Get GetDate() As Range
    Set GetDate = Me.GetRange.Cells(1)
End Property

Property Get GetBloombergTicker() As Range
    Set GetBloombergTicker = Me.GetRange.Cells(2)
End Property

Property Get GetBody() As Range
    Set GetBody = Me.GetRange.Cells(3)
End Property

Property Get GetEqyFundTicker() As Range
    Set GetEqyFundTicker = Me.GetRange.Cells(4)
End Property

Property Get GetAlertType() As Range
    Set GetAlertType = Me.GetRange.Cells(5)
End Property

' --------------------------------------------------------------------------------------------------
' Setters
' --------------------------------------------------------------------------------------------------
Property Let SetDate(Arg As String)
    Let Me.GetDate.Value2 = Arg
End Property

Property Let SetBloombergTicker(Arg As String)
    Let Me.GetBloombergTicker.Value2 = Arg
End Property

Property Let SetBody(Arg As String)
    Let Me.GetBody.Value2 = Arg
End Property

Property Let SetEqyFundTicker(Arg As String)
    Let Me.GetEqyFundTicker.Value2 = Arg
End Property

Property Let SetAlertType(Arg As String)
    Let Me.GetAlertType.Value2 = Arg
End Property

