VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EmsxTrade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Base 1
Option Explicit

' Private instance variable to hold the trade database
Private RangeDb As Range

' Custom initializer. The first row is ignored because it is expected to be the
' header row
Public Sub InitializeWithRange(ARange As Range)
    Set RangeDb = ARange.Resize(1, Me.GetNumberOfColumns)
End Sub

' Getters
' This returns the range holding the asset's value. This could be useful when trying to get the address
' of the range.
Property Get GetRange() As Range
    Set GetRange = RangeDb
End Property

Property Get GetHeaders() As Variant
    Let GetHeaders = Array("Date", "Time", "AssetClassCode", "StrategyCode", "Custodian", "FundId", "BloombergTicker", _
                           "BloombergExchangeCode", "CurrencyCode", "OrderAsString", "Side", "Units", "Limit", "Tif", _
                           "Account", "OptimalDateStamp", "OptimalTimeStamp", "HoldingsDateStamp", "HoldingsTimeStamp", _
                           "AllocationsDateStamp", "AllocationsTimeStamp", "RunNumber")
End Property

Property Get GetNumberOfColumns() As Integer
    Let GetNumberOfColumns = GetArrayLength(Me.GetHeaders)
End Property

Property Get GetDate() As Range
    Set GetDate = GetRange.Cells(1)
End Property

Property Get GetTime() As Range
    Set GetTime = GetRange.Cells(2)
End Property

Property Get GetAssetClassCode() As Range
    Set GetAssetClassCode = GetRange.Cells(3)
End Property

Property Get GetStrategyCode() As Range
    Set GetStrategyCode = GetRange.Cells(4)
End Property

Property Get GetCustodian() As Range
    Set GetCustodian = GetRange.Cells(5)
End Property

Property Get GetFundId() As Range
    Set GetFundId = GetRange.Cells(6)
End Property

' This ticker includes the exchange code, but it does not include the "Equity" qualifier
Property Get GetBloombergTicker() As Range
    Set GetBloombergTicker = GetRange.Cells(7)
End Property

Property Get GetBloombergExchangeCode() As Range
    Set GetBloombergExchangeCode = GetRange.Cells(8)
End Property

Property Get GetCurrencyCode() As Range
    Set GetCurrencyCode = GetRange.Cells(9)
End Property

Property Get GetOrder() As Range
    Set GetOrder = GetRange.Cells(10)
End Property

Property Get GetSide() As Range
    Set GetSide = GetRange.Cells(11)
End Property

Property Get GetUnits() As Range
    Set GetUnits = GetRange.Cells(12)
End Property

Property Get GetLimit() As Range
    Set GetLimit = GetRange.Cells(13)
End Property

Property Get GetTif() As Range
    Set GetTif = GetRange.Cells(14)
End Property

Property Get GetAccount() As Range
    Set GetAccount = GetRange.Cells(15)
End Property

Property Get GetOptimalDateStamp() As Range
    Set GetOptimalDateStamp = GetRange.Cells(16)
End Property

Property Get GetOptimalTimeStamp() As Range
    Set GetOptimalTimeStamp = GetRange.Cells(17)
End Property

Property Get GetHoldingsDateStamp() As Range
    Set GetHoldingsDateStamp = GetRange.Cells(18)
End Property

Property Get GetHoldingsTimeStamp() As Range
    Set GetHoldingsTimeStamp = GetRange.Cells(19)
End Property

Property Get GetAllocationsDateStamp() As Range
    Set GetAllocationsDateStamp = GetRange.Cells(20)
End Property

Property Get GetAllocationsTimeStamp() As Range
    Set GetAllocationsTimeStamp = GetRange.Cells(21)
End Property

Property Get GetRunNumber() As Range
    Set GetRunNumber = GetRange.Cells(22)
End Property

' Setters
' This ticker includes the exchange code, but it does not include the "Equity" qualifier
Property Let SetAssetClassCode(Arg As String)
    Let Me.GetAssetClassCode.Value2 = Arg
End Property

Property Let SetStrategyCode(Arg As String)
    Let Me.GetStrategyCode.Value2 = Arg
End Property

Property Let SetCustodian(Arg As String)
    Let Me.GetCustodian.Value2 = Arg
End Property

Property Let SetFundId(Arg As String)
    Let Me.GetFundId.Value2 = Arg
End Property

Property Let SetDate(Arg As Long)
    Let Me.GetDate.Value2 = Arg
End Property

Property Let SetTime(Arg As Long)
    Let Me.GetTime.Value2 = Arg
End Property

Property Let SetBloombergTicker(Arg As String)
    Let Me.GetBloombergTicker.Value2 = Arg
End Property

Property Let SetBloombergExchangeCode(Arg As String)
    Let Me.GetBloombergExchangeCode.Value2 = Arg
End Property

Property Let SetCurrencyCode(Arg As String)
    Let Me.GetCurrencyCode.Value2 = Arg
End Property

Property Let SetOrder(Arg As String)
    Let Me.GetOrder.Value2 = Arg
End Property

Property Let SetSide(Arg As String)
    Let Me.GetSide.Value2 = Arg
End Property

Property Let SetUnits(Arg As Double)
    Let Me.GetUnits.Value2 = Arg
End Property

Property Let SetLimit(Arg As String)
    Let Me.GetLimit.Value2 = Arg
End Property

Property Let SetTif(Arg As String)
    Let Me.GetTif.Value2 = Arg
End Property

Property Let SetAccount(Arg As String)
    Let Me.GetAccount.Value2 = Arg
End Property

Property Let SetOptimalDateStamp(Arg As Long)
    Let Me.GetOptimalDateStamp = Arg
End Property

Property Let SetOptimalTimeStamp(Arg As Long)
    Let Me.GetOptimalTimeStamp = Arg
End Property

Property Let SetHoldingsDateStamp(Arg As Long)
    Let Me.GetHoldingsDateStamp = Arg
End Property

Property Let SetHoldingsTimeStamp(Arg As Long)
    Let Me.GetHoldingsTimeStamp = Arg
End Property

Property Let SetAllocationsDateStamp(Arg As Long)
    Let Me.GetAllocationsDateStamp = Arg
End Property

Property Let SetAllocationsTimeStamp(Arg As Long)
    Let Me.GetAllocationsTimeStamp = Arg
End Property

Property Let SetRunNumber(Arg As Integer)
    Let Me.GetRunNumber = Arg
End Property