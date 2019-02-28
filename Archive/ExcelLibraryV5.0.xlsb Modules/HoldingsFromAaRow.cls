VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HoldingsFromAaRow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Base 1
Option Explicit

' Private instance variable to hold the trade database
Private RangeDb As Range

' Initializer
Public Sub InitializeWithRange(ARange As Range)
    Set RangeDb = ARange.Resize(1, Me.GetNumberOfColumns)
End Sub

' Getters
Property Get GetNumberOfColumns() As Integer
    Let GetNumberOfColumns = GetArrayLength(GetHeaders())
End Property

' This returns the range holding the asset's value. This could be useful when trying to get the address
' of the range.
Property Get GetRange() As Range
    Set GetRange = RangeDb
End Property

Property Get GetDate() As Range
    Set GetDate = Me.GetRange.Cells(1)
End Property

Property Get GetTime() As Range
    Set GetTime = Me.GetRange.Cells(2)
End Property

Property Get GetAssetClassCode() As Range
    Set GetAssetClassCode = Me.GetRange.Cells(3)
End Property

Property Get GetStrategyCode() As Range
    Set GetStrategyCode = Me.GetRange.Cells(4)
End Property

Property Get GetFundId() As Range
    Set GetFundId = Me.GetRange.Cells(5)
End Property

Property Get Get7CharSedol() As Range
    Set Get7CharSedol = Me.GetRange.Cells(6)
End Property

Property Get GetIsin() As Range
    Set GetIsin = Me.GetRange.Cells(7)
End Property

Property Get GetBloombergTicker() As Range
    Set GetBloombergTicker = Me.GetRange.Cells(8)
End Property

Property Get GetSecurityName() As Range
    Set GetSecurityName = Me.GetRange.Cells(9)
End Property

Property Get GetAmount() As Range
    Set GetAmount = Me.GetRange.Cells(10)
End Property

Property Get GetUsdPositionValue() As Range
    Set GetUsdPositionValue = Me.GetRange.Cells(11)
End Property

Property Get GetCustodian() As Range
    Set GetCustodian = Me.GetRange.Cells(12)
End Property

Property Get GetUnavailabilityFlag() As Range
    Set GetUnavailabilityFlag = Me.GetRange.Cells(13)
End Property

Property Get GetPortfolioDate() As Range
    Set GetPortfolioDate = Me.GetRange.Cells(14)
End Property

Property Get GetRunNumber() As Range
    Set GetRunNumber = Me.GetRange.Cells(15)
End Property

Property Get GetClassification() As Range
    Set GetClassification = Me.GetRange.Cells(16)
End Property

Property Get GetHeaders() As Variant
    Let GetHeaders = Array("Date", "Time", "Asset Class", "Strategy", "FundID", "SEDOL", "ISIN", "Bloomberg Ticker", _
                           "Security Name", "Amount", "USD Value", "Custodian", "Unavailability Flag", _
                           "Portfolio Date", "Run Number", "Classification")
End Property

' Setters
Property Let SetDate(arg As Long)
    Let Me.GetDate.Value2 = arg
End Property

Property Let SetTime(arg As Long)
    Set Me.GetTime.Value2 = arg
End Property

Property Let SetAssetClassCode(arg As String)
    Let Me.GetAssetClassCode.Value2 = arg
End Property

Property Let SetStrategyCode(arg As String)
    Let Me.GetStrategyCode.Value2 = arg
End Property

Property Let SetFundId(arg As String)
    Let Me.GetFundId.Value2 = arg
End Property

Property Let Set7CharSedol(arg As String)
    Set Me.Get7CharSedol.Value2 = arg
End Property

Property Let SetIsin(arg As String)
    Let Me.GetIsin.Value2 = arg
End Property

Property Let SetBloombergTicker(arg As String)
    Let Me.GetBloombergTicker.Value2 = arg
End Property

Property Let SetSecurityName(arg As String)
    Let Me.GetSecurityName.Value2 = arg
End Property

Property Let SetAmount(arg As Double)
    Let Me.GetAmount.Value2 = arg
End Property

Property Let SetUsdPositionValue(arg As Double)
    Let Me.GetUsdPositionValue.Value2 = arg
End Property

Property Let SetCustodian(arg As String)
    Let Me.GetCustodian.Value2 = arg
End Property

Property Let SetUnavailabilityFlag(arg As Integer)
    Let Me.GetUnavailabilityFlag.Value2 = arg
End Property

Property Let SetPortfolioDate(arg As Long)
    Let Me.GetPortfolioDate.Value2 = arg
End Property

Property Let SetRunNumber(arg As Integer)
    Let Me.GetRunNumber.Value2 = arg
End Property

Property Let SetClassification(arg As String)
    Let Me.GetClassification.Value2 = arg
End Property
