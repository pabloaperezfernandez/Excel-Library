VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PrivateFileRow"
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
    Let GetNumberOfColumns = GetArrayLength(Me.GetHeaders)
End Property

Property Get GetHeaders() As Variant
    Let GetHeaders = Array("Date", "Country of Domicile", "Rank", "USD Market Cap (mn)", "Company Name", "PIAM Sector", "FY1 P/E", "Industry Code", "Net Alpha", "Beta", "Long-term Forward Growth", "Dividend Yield", "Shares Outstanding", "ISIN", "Avg One-month Trading Volume", "USD Trade Lot Size", "Tradable Country Flag", "Bloomberg Ticker", "Investable SEDOL", "Investable Bloomberg Ticker", "Investable ISIN", "7-char SEDOL", "Primary Listing's Bloomberg Ticker", "Classification", "Region", "Size", "Style")
End Property

' This returns the range holding the asset's value. This could be useful when trying to get the address
' of the range.
Property Get GetRange() As Range
    Set GetRange = RangeDb
End Property

Property Get GetDate() As Range
    Set GetDate = Me.GetRange.Cells(1)
End Property

Property Get GetCountryOfDomicile() As Range
    Set GetCountryOfDomicile = Me.GetRange.Cells(2)
End Property

Property Get GetRank() As Range
    Set GetRank = Me.GetRange.Cells(3)
End Property

Property Get GetUsdMarketCap() As Range
    Set GetUsdMarketCap = Me.GetRange.Cells(4)
End Property

Property Get GetName() As Range
    Set GetName = Me.GetRange.Cells(5)
End Property

Property Get GetPiamSector() As Range
    Set GetPiamSector = Me.GetRange.Cells(6)
End Property

Property Get GetFy1Pe() As Range
    Set GetFy1Pe = Me.GetRange.Cells(7)
End Property

Property Get GetIndustryCode() As Range
    Set GetIndustryCode = Me.GetRange.Cells(8)
End Property

Property Get GetNetAlpha() As Range
    Set GetNetAlpha = Me.GetRange.Cells(9)
End Property

Property Get GetBeta() As Range
    Set GetBeta = Me.GetRange.Cells(10)
End Property

Property Get GetLongTermForwardGrowth() As Range
    Set GetLongTermForwardGrowth = Me.GetRange.Cells(11)
End Property

Property Get GetDividendYield() As Range
    Set GetDividendYield = Me.GetRange.Cells(12)
End Property

Property Get GetSharesOutstanding() As Range
    Set GetSharesOutstanding = Me.GetRange.Cells(13)
End Property

Property Get GetIsin() As Range
    Set GetIsin = Me.GetRange.Cells(14)
End Property

Property Get GetAverageOneMonthTradingVolume() As Range
    Set GetAverageOneMonthTradingVolume = Me.GetRange.Cells(15)
End Property

Property Get GetUsdTradeLotSize() As Range
    Set GetUsdTradeLotSize = Me.GetRange.Cells(16)
End Property

Property Get GetTradableCountryFlag() As Range
    Set GetTradableCountryFlag = Me.GetRange.Cells(17)
End Property

Property Get GetBloombergTicker() As Range
    Set GetBloombergTicker = Me.GetRange.Cells(18)
End Property

Property Get GetInvestableSedol() As Range
    Set GetInvestableSedol = Me.GetRange.Cells(19)
End Property

Property Get GetInvestableBloombergTicker() As Range
    Set GetInvestableBloombergTicker = Me.GetRange.Cells(20)
End Property

Property Get GetInvestableIsin() As Range
    Set GetInvestableIsin = Me.GetRange.Cells(21)
End Property

Property Get GetSedol() As Range
    Set GetSedol = Me.GetRange.Cells(22)
End Property

Property Get GetPrimaryBloombergTicker() As Range
    Set GetPrimaryBloombergTicker = Me.GetRange.Cells(23)
End Property

Property Get GetClassification() As Range
    Set GetClassification = Me.GetRange.Cells(24)
End Property

Property Get GetRegion() As Range
    Set GetRegion = GetRange.Cells(25)
End Property

Property Get GetSize() As Range
    Set GetSize = GetRange.Cells(26)
End Property

Property Get GetStyle() As Range
    Set GetStyle = GetRange.Cells(27)
End Property

' Setters
Property Let SetDate(arg As Long)
    Let Me.GetDate.Value2 = arg
End Property

Property Let SetCountryOfDomicile(arg As String)
    Let Me.GetCountryOfDomicile.Value2 = arg
End Property

Property Let SetRank(arg As Integer)
    Let Me.GetRank.Value2 = arg
End Property

Property Let SetUsdMarketCap(arg As Double)
    Let Me.GetUsdMarketCap.Value2 = arg
End Property

Property Let SetName(arg As String)
    Let Me.GetName.Value2 = arg
End Property

Property Let SetPiamSector(arg As String)
    Let Me.GetPiamSector.Value2 = arg
End Property

Property Let SetFy1Pe(arg As Double)
    Let Me.GetFy1Pe.Value2 = arg
End Property

Property Let SetIndustryCode(arg As Long)
    Let Me.GetIndustryCode.Value2 = arg
End Property

Property Let SetNetAlpha(arg As Double)
    Let Me.GetNetAlpha.Value2 = arg
End Property

Property Let SetBeta(arg As Double)
    Let Me.GetBeta.Value2 = arg
End Property

Property Let SetLongTermForwardGrowth(arg As Double)
    Let Me.GetLongTermForwardGrowth.Value2 = arg
End Property

Property Let SetDividendYield(arg As Double)
    Let Me.GetDividendYield.Value2 = arg
End Property

Property Let SetSharesOutstanding(arg As Double)
    Let Me.GetSharesOutstanding.Value2 = arg
End Property

Property Let SetIsin(arg As String)
    Let Me.GetIsin.Value2 = arg
End Property

Property Let SetAverageOneMonthTradingVolume(arg As Double)
    Let Me.GetAverageOneMonthTradingVolume.Value2 = arg
End Property

Property Let SetUsdTradeLotSize(arg As Double)
    Let Me.GetUsdTradeLotSize.Value2 = arg
End Property

Property Let SetTradableCountryFlag(arg As Boolean)
    Let Me.GetTradableCountryFlag.Value2 = arg
End Property

Property Let SetBloombergTicker(arg As String)
    Let Me.GetBloombergTicker.Value2 = arg
End Property

Property Let SetInvestableSedol(arg As String)
    Let Me.GetInvestableSedol.Value2 = arg
End Property

Property Let SetInvestableBloombergTicker(arg As String)
    Let Me.GetInvestableBloombergTicker.Value2 = arg
End Property

Property Let SetInvestableIsin(arg As String)
    Let Me.GetInvestableIsin.Value2 = arg
End Property

Property Let SetSedol(arg As String)
    Let Me.GetSedol.Value2 = arg
End Property

Property Let SetPrimaryBloombergTicker(arg As String)
    Let Me.GetPrimaryBloombergTicker.Value2 = arg
End Property

Property Let SetClassification(arg As String)
    Let Me.GetClassification.Value2 = arg
End Property

Property Let SetRegion(arg As String)
    Let Me.GetRegion.Value2 = arg
End Property

Property Let SetSize(arg As String)
    Let Me.GetSize.Value2 = arg
End Property

Property Let SetStyle(arg As String)
    Let Me.GetStyle.Value2 = arg
End Property
