VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MasterFileRow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Base 1
Option Explicit

' Private instance variable to hold the trade database
Private RangeDb As Range

' Custom initializer. aRange must be a row range.  All columns to the right of number 11 are ignored.
Public Sub InitializeWithRange(ARange As Range)
    Set RangeDb = ARange.Resize(1, 48)
End Sub

' Getters
' This returns the range holding the asset's value. This could be useful when trying to get the address
' of the range.
Property Get GetRange() As Range
    Set GetRange = RangeDb
End Property

Property Get GetIsin() As Range
    Set GetIsin = GetRange.Cells(1)
End Property

Property Get Get6DigitSedol() As Range
    Set Get6DigitSedol = GetRange.Cells(2)
End Property

Property Get Get7DigitSedol() As Range
    Set Get7DigitSedol = GetRange.Cells(3)
End Property

Property Get GetParsekeableDesSource() As Range
    Set GetParsekeableDesSource = GetRange.Cells(4)
End Property

Property Get GetBloombergTicker() As Range
    Set GetBloombergTicker = GetRange.Cells(5)
End Property

Property Get GetExchCode() As Range
    Set GetExchCode = GetRange.Cells(6)
End Property

Property Get GetCrncy() As Range
    Set GetCrncy = GetRange.Cells(7)
End Property

Property Get GetCntryOfDomicile() As Range
    Set GetCntryOfDomicile = GetRange.Cells(8)
End Property

Property Get GetGicsSubIndustry() As Range
    Set GetGicsSubIndustry = GetRange.Cells(9)
End Property

Property Get GetMsciIndustryCode() As Range
    Set GetMsciIndustryCode = GetRange.Cells(10)
End Property

Property Get GetPalladyneSectorCode() As Range
    Set GetPalladyneSectorCode = GetRange.Cells(11)
End Property

Property Get GetRegion() As Range
    Set GetRegion = GetRange.Cells(12)
End Property

Property Get GetName() As Range
    Set GetName = GetRange.Cells(13)
End Property

Property Get GetUsdMarketValue() As Range
    Set GetUsdMarketValue = GetRange.Cells(14)
End Property

Property Get GetMarketStatus() As Range
    Set GetMarketStatus = GetRange.Cells(15)
End Property

Property Get GetSecurityType() As Range
    Set GetSecurityType = GetRange.Cells(16)
End Property

Property Get GetCountryTradableFlag() As Range
    Set GetCountryTradableFlag = GetRange.Cells(17)
End Property

Property Get GetPiamTradableCountryName() As Range
    Set GetPiamTradableCountryName = GetRange.Cells(18)
End Property

Property Get Get30DayAverageTradingVolume() As Range
    Set Get30DayAverageTradingVolume = GetRange.Cells(19)
End Property

Property Get Get30DayAverageTradingValue() As Range
    Set Get30DayAverageTradingValue = GetRange.Cells(20)
End Property

Property Get GetCountryFullName() As Range
    Set GetCountryFullName = GetRange.Cells(21)
End Property

Property Get GetPxTradeLotSize() As Range
    Set GetTradeLotSize = GetRange.Cells(22)
End Property

Property Get GetPxRoundLotSize() As Range
    Set GetTradeLotSize = GetRange.Cells(23)
End Property

Property Get GetCountryRiskIsoCode() As Range
    Set GetCountryRiskIsoCode = GetRange.Cells(24)
End Property

' Returns one of the numerical codes below indicating Bloomberg's classification for the secuirity
' COMMON STOCK==0;  ADR/ GDR/ ADS==1; PRFD STOCK==2; CLOSED-END FUND==3; CONV PRFD==4; REIT==5; MISC==6
Property Get GetAdrGdrAdsPrfdClosedConvReitMiscFlag() As Range
    Set GetAdrGdrAdsPrfdClosedConvReitMiscFlag = GetRange.Cells(26)
End Property

Property Get GetGet30DayAverageTradingVolume() As Range
    Set GetGet30DayAverageTradingVolume = GetRange.Cells(32)
End Property

Property Get GetSecurityTyp2() As Range
    Set GetSecurityTyp2 = GetRange.Cells(33)
End Property

Property Get GetCommonSharesOutstandingPerFactSet() As Range
    Set GetCommonSharesOutstandingPerFactSet = GetRange.Cells(34)
End Property

Property Get GetCieDes() As Range
    Set GetCieDes = GetRange.Cells(36)
End Property

Property Get GetEqyFundTicker() As Range
    Set GetEqyFundTicker = GetRange.Cells(37)
End Property

Property Get GetWacc() As Range
    Set GetWacc = GetRange.Cells(38)
End Property

Property Get GetUsdTradeLotValue() As Range
    Set GetUsdTradeLotValue = GetRange.Cells(39)
End Property

Property Get GetPriceTimes1MAverageTradingVolume() As Range
    Set GetUsdAverageTradingValue = GetRange.Cells(41)
End Property

Property Get GetAboveMinimumUsdAverageTradingValueFlag() As Range
    Set GetUsdAverageTradingValue = GetRange.Cells(42)
End Property

Property Get GetBloombergTickerPlusAboveMinimumUsdAverageTradingValueFlag() As Range
    Set GetBloombergTickerPlusAboveMinimumUsdAverageTradingValue = GetRange.Cells(43)
End Property

Property Get GetTradableAlternativeBloombergTicker() As Range
    Set GetTradableAlternativeBloombergTicker = GetRange.Cells(45)
End Property

Property Get GetTradableAlternative6DigitSedol() As Range
    Set GetTradableAlternative6DigitSedol = GetRange.Cells(46)
End Property

Property Get GetPalladyneSector() As Range
    Set GetPalladyneSector = GetRange.Cells(47)
End Property

Property Get GetInvestibleSedol() As Range
    Set GetPalladyneSector = GetRange.Cells(48)
End Property

Property Get GetInvestibleBloombergTicker() As Range
    Set GetPalladyneSector = GetRange.Cells(49)
End Property

' Setters

Property Let SetIsin(arg As String)
    Let GetIsins.Value2 = arg
End Property

Property Let Set6DigitSedol(arg As String)
    Let Get6DigitSedol.Value2 = arg
End Property

Property Let Set7DigitSedol(arg As String)
    Let Get7DigitSedol.Value2 = arg
End Property

Property Let SetParsekeableDesSource(arg As String)
    Let GetParsekeableDesSource.Value2 = arg
End Property

Property Let SetBloombergTicker(arg As String)
    Let GetBloombergTicker.Value2 = arg
End Property

Property Let SetExchCode(arg As String)
    Let GetExchCode.Value2 = arg
End Property

Property Let SetCrncy(arg As String)
    Let GetCrncy.Value2 = arg
End Property

Property Let SetCntryOfDomicile(arg As String)
    Let GetCntryOfDomicile.Value2 = arg
End Property

Property Let SetGicsSubIndustry(arg As String)
    Let GetGicsSubIndustry.Value2 = arg
End Property

Property Let SetMsciIndustryCode(arg As String)
    Let GetMsciIndustryCode.Value2 = arg
End Property

Property Let SetPalladyneSectorCode(arg As String)
    Let GetPalladyneSectorCode.Value2 = arg
End Property

Property Let SetRegion(arg As String)
    Let GetRegion.Value2 = arg
End Property

Property Let SetName(arg As String)
    Let GetName.Value2 = arg
End Property

Property Let SetUsdMarketValue(arg As Double)
    Let GetUsdMarketValue.Value2 = arg
End Property

Property Let SetMarketStatus(arg As String)
    Let GetMarketStatus.Value2 = arg
End Property

Property Let SetSecurityType(arg As String)
    Let GetSecurityType.Value2 = arg
End Property

Property Let SetCountryTradableFlag(arg As String)
    Let GetCountryTradableFlag.Value2 = arg
End Property

Property Let SetPiamTradableCountryName(arg As String)
    Let GetPiamTradableCountryName.Value2 = arg
End Property

Property Let Set30DayAverageTradingVolume(arg As Long)
    Let Get30DayAverageTradingVolume.Value2 = arg
End Property

Property Let Set30DayAverageTradingValue(arg As Double)
    Let Get30DayAverageTradingValue.Value2 = arg
End Property

Property Let SetCountryFullName(arg As String)
    Let GetCountryFullName.Value2 = arg
End Property

Property Let SetTradeLotSize(arg As Long)
    Let GetTradeLotSize.Value2 = arg
End Property

Property Let SetCountryRiskIsoCode(arg As String)
    Let GetCountryRiskIsoCode.Value2 = arg
End Property

Property Let SetName2(arg As String)
    Let GetName.Value2 = arg
End Property

Property Let SetAdrGdrAdsPrfdClosedConvReitMiscFlag(arg As Integer)
    Let GetAdrGdrAdsPrfdClosedConvReitMiscFlag.Value2 = arg
End Property

Property Let SetGet30DayAverageTradingVolume(arg As Long)
    Let GetGet30DayAverageTradingVolume.Value2 = arg
End Property

Property Let SetSecurityTyp2(arg As String)
    Let GetSecurityTyp2.Value2 = arg
End Property

Property Let SetCieDes(arg As String)
    Let GetCieDes.Value2 = arg
End Property

Property Let SetEqyFundTicker(arg As String)
    Let GetEqyFundTicker.Value2 = arg
End Property

Property Let SetWacc(arg As Double)
    Let GetWacc.Value2 = arg
End Property

Property Let SetUsdTradeLotValue(arg As Double)
    Let GetUsdTradeLotValue.Value2 = arg
End Property

Property Let SetPriceTimes1MAverageTradingVolume(arg As Double)
    Let GetUsdAverageTradingValue.Value2 = arg
End Property

Property Let SetAboveMinimumUsdAverageTradingValueFlag(arg As String)
    Let GetUsdAverageTradingValue.Value2 = arg
End Property

Property Let SetBloombergTickerPlusAboveMinimumUsdAverageTradingValueFlag(arg As String)
    Let GetBloombergTickerPlusAboveMinimumUsdAverageTradingValue.Value2 = arg
End Property

Property Let SetTradableAlternativeBloombergTicker(arg As String)
    Let GetTradableAlternativeBloombergTicker.Value2 = arg
End Property

Property Let SetTradableAlternative6DigitSedol(arg As String)
    Let GetTradableAlternative6DigitSedol.Value2 = arg
End Property

Property Let SetPalladyneSector(arg As String)
    Let GetPalladyneSector.Value2 = arg
End Property

Property Let SetInvestibleSedol(arg As String)
    Let GetPalladyneSector.Value2 = arg
End Property

Property Let SetInvestibleBloombergTicker(arg As String)
    Let GetPalladyneSector.Value2 = arg
End Property