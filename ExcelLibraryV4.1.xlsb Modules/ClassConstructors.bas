Attribute VB_Name = "ClassConstructors"
Option Base 1
Option Explicit

Public Function NewCountryCodeRegionalCodeMap() As CountryCodeRegionalCodeMap
    Set NewCountryCodeRegionalCodeMap = New CountryCodeRegionalCodeMap
End Function

Public Function NewEmsxTrade()
    Set NewEmsxTrade = New EmsxTrade
End Function

Public Function NewEmsxTradeList()
    Set NewEmsxTradeList = New EmsxTradeList
End Function

Public Function NewTimeSeries() As TimeSeries
    Set NewTimeSeries = New TimeSeries
End Function

Public Function NewHoldingsFromAa() As HoldingsFromAa
    Set NewHoldingsFromAa = New HoldingsFromAa
End Function

Public Function NewHoldingsFromAaRow() As HoldingsFromAaRow
    Set NewHoldingsFromAaRow = New HoldingsFromAaRow
End Function

Public Function NewOptimalAsset() As OptimalAsset
    Set NewOptimalAsset = New OptimalAsset
End Function

Public Function NewOptimalPortfolio() As OptimalPortfolio
    Set NewOptimalPortfolio = New OptimalPortfolio
End Function

Public Function NewTargetAssetAllocations() As TargetAssetAllocations
    Set NewTargetAssetAllocations = New TargetAssetAllocations
End Function

Public Function NewTargetAssetAllocationRow() As TargetAssetAllocationRow
    Set NewTargetAssetAllocationRow = New TargetAssetAllocationRow
End Function

Public Function NewPrivateFile() As PrivateFile
    Set NewPrivateFile = New PrivateFile
End Function

Public Function NewPrivateFileRow() As PrivateFileRow
    Set NewPrivateFileRow = New PrivateFileRow
End Function

Public Function NewEquityDbHandler() As EquityDbHandler
    Set NewEquityDbHandler = New EquityDbHandler
End Function

Public Function NewEquityPlotHandler() As EquityPlotHandler
    Set NewEquityPlotHandler = New EquityPlotHandler
End Function

Public Function NewPostTradingPortfolioAsset() As PostTradingPortfolioAsset
    Set NewPostTradingPortfolioAsset = New PostTradingPortfolioAsset
End Function

Public Function NewPostTradingPortfolio() As PostTradingPortfolio
    Set NewPostTradingPortfolio = New PostTradingPortfolio
End Function

Public Function NewBloombergEquityAlerts() As BloombergEquityAlerts
    Set NewBloombergEquityAlerts = New BloombergEquityAlerts
End Function

Public Function NewBloombergEquityAlertRow() As BloombergEquityAlertRow
    Set NewBloombergEquityAlertRow = New BloombergEquityAlertRow
End Function

Public Function NewBloombergNewsAlerts() As BloombergNewsAlerts
    Set NewBloombergNewsAlerts = New BloombergNewsAlerts
End Function

Public Function NewBloombergNewsAlertRow() As BloombergNewsAlertRow
    Set NewBloombergNewsAlertRow = New BloombergNewsAlertRow
End Function

Public Function NewCorporateAction() As CorporateAction
    Set NewCorporateAction = New CorporateAction
End Function

Public Function NewSeimRecord() As SeimRecord
    Set NewSeimRecord = New SeimRecord
End Function

Public Function NewSeimRecordSet() As SeimRecordSet
    Set NewSeimRecordSet = New SeimRecordSet
End Function

Public Function NewBrokerAllocation() As BrokerAllocation
    Set NewBrokerAllocation = New BrokerAllocation
End Function

Public Function NewBrokerAllocationRow() As BrokerAllocationRow
    Set NewBrokerAllocationRow = New BrokerAllocationRow
End Function
