VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PostTradingPortfolioAsset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Base 1
Option Explicit

' Private instance variable to hold the trade database
Private RangeDb As Range

' Custom initializer. aRange must be a row range.  All columns to the right of number 15 are ignored.
Public Sub InitializeWithRange(ARange As Range)
    Set RangeDb = ARange.Resize(1, Me.GetNumberOfColumns)
End Sub

' Getters
Property Get GetRange() As Range
    Set GetRange = RangeDb
End Property

Property Get GetHeaders() As Variant
    Let GetHeaders = Array("SecurityName", "BloombergTicker", "FinalUnits", "FinalUsdValue", "Weight", "FundId", _
                           "AssetClass", "Strategy", "CreationDate", "CreationTime", "Date", "Time")
End Property

Property Get GetNumberOfColumns() As Integer
    Let GetNumberOfColumns = GetArrayLength(GetHeaders())
End Property

Property Get GetSecurityName() As Range
    Set GetSecurityName = Me.GetRange.Cells(1)
End Property

Property Get GetBloombergTicker() As Range
    Set GetBloombergTicker = Me.GetRange.Cells(2)
End Property

Property Get GetFinalUnits() As Range
    Set GetFinalUnits = Me.GetRange.Cells(3)
End Property

Property Get GetFinalUsdValue() As Range
    Set GetFinalUsdValue = Me.GetRange.Cells(4)
End Property

Property Get GetWeight() As Range
    Set GetWeight = Me.GetRange.Cells(5)
End Property

Property Get GetFundId() As Range
    Set GetFundId = Me.GetRange.Cells(6)
End Property

Property Get GetAssetClass() As Range
    Set GetAssetClass = Me.GetRange.Cells(7)
End Property

Property Get GetStrategy() As Range
    Set GetStrategy = Me.GetRange.Cells(8)
End Property

Property Get GetCreationDate() As Range
    Set GetCreationDate = Me.GetRange.Cells(9)
End Property

Property Get GetCreationTime() As Range
    Set GetCreationTime = Me.GetRange.Cells(10)
End Property

Property Get GetDate() As Range
    Set GetDate = Me.GetRange.Cells(11)
End Property

Property Get GetTime() As Range
    Set GetTime = Me.GetRange.Cells(12)
End Property

' Setters
Property Let SetSecurityName(Arg As String)
    Let Me.GetSecurityName.Value2 = Arg
End Property

Property Let SetBloombergTicker(Arg As Variant)
    Let Me.GetBloombergTicker.Value2 = Arg
End Property

Property Let SetFinalUnits(Arg As Variant)
    Let Me.GetFinalUnits.Value2 = Arg
End Property

Property Let SetFinalUsdValue(Arg As Variant)
    Let Me.GetFinalUsdValue.Value2 = Arg
End Property

Property Let SetWeight(Arg As Variant)
    Let Me.GetWeight.Value2 = Arg
End Property

Property Let SetFundId(Arg As Variant)
    Let Me.GetFundId.Value2 = Arg
End Property

Property Let SetAssetClass(Arg As Variant)
    Let Me.GetAssetClass.Value2 = Arg
End Property

Property Let SetStrategy(Arg As Variant)
    Let Me.GetStrategy.Value2 = Arg
End Property

Property Let SetCreationDate(Arg As Variant)
    Let Me.GetCreationDate.Value2 = Arg
End Property

Property Let SetCreationTime(Arg As Variant)
    Let Me.GetCreationTime.Value2 = Arg
End Property

Property Let SetDate(Arg As Variant)
    Let Me.GetDate.Value2 = Arg
End Property

Property Let SetTime(Arg As Long)
    Let Me.GetTime.Value2.Value2 = Arg
End Property
