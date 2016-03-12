VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BrokerAllocationRow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' Private instance variable to hold the portfolio's data
Public RangeDb As Range

' Custom initializer. Parameter aRange must include a header row.  There is no attempt to consolidate
' multiple occurences of the same security
Public Sub InitializeWithRange(ARange As Range)
    If Not (ARange Is Nothing) Then
        Set Me.RangeDb = ARange.Resize(1, Me.GetNumberOfColumns)
    End If
End Sub

' Getters
' This returns the range holding the asset's value. This could be useful when trying to get the address
' of the range.
Property Get GetRange() As Range
    Set GetRange = Me.RangeDb
End Property

Property Get GetNumberOfColumns() As Integer
    Let GetNumberOfColumns = GetArrayLength(Me.GetHeaders)
End Property

Property Get GetCurrencyCode() As Range
    Set GetCurrencyCode = Me.GetRange.Columns(1)
End Property

Property Get GetBloombergExchangeCode() As Range
    Set GetBloombergExchangeCode = Me.GetRange.Columns(2)
End Property

Property Get GetBrokerCode() As Range
    Set GetBrokerCode = Me.GetRange.Columns(3)
End Property

Property Get GetAllocationPercentage() As Range
    Set GetAllocationPercentage = Me.GetRange.Columns(4)
End Property

Property Get GetAssetClassCode() As Range
    Set GetAssetClassCode = Me.GetRange.Columns(5)
End Property

Property Get GetStrategyCode() As Range
    Set GetStrategyCode = Me.GetRange.Columns(6)
End Property

Property Get GetDate() As Range
    Set GetDate = Me.GetRange.Columns(7)
End Property

Property Get GetTime() As Range
    Set GetTime = Me.GetRange.Columns(8)
End Property

Property Get GetRunNumber() As Range
    Set GetRunNumber = Me.GetRange.Columns(9)
End Property

Property Get GetPrimaryKey() As Range
    Set GetPrimaryKey = Me.GetRange.Columns(10)
End Property

' This function returns a 1D array with a Portfolio's headers
Public Function GetHeaders() As Variant
    Let GetHeaders = Array("CurrencyCode", "BloombergExchangeCode", "BrokerCode", "AllocationPercentage", _
                           "AssetClassCode", "StrategyCode", "Date", "Time", "RunNumber", "PrimnaryKey")
End Function

' Setters
Property Let SetCurrencyCode(Arg As String)
    Let Me.GetCurrencyCode.Value2 = Arg
End Property

Property Let SetBloombergExchangeCode(Arg As String)
    Let Me.GetBloombergExchangeCode.Value2 = Arg
End Property

Property Let SetBrokerCode(Arg As String)
    Let Me.GetBrokerCode.Value2 = Arg
End Property

Property Let SetAllocationPercentage(Arg As Double)
    Let Me.GetAllocationPercentage.Value2 = Arg
End Property

Property Let SetAssetCode(Arg As String)
    Let Me.GetAssetCode.Value2 = Arg
End Property

Property Let SetStrategyCode(Arg As String)
    Let Me.GetStrategyCode.Value2 = Arg
End Property

Property Let SetDate(Arg As Long)
    Let Me.GetDate.Value2 = Arg
End Property

Property Let SetTime(Arg As Long)
    Let Me.GetTime.Value2 = Arg
End Property

Property Let SetRunNumber(Arg As Integer)
    Let Me.GetRunNumber.Value2 = Arg
End Property
