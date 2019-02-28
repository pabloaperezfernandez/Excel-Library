VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TargetAssetAllocationRow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Base 1
Option Explicit

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

Property Get GetDate() As Range
    Set GetDate = Me.GetRange.Cells(1).Cells
End Property

Property Get GetTime() As Range
    Set GetTime = Me.GetRange.Columns(2).Cells
End Property

Property Get GetFundId() As Range
    Set GetFundId = Me.GetRange.Columns(3).Cells
End Property

Property Get GetAssetClass() As Range
    Set GetAssetClass = Me.GetRange.Columns(4).Cells
End Property

Property Get GetStrategyCode() As Range
    Set GetStrategyCode = Me.GetRange.Columns(5).Cells
End Property

Property Get GetPiamClientId() As Range
    Set GetPiamClientId = Me.GetRange.Columns(6).Cells
End Property

Property Get GetTargetUsdValue() As Range
    Set GetTargetUsdValue = Me.GetRange.Columns(7).Cells
End Property

Property Get GetCustodian() As Range
    Set GetCustodian = Me.GetRange.Columns(8).Cells
End Property

' This function returns a 1D array with a Portfolio's headers
Public Function GetHeaders() As Variant
    Let GetHeaders = Array("date", "time", "fund id", "asset class", "strategy code", "piam client id", _
                           "target usd value", "custodian")
End Function

' Setters
Property Let SetDate(Arg As Long)
    Let Me.GetDate.Value2 = Arg
End Property

Property Let SetTime(Arg As Long)
    Let Me.SetTime.Value2 = Arg
End Property

Property Let SetFundId(Arg As String)
    Let Me.GetFundId.Value2 = Arg
End Property

Property Let SetAssetClass(Arg As String)
    Let Me.GetAssetClass.Value2 = Arg
End Property

Property Let SetStrategyCode(Arg As String)
    Let Me.GetStrategyCode.Value2 = Arg
End Property

Property Let SetPiamClientId(Arg As String)
    Let Me.SetPiamClientId.Value2 = Arg
End Property

Property Let SetTargetUsdValue(Arg As String)
    Let Me.GetTargetUsdValue.Value2 = Arg
End Property

Property Let SetCustodian(Arg As String)
    Let Me.GetCustodian.Value2 = Arg
End Property
