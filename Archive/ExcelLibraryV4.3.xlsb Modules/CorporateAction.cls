VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CorporateAction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Option Base 1

Private ActionType As CorporateActionType
Private ActionUpdateQ As Boolean

Property Get GetType() As CorporateActionType
    Let GetType = ActionType
End Property

Property Let SetType(Arg As CorporateActionType)
    Let ActionType = Arg
End Property

Property Get UpdateQ() As Boolean
    Let UpdateQ = Me.ActionUpdateQ
End Property

Property Let SetUpdateQ(Arg As Boolean)
    Let Me.UpdateQ = Arg
End Property

Property Let SetTypeFromString(Arg As String)
    Select Case Arg
        Case "Dividend"
            Let ActionType = Dividend
        Case "StockSplit"
            Let ActionType = StockSplit
        Case "ReversedSplit"
            Let ActionType = ReversedSplit
        Case "Acquisition"
            Let ActionType = Acquisition
        Case "Sale"
            Let ActionType = Sale
        Case "StockBuyBack"
            Let ActionType = StockBuyBack
        Case "Divestiture"
            Let ActionType = Divestiture
        Case "EquityOffering"
            Let ActionType = EquityOffering
        Case "Delisting"
            Let ActionType = Delisting
        Case "NameChange"
            Let ActionType = NameChange
        Case "TickerSymbolChange"
            Let ActionType = TickerSymbolChange
        Case "IdNumberChange"
            Let ActionType = IdNumberChange
        Case "DebtOffering"
            Let ActionType = DebtOffering
        Case "RightsIssue"
            Let ActionType = RightsIssue
        Case "StockDividend"
            Let ActionType = StockDividend
        Case "DebtRedemption"
            Let ActionType = DebtRedemption
        Case Else
            Let ActionType = Other
    End Select
End Property

Property Get GetTypeAsString() As String
    Select Case Me.GetType
        Case Dividend
            Let GetTypeAsString = "Dividend"
        Case StockSplit
            Let GetTypeAsString = "StockSplit"
        Case ReversedSplit
            Let GetTypeAsString = "ReversedSplit"
        Case Acquisition
            Let GetTypeAsString = "Acquisition"
        Case Sale
            Let GetTypeAsString = "Sale"
        Case StockBuyBack
            Let GetTypeAsString = "StockBuyBack"
        Case Divestiture
            Let GetTypeAsString = "Divestiture"
        Case EquityOffering
            Let GetTypeAsString = "EquityOffering"
        Case Delisting
            Let GetTypeAsString = "Delisting"
        Case NameChange
            Let GetTypeAsString = "NameChange"
        Case TickerSymbolChange
            Let GetTypeAsString = "TickerSymbolChange"
        Case IdNumberChange
            Let GetTypeAsString = "IdNumberChange"
        Case DebtOffering
            Let GetTypeAsString = "DebtOffering"
        Case RightsIssue
            Let GetTypeAsString = "RightsIssue"
        Case StockDividend
            Let GetTypeAsString = "StockDividend"
        Case DebtRedemption
            Let GetTypeAsString = "DebtRedemption"
        Case Other
            Let GetTypeAsString = "Other"
    End Select
End Property

