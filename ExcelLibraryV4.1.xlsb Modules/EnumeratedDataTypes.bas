Attribute VB_Name = "EnumeratedDataTypes"
' This is the enumerated data type for type of quantiling. For example: quartile, quintile, decile, percentile, etc.
Enum QuantileType
    Quintile
    Decile
End Enum

Enum AssetClassType
    FX
    Equity
    FixedIncome
    Convertible
    StructureProduct
End Enum

' This is used to denote the screen type with Thomson Reuters Spreadsheet Link (TRSL)
Enum PrimarySecondaryFlag
    Primary
    Secondary
    All
End Enum

' This is used for dealing with Bloomberg's corporate action alerts
Enum CorporateActionType
    Dividend
    StockSplit
    ReversedSplit
    Acquisition
    Sale
    StockBuyBack
    Divestiture
    EquityOffering
    Delisting
    NameChange
    TickerSymbolChange
    IdNumberChange
    DebtOffering
    RightsIssue
    StockDividend
    DebtRedemption
    Other
End Enum

' This is used to indicate if ArrayFormulas.ConvertRecordSetPayloadToMatrix should return headers, headers and body, or just the body
Enum ConvertRecordSetPayloadToMatrixOptionsType
    Headers
    Body
    HeadersAndBody
End Enum

' This is used to convert instances of type AssetClassType to their string
' representations.
Public Function ConvertPrimarySecondaryFlagToString(TheAssetType As AssetClassType) As String
    Select Case PrimarySecondaryFlag
        Case Primary
            Let ConvertPrimarySecondaryFlagToString = "Primary"
        Case Secondary
            Let ConvertPrimarySecondaryFlagToString = "Not Primary"
        Case Else
            Let ConvertPrimarySecondaryFlagToString = "All"
    End Select
End Function

' This is done to convert a string to an instance of data type AssetClassType
Public Function ConvertStringToPrimarySecondaryFlag(TheAssetClassString As String) As PrimarySecondaryFlag
    Select Case TheAssetClassString
        Case "Primary"
            Let ConvertStringToPrimarySecondaryFlag = Primary
        Case "Not Primary"
            Let ConvertStringToPrimarySecondaryFlag = Secondary
        Case Else
            Let ConvertStringToPrimarySecondaryFlag = All
    End Select
End Function

' This is used to convert instances of type AssetClassType to their string
' representations.
Public Function ConvertAssetClassTypeToString(TheAssetType As AssetClassType) As String
    Select Case TheAssetType
        Case FX
            Let ConvertAssetClassTypeToString = "FX"
        Case Equity
            Let ConvertAssetClassTypeToString = "Equity"
        Case FixedIncome
            Let ConvertAssetClassTypeToString = "Fixed Income"
        Case Convertible
            Let ConvertAssetClassTypeToString = "Convertible"
        Case Else
            Let ConvertAssetClassTypeToString = "Structured Product"
    End Select
End Function

' This is done to convert a string to an instance of data type AssetClassType
Public Function ConvertStringToAssetClassType(TheAssetClassString As String) As AssetClassType
    Select Case TheAssetClassString
        Case "FX"
            Let ConvertAssetClConvertStringToAssetClassTypeassTypeToString = FX
        Case "Equity"
            Let ConvertStringToAssetClassType = Equity
        Case "Fixed Income"
            Let ConvertStringToAssetClassType = FixedIncome
        Case "Convertible"
            Let ConvertStringToAssetClassType = Convertible
        Case Else
            Let ConvertStringToAssetClassType = StructureProduct
    End Select
End Function

