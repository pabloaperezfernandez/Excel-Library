VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MasterFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Base 1
Option Explicit

Public UnderlyingWorkbook As Workbook
' Private instance variable to hold the trade database
Private RangeDb As Range
' The following two store the row positions of the corresponding identifiers
Private SedolRowPositionDictionary As Dictionary
Private BloombergTickerRowPositionDictionary As Dictionary
Private IsinRowPositionDictionary As Dictionary

' This is the class destructor.
Public Sub Class_Terminate()
    ' Determine if this was created from a file.  If it was, close the source workbook without saving any
    ' changes
    If Not Me.GetRange.Worksheet.Parent Is ThisWorkbook Then
        Call Me.GetRange.Worksheet.Parent.Close(SaveChanges:=False)
    End If
End Sub

' Sub to save changes
Public Sub SaveChanges()
    Call Me.GetRange.Worksheet.Parent.Save
End Sub

' This returns the range holding the asset's value. This could be useful when trying to get the address
' of the range.
Property Get GetRange() As Range
    Set GetRange = RangeDb
End Property

' Returns the number of securities in the master file
Property Get GetDbLength() As Long
    Let GetDbLength = GetRange.Rows.Count
End Property

' Returns the ith row in the master file as an instance of PrivateFileRecord
Property Get GetIthMasterFileRow(i As Long) As MasterFileRow
    Dim AMasterFileRow As MasterFileRow
    
    ' Instantiate PrivateFileRecord
    Set AMasterFileRow = New MasterFileRow
    
    ' Initialize instance with data from appropriate master
    Call AMasterFileRow.InitializeWithRange(GetRange.Rows(i))
    
    Set GetIthMasterFileRow = AMasterFileRow
End Property

' Returns an instance of PrivateFileRecord corresponding to the security with the given 7-digit SEDOL
Property Get GetMasterFileRowWith7DigitSedol(TheSedol As String) As MasterFileRow
    Dim AMasterFileRow As MasterFileRow
    Dim RowIndex As Long
    
    ' Exit returning Nothing if TheSedol is not in the master file
    If Not Me.SevenDigitSedolExistsQ(TheSedol) Then
        Set GetMasterFileRowWith7DigitSedol = Nothing
        
        Exit Property
    End If
    
    ' Instantiate MasterFileRow
    Set AMasterFileRow = New MasterFileRow
    
    ' Get row index of of the given SEDOL in the mater file
    Let RowIndex = SedolRowPositionDictionary.Item(TheSedol)
    
    ' Initialize MasterFileRow instance with the range holding master file data for this SEDOL
    Call AMasterFileRow.InitializeWithRange(GetIthMasterFileRow(RowIndex).GetRange)
    
    ' Return the reference to the instance of MasterFileRow holding this row of the mater file
    Set GetMasterFileRowWith7DigitSedol = AMasterFileRow
End Property

' Returns an instance of PrivateFileRecord corresponding to the security with the given ISIN.
' If the ISIN is not in the master file, this function returns Nothing.
Property Get GetMasterFileRowWithIsin(TheIsin As String) As MasterFileRow
    Dim AMasterFileRow As MasterFileRow
    Dim RowIndex As Long

    ' Exit returning Nothing if TheIsin is not in the master file
    If Not Me.SevenDigitSedolExistsQ(TheIsin) Then
        Set GetMasterFileRowWith7DigitSedol = Nothing
        
        Exit Property
    End If
    
    ' Instantiate MasterFileRow
    Set AMasterFileRow = New MasterFileRow

    ' Get row index of of the given SEDOL in the mater file
    Let RowIndex = IsinRowPositionDictionary.Item(TheIsin)
    
    ' Initialize MasterFileRow instance with the range holding master file data for this SEDOL
    If RowIndex = 0 Then
        ' Return nothing since this ISIN is not in the master file
        Set GetMasterFileRowWithIsin = Nothing
    Else
        Call AMasterFileRow.InitializeWithRange(GetIthMasterFileRow(RowIndex).GetRange)
        
        ' Return the reference to the instance of MasterFileRow holding this row of the mater file
        Set GetMasterFileRowWithIsin = AMasterFileRow
    End If
End Property

' Returns an instance of PrivateFileRecord corresponding to the security with the given Bloomberg ticker
Property Get GetMasterFileRowWithBloombergTicker(BloombergTicker As String) As MasterFileRow
    Dim AMasterFileRow As MasterFileRow
    Dim RowIndex As Long

    ' Exit returning Nothing if BloombergTicker is not in the master file
    If Not Me.SevenDigitSedolExistsQ(BloombergTicker) Then
        Set GetMasterFileRowWith7DigitSedol = Nothing
        
        Exit Property
    End If

    ' Instantiate MasterFileRow
    Set AMasterFileRow = New MasterFileRow

    ' Get row index of of the given SEDOL in the mater file
    Let RowIndex = BloombergTickerRowPositionDictionary.Item(BloombergTicker)

    ' Initialize MasterFileRow instance with the range holding master file data for this SEDOL
    Call AMasterFileRow.InitializeWithRange(GetIthMasterFileRow(RowIndex).GetRange)
    
    ' Return the reference to the instance of MasterFileRow holding this row of the mater file
    Set GetMasterFileRowWithBloombergTicker = AMasterFileRow
End Property

' Returns the primary trading security corresponding to the given ISIN as an instance of MasterFileRow.
' This correspondings to the MasterFileRow whose BloombergTicker is the EQY_FUND_TICKER corresponding
' to the given ISIN
Property Get GetPrimaryMasterFileRowWithIsin(TheIsin As String) As MasterFileRow
    Dim PrimaryMasterFileRow As MasterFileRow
    Dim RowIndex As Long
    
    ' Instantiate MasterFileRow
    Set PrimaryMasterFileRow = New MasterFileRow
    
    ' Get row index of the given SEDOL in the mater file
    Let RowIndex = IsinRowPositionDictionary.Item(TheIsin)

    ' Initialize MasterFileRow instance with the range holding master file data for this SEDOL
    Call PrimaryMasterFileRow.InitializeWithRange(GetIthMasterFileRow(RowIndex).GetRange)
    
    ' Get row index of the EQY_FUND_TICKER of the given ISIN
    Let RowIndex = BloombergTickerRowPositionDictionary.Item(PrimaryMasterFileRow.GetEqyFundTicker.Value)
    
    ' Initialize MasterFileRow instance with the range holding master file data for this EQY_FUND_TICKER
    Call PrimaryMasterFileRow.InitializeWithRange(GetIthMasterFileRow(RowIndex).GetRange)
    
    ' Return the reference to the instance of MasterFileRow holding this row of the mater file
    Set GetPrimaryMasterFileRowWithIsin = PrimaryMasterFileRow
End Property

' Custom initializer. aRange must be the master file table, including the headers row
' The first column of the source range in the master file is ignored because it contains redundant information
Public Sub InitializeWithRange(ARange As Range)
    Dim NumberOfSecurities As Long
    Dim LastRow As Long
    
    ' Exit if aRange is Nothing
    If ARange Is Nothing Then
        Exit Sub
    End If
    
    ' Set UnderlyingWorkbook
    Set UnderlyingWorkbook = ARange.Worksheet.Parent
    
    Let LastRow = LastNonBlankRowIndexInColumnNotEqualTo(ARange.Worksheet.Cells(1, 1), Array("", " "))
    
    ' Set reference to the underlying range holding the data. We use the column
    ' position of the last column as the size of the underlying range.
    Set RangeDb = ARange.Offset(1, 1).Resize(LastRow - 1, 48)
    
    ' Initialize row position dictionaries
    Call InitializeRowPositionDictionaries
End Sub

' Initiaizes the MasterFile object with the master file data contained in a file
Public Sub InitializeWithFile(AFileName As String)
    Dim MasterSourceWorkbook As Workbook
    Dim MasterSourceWorksheet As Worksheet
    Dim MasterSourceRange As Range
    Dim TargetRange As Range
    Dim LastRow As Long

    ' Open workbook containing map and a save a reference
    Set MasterSourceWorkbook = Application.Workbooks.Open(Filename:=AFileName, ReadOnly:=True)
    
    ' Set reference to worksheet containing master file data
    Set MasterSourceWorksheet = MasterSourceWorkbook.Worksheets("data cleaned")
    
    ' Compute the number of securities
    Let LastRow = LastNonBlankRowIndexInColumnNotEqualTo(MasterSourceWorksheet.Cells(1, 1), Array("", " "))
    
    ' Exit if lastrow < 2
    If LastRow < 2 Then
        Exit Sub
    End If
    
    ' Set reference to range containing the master file data
    Set MasterSourceRange = MasterSourceWorksheet.Range("A1").Resize(LastRow, 48)
    
    ' Initialize MasterFile instance with this copy of the data range
    Call InitializeWithRange(MasterSourceRange)
End Sub

' This is private function is used by InitializeWithRange() to initialize the row index
' dictionaries
Private Sub InitializeRowPositionDictionaries()
    Dim i As Long
    
    ' Instantiate the dictionary objects
    Set SedolRowPositionDictionary = New Dictionary
    Set BloombergTickerRowPositionDictionary = New Dictionary
    Set IsinRowPositionDictionary = New Dictionary
    
    ' Initialize the dictionaries
    For i = 1 To GetDbLength
        If Not SedolRowPositionDictionary.Exists(GetIthMasterFileRow(i).Get7DigitSedol.Value) Then
            Call SedolRowPositionDictionary.Add(GetIthMasterFileRow(i).Get7DigitSedol.Value, i)
        End If
        
        If Not BloombergTickerRowPositionDictionary.Exists(GetIthMasterFileRow(i).GetBloombergTicker.Value) Then
            Call BloombergTickerRowPositionDictionary.Add(GetIthMasterFileRow(i).GetBloombergTicker.Value, i)
        End If
        
        If Not IsinRowPositionDictionary.Exists(GetIthMasterFileRow(i).GetIsin.Value) Then
            Call IsinRowPositionDictionary.Add(GetIthMasterFileRow(i).GetIsin.Value, i)
        End If
    Next i
End Sub

' Getters
Property Get GetIsins() As Range
    Set GetIsins = GetRange.Columns(1).Cells
End Property

Property Get Get6DigitSedols() As Range
    Set Get6DigitSedols = GetRange.Columns(2).Cells
End Property

Property Get Get7DigitSedols() As Range
    Set Get7DigitSedols = GetRange.Columns(3).Cells
End Property

Property Get GetParsekeableDesSources() As Range
    Set GetParsekeableDesSources = GetRange.Columns(4).Cells
End Property

Property Get GetBloombergTickers() As Range
    Set GetBloombergTickers = GetRange.Columns(5).Cells
End Property

Property Get GetExchCodes() As Range
    Set GetExchCodes = GetRange.Columns(6).Cells
End Property

Property Get GetCrncies() As Range
    Set GetCrncy = GetRange.Columns(7).Cells
End Property

Property Get GetCntryOfDomiciles() As Range
    Set GetCntryOfDomicile = GetRange.Columns(8).Cells
End Property

Property Get GetGicsSubIndustries() As Range
    Set GetGicsSubIndustry = GetRange.Columns(9).Cells
End Property

Property Get GetMsciIndustryCodes() As Range
    Set GetMsciIndustryCode = GetRange.Columns(10).Cells
End Property

Property Get GetPalladyneSectorCodes() As Range
    Set GetPalladyneSectorCode = GetRange.Columns(11).Cells
End Property

Property Get GetRegions() As Range
    Set GetRegion = GetRange.Columns(12).Cells
End Property

Property Get GetNames() As Range
    Set GetName = GetRange.Columns(13).Cells
End Property

Property Get GetUsdMarketValues() As Range
    Set GetUsdMarketValue = GetRange.Columns(14).Cells
End Property

Property Get GetMarketStatuses() As Range
    Set GetMarketStatus = GetRange.Columns(15).Cells
End Property

Property Get GetSecurityTypes() As Range
    Set GetSecurityType = GetRange.Columns(16).Cells
End Property

Property Get GetCountryTradableFlags() As Range
    Set GetCountryTradableFlags = GetRange.Columns(17).Cells
End Property

Property Get GetPiamTradableCountryNames() As Range
    Set GetPiamTradableCountryNames = GetRange.Columns(18).Cells
End Property

Property Get Get30DayAverageTradingVolumes() As Range
    Set Get30DayAverageTradingVolumes = GetRange.Columns(19).Cells
End Property

Property Get Get30DayAverageTradingValues() As Range
    Set Get30DayAverageTradingValues = GetRange.Columns(20).Cells
End Property

Property Get GetCountryFullNames() As Range
    Set GetCountryFullNames = GetRange.Columns(21).Cells
End Property

Property Get GetTradeLotSizes() As Range
    Set GetTradeLotSizes = GetRange.Columns(22).Cells
End Property

Property Get GetCountryRiskIsoCodes() As Range
    Set GetCountryRiskIsoCodes = GetRange.Columns(23).Cells
End Property

Property Get GetNames2() As Range
    Set GetNames2 = GetRange.Columns(24).Cells
End Property

Property Get GetAdrGdrAdsPrfdClosedConvReitMiscFlags() As Range
    Set GetAdrGdrAdsPrfdClosedConvReitMiscFlags = GetRange.Columns(25).Cells
End Property

Property Get GetGet30DayAverageTradingVolumes() As Range
    Set GetGet30DayAverageTradingVolumes = GetRange.Columns(31).Cells
End Property

Property Get GetSecurityTyp2s() As Range
    Set GetSecurityTyp2s = GetRange.Columns(32).Cells
End Property

Property Get GetCieDes() As Range
    Set GetCieDes = GetRange.Columns(36).Cells
End Property

Property Get GetEqyFundTickers() As Range
    Set GetEqyFundTickers = GetRange.Columns(37).Cells
End Property

Property Get GetWaccs() As Range
    Set GetWaccs = GetRange.Columns(38).Cells
End Property

Property Get GetUsdTradeLotValues() As Range
    Set GetUsdTradeLotValues = GetRange.Columns(39).Cells
End Property

Property Get GetPriceTimes1MAverageTradingVolumes() As Range
    Set GetPriceTimes1MAverageTradingVolumes = GetRange.Columns(40).Cells
End Property

Property Get GetAboveMinimumUsdAverageTradingValueFlags() As Range
    Set GetAboveMinimumUsdAverageTradingValueFlags = GetRange.Columns(41).Cells
End Property

Property Get GetBloombergTickerPlusAboveMinimumUsdAverageTradingValueFlags() As Range
    Set GetBloombergTickerPlusAboveMinimumUsdAverageTradingValueFlags = GetRange.Columns(42).Cells
End Property

Property Get GetTradableAlternativeBloombergTickers() As Range
    Set GetTradableAlternativeBloombergTickers = GetRange.Columns(44).Cells
End Property

Property Get GetTradableAlternative6DigitSedols() As Range
    Set GetTradableAlternative6DigitSedols = GetRange.Columns(45).Cells
End Property

Property Get GetPalladyneSectors() As Range
    Set GetPalladyneSectors = GetRange.Columns(46).Cells
End Property

Property Get GetInvestibleSedols() As Range
    Set GetInvestibleSedols = GetRange.Columns(47).Cells
End Property

Property Get GetInvestibleBloombergTickers() As Range
    Set GetInvestibleBloombergTickers = GetRange.Columns(48).Cells
End Property

' This function returns a 1D array of strings containing all master file ISINs whose EQY_FUND_TICKER equals the given EQY_FUND_TICKER ticker.  If
' the given EQY_FUND_TICKER is not in the master file, the property returns an empty array (e.g. Array())
' property will fail.
Property Get GetAllIsinsWithGivenEqyFundTicker(EqyFundTicker As String) As Variant
    Dim tmpSht As Worksheet
    Dim NumberOfIdentifiers As Long
    Dim i As Long
    
    If Not Me.BloombergTickerExistsQ(EqyFundTicker) Then
        Exit Property
    End If
    
    ' Set reference to ThisWorkbook.Worksheets("TempComputation")
    Set tmpSht = ThisWorkbook.Worksheets("TempComputation")
    
    ' Clear ThisWorkbook.TempComputation worksheet
    Call tmpSht.UsedRange.ClearContents
        
    ' Autofilter the master file to get all rows with the given EQY_FUND_TICKER
    Let GetRange.Worksheet.AutoFilterMode = False
    
    ' Autofiter the data to show only the first instance of each element
    GetRange.Worksheet.Range("A1").CurrentRegion.AutoFilter Field:=Me.GetEqyFundTickers.Column, _
                                                            Criteria1:=EqyFundTicker
        
    ' Copy visible cells so they collapse into one contiguous range.
    Call GetRange.Worksheet.AutoFilter.Range.SpecialCells(xlCellTypeVisible).Copy
    
    ' Clear worksheet TempComputation and paste the filtered data
    Call ThisWorkbook.Worksheets("TempComputation").Range("A1").PasteSpecial(xlPasteValues)
    
    ' Clear autofilters
    Let GetRange.Worksheet.AutoFilterMode = False

    ' Extract the list of ISINs and packed into an array
    Let NumberOfIdentifiers = LastNonBlankRowIndexInColumnNotEqualTo(tmpSht.Cells(1, GetIsins.Column), Array("", " ")) - 1
    
    ' Return 1D array of mf ISINs whose EQY_FUND_TICKER equals aBbTicker
    If NumberOfIdentifiers > 0 Then
        Let GetAllIsinsWithGivenEqyFundTicker = UniqueSubset(ConvertTo1DArray(tmpSht.Cells(1, GetIsins.Column).Offset(1, 0).Resize(NumberOfIdentifiers, 1).Value2))
    Else
        Let GetAllIsinsWithGivenEqyFundTicker = Array()
    End If
End Property

' This function returns a 1D array of strings containing all master file 7-char SEDOLs whose EQY_FUND_TICKER equals the given EQY_FUND_TICKER ticker.  If
' the given EQY_FUND_TICKER is not in the master file, the property returns an empty array (e.g. Array())
Property Get GetAll7CharSedolsWithGivenEqyFundTicker(EqyFundTicker As String) As Variant
    Dim tmpSht As Worksheet
    Dim NumberOfIdentifiers As Long
    Dim i As Long
    
    If Not Me.BloombergTickerExistsQ(EqyFundTicker) Then
        Exit Property
    End If
    
    ' Set reference to ThisWorkbook.Worksheets("TempComputation")
    Set tmpSht = ThisWorkbook.Worksheets("TempComputation")
    
    ' Clear ThisWorkbook.TempComputation worksheet
    Call tmpSht.UsedRange.ClearContents
        
    ' Autofilter the master file to get all rows with the given EQY_FUND_TICKER
    Let GetRange.Worksheet.AutoFilterMode = False
    
    ' Autofiter the data to show only the first instance of each element
    GetRange.Worksheet.Range("A1").CurrentRegion.AutoFilter Field:=Me.GetEqyFundTickers.Column, _
                                                            Criteria1:=EqyFundTicker
        
    ' Copy visible cells so they collapse into one contiguous range.
    Call GetRange.Worksheet.AutoFilter.Range.SpecialCells(xlCellTypeVisible).Copy
    
    ' Clear worksheet TempComputation and paste the filtered data
    Call ThisWorkbook.Worksheets("TempComputation").Range("A1").PasteSpecial(xlPasteValues)
    
    ' Clear autofilters
    Let GetRange.Worksheet.AutoFilterMode = False

    ' Extract the list of 7-char SEDOLs and packed into an array
    Let NumberOfIdentifiers = LastNonBlankRowIndexInColumnNotEqualTo(tmpSht.Cells(1, Me.Get7DigitSedols.Column), Array("", " ")) - 1
    
    ' Return 1D array of mf 7-char SEDOLs whose EQY_FUND_TICKER equals aBbTicker
    If NumberOfIdentifiers > 0 Then
        Let GetAll7CharSedolsWithGivenEqyFundTicker = UniqueSubset(ConvertTo1DArray(tmpSht.Cells(1, Me.Get7DigitSedols.Column).Offset(1, 0).Resize(NumberOfIdentifiers, 1).Value2))
    Else
        Let GetAll7CharSedolsWithGivenEqyFundTicker = Array()
    End If
End Property

' This function returns a 1D array of strings containing all master file Bloomberg Tickers whose EQY_FUND_TICKER equals the given EQY_FUND_TICKER.  If
' the given EQY_FUND_TICKER is not in the master file, the property returns an empty array (e.g. Array())
Property Get GetAllBloombergTickersWithGivenEqyFundTicker(EqyFundTicker As String) As Variant
    Dim tmpSht As Worksheet
    Dim NumberOfIdentifiers As Long
    Dim i As Long
    
    If Not Me.BloombergTickerExistsQ(EqyFundTicker) Then
        Exit Property
    End If
    
    ' Set reference to ThisWorkbook.Worksheets("TempComputation")
    Set tmpSht = ThisWorkbook.Worksheets("TempComputation")
    
    ' Clear ThisWorkbook.TempComputation worksheet
    Call tmpSht.UsedRange.ClearContents
        
    ' Autofilter the master file to get all rows with the given EQY_FUND_TICKER
    Let GetRange.Worksheet.AutoFilterMode = False
    
    ' Autofiter the data to show only the first instance of each element
    GetRange.Worksheet.Range("A1").CurrentRegion.AutoFilter Field:=Me.GetEqyFundTickers.Column, _
                                                            Criteria1:=EqyFundTicker
        
    ' Copy visible cells so they collapse into one contiguous range.
    Call GetRange.Worksheet.AutoFilter.Range.SpecialCells(xlCellTypeVisible).Copy
    
    ' Clear worksheet TempComputation and paste the filtered data
    Call ThisWorkbook.Worksheets("TempComputation").Range("A1").PasteSpecial(xlPasteValues)
    
    ' Clear autofilters
    Let GetRange.Worksheet.AutoFilterMode = False

    ' Extract the list of 7-char SEDOLs and packed into an array
    Let NumberOfIdentifiers = LastNonBlankRowIndexInColumnNotEqualTo(tmpSht.Cells(1, Me.GetBloombergTickers.Column), Array("", " ")) - 1
    
    ' Return 1D array of mf 7-char SEDOLs whose EQY_FUND_TICKER equals aBbTicker
    If NumberOfIdentifiers > 0 Then
        Let GetAllBloombergTickersWithGivenEqyFundTicker = UniqueSubset(ConvertTo1DArray(tmpSht.Cells(1, Me.GetBloombergTickers.Column).Offset(1, 0).Resize(NumberOfIdentifiers, 1).Value2))
    Else
        Let GetAllBloombergTickersWithGivenEqyFundTicker = Array()
    End If
End Property

' Predicates
' Returns True or False depending on whether or not the give SEDOL corresponds to
' a security in the master file.
Public Function SevenDigitSedolExistsQ(TheSedol As String) As Boolean
    Let SevenDigitSedolExistsQ = SedolRowPositionDictionary.Exists(TheSedol)
End Function

' Returns True or False depending on whether or not the give ISIN corresponds to
' a security in the master file.
Public Function IsinExistsQ(TheIsin As String) As Boolean
    Let IsinExistsQ = IsinRowPositionDictionary.Exists(TheIsin)
End Function

' Returns True or False depending on whether or not the give Bloomberg ticker corresponds to
' a security in the master file.
Public Function BloombergTickerExistsQ(TheTicker As String) As Boolean
    Let BloombergTickerExistsQ = BloombergTickerRowPositionDictionary.Exists(TheTicker)
End Function


