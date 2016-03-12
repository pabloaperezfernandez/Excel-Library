VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PrivateFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Base 1
Option Explicit

' Private instance variable to hold the portfolio's data
Public RangeDb As Range
' Private instance variable to hold row index of each SEDOL in RangeDb
Public BlooombergTickerRowIndicesDictionary As Dictionary
' Private instance variable to hold row index of each SEDOL in RangeDb
Public SedolRowIndicesDictionary As Dictionary
' Private instance variable to hold row index of each primary Bloomberg ticker in RangeDb
Public PrimaryBloombergTickerRowIndicesDictionary As Dictionary
' This is the workbook underlying the data
Public UnderlyingWorkbook As Workbook
' Private instance variable to hold the country code-to-region map
Public CountryRegionMapping As CountryCodeRegionalCodeMap
' Private instance variable to hold the regional size thresholds
Public SizeThresholds As Dictionary
' Private instance variable to hold the regional/sectoral style thresholds
Public StyleThresholds As Dictionary
' Private instance variable to hold GICS-to-MSCI Sub Industry Code mapping
Public GicsToMsciSubIndustryCodeMapping As Dictionary

' This destructor closes (without saving changes) the workbook holding the underlying data
' This is the class destructor.
Public Sub Class_Terminate()
    If Not (Me.UnderlyingWorkbook Is Nothing) Then
        If Not Me.UnderlyingWorkbook Is ThisWorkbook Then
            Call Me.UnderlyingWorkbook.Close(SaveChanges:=False)
        End If
    End If
End Sub

' Custom initializer. Parameter aRange must include a header row.
' The same 7-char SEDOL CANNOT appear in more than one row.
Public Sub InitializeWithRange(ARange As Range)
    Dim LastRowIndex As Long

    Let LastRowIndex = LastNonBlankRowIndexInColumnNotEqualTo(ARange.Columns(1), "")

    If Not (ARange Is Nothing) Then
        ' Set UnderlyingWorkbook
        Set Me.UnderlyingWorkbook = ARange.Worksheet.Parent
    
        ' As long as aRange has more than one row, initialize RangeDb and RowIndicesDictionary
        If LastRowIndex > 1 Then
            ' Exclude headers row
            Set Me.RangeDb = ARange.Offset(1, 0).Resize(LastRowIndex - 1, Me.GetNumberOfColumns)
    
            ' Initialize row indices dictionaries
            Call InitializeRowIndicesDictionaries
        End If
    End If
End Sub

' This sub uses worksheet Classification in workbook Equity1Dot4 to fill the classifications column of this instance
' Clearly, this classification can only be run after the instance has been initialized.
Public Sub ClassifyWithEquity1Dot4(AFileName As String)
    Dim i As Long
    Dim wsht As Worksheet
    Dim NRows As Long
    Dim BbDict As New Dictionary
    Dim SedolDict As New Dictionary
    Dim IsinDict As New Dictionary
    Dim TheClassifications() As String
    
    ' Open and reference Classification worksheet in Equity1Dot4
    Set wsht = Application.Workbooks.Open(Filename:=AFileName, UpdateLinks:=False, ReadOnly:=True).Worksheets("Classification")
    
    ' Count number of rows in Classification worksheet
    Let NRows = Application.CountA(wsht.Range("A:A")) - 1
    
    ' Initialize dictionaries
    For i = 2 To NRows
        If Not BbDict.Exists(Key:=wsht.Cells(i, 23).Value2) Then
            Call BbDict.Add(Key:=wsht.Cells(i, 23).Value2, Item:=wsht.Cells(i, 13).Value2)
        End If
        
        If Not SedolDict.Exists(Key:=wsht.Cells(i, 1).Value2) Then
            Call SedolDict.Add(Key:=wsht.Cells(i, 1).Value2, Item:=wsht.Cells(i, 13).Value2)
        End If
        
        If Not IsinDict.Exists(Key:=wsht.Cells(i, 21).Value2) Then
            Call IsinDict.Add(Key:=wsht.Cells(i, 21).Value2, Item:=wsht.Cells(i, 13).Value2)
        End If
    Next i
    
    ReDim TheClassifications(Me.GetDbLength)
    
    For i = 1 To Me.GetDbLength
        If BbDict.Exists(Key:=Me.GetIthRow(i).GetBloombergTicker.Value2) Then
            Let TheClassifications(i) = BbDict.Item(Key:=Me.GetIthRow(i).GetBloombergTicker.Value2)
        ElseIf BbDict.Exists(Key:=Me.GetIthRow(i).GetInvestableBloombergTicker.Value2) Then
            Let TheClassifications(i) = BbDict.Item(Key:=Me.GetIthRow(i).GetInvestableBloombergTicker.Value2)
        ElseIf SedolDict.Exists(Key:=Me.GetIthRow(i).GetSedol.Value2) Then
            Let TheClassifications(i) = BbDict.Item(Key:=Me.GetIthRow(i).GetSedol.Value2)
        ElseIf SedolDict.Exists(Key:=Me.GetIthRow(i).GetInvestableSedol.Value2) Then
            Let TheClassifications(i) = BbDict.Item(Key:=Me.GetIthRow(i).GetInvestableSedol.Value2)
        ElseIf IsinDict.Exists(Key:=Me.GetIthRow(i).GetIsin.Value2) Then
            Let TheClassifications(i) = BbDict.Item(Key:=Me.GetIthRow(i).GetIsin.Value2)
        ElseIf IsinDict.Exists(Key:=Me.GetIthRow(i).GetInvestableIsin.Value2) Then
            Let TheClassifications(i) = BbDict.Item(Key:=Me.GetIthRow(i).GetInvestableIsin.Value2)
        Else
            Let TheClassifications(i) = "NULL"
        End If
    Next i
    
    ' Close equity1dot4
    Call wsht.Parent.Close(SaveChanges:=False)
    
    ' Dump classifications in the appropriate column
    Let Me.GetClassifications.Value2 = Application.Transpose(TheClassifications)
End Sub

' Custom initializer.  The file must not have repeated securities (e.g. the same 7-char SEDOL must not appear twice).
' The filename must be fully qualified
' Custom initializer.  The file must not have repeated securities (e.g. the same 7-char SEDOL must not appear twice).
' The filename must be fully qualified
Public Sub InitializeWithPrivateFile(AFileName As String, Optional OptionalDate As Variant)
    Dim i As Integer
    Dim SourceRange As Range
    Dim sourceSht As Worksheet
    Dim SourceToTargetColumnIndexMapping As Variant
    Dim targetSht As Worksheet
    Dim TargetRange As Range
    Dim TheDate As Long
    Dim ARow As PrivateFileRow
    Dim TheHeaders As Variant
    Dim LastRowIndex As Long
    
    ' Create a dummy instance HoldingsFromAaRow to use for column positions, etc.
    Set ARow = New PrivateFileRow
    Call ARow.InitializeWithRange(ThisWorkbook.Worksheets("TempComputation").Range("A1").Resize(1, ARow.GetNumberOfColumns))
    
    ' Set the mapping from Tradar's table column position to the DB's table column position
    Let SourceToTargetColumnIndexMapping = [{2, 2; 4, 3; 5, 4; 6, 5; 8, 6; 9, 7; 10, 8; 11, 9; 12, 10; 13, 11; 14, 12; 19, 13; 20, 14; 21, 15; 22, 16; 24, 17; 28, 18; 31, 19; 32, 20; 33, 21; 34, 22; 35, 23}]
    
    ' Open the workbook holding the Tradar view and extract the data
    Set sourceSht = Application.Workbooks.Open(Filename:=AFileName, UpdateLinks:=False, ReadOnly:=True).Worksheets(1)
    
    ' Set aRange to the range holding the underlying data
    Let LastRowIndex = LastNonBlankRowIndexInColumnNotEqualTo(aColumnRange:=sourceSht.Range("A1").EntireColumn, _
                                                              ComparisonScalars:="")
    Set SourceRange = sourceSht.Range("A1").Resize(LastRowIndex, Me.GetNumberOfColumns)
    
    ' Create new workbook to hold the database underlying the DB contents
    Set targetSht = Application.Workbooks.Add.Worksheets(1)
    Call RemoveAllOtherWorksheets(targetSht)
    
    ' Format as text target columns holding 7-char SEDOLS
    Let targetSht.Cells(1, ARow.GetSedol.Column).EntireColumn.NumberFormat = "@"
    Let targetSht.Cells(1, ARow.GetInvestableSedol.Column).EntireColumn.NumberFormat = "@"
    
    ' Move columns from source range to the expected location in the target worksheet
    For i = LBound(SourceToTargetColumnIndexMapping, 1) To UBound(SourceToTargetColumnIndexMapping, 1)
        Set TargetRange = targetSht.Cells(1, SourceToTargetColumnIndexMapping(i, 2)).Resize(SourceRange.Rows.Count, 1)
        Let TargetRange.Value2 = SourceRange.Columns(SourceToTargetColumnIndexMapping(i, 1)).Value2
    Next i
    
    ' Set targetRange before initializing this instance with targetRage
    Set TargetRange = targetSht.Range("A1").Resize(SourceRange.Rows.Count, ARow.GetNumberOfColumns)

    ' Close the source workbook now that all data has been transfered to the target workbook
    Call sourceSht.Parent.Close(SaveChanges:=False)

    ' Set UnderlyingWorkbook
    Set Me.UnderlyingWorkbook = TargetRange.Worksheet.Parent

    ' Fill in date and time column headers for targetRange
    Let TargetRange.Worksheet.Range("A1").Resize(1, GetArrayLength(ARow.GetHeaders)).Value2 = ARow.GetHeaders
        
    ' Fill in the date and time columns with today's date and the current time
    If TargetRange.Rows.Count > 1 Then
        If IsMissing(OptionalDate) Then
            Let TheDate = CLng(Right(StringSplit(AFileName, ".", GetArrayLength(Split(AFileName, ".")) - 1), 8))
        Else
            Let TheDate = CLng(OptionalDate)
        End If
        
        ' Insert NULL in the classification column since Excel private files do not have classifications (by design)
        Let TargetRange.Resize(TargetRange.Rows.Count - 1, 1).Offset(1, ARow.GetClassification.Column - 1).Value2 = "NULL"
    
        ' Insert current date
        Let TargetRange.Resize(TargetRange.Rows.Count - 1, 1).Offset(1, ARow.GetDate.Column - 1).Value2 = TheDate
    
        ' Initialize with the target in the target worksheet
        Call Me.InitializeWithRange(TargetRange)
    End If
End Sub

' This injects this instance into the DB by making a copy of the private file instance and saving it as a TSV file.
' Then, this is injected using LOAD DATA LOCAL INFILE...
Public Sub InjectIntoMySql(ServerName As String, DatabaseName As String, TableName As String, UserName As String, Password As String)
    Dim TempTsvFileName As String
    Dim wsht As Worksheet
    Dim wbk As Workbook
    Dim NumberOfWorksheets As Integer
    
    If Me.GetDbLength > 0 Then
        ' Make a copy of the worksheet underlying this instance in a new workbook
        Set wbk = Application.Workbooks.Add
        Set wsht = wbk.Worksheets(wbk.Worksheets.Count)
        Let NumberOfWorksheets = wbk.Worksheets.Count
        Call Me.GetRange.Worksheet.Copy(After:=wsht)
        Set wsht = wsht.Parent.Worksheets(NumberOfWorksheets + 1)
        Call RemoveAllOtherWorksheets(wsht)
        Call wsht.Range("1:1").Delete
        Call wsht.Range(Cells(1, Me.GetClassifications.Column + 1), Cells(1, wsht.Columns.Count)).EntireColumn.ClearContents
        
        ' Set the filename
        Let TempTsvFileName = "c:\users\" & ReturnUserName() & "\Desktop\TempTsv.tsv"
        
        ' Replace blank cells by NULL in those columns that are allowed to be NULL
        Call wsht.Range("R2").Resize(Me.GetDbLength - 1, 4).Replace(What:="", Replacement:="NULL")
        
        ' Save temp TSV file and close the file
        Call wsht.Parent.SaveAs(Filename:=TempTsvFileName, FileFormat:=xlText)
        Call wsht.Parent.Close(SaveChanges:=False)
        
        ' Inject private instance into DB
        Call RunQuery("LOAD DATA LOCAL INFILE '" & Replace(TempTsvFileName, "\", "\\") & "' INTO TABLE `" & DatabaseName & "`.`" & TableName & "` FIELDS TERMINATED BY '\t' OPTIONALLY ENCLOSED BY '\""' LINES TERMINATED BY '\r\n';", ServerName, DatabaseName, UserName, Password)
        
        ' Delete temporary file
        Kill TempTsvFileName
    End If
End Sub

' Custom Initializer.
Public Sub InitializeWithMySql(ServerName As String, DatabaseName As String, TableName As String, TheDate As Long, _
                               UserName As String, Password As String)
    Dim TheQuery As String
    Dim TheResults As Variant
    Dim ARow As New PrivateFileRow
       
    ' Set the query to get the latest portfolio on this date
    Let TheQuery = "SELECT * FROM `" & DatabaseName & "`.`" & TableName & "` WHERE "
    Let TheQuery = TheQuery & "`date` = " & TheDate & ";"
    Let TheResults = ConnectAndSelect(TheQuery, DatabaseName, ServerName, UserName, Password)
            
    ' Create a new workbook to hold data and dump data in a worksheet
    Set Me.UnderlyingWorkbook = Application.Workbooks.Add
    Call RemoveAllOtherWorksheets(UnderlyingWorkbook.Worksheets(1))
    
    ' Instantiate aRow
    Call ARow.InitializeWithRange(Me.UnderlyingWorkbook.Worksheets(1).Range("A1").Resize(1, ARow.GetNumberOfColumns))
    
    ' Format as text target columns holding 7-char SEDOLS
    Let Me.UnderlyingWorkbook.Worksheets(1).Cells(1, ARow.GetSedol.Column).EntireColumn.NumberFormat = "@"
    Let Me.UnderlyingWorkbook.Worksheets(1).Cells(1, ARow.GetInvestableSedol.Column).EntireColumn.NumberFormat = "@"
    
    Call DumpInTempPositionWithoutFirstClearing(TheResults, UnderlyingWorkbook.Worksheets(1).Range("A1"))
    
    ' Initialize with the data now held in the range
    Call Me.InitializeWithRange(UnderlyingWorkbook.Worksheets(1).Range("A1").CurrentRegion)
End Sub

' This initializes all row indices dictionary (e.g. sedol, bloomberg ticker, and primary blommberg ticker).
' This sub requires RangeDb to have been set.
Public Sub InitializeRowIndicesDictionaries()
    Dim i As Integer
    Dim ARow As PrivateFileRow
    
    ' Instantiate a new, empty dictionary
    Set Me.SedolRowIndicesDictionary = New Dictionary
    Set Me.BlooombergTickerRowIndicesDictionary = New Dictionary
    Set Me.PrimaryBloombergTickerRowIndicesDictionary = New Dictionary
    
    If Me.GetDbLength > 0 Then
        Set ARow = Me.GetIthRow(1)
    
        For i = 1 To Me.GetDbLength
            If Not Me.SedolRowIndicesDictionary.Exists(Me.GetRange(i, ARow.GetSedol.Column - ARow.GetRange.Cells(1, 1).Column + 1).Value2) Then
                Call Me.SedolRowIndicesDictionary.Add(Key:=Me.GetRange(i, ARow.GetSedol.Column - ARow.GetRange.Cells(1, 1).Column + 1).Value2, Item:=i)
            End If

            If Not Me.BlooombergTickerRowIndicesDictionary.Exists(Me.GetRange(i, ARow.GetBloombergTicker.Column - ARow.GetRange.Cells(1, 1).Column + 1).Value2) Then
                Call Me.BlooombergTickerRowIndicesDictionary.Add(Key:=Me.GetRange(i, ARow.GetBloombergTicker.Column - ARow.GetRange.Cells(1, 1).Column + 1).Value2, Item:=i)
            End If
            
            If Not Me.PrimaryBloombergTickerRowIndicesDictionary.Exists(Me.GetRange(i, ARow.GetPrimaryBloombergTicker.Column - ARow.GetRange.Cells(1, 1).Column + 1).Value2) Then
                Call Me.PrimaryBloombergTickerRowIndicesDictionary.Add(Key:=Me.GetRange(i, ARow.GetPrimaryBloombergTicker.Column - ARow.GetRange.Cells(1, 1).Column + 1).Value2, Item:=i)
            End If
        Next i
    End If
End Sub

' This initializes the Sedols row indices dictionary.  This sub requires RangeDb to have been set.
Public Sub InitializeSedolRowIndicesDictionary()
    Dim i As Integer
    Dim ARow As PrivateFileRow
    
    ' Instantiate a new, empty dictionary
    Set Me.SedolRowIndicesDictionary = New Dictionary
    
    If Me.GetDbLength > 0 Then
        Set ARow = Me.GetIthRow(1)
    
        For i = 1 To Me.GetRange.Rows.Count
            If Not Me.SedolRowIndicesDictionary.Exists(Me.GetRange(i, ARow.GetBloombergTicker.Column - ARow.GetRange.Cells(1, 1).Column + 1).Value2) Then
                Call Me.SedolRowIndicesDictionary.Add(Key:=Me.GetRange(i, ARow.GetBloombergTicker.Column - ARow.GetRange.Cells(1, 1).Column + 1).Value2, Item:=i)
            End If
        Next i
    End If
End Sub

' This initializes the primary Bloomberg tickers row indices dictionary.  This sub requires RangeDb to have been set.
Public Sub InitializePrimaryBloombergTickerRowIndicesDictionary()
    Dim ARow As PrivateFileRow
    Dim i As Long
    
    ' Instantiate a new, empty dictionary
    Set Me.PrimaryBloombergTickerRowIndicesDictionary = New Dictionary
    
    If Me.GetDbLength > 0 Then
        Set ARow = Me.GetIthRow(1)
    
        For i = 1 To Me.GetRange.Rows.Count
            If Not Me.PrimaryBloombergTickerRowIndicesDictionary.Exists(Me.GetRange(i, ARow.GetPrimaryBloombergTicker.Column - ARow.GetRange.Cells(1, 1).Column + 1).Value2) Then
                Call Me.PrimaryBloombergTickerRowIndicesDictionary.Add(Key:=Me.GetRange(i, ARow.GetPrimaryBloombergTicker.Column - ARow.GetRange.Cells(1, 1).Column + 1).Value2, Item:=i)
            End If
        Next i
    End If
End Sub

' This initializes the primary Bloomberg tickers row indices dictionary.  This sub requires RangeDb to have been set.
Public Sub InitializeBloombergTickerRowIndicesDictionary()
    Dim ARow As PrivateFileRow
    Dim i As Long
    
    ' Instantiate a new, empty dictionary
    Set Me.BloombergTickerRowIndicesDictionary = New Dictionary
    
    If Me.GetDbLength > 0 Then
        Set ARow = Me.GetIthRow(1)
    
        For i = 1 To Me.GetRange.Rows.Count
            If Not Me.BloombergTickerRowIndicesDictionary.Exists(Me.GetRange(i, ARow.GetBloombergTicker.Column - ARow.GetRange.Cells(1, 1).Column + 1 - ARow.GetRange.Cells(1, 1).Column + 1).Value2) Then
                Call Me.BlooombergTickerRowIndicesDictionary.Add(Key:=Me.GetRange(i, ARow.GetBloombergTicker.Column).Value2, Item:=i)
            End If
        Next i
    End If
End Sub

' This initializes the country-to-region mapping instance that is private to this object
Public Sub InitializeCountryToRegionMap(AFileName As String)
    ' Instantiate a new CountryCodeRegionalCodeMap
    Set Me.CountryRegionMapping = Nothing
    Set Me.CountryRegionMapping = New CountryCodeRegionalCodeMap

    ' Initialize the map with the data in file aFileName
    Call Me.CountryRegionMapping.InitializeWithFile(AFileName)
End Sub

' This function initializes a dictionary indexing the GICS code-to-MSCI sub-industry code
Public Sub InitializeGicsToMsciSubIndustryCodeMappingWithFile(AFileName As String)
    Dim MsciWorksheet As Worksheet
    Dim TheRange As Range
    Dim NumberOfGicsCodes As Integer
    
    ' Open workbook containing map and save reference. Set reference to worksheet holding the data
    Set MsciWorksheet = Application.Workbooks.Open(AFileName).Worksheets(1)
    
    ' Extract range holding the GISC-to-MSCI translation map
    Set TheRange = MsciWorksheet.Range("A1").CurrentRegion
    
    ' Determine the number of SEDOLs to avoid problems stemming from various
    ' columns having different lengths.
    Let NumberOfGicsCodes = TheRange.Columns(1).Cells.Count

    ' Repoint TheRange to the worksheet we created (e.g. PrivateFileWorksheet)
    Set TheRange = MsciWorksheet.Range("A1").CurrentRegion
    
    ' Pick only as many columns as there are SEDOLs plus the header row.
    Call Me.InitializeGicsToMsciSubIndustryCodeMappingWithRange(TheRange.Resize(NumberOfGicsCodes, TheRange.Columns.Count))
    
    ' Close workbook holding the map
    Call MsciWorksheet.Parent.Close
End Sub

' This function initializes the dictionary of GICS codes-to-MSCI sub-industry codes
Public Sub InitializeGicsToMsciSubIndustryCodeMappingWithRange(ARange As Range)
    Dim GicsCodes As Variant
    Dim MsciSubIndustryCodes As Variant
    Dim NumberOfGicsCodes As Integer
    Dim i As Integer
    
    ' Extract the GICS codes array
    Let GicsCodes = ConvertTo1DArray(ARange.Offset(1, 0).Resize(ARange.Rows.Count - 1, 1).Value2)
    Let MsciSubIndustryCodes = ConvertTo1DArray(ARange.Offset(1, 2).Resize(ARange.Rows.Count - 1, 1).Value2)
    
    ' Compute the number of GICS codes in the map
    Let NumberOfGicsCodes = UBound(GicsCodes)
    
    ' Instantiate a new dictionary
    Set Me.GicsToMsciSubIndustryCodeMapping = Nothing
    Set Me.GicsToMsciSubIndustryCodeMapping = New Dictionary
    
    ' Insert map into dictionary
    For i = 1 To NumberOfGicsCodes
        Call Me.GicsToMsciSubIndustryCodeMapping.Add(GicsCodes(i), MsciSubIndustryCodes(i))
    Next i
End Sub

' This function maps GICS codes to MSCI sub-industry codes for all securities and insert them
' into the underlying range
Public Sub ComputeMsciSubIndustryCodes()
    Dim MsicSubIndustryCodes() As String
    Dim i As Long
    
    ' Pre-allocate the array of regional codes
    ReDim MsicSubIndustryCodes(GetDbLength)
    
    ' Get the regional code of all securities
    For i = 1 To GetDbLength
        Let MsicSubIndustryCodes(i) = Me.GetGicsToMsciSubIndustryCodeMapping.Item(Me.GetIthSecurity(i).GetIndustryCode.Value)
    Next i
    
    ' Dump the column array of regional codes in the expected column
    Let Me.GetGicsCodes.Value2 = Application.Transpose(MsicSubIndustryCodes)
End Sub

' This initializes SizeThresholds
Public Sub InitializeSizeThresholds()
    Dim NaMarketCap As Double
    Dim EurMarketCap As Double
    Dim RwMarketCap As Double
    
    ' Instantiate a new dictionary
    Set SizeThresholds = New Dictionary
    
    ' Compute total US market cap
    Let NaMarketCap = Application.SumIf(Me.GetRegions, "USA", Me.GetUsdMarketCaps)
    Let EurMarketCap = Application.SumIf(Me.GetRegions, "EUR", Me.GetUsdMarketCaps)
    Let RwMarketCap = Application.SumIf(Me.GetRegions, "RW", Me.GetUsdMarketCaps)
    
    ' Add size thresholds to the dictionary
    Call Me.SizeThresholds.Add("USA", NaMarketCap * 0.0009)
    Call Me.SizeThresholds.Add("EUR", EurMarketCap * 0.001)
    Call Me.SizeThresholds.Add("RW", RwMarketCap * 0.001)
End Sub

' This initializes StyleThresholds
Public Sub InitializeStyleThresholds()
    Dim i As Integer
    Dim j As Integer
    Dim SectorRegionCoordinates() As String
    Dim Regions As Variant
    Dim Sectors As Variant
    Dim MedianPe As Double
    
    ' Pre-allocate Sector/Region threshold array (e.g. 3 regions and 12 sectors)
    ReDim SectorRegionCoordinates(UBound(PalladyneRegions()) * UBound(PalladyneSectors()))

    ' Set regions arraypaz
    Let Regions = PalladyneRegions()
    
    ' Set sectors array
    Let Sectors = PalladyneSectors()

    ' Temporarily insert region/sector coordinates for each security
    Let Me.GetSizes.FormulaR1C1 = "=R[0]C[" & -(Me.GetSizes.Column - Me.GetRegions.Column) & "] &" & Chr(34) & "-" & Chr(34) & "& R[0]C[" & -(Me.GetSizes.Column - Me.GetPiamSectors.Column) & "]"

    ' Instantiate a new dictionary
    Set Me.StyleThresholds = New Dictionary
    
    ' Create sector/region coordinate pairs
    For i = 1 To UBound(PalladyneSectors())
        For j = 1 To UBound(PalladyneRegions())
            Let SectorRegionCoordinates(i + (j - 1) * 12) = Regions(j) & "-" & Sectors(i)
        Next j
    Next i
    
    ' Compute the media PE of each region/sector pair and add them to the dictionary
    For i = 1 To UBound(Regions) * UBound(Sectors)
        Let Me.GetRange(i, Me.GetRange.Columns.Count).Offset(0, 1).FormulaArray = "=median(if(" & Me.GetSizes.Address & "=" & Chr(34) & SectorRegionCoordinates(i) & Chr(34) & ", " & Me.GetFy1Pes.Address & ", " & Chr(34) & Chr(34) & "))"
        Let MedianPe = Me.GetRange(i, Me.GetNumberOfColumns).Offset(0, 1).Value
        Call Me.StyleThresholds.Add(SectorRegionCoordinates(i), MedianPe)
    Next i
    
    ' Clear the temp columns we inserted to compute the style thresholds
    Call Me.GetSizes.Cells.ClearContents
    Call Me.GetRange.Columns(Me.GetNumberOfColumns).Offset(0, 1).Cells.ClearContents
End Sub

' This function computes and sets the regions corresponding to the securities in the private file
Public Sub ComputeRegionalCodes()
    Dim RegionalCodes() As String
    Dim i As Long
    
    ' Pre-allocate the array of regional codes
    ReDim RegionalCodes(Me.GetDbLength)
    
    ' Get the regional code of all securities
    For i = 1 To GetDbLength
        Let RegionalCodes(i) = Me.CountryRegionMapping.GetRegionalCode(Me.GetIthRow(i).GetCountryOfDomicile.Value)
    Next i
    
    ' Return the array of regional codes
    Let Me.GetRegions.Value = Application.Transpose(RegionalCodes)
End Sub

' This returns an array with the size classifications of all securities
Public Sub ComputesSizes()
    Dim ResultArray() As String
    Dim i As Long
    
    ' Pre-allocate the array of regional codes
    ReDim ResultArray(Me.GetDbLength)
    
    ' Get the regional code of all securities
    For i = 1 To Me.GetDbLength
        If Me.GetIthRow(i).GetUsdMarketCap.Value2 <= Me.SizeThresholds(Me.GetIthRow(i).GetRegion.Value2) Then
            Let ResultArray(i) = "S"
        Else
            Let ResultArray(i) = "L"
        End If
    Next i
    
    ' Populate secutities` size classifications
    Let GetSizes.Value = Application.Transpose(ResultArray)
End Sub

' This computes and populates the size classifications of all securities
Public Sub ComputesStyles()
    Dim ResultArray() As String
    Dim i As Long
    
    ' Pre-allocate the array of regional codes
    ReDim ResultArray(GetDbLength)
    
    ' Get the regional code of all securities
    For i = 1 To GetDbLength
        If Me.GetIthRow(i).GetFy1Pe.Value2 <= Me.StyleThresholds(Me.GetIthRow(i).GetRegion.Value2 & "-" & Me.GetIthRow(i).GetPiamSector.Value2) Then
            Let ResultArray(i) = "V"
        Else
            Let ResultArray(i) = "G"
        End If
    Next i
    
    ' Populate securities` style classifications
    Let GetStyles.Value = Application.Transpose(ResultArray)
End Sub

' This function drops securities with non-positive and extremely high PEs (e.g. PE > 120) and
' returns a 2D array with the deleted rows.
Public Function FilterOutUnacceptablePes() As Variant
    Dim Headers As Variant
    Dim TempValues As Variant
    Dim ReturnArray As Variant
    Dim ARow As PrivateFileRow
    
    ' Instantiate and initialize PrivateFileRow
    Set ARow = New PrivateFileRow
    Call ARow.InitializeWithRange(ThisWorkbook.Worksheets("TempComputation").Range("A1").Resize(1, ARow.GetNumberOfColumns))
    
    ' Clear all previous autofilter conditions
    Let Me.GetRange.Worksheet.AutoFilterMode = False
    
    ' Store header row
    Let Headers = GetRange.Rows(1).Offset(-1, 0).Cells.Value
    
    ' Do autofilters to select assets to delete
    Call Me.GetRange.Offset(-1, 0).Resize(Me.GetDbLength + 1, Me.GetNumberOfColumns).AutoFilter(Field:=Me.GetFy1Pes.Column, Criteria1:="<=0", Operator:=xlOr, Criteria2:=">120")
    
    ' Extract the rows to be deleted before they are deleted. They must be copied and pasted to transform
    ' the set into a contiguous range of cells from a possibly disconnected set or rows.
    Call ThisWorkbook.Worksheets("TempComputation").UsedRange.ClearContents
    Call Me.GetRange.Worksheet.AutoFilter.Range.SpecialCells(xlCellTypeVisible).Copy
    Call ThisWorkbook.Worksheets("TempComputation").Range("A1").PasteSpecial(Paste:=xlPasteValues)
    Let ReturnArray = ThisWorkbook.Worksheets("TempComputation").Range("A1").CurrentRegion.Value2
    Call ThisWorkbook.Worksheets("TempComputation").UsedRange.ClearContents
    
    ' Select all visible rows to delete them and delete them
    Call Me.GetRange.Worksheet.AutoFilter.Range.SpecialCells(xlCellTypeVisible).Delete(Shift:=xlUp)
    
    ' Clear all previous autofilter conditions
    Let GetRange.Worksheet.AutoFilterMode = False
    
    ' Store securities left after the autofilter
    Let TempValues = Me.GetRange.Worksheet.Range("A1").Resize(Me.GetDbLength, Me.GetNumberOfColumns).Value
    ' Clear region before pasting assets left over
    Call Me.GetRange.Worksheet.UsedRange.ClearContents
    ' Paste headers
    Let Me.GetRange.Rows(1).Value = ARow.GetHeaders
    ' Paste assets left over
    Let Me.GetRange.Worksheet.Range("A2").Resize(UBound(TempValues, 1), UBound(TempValues, 2)).Value = TempValues
    
    ' Repoint asset to new region
    Call InitializeWithRange(Me.GetRange.Worksheet.Range("A1").Resize(Me.GetDbLength, ARow.GetNumberOfColumns))
    
    ' Return rows that were deleted
    Let FilterOutUnacceptablePes = ReturnArray
End Function

' This function drops securities with missing countries and returns a 2D array with the deleted rows
Public Function FilterOutMissingCountries() As Variant
    Dim Headers As Variant
    Dim TempValues As Variant
    Dim NumberOfColumns As Integer
    Dim ReturnArray As Variant
    
    ' Store the number of columns in the object's underlying range
    Let NumberOfColumns = GetRange.Columns.Count
    
    ' Clear all previous autofilter conditions
    Let GetRange.Worksheet.AutoFilterMode = False
    
    ' Store header row
    Let Headers = GetRange.Rows(1).Offset(-1, 0).Cells.Value
    
    ' Do autofilters to select assets to delete
    GetRange.Offset(-1, 0).Resize(Me.GetDbLength + 1, GetRange.Columns.Count).AutoFilter _
        Field:=Me.GetCountryOfDomiciles.Column, Criteria1:="="

    ' Extract the rows to be deleted before they are deleted. They must be copied and pasted to transform
    ' the set into a contiguous range of cells from a possibly disconnected set or rows.
    Call ThisWorkbook.Worksheets("TempComputation").UsedRange.ClearContents
    Call GetRange.Worksheet.AutoFilter.Range.SpecialCells(xlCellTypeVisible).Copy
    ThisWorkbook.Worksheets("TempComputation").Range("A1").PasteSpecial Paste:=xlPasteValues
    Let ReturnArray = ThisWorkbook.Worksheets("TempComputation").Range("A1").CurrentRegion.Value
    Call ThisWorkbook.Worksheets("TempComputation").UsedRange.ClearContents

    ' Select all visible rows to delete them and delete them
    GetRange.Worksheet.AutoFilter.Range.SpecialCells(xlCellTypeVisible).Delete Shift:=xlUp
    
    ' Clear all previous autofilter conditions
    Let GetRange.Worksheet.AutoFilterMode = False
    
    ' Store assets left
    Let TempValues = GetRange.Worksheet.Range("A1").Resize(GetRange.Worksheet.Range("A1").CurrentRegion.Rows.Count, NumberOfColumns).Value
    ' Clear region before pasting assets left over
    Call GetRange.Worksheet.UsedRange.ClearContents
    ' Paste headers
    Let GetRange.Rows(1).Value = Headers
    ' Paste assets left over
    Let GetRange.Worksheet.Range("A2").Resize(UBound(TempValues, 1), UBound(TempValues, 2)).Value = TempValues
    
    ' Repoint asset to new region
    Call InitializeWithRange(GetRange.Worksheet.Range("A1").Resize(GetRange.Worksheet.Range("A1").CurrentRegion.Columns(1).Cells.Count, _
                                                                   GetRange.Worksheet.Range("A1").CurrentRegion.Rows(1).Cells.Count))
    
    ' Return rows that were deleted
    Let FilterOutMissingCountries = ReturnArray
End Function

' This function returns a 2D array (Nx2) with the number of securities in each cell.
' The first column contains the cellular coordinates and the second the the weights.
Public Function ComputeCellularSecurityCounts() As Variant
    Dim Sectors As Variant
    Dim Regions As Variant
    Dim Styles As Variant
    Dim Sizes As Variant
    Dim Sector As Variant
    Dim Region As Variant
    Dim Style As Variant
    Dim Size As Variant
    Dim CoordinateTuples As Dictionary
    Dim ColPos As Integer
    Dim Counts As Double
    Dim CurrentCoordinateTuple As String

    ' Create a new dictionary instance
    Set CoordinateTuples = New Dictionary

    ' Set regions, sectors, styles, and sizes arrays
    Let Regions = PalladyneRegions()
    Let Sectors = PalladyneSectors()
    Let Styles = PalladyneStyles()
    Let Sizes = PalladyneSizes()
    
    ' Compute column position of first column to the right of the underlying range
    Let ColPos = Me.GetRange.Columns(Me.GetNumberOfColumns).Offset(0, 1).Column
    
    ' Insert column with coordinate tuples to the right of the underlying range of the PrivateFile instance
    Let Me.GetRange.Columns(Me.GetNumberOfColumns).Offset(0, 1).FormulaR1C1 = _
        "=R[0]C[" & -(ColPos - Me.GetPiamSectors.Column) & "]&" & Chr(34) & "-" & Chr(34) & _
         "&R[0]C[" & -(ColPos - Me.GetRegions.Column) & "]&" & Chr(34) & "-" & Chr(34) & _
         "&R[0]C[" & -(ColPos - Me.GetSizes.Column) & "]&" & Chr(34) & "-" & Chr(34) & _
         "&R[0]C[" & -(ColPos - Me.GetStyles.Column) & "]"

    ' Create array of all coordinate tuples and store the weight of each cell
    For Each Sector In Sectors
        For Each Region In Regions
            For Each Size In Sizes
                For Each Style In Styles
                    ' Create the string representing the current coordinate tuple
                    Let CurrentCoordinateTuple = Sector & "-" & Region & "-" & Size & "-" & Style
                
                    ' Compute the weight of the cell with the current coordinate 4-tuple
                    Let Counts = Application.CountIf(Me.GetRange.Range("A1").Offset(0, ColPos - 1).Resize(GetRange.Rows.Count, 1), _
                                                   "=" & CurrentCoordinateTuple)
                                        
                    ' Store this weight in the dictionary
                    CoordinateTuples.Add CurrentCoordinateTuple, Counts
                Next Style
            Next Size
        Next Region
    Next Sector
    
    ' Clear helper columns
    Call Me.GetRange.Columns(Me.GetNumberOfColumns).Offset(0, 1).Resize(GetRange.Rows.Count, 1).Cells.ClearContents
    
    ' Return dictionary object containing weights
    Let ComputeCellularSecurityCounts = Application.Transpose(Array(CoordinateTuples.Keys, CoordinateTuples.Items))
End Function

' This function returns a 2D array (Nx2) with the straight, market-relative weights of each cell.
' The first column contains the cellular coordinates and the second the the weights.
Public Function ComputeCellularWeights() As Variant
    Dim Sectors As Variant
    Dim Regions As Variant
    Dim Styles As Variant
    Dim Sizes As Variant
    Dim Sector As Variant
    Dim Region As Variant
    Dim Style As Variant
    Dim Size As Variant
    Dim CoordinateTuples As Dictionary
    Dim ColPos As Integer
    Dim Weight As Double
    Dim TotalUniverseMarketCap As Double
    Dim CurrentCoordinateTuple As String

    ' Create a new dictionary instance
    Set CoordinateTuples = New Dictionary

    ' Set regions, sectors, styles, and sizes arrays
    Let Regions = PalladyneRegions()
    Let Sectors = PalladyneSectors()
    Let Styles = PalladyneStyles()
    Let Sizes = PalladyneSizes()
    
    ' Compute column position of first column to the right of the underlying range
    Let ColPos = Me.GetRange.Columns(Me.GetNumberOfColumns).Offset(0, 1).Column
    
    ' Insert column with coordinate tuples to the right of the underlying range of the PrivateFile instance
    Let GetRange.Columns(Me.GetNumberOfColumns).Offset(0, 1).FormulaR1C1 = _
        "=R[0]C[" & -(ColPos - Me.GetPiamSectors.Column) & "]&" & Chr(34) & "-" & Chr(34) & _
         "&R[0]C[" & -(ColPos - Me.GetRegions.Column) & "]&" & Chr(34) & "-" & Chr(34) & _
         "&R[0]C[" & -(ColPos - Me.GetSizes.Column) & "]&" & Chr(34) & "-" & Chr(34) & _
         "&R[0]C[" & -(ColPos - Me.GetStyles.Column) & "]"

    ' Compute the total market cap of the universe
    Let TotalUniverseMarketCap = Application.Sum(Me.GetUsdMarketCaps.Value2)

    ' Create array of all coordinate tuples and store the weight of each cell
    For Each Sector In Sectors
        For Each Region In Regions
            For Each Size In Sizes
                For Each Style In Styles
                    ' Create the string representing the current coordinate tuple
                    Let CurrentCoordinateTuple = Sector & "-" & Region & "-" & Size & "-" & Style
                
                    ' Compute the weight of the cell with the current coordinate 4-tuple
                    Let Weight = Application.SumIf(Me.GetRange.Range("A1").Offset(0, ColPos - 1).Resize(Me.GetRange.Rows.Count, 1), _
                                                   "=" & CurrentCoordinateTuple, _
                                                   Me.GetUsdMarketCaps) _
                                 / TotalUniverseMarketCap
                                        
                    ' Store this weight in the dictionary
                    CoordinateTuples.Add CurrentCoordinateTuple, Weight
                Next Style
            Next Size
        Next Region
    Next Sector
    
    ' Clear helper columns
    Call Me.GetRange.Columns(Me.GetNumberOfColumns).Offset(0, 1).Resize(GetRange.Rows.Count, 1).Cells.ClearContents
    
    ' Return dictionary object containing weights
    Let ComputeCellularWeights = Application.Transpose(Array(CoordinateTuples.Keys, CoordinateTuples.Items))
End Function

' This function returns a 2D array (Nx2) with the market caps of each cell.
' The first column contains the cellular coordinates and the second the the weights.
Public Function ComputeCellularMarketCaps() As Variant
    Dim Sectors As Variant
    Dim Regions As Variant
    Dim Styles As Variant
    Dim Sizes As Variant
    Dim Sector As Variant
    Dim Region As Variant
    Dim Style As Variant
    Dim Size As Variant
    Dim CoordinateTuples As Dictionary
    Dim ColPos As Integer
    Dim MarketCap As Double
    Dim CurrentCoordinateTuple As String

    ' Create a new dictionary instance
    Set CoordinateTuples = New Dictionary

    ' Set regions, sectors, styles, and sizes arrays
    Let Regions = PalladyneRegions()
    Let Sectors = PalladyneSectors()
    Let Styles = PalladyneStyles()
    Let Sizes = PalladyneSizes()
    
    ' Compute column position of first column to the right of the underlying range
    Let ColPos = GetRange.Columns(GetRange.Columns.Count).Offset(0, 1).Column
    
    ' Insert column with coordinate tuples to the right of the underlying range of the PrivateFile instance
    Let GetRange.Columns(GetRange.Columns.Count).Offset(0, 1).FormulaR1C1 = _
        "=R[0]C[" & -(ColPos - GetSectors.Column) & "]&" & Chr(34) & "-" & Chr(34) & _
         "&R[0]C[" & -(ColPos - GetRegions.Column) & "]&" & Chr(34) & "-" & Chr(34) & _
         "&R[0]C[" & -(ColPos - GetSizes.Column) & "]&" & Chr(34) & "-" & Chr(34) & _
         "&R[0]C[" & -(ColPos - GetStyles.Column) & "]"

    ' Create array of all coordinate tuples and store the weight of each cell
    For Each Sector In Sectors
        For Each Region In Regions
            For Each Size In Sizes
                For Each Style In Styles
                    ' Create the string representing the current coordinate tuple
                    Let CurrentCoordinateTuple = Sector & "-" & Region & "-" & Size & "-" & Style
                
                    ' Compute the weight of the cell with the current coordinate 4-tuple
                    Let MarketCap = Application.SumIf(GetRange.Range("A1").Offset(0, ColPos - 1).Resize(GetRange.Rows.Count, 1), _
                                                      "=" & CurrentCoordinateTuple, _
                                                      GetMarketCaps)
                                        
                    ' Store this weight in the dictionary
                    CoordinateTuples.Add CurrentCoordinateTuple, MarketCap
                Next Style
            Next Size
        Next Region
    Next Sector
    
    ' Clear helper columns
    GetRange.Columns(GetRange.Columns.Count).Offset(0, 1).Resize(GetRange.Rows.Count, 1).Cells.ClearContents
    
    ' Return dictionary object containing weights
    Let ComputeCellularMarketCaps = Application.Transpose(Array(CoordinateTuples.Keys, CoordinateTuples.Items))
End Function

' This function returns a 2D array (Nx2) with the straight, market-relative weights of each
' region. The first column contains the regional coordinates and the second the weights.
Public Function ComputeRegionalWeights() As Variant
    Dim Regions As Variant
    Dim Region As Variant
    Dim CoordinateTuples As Dictionary
    Dim ColPos As Integer
    Dim Weight As Double
    Dim TotalUniverseMarketCap As Double
    Dim CurrentCoordinateTuple As String

    ' Create a new dictionary instance
    Set CoordinateTuples = New Dictionary

    ' Set regions, sectors, styles, and sizes arrays
    Let Regions = PalladyneRegions()
    
    ' Compute column position of first column to the right of the underlying range
    Let ColPos = Me.GetRange.Columns(Me.GetNumberOfColumns).Offset(0, 1).Column
    
    ' Insert column with coordinate tuples to the right of the underlying range of the PrivateFile instance
    Let GetRange.Columns(Me.GetNumberOfColumns).Offset(0, 1).FormulaR1C1 = _
        "=R[0]C[" & -(ColPos - Me.GetRegions.Column) & "]"

    ' Compute the total market cap of the universe
    Let TotalUniverseMarketCap = Application.Sum(Me.GetUsdMarketCaps.Value2)

    ' Create array of all coordinate tuples and store the weight of each cell
    For Each Region In Regions
        ' Create the string representing the current coordinate tuple
        Let CurrentCoordinateTuple = Region

        ' Compute the weight of the cell with the current coordinate 4-tuple
        Let Weight = Application.SumIf(Me.GetRange.Range("A1").Offset(0, ColPos - 1).Resize(Me.GetRange.Rows.Count, 1), _
                                       "=" & CurrentCoordinateTuple, _
                                       Me.GetUsdMarketCaps) _
                     / TotalUniverseMarketCap
                                        
        ' Store this weight in the dictionary
        CoordinateTuples.Add CurrentCoordinateTuple, Weight
    Next Region
    
    ' Clear helper columns
    Call Me.GetRange.Columns(Me.GetNumberOfColumns).Offset(0, 1).Resize(GetRange.Rows.Count, 1).Cells.ClearContents
    
    ' Return dictionary object containing weights
    Let ComputeRegionalWeights = Application.Transpose(Array(CoordinateTuples.Keys, CoordinateTuples.Items))
End Function

' This function returns a 2D array (Nx2) with the regional security counts of each
' region. The first column contains the regional coordinates and the second the counts.
Public Function ComputeRegionalSecurityCounts() As Variant
    Dim Regions As Variant
    Dim Region As Variant
    Dim CoordinateTuples As Dictionary
    Dim ColPos As Integer
    Dim Counts As Double
    Dim CurrentCoordinateTuple As String

    ' Create a new dictionary instance
    Set CoordinateTuples = New Dictionary

    ' Set regions, sectors, styles, and sizes arrays
    Let Regions = PalladyneRegions()
    
    ' Compute column position of first column to the right of the underlying range
    Let ColPos = Me.GetRange.Columns(Me.GetNumberOfColumns).Offset(0, 1).Column
    
    ' Insert column with coordinate tuples to the right of the underlying range of the PrivateFile instance
    Let Me.GetRange.Columns(Me.GetNumberOfColumns).Offset(0, 1).FormulaR1C1 = _
        "=R[0]C[" & -(ColPos - Me.GetRegions.Column) & "]"

    ' Create array of all coordinate tuples and store the weight of each cell
    For Each Region In Regions
        ' Create the string representing the current coordinate tuple
        Let CurrentCoordinateTuple = Region

        ' Compute the weight of the cell with the current coordinate 4-tuple
        Let Counts = Application.CountIf(Me.GetRange.Range("A1").Offset(0, ColPos - 1).Resize(Me.GetRange.Rows.Count, 1), _
                                       "=" & CurrentCoordinateTuple)
                                        
        ' Store this weight in the dictionary
        CoordinateTuples.Add CurrentCoordinateTuple, Counts
    Next Region
    
    ' Clear helper columns
    Call Me.GetRange.Columns(Me.GetNumberOfColumns).Offset(0, 1).Resize(GetRange.Rows.Count, 1).Cells.ClearContents
    
    ' Return dictionary object containing weights
    Let ComputeRegionalSecurityCounts = Application.Transpose(Array(CoordinateTuples.Keys, CoordinateTuples.Items))
End Function

' This function returns a 2D array (Nx2) with the straight, market-relative weights of each sector.
' The first column contains the sectoral coordinates and the second the the weights.
Public Function ComputeSectoralWeights() As Variant
    Dim Sectors As Variant
    Dim Sector As Variant
    Dim CoordinateTuples As Dictionary
    Dim ColPos As Integer
    Dim Weight As Double
    Dim TotalUniverseMarketCap As Double
    Dim CurrentCoordinateTuple As String

    ' Create a new dictionary instance
    Set CoordinateTuples = New Dictionary

    ' Set regions, sectors, styles, and sizes arrays
    Let Sectors = PalladyneSectors()
    
    ' Compute column position of first column to the right of the underlying range
    Let ColPos = Me.GetRange.Columns(Me.GetNumberOfColumns).Offset(0, 1).Column
    
    ' Insert column with coordinate tuples to the right of the underlying range of the PrivateFile instance
    Let GetRange.Columns(Me.GetNumberOfColumns).Offset(0, 1).FormulaR1C1 = _
        "=R[0]C[" & -(ColPos - Me.GetPiamSectors.Column) & "]"

    ' Compute the total market cap of the universe
    Let TotalUniverseMarketCap = Application.Sum(Me.GetUsdMarketCaps.Value2)

    ' Create array of all coordinate tuples and store the weight of each cell
    For Each Sector In Sectors
        ' Create the string representing the current coordinate tuple
        Let CurrentCoordinateTuple = Sector

        ' Compute the weight of the cell with the current coordinate 4-tuple
        Let Weight = Application.SumIf(GetRange.Range("A1").Offset(0, ColPos - 1).Resize(GetRange.Rows.Count, 1), _
                                       "=" & CurrentCoordinateTuple, _
                                       Me.GetUsdMarketCaps) _
                     / TotalUniverseMarketCap
                                        
                    ' Store this weight in the dictionary
                    CoordinateTuples.Add CurrentCoordinateTuple, Weight
    Next Sector
    
    ' Clear helper columns
    Call Me.GetRange.Columns(Me.GetNumberOfColumns).Offset(0, 1).Resize(GetRange.Rows.Count, 1).Cells.ClearContents
    
    ' Return dictionary object containing weights
    Let ComputeSectoralWeights = Application.Transpose(Array(CoordinateTuples.Keys, CoordinateTuples.Items))
End Function

' This function returns a 2D array (Nx2) with the security counts of each sector.
' The first column contains the sectoral coordinates and the second the counts.
Public Function ComputeSectoralSecurityCounts() As Variant
    Dim Sectors As Variant
    Dim Sector As Variant
    Dim CoordinateTuples As Dictionary
    Dim ColPos As Integer
    Dim Counts As Double
    Dim CurrentCoordinateTuple As String

    ' Create a new dictionary instance
    Set CoordinateTuples = New Dictionary

    ' Set regions, sectors, styles, and sizes arrays
    Let Sectors = PalladyneSectors()
    
    ' Compute column position of first column to the right of the underlying range
    Let ColPos = Me.GetRange.Columns(Me.GetNumberOfColumns).Offset(0, 1).Column
    
    ' Insert column with coordinate tuples to the right of the underlying range of the PrivateFile instance
    Let Me.GetRange.Columns(Me.GetNumberOfColumns).Offset(0, 1).FormulaR1C1 = _
        "=R[0]C[" & -(ColPos - Me.GetPiamSectors.Column) & "]"

    ' Create array of all coordinate tuples and store the weight of each cell
    For Each Sector In Sectors
        ' Create the string representing the current coordinate tuple
        Let CurrentCoordinateTuple = Sector

        ' Compute the weight of the cell with the current coordinate 4-tuple
        Let Counts = Application.CountIf(Me.GetRange.Range("A1").Offset(0, ColPos - 1).Resize(Me.GetRange.Rows.Count, 1), _
                                       "=" & CurrentCoordinateTuple)
                                        
                    ' Store this weight in the dictionary
                    CoordinateTuples.Add CurrentCoordinateTuple, Counts
    Next Sector
    
    ' Clear helper columns
    Call Me.GetRange.Columns(Me.GetNumberOfColumns).Offset(0, 1).Resize(GetRange.Rows.Count, 1).Cells.ClearContents
    
    ' Return dictionary object containing weights
    Let ComputeSectoralSecurityCounts = Application.Transpose(Array(CoordinateTuples.Keys, CoordinateTuples.Items))
End Function

' This function returns a 2D array (Nx2) with the straight, market-relative sectoral/regional weights of
' each cell. The first column contains the sectoral/regional coordinates and the second the weights.
Public Function ComputeSectoralRegionalWeights() As Variant
    Dim Sectors As Variant
    Dim Regions As Variant
    Dim Sector As Variant
    Dim Region As Variant
    Dim CoordinateTuples As Dictionary
    Dim ColPos As Integer
    Dim Weight As Double
    Dim TotalUniverseMarketCap As Double
    Dim CurrentCoordinateTuple As String

    ' Create a new dictionary instance
    Set CoordinateTuples = New Dictionary

    ' Set regions, sectors, styles, and sizes arrays
    Let Regions = PalladyneRegions()
    Let Sectors = PalladyneSectors()
    
    ' Compute column position of first column to the right of the underlying range
    Let ColPos = Me.GetRange.Columns(Me.GetNumberOfColumns).Offset(0, 1).Column
    
    ' Insert column with coordinate tuples to the right of the underlying range of the PrivateFile instance
    Let Me.GetRange.Columns(Me.GetNumberOfColumns).Offset(0, 1).FormulaR1C1 = _
        "=R[0]C[" & -(ColPos - Me.GetPiamSectors.Column) & "]&" & Chr(34) & "-" & Chr(34) & _
         "&R[0]C[" & -(ColPos - Me.GetRegions.Column) & "]"

    ' Compute the total market cap of the universe
    Let TotalUniverseMarketCap = Application.Sum(Me.GetUsdMarketCaps.Value2)

    ' Create array of all coordinate tuples and store the weight of each cell
    For Each Sector In Sectors
        For Each Region In Regions
            ' Create the string representing the current coordinate tuple
            Let CurrentCoordinateTuple = Sector & "-" & Region
            
            ' Compute the weight of the cell with the current coordinate 4-tuple
            Let Weight = Application.SumIf(Me.GetRange.Range("A1").Offset(0, ColPos - 1).Resize(Me.GetRange.Rows.Count, 1), _
                                           "=" & CurrentCoordinateTuple, _
                                           Me.GetUsdMarketCaps) _
                         / TotalUniverseMarketCap
                                        
            ' Store this weight in the dictionary
            CoordinateTuples.Add CurrentCoordinateTuple, Weight
        Next Region
    Next Sector
    
    ' Clear helper columns
    Call Me.GetRange.Columns(Me.GetNumberOfColumns).Offset(0, 1).Resize(GetRange.Rows.Count, 1).Cells.ClearContents
    
    ' Return dictionary object containing weights
    Let ComputeSectoralRegionalWeights = Application.Transpose(Array(CoordinateTuples.Keys, CoordinateTuples.Items))
End Function

' This function returns a 2D array (Nx2) with the straight, market-relative sectoral/regional weights of
' each cell. The first column contains the sectoral/regional coordinates and the second the weights.
Public Function ComputeSectoralRegionalSecurityCounts() As Variant
    Dim Sectors As Variant
    Dim Regions As Variant
    Dim Sector As Variant
    Dim Region As Variant
    Dim CoordinateTuples As Dictionary
    Dim ColPos As Integer
    Dim Counts As Double
    Dim CurrentCoordinateTuple As String

    ' Create a new dictionary instance
    Set CoordinateTuples = New Dictionary

    ' Set regions, sectors, styles, and sizes arrays
    Let Regions = PalladyneRegions()
    Let Sectors = PalladyneSectors()
    
    ' Compute column position of first column to the right of the underlying range
    Let ColPos = GetRange.Columns(GetRange.Columns.Count).Offset(0, 1).Column
    
    ' Insert column with coordinate tuples to the right of the underlying range of the PrivateFile instance
    Let GetRange.Columns(GetRange.Columns.Count).Offset(0, 1).FormulaR1C1 = _
        "=R[0]C[" & -(ColPos - GetSectors.Column) & "]&" & Chr(34) & "-" & Chr(34) & _
         "&R[0]C[" & -(ColPos - GetRegions.Column) & "]"
    
    ' Create array of all coordinate tuples and store the weight of each cell
    For Each Sector In Sectors
        For Each Region In Regions
            ' Create the string representing the current coordinate tuple
            Let CurrentCoordinateTuple = Sector & "-" & Region
            
            ' Compute the weight of the cell with the current coordinate 4-tuple
            Let Counts = Application.CountIf(GetRange.Range("A1").Offset(0, ColPos - 1).Resize(GetRange.Rows.Count, 1), _
                                             "=" & CurrentCoordinateTuple)
                                        
            ' Store this weight in the dictionary
            CoordinateTuples.Add CurrentCoordinateTuple, Counts
        Next Region
    Next Sector
    
    ' Clear helper columns
    GetRange.Columns(GetRange.Columns.Count).Offset(0, 1).Resize(GetRange.Rows.Count, 1).Cells.ClearContents
    
    ' Return dictionary object containing weights
    Let ComputeSectoralRegionalSecurityCounts = Application.Transpose(Array(CoordinateTuples.Keys, CoordinateTuples.Items))
End Function

' This function returns a 2D array (Nx2) with the straight, market-relative size/style weights of
' each cell. The first column contains the size/style coordinates and the second the weights.
Public Function ComputeSizeStyleWeights() As Variant
    Dim Sizes As Variant
    Dim Styles As Variant
    Dim Size As Variant
    Dim Style As Variant
    Dim CoordinateTuples As Dictionary
    Dim ColPos As Integer
    Dim Weight As Double
    Dim TotalUniverseMarketCap As Double
    Dim CurrentCoordinateTuple As String

    ' Create a new dictionary instance
    Set CoordinateTuples = New Dictionary

    ' Set regions, sectors, styles, and sizes arrays
    Let Sizes = PalladyneSizes()
    Let Styles = PalladyneStyles()
    
    ' Compute column position of first column to the right of the underlying range
    Let ColPos = Me.GetRange.Columns(Me.GetNumberOfColumns).Offset(0, 1).Column
    
    ' Insert column with coordinate tuples to the right of the underlying range of the PrivateFile instance
    Let Me.GetRange.Columns(Me.GetNumberOfColumns).Offset(0, 1).FormulaR1C1 = _
        "=R[0]C[" & -(ColPos - Me.GetSizes.Column) & "]&" & Chr(34) & "-" & Chr(34) & _
         "&R[0]C[" & -(ColPos - Me.GetStyles.Column) & "]"

    ' Compute the total market cap of the universe
    Let TotalUniverseMarketCap = Application.Sum(Me.GetUsdMarketCaps.Value2)

    ' Create array of all coordinate tuples and store the weight of each cell
    For Each Size In Sizes
        For Each Style In Styles
            ' Create the string representing the current coordinate tuple
            Let CurrentCoordinateTuple = Size & "-" & Style
            
            ' Compute the weight of the cell with the current coordinate 4-tuple
            Let Weight = Application.SumIf(Me.GetRange.Range("A1").Offset(0, ColPos - 1).Resize(Me.GetRange.Rows.Count, 1), _
                                           "=" & CurrentCoordinateTuple, _
                                           Me.GetUsdMarketCaps) _
                         / TotalUniverseMarketCap
                                        
            ' Store this weight in the dictionary
            CoordinateTuples.Add CurrentCoordinateTuple, Weight
        Next Style
    Next Size
    
    ' Clear helper columns
    Call Me.GetRange.Columns(Me.GetNumberOfColumns).Offset(0, 1).Resize(GetRange.Rows.Count, 1).Cells.ClearContents
    
    ' Return dictionary object containing weights
    Let ComputeSizeStyleWeights = Application.Transpose(Array(CoordinateTuples.Keys, CoordinateTuples.Items))
End Function

' This function returns a 2D array (Nx2) with the size/style security counts of
' each cell. The first column contains the size/style coordinates and the second the counts.
Public Function ComputeSizeStyleSecurityCounts() As Variant
    Dim Sizes As Variant
    Dim Styles As Variant
    Dim Size As Variant
    Dim Style As Variant
    Dim CoordinateTuples As Dictionary
    Dim ColPos As Integer
    Dim Counts As Double
    Dim CurrentCoordinateTuple As String

    ' Create a new dictionary instance
    Set CoordinateTuples = New Dictionary

    ' Set regions, sectors, styles, and sizes arrays
    Let Sizes = PalladyneSizes()
    Let Styles = PalladyneStyles()
    
    ' Compute column position of first column to the right of the underlying range
    Let ColPos = GetRange.Columns(GetRange.Columns.Count).Offset(0, 1).Column
    
    ' Insert column with coordinate tuples to the right of the underlying range of the PrivateFile instance
    Let GetRange.Columns(GetRange.Columns.Count).Offset(0, 1).FormulaR1C1 = _
        "=R[0]C[" & -(ColPos - GetSizes.Column) & "]&" & Chr(34) & "-" & Chr(34) & _
         "&R[0]C[" & -(ColPos - GetStyles.Column) & "]"

    ' Create array of all coordinate tuples and store the weight of each cell
    For Each Size In Sizes
        For Each Style In Styles
            ' Create the string representing the current coordinate tuple
            Let CurrentCoordinateTuple = Size & "-" & Style
            
            ' Compute the weight of the cell with the current coordinate 4-tuple
            Let Counts = Application.CountIf(GetRange.Range("A1").Offset(0, ColPos - 1).Resize(GetRange.Rows.Count, 1), _
                                             "=" & CurrentCoordinateTuple)
                                        
            ' Store this weight in the dictionary
            CoordinateTuples.Add CurrentCoordinateTuple, Counts
        Next Style
    Next Size
    
    ' Clear helper columns
    GetRange.Columns(GetRange.Columns.Count).Offset(0, 1).Resize(GetRange.Rows.Count, 1).Cells.ClearContents
    
    ' Return dictionary object containing weights
    Let ComputeSizeStyleSecurityCounts = Application.Transpose(Array(CoordinateTuples.Keys, CoordinateTuples.Items))
End Function

' Compute cellular, market-weighted ranks
Public Function ComputeCellularMarketCapWeightedRanks() As Variant
    Dim Sector As Variant
    Dim Region As Variant
    Dim Style As Variant
    Dim Size As Variant
    Dim ColPos As Integer
    Dim CurrentCoordinateTuple As String
    Dim SumMarketCapWeightedRanks As Double
    Dim TotalMarketCap As Double
    Dim CellularRanks  As Dictionary

    ' Create a new dictionary instance
    Set CellularRanks = New Dictionary

    ' Compute column position of first column to the right of the underlying range
    Let ColPos = GetRange.Columns(Me.GetNumberOfColumns).Offset(0, 1).Column
    
    ' Insert column with coordinate tuples to the right of the underlying range of the PrivateFile instance
    Let Me.GetRange.Columns(Me.GetNumberOfColumns).Offset(0, 1).FormulaR1C1 = _
        "=R[0]C[" & -(ColPos - Me.GetPiamSectors.Column) & "]&" & Chr(34) & "-" & Chr(34) & _
         "&R[0]C[" & -(ColPos - Me.GetRegions.Column) & "]&" & Chr(34) & "-" & Chr(34) & _
         "&R[0]C[" & -(ColPos - Me.GetSizes.Column) & "]&" & Chr(34) & "-" & Chr(34) & _
         "&R[0]C[" & -(ColPos - Me.GetStyles.Column) & "]"
         
    ' Insert column with formulas for rank * market cap for every security
    Let GetRange.Columns(Me.GetNumberOfColumns).Offset(0, 2).FormulaR1C1 = _
        "=R[0]C[" & -(ColPos + 1 - Me.GetUsdMarketCaps.Column) & "] * R[0]C[" & -(ColPos + 1 - Me.GetRanks.Column) & "]"

    ' Cycle through all cell, computing the market cap-weighted rank of each
    For Each Sector In PalladyneSectors()
        For Each Region In PalladyneRegions()
            For Each Size In PalladyneSizes()
                For Each Style In PalladyneStyles()
                    ' Create the string representing the current coordinate tuple
                    Let CurrentCoordinateTuple = Sector & "-" & Region & "-" & Size & "-" & Style
    
                    ' Sum the rank * market cap of all stocks in this coordinate
                    Let TotalMarketCap = Application.SumIf(Me.GetRange.Columns(Me.GetNumberOfColumns).Offset(0, 1), _
                                                           "=" & CurrentCoordinateTuple, _
                                                           Me.GetUsdMarketCaps)
    
                    ' Compute the marketcap-weight rank
                    Let SumMarketCapWeightedRanks = _
                        Application.SumIf(Me.GetRange.Columns(Me.GetNumberOfColumns).Offset(0, 1), _
                                          "=" & CurrentCoordinateTuple, _
                                          Me.GetRange.Columns(Me.GetNumberOfColumns).Offset(0, 2))
    
                    ' Store this cellular rank to the dictionary
                    If TotalMarketCap = 0 Then
                        CellularRanks.Add CurrentCoordinateTuple, 0
                    Else
                        CellularRanks.Add CurrentCoordinateTuple, SumMarketCapWeightedRanks / TotalMarketCap
                    End If

                Next Style
            Next Size
        Next Region
    Next Sector
    
    ' Clear helper columns
    Call Me.GetRange.Columns(Me.GetNumberOfColumns).Offset(0, 1).ClearContents
    Call Me.GetRange.Columns(Me.GetNumberOfColumns).Offset(0, 2).ClearContents
    
    ' Return 2D array with the market cap-weighted ranks
    Let ComputeCellularMarketCapWeightedRanks = Application.Transpose(Array(CellularRanks.Keys, CellularRanks.Items))
End Function

' Compute cellular, market-weighted PEs
Public Function ComputeCellularMarketCapWeightedPes() As Variant
    Dim Sector As Variant
    Dim Region As Variant
    Dim Style As Variant
    Dim Size As Variant
    Dim ColPos As Integer
    Dim CurrentCoordinateTuple As String
    Dim SumMarketCapWeightedPes As Double
    Dim TotalMarketCap As Double
    Dim CellularPes  As Dictionary

    ' Create a new dictionary instance
    Set CellularPes = New Dictionary

    ' Compute column position of first column to the right of the underlying range
    Let ColPos = GetRange.Columns(GetRange.Columns.Count).Offset(0, 1).Column
    
    ' Insert column with coordinate tuples to the right of the underlying range of the PrivateFile instance
    Let GetRange.Columns(GetRange.Columns.Count).Offset(0, 1).FormulaR1C1 = _
        "=R[0]C[" & -(ColPos - GetSectors.Column) & "]&" & Chr(34) & "-" & Chr(34) & _
         "&R[0]C[" & -(ColPos - GetRegions.Column) & "]&" & Chr(34) & "-" & Chr(34) & _
         "&R[0]C[" & -(ColPos - GetSizes.Column) & "]&" & Chr(34) & "-" & Chr(34) & _
         "&R[0]C[" & -(ColPos - GetStyles.Column) & "]"
         
    ' Insert column with formulas for PE * market cap for every security
    Let GetRange.Columns(GetRange.Columns.Count).Offset(0, 2).FormulaR1C1 = _
        "=R[0]C[" & -(ColPos + 1 - GetMarketCaps.Column) & "] * R[0]C[" & -(ColPos + 1 - GetF1ForwardPes.Column) & "]"

    ' Cycle through all cell, computing the market cap-weighted rank of each
    For Each Sector In PalladyneSectors()
        For Each Region In PalladyneRegions()
            For Each Size In PalladyneSizes()
                For Each Style In PalladyneStyles()
                    ' Create the string representing the current coordinate tuple
                    Let CurrentCoordinateTuple = Sector & "-" & Region & "-" & Size & "-" & Style
    
                    ' Sum the rank * market cap of all stocks in this coordinate
                    Let TotalMarketCap = Application.SumIf(GetRange.Columns(GetRange.Columns.Count).Offset(0, 1), _
                                                           "=" & CurrentCoordinateTuple, _
                                                           GetMarketCaps)
    
                    ' Compute the marketcap-weight rank
                    Let SumMarketCapWeightedPes = _
                        Application.SumIf(GetRange.Columns(GetRange.Columns.Count).Offset(0, 1), _
                                          "=" & CurrentCoordinateTuple, _
                                          GetRange.Columns(GetRange.Columns.Count).Offset(0, 2))
    
                    ' Store this cellular rank to the dictionary
                    If TotalMarketCap = 0 Then
                        CellularPes.Add CurrentCoordinateTuple, 0
                    Else
                        CellularPes.Add CurrentCoordinateTuple, SumMarketCapWeightedPes / TotalMarketCap
                    End If

                Next Style
            Next Size
        Next Region
    Next Sector
    
    ' Clear helper columns
    GetRange.Columns(GetRange.Columns.Count).Offset(0, 1).ClearContents
    GetRange.Columns(GetRange.Columns.Count).Offset(0, 2).ClearContents
    
    ' Return 2D array with the market cap-weighted ranks
    Let ComputeCellularMarketCapWeightedPes = Application.Transpose(Array(CellularPes.Keys, CellularPes.Items))
End Function

' Compute sectoral/regional, market-weighted ranks
Public Function ComputeSectoralRegionalMarketCapWeightedRanks() As Variant
    Dim Sector As Variant
    Dim Region As Variant
    Dim ColPos As Integer
    Dim CurrentCoordinateTuple As String
    Dim SumMarketCapWeightedRanks As Double
    Dim TotalMarketCap As Double
    Dim CellularRanks  As Dictionary

    ' Create a new dictionary instance
    Set CellularRanks = New Dictionary

    ' Compute column position of first column to the right of the underlying range
    Let ColPos = Me.GetRange.Columns(Me.GetNumberOfColumns).Offset(0, 1).Column
    
    ' Insert column with sector/region coordinates to the right of the underlying range of the PrivateFile instance
    Let GetRange.Columns(Me.GetNumberOfColumns).Offset(0, 1).FormulaR1C1 = _
        "=R[0]C[" & -(ColPos - Me.GetPiamSectors.Column) & "]&" & Chr(34) & "-" & Chr(34) & _
         "&R[0]C[" & -(ColPos - Me.GetRegions.Column) & "]"
         
    ' Insert column with formulas for rank * market cap for every security
    Let GetRange.Columns(Me.GetNumberOfColumns).Offset(0, 2).FormulaR1C1 = _
        "=R[0]C[" & -(ColPos + 1 - Me.GetUsdMarketCaps.Column) & "] * R[0]C[" & -(ColPos + 1 - Me.GetRanks.Column) & "]"

    ' Cycle through all cell, computing the market cap-weighted rank of each
    For Each Sector In PalladyneSectors()
        For Each Region In PalladyneRegions()
                    ' Create the string representing the current coordinate
                    Let CurrentCoordinateTuple = Sector & "-" & Region
    
                    ' Sum the rank * market cap of all stocks in this coordinate
                    Let TotalMarketCap = Application.SumIf(Me.GetRange.Columns(Me.GetNumberOfColumns).Offset(0, 1), _
                                                           "=" & CurrentCoordinateTuple, _
                                                           Me.GetUsdMarketCaps)
    
                    ' Compute the marketcap-weight rank
                    Let SumMarketCapWeightedRanks = _
                        Application.SumIf(Me.GetRange.Columns(Me.GetNumberOfColumns).Offset(0, 1), _
                                          "=" & CurrentCoordinateTuple, _
                                          Me.GetRange.Columns(Me.GetNumberOfColumns).Offset(0, 2))
    
                    ' Store this cellular rank to the dictionary
                    If TotalMarketCap = 0 Then
                        CellularRanks.Add CurrentCoordinateTuple, 0
                    Else
                        CellularRanks.Add CurrentCoordinateTuple, SumMarketCapWeightedRanks / TotalMarketCap
                    End If
        Next Region
    Next Sector
    
    ' Clear helper columns
    Call Me.GetRange.Columns(Me.GetNumberOfColumns).Offset(0, 1).ClearContents
    Call Me.GetRange.Columns(Me.GetNumberOfColumns).Offset(0, 2).ClearContents
    
    ' Return 2D array with the market cap-weighted ranks
    Let ComputeSectoralRegionalMarketCapWeightedRanks = Application.Transpose(Array(CellularRanks.Keys, CellularRanks.Items))
End Function

' Compute sectoral/regional, market-weighted PEs
Public Function ComputeSectoralRegionalMarketCapWeightedPEs() As Variant
    Dim Sector As Variant
    Dim Region As Variant
    Dim ColPos As Integer
    Dim CurrentCoordinateTuple As String
    Dim SumMarketCapWeightedPes As Double
    Dim TotalMarketCap As Double
    Dim CellularPes  As Dictionary

    ' Create a new dictionary instance
    Set CellularPes = New Dictionary

    ' Compute column position of first column to the right of the underlying range
    Let ColPos = GetRange.Columns(GetRange.Columns.Count).Offset(0, 1).Column
    
    ' Insert column with sector/region coordinates to the right of the underlying range of the PrivateFile instance
    Let GetRange.Columns(GetRange.Columns.Count).Offset(0, 1).FormulaR1C1 = _
        "=R[0]C[" & -(ColPos - GetSectors.Column) & "]&" & Chr(34) & "-" & Chr(34) & _
         "&R[0]C[" & -(ColPos - GetRegions.Column) & "]"
         
    ' Insert column with formulas for rank * market cap for every security
    Let GetRange.Columns(GetRange.Columns.Count).Offset(0, 2).FormulaR1C1 = _
        "=R[0]C[" & -(ColPos + 1 - GetMarketCaps.Column) & "] * R[0]C[" & -(ColPos + 1 - GetF1ForwardPes.Column) & "]"

    ' Cycle through all cell, computing the market cap-weighted rank of each
    For Each Sector In PalladyneSectors()
        For Each Region In PalladyneRegions()
                    ' Create the string representing the current coordinate
                    Let CurrentCoordinateTuple = Sector & "-" & Region
    
                    ' Sum the rank * market cap of all stocks in this coordinate
                    Let TotalMarketCap = Application.SumIf(GetRange.Columns(GetRange.Columns.Count).Offset(0, 1), _
                                                           "=" & CurrentCoordinateTuple, _
                                                           GetMarketCaps)
    
                    ' Compute the marketcap-weight rank
                    Let SumMarketCapWeightedPes = _
                        Application.SumIf(GetRange.Columns(GetRange.Columns.Count).Offset(0, 1), _
                                          "=" & CurrentCoordinateTuple, _
                                          GetRange.Columns(GetRange.Columns.Count).Offset(0, 2))
    
                    ' Store this cellular rank to the dictionary
                    If TotalMarketCap = 0 Then
                        CellularPes.Add CurrentCoordinateTuple, 0
                    Else
                        CellularPes.Add CurrentCoordinateTuple, SumMarketCapWeightedPes / TotalMarketCap
                    End If
        Next Region
    Next Sector
    
    ' Clear helper columns
    GetRange.Columns(GetRange.Columns.Count).Offset(0, 1).ClearContents
    GetRange.Columns(GetRange.Columns.Count).Offset(0, 2).ClearContents
    
    ' Return 2D array with the market cap-weighted PEs
    Let ComputeSectoralRegionalMarketCapWeightedPEs = Application.Transpose(Array(CellularPes.Keys, CellularPes.Items))
End Function

' Compute size/style, market-weighted ranks
Public Function ComputeSizeStyleMarketCapWeightedRanks() As Variant
    Dim Style As Variant
    Dim Size As Variant
    Dim ColPos As Integer
    Dim CurrentCoordinateTuple As String
    Dim SumMarketCapWeightedRanks As Double
    Dim TotalMarketCap As Double
    Dim CellularRanks  As Dictionary

    ' Create a new dictionary instance
    Set CellularRanks = New Dictionary

    ' Compute column position of first column to the right of the underlying range
    Let ColPos = Me.GetRange.Columns(Me.GetNumberOfColumns).Offset(0, 1).Column
    
    ' Insert column with sector/region coordinates to the right of the underlying range of the PrivateFile instance
    Let GetRange.Columns(Me.GetNumberOfColumns).Offset(0, 1).FormulaR1C1 = _
        "=R[0]C[" & -(ColPos - Me.GetSizes.Column) & "]&" & Chr(34) & "-" & Chr(34) & _
         "&R[0]C[" & -(ColPos - Me.GetStyles.Column) & "]"
         
    ' Insert column with formulas for rank * market cap for every security
    Let GetRange.Columns(Me.GetNumberOfColumns).Offset(0, 2).FormulaR1C1 = _
        "=R[0]C[" & -(ColPos + 1 - Me.GetUsdMarketCaps.Column) & "] * R[0]C[" & -(ColPos + 1 - Me.GetRanks.Column) & "]"

    ' Cycle through all cell, computing the market cap-weighted rank of each
    For Each Size In PalladyneSizes()
        For Each Style In PalladyneStyles()
                    ' Create the string representing the current coordinate tuple
                    Let CurrentCoordinateTuple = Size & "-" & Style
    
                    ' Sum the rank * market cap of all stocks in this coordinate
                    Let TotalMarketCap = Application.SumIf(Me.GetRange.Columns(Me.GetNumberOfColumns).Offset(0, 1), _
                                                           "=" & CurrentCoordinateTuple, _
                                                           Me.GetUsdMarketCaps)
    
                    ' Compute the marketcap-weight rank
                    Let SumMarketCapWeightedRanks = _
                        Application.SumIf(Me.GetRange.Columns(Me.GetNumberOfColumns).Offset(0, 1), _
                                          "=" & CurrentCoordinateTuple, _
                                          Me.GetRange.Columns(Me.GetNumberOfColumns).Offset(0, 2))
    
                    ' Store this cellular rank to the dictionary
                    If TotalMarketCap = 0 Then
                        CellularRanks.Add CurrentCoordinateTuple, 0
                    Else
                        CellularRanks.Add CurrentCoordinateTuple, SumMarketCapWeightedRanks / TotalMarketCap
                    End If
        Next Style
    Next Size
    
    ' Clear helper columns
    Call Me.GetRange.Columns(Me.GetNumberOfColumns).Offset(0, 1).ClearContents
    Call Me.GetRange.Columns(Me.GetNumberOfColumns).Offset(0, 2).ClearContents
    
    ' Return 2D array with the market cap-weighted ranks
    Let ComputeSizeStyleMarketCapWeightedRanks = Application.Transpose(Array(CellularRanks.Keys, CellularRanks.Items))
End Function

' Compute size/style, market-weighted PEs
Public Function ComputeSizeStyleMarketCapWeightedPes() As Variant
    Dim Style As Variant
    Dim Size As Variant
    Dim ColPos As Integer
    Dim CurrentCoordinateTuple As String
    Dim SumMarketCapWeightedPes As Double
    Dim TotalMarketCap As Double
    Dim CellularPes  As Dictionary

    ' Create a new dictionary instance
    Set CellularPes = New Dictionary

    ' Compute column position of first column to the right of the underlying range
    Let ColPos = GetRange.Columns(GetRange.Columns.Count).Offset(0, 1).Column
    
    ' Insert column with sector/region coordinates to the right of the underlying range of the PrivateFile instance
    Let GetRange.Columns(GetRange.Columns.Count).Offset(0, 1).FormulaR1C1 = _
        "=R[0]C[" & -(ColPos - GetSizes.Column) & "]&" & Chr(34) & "-" & Chr(34) & _
         "&R[0]C[" & -(ColPos - GetStyles.Column) & "]"
         
    ' Insert column with formulas for rank * market cap for every security
    Let GetRange.Columns(GetRange.Columns.Count).Offset(0, 2).FormulaR1C1 = _
        "=R[0]C[" & -(ColPos + 1 - GetMarketCaps.Column) & "] * R[0]C[" & -(ColPos + 1 - GetF1ForwardPes.Column) & "]"

    ' Cycle through all cell, computing the market cap-weighted rank of each
    For Each Size In PalladyneSizes()
        For Each Style In PalladyneStyles()
                    ' Create the string representing the current coordinate tuple
                    Let CurrentCoordinateTuple = Size & "-" & Style
    
                    ' Sum the rank * market cap of all stocks in this coordinate
                    Let TotalMarketCap = Application.SumIf(GetRange.Columns(GetRange.Columns.Count).Offset(0, 1), _
                                                           "=" & CurrentCoordinateTuple, _
                                                           GetMarketCaps)
    
                    ' Compute the marketcap-weight rank
                    Let SumMarketCapWeightedPes = _
                        Application.SumIf(GetRange.Columns(GetRange.Columns.Count).Offset(0, 1), _
                                          "=" & CurrentCoordinateTuple, _
                                          GetRange.Columns(GetRange.Columns.Count).Offset(0, 2))
    
                    ' Store this cellular rank to the dictionary
                    If TotalMarketCap = 0 Then
                        CellularPes.Add CurrentCoordinateTuple, 0
                    Else
                        CellularPes.Add CurrentCoordinateTuple, SumMarketCapWeightedPes / TotalMarketCap
                    End If
        Next Style
    Next Size
    
    ' Clear helper columns
    GetRange.Columns(GetRange.Columns.Count).Offset(0, 1).ClearContents
    GetRange.Columns(GetRange.Columns.Count).Offset(0, 2).ClearContents
    
    ' Return 2D array with the market cap-weighted ranks
    Let ComputeSizeStyleMarketCapWeightedPes = Application.Transpose(Array(CellularPes.Keys, CellularPes.Items))
End Function

' Compute sectoral market-weighted ranks
Public Function ComputeSectoralMarketCapWeightedRanks() As Variant
    Dim Sector As Variant
    Dim ColPos As Integer
    Dim CurrentCoordinateTuple As String
    Dim SumMarketCapWeightedRanks As Double
    Dim TotalMarketCap As Double
    Dim CellularRanks  As Dictionary

    ' Create a new dictionary instance
    Set CellularRanks = New Dictionary

    ' Compute column position of first column to the right of the underlying range
    Let ColPos = GetRange.Columns(GetRange.Columns.Count).Offset(0, 1).Column
         
    ' Insert column with formulas for rank * market cap for every security
    Let Me.GetRange.Columns(Me.GetRange.Columns.Count).Offset(0, 1).FormulaR1C1 = _
        "=R[0]C[" & -(ColPos - Me.GetMarketCaps.Column) & "] * R[0]C[" & -(ColPos - Me.GetRanks.Column) & "]"

    ' Cycle through all cell, computing the market cap-weighted rank of each
    For Each Sector In PalladyneSectors()
        ' Create the string representing the current coordinate 4-tuple
        Let CurrentCoordinateTuple = Size & "-" & Style
    
        ' Sum the rank * market cap of all stocks in this coordinate
        Let TotalMarketCap = Application.SumIf(Me.GetRange.Columns(Me.GetRange.Columns.Count).Offset(0, 1), _
                                               CurrentCoordinate4Tuple, _
                                               Me.GetMarketCaps)
    
        ' Compute the marketcap-weight rank
        Let SumMarketCapWeightedRanks = _
            Application.SumIf(Me.GetRange.Columns(Me.GetRange.Columns.Count).Offset(0, 1), _
                              CurrentCoordinate4Tuple, _
                              Me.GetRange.Columns(Me.GetRange.Columns.Count).Offset(0, 2))
    
        ' Store this cellular rank to the dictionary
        If TotalMarketCap = 0 Then
            Call CellularRanks.Add(CurrentCoordinateTuple, 0)
        Else
            Call CellularRanks.Add(CurrentCoordinateTuple, SumMarketCapWeightedRanks / TotalMarketCap)
        End If
    Next Sector
    
    ' Clear helper columns
    Call Me.GetRange.Columns(Me.GetRange.Columns.Count).Offset(0, 1).ClearContents
    
    ' Return 2D array with the market cap-weighted ranks
    Let ComputeSectoralMarketCapWeightedRanks = Application.Transpose(Array(CellularRanks.Keys, CellularRanks.Items))
End Function

' Compute sectoral market-weighted PEs
Public Function ComputeSectoralMarketCapWeightedPes() As Variant
    Dim Sector As Variant
    Dim ColPos As Integer
    Dim CurrentCoordinateTuple As String
    Dim SumMarketCapWeightedPes As Double
    Dim TotalMarketCap As Double
    Dim CellularPes  As Dictionary

    ' Create a new dictionary instance
    Set CellularPes = New Dictionary

    ' Compute column position of first column to the right of the underlying range
    Let ColPos = Me.GetRange.Columns(Me.GetRange.Columns.Count).Offset(0, 1).Column
         
    ' Insert column with formulas for rank * market cap for every security
    Let Me.GetRange.Columns(Me.GetRange.Columns.Count).Offset(0, 1).FormulaR1C1 = _
        "=R[0]C[" & -(ColPos - Me.GetMarketCaps.Column) & "] * R[0]C[" & -(ColPos - Me.GetFy1Pes.Column) & "]"

    ' Cycle through all cell, computing the market cap-weighted rank of each
    For Each Sector In PalladyneSectors()
        ' Create the string representing the current coordinate 4-tuple
        Let CurrentCoordinateTuple = Size & "-" & Style
    
        ' Sum the rank * market cap of all stocks in this coordinate
        Let TotalMarketCap = Application.SumIf(Me.GetRange.Columns(GetRange.Columns.Count).Offset(0, 1), _
                                               CurrentCoordinate4Tuple, _
                                               Me.GetMarketCaps)
    
        ' Compute the marketcap-weight rank
        Let SumMarketCapWeightedPes = _
            Application.SumIf(Me.GetRange.Columns(GetRange.Columns.Count).Offset(0, 1), _
                              CurrentCoordinate4Tuple, _
                              Me.GetRange.Columns(GetRange.Columns.Count).Offset(0, 2))
    
        ' Store this cellular rank to the dictionary
        If TotalMarketCap = 0 Then
            Call Me.CellularPes.Add(CurrentCoordinateTuple, 0)
        Else
            Call CellularPes.Add(CurrentCoordinateTuple, SumMarketCapWeightedPes / TotalMarketCap)
        End If
    Next Sector
    
    ' Clear helper columns
    Me.GetRange.Columns(Me.GetRange.Columns.Count).Offset(0, 1).ClearContents
    
    ' Return 2D array with the market cap-weighted ranks
    Let ComputeSectoralMarketCapWeightedPes = Application.Transpose(Array(Me.CellularPes.Keys, Me.CellularPes.Items))
End Function

' Compute the portfolio's, capital-weighted rank
Public Function ComputePortfolioMarketCapWeightedRank() As Double
    Let ComputePortfolioMarketCapWeightedRank = Application.SumProduct(Me.GetRanks, Me.GetUsdMarketCaps) / Application.Sum(Me.GetUsdMarketCaps)
End Function

' This function initializes this instance with the contents of a private file. It then classifies the instance by doing
' the following:
' 1. Initialize the underlying range with the given range
' 2. Load and initialize the country-to-region mapping
' 3. Load and initialize the GICS code-to-MSCI sub-industry code mapping
' 4. Classify and fill-in region, size, style, and MSCI sub-industry code and MSCI sub-industry name for all stocks
' The optional boolean parameter FilteredOutUnacceptablePesAndCountries specifies whether or not filtered
' securities should be referenced by the underlying range.
' This function is the equivalent of PrivateFile.InitializeAndClassifyPrivateFile() below, starting from a filename
' instead of a source range.
Public Sub LoadAndProcessPrivateFile(PrivateFileFileName As String, _
                                     Optional FilteredOutUnacceptablePesAndCountries As Boolean = True)
    ' Open first private file. Set source and destination ranges for first private file and classify and initialize it
    Call Me.InitializeWithPrivateFile(PrivateFileFileName)
        
    Call HelperSubForLoadAndProcessPrivateFile(CountryRegionRelationMapFileName(), _
                                               GicsCodeToMsciSubIndustryCodeMapFileName(), _
                                               FilteredOutUnacceptablePesAndCountries)
End Sub

Private Sub HelperSubForLoadAndProcessPrivateFile(TheCountryToRegionMapFilename As String, _
                                                  TheGicsToMsciCodeMapFilename As String, _
                                                  Optional FilteredOutUnacceptablePesAndCountries As Boolean = True)
    Dim PrivateFileRowsFilteredForPes As Variant
    Dim PrivateFileRowsFilteredForCountries As Variant
    Dim NumberOfAcceptableSecurities As Integer
    Dim NumberOfSedols As Long
    Dim RangeOfAcceptableSecurities As Range
    Dim SubTargetRange As Range
    Dim TargetRange As Range
    Dim i As Integer
           
    ' Load and initialize country-to-region and GICS-to-MSCI mappings
    Call Me.InitializeCountryToRegionMap(TheCountryToRegionMapFilename)
    Call Me.InitializeGicsToMsciSubIndustryCodeMappingWithFile(TheGicsToMsciCodeMapFilename)
    
    ' Filter out non-positive and extreme PEs.  Doing this step here deletes approximately 2,500 securities, which
    ' speeds up all subsequent steps. Get rid of the header row.
    Let PrivateFileRowsFilteredForPes = Me.FilterOutUnacceptablePes()
    Call ToTemp(PrivateFileRowsFilteredForPes)
    Set PrivateFileRowsFilteredForPes = ThisWorkbook.Worksheets("TempComputation").Range("A1").CurrentRegion
    Set PrivateFileRowsFilteredForPes = _
        PrivateFileRowsFilteredForPes.Resize(PrivateFileRowsFilteredForPes.Rows.Count - 1, Me.GetNumberOfColumns)
    Let PrivateFileRowsFilteredForPes = PrivateFileRowsFilteredForPes.Offset(1, 0).Value2
    
    ' Filter out securities with missing country codes. Get rid of the header row.
    Let PrivateFileRowsFilteredForCountries = Me.FilterOutMissingCountries()
    If UBound(PrivateFileRowsFilteredForCountries) > 1 Then
        Call ToTemp(PrivateFileRowsFilteredForCountries)
        Set PrivateFileRowsFilteredForCountries = ThisWorkbook.Worksheets("TempComputation").Range("A1").CurrentRegion
        Set PrivateFileRowsFilteredForCountries = _
            PrivateFileRowsFilteredForCountries.Resize(PrivateFileRowsFilteredForCountries.Rows.Count - 1, _
                                                       PrivateFileRowsFilteredForCountries.Columns.Count)
        Let PrivateFileRowsFilteredForCountries = PrivateFileRowsFilteredForCountries.Offset(1, 0).Value
    End If
 
    ' Compute and load regional codes.  These must be done before initializing size and style threshold tables
    Call Me.ComputeRegionalCodes
    
    ' Compute threshold and size tables.  These must be done after filtering unacceptable PEs. Initialize GICS-to-MSCI mapping
    Call Me.InitializeSizeThresholds
    Call Me.InitializeStyleThresholds

    ' Compute and populate regions, sizes, and styles for all securities.
    Call Me.ComputesSizes
    Call Me.ComputesStyles
    Call Me.ComputeClassifications
    
    ' Compute the number of securities for which the PEs are in the acceptable range and the country codes are not missing
    Let NumberOfAcceptableSecurities = LastNonBlankRowIndexInColumnNotEqualTo(Me.GetRange.Columns(1).EntireColumn, "") - 1
    Set RangeOfAcceptableSecurities = Me.GetRange.Worksheet.Range("A1").Resize(NumberOfAcceptableSecurities + 1, _
                                                                               Me.GetNumberOfColumns)
    
    ' If filtered out securities are to be dropped from the file, simply set the underlying range to reference
    ' the securities that passed the filter.
    If FilteredOutUnacceptablePesAndCountries Then
        ' Initialize the Private File instance with the sub target range
        Call Me.InitializeWithRange(RangeOfAcceptableSecurities)
    Else
        ' Append to the underlying range the rows deleted for unacceptable PEs
        If UBound(PrivateFileRowsFilteredForPes) > 1 Then
            Set SubTargetRange = Me.GetRange.Range("A1").Offset(GetRange.Rows.Count, 0)
            Set SubTargetRange = SubTargetRange.Resize(UBound(PrivateFileRowsFilteredForPes, 1), _
                                                       UBound(PrivateFileRowsFilteredForPes, 2))
            Let SubTargetRange.Value2 = PrivateFileRowsFilteredForPes
        End If

        ' Append to the underlying range the rows deleted for missing countries
        If UBound(PrivateFileRowsFilteredForCountries) > 1 Then
            If UBound(PrivateFileRowsFilteredForPes) > 1 Then
                Set SubTargetRange = Me.GetRange.Range("A1").Offset(GetRange.Rows.Count + SubTargetRange.Rows.Count, 0)
                Set SubTargetRange = SubTargetRange.Resize(UBound(PrivateFileRowsFilteredForCountries, 1), _
                                                           UBound(PrivateFileRowsFilteredForCountries, 2))
                Let SubTargetRange.Value = PrivateFileRowsFilteredForCountries
            Else
                Set SubTargetRange = GetRange.Range("A1").Offset(GetRange.Rows.Count, 0)
                Set SubTargetRange = SubTargetRange.Resize(UBound(PrivateFileRowsFilteredForCountries, 1), _
                                                           UBound(PrivateFileRowsFilteredForCountries, 2))
                Let SubTargetRange.Value = PrivateFileRowsFilteredForCountries
            End If
        End If

        ' Compute the regions, sizes, styles, styles, and MSCI sub-industries for the appended rows.
        Set SubTargetRange = GetRange.Range("A1").Offset(GetRange.Rows.Count - 1, 0)
        Set SubTargetRange = _
            SubTargetRange.Resize(UBound(PrivateFileRowsFilteredForPes, 1) + UBound(PrivateFileRowsFilteredForCountries, 1), _
                                  UBound(PrivateFileRowsFilteredForCountries, 2))
    
        ' Re-initialize this PrivateFile instance with the range of appended rows
        Call Me.InitializeWithRange(SubTargetRange)
    
        ' Loop through each of the rows appended adding region, size, style, and MSCI sub-industry,
        ' when possible
        Call Me.ComputeRegionalCodes
        Call Me.ComputesSizes
        Call Me.ComputesStyles
        Call Me.ComputeClassifications
        
        ' Repoint the underlying range to refererence all rows
        Let NumberOfSedols = LastNonBlankRowIndexInColumnNotEqualTo(Me.GetRange.Columns(1).EntireColumn, "") - 1
        Call Me.InitializeWithRange(Me.GetRange.Worksheet.Range("A1").Resize(NumberOfSedols + 1, Me.GetNumberOfColumns))
    End If
End Sub

' Returns a 2D matrix with the style thresholds.  The first column contains the coordinates and the second
' the threholds
Public Property Get GetStyleThresholdsMatrix() As Variant
    Let GetStyleThresholdsMatrix = Application.Transpose(Array(Me.StyleThresholds.Keys, Me.StyleThresholds.Items))
End Property

' Returns a 2D matrix with the style thresholds.  The first column contains the coordinates and the second
' the threholds
Public Property Get GetSyzeThresholdsMatrix() As Variant
    Let GetSyzeThresholdsMatrix = Application.Transpose(Array(Me.SizeThresholds.Keys, Me.SizeThresholds.Items))
End Property

' Returns the ith row in the private file
Property Get GetIthRow(i As Long) As PrivateFileRow
    Dim ARow As PrivateFileRow
    
    Set ARow = New PrivateFileRow
    
    Call ARow.InitializeWithRange(Me.GetRange.Rows(i))
    
    Set GetIthRow = ARow
End Property

' This function populates the classification column.  It can be run only after region, size, and style have been
' populated
Public Sub ComputeClassifications()
    Dim i As Long
    Dim Results() As String
    
    ReDim Results(Me.GetDbLength)
    
    For i = 1 To Me.GetDbLength
        Let Results(i) = Me.GetIthRow(i).GetPiamSector.Value2 & "-" & Me.GetIthRow(i).GetRegion.Value2 & "-" & Me.GetIthRow(i).GetSize.Value2 & "-" & Me.GetIthRow(i).GetStyle.Value2
    Next i
    
    Let Me.GetClassifications.Value2 = Application.Transpose(Results)
End Sub

' This returns the private dictionary instance used to store the database of row
' positions for the SEDOLs in the private file.
Property Get GetSedolRowIndicesDictionary() As Dictionary
    Set GetSedolRowIndicesDictionary = Me.SedolRowIndicesDictionary
End Property

' This returns the private dictionary instance used to store the database of row
' positions for the primary Bloomberg tickers in the private file.
Property Get GetPrimaryBloombergTickerRowIndicesDictionary() As Dictionary
    Set GetPrimaryBloombergTickerRowIndicesDictionary = Me.PrimaryBloombergTickerRowIndicesDictionary
End Property

' This returns the private dictionary instance used to store the database of row
' positions for the Bloomberg tickers in the private file.
Property Get GetBloombergTickerRowIndicesDictionary() As Dictionary
    Set GetBloombergTickerRowIndicesDictionary = Me.BlooombergTickerRowIndicesDictionary
End Property

' Returns the row index in RangeDb of the security with the given 7-char SEDOL
' Rows indices for our private file begin with 1 for the first row below the header row.
' A returned value of 0 indicates the given SEDOL is not in the private file.
Property Get GetSedolRowIndex(TheSedol As String) As Long
    Dim TheRowIndex As Long
    
    Let TheRowIndex = Me.GetSedolRowIndicesDictionary.Item(TheSedol)
    
    If TheRowIndex = Empty Then
        Let GetSedolRowIndex = 0
    Else
        Let GetSedolRowIndex = CLng(TheRowIndex)
    End If
End Property

' Returns the row index in RangeDb of the security with the given primary Bloomberg ticker
' Rows indices for our private file begin with 1 for the first row below the header row.
' A returned value of 0 indicates the given primary BB ticker is not in the private file.
Property Get GetPrimaryBloombergTickerIndex(TheBbTicker As String) As Long
    Dim TheRowIndex As Long

    Let TheRowIndex = Me.GetPrimaryBloombergTickerRowIndicesDictionary.Item(TheBbTicker)
    
    If TheRowIndex = Empty Then
        Let GetPrimaryBloombergTickerIndex = 0
    Else
        Let GetPrimaryBloombergTickerIndex = CLng(TheRowIndex)
    End If
End Property

' Returns the row index in RangeDb of the security with the given primary Bloomberg ticker
' Rows indices for our private file begin with 1 for the first row below the header row.
' A returned value of 0 indicates the given primary BB ticker is not in the private file.
Property Get GetBloombergTickerIndex(TheBbTicker As String) As Long
    Dim TheRowIndex As Long

    Let TheRowIndex = Me.GetBloombergTickerRowIndicesDictionary.Item(TheBbTicker)
    
    If TheRowIndex = Empty Then
        Let GetBloombergTickerIndex = 0
    Else
        Let GetBloombergTickerIndex = CLng(TheRowIndex)
    End If
End Property

' Returns a PrivateFileRecord instance for the private file row corresponding to the security with the given primary Bloomberg Ticker
Property Get GetRowWithPrimaryBloombergTicker(TheBbTicker As String) As PrivateFileRow
    Dim TheRowIndex As Long
    
    Let TheRowIndex = Me.GetPrimaryBloombergTickerIndex(TheBbTicker)
    
    If TheRowIndex = 0 Then
        Set GetRowWithPrimaryBloombergTicker = New PrivateFileRow
    Else
        Set GetRowWithPrimaryBloombergTicker = Me.GetIthRow(TheRowIndex)
    End If
End Property

' Returns a PrivateFileRecord instance for the private file row corresponding to the security with the given Bloomberg Ticker
Property Get GetRowWithBloombergTicker(TheBbTicker As String) As PrivateFileRow
    Dim TheRowIndex As Long
    
    Let TheRowIndex = Me.GetBloombergTickerIndex(TheBbTicker)
    
    If TheRowIndex = 0 Then
        Set GetRowWithBloombergTicker = New PrivateFileRow
    Else
        Set GetRowWithBloombergTicker = Me.GetIthRow(TheRowIndex)
    End If
End Property

' Returns a PrivateFileRecord instance for the private file row corresponding to the security with the given 7-char SEDOL
Property Get GetRowWithSedol(TheSedol As String) As PrivateFileRow
    Dim TheRowIndex As Long
    
    Let TheRowIndex = Me.GetSedolRowIndex(TheSedol)
    
    If TheRowIndex = 0 Then
        Set GetRowWithSedol = New PrivateFileRow
    Else
        Set GetRowWithSedol = Me.GetIthRow(TheRowIndex)
    End If
End Property

' Returns True if there is a row with the given SEDOL in the private file. Otherwise, it returns False.
Property Get SedolExists(TheSedol As String) As Boolean
    If Me.GetSedolRowIndex(TheSedol) = 0 Then
        Let SedolExists = False
    Else
        Let SedolExists = True
    End If
End Property

' Returns True if there is a row with the given primary Bloomberg ticker in the private file. Otherwise, it returns False.
Property Get PrimaryBloombergTickerExists(TheBloombergTicker As String) As Boolean
    If Me.GetPrimaryBloombergTickerIndex(TheBloombergTicker) = 0 Then
        Let PrimaryBloombergTickerExists = False
    Else
        Let PrimaryBloombergTickerExists = True
    End If
End Property

' Returns True if there is a row with the given primary Bloomberg ticker in the private file. Otherwise, it returns False.
Property Get BloombergTickerExists(TheBloombergTicker As String) As Boolean
    If Me.GetBloombergTickerIndex(TheBloombergTicker) = 0 Then
        Let BloombergTickerExists = False
    Else
        Let BloombergTickerExists = True
    End If
End Property

' Getters
' This returns the range holding the asset's value. This could be useful when trying to get the address
' of the range.
Property Get GetRange() As Range
    Set GetRange = RangeDb
End Property

Property Get GetDbLength() As Long
    If Me.GetRange Is Nothing Then
        Let GetDbLength = 0
    Else
        Let GetDbLength = Me.GetRange.Rows.Count
    End If
End Property

Property Get GetNumberOfColumns() As Integer
    Dim ARow As New PrivateFileRow

    Let GetNumberOfColumns = ARow.GetNumberOfColumns
End Property

Property Get GetDates() As Range
    Set GetDates = Me.GetRange.Columns(1).Cells
End Property

Property Get GetCountryOfDomiciles() As Range
    Set GetCountryOfDomiciles = Me.GetRange.Columns(2).Cells
End Property

Property Get GetRanks() As Range
    Set GetRanks = Me.GetRange.Columns(3).Cells
End Property

Property Get GetUsdMarketCaps() As Range
    Set GetUsdMarketCaps = Me.GetRange.Columns(4).Cells
End Property

Property Get GetNames() As Range
    Set GetNames = Me.GetRange.Columns(5).Cells
End Property

Property Get GetPiamSectors() As Range
    Set GetPiamSectors = Me.GetRange.Columns(6).Cells
End Property

Property Get GetFy1Pes() As Range
    Set GetFy1Pes = Me.GetRange.Columns(7).Cells
End Property

Property Get GetIndustryCodes() As Range
    Set GetIndustryCodes = Me.GetRange.Columns(8).Cells
End Property

Property Get GetNetAlphas() As Range
    Set GetNetAlphas = Me.GetRange.Columns(9).Cells
End Property

Property Get GetBetas() As Range
    Set GetBetas = Me.GetRange.Columns(10).Cells
End Property

Property Get GetLongTermForwardGrowths() As Range
    Set GetLongTermForwardGrowths = Me.GetRange.Columns(11).Cells
End Property

Property Get GetDividendYields() As Range
    Set GetDividendYields = Me.GetRange.Columns(12).Cells
End Property

Property Get GetSharesOutstanding() As Range
    Set GetSharesOutstanding = Me.GetRange.Columns(13).Cells
End Property

Property Get GetIsins() As Range
    Set GetIsins = Me.GetRange.Columns(14).Cells
End Property

Property Get GetAverageOneMonthTradingVolumes() As Range
    Set GetAverageOneMonthTradingVolumes = Me.GetRange.Columns(15).Cells
End Property

Property Get GetUsdTradeLotSizes() As Range
    Set GetUsdTradeLotSizes = Me.GetRange.Columns(16).Cells
End Property

Property Get GetTradableCountryFlags() As Range
    Set GetTradableCountryFlags = Me.GetRange.Columns(17).Cells
End Property

Property Get GetBloombergTickers() As Range
    Set GetBloombergTickers = Me.GetRange.Columns(18).Cells
End Property

Property Get GetInvestableSedols() As Range
    Set GetInvestableSedols = Me.GetRange.Columns(19).Cells
End Property

Property Get GetInvestableBloombergTickers() As Range
    Set GetInvestableBloombergTickers = Me.GetRange.Columns(20).Cells
End Property

Property Get GetInvestableIsins() As Range
    Set GetInvestableIsins = Me.GetRange.Columns(21).Cells
End Property

Property Get GetSedols() As Range
    Set GetSedols = Me.GetRange.Columns(22).Cells
End Property

Property Get GetPrimaryBloombergTickers() As Range
    Set GetPrimaryBloombergTickers = Me.GetRange.Columns(23).Cells
End Property

Property Get GetClassifications() As Range
    Set GetClassifications = Me.GetRange.Columns(24).Cells
End Property

Property Get GetRegions() As Range
    Set GetRegions = GetRange.Columns(25).Cells
End Property

Property Get GetSizes() As Range
    Set GetSizes = GetRange.Columns(26).Cells
End Property

Property Get GetStyles() As Range
    Set GetStyles = GetRange.Columns(27).Cells
End Property
