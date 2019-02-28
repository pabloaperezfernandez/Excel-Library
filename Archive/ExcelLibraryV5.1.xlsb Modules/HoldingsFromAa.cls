VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HoldingsFromAa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Base 1
Option Explicit

' Private instance variable to hold the portfolio's data
Public RangeDb As Range
' Private instance variable to hold row index of each BloombergTicker in RangeDb
Public BloombergTickerRowIndicesDictionary As Dictionary
' Private instance variable to hold row index of each 7-char SEDOL in RangeDb
Public SedolRowIndicesDictionary As Dictionary
' This is the workbook underlying the data
Public UnderlyingWorkbook As Workbook

' This destructor closes (without saving changes) the workbook holding the underlying data
' This is the class destructor.
Public Sub Class_Terminate()
    If Not (Me.UnderlyingWorkbook Is Nothing) Then
        If Not Me.UnderlyingWorkbook Is ThisWorkbook Then
            Call Me.UnderlyingWorkbook.Close(SaveChanges:=False)
        End If
    End If
End Sub

' Custom initializer. Parameter aRange must include a header row.  There is no attempt to consolidate
' multiple occurences of the same security
Public Sub InitializeWithRange(ARange As Range)
    If ARange Is Nothing Then
        Exit Sub
    End If
        
    ' Set UnderlyingWorkbook
    Set Me.UnderlyingWorkbook = ARange.Worksheet.Parent

    ' As long as aRange has more than one row (the headers row), initialize RangeDb and RowIndicesDictionary
    If ARange.Rows.Count > 1 Then
        ' Exclude headers row
        Set Me.RangeDb = ARange.Offset(1, 0).Resize(ARange.Rows.Count - 1, Me.GetNumberOfColumns)

        ' Initialize RowIndicesDictionary
        Call Me.InitializeRowIndicesDictionaries
    End If
End Sub

' Custom initializer. Parameter aRange must include a header row.  There is no attempt to consolidate
' multiple occurences of the same security
Public Sub InitializeWithTradarView(AFileName As String)
    Dim i As Integer
    Dim SourceRange As Range
    Dim sourceSht As Worksheet
    Dim TradarToDbIndexMapping As Variant
    Dim targetSht As Worksheet
    Dim TargetRange As Range
    Dim PortfolioDate As Long
    Dim anAsset As HoldingsFromAaRow
    Dim TheHeaders As Variant
    Dim NumberRowsBeforeFiltering As Integer
    Dim NumberRowsAfterFiltering As Integer
    
    ' Create a dummy instance HoldingsFromAaRow to use for column positions, etc.
    Set anAsset = New HoldingsFromAaRow
    Call anAsset.InitializeWithRange(ThisWorkbook.Worksheets("TempComputation").Range("A1").Resize(1, anAsset.GetNumberOfColumns))
    
    ' Set the mapping from Tradar's table column position to the DB's table column position
    Let TradarToDbIndexMapping = [{1,10; 2,8; 3,7; 4,6; 5,9; 6, 5; 7, 12; 8, 3; 9, 4; 10, 13; 11, 11}]
    
    ' Open the workbook holding the Tradar view and extract the data
    Set sourceSht = Application.Workbooks.Open(Filename:=AFileName, UpdateLinks:=False, ReadOnly:=True).Worksheets(1)
    
    ' Set aRange to the range holding the underlying data
    Set SourceRange = sourceSht.Range("A1").CurrentRegion
    
    ' Create new workbook to hold the database underlying the DB contents
    Set targetSht = Application.Workbooks.Add.Worksheets(1)
    Call RemoveAllOtherWorksheets(targetSht)
    ' Format target SEDOL column as text
    Let targetSht.Cells(1, anAsset.Get7CharSedol.Column).Resize(SourceRange.Rows.Count, 1).NumberFormat = "@"
    
    ' Move columns from source range to the expected location in the target worksheet
    For i = LBound(TradarToDbIndexMapping, 1) To UBound(TradarToDbIndexMapping, 1)
        Set TargetRange = targetSht.Cells(1, TradarToDbIndexMapping(i, 2)).Resize(SourceRange.Rows.Count, 1)
        Let TargetRange.Value2 = SourceRange.Columns(TradarToDbIndexMapping(i, 1)).Value2
    Next i
    
    ' Copy over the strategy column (e.g. 8) to the sub-strategy column (e.g. 9) if the former is empty.
    If SourceRange.Rows.Count > 1 Then
        If IsEmpty(SourceRange.Cells(2, 9).Value2) Then
            Let targetSht.Cells(2, 4).Resize(SourceRange.Rows.Count - 1, 1).Value2 = _
                targetSht.Cells(2, 3).Resize(SourceRange.Rows.Count - 1, 1).Value2
        End If
    End If
    
    ' Set targetRange before initializing this instance with targetRage
    Set TargetRange = targetSht.Range("A1").Resize(SourceRange.Rows.Count, anAsset.GetNumberOfColumns)
    
    ' Record the number of rows in the source range before filtering our untradable securities
    Let NumberRowsBeforeFiltering = SourceRange.Rows.Count

    ' Close the source workbook now that all data has been transfered to the target workbook
    Call sourceSht.Parent.Close(SaveChanges:=False)

    ' Set UnderlyingWorkbook
    Set Me.UnderlyingWorkbook = TargetRange.Worksheet.Parent

    ' Insert headers in row 1
    Let TargetRange.Rows(1).Value2 = Me.GetHeaders()
        
    ' Fill in the date and time columns with today's date and the current time
    If TargetRange.Rows.Count > 1 Then
        ' Insert current date
        Let TargetRange.Resize(TargetRange.Rows.Count - 1, 1).Offset(1, anAsset.GetDate.Column - 1).Value2 = ConvertDateToSerial(Date)
            
        ' Insert current time
        Let TargetRange.Resize(TargetRange.Rows.Count - 1, 1).Offset(1, anAsset.GetTime.Column - 1).Value2 = ConvertTimeToSerial(Time)
    
        ' Extract portfolio date from the file name
        Let PortfolioDate = CLng(Right(ExtractFileNameWithNoExtensionFromFullPathFileName(AFileName), 8))
    
        ' Insert portfolio date
        Let TargetRange.Resize(TargetRange.Rows.Count - 1, 1).Offset(1, anAsset.GetPortfolioDate.Column - 1).Value2 = PortfolioDate
            
        ' Insert NULL in the classifications column
        Let TargetRange.Resize(TargetRange.Rows.Count - 1, 1).Offset(1, anAsset.GetClassification.Column - 1).Value2 = "NULL"
            
        ' Filter out un-tradable securities
        Let TargetRange.Worksheet.AutoFilterMode = False
        Let TheHeaders = TargetRange.Range("A1").CurrentRegion.Rows(1).Value2
        Call TargetRange.Range("A1").CurrentRegion.AutoFilter(Field:=13, Criteria1:="=1", Operator:=xlOr, Criteria2:="TRUE")
        
        ' Record number of rows after filtering untradable securities
        Let NumberRowsAfterFiltering = TargetRange.Worksheet.AutoFilter.Range.SpecialCells(xlCellTypeVisible).Rows.Count
        
        If NumberRowsAfterFiltering = NumberRowsBeforeFiltering Then
            Exit Sub
        Else
            Call TargetRange.Worksheet.AutoFilter.Range.SpecialCells(xlCellTypeVisible).Delete
            Let TargetRange.Worksheet.AutoFilterMode = False
            Call TargetRange.Worksheet.Range("1:1").Insert
            Call DumpInTempPositionWithoutFirstClearing(TheHeaders, TargetRange.Worksheet.Range("A1"))
            Set TargetRange = TargetRange.Worksheet.Range("A1").CurrentRegion
        
            ' Initialized with range (include header row when calling function)
            Call Me.InitializeWithRange(TargetRange)
        End If
    End If
End Sub

' This injects this instance into the DB
Public Sub InjectIntoMySql(ServerName As String, DatabaseName As String, TableName As String, UserName As String, Password As String)
    Dim TheArray As Variant
    Dim TheHeaders As Variant
    Dim i As Integer

    If Me.GetRange.Rows.Count > 0 Then
        ' Get the data held in the underlying range
        Let TheArray = Me.GetRange.Value2
        
        ' User headers as defined in the DB table, exclusing the last one (e.g. primarykey), which is auto-increment
        Let TheHeaders = Most(ConvertTo1DArray(AddSingleBackQuotesToAllArrayElements(GetRow(ConnectAndSelect("SELECT * FROM `" & DatabaseName & "`.`" & TableName & "`;", DatabaseName, ServerName, UserName, Password), 1))))
        
        ' Add quote to every element of the following columns
        Let TheArray = AddQuotesToAllSubArrayElements(TheArray, 3, 1, Me.GetDbLength, 7)
        Let TheArray = AddQuotesToAllSubArrayElements(TheArray, 12, 1, Me.GetDbLength, 1)
        
        ' Set empty columns equal to NULL
        For i = 1 To Me.GetDbLength
            Let TheArray(i, Me.GetUnavailabilityFlag.Column) = IIf(TheArray(i, Me.GetUnavailabilityFlag.Column) = Empty, "NULL", 1)
        Next i
    
        ' Connect to DB and inject into DB
        Call ConnectAndExecuteInsertQuery(TheArray, TheHeaders, TableName, ServerName, DatabaseName, UserName, Password)
    End If
End Sub

' Custom Initializer.
' If FundId = "ALL", then the entire holdings for the given account are returned on the given asset class, strategy,
' custodian, and date.
Public Sub InitializeWithMySql(ServerName As String, DatabaseName As String, TableName As String, FundId As String, _
                               Custodian As String, UserName As String, Password As String, AssetClass As String, _
                               Strategy As String, ADate As Long)
    Dim TheQuery As String
    Dim TheResults As Variant
    Dim LatestTime As Long
    
    ' Set the query to determine the latest portfolio available on the given date for this client and asset class
    If FundId = "ALL" Then
        Let TheQuery = "SELECT MAX(`time`) FROM `" & DatabaseName & "`.`" & TableName & "` WHERE "
        Let TheQuery = TheQuery & "`portfoliodate` = " & ADate & " AND `custodian` = """
        Let TheQuery = TheQuery & Custodian & """ AND `strategycode` = """ & Strategy & """ AND `assetclass` = """ & AssetClass & """;"
    Else
        Let TheQuery = "SELECT MAX(`time`) FROM `" & DatabaseName & "`.`" & TableName & "` WHERE "
        Let TheQuery = TheQuery & "`portfoliodate` = " & ADate & " AND `fundid` = """ & FundId & """ AND `custodian` = """
        Let TheQuery = TheQuery & Custodian & """ AND `strategycode` = """ & Strategy & """ AND `assetclass` = """ & AssetClass & """;"
    End If
    
    ' Select the data and get the latest time available for this date
    Let TheResults = ConnectAndSelect(TheQuery, DatabaseName, ServerName, UserName, Password)
    
    If Not (IsArray(TheResults)) Then
        Exit Sub
    End If
    
    If GetNumberOfRows(TheResults) > 1 Then
        Let LatestTime = TheResults(2, 1)
    
        ' Set the query to get the latest portfolio on this date
        If FundId = "ALL" Then
            Let TheQuery = "SELECT * FROM `" & DatabaseName & "`.`" & TableName & "` WHERE "
            Let TheQuery = TheQuery & "`portfoliodate` = " & ADate & " AND `custodian` = """ & Custodian & """ "
            Let TheQuery = TheQuery & "AND `strategycode` = """ & Strategy & """ AND `assetclass` = """ & AssetClass & """ AND `time` = " & LatestTime & " AND `holdingsfromaa`.`unavailabilityflag` IS NULL;"
        Else
            Let TheQuery = "SELECT * FROM `" & DatabaseName & "`.`" & TableName & "` WHERE "
            Let TheQuery = TheQuery & "`portfoliodate` = " & ADate & " AND `fundid` = """ & FundId & """ AND `custodian` = """ & Custodian & """ "
            Let TheQuery = TheQuery & "AND `strategycode` = """ & Strategy & """ AND `assetclass` = """ & AssetClass & """ AND `time` = " & LatestTime & " AND `holdingsfromaa`.`unavailabilityflag` IS NULL;"
        End If
        
        ' Select the data
        Let TheResults = ConnectAndSelect(TheQuery, DatabaseName, ServerName, UserName, Password)
                
        ' Create a new workbook to hold data and dump data in a worksheet
        Set UnderlyingWorkbook = Application.Workbooks.Add
        Call RemoveAllOtherWorksheets(UnderlyingWorkbook.Worksheets(1))
        ' Format target column holding 7-char SEDOLS as text
        Let UnderlyingWorkbook.Worksheets(1).Range("A1").Offset(1, 5).Resize(GetNumberOfRows(TheResults) - 1, 1).NumberFormat = "@"
        
        Call DumpInTempPositionWithoutFirstClearing(TheResults, UnderlyingWorkbook.Worksheets(1).Range("A1"))
        
        ' Initialize with the data now held in the range
        Call Me.InitializeWithRange(UnderlyingWorkbook.Worksheets(1).Range("A1").CurrentRegion)
    End If
End Sub

' This initializes the row indices dictionaries.  This sub requires RangeDb to have been set.
' At the moment we initialize dictionaries for 7-char SEDOLs and Bloomberg tickers
Public Sub InitializeRowIndicesDictionaries()
    Dim i As Integer
    Dim anAsset As HoldingsFromAaRow
    
    Set Me.BloombergTickerRowIndicesDictionary = New Dictionary
    Set Me.SedolRowIndicesDictionary = New Dictionary

    If Me.GetDbLength > 0 Then
        Set anAsset = Me.GetIthAsset(1)
    
        For i = 1 To Me.GetRange.Rows.Count
            If Not Me.BloombergTickerRowIndicesDictionary.Exists(Me.GetRange(i, anAsset.GetBloombergTicker.Column).Value2) Then
                Call Me.BloombergTickerRowIndicesDictionary.Add(Key:=Me.GetRange(i, anAsset.GetBloombergTicker.Column).Value2, Item:=i)
            End If

            If Not Me.SedolRowIndicesDictionary.Exists(Me.GetRange(i, anAsset.Get7CharSedol.Column).Value2) Then
                Call Me.SedolRowIndicesDictionary.Add(Key:=Me.GetRange(i, anAsset.Get7CharSedol.Column).Value2, Item:=i)
            End If
        Next i
    End If
End Sub

' Returns the ith asset in the portfolio
Property Get GetIthAsset(i As Long) As HoldingsFromAaRow
    Dim anAsset As HoldingsFromAaRow
    
    Set anAsset = New HoldingsFromAaRow
    
    Call anAsset.InitializeWithRange(Me.GetRange.Rows(i))
    
    Set GetIthAsset = anAsset
End Property

' This returns the private dictionary instance used to store the database of row
' positions for the Bloomberg tickers in the portfolio.
Property Get GetBloombergTickerRowIndicesDb() As Dictionary
    Set GetBloombergTickerRowIndicesDb = Me.BloombergTickerRowIndicesDictionary
End Property

' This returns the private dictionary instance used to store the database of row
' positions for the 7-char SEDOLs in the portfolio.
Property Get GetSedolRowIndicesDb() As Dictionary
    Set GetSedolRowIndicesDb = Me.SedolRowIndicesDictionary
End Property

' Returns the row index in RangeDb of the security with the given Bloomberg Ticker
Property Get GetBloombergTickerRowIndex(TheBbTicker As String)
    Let GetBloombergTickerRowIndex = Me.BloombergTickerRowIndicesDictionary.Item(TheBbTicker)
End Property

' Returns the row index in RangeDb of the security with the given Bloomberg Ticker
Property Get GetSedolRowIndex(TheSedol As String)
    Let GetSedolRowIndex = Me.SedolRowIndicesDictionary.Item(TheSedol)
End Property

' Returns a PrivateFileRecord instance for the private file row corresponding to the security with the given Bloomberg Ticker
Property Get GetAssetWithBloombergTicker(TheBbTicker As String) As HoldingsFromAaRow
    Dim PrivateFileRecordInstance
    
    Set GetAssetWithBloombergTicker = Me.GetIthAsset(Me.GetBloombergTickerRowIndex(TheBbTicker))
End Property

' Returns a PrivateFileRecord instance for the private file row corresponding to the security with the given Bloomberg Ticker
Property Get GetAssetWithSedol(TheSedol As String) As HoldingsFromAaRow
    Dim PrivateFileRecordInstance
    
    Set GetAssetWithSedol = Me.GetIthAsset(Me.GetSedolRowIndex(TheBbTicker))
End Property

' Returns True if the Bloomberg ticker is in the private file. Otherwise, it returns False.
Property Get BloombergTickerExists(TheBloombergTicker As String) As Boolean
    If Me.GetBloombergTickerRowIndex(TheBloombergTicker) = Empty Then
        Let BloombergTickerExists = False
    Else
        Let BloombergTickerExists = True
    End If
End Property

' Returns True if the Bloomberg ticker is in the private file. Otherwise, it returns False.
Property Get SedolExists(TheSedol As String) As Boolean
    If Me.GetSedolRowIndex(TheSedol) = Empty Then
        Let SedolExists = False
    Else
        Let SedolExists = True
    End If
End Property

' Getters
' This returns the range holding the asset's value. This could be useful when trying to get the address
' of the range.
Property Get GetRange() As Range
    Set GetRange = RangeDb
End Property

Property Get GetDbLength() As Long
    If GetRange Is Nothing Then
        Let GetDbLength = 0
    Else
        Let GetDbLength = Me.GetRange.Rows.Count
    End If
End Property

Property Get GetHeaders()
    Dim ARow As HoldingsFromAaRow
    
    Set ARow = New HoldingsFromAaRow
    
    Let GetHeaders = ARow.GetHeaders()
End Property

Property Get GetNumberOfColumns() As Integer
    Dim ARow As HoldingsFromAaRow
    
    Set ARow = New HoldingsFromAaRow

    Let GetNumberOfColumns = ARow.GetNumberOfColumns
End Property

Property Get GetDates() As Range
    Set GetDates = Me.GetRange.Columns(1).Cells
End Property

Property Get GetTimes() As Range
    Set GetTimes = Me.GetRange.Columns(2).Cells
End Property

Property Get GetAssetClasses() As Range
    Set GetAssetClasses = Me.GetRange.Columns(3).Cells
End Property

Property Get GetStrategies() As Range
    Set GetStrategies = Me.GetRange.Columns(4).Cells
End Property

Property Get GetFundIds() As Range
    Set GetFundIds = Me.GetRange.Columns(5).Cells
End Property

Property Get Get7CharSedols() As Range
    Set Get7CharSedols = Me.GetRange.Columns(6).Cells
End Property

Property Get GetIsins() As Range
    Set GetIsins = Me.GetRange.Columns(7).Cells
End Property

Property Get GetBloombergTickers() As Range
    Set GetBloombergTickers = Me.GetRange.Columns(8).Cells
End Property

Property Get GetSecurityNames() As Range
    Set GetSecurityNames = Me.GetRange.Columns(9).Cells
End Property

Property Get GetAmounts() As Range
    Set GetAmounts = Me.GetRange.Columns(10).Cells
End Property

Property Get GetUsdPositionValues() As Range
    Set GetUsdPositionValues = Me.GetRange.Columns(11).Cells
End Property

Property Get GetCustodian() As Range
    Set GetCustodian = Me.GetRange.Columns(12).Cells
End Property

Property Get GetUnavailabilityFlag() As Range
    Set GetUnavailabilityFlag = Me.GetRange.Columns(13).Cells
End Property

Property Get GetPortfolioDates() As Range
    Set GetPortfolioDates = Me.GetRange.Columns(14).Cells
End Property

Property Get GetRunNumbers() As Range
    Set GetRunNumbers = Me.GetRange.Columns(15).Cells
End Property

Property Get GetClassifications() As Range
    Set GetClassifications = Me.GetRange.Columns(16).Cells
End Property

Property Get GetWeights() As Variant
    Dim TheWeights() As Variant
    Dim TotalCapital As Double
    Dim i As Long
    
    If Me.GetDbLength > 1 Then
        ReDim TheWeights(1 To Me.GetDbLength)
        Let TotalCapital = Application.Sum(Me.GetUsdPositionValues)
    
        For i = 1 To Me.GetDbLength
            Let TheWeights(i) = Me.GetIthAsset(i).GetUsdPositionValue.Value2 / TotalCapital
        Next i

        Let GetWeights = TheWeights
    Else
        Let GetWeights = Array()
    End If
End Property



