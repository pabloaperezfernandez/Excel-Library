VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PostTradingPortfolio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Option Base 1

Public UnderlyingWorkbook As Workbook
Public RangeDb As Range

' This destructor closes (without saving changes) the workbook holding the underlying data
' This is the class destructor.
Public Sub Class_Terminate()
    If Not (Me.UnderlyingWorkbook Is Nothing) Then
        If Not Me.UnderlyingWorkbook Is ThisWorkbook Then
            Call Me.UnderlyingWorkbook.Close(SaveChanges:=False)
        End If
    End If
End Sub

Property Get GetRange() As Range
    Set GetRange = Me.RangeDb
End Property

Property Get GetDbLength() As Long
    If GetRange Is Nothing Then
        Let GetDbLength = 0
    Else
        Let GetDbLength = Me.GetRange.Rows.Count
    End If
End Property

Property Get GetNumberOfColumns() As Integer
    Dim aPostTradingPortfolioAsset As New PostTradingPortfolioAsset
    
    Let GetNumberOfColumns = GetArrayLength(aPostTradingPortfolioAsset.GetHeaders())
End Property

Property Get GetSecurityNames() As Range
    Set GetSecurityNames = Me.GetRange.Columns(1).Cells

End Property

Property Get GetBloombergTickers() As Range
    Set GetBloombergTickers = Me.GetRange.Columns(2).Cells
End Property

Property Get GetFinalUnits() As Range
    Set GetFinalUnits = Me.GetRange.Columns(3).Cells
End Property

Property Get GetFinalUsdValues() As Range
    Set GetFinalUsdValues = Me.GetRange.Columns(4).Cells
End Property

Property Get GetWeights() As Range
    Set GetWeights = Me.GetRange.Columns(5).Cells
End Property

Property Get GetFundIds() As Range
    Set GetFundIds = Me.GetRange.Columns(6).Cells
End Property

Property Get GetAssetClasses() As Range
    Set GetAssetClasses = Me.GetRange.Columns(7).Cells
End Property

Property Get GetStrategies() As Range
    Set GetStrategies = Me.GetRange.Columns(8).Cells
End Property

Property Get GetCreationDates() As Range
    Set GetCreationDates = Me.GetRange.Columns(9).Cells
End Property

Property Get GetCreationTimes() As Range
    Set GetCreationTimes = Me.GetRange.Columns(10).Cells
End Property

Property Get GetDates() As Range
    Set GetDates = Me.GetRange.Columns(11).Cells
End Property

Property Get GetTimes() As Range
    Set GetTimes = Me.GetRange.Columns(12).Cells
End Property

' This injects this instance into the DB
Public Sub InjectIntoMySql(ServerName As String, DatabaseName As String, TableName As String, UserName As String, Password As String)
    Dim TheArray As Variant
    Dim TheHeaders As Variant

    If Me.GetDbLength > 0 Then
        ' Pull instances data into an array
        Let TheArray = Me.GetRange.Value2
        
        ' Add quotes where needed based on datatype and pull NULLs in empty cells
        Let TheArray = DoubleQuote2DArray(TheArray)
        
        ' Get the list of headers.  Most() is used to drop the last column (`PrimaryKey` - PK, AI)
        Let TheHeaders = Most(ConvertTo1DArray(AddSingleBackQuotesToAllArrayElements(GetRow(ConnectAndSelect("SELECT * FROM `" & DatabaseName & "`.`" & TableName & "`;", DatabaseName, ServerName, UserName, Password), 1))))
        
        ' Connect to DB and inject into DB
        Call ConnectAndExecuteInsertQuery(TheArray, TheHeaders, TableName, ServerName, DatabaseName, UserName, Password)
    End If
End Sub

Public Sub InitializeWithFiles(FileNameList As Variant)
    Dim TheResults As Variant
    Dim WshtArray() As Worksheet
    Dim N As Integer
    Dim LastRow As Long
    Dim FileClient As String
    Dim FileAssetClass As String
    Dim FileStrategy As String
    Dim FileDate As String
    Dim filetime As String
    Dim WshtResults As Worksheet
    
    ' Pre-allocate the array used to hold the references of the worksheets holding the post-trading portfolios to consolidate
    ReDim WshtArray(1 To GetArrayLength(FileNameList))
    
    ' Instantiate a worksheet in a new workbook to hold the data underlying this instance
    Set WshtResults = Application.Workbooks.Add.Worksheets(1)
    Call RemoveAllOtherWorksheets(WshtResults)
    
    ' Set UnderlyingWorkbook to point to the workbook holding WshtResults
    Set UnderlyingWorkbook = WshtResults.Parent

    For N = 1 To GetArrayLength(FileNameList)
        ' Extract data embedded in the filename
        Let FileClient = Split(Split(ExtractFilenameAndExtensionFromFullPathFileName(CStr(FileNameList(N))), ".")(0), "-")(1)
        Let FileAssetClass = Split(Split(ExtractFilenameAndExtensionFromFullPathFileName(CStr(FileNameList(N))), ".")(0), "-")(2)
        Let FileStrategy = Split(Split(ExtractFilenameAndExtensionFromFullPathFileName(CStr(FileNameList(N))), ".")(0), "-")(3)
        Let FileDate = Split(Split(ExtractFilenameAndExtensionFromFullPathFileName(CStr(FileNameList(N))), ".")(0), "-")(4)
        Let filetime = Split(Split(ExtractFilenameAndExtensionFromFullPathFileName(CStr(FileNameList(N))), ".")(0), "-")(5)

        ' Open trade list file and set reference to the post-trading portfolio worksheet
        Set WshtArray(N) = Application.Workbooks.Open(Filename:=CStr(FileNameList(N)), UpdateLinks:=False, ReadOnly:=True).Worksheets("PostTradingPortfolio")
        
        ' Convert everything to values to avoid problems
        Let WshtArray(N).UsedRange.Value2 = WshtArray(N).UsedRange.Value2
        
        ' Find out how many rows there are
        Let LastRow = LastNonBlankRowIndexInColumnNotEqualTo(WshtArray(N).Range("A:A"), "")
        
        ' Insert the customer's account number
        Let WshtArray(N).Range("S3").Value2 = "FundId"
        Let WshtArray(N).Range("S4").Font.Bold = True
        Let WshtArray(N).Range("S4").Resize(LastRow - 3, 1).Value2 = FileClient
        
        ' Insert column with file's asset class
        Let WshtArray(N).Range("T3").Value2 = "AssetClass"
        Let WshtArray(N).Range("T4").Font.Bold = True
        Let WshtArray(N).Range("T4").Resize(LastRow - 3, 1).Value2 = FileAssetClass
        
        ' Insert column with file's strategy
        Let WshtArray(N).Range("U3").Value2 = "Strategy"
        Let WshtArray(N).Range("U4").Font.Bold = True
        Let WshtArray(N).Range("U4").Resize(LastRow - 3, 1).Value2 = FileStrategy
        
        ' Insert column with file's date
        Let WshtArray(N).Range("V3").Value2 = "CreationDate"
        Let WshtArray(N).Range("V4").Font.Bold = True
        Let WshtArray(N).Range("V4").Resize(LastRow - 3, 1).Value2 = FileDate
        
        ' Insert column with file's time
        Let WshtArray(N).Range("W3").Value2 = "CreationTime"
        Let WshtArray(N).Range("W4").Font.Bold = True
        Let WshtArray(N).Range("W4").Resize(LastRow - 3, 1).Value2 = filetime
        
        ' Insert column with DB-injection date
        Let WshtArray(N).Range("X3").Value2 = "CreationTime"
        Let WshtArray(N).Range("X4").Font.Bold = True
        Let WshtArray(N).Range("X4").Resize(LastRow - 3, 1).Value2 = ConvertDateToSerial(Date)
        
        ' Insert column with DB-injection time
        Let WshtArray(N).Range("Y3").Value2 = "CreationTime"
        Let WshtArray(N).Range("Y4").Font.Bold = True
        Let WshtArray(N).Range("Y4").Resize(LastRow - 3, 1).Value2 = ConvertTimeToSerial(Time)
        
        ' Delete columns that are not needed
        Call WshtArray(N).Range("R:R").Delete
        Call WshtArray(N).Range("P:P").Delete
        Call WshtArray(N).Range("L:N").Delete
        Call WshtArray(N).Range("H:J").Delete
        Call WshtArray(N).Range("E:G").Delete
        Call WshtArray(N).Range("A:B").Delete

        ' Delete rows 1 and 2
        Call WshtArray(N).Range("1:2").Delete
    Next N
        
    ' Consolidate the worksheets
    Let TheResults = ConsolidateWorksheets(WshtArray, StartingRow:=2)
    
    ' Close the sourceworkbooks
    For N = 1 To GetArrayLength(FileNameList)
        Call WshtArray(N).Parent.Close(SaveChanges:=False)
    Next N
    
    ' Inject consolidated data in the worksheet created to hold the data
    Call DumpInTempPositionWithoutFirstClearing(TheResults, WshtResults.Range("A1"))
    
    ' Format underlying range properly
    Let WshtResults.Range("A:A").NumberFormat = "@"
    Let WshtResults.Range("B:B").NumberFormat = "@"
    Let WshtResults.Range("C:C").NumberFormat = "#,##0.00"
    Let WshtResults.Range("D:D").NumberFormat = "$#,##0_);($#,##0)"
    Let WshtResults.Range("F:F").NumberFormat = "0"
    Let WshtResults.Range("E:E").NumberFormat = "0.00%"
    Let WshtResults.Range("I:J").NumberFormat = "0"
    
    ' Autofit the underlying range
    Call WshtResults.Range("A1").CurrentRegion.EntireColumn.AutoFit
    
    ' Center columns B through J horizontally
    Let WshtResults.Range("B:J").EntireColumn.HorizontalAlignment = xlCenter
    
    ' Initialize the object with the range
    Call Me.InitializeWithRange(WshtResults.Range("A1").CurrentRegion)
End Sub

Public Sub InitializeWithRange(ARange As Range)
    Dim SourceRange As Range
    
    Set SourceRange = ARange.Worksheet.Range("A1").CurrentRegion

    If SourceRange.Rows.Count < 2 Then
        Exit Sub
    End If
    
    Set Me.RangeDb = SourceRange.Range("A2").Resize(SourceRange.Rows.Count - 1, Me.GetNumberOfColumns)
End Sub

' Parameter FundId can be either any valid account identifier or "ALL".
' If omitted, FundId = "ALL" is used by default.
' If ALL is passed, this instance is initialized with all accounts available on the requerested date.
' Each account is loaded using the latest `creationtime` on the requested `creationdate` (which must match param TradeDate)
Public Sub InitializeWithMySql(ServerName As String, DatabaseName As String, TableName As String, UserName As String, _
                               Password As String, AssetClass As String, Strategy As String, TradeDate As Long, Optional FundId As Variant)
    Dim TheQuery As String
    Dim TheResults As Variant
        
    ' Set the query to determine the latest portfolio available on the given date for this client and asset class
    If IsMissing(FundId) Then
        Let TheQuery = "SELECT * FROM " & vbCrLf
        Let TheQuery = TheQuery & "`" & DatabaseName & "`.`" & TableName & "` AS A," & vbCrLf
        Let TheQuery = TheQuery & "(SELECT `fundid`, MAX(`creationtime`) AS `latestcreationtime` FROM `" & DatabaseName & "`.`" & TableName & "` WHERE `creationdate` = " & TradeDate & " AND `assetclass` = '" & AssetClass & "' AND `strategy` = '" & Strategy & "' GROUP BY `fundid`) AS B" & vbCrLf
        Let TheQuery = TheQuery & "WHERE" & vbCrLf
        Let TheQuery = TheQuery & "A.`fundid` = B.`fundid` AND" & vbCrLf
        Let TheQuery = TheQuery & "A.`creationdate` = " & TradeDate & " AND" & vbCrLf
        Let TheQuery = TheQuery & "A.`creationtime` = B.`latestcreationtime` AND" & vbCrLf
        Let TheQuery = TheQuery & "A.`assetclass` = '" & AssetClass & "' AND" & vbCrLf
        Let TheQuery = TheQuery & "A.`Strategy` = '" & Strategy & "';"
    Else
        Let TheQuery = "SELECT * FROM " & vbCrLf
        Let TheQuery = TheQuery & "`" & DatabaseName & "`.`" & TableName & "` AS A," & vbCrLf
        Let TheQuery = TheQuery & "(SELECT `fundid`, MAX(`creationtime`) AS `latestcreationtime` FROM `" & DatabaseName & "`.`" & TableName & "` WHERE `creationdate` = " & TradeDate & " AND `assetclass` = '" & AssetClass & "' AND `strategy` = '" & Strategy & "' GROUP BY `fundid`) AS B" & vbCrLf
        Let TheQuery = TheQuery & "WHERE" & vbCrLf
        Let TheQuery = TheQuery & "A.`fundid` = '" & FundId & "' AND" & vbCrLf
        Let TheQuery = TheQuery & "A.`fundid` = B.`fundid` AND" & vbCrLf
        Let TheQuery = TheQuery & "A.`creationdate` = " & TradeDate & " AND" & vbCrLf
        Let TheQuery = TheQuery & "A.`creationtime` = B.`latestcreationtime` AND" & vbCrLf
        Let TheQuery = TheQuery & "A.`assetclass` = '" & AssetClass & "' AND" & vbCrLf
        Let TheQuery = TheQuery & "A.`Strategy` = '" & Strategy & "';"
    End If
    
    ' Select the data and get the latest time available for this date
    Let TheResults = ConnectAndSelect(TheQuery, DatabaseName, ServerName, UserName, Password)
    
    ' Exit if no clients have any data on the requested date
    If GetNumberOfRows(TheResults) < 2 Then
        Exit Sub
    End If
    
    ' Throw away the last column (e.g. PrimaryKey)
    Let TheResults = GetSubMatrix(TheResults, 1, GetNumberOfRows(TheResults), 1, Me.GetNumberOfColumns())
    
    ' Create a new workbook to hold data and dump data in a worksheet
    Set UnderlyingWorkbook = Application.Workbooks.Add
    Call RemoveAllOtherWorksheets(UnderlyingWorkbook.Worksheets(1))

    ' Format columns appropriately before dumping data
    Let UnderlyingWorkbook.Worksheets(1).Range("A:A").NumberFormat = "@"
    Let UnderlyingWorkbook.Worksheets(1).Range("B:B").NumberFormat = "@"
    Let UnderlyingWorkbook.Worksheets(1).Range("C:C").NumberFormat = "#,##0.00"
    Let UnderlyingWorkbook.Worksheets(1).Range("D:D").NumberFormat = "$#,##0_);($#,##0)"
    Let UnderlyingWorkbook.Worksheets(1).Range("F:F").NumberFormat = "0"
    Let UnderlyingWorkbook.Worksheets(1).Range("E:E").NumberFormat = "0.00%"
    Let UnderlyingWorkbook.Worksheets(1).Range("I:J").NumberFormat = "0"
    Call UnderlyingWorkbook.Worksheets(1).Range("A1").CurrentRegion.EntireColumn.AutoFit

    Call DumpInTempPositionWithoutFirstClearing(TheResults, UnderlyingWorkbook.Worksheets(1).Range("A1"))
    
    ' Initialize with the data now held in the range
    Call Me.InitializeWithRange(UnderlyingWorkbook.Worksheets(1).Range("A1").CurrentRegion)
End Sub

' This method spawns an instance of this class that consolidates each identifier across all fundids
' Column `fundid` is set to "ALL"
Public Function GetConsolidatePostTradingPortfolio() As PostTradingPortfolio
    Dim ConsolidationDict As Dictionary
    Dim i As Integer
    Dim wsht As Worksheet
    Dim AssetData As Variant
    Dim anAsset As PostTradingPortfolioAsset
    Dim ConsolidatedPostTradedPortfolio As PostTradingPortfolio
    Dim TotalUsdMarketValue As Double
    Dim NewWeights() As Double
    
    Set ConsolidatedPostTradedPortfolio = New PostTradingPortfolio
    Set ConsolidationDict = New Dictionary
    
    ' Instantiate a worksheet to hold the consolidate post-trading portfolio
    Set wsht = Application.Workbooks.Add.Worksheets(1)
    Call RemoveAllOtherWorksheets(wsht)
    
    ' Consolidate data
    For i = 1 To Me.GetDbLength
        ' Set a reference to the current asset
        Set anAsset = Me.GetIthAsset(CLng(i))
        
        If ConsolidationDict.Exists(Key:=anAsset.GetBloombergTicker.Value2) Then
            ' Pull the stored, consolidated data
            Let AssetData = ConsolidationDict.Item(Key:=anAsset.GetBloombergTicker.Value2)
            
            ' Update stored, consolidated data with the data from this asset
            Let AssetData(anAsset.GetFinalUnits.Column) = AssetData(anAsset.GetFinalUnits.Column) + anAsset.GetFinalUnits.Value2
            Let AssetData(anAsset.GetFinalUsdValue.Column) = AssetData(anAsset.GetFinalUsdValue.Column) + anAsset.GetFinalUsdValue.Value2
            
            ' Replace current dictionary entry with new one
            Call ConsolidationDict.Remove(Key:=anAsset.GetBloombergTicker.Value2)
            Call ConsolidationDict.Add(Key:=anAsset.GetBloombergTicker.Value2, Item:=AssetData)
        Else
            ' Store the current asset's underlying data
            Call ConsolidationDict.Add(Key:=anAsset.GetBloombergTicker.Value2, Item:=ConvertTo1DArray(anAsset.GetRange.Value2))
        End If
    Next i
    
    ' Get the consolidated portfolio's total USD market value
    ' Application.Transpose(Application.Transpose()) is used as a hack to change indices from 0 to n-1 to 1 to n in ConsolidationDict.Items
    Let TotalUsdMarketValue = Application.Sum(Application.Index(Application.Transpose(Application.Transpose(ConsolidationDict.Items)), 0, anAsset.GetFinalUsdValue.Column))
    
    ' Recompute weights for this consolidated portfolio
    ReDim NewWeights(LBound(ConsolidationDict.Keys) To UBound(ConsolidationDict.Keys))
    For i = LBound(ConsolidationDict.Keys) To UBound(ConsolidationDict.Keys)
        Let NewWeights(i) = ConsolidationDict.Item(ConsolidationDict.Keys(i))(anAsset.GetFinalUsdValue.Column) / TotalUsdMarketValue
    Next i
    
    ' Dump consolidated data into consolidation worksheet
    ' Application.Transpose(Application.Transpose()) is used as a hack to change indices from 0 to n-1 to 1 to n in ConsolidationDict.Items
    Call DumpInTempPositionWithoutFirstClearing(anAsset.GetHeaders, wsht.Range("A1"))
    Call DumpInTempPositionWithoutFirstClearing(Application.Transpose(Application.Transpose(ConsolidationDict.Items)), wsht.Range("A2"))
    
    Set ConsolidatedPostTradedPortfolio.UnderlyingWorkbook = wsht.Parent
    Call ConsolidatedPostTradedPortfolio.InitializeWithRange(wsht.Range("A1").CurrentRegion)
    
    ' Clear out column `creationtime`
    Call ConsolidatedPostTradedPortfolio.GetCreationTimes.ClearContents
    Let ConsolidatedPostTradedPortfolio.GetFundIds.Value2 = "ALL"
    
    ' Drop new weights
    Let ConsolidatedPostTradedPortfolio.GetWeights.Value2 = Application.Transpose(NewWeights)
    
    ' Return instance with consolidated data
    Set GetConsolidatePostTradingPortfolio = ConsolidatedPostTradedPortfolio
End Function

' Returns the ith asset in the portfolio
Property Get GetIthAsset(i As Long) As PostTradingPortfolioAsset
    Dim anAsset As PostTradingPortfolioAsset
    
    ' Instantiate the object
    Set anAsset = New PostTradingPortfolioAsset
    
    ' Initialize the object with the range where the object is held
    Call anAsset.InitializeWithRange(Me.GetRange.Rows(i))
    
    ' return a reference to the object
    Set GetIthAsset = anAsset
End Property

' Returns the list of file names with lastest time stamp on the given date that also match the given asset class and strategy
Public Function GetFileList(TheDir As String, TheAssetClass As String, TheStrategy As String, TheDate As Long) As Variant
    Dim FileNameList As Variant
    Dim AFileName As Variant
    Dim FileListDict As Dictionary
    Dim obj As Object
    Dim FileClient As String
    Dim FileAssetClass As String
    Dim FileStrategy As String
    Dim FileDate As String
    Dim filetime As String
    Dim FileExtension As String
    Dim TheKey As String
    Dim FileNames As Variant
    Dim TimeStamps As Variant
    Dim ReconstructedFilenames() As String
    Dim i As Integer
    
    ' Instantiate dictionary to hold chosen files
    Set FileListDict = New Dictionary
    
    ' Get the list of files
    Let FileNameList = GetFileNames(TheDir)
    
    ' Select only the files corresponding to the chosen date
    For Each AFileName In FileNameList
        Let FileExtension = UCase(Right(CStr(AFileName), 4))
    
        If GetArrayLength(Split(AFileName, "-")) = 6 And Left(CStr(AFileName), 10) = "TradeList-" And FileExtension = "XLSX" Then
            Let FileClient = Split(Split(ExtractFilenameAndExtensionFromFullPathFileName(CStr(AFileName)), ".")(0), "-")(1)
            Let FileAssetClass = Split(Split(ExtractFilenameAndExtensionFromFullPathFileName(CStr(AFileName)), ".")(0), "-")(2)
            Let FileStrategy = Split(Split(ExtractFilenameAndExtensionFromFullPathFileName(CStr(AFileName)), ".")(0), "-")(3)
            Let FileDate = Split(Split(ExtractFilenameAndExtensionFromFullPathFileName(CStr(AFileName)), ".")(0), "-")(4)
            Let filetime = Split(Split(ExtractFilenameAndExtensionFromFullPathFileName(CStr(AFileName)), ".")(0), "-")(5)

            If IsNumeric(FileDate) And Len(FileDate) = 8 And IsNumeric(filetime) And Len(filetime) <= 6 Then
                If CLng(TheDate) = CLng(FileDate) And TheAssetClass = FileAssetClass And TheStrategy = FileStrategy Then
                    Let TheKey = FileClient & "-" & FileAssetClass & "-" & FileStrategy & "-" & FileDate
                    
                    If FileListDict.Exists(Key:=TheKey) Then
                        If CLng(filetime) > FileListDict.Item(Key:=TheKey) Then
                            Call FileListDict.Remove(Key:=TheKey)
                            Call FileListDict.Add(Key:=TheKey, Item:=CLng(filetime))
                        End If
                    Else
                        Call FileListDict.Add(Key:=TheKey, Item:=CLng(filetime))
                    End If
                End If
            End If
        End If
    Next

    ' Return an empty array if no matching files were found
    If FileListDict.Count = 0 Then
        Let GetFileList = Array()
        
        Exit Function
    End If
    
    ' Reconstruct filenames to include time stamps
    ReDim ReconstructedFilenames(1 To FileListDict.Count)
    Let FileNames = FileListDict.Keys
    Let TimeStamps = FileListDict.Items
    
    For i = 1 To FileListDict.Count
        Let ReconstructedFilenames(i) = TheDir & "TradeList-" & FileListDict.Keys(i - 1) & "-" & FileListDict.Items(i - 1) & ".xlsx"
    Next i
    
    Let GetFileList = ReconstructedFilenames
End Function

' This predicate determine if the given file name has a properly formatted trade list format
Public Function ValidTradeListFileNameQ(AFileName As String) As Boolean
    Dim FileClient As String
    Dim FileAssetClass As String
    Dim FileStrategy As String
    Dim FileDate As String
    Dim filetime As String
    Dim FileExtension As String
    Dim ValidClients As Variant
    Dim ValidAssetClasses As Variant
    Dim ValidStrategies As Variant

    ' Set default value
    Let ValidTradeListFileNameQ = True

    ' Get the file's extension
    Let FileExtension = UCase(Right(CStr(AFileName), 4))

    If GetArrayLength(Split(AFileName, "-")) = 6 And Left(CStr(AFileName), 10) = "TradeList-" And FileExtension = "XLSX" Then
        Let FileClient = Split(Split(CStr(AFileName), ".")(0), "-")(1)
        Let FileAssetClass = Split(Split(CStr(AFileName), ".")(0), "-")(2)
        Let FileStrategy = Split(Split(CStr(AFileName), ".")(0), "-")(3)
        Let FileDate = Split(Split(CStr(AFileName), ".")(0), "-")(4)
        Let filetime = Split(Split(CStr(AFileName), ".")(0), "-")(5)

        If IsNumeric(FileDate) And Len(FileDate) = 8 And IsNumeric(filetime) And Len(filetime) <= 6 Then
            ' Get the list of approved clients
            Let ValidClients = ConnectAndSelect("SELECT DISTINCT(`fundid`) FROM `etwip2dot0`.`approvedclientcodes`;", "etwip2dot0", DbServerAddress, DbUserName, DbPassword)
            
            If NumberOfDimensions(ValidClients) = 0 Then
                Let ValidClients = Array(ValidClients)
            End If
            
            ' Get the list of approved asset classes
            Let ValidAssetClasses = ConnectAndSelect("SELECT DISTINCT(`assetclasscode`) FROM `etwip2dot0`.`assetclassandstrategycodes`;", "etwip2dot0", DbServerAddress, DbUserName, DbPassword)
            
            If NumberOfDimensions(ValidClients) = 0 Then
                Let ValidClients = Array(ValidClients)
            End If
            
            ' Get the list of approved stragegies
            Let ValidStrategies = ConnectAndSelect("SELECT DISTINCT(`strategycode`) FROM `etwip2dot0`.`assetclassandstrategycodes` WHERE `assetclasscode` = '" & FileAssetClass & "';", _
                                                   "etwip2dot0", DbServerAddress, DbUserName, DbPassword)
    
            If NumberOfDimensions(ValidClients) = 0 Then
                Let ValidClients = Array(ValidClients)
            End If
            
            ' Verify that client, asset class, and strategy are all valid
            If FreeQ(ValidClients, FileClient) Or FreeQ(ValidAssetClasses, FileAssetClass) Or FreeQ(ValidStrategies, FileStrategy) Then
                Let ValidTradeListFileNameQ = False
            End If
        Else
            Let ValidTradeListFileNameQ = False
        End If
    Else
        Let ValidTradeListFileNameQ = False
    End If
End Function
