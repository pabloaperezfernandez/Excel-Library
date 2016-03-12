VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OptimalPortfolio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Base 1
Option Explicit

' Private instance variable to hold the portfolio's data
Public RangeDb As Range
' Private instance variable to hold row index of each SEDOL in RangeDb
Public RowIndicesDictionary As Dictionary
' Workbook holdings this instance's data
Public UnderlyingWorkbook As Workbook

' Custom initializer. Parameter aRange must include a header row.  There is no attempt to consolidate
' multiple occurences of the same security
Public Sub InitializeWithRange(ARange As Range)
    Dim i As Integer
    Dim anAsset As OptimalAsset
    
    Set anAsset = New OptimalAsset

    If Not (ARange Is Nothing) Then
        ' Set UnderlyingWorkbook
        Set Me.UnderlyingWorkbook = ARange.Worksheet.Parent
    
        ' As long as aRange has more than one row, initialize RangeDb and RowIndicesDictionary
        If ARange.Rows.Count > 1 Then
            ' Exclude headers row
            Set Me.RangeDb = ARange.Offset(1, 0).Resize(ARange.Rows.Count - 1, anAsset.GetNumberOfColumns)
    
            ' Initialize RowIndicesDictionary
            Call InitializeRowIndicesDictionary
        End If
    End If
End Sub

' Custom initializer. The source file must contain exactly one worksheet.
' There is no attempt to consolidate multiple occurences of the same security
' The classification column is set to NULL since R&D does not pass this data
Public Sub InitializeWithFile(AFileName As String)
    Dim i As Integer
    Dim SourceRange As Range
    Dim sourceSht As Worksheet
    Dim FileToDbIndexMapping As Variant
    Dim targetSht As Worksheet
    Dim TargetRange As Range
    Dim PortfolioDate As Long
    Dim anAsset As OptimalAsset
    Dim TheAssetClass As String
    Dim TheStrategy As String
    
    ' Instantiate dummy asset to get its properties such as column positions
    Set anAsset = New OptimalAsset
    Call anAsset.InitializeWithRange(ThisWorkbook.Worksheets("TempComputation").Range("A1").Resize(1, anAsset.GetNumberOfColumns))
    
    ' Set the mapping from source optimal table column position to the DB's table column position
    Let FileToDbIndexMapping = [{1,5; 2,6; 3,7; 4,8; 5,9; 6, 11; 7,14; 8, 15}]
    
    ' Get asset class and strategy from name
    Let TheAssetClass = Split(ExtractFilenameAndExtensionFromFullPathFileName(AFileName), "-")(1)
    Let TheStrategy = Split(ExtractFilenameAndExtensionFromFullPathFileName(AFileName), "-")(2)
    
    ' Open the workbook holding the Tradar view and extract the data
    Set sourceSht = Application.Workbooks.Open(Filename:=AFileName, UpdateLinks:=False, ReadOnly:=True).Worksheets(1)
    
    ' Set aRange to the range holding the underlying data
    Set SourceRange = sourceSht.Range("A1").CurrentRegion
    
    ' Create new workbook to hold the database underlying the DB contents
    Set targetSht = Application.Workbooks.Add.Worksheets(1)
    Call RemoveAllOtherWorksheets(targetSht)
    
    ' Format target SEDOL column as text
    Let targetSht.Cells(1, 6).Resize(SourceRange.Rows.Count, 1).NumberFormat = "@"
    
    ' Move columns from source range to the expected location in the target worksheet
    For i = LBound(FileToDbIndexMapping, 1) To UBound(FileToDbIndexMapping, 1)
        Set TargetRange = targetSht.Cells(1, FileToDbIndexMapping(i, 2)).Resize(SourceRange.Rows.Count, 1)
        Let TargetRange.Value2 = SourceRange.Columns(FileToDbIndexMapping(i, 1)).Value2
    Next i
    
    ' Set targetRange before initializing this instance with targetRage
    Set TargetRange = targetSht.Range("A1").Resize(SourceRange.Rows.Count, anAsset.GetNumberOfColumns)

    ' Close the source workbook now that all data has been transfered to the target workbook
    Call sourceSht.Parent.Close(SaveChanges:=False)

    ' Set UnderlyingWorkbook
    Set Me.UnderlyingWorkbook = TargetRange.Worksheet.Parent

    ' Fill in headers row in targetRange
    Let TargetRange.Rows(1).Value2 = Me.GetHeaders()
    
    ' Fill in the date and time columns with today's date, the current time, and the current run number
    If TargetRange.Rows.Count > 1 Then
        ' Insert current date
        Let TargetRange.Resize(TargetRange.Rows.Count - 1, 1).Offset(1, anAsset.GetDate.Column - 1).Value2 = ConvertDateToSerial(Date)
            
        ' Insert current time
        Let TargetRange.Resize(TargetRange.Rows.Count - 1, 1).Offset(1, anAsset.GetTime.Column - 1).Value2 = ConvertTimeToSerial(Time)
    
        ' Extract portfolio date from the file name
        Let PortfolioDate = CLng(Right(ExtractFileNameWithNoExtensionFromFullPathFileName(AFileName), 8))
    
        ' Insert portfolio date
        Let TargetRange.Resize(TargetRange.Rows.Count - 1, 1).Offset(1, anAsset.GetPortfolioDate.Column - 1).Value2 = PortfolioDate
        
        ' Insert the asset class and the strategy
        Let TargetRange.Resize(TargetRange.Rows.Count - 1, 1).Offset(1, anAsset.GetAssetClassCode.Column - 1).Value2 = TheAssetClass
        Let TargetRange.Resize(TargetRange.Rows.Count - 1, 1).Offset(1, anAsset.GetStrategyCode.Column - 1).Value2 = TheStrategy
        
        ' Exclude headers row
        Call Me.InitializeWithRange(TargetRange)
    End If
End Sub

' Custom Initializer.
Public Sub InitializeWithMySql(ServerName As String, DatabaseName As String, TableName As String, UserName As String, _
                                Password As String, AssetClass As String, Strategy As String, ADate As Long)
    Dim TheQuery As String
    Dim TheResults As Variant
    Dim LatestTime As Long
    
    ' Set the query to determine the latest portfolio available on the given date for this client and asset class
    Let TheQuery = "SELECT MAX(`time`) FROM `" & DatabaseName & "`.`" & TableName & "` WHERE "
    Let TheQuery = TheQuery & "`portfoliodate` = " & ADate & " AND "
    Let TheQuery = TheQuery & "`strategycode` = """ & Strategy & """ AND `assetclass` = """ & AssetClass & """;"
    
    ' Select the data and get the latest time available for this date
    Let TheResults = ConnectAndSelect(TheQuery, DatabaseName, ServerName, UserName, Password)
    
    If Not (IsArray(TheResults)) Then
        Exit Sub
    End If
    
    If GetNumberOfRows(TheResults) > 1 Then
        Let LatestTime = TheResults(2, 1)
    
        ' Set the query to get the latest portfolio on this date
        Let TheQuery = "SELECT * FROM `" & DatabaseName & "`.`" & TableName & "` WHERE "
        Let TheQuery = TheQuery & "`portfoliodate` = " & ADate & " "
        Let TheQuery = TheQuery & "AND `strategycode` = """ & Strategy & """ AND `assetclass` = """ & AssetClass & """ AND `time` = " & LatestTime & ";"
    
        ' Select the data
        Let TheResults = ConnectAndSelect(TheQuery, DatabaseName, ServerName, UserName, Password)
        
        ' Create a new workbook to hold data and dump data in a worksheet
        Set UnderlyingWorkbook = Application.Workbooks.Add
        Call RemoveAllOtherWorksheets(UnderlyingWorkbook.Worksheets(1))
        Let UnderlyingWorkbook.Worksheets(1).Columns(6).NumberFormat = "@"
        Call DumpInTempPositionWithoutFirstClearing(TheResults, UnderlyingWorkbook.Worksheets(1).Range("A1"))
        
        ' Initialize with the data now held in the range
        Call Me.InitializeWithRange(UnderlyingWorkbook.Worksheets(1).Range("A1").CurrentRegion)
    End If
End Sub
' This injects this instance into the DB
Public Sub InjectIntoMySql(ServerName As String, DatabaseName As String, TableName As String, UserName As String, Password As String)
    Dim TheArray As Variant
    Dim TheHeaders As Variant

    If Me.GetDbLength > 0 Then
        ' Pull instances data into an array
        Let TheArray = Me.GetRange.Value2
        
        ' Get the list of headers
        Let TheHeaders = Most(ConvertTo1DArray(AddSingleBackQuotesToAllArrayElements(GetRow(ConnectAndSelect("SELECT * FROM `" & DatabaseName & "`.`" & TableName & "`;", DatabaseName, ServerName, UserName, Password), 1))))
        
        ' Add quote to every element of the following columns
        Let TheArray = AddQuotesToAllSubArrayElements(TheArray, 3, 1, Me.GetDbLength, 2)
        Let TheArray = AddQuotesToAllSubArrayElements(TheArray, 6, 1, Me.GetDbLength, 6)
        Let TheArray = AddQuotesToAllSubArrayElements(TheArray, 14, 1, Me.GetDbLength, 1)
        
        ' Connect to DB and inject into DB
        Call ConnectAndExecuteInsertQuery(TheArray, TheHeaders, TableName, ServerName, DatabaseName, UserName, Password)
    End If
End Sub

' This initializes the row indices dictionary.  This sub requires RangeDb to have been set.
Public Sub InitializeRowIndicesDictionary()
    Dim i As Integer
    Dim anAsset As OptimalAsset
    
    Set Me.RowIndicesDictionary = New Dictionary
    
    Set anAsset = New OptimalAsset
    Call anAsset.InitializeWithRange(ThisWorkbook.Worksheets("TempComputation").Range("A1").Resize(1, anAsset.GetNumberOfColumns))

    If Me.GetRange.Rows.Count > 1 Then
        For i = 1 To Me.GetRange.Rows.Count
            If Not Me.RowIndicesDictionary.Exists(Me.GetRange(i, anAsset.GetBloombergTicker.Column).Value2) Then
                Call Me.RowIndicesDictionary.Add(Key:=Me.GetRange(i, anAsset.GetBloombergTicker.Column).Value2, Item:=i)
            End If
        Next i
    End If
End Sub

' This destructor closes (without saving changes) the workbook holding the underlying data
' This is the class destructor.
Public Sub Class_Terminate()
    If Not (Me.UnderlyingWorkbook Is Nothing) Then
        If Not Me.UnderlyingWorkbook Is ThisWorkbook Then
            Call Me.UnderlyingWorkbook.Close(SaveChanges:=False)
        End If
    End If
End Sub

' Getters
' This returns the range holding the asset's value. This could be useful when trying to get the address
' of the range.
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
    Dim AnOptimalAsset As New OptimalAsset
    
    Let GetNumberOfColumns = AnOptimalAsset.GetNumberOfColumns
End Property

Property Get GetDates() As Range
    Set GetDates = Me.GetRange.Columns(1).Cells
End Property

Property Get GetTimes() As Range
    Set GetTimes = Me.GetRange.Columns(2).Cells
End Property

Property Get GetAssetClassCodes() As Range
    Set GetAssetClassCodes = Me.GetRange.Columns(3).Cells
End Property

Property Get GetStrategyCodes() As Range
    Set GetStrategyCodes = Me.GetRange.Columns(4).Cells
End Property

Property Get GetWeights() As Range
    Set GetWeights = Me.GetRange.Columns(5).Cells
End Property

Property Get Get7CharSedols() As Range
    Set Get7CharSedols = Me.GetRange.Columns(6).Cells
End Property

Property Get GetSecurityNames() As Range
    Set GetSecurityNames = Me.GetRange.Columns(7).Cells
End Property

Property Get GetIsins() As Range
    Set GetIsins = Me.GetRange.Columns(8).Cells
End Property

Property Get GetCountryCodes() As Range
    Set GetCountryCodes = Me.GetRange.Columns(9).Cells
End Property

Property Get GetCurrencyCodes() As Range
    Set GetCurrencyCodes = Me.GetRange.Columns(10).Cells
End Property

Property Get GetBloombergTickers() As Range
    Set GetBloombergTickers = Me.GetRange.Columns(11).Cells
End Property

Property Get GetPortfolioDates() As Range
    Set GetPortfolioDates = Me.GetRange.Columns(12).Cells
End Property

Property Get GetRunNumbers() As Range
    Set GetRunNumbers = Me.GetRange.Columns(13).Cells
End Property

Property Get GetClassifications() As Range
    Set GetClassifications = Me.GetRange.Columns(14).Cells
End Property

' Returns the ith asset in the portfolio
Property Get GetIthAsset(i As Long) As OptimalAsset
    Dim anAsset As OptimalAsset
    
    ' Instantiate the object
    Set anAsset = New OptimalAsset
    
    ' Initialize the object with the range where the object is held
    Call anAsset.InitializeWithRange(Me.GetRange.Rows(i))
    
    ' return a reference to the object
    Set GetIthAsset = anAsset
End Property

' This returns the private dictionary instance used to store the database of row
' positions for the SEDOLs in the portfolio.
Property Get GetRowIndicesDb() As Dictionary
    Set GetRowIndicesDb = Me.RowIndicesDictionary
End Property

' Returns the row index in RangeDb of the security with the given ISIN
Property Get GetRowIndex(TheBloombergTicker As String)
    Let GetRowIndex = Me.GetRowIndicesDb.Item(TheBloombergTicker)
End Property

' Returns a PrivateFileRecord instance for the private file row corresponding to the security with the given ISIN
Property Get GetAssetWithBloombergTicker(TheBloombergTicker As String) As OptimalAsset
    Set GetAssetWithBloombergTicker = Me.GetIthAsset(Me.GetRowIndex(TheBloombergTicker))
End Property

' Returns True if the ISIN is in the private file. Otherwise, it returns False.
Property Get BloombergTickerExists(TheBloombergTicker As String) As Boolean
    If Me.GetRowIndex(TheBloombergTicker) = Empty Then
        Let BloombergTickerExists = False
    Else
        Let BloombergTickerExists = True
    End If
End Property

' This function returns a 1D array with a Portfolio's headers
Property Get GetHeaders() As Variant
    Dim anAsset As New OptimalAsset
    
    Let GetHeaders = anAsset.GetHeaders
End Property


