VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BrokerAllocation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Base 1
Option Explicit

' Private instance variable to hold the portfolio's data
Public RangeDb As Range
Public UnderlyingWorkbook As Workbook

' Custom initializer. Parameter aRange must include a header row.  There is no attempt to consolidate
' multiple occurences of the same security
Public Sub InitializeWithRange(ARange As Range)
    Dim i As Integer
    Dim obj As New BrokerAllocationRow
    
    If Not (ARange Is Nothing) Then
        ' Set UnderlyingWorkbook
        Set Me.UnderlyingWorkbook = ARange.Worksheet.Parent
    
        ' As long as aRange has more than one row, initialize RangeDb
        If ARange.Rows.Count > 1 Then
            ' Exclude headers row
            Set Me.RangeDb = ARange.Offset(1, 0).Resize(ARange.Rows.Count - 1, obj.GetNumberOfColumns)
            
            ' Format underlying range
            Call Me.FormatUnderlyingRange
        End If
    End If
End Sub

' Custom Initializer.  This initializes the object with the latest data (latest date and time) for a given run number (if provided)
' or the latest date and time in the latest run number.
Public Sub InitializeWithMySql(ServerName As String, DatabaseName As String, TableName As String, UserName As String, _
                                Password As String, AssetClass As String, Strategy As String, Optional RunNumber As Variant)
    Dim TheQuery As String
    Dim TheResults As Variant
    
    ' Set the query to determine the latest portfolio available on the given date for this client and asset class
    If IsMissing(RunNumber) Then
        Let TheQuery = TheQuery & "SELECT D.* FROM "
        Let TheQuery = TheQuery & "(SELECT @MaxDate:=MAX(`date`) FROM `" & DatabaseName & "`.`" & TableName & "` WHERE `assetclasscode` = '" & AssetClass & "' AND `strategycode` = '" & Strategy & "') AS B, "
        Let TheQuery = TheQuery & "(SELECT @MaxTime:=MAX(`time`) FROM `" & DatabaseName & "`.`" & TableName & "` WHERE `date` = @MaxDate AND `assetclasscode` = '" & AssetClass & "' AND `strategycode` = '" & Strategy & "') AS C, "
        Let TheQuery = TheQuery & "`" & DatabaseName & "`.`" & TableName & "` AS D "
        Let TheQuery = TheQuery & " WHERE "
        Let TheQuery = TheQuery & "`date` = @MaxDate AND `time` = @MaxTime AND `strategycode` = '" & Strategy & "'"
        Let TheQuery = TheQuery & " AND `assetclasscode` = '" & AssetClass & "';"
    Else
        Let TheQuery = TheQuery & "SELECT D.* FROM "
        Let TheQuery = TheQuery & "(SELECT @MaxDate:=MAX(`date`) FROM `" & DatabaseName & "`.`" & TableName & "` WHERE `runnumber` = " & RunNumber & " AND `assetclasscode` = '" & AssetClass & "' AND `strategycode` = '" & Strategy & "') AS B, "
        Let TheQuery = TheQuery & "(SELECT @MaxTime:=MAX(`time`) FROM `" & DatabaseName & "`.`" & TableName & "` WHERE `runnumber` = " & RunNumber & " AND `date` = @MaxDate AND `assetclasscode` = '" & AssetClass & "' AND `strategycode` = '" & Strategy & "') AS C, "
        Let TheQuery = TheQuery & "`" & DatabaseName & "`.`" & TableName & "` AS D "
        Let TheQuery = TheQuery & " WHERE "
        Let TheQuery = TheQuery & "`runnumber` = " & RunNumber & " AND `date` = @MaxDate AND `time` = @MaxTime AND `strategycode` = '" & Strategy & "'"
        Let TheQuery = TheQuery & " AND `assetclasscode` = '" & AssetClass & "';"
    End If
    
    ' Select the data and get the latest time available for this date
    Let TheResults = ConnectAndSelect(TheQuery, DatabaseName, ServerName, UserName, Password)
    
    If GetNumberOfRows(TheResults) > 1 Then
        ' Create a new workbook to hold data and dump data in a worksheet
        Set UnderlyingWorkbook = Application.Workbooks.Add
        Call RemoveAllOtherWorksheets(UnderlyingWorkbook.Worksheets(1))
        Call DumpInSheet(TheResults, UnderlyingWorkbook.Worksheets(1).Range("A1"))
        
        ' Initialize with the data now held in the range
        Call Me.InitializeWithRange(UnderlyingWorkbook.Worksheets(1).Range("A1").CurrentRegion)
    End If
End Sub

' This injects this instance into the DB
Public Sub InjectIntoMySql(ServerName As String, DatabaseName As String, TableName As String, UserName As String, Password As String)
    Dim TheArray As Variant
    Dim TheHeaders As Variant

    If Me.GetRange.Rows.Count > 0 Then
        ' Insert the current date and time into the first two columns
        Let Me.GetDates.Value2 = ConvertDateToSerial(Date)
        Let Me.GetTimes.Value2 = ConvertTimeToSerial(Time)

        ' Get the header row.  Ignore the last column (primarykey)
        Let TheHeaders = Most(ConvertTo1DArray(AddSingleBackQuotesToAllArrayElements(GetRow(ConnectAndSelect("SELECT * FROM `" & DatabaseName & "`.`" & TableName & "`;", DatabaseName, ServerName, UserName, Password), 1))))
        ' Get all columns but the last one (primarykey)
        Let TheArray = DoubleQuote2DArray(Me.GetRange.Resize(Me.GetDbLength, Me.GetNumberOfColumns - 1).Value2)
        
        ' Connect to DB and inject into DB
        Call ConnectAndExecuteInsertQuery(TheArray, TheHeaders, TableName, ServerName, DatabaseName, UserName, Password)
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
    If Me.GetRange Is Nothing Then
        Let GetDbLength = 0
    Else
        Let GetDbLength = Me.GetRange.Rows.Count
    End If
End Property

Property Get GetNumberOfColumns() As Integer
    Dim obj As New BrokerAllocationRow

    Call obj.InitializeWithRange(ThisWorkbook.Worksheets("TempComputation").Range("A1"))
    
    Let GetNumberOfColumns = obj.GetNumberOfColumns
End Property

Property Get GetCurrencyCodes() As Range
    Dim obj As New BrokerAllocationRow

    Call obj.InitializeWithRange(ThisWorkbook.Worksheets("TempComputation").Range("A1"))
    
    Set GetCurrencyCodes = Me.GetRange.Columns(obj.GetCurrencyCode.Column).Cells
End Property

Property Get GetBloombergExchangeCodes() As Range
    Dim obj As New BrokerAllocationRow

    Call obj.InitializeWithRange(ThisWorkbook.Worksheets("TempComputation").Range("A1"))
    
    Set GetBloombergExchangeCodes = Me.GetRange.Columns(obj.GetBloombergExchangeCode.Column).Cells
End Property

Property Get GetBrokerCodes() As Range
    Dim obj As New BrokerAllocationRow

    Call obj.InitializeWithRange(ThisWorkbook.Worksheets("TempComputation").Range("A1"))
    
    Set GetBrokerCodes = Me.GetRange.Columns(obj.GetBrokerCode.Column).Cells
End Property

Property Get GetAllocationPercentages() As Range
    Dim obj As New BrokerAllocationRow

    Call obj.InitializeWithRange(ThisWorkbook.Worksheets("TempComputation").Range("A1"))
    
    Set GetAllocationPercentages = Me.GetRange.Columns(obj.GetAllocationPercentage.Column).Cells
End Property

Property Get GetAssetClassCodes() As Range
    Dim obj As New BrokerAllocationRow

    Call obj.InitializeWithRange(ThisWorkbook.Worksheets("TempComputation").Range("A1"))
    
    Set GetAssetClassCodes = Me.GetRange.Columns(obj.GetAssetClassCode.Column).Cells
End Property

Property Get GetStrategyCodes() As Range
    Dim obj As New BrokerAllocationRow

    Call obj.InitializeWithRange(ThisWorkbook.Worksheets("TempComputation").Range("A1"))
    
    Set GetStrategyCodes = Me.GetRange.Columns(obj.GetStrategyCode.Column).Cells
End Property

Property Get GetDates() As Range
    Dim obj As New BrokerAllocationRow

    Call obj.InitializeWithRange(ThisWorkbook.Worksheets("TempComputation").Range("A1"))
    
    Set GetDates = Me.GetRange.Columns(obj.GetDate.Column).Cells
End Property

Property Get GetTimes() As Range
    Dim obj As New BrokerAllocationRow

    Call obj.InitializeWithRange(ThisWorkbook.Worksheets("TempComputation").Range("A1"))
    
    Set GetTimes = Me.GetRange.Columns(obj.GetTime.Column).Cells
End Property

Property Get GetRunNumbers() As Range
    Dim obj As New BrokerAllocationRow

    Call obj.InitializeWithRange(ThisWorkbook.Worksheets("TempComputation").Range("A1"))
    
    Set GetRunNumbers = Me.GetRange.Columns(obj.GetRunNumber.Column).Cells
End Property

' Returns the ith asset in the portfolio
Property Get GetIthRow(i As Long) As BrokerAllocationRow
    Dim obj As New BrokerAllocationRow
    
    ' Initialize the object with the range where the object is held
    Call obj.InitializeWithRange(Me.GetRange.Rows(1))
    
    ' return a reference to the object
    Set GetIthRow = obj
End Property

' This function returns a 1D array with a Portfolio's headers
Public Function GetHeaders() As Variant
    Dim obj As New BrokerAllocationRow
    
    Let GetHeaders = obj.GetHeaders
End Function

Public Sub FormatUnderlyingRange()
    Dim ColumnPositions As Variant
    Dim ColPos As Variant
    Dim obj As New BrokerAllocationRow
    
    ' Initialize the object with the range where the object is held
    Call obj.InitializeWithRange(ThisWorkbook.Worksheets("TempComputation").Range("A1"))
    
    ' Set array of columns to center
    Let ColumnPositions = CreateSequentialArray(1, obj.GetNumberOfColumns)
    For Each ColPos In ColumnPositions
        Let Me.GetRange.Columns(ColPos).EntireColumn.HorizontalAlignment = xlCenter
    Next
End Sub


