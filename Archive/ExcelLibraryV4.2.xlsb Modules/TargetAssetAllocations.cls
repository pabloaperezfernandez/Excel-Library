VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TargetAssetAllocations"
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
    Dim AnAssetAllocationRow As TargetAssetAllocationRow
    
    Set AnAssetAllocationRow = New TargetAssetAllocationRow

    If Not (ARange Is Nothing) Then
        ' Set UnderlyingWorkbook
        Set Me.UnderlyingWorkbook = ARange.Worksheet.Parent
    
        ' As long as aRange has more than one row, initialize RangeDb
        If ARange.Rows.Count > 1 Then
            ' Exclude headers row
            Set Me.RangeDb = ARange.Offset(1, 0).Resize(ARange.Rows.Count - 1, AnAssetAllocationRow.GetNumberOfColumns)
        End If
    End If
End Sub

' Custom initializer. Parameter aRange must include a header row.
Public Sub InitializeWithFile(AFileName As String)
    Dim i As Integer
    Dim aTargetAllocationsRow As New TargetAssetAllocationRow
    Dim SourceRange As Range
    
    ' Open the workbook holding the Tradar view and extract the data
    Set SourceRange = Application.Workbooks.Open(Filename:=AFileName, UpdateLinks:=False, ReadOnly:=True).Worksheets(1).Range("A1").CurrentRegion

    ' Set UnderlyingWorkbook
    Set Me.UnderlyingWorkbook = SourceRange.Worksheet.Parent

    ' Fill in the date and time columns with today's date and the current time
    If SourceRange.Rows.Count > 1 Then
        Let SourceRange.Resize(SourceRange.Rows.Count - 1, 1).Offset(1, 0).Columns(1).Value2 = _
            ConvertDateToSerial(Date)
        Let SourceRange.Resize(SourceRange.Rows.Count - 1, 1).Offset(1, 0).Columns(2).Value2 = _
            ConvertTimeToSerial(Time)
    
        ' Exclude headers row
        Call Me.InitializeWithRange(SourceRange)
    End If
End Sub

' Custom Initializer.
Public Sub InitializeWithMySql(ServerName As String, DatabaseName As String, TableName As String, UserName As String, _
                                Password As String, AssetClass As String, Strategy As String, Optional ADate As Variant)
    Dim TheQuery As String
    Dim TheResults As Variant
    
    ' Set the query to determine the latest portfolio available on the given date for this client and asset class
    If IsMissing(ADate) Then
        Let TheQuery = "SELECT * FROM `" & DatabaseName & "`.`" & TableName & "` WHERE "
        Let TheQuery = TheQuery & "`strategycode` = """ & Strategy & """"
        Let TheQuery = TheQuery & " AND `assetclass` = """ & AssetClass & """;"
    ElseIf IsNumeric(ADate) Then
        Let TheQuery = "SELECT * FROM `" & DatabaseName & "`.`" & TableName & "` WHERE "
        Let TheQuery = TheQuery & "`date` = " & ADate & " AND `strategycode` = """ & Strategy & """"
        Let TheQuery = TheQuery & " AND `assetclass` = """ & AssetClass & """;"
    Else
        Exit Sub
    End If
    
    ' Select the data and get the latest time available for this date
    Let TheResults = ConnectAndSelect(TheQuery, DatabaseName, ServerName, UserName, Password)
    
    If GetNumberOfRows(TheResults) > 1 Then
        ' Create a new workbook to hold data and dump data in a worksheet
        Set UnderlyingWorkbook = Application.Workbooks.Add
        Call RemoveAllOtherWorksheets(UnderlyingWorkbook.Worksheets(1))
        Call DumpInTempPositionWithoutFirstClearing(TheResults, UnderlyingWorkbook.Worksheets(1).Range("A1"))
        
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

        Let TheArray = Me.GetRange.Value2
        Let TheHeaders = Most(ConvertTo1DArray(AddSingleBackQuotesToAllArrayElements(GetRow(ConnectAndSelect("SELECT * FROM `" & DatabaseName & "`.`" & TableName & "`;", DatabaseName, ServerName, UserName, Password), 1))))
        
        ' Add quote to every element of the following columns
        Let TheArray = AddQuotesToAllSubArrayElements(TheArray, 3, 1, Me.GetDbLength, 4)
        Let TheArray = AddQuotesToAllSubArrayElements(TheArray, 8, 1, Me.GetDbLength, 1)
        
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
    Dim obj As New TargetAssetAllocationRow

    Call obj.InitializeWithRange(ThisWorkbook.Worksheets("TempComputation").Range("A1"))
    
    Let GetNumberOfColumns = obj.GetNumberOfColumns
End Property

Property Get GetDates() As Range
    Dim obj As New TargetAssetAllocationRow

    Call obj.InitializeWithRange(ThisWorkbook.Worksheets("TempComputation").Range("A1"))
    
    Set GetDates = Me.GetRange.Columns(obj.GetDate.Column).Cells
End Property

Property Get GetTimes() As Range
    Dim obj As New TargetAssetAllocationRow

    Call obj.InitializeWithRange(ThisWorkbook.Worksheets("TempComputation").Range("A1"))

    Set GetTimes = Me.GetRange.Columns(obj.GetTime.Column).Cells
End Property

Property Get GetFundIds() As Range
    Dim obj As New TargetAssetAllocationRow

    Call obj.InitializeWithRange(ThisWorkbook.Worksheets("TempComputation").Range("A1"))
    
    Set GetFundIds = Me.GetRange.Columns(obj.GetFundId.Column).Cells
End Property

Property Get GetAssetClasses() As Range
    Dim obj As New TargetAssetAllocationRow

    Call obj.InitializeWithRange(ThisWorkbook.Worksheets("TempComputation").Range("A1"))
    
    Set GetAssetClasses = Me.GetRange.Columns(obj.GetAssetClass.Column).Cells
End Property

Property Get GetStrategyCodes() As Range
    Dim obj As New TargetAssetAllocationRow

    Call obj.InitializeWithRange(ThisWorkbook.Worksheets("TempComputation").Range("A1"))
    
    Set GetStrategyCodes = Me.GetRange.Columns(obj.GetStrategyCode.Column).Cells
End Property

Property Get GetPiamClientIds() As Range
    Dim obj As New TargetAssetAllocationRow

    Call obj.InitializeWithRange(ThisWorkbook.Worksheets("TempComputation").Range("A1"))
    
    Set GetPiamClientIds = Me.GetRange.Columns(obj.GetPiamClientId.Column).Cells
End Property

Property Get GetTargetUsdValues() As Range
    Dim obj As New TargetAssetAllocationRow

    Call obj.InitializeWithRange(ThisWorkbook.Worksheets("TempComputation").Range("A1"))
    
    Set GetTargetUsdValues = Me.GetRange.Columns(obj.GetTargetUsdValue.Column).Cells
End Property

Property Get GetCustodians() As Range
    Dim obj As New TargetAssetAllocationRow

    Call obj.InitializeWithRange(ThisWorkbook.Worksheets("TempComputation").Range("A1"))
    
    Set GetCustodians = Me.GetRange.Columns(obj.GetCustodian.Column).Cells
End Property

' Returns the ith asset in the portfolio
Property Get GetIthTargetAssetAllocationRow(i As Long) As TargetAssetAllocationRow
    Dim obj As New TargetAssetAllocationRow
    
    ' Initialize the object with the range where the object is held
    Call obj.InitializeWithRange(Me.GetRange.Rows(i))
    
    ' return a reference to the object
    Set GetIthTargetAssetAllocationRow = obj
End Property

' This function returns a 1D array with a Portfolio's headers
Public Function GetHeaders() As Variant
    Dim obj As New TargetAssetAllocationRow
    
    ' Initialize the object with the range where the object is held
    Call obj.InitializeWithRange(Me.GetRange.Rows(i))
    
    Let GetHeaders = obj.GetHeaders
End Function

Public Sub FormatUnderlyingRange()
    Dim ColumnPositions As Variant
    Dim ColPos As Variant
    Dim obj As New TargetAssetAllocationRow
    
    ' Initialize the object with the range where the object is held
    Call obj.InitializeWithRange(ThisWorkbook.Worksheets("TempComputation").Range("A1"))
    
    ' Set array of columns to center
    Let ColumnPositions = CreateSequentialArray(1, obj.GetNumberOfColumns)
    For Each ColPos In ColumnPositions
        Let Me.GetRange.Columns(ColPos).EntireColumn.HorizontalAlignment = xlCenter
    Next
    
    ' Set columns to format as "0"
    Let ColumnPositions = Array(obj.GetDate.Column, obj.GetTime.Column)
    For Each ColPos In ColumnPositions
        Let Me.GetRange.Columns(ColPos).NumberFormat = "0"
    Next
    
    ' Set target USD column as USD
    Let Me.GetRange.Columns(obj.GetTargetUsdValue.Column).Cells.NumberFormat = "$#,##0.00_);($#,##0.00)"
    
     ' Set columns to format as text
    Let ColumnPositions = Array(obj.GetFundId.Column, obj.GetAssetClass.Column, obj.GetStrategyCode.Column, obj.GetPiamClientId.Column, obj.GetCustodian.Column)
    For Each ColPos In ColumnPositions
        Let Me.GetRange.Columns(ColPos).Cells.NumberFormat = "@"
    Next
End Sub
