VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EmsxTradeList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Base 1
Option Explicit

' Private instance variable to hold the trade database
Public RangeDb As Range
Public UnderlyingWorkbook As Workbook

' Custom initializer. This creates a new workbook and inserts
' the headers row.  It sets RangeDb to Nothing because initially there
' are no securities in the trade list
Public Sub InitializeWithEmptyWorkbook()
    Dim i As Integer
    Dim AWorkBook As Workbook
    Dim WorksheetNames() As String
    Dim TargetWorksheetName As String
    
    ' Create a new workbook to hold this template
    Set AWorkBook = Application.Workbooks.Add
    
    ' Delete any other worksheets beside the first one if there is more than one
    Call RemoveAllOtherWorksheets(AWorkBook.Worksheets(1))
    
    ' Insert headers row
    Let AWorkBook.Worksheets(1).Range("A1").Resize(1, UBound(GetHeaders)).Value = GetHeaders
    
    ' Set a reference to the underlying workbook
    Set UnderlyingWorkbook = AWorkBook
End Sub

' Custom initializer. aRange must include a header row
Public Sub InitializeWithRange(ARange As Range)
    Set UnderlyingWorkbook = ARange.Worksheet.Parent
    Set RangeDb = ARange.Resize(ARange.Rows.Count - 1, UBound(GetHeaders)).Offset(1, 0)
End Sub

Public Sub InitializeWithMySql(ServerName As String, DatabaseName As String, TableName As String, FundId As String, _
                               Custodian As String, UserName As String, Password As String, AssetClass As String, _
                               Strategy As String, ADate As Long)
    Dim TheQuery As String
    Dim TheResults As Variant
    Dim LatestTime As Long
    
    ' Set the query to determine the latest portfolio available on the given date for this client and asset class
    Let TheQuery = "SELECT MAX(`time`) FROM `" & DatabaseName & "`.`" & TableName & "` WHERE "
    Let TheQuery = TheQuery & "`date` = " & ADate & " AND `strategycode` = '" & Strategy & "' AND `assetclasscode` = '" & AssetClass & "'"
    If FundId = "ALL" And Custodian <> "ALL" Then
        Let TheQuery = TheQuery & " AND `custodian` = '" & Custodian & "'"
    ElseIf FundId <> "ALL" And Custodian = "ALL" Then
        Let TheQuery = TheQuery & " AND `fundid` = '" & FundId & "'"
    Else
        Let TheQuery = TheQuery & " AND `fundid` = '" & FundId & "' AND `custodian` = '" & Custodian & "'"
    End If
    Let TheQuery = TheQuery & ";"
    
    ' Select the data and get the latest time available for this date
    Let TheResults = ConnectAndSelect(TheQuery, DatabaseName, ServerName, UserName, Password)
    
    If Not IsArray(TheResults) Then
        Exit Sub
    End If
    
    If IsNull(TheResults(2, 1)) Then
        Exit Sub
    End If
    
    If GetNumberOfRows(TheResults) > 1 Then
        Let LatestTime = TheResults(2, 1)
    
        ' Set the query to get the latest portfolio on this date
        Let TheQuery = "SELECT * FROM `" & DatabaseName & "`.`" & TableName & "` WHERE "
        Let TheQuery = TheQuery & "`date` = " & ADate & " AND `strategycode` = '" & Strategy & "' AND `assetclasscode` = '" & AssetClass & "' AND `time` = " & LatestTime
        If FundId = "ALL" And Custodian <> "ALL" Then
            Let TheQuery = TheQuery & " AND `custodian` = '" & Custodian & "'"
        ElseIf FundId <> "ALL" And Custodian = "ALL" Then
            Let TheQuery = TheQuery & " AND `fundid` = '" & FundId & "'"
        Else
            Let TheQuery = TheQuery & " AND `fundid` = '" & FundId & "' AND `custodian` = '" & Custodian & "'"
        End If
        Let TheQuery = TheQuery & ";"
    
        ' Select the data
        Let TheResults = ConnectAndSelect(TheQuery, DatabaseName, ServerName, UserName, Password)
        
        ' Create a new workbook to hold data and dump data in a worksheet
        Set UnderlyingWorkbook = Application.Workbooks.Add
        Call RemoveAllOtherWorksheets(UnderlyingWorkbook.Worksheets(1))
        Call DumpInTempPositionWithoutFirstClearing(TheResults, UnderlyingWorkbook.Worksheets(1).Range("A1"))
        
        ' Initialize with the data now held in the range
        Call Me.InitializeWithRange(UnderlyingWorkbook.Worksheets(1).Range("A1").CurrentRegion)
    End If
End Sub


' Getters
' This returns the range holding the asset's value. This could be useful when trying to get the address
' of the range.
Property Get GetRange() As Range
    Set GetRange = RangeDb
End Property

Property Get GetNumberOfColumns() As Integer
    Dim obj As New EmsxTrade
    
    Let GetNumberOfColumns = obj.GetNumberOfColumns
End Property

' Returns the number of securities in the BloombergEqs instance
Property Get GetDbLength() As Long
    If GetRange Is Nothing Then
        Let GetDbLength = 0
    Else
        Let GetDbLength = GetRange.Rows.Count
    End If
End Property

Property Get GetDates() As Range
    Dim obj As New EmsxTrade

    Call obj.InitializeWithRange(ThisWorkbook.Worksheets("TempComputation").Range("A1"))
    
    Set GetDates = GetRange.Columns(obj.GetDate.Column)
End Property

Property Get GetTimes() As Range
    Dim obj As New EmsxTrade
    
    Call obj.InitializeWithRange(ThisWorkbook.Worksheets("TempComputation").Range("A1"))
    
    Set GetTimes = GetRange.Columns(obj.GetTime.Column)
End Property

Property Get GetAssetClassCodes() As Range
    Dim obj As New EmsxTrade
    
    Call obj.InitializeWithRange(ThisWorkbook.Worksheets("TempComputation").Range("A1"))
    
    Set GetAssetClassCodes = GetRange.Columns(obj.GetAssetClassCode.Column)
End Property

Property Get GetStrategyCodes() As Range
    Dim obj As New EmsxTrade
    
    Call obj.InitializeWithRange(ThisWorkbook.Worksheets("TempComputation").Range("A1"))

    Set GetStrategyCodes = GetRange.Columns(obj.GetStrategyCode.Column)
End Property

Property Get GetCustodians() As Range
    Dim obj As New EmsxTrade
    
    Call obj.InitializeWithRange(ThisWorkbook.Worksheets("TempComputation").Range("A1"))

    Set GetCustodians = GetRange.Columns(obj.GetCustodian.Column)
End Property

Property Get GetFundIds() As Range
    Dim obj As New EmsxTrade
    
    Call obj.InitializeWithRange(ThisWorkbook.Worksheets("TempComputation").Range("A1"))

    Set GetFundIds = GetRange.Columns(obj.GetFundId.Column)
End Property

Property Get GetBloombergTickers() As Range
    Dim obj As New EmsxTrade
    
    Call obj.InitializeWithRange(ThisWorkbook.Worksheets("TempComputation").Range("A1"))

    Set GetBloombergTickers = GetRange.Columns(obj.GetBloombergTicker.Column)
End Property

Property Get GetExchangeCodes() As Range
    Dim obj As New EmsxTrade
    
    Call obj.InitializeWithRange(ThisWorkbook.Worksheets("TempComputation").Range("A1"))

    Set GetExchangeCodes = GetRange.Columns(obj.GetBloombergExchangeCode.Column)
End Property

Property Get GetCurrencyCodes() As Range
    Dim obj As New EmsxTrade
    
    Call obj.InitializeWithRange(ThisWorkbook.Worksheets("TempComputation").Range("A1"))

    Set GetCurrencyCodes = GetRange.Columns(obj.GetCurrencyCode.Column)
End Property

Property Get GetOrders() As Range
    Dim obj As New EmsxTrade
    
    Call obj.InitializeWithRange(ThisWorkbook.Worksheets("TempComputation").Range("A1"))

    Set GetOrders = GetRange.Columns(obj.GetOrder.Column)
End Property

Property Get GetSides() As Range
    Dim obj As New EmsxTrade
    
    Call obj.InitializeWithRange(ThisWorkbook.Worksheets("TempComputation").Range("A1"))

    Set GetSides = GetRange.Columns(obj.GetSide.Column)
End Property

Property Get GetUnits() As Range
    Dim obj As New EmsxTrade
    
    Call obj.InitializeWithRange(ThisWorkbook.Worksheets("TempComputation").Range("A1"))

    Set GetUnits = GetRange.Columns(obj.GetUnits.Column)
End Property

Property Get GetLimits() As Range
    Dim obj As New EmsxTrade
    
    Call obj.InitializeWithRange(ThisWorkbook.Worksheets("TempComputation").Range("A1"))

    Set GetLimits = GetRange.Columns(obj.GetLimit.Column)
End Property

Property Get GetTifs() As Range
    Dim obj As New EmsxTrade
    
    Call obj.InitializeWithRange(ThisWorkbook.Worksheets("TempComputation").Range("A1"))

    Set GetTifs = GetRange.Columns(obj.GetTif.Column)
End Property

Property Get GetAccounts() As Range
    Dim obj As New EmsxTrade
    
    Call obj.InitializeWithRange(ThisWorkbook.Worksheets("TempComputation").Range("A1"))
    
    Set GetAccounts = GetRange.Columns(obj.GetAccount.Column)
End Property

Property Get GetOptimalDateStamps() As Range
    Dim obj As New EmsxTrade
    
    Call obj.InitializeWithRange(ThisWorkbook.Worksheets("TempComputation").Range("A1"))

    Set GetOptimalDateStamps = GetRange.Columns(obj.GetOptimalDateStamp.Column)
End Property

Property Get GetOptimalTimeStamps() As Range
    Dim obj As New EmsxTrade
    
    Call obj.InitializeWithRange(ThisWorkbook.Worksheets("TempComputation").Range("A1"))

    Set GetOptimalTimeStamps = GetRange.Columns(obj.GetOptimalTimeStamp.Column)
End Property

Property Get GetHoldingsDateStamps() As Range
    Dim obj As New EmsxTrade
    
    Call obj.InitializeWithRange(ThisWorkbook.Worksheets("TempComputation").Range("A1"))

    Set GetHoldingsDateStamps = GetRange.Columns(obj.GetHoldingsDateStamp.Column)
End Property

Property Get GetHoldingsTimeStamps() As Range
    Dim obj As New EmsxTrade
    
    Call obj.InitializeWithRange(ThisWorkbook.Worksheets("TempComputation").Range("A1"))

    Set GetHoldingsTimeStamps = GetRange.Columns(obj.GetHoldingsTimeStamp.Column)
End Property

Property Get GetAllocationsDateStamps() As Range
    Dim obj As New EmsxTrade
    
    Call obj.InitializeWithRange(ThisWorkbook.Worksheets("TempComputation").Range("A1"))

    Set GetAllocationsDateStamps = GetRange.Columns(obj.GetAllocationsDateStamp.Column)
End Property

Property Get GetAllocationsTimeStamps() As Range
    Dim obj As New EmsxTrade
    
    Call obj.InitializeWithRange(ThisWorkbook.Worksheets("TempComputation").Range("A1"))

    Set GetAllocationsTimeStamps = GetRange.Columns(obj.GetAllocationsTimeStamp.Column)
End Property

Property Get GetRunNumbers() As Range
    Dim obj As New EmsxTrade
    
    Call obj.InitializeWithRange(ThisWorkbook.Worksheets("TempComputation").Range("A1"))

    Set GetRunNumbers = GetRange.Columns(obj.GetRunNumber.Column)
End Property

' This extracts the ith trade in the list and returns it as an instance
' of class EmsxTrade
Property Get GetIthTrade(i As Integer) As EmsxTrade
    Dim ATrade As EmsxTrade
    
    Set ATrade = New EmsxTrade
    
    Call ATrade.InitializeWithRange(GetRange.Rows(i))
    
    Set GetIthTrade = ATrade
End Property

' This property gets the trade corresponding to the given Bloomberg ticker,
' number of units, and side.
Property Get GetTrade(BloombergTicker As String, IsoCrncyCode As String, Side As String, Units As Long) As EmsxTrade
    Dim i As Integer
    
    Let i = 1
    While i <= GetDbLength
        If GetIthTrade(i).GetBloombergTicker.Value2 = BloombergTicker And GetIthTrade(i).GetCurrencyCode.Value2 = IsoCrncyCode _
           And GetIthTrade(i).GetSide.Value2 = Side And GetIthTrade(i).GetUnits.Value2 = Units Then
           
           Set GetTrade = GetIthTrade(i)
           
           Exit Property
        End If
        
        ' Proceed to next trade
        Let i = i + 1
    Wend
    
    ' If the code gets to this point, the trade was not found.  Return nothing
    Set GetTrade = Nothing
End Property

' Subs
' This sub adds a new trade to the trade list
Public Sub AddTrade(BloombergTicker As Variant, _
                    BloombergExchangeCode As Variant, _
                    CurrencyCode As Variant, _
                    Order As Variant, _
                    Side As Variant, _
                    Units As Variant, _
                    Limit As Variant, _
                    Tif As Variant, _
                    Account As Variant)
    Dim ATrade As EmsxTrade
    
    ' Instantiate a EmsxTrade instance to hold the new trade
    Set ATrade = New EmsxTrade
    
    ' Resize the underlying range to hold this new trade and initialize aTrade with the new
    ' row in the underlying range of this trade list
    If Me.GetRange Is Nothing Then
        Set RangeDb = UnderlyingWorkbook.Worksheets(1).Range("A2").Resize(1, GetArrayLength(Me.GetHeaders()))
        Call ATrade.InitializeWithRange(GetRange.Rows(GetDbLength))
    Else
        Set RangeDb = RangeDb.Resize(RangeDb.Rows.Count + 1, RangeDb.Columns.Count)
        Call ATrade.InitializeWithRange(GetRange.Rows(GetDbLength))
    End If
    
    ' Populate the range row underlying the new trade with the given information
    Let GetIthTrade(GetDbLength).SetBloombergTicker = BloombergTicker
    Let GetIthTrade(GetDbLength).SetBloombergExchangeCode = BloombergExchangeCode
    Let GetIthTrade(GetDbLength).SetCurrencyCode = CurrencyCode
    Let GetIthTrade(GetDbLength).SetOrder = Order
    Let GetIthTrade(GetDbLength).SetSide = Side
    Let GetIthTrade(GetDbLength).SetUnits = Units
    Let GetIthTrade(GetDbLength).SetLimit = Limit
    Let GetIthTrade(GetDbLength).SetTif = Tif
    Let GetIthTrade(GetDbLength).SetAccount = Account
End Sub

' Deletes the ith trade from the database
Public Sub DeleteIthTrade(i As Integer)
    Call GetIthTrade(i).GetRange.Delete(Shift:=xlShiftUp)
End Sub

' Deletes a trade from the database, with uniqueness defined as equal tickers, currency code, trade side, and units traded
Public Sub DeleteTrade(ATrade As EmsxTrade)
    Set ATrade = GetTrade(ATrade.GetBloombergTicker.Value2, ATrade.GetCurrencyCode.Value2, ATrade.GetSide.Value2, ATrade.GetUnits.Value2)
    
    Call ATrade.GetRange.Delete(Shift:=xlShiftUp)
End Sub

' This saves the workbook holding the trades
Public Sub SaveChanges(AFileName As String)
    Call GetRange.Worksheet.Parent.SaveAs(Filename:=AFileName, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False)
End Sub

' This is the class destructor.
Public Sub Class_Terminate()
    If Not (Me.UnderlyingWorkbook Is Nothing) Then
        If Not Me.UnderlyingWorkbook Is ThisWorkbook Then
            Call Me.UnderlyingWorkbook.Close(SaveChanges:=False)
        End If
    End If
End Sub

Property Get GetHeaders() As Variant
    Dim ATrade As New EmsxTrade

    Let GetHeaders = ATrade.GetHeaders
End Property


