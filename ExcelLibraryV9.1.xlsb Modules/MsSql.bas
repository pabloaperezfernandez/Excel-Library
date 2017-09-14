Attribute VB_Name = "MsSql"
Option Explicit
Option Base 1

' This sub-routine injects CSV file with name CsvFileName in directory TheDir into table with TargetTable in
' database DbName. The function logs into server ServerNameOrIpAddress using username UserName with password
' ThePassword.  ColumnsNames is a string like (`TheDate`, `Identifier`, `Value`) indicating the fields to
' populate.
Public Sub InjectCsvFileIntoMsSql(TheDir As String, CsvFileName As String, _
                                  DbName As String, TargetTable As String, _
                                  ServerNameOrIpAddress As String, _
                                  ColumnNames As String)
    Dim SQLStr As String
    Dim cn As ADODB.Connection
    Dim QuerryString As String
    
    ' Set the database connection and recordset objects
    Set cn = New ADODB.Connection
    
    ' Open the database connection
    cnt.Open "Driver={SQL Server};Server=" & ServerNameOrIpAddress & ";Database=" & DbName & ";Trusted_Connection=Yes;"
        
    ' Set up the SQL querry.
    Let SQLStr = "bulk insert " & TargetTable & _
                 "from '" & CsvFileName & "' " & _
                 "with (fieldterminator = ',', rowterminator = '\n') " & _
                 "go"
        
    ' Execute the SQL querry.
    cn.Execute SQLStr

    ' Close the recordset and connection.
    cn.Close
End Sub

' Executes and returns (as 2D array with headers) the result of a SELECT query
' Example: printarray ConnectToMsSqlAndExecuteSelectQuery("SELECT * FROM TestTable;", "TestDb", "PabloTablet\TestInstance")
Public Function ConnectToMsSqlAndExecuteSelectQuery(TheSqlQuery As String, _
                                                    DbName As String, _
                                                    ServerNameOrIpAddress As String) As Variant
    Dim cnt As ADODB.Connection
    Dim rst As ADODB.Recordset
    Dim i As Integer
    Dim Headers() As String

    'Instantiate the ADO-objects.
    Set cnt = New ADODB.Connection
    Set rst = New ADODB.Recordset

    ' Open the database connection
    cnt.Open "Driver={SQL Server};Server=" & ServerNameOrIpAddress & ";Database=" & DbName & ";Trusted_Connection=Yes;"

    Call rst.Open(TheSqlQuery, cnt) 'Create the recordset.
    
    ' Dump results of query in TempComputation worksheet
    Call TempComputation.UsedRange.ClearContents
    
    ReDim Headers(rst.Fields.Count)
    For i = 1 To rst.Fields.Count
        Let Headers(i) = rst.Fields(i - 1).Name
    Next i
    Let TempComputation.Cells(1, 1).Resize(1, rst.Fields.Count).Value2 = Headers
    Call TempComputation.Cells(2, 1).CopyFromRecordset(rst)
          
    'Release objects from the memory.
    Call cnt.Close
    Set rst = Nothing
    Set cnt = Nothing
    
    ' Pull data to return
    Let ConnectToMsSqlAndExecuteSelectQuery = TempComputation.Range("A1").CurrentRegion.Value2
    
    ' Clear TempComputation worksheet
    Call TempComputation.UsedRange.ClearContents
End Function

' Execute a query that returns no data.
Public Sub ExecuteMsSqlQuery(TheQuery As String, ServerNameOrIpAddress As String, DbName As String)
    Dim cn As ADODB.Connection
    ' Set the database connection and recordset objects
    Set cn = New ADODB.Connection
       
    ' Open the database connection
    cn.Open "Driver={SQL Server};Server=" & ServerNameOrIpAddress & ";Database=" & DbName & ";Trusted_Connection=Yes;"
        
    Call cn.Execute(TheQuery)
    
    ' Close the recordset and connection.
    Call cn.Close
    
    ' Destroy the recordset and connection objects.
    Set cn = Nothing
End Sub


