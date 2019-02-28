Attribute VB_Name = "MySql"
Option Explicit
Option Base 1

Public Const DbServerAddress As String = "localhost"
Public Const DbUserName As String = "root"
Public Const DbPassword As String = ""
Public Const DbDriverString As String = "{MySQL ODBC 5.2 ANSI Driver}"
Public Const SslCa As String = ""
Public Const SslCert As String = ""
Public Const SslKey As String = ""
Public Const SslVerify As Integer = 0

' This function returns a 1D string array with the names of the ODBC drivers installed in the system
Public Function GetOdbcDeviceDrivers() As String()
    Const HKEY_LOCAL_MACHINE = &H80000002
    Dim strComputer As String
    Dim strKeyPath As String
    Dim arrValueNames As Variant
    Dim strValueName As String
    Dim arrValueTypesa As Variant
    Dim strValue As Variant
    Dim objRegistry As Object
    Dim i As Integer
    Dim ResultArray() As String
     
    Let strComputer = "."
     
    Set objRegistry = GetObject("winmgmts:\\" & strComputer & "\root\default:StdRegProv")
     
    Let strKeyPath = "SOFTWARE\ODBC\ODBCINST.INI\ODBC Drivers"
    Call objRegistry.EnumValues(HKEY_LOCAL_MACHINE, strKeyPath, arrValueNames, arrValueTypesa)
     
    ReDim ResultArray(1 To Length(arrValueNames))
    For i = 0 To UBound(arrValueNames)
        Let strValueName = arrValueNames(i)
        Call objRegistry.GetStringValue(HKEY_LOCAL_MACHINE, strKeyPath, strValueName, strValue)
        Let ResultArray(i + 1) = arrValueNames(i)
    Next

    Let GetOdbcDeviceDrivers = ResultArray
End Function

' This sub-routine injects CSV file with name CsvFileName in directory TheDir into table with TargetTable in
' database DbName. The function logs into server ServerNameOrIpAddress using username UserName with password
' ThePassword.  ColumnsNames is a string like (`TheDate`, `Identifier`, `Value`) indicating the fields to
' populate.
Public Sub InjectCsvFileIntoMySql(TheDir As String, _
                                  CsvFileName As String, _
                                  DbName As String, _
                                  TargetTable As String, _
                                  ServerNameOrIpAddress As String, _
                                  UserName As String, _
                                  ThePassword As String, _
                                  ColumnNames As String, _
                                  Optional UseSllQ As Boolean = False)

    Call InjectFileIntoMysql(TheDir, CsvFileName, DbName, TargetTable, ServerNameOrIpAddress, UserName, _
                             ThePassword, ColumnNames, ",", UseSllQ)
End Sub

' This sub-routine injects a Delimiter-separated files with name AFileName in directory TheDir into table with TargetTable in
' database DbName. The function logs into server ServerNameOrIpAddress using username UserName with password
' ThePassword.  ColumnsNames is a string like (`TheDate`, `Identifier`, `Value`) indicating the fields to
' populate.
Public Sub InjectFileIntoMysql(TheDir As String, _
                               AFileName As String, _
                               DbName As String, _
                               TargetTable As String, _
                               ServerNameOrIpAddress As String, _
                               UserName As String, _
                               ThePassword As String, _
                               ColumnNames As String, _
                               ColumnSeparator As String, _
                               Optional FieldEncloser As String = "\""", _
                               Optional UseSllQ As Boolean = False)
    Dim SQLStr As String
    Dim cn As ADODB.Connection
    Dim QuerryString As String
    
    ' Set the database connection and recordset objects
    Set cn = New ADODB.Connection
       
    ' Open the database connection
    cn.Open CreateConnectionString(ServerNameOrIpAddress, DbName, UserName, ThePassword, UseSllQ)

    Let QuerryString = "FIELDS TERMINATED BY '" & ColumnSeparator & "' " & _
                       "ENCLOSED BY '" & FieldEncloser & "' " & _
                       "LINES TERMINATED BY '\r\n' " & _
                       ColumnNames & ";"
        
    ' Set up the SQL querry.
   Let SQLStr = "LOAD DATA LOCAL INFILE '" & ReplaceForwardSlashWithBackSlash(TheDir & AFileName) & _
                 "' INTO TABLE " & TargetTable & " " & QuerryString
        
    ' Execute the SQL querry.
    cn.Execute SQLStr

    ' Close the recordset and connection.
    cn.Close
End Sub

' Returns a 1D array with the given table's headers
Public Function GetTableHeaders(TableName As String, _
                                DbName As String, _
                                ServerNameOrIpAddress As String, _
                                UserName As String, _
                                ThePassword As String, _
                                Optional UseSllQ As Boolean = False) As Variant
    Let GetTableHeaders = Flatten(ConnectAndSelect("SELECT * FROM `" & DbName & "`.`" & TableName & "` LIMIT 0,0;", _
                                                   DbName, _
                                                   ServerNameOrIpAddress, _
                                                   UserName, _
                                                   ThePassword, _
                                                   UseSllQ))
End Function

' Executes and returns (as 2D array with headers) the result of a SELECT query
Public Function ConnectAndSelect(TheSqlQuery As String, _
                                 DbName As String, _
                                 ServerNameOrIpAddress As String, _
                                 UserName As String, _
                                 ThePassword As String, _
                                 Optional UseSllQ As Boolean = False) As Variant
    Dim cnt As ADODB.Connection
    Dim rst As ADODB.Recordset
    Dim i As Integer
    Dim Headers() As String

    'Instantiate the ADO-objects.
    Set cnt = New ADODB.Connection
    Set rst = New ADODB.Recordset

    ' Open the database connection
    cnt.Open CreateConnectionString(ServerNameOrIpAddress, DbName, UserName, ThePassword, UseSllQ)

    Call rst.Open(TheSqlQuery, cnt) 'Create the recordset.
    
    Let ConnectAndSelect = ConvertRecordSetToMatrix(rst)
    
    'Release objects from the memory.
    Call cnt.Close
    Set rst = Nothing
    Set cnt = Nothing
End Function

' Given an opened connection, execute a SELECT query and return the result as a 2D array
Public Function ExecuteQuery(cnt As ADODB.Connection, TheSqlQuery As String) As Variant
    Dim rst As ADODB.Recordset

    'Instantiate the ADO-objects.
    Set rst = New ADODB.Recordset

    Call rst.Open(TheSqlQuery, cnt) 'Create the recordset.
    
    Let ExecuteQuery = ConvertRecordSetToMatrix(rst)
End Function

' This sub creates a connection, properly quotes headers and data set, and then injects it in the DB
' The sub splits very large matrices into manageable chunks.
Public Function InjectMatrixIntoMySql(ValuesMatrix As Variant, _
                                      FieldNames As Variant, _
                                      TableName As String, _
                                      ServerAddress As String, _
                                      DbName As String, _
                                      UserName As String, _
                                      ThePassword As String, _
                                      Optional StatusBarMsgsFlag As Boolean = False, _
                                      Optional UseSllQ As Boolean = False) As Boolean
    Dim i As Long
    Dim NumberOfChunks As Integer
    Dim QuotedValuesMatrix As Variant
    Dim QuotedFieldNames As Variant
    Const ChunkSize As Integer = 1000
    
    Let InjectMatrixIntoMySql = True
    
    If EmptyArrayQ(ValuesMatrix) Or EmptyArrayQ(FieldNames) Then
        Let InjectMatrixIntoMySql = False
    
        Exit Function
    End If
    
    ' Sinble-back-quote FieldNames
    Let QuotedFieldNames = AddSingleBackQuotesToAllArrayElements(FieldNames)
    
    ' Compute the number of chunks (e.g. number of insert statements we need)
    Let NumberOfChunks = Application.Floor_Precise(NumberOfRows(ValuesMatrix) / ChunkSize)
    
    ' Insert each chunk
    For i = 1 To NumberOfChunks
        ' Single-quote ValuesMatrix
        Let QuotedValuesMatrix = Part(ValuesMatrix, Span(1 + (i - 1) * ChunkSize, i * ChunkSize))
        Let QuotedValuesMatrix = AddSingleQuotesToAllArrayElements(QuotedValuesMatrix)

        Call ConnectAndExecuteInsertQuery(QuotedValuesMatrix, _
                                          QuotedFieldNames, _
                                          TableName, _
                                          ServerAddress, _
                                          DbName, _
                                          UserName, _
                                          ThePassword, _
                                          UseSllQ)
                                          
        If StatusBarMsgsFlag Then
            Let Application.StatusBar = "Injected chunk " & i & " out of " & IIf(NumberOfRows(ValuesMatrix) Mod ChunkSize = 0, 0, 1)
        End If
    Next i
    
    ' Insert the remainder after chunking the data set
    If NumberOfRows(ValuesMatrix) Mod ChunkSize > 0 Then
        Let QuotedValuesMatrix = Part(ValuesMatrix, Span(1 + CLng(NumberOfChunks) * ChunkSize, Length(ValuesMatrix)))
        Let QuotedValuesMatrix = AddSingleQuotesToAllArrayElements(QuotedValuesMatrix)
        
        Call ConnectAndExecuteInsertQuery(QuotedValuesMatrix, _
                                          QuotedFieldNames, _
                                          TableName, _
                                          ServerAddress, _
                                          DbName, _
                                          UserName, _
                                          ThePassword, _
                                          UseSllQ)
    End If
    
    Let Application.StatusBar = "Ready"
End Function

' This sub creates a connection, properly quotes headers and data set, and then injects it in the DB
' The sub splits very large matrices into manageable chunks.  The data is contained in a list object
Public Function InjectListObjectIntoMySql(aListObject As ListObject, _
                                          FieldNames As Variant, _
                                          TableName As String, _
                                          ServerAddress As String, _
                                          DbName As String, _
                                          UserName As String, _
                                          ThePassword As String, _
                                          Optional StatusBarMsgsFlag As Boolean = False, _
                                          Optional UseSllQ As Boolean = False) As Boolean
    Dim i As Long
    Dim NumberOfChunks As Integer
    Dim QuotedValuesMatrix As Variant
    Dim QuotedFieldNames As Variant
    Const ChunkSize As Integer = 500
    
    Let InjectListObjectIntoMySql = True
    
    If aListObject.ListRows.Count = 0 Or EmptyArrayQ(FieldNames) Then
        Let InjectListObjectIntoMySql = False
    
        Exit Function
    End If
    
    ' Sinble-back-quote FieldNames
    Let QuotedFieldNames = AddSingleBackQuotesToAllArrayElements(FieldNames)
    
    ' Compute the number of chunks (e.g. number of insert statements we need)
    Let NumberOfChunks = Application.Floor_Precise(aListObject.ListRows.Count / ChunkSize)
    
    ' Insert each chunk
    For i = 1 To NumberOfChunks
        ' Single-quote ValuesMatrix
        Let QuotedValuesMatrix = aListObject.ListRows(1 + (i - 1) * ChunkSize).Range.Resize(ChunkSize, aListObject.ListColumns.Count).Value2
        Let QuotedValuesMatrix = AddSingleQuotesToAllArrayElements(QuotedValuesMatrix)

        Call ConnectAndExecuteInsertQuery(QuotedValuesMatrix, _
                                          QuotedFieldNames, _
                                          TableName, _
                                          ServerAddress, _
                                          DbName, _
                                          UserName, _
                                          ThePassword, _
                                          UseSllQ)
                                          
        If StatusBarMsgsFlag Then
            Let Application.StatusBar = "Injected chunk " & i & " out of " & IIf(aListObject.ListRows.Count Mod ChunkSize = 0, 0, 1)
        End If
    Next i
    
    ' Insert the remainder after chunking the data set
    If aListObject.ListRows.Count Mod ChunkSize > 0 Then
        Let QuotedValuesMatrix = aListObject.ListRows(1 + CLng(NumberOfChunks) * ChunkSize).Range.Resize(aListObject.ListRows.Count - (1 + CLng(NumberOfChunks) * ChunkSize) + 1, aListObject.ListColumns.Count).Value2

        Let QuotedValuesMatrix = AddSingleQuotesToAllArrayElements(QuotedValuesMatrix)
        
        Call ConnectAndExecuteInsertQuery(QuotedValuesMatrix, _
                                          QuotedFieldNames, _
                                          TableName, _
                                          ServerAddress, _
                                          DbName, _
                                          UserName, _
                                          ThePassword, _
                                          UseSllQ)
    End If
    
    Let InjectListObjectIntoMySql = True
    
    Let Application.StatusBar = "Ready"
End Function

' Execute a query that returns no data.
Public Sub ConnectAndExecuteInsertQuery(ValuesMatrix As Variant, FieldNames As Variant, TableName As String, _
                                        ServerAddress As String, DbName As String, UserName As String, _
                                        ThePassword As String, Optional UseSllQ As Boolean = False)
    Dim cn As ADODB.Connection
    Dim TheQuery As String
    Dim i As Integer
    Dim j As Integer
    Dim TheRowArray() As Variant
    
    ' Set the database connection and recordset objects
    Set cn = New ADODB.Connection
       
    ' Open the database connection
    cn.Open CreateConnectionString(ServerNameOrIpAddress, DbName, UserName, ThePassword, UseSllQ)
            
    ' Create query
    Let TheQuery = "INSERT INTO `" & DbName & "`.`" & TableName & "` " & vbCrLf
    Let TheQuery = TheQuery & Convert1DArrayIntoParentheticalExpression(FieldNames) & vbCrLf
    Let TheQuery = TheQuery & " VALUES " & vbCrLf
    
    ReDim TheRowArray(1 To NumberOfColumns(ValuesMatrix))
    For i = 1 To NumberOfRows(ValuesMatrix)
        For j = 1 To NumberOfColumns(ValuesMatrix)
            Let TheRowArray(j) = ValuesMatrix(i, j)
        Next j
    
        Let TheQuery = TheQuery & Convert1DArrayIntoParentheticalExpression(TheRowArray)
        
        If i < NumberOfRows(ValuesMatrix) Then
            Let TheQuery = TheQuery & ", " & vbCrLf
        End If
    Next i
    Let TheQuery = TheQuery & ";"
    
    ' Execute the query
    Call cn.Execute(TheQuery)
    
    ' Close the recordset and connection.
    Call cn.Close
    
    ' Destroy the recordset and connection objects.
    Set cn = Nothing
End Sub

' Execute a query that returns no data.
Public Sub RunQuery(TheQuery As String, ServerName As String, DbName As String, UserName As String, ThePassword As String, _
                    Optional UseSllQ As Boolean = False)
    Dim cn As ADODB.Connection
    ' Set the database connection and recordset objects
    Set cn = New ADODB.Connection
       
    ' Open the database connection
    cn.Open CreateConnectionString(ServerNameOrIpAddress, DbName, UserName, ThePassword, UseSllQ)
        
    Call cn.Execute(TheQuery)
    
    ' Close the recordset and connection.
    Call cn.Close
    
    ' Destroy the recordset and connection objects.
    Set cn = Nothing
End Sub

' This function returns TRUE if the DB connection works and FALSE otherwise
Public Function DbConnectionOkayQ(txtHostname As String, txtDatabase As String, txtUsername As String, _
                                  txtPassword As String, Optional UseSllQ As Boolean = False) As Boolean
    Dim oConn As ADODB.Connection

    On Error GoTo ErrHandler

    Set oConn = New ADODB.Connection
    oConn.Open CreateConnectionString(ServerNameOrIpAddress, DbName, UserName, ThePassword, UseSllQ)
    
    Let DbConnectionOkayQ = True
        
    Call oConn.Close
    Exit Function

ErrHandler:
    Let DbConnectionOkayQ = False
End Function

' This sub populates a DB table with the contents of the same table on a different server
' We assume the source and target servers both have the DB and table with tha same name and the table with the
' same exact structure.  The target table is truncated before copying over the data.
Public Sub CopyTableFromOneDbServerToAnother(SourceServerAddress As String, TheDbName As String, DbTableName As String, _
                                             SourceDbUsername As String, SourceDbPassword As String, _
                                             TargetServerAddress As String, TargetDbUsername As String, _
                                             TargetDbPassword As String, Optional UseSllQ As Boolean = False)
    Dim TheData As Variant
    
    Let TheData = ConnectAndSelect("SELECT * FROM `" & TheDbName & "`.`" & DbTableName & "`;", _
                                   TheDbName, SourceServerAddress, SourceDbUsername, SourceDbPassword, UseSllQ)
    Let TheData = Rest(TheData)
    
    Let TheData = Drop(TransposeMatrix(TheData), Array(NumberOfColumns(TheData)))
    Let TheData = TransposeMatrix(TheData)
    
    Call MySql.RunQuery("TRUNCATE `" & TheDbName & "`.`" & DbTableName & "`;", TargetServerAddress, TheDbName, _
                        TargetDbUsername, TargetDbPassword, UseSllQ)
    
    Call InjectMatrixIntoMySql(ToTemp(TheData, True).Value2, Most(GetTableHeaders(DbTableName, TheDbName, SourceServerAddress, SourceDbUsername, SourceDbPassword)), DbTableName, _
                               TargetServerAddress, TheDbName, TargetDbUsername, TargetDbPassword, UseSllQ)
    
End Sub

' This function does the same thing as Worksheet.CopyFromRecordSet, but it does not fail when a column has an entry more than 255 characters long.
Public Function ConvertRecordSetToMatrix(rst As ADODB.Recordset, _
                                         Optional HeadersAndBodyQ As Boolean = True) As Variant
    Dim TheResults() As Variant
    Dim CurrentRow As Long
    Dim RowCount As Long
    Dim NColumns As Long
    Dim FirstRow As Long
    Dim r As Long
    Dim c As Long
    Dim h As Long
    
    Let NColumns = rst.Fields.Count
    Select Case HeadersAndBodyQ
        Case True
            Let FirstRow = 2
        Case Else
            Let FirstRow = 1
    End Select
    
    ReDim TheResults(1 To NColumns, 1 To 1)
    Let RowCount = 1
    
    If HeadersAndBodyQ Then
        For h = 0 To NColumns - 1
            Let TheResults(h + 1, 1) = rst.Fields(h).Name
        Next h
    End If
    
    If Not HeadersAndBodyQ Then
        Let CurrentRow = 1
    Else
        Let CurrentRow = 2
    End If
    
    While Not rst.EOF
        Let RowCount = RowCount + 1
        ReDim Preserve TheResults(1 To NColumns, 1 To RowCount)
    
        For c = 0 To NColumns - 1
            Let TheResults(c + 1, CurrentRow) = rst.Fields(c).Value
        Next c
        
        Call rst.MoveNext
        Let CurrentRow = CurrentRow + 1
    Wend
    
    Let ConvertRecordSetToMatrix = TransposeMatrix(TheResults, False)
End Function

Public Function CreateConnectionString(ServerAddress As String, _
                                       DatabaseName As String, _
                                       TheUserName As String, _
                                       ThePassword As String, _
                                       Optional UseSllQ As Boolean = False) As String
    Let CreateConnectionString = "Driver=" & DbDriverString & ";" & _
                                 "Server=" & ServerAddress & ";" & _
                                 "Database=" & DatabaseName & ";" & _
                                 "Uid=" & TheUserName & ";" & _
                                 "Pwd=" & ThePassword & ";" & _
                                 IIf(SslCa = vbNullString, vbNullString, "sslca=" & SslCa & ";") & _
                                 IIf(SslCert = vbNullString, vbNullString, "slcert=" & SslCert & ";") & _
                                 IIf(SslKey = vbNullString, vbNullString, "SslKey = " & SslKey & ";") & _
                                 IIf(UseSllQ, "sslverify=1;", vbNullString)
End Function
