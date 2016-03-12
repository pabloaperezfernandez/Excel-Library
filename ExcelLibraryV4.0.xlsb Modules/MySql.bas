Attribute VB_Name = "MySql"
Option Explicit
Option Base 1

Public Const DbServerAddress As String = "localhost"
Public Const DbUserName As String = "root"
Public Const DbPassword As String = ""
Public Const DbDriverString As String = "{MySQL ODBC 5.2 ANSI Driver}"

'Public Const DbServerAddress As String = "localhost"
'Public Const DbUserName As String = "root"
'Public Const DbPassword As String = ""

' This sub-routine injects CSV file with name CsvFileName in directory TheDir into table with TargetTable in
' database DbName. The function logs into server ServerNameOrIpAddress using username UserName with password
' ThePassword.  ColumnsNames is a string like (`TheDate`, `Identifier`, `Value`) indicating the fields to
' populate.
Public Sub InjectCsvFileIntoMySql(TheDir As String, CsvFileName As String, _
                                  DbName As String, TargetTable As String, _
                                  ServerNameOrIpAddress As String, _
                                  UserName As String, ThePassword As String, _
                                  ColumnNames As String)

    Call InjectFileIntoMysql(TheDir, CsvFileName, DbName, TargetTable, ServerNameOrIpAddress, UserName, ThePassword, ColumnNames, ",")
End Sub

' This sub-routine injects a Delimiter-separated files with name AFileName in directory TheDir into table with TargetTable in
' database DbName. The function logs into server ServerNameOrIpAddress using username UserName with password
' ThePassword.  ColumnsNames is a string like (`TheDate`, `Identifier`, `Value`) indicating the fields to
' populate.
Public Sub InjectFileIntoMysql(TheDir As String, AFileName As String, _
                                  DbName As String, TargetTable As String, _
                                  ServerNameOrIpAddress As String, _
                                  UserName As String, ThePassword As String, _
                                  ColumnNames As String, _
                                  Delimiter As String)
    Dim SQLStr As String
    Dim Cn As ADODB.Connection
    Dim QuerryString As String
    
    ' Set the database connection and recordset objects
    Set Cn = New ADODB.Connection
       
    ' Open the database connection
    Cn.Open "Driver=" & DbDriverString & ";" & _
            "Server=" & ServerNameOrIpAddress & ";" & _
            "Database=" & DbName & ";" & _
            "Uid=" & UserName & ";" & _
            "Pwd=" & ThePassword & ";"
        
    Let QuerryString = "FIELDS TERMINATED BY '" & Delimiter & "' " & _
                       "ENCLOSED BY '\""' " & _
                       "LINES TERMINATED BY '\r\n' " & _
                       ColumnNames & ";"
        
    ' Set up the SQL querry.
   Let SQLStr = "LOAD DATA LOCAL INFILE '" & ReplaceForwardSlashWithBackSlash(TheDir & AFileName) & _
                 "' INTO TABLE " & TargetTable & " " & QuerryString
        
    ' Execute the SQL querry.
    Cn.Execute SQLStr

    ' Close the recordset and connection.
    Cn.Close
End Sub

' Returns a 1D array with the given table's headers
Public Function GetTableHeaders(TableName As String, _
                                DbName As String, _
                                ServerNameOrIpAddress As String, _
                                UserName As String, ThePassword As String) As Variant
    Let GetTableHeaders = ConvertTo1DArray(ConnectAndSelect("SELECT * FROM `" & DbName & "`.`" & TableName & "` LIMIT 0,0;", DbName, ServerNameOrIpAddress, UserName, ThePassword))
End Function

' Executes and returns (as 2D array with headers) the result of a SELECT query
Public Function ConnectAndSelect(TheSqlQuery As String, _
                                 DbName As String, _
                                 ServerNameOrIpAddress As String, _
                                 UserName As String, ThePassword As String) As Variant
    Dim cnt As ADODB.Connection
    Dim rst As ADODB.Recordset
    Dim i As Integer
    Dim Headers() As String

    'Instantiate the ADO-objects.
    Set cnt = New ADODB.Connection
    Set rst = New ADODB.Recordset

    ' Open the database connection
    cnt.Open "Driver=" & DbDriverString & ";" & _
            "Server=" & ServerNameOrIpAddress & ";" & _
            "Database=" & DbName & ";" & _
            "Uid=" & UserName & ";" & _
            "Pwd=" & ThePassword & ";"

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

' This sub creates a connection, properly quote headers and data set, and then injects it in the DB
' The sub splits very large matrices into manageable chunks.
Public Function InjectMatrixIntoMySql(ValuesMatrix As Variant, FieldNames As Variant, TableName As String, ServerAddress As String, _
                                      DbName As String, UserName As String, ThePassword As String) As Boolean
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
    
    ' Single-quote ValuesMatrix
    Let QuotedValuesMatrix = AddSingleQuotesToAllArrayElements(ValuesMatrix)
    
    ' Sinble-back-quote FieldNames
    Let QuotedFieldNames = AddSingleBackQuotesToAllArrayElements(FieldNames)
    
    ' Compute the number of chunks (e.g. number of insert statements we need)
    Let NumberOfChunks = Application.Floor_Precise(GetNumberOfRows(ValuesMatrix) / ChunkSize)
    
    ' Insert each chunk
    For i = 1 To NumberOfChunks
        Call ConnectAndExecuteInsertQuery(GetSubMatrix(QuotedValuesMatrix, 1 + (i - 1) * ChunkSize, i * ChunkSize, 1, GetNumberOfColumns(QuotedValuesMatrix)), _
                                          QuotedFieldNames, TableName, ServerAddress, DbName, UserName, ThePassword)
    Next i
    
    ' Insert the remainder after chunking the data set
    If GetNumberOfRows(ValuesMatrix) Mod ChunkSize > 0 Then
        Call ConnectAndExecuteInsertQuery(GetSubMatrix(QuotedValuesMatrix, 1 + CLng(NumberOfChunks) * ChunkSize, GetNumberOfRows(ValuesMatrix), 1, GetNumberOfColumns(QuotedValuesMatrix)), _
                                          QuotedFieldNames, TableName, ServerAddress, DbName, UserName, ThePassword)
    End If
End Function

' Execute a query that returns no data.
Public Sub ConnectAndExecuteInsertQuery(ValuesMatrix As Variant, FieldNames As Variant, TableName As String, ServerAddress As String, _
                                        DbName As String, UserName As String, ThePassword As String)
    Dim Cn As ADODB.Connection
    Dim TheQuery As String
    Dim i As Integer
    Dim j As Integer
    Dim TheRowArray() As Variant
    
    ' Set the database connection and recordset objects
    Set Cn = New ADODB.Connection
       
    ' Open the database connection
    Cn.Open "Driver=" & DbDriverString & ";" & _
            "Server=" & ServerAddress & ";" & _
            "Database=" & DbName & ";" & _
            "Uid=" & UserName & ";" & _
            "Pwd=" & ThePassword & ";"
            
    ' Create query
    Let TheQuery = "INSERT INTO `" & DbName & "`.`" & TableName & "` " & vbCrLf
    Let TheQuery = TheQuery & Convert1DArrayIntoParentheticalExpression(FieldNames) & vbCrLf
    Let TheQuery = TheQuery & " VALUES " & vbCrLf
    
    ReDim TheRowArray(1 To GetNumberOfColumns(ValuesMatrix))
    For i = 1 To GetNumberOfRows(ValuesMatrix)
        For j = 1 To GetNumberOfColumns(ValuesMatrix)
            Let TheRowArray(j) = ValuesMatrix(i, j)
        Next j
    
        Let TheQuery = TheQuery & Convert1DArrayIntoParentheticalExpression(TheRowArray)
        
        If i < GetNumberOfRows(ValuesMatrix) Then
            Let TheQuery = TheQuery & ", " & vbCrLf
        End If
    Next i
    Let TheQuery = TheQuery & ";"
    
    ' Execute the query
    Call Cn.Execute(TheQuery)
    
    ' Close the recordset and connection.
    Call Cn.Close
    
    ' Destroy the recordset and connection objects.
    Set Cn = Nothing
End Sub

' Execute a query that returns no data.
Public Sub RunQuery(TheQuery As String, ServerName As String, DbName As String, UserName As String, ThePassword As String)
    Dim Cn As ADODB.Connection
    ' Set the database connection and recordset objects
    Set Cn = New ADODB.Connection
       
    ' Open the database connection
    Cn.Open "Driver=" & DbDriverString & ";" & _
            "Server=" & ServerName & ";" & _
            "Database=" & DbName & ";" & _
            "Uid=" & UserName & ";" & _
            "Pwd=" & ThePassword & ";"
        
    Call Cn.Execute(TheQuery)
    
    ' Close the recordset and connection.
    Call Cn.Close
    
    ' Destroy the recordset and connection objects.
    Set Cn = Nothing
End Sub

' This function returns TRUE if the DB connection works and FALSE otherwise
Public Function DbConnectionOkayQ(txtHostname As String, txtDatabase As String, txtUsername As String, txtPassword As String) As Boolean
    Dim oConn As ADODB.Connection

    On Error GoTo ErrHandler

    Set oConn = New ADODB.Connection
    oConn.Open "DRIVER=" & DbDriverString & ";" & _
                "SERVER=" & Trim(txtHostname) & ";" & _
                "DATABASE=" & Trim(txtDatabase) & ";" & _
                "USER=" & Trim(txtUsername) & ";" & _
                "PASSWORD=" & Trim(txtPassword) & ";" & _
                "Option=3"
    
    Let DbConnectionOkayQ = True
        
    Call oConn.Close
    Exit Function

ErrHandler:
    Let DbConnectionOkayQ = False
End Function

' This sub populates a DB table with the contents of the same table on a different server
' We assume the source and target servers both have the DB and table with tha same name and the table with the
' same exact structure.  The target table is truncated before copying over the data.
Public Sub CopyTableFromOneDbServerToAnother(SourceServerAddress As String, TheDbName As String, DbTableName As String, SourceDbUsername As String, SourceDbPassword As String, _
                                             TargetServerAddress As String, TargetDbUsername As String, TargetDbPassword As String)
    Dim TheData As Variant
    
    Let TheData = ConnectAndSelect("SELECT * FROM `" & TheDbName & "`.`" & DbTableName & "`;", TheDbName, SourceServerAddress, SourceDbUsername, SourceDbPassword)
    Let TheData = Rest(TheData)
    
    Let TheData = Drop(TransposeMatrix(TheData), Array(GetNumberOfColumns(TheData)))
    Let TheData = TransposeMatrix(TheData)
    
    Call MySql.RunQuery("TRUNCATE `" & TheDbName & "`.`" & DbTableName & "`;", TargetServerAddress, TheDbName, TargetDbUsername, TargetDbPassword)
    
    Call InjectMatrixIntoMySql(ToTemp(TheData, True).Value2, Most(GetTableHeaders(DbTableName, TheDbName, SourceServerAddress, SourceDbUsername, SourceDbPassword)), DbTableName, _
                               TargetServerAddress, TheDbName, TargetDbUsername, TargetDbPassword)
    
End Sub

