Attribute VB_Name = "ExcelSql"
Option Explicit
Option Base 1

' Executes a SELECT statement and returns the results either to the optional target range
' or as an array with or without headers
Public Function SelectUsingSql(SqlQuery As String, _
                               FullPathFileName As String, _
                               Optional ReturnAsArrayQ As Boolean = True, _
                               Optional IncludeHeadersQ As Boolean = True, _
                               Optional TargetRangeUpperLeftCorner As Range = Nothing) As Variant
    Dim cnn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim TheHeaders() As String
    Dim i As Integer
    Dim TheResults As Variant
    Dim RowOffset As Integer
    
    ' Parameter consitency check
    If Not ReturnAsArrayQ And TargetRangeUpperLeftCorner Is Nothing Then
        Let SelectUsingSql = Null
        Exit Function
    End If
    
    Set cnn = CreateConnectionToExcelFile(FullPathFileName)
    Set rs = New ADODB.Recordset
    
    Call rs.Open(SqlQuery, cnn, adOpenKeyset, adLockOptimistic)
    
    If IncludeHeadersQ Then
        ReDim TheHeaders(1 To rs.Fields.Count)
        For i = 0 To rs.Fields.Count - 1
            Let TheHeaders(i + 1) = rs.Fields.Item(i).Name
        Next i
    End If
    
    If ReturnAsArrayQ Then
        Let TheResults = TransposeMatrix(rs.GetRows(rs.RecordCount, 1))
        If IncludeHeadersQ Then
            Let TheResults = Prepend(TheResults, TheHeaders)
        End If
        
        Let SelectUsingSql = TheResults
    Else
        Let RowOffset = IIf(IncludeHeadersQ, 1, 0)
        If IncludeHeadersQ Then
            Call DumpInSheet(TheHeaders, TargetRangeUpperLeftCorner.Range("A1"))
        End If
        
        Call TargetRangeUpperLeftCorner.Offset(RowOffset, 0).CopyFromRecordset(rs)
    End If
    
    Call CloseRecordSet(rs)
    Call CloseConnection(rs)
End Function

' Executes a SELECT statement and returns the results either to the optional target range
' or as an array with or without headers
Public Sub UpdateUsingSql(SqlQuery As String, _
                          FullPathFileName As String)
    Dim cnn As ADODB.Connection
    Dim rs As ADODB.Recordset
    
    Set cnn = CreateConnectionToExcelFile(FullPathFileName)
    Set rs = New ADODB.Recordset
    
    Call rs.Open(SqlQuery, cnn, adOpenKeyset, adLockOptimistic)
        
    Call CloseRecordSet(rs)
    Call CloseConnection(rs)
End Sub

' Instantiates a connection and connects to an excel file.
Private Function CreateConnectionToExcelFile(FullPathFileName As String) As ADODB.Connection
    Dim cnn As ADODB.Connection
    
    Set cnn = New ADODB.Connection
    
    Let cnn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                               "Data Source=" & FullPathFileName & ";" & _
                               "Extended Properties=""Excel 8.0;"""
    
    Call cnn.Open
    
    Set CreateConnectionToExcelFile = cnn
End Function

' Closes a record set and sets the reference to nothing
Private Sub CloseRecordSet(rs As ADODB.Recordset)
    If rs Is Nothing Then
        Exit Sub
    End If

    If rs.State = adStateOpen Then
        Call rs.Close
    End If
    
    Let rs.CursorLocation = adUseClient
    
    Set rs = Nothing
End Sub

' Closes a record set and sets the reference to nothing
Private Sub CloseConnection(cn As ADODB.Connection)
    If cn Is Nothing Then
        Exit Sub
    End If
    
    Set cn = Nothing
End Sub
