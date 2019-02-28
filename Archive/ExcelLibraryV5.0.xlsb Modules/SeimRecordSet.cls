VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SeimRecordSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Option Base 1

' Private instance variable to hold the trade database
Public RangeDb As Range
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

' Custom initializer. The first row is ignored because it is expected to be the
' header row
Public Sub InitializeWithRange(ARange As Range)
    Set Me.UnderlyingWorkbook = ARange.Worksheet.Parent
    
    If ARange.Rows.Count < 2 Then
        Exit Sub
    End If
    
    Set RangeDb = ARange.Offset(1, 0).Resize(ARange.Rows.Count - 1, Me.GetNumberOfColumns)
    Call FormatUnderlyingRange
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

Property Get GetHeaders() As Variant
    Dim ASeimRecord As New SeimRecord

    Let GetHeaders = ASeimRecord.GetHeaders
End Property

Property Get GetNumberOfColumns() As Integer
    Dim ASeimRecord As New SeimRecord
    
    Let GetNumberOfColumns = ASeimRecord.GetNumberOfColumns
End Property

Property Get GetIthRecord(i As Integer) As SeimRecord
    Dim AnAlert As SeimRecord
    
    Set AnAlert = New SeimRecord
    
    Call AnAlert.InitializeWithRange(Me.GetRange.Rows(i))
    
    Set GetIthRecord = AnAlert
End Property

Property Get GetDateTimes() As Range
    Dim ARecord As New SeimRecord
    
    Call ARecord.InitializeWithRange(Me.GetRange.Rows(1))
    
    Set GetDateTimes = Me.GetRange.Columns(ARecord.GetDateTime.Column)
End Property

Property Get GetUsers() As Range
    Dim ARecord As New SeimRecord
    
    Call ARecord.InitializeWithRange(Me.GetRange.Rows(1))
    
    Set GetUsers = Me.GetRange.Columns(ARecord.GetUser.Column)
End Property

Property Get GetTypes() As Range
    Dim ARecord As New SeimRecord
    
    Call ARecord.InitializeWithRange(Me.GetRange.Rows(1))
    
    Set GetTypes = Me.GetRange.Columns(ARecord.GetType.Column)
End Property

Property Get GetActions() As Range
    Dim ARecord As New SeimRecord
    
    Call ARecord.InitializeWithRange(Me.GetRange.Rows(1))
    
    Set GetActions = Me.GetRange.Columns(ARecord.GetAction.Column)
End Property

Property Get GetSourceTypes() As Range
    Dim ARecord As New SeimRecord
    
    Call ARecord.InitializeWithRange(Me.GetRange.Rows(1))
    
    Set GetSourceTypes = Me.GetRange.Columns(ARecord.GetSourceType.Column)
End Property

Property Get GetSourceIds() As Range
    Dim ARecord As New SeimRecord
    
    Call ARecord.InitializeWithRange(Me.GetRange.Rows(1))
    
    Set GetSourceId = Me.GetRange.Columns(ARecord.GetSourceId.Column)
End Property

Property Get GetAppliances() As Range
    Dim ARecord As New SeimRecord
    
    Call ARecord.InitializeWithRange(Me.GetRange.Rows(1))
    
    Set GetAppliances = Me.GetRange.Columns(ARecord.GetAppliance.Column)
End Property

Property Get GetXmls() As Range
    Dim ARecord As New SeimRecord
    
    Call ARecord.InitializeWithRange(Me.GetRange.Rows(1))
    
    Set GetXmls = Me.GetRange.Columns(ARecord.GetXml.Column)
End Property

Property Get GetSourceFileNames() As Range
    Dim ARecord As New SeimRecord
    
    Call ARecord.InitializeWithRange(Me.GetRange.Rows(1))
    
    Set GetSourceFileNames = Me.GetRange.Columns(ARecord.GetSourceFileName.Column)
End Property

Property Get GetComplianceComments() As Range
    Dim ARecord As New SeimRecord
    
    Call ARecord.InitializeWithRange(Me.GetRange.Rows(1))
    
    Set GetComplianceComments = Me.GetRange.Columns(ARecord.GetComplianceComments.Column)
End Property

Property Get GetPrimaryKeys() As Range
    Dim ARecord As New SeimRecord
    
    Call ARecord.InitializeWithRange(Me.GetRange.Rows(1))
    
    Set GetPrimaryKeys = Me.GetRange.Columns(ARecord.GetPrimaryKey.Column)
End Property

Public Sub InitializedWithLogFile(TheFileName As String)
    Dim LogDictionary As Dictionary
    
    Set Me.UnderlyingWorkbook = Application.Workbooks.Add
    Call RemoveAllOtherWorksheets(Me.UnderlyingWorkbook.Worksheets(1))
    
    Set LogDictionary = GetLogFileContents(TheFileName)
        
    If LogDictionary.Count > 0 Then
        Call DumpInSheet(Me.GetHeaders, Me.UnderlyingWorkbook.Worksheets(1).Range("A1"))
        Call DumpInSheet(Me.ConvertDictionaryTo2DArray(LogDictionary), Me.UnderlyingWorkbook.Worksheets(1).Range("A2"))
        Call Me.InitializeWithRange(Me.UnderlyingWorkbook.Worksheets(1).Range("A1").CurrentRegion)
        Let Me.GetComplianceComments.Value2 = "NULL"
        Let Me.GetSourceFileNames.Value2 = ExtractFilenameAndExtensionFromFullPathFileName(TheFileName)
    End If
    
    Call Me.CorrectDateFormats
End Sub

Public Function ConvertDictionaryTo2DArray(LogDictionary As Dictionary) As Variant
    Dim ARecord As New SeimRecord
    Dim TheResults() As Variant
    Dim i As Integer
    Dim xDoc As MSXML2.DOMDocument60
    Dim NodeValue As Variant
    
    ' Initialize helper range
    Call ARecord.InitializeWithRange(ThisWorkbook.Worksheets(1).Range("A1"))
    
    ' Instantiate a DOMDocument object
    Set xDoc = New MSXML2.DOMDocument60

    If LogDictionary.Count = 0 Then
        Let ConvertDictionaryTo2DArray = Array()
        Exit Function
    End If
    
    ReDim TheResults(LogDictionary.Count, Me.GetNumberOfColumns - 1)
    
    For i = 1 To LogDictionary.Count
        ' Initialize the DOMDocument with the XML data
        Call xDoc.LoadXML(LogDictionary.Item(Key:=i))
        
        Let TheResults(i, ARecord.GetDateTime.Column) = GetNodeValueOrEmpty(xDoc.FirstChild.SelectSingleNode("TIMESTAMP"))
        Let TheResults(i, ARecord.GetUser.Column) = GetNodeValueOrEmpty(xDoc.FirstChild.SelectSingleNode("USER"))
        Let TheResults(i, ARecord.GetType.Column) = GetNodeValueOrEmpty(xDoc.FirstChild.SelectSingleNode("TYPE"))
        Let TheResults(i, ARecord.GetAction.Column) = GetNodeValueOrEmpty(xDoc.FirstChild.SelectSingleNode("ACTION"))
        Let TheResults(i, ARecord.GetSourceType.Column) = GetNodeValueOrEmpty(xDoc.FirstChild.SelectSingleNode("SOURCETYPE"))
        Let TheResults(i, ARecord.GetSourceId.Column) = GetNodeValueOrEmpty(xDoc.FirstChild.SelectSingleNode("SOURCEID"))
        Let TheResults(i, ARecord.GetAppliance.Column) = GetNodeValueOrEmpty(xDoc.FirstChild.SelectSingleNode("APPLIANCE"))
        Let TheResults(i, ARecord.GetXml.Column) = xDoc.XML
    Next i
    
    Let ConvertDictionaryTo2DArray = TheResults
End Function

Private Function GetNodeValueOrEmpty(arg As Variant)
    If arg Is Nothing Then
        Let GetNodeValueOrEmpty = Empty
    Else
        Let GetNodeValueOrEmpty = arg.nodeTypedValue
    End If
End Function

Public Function GetLogFileContents(TheFileName As String) As Dictionary
    Dim FsoObj As Scripting.FileSystemObject
    Dim FileObj As Scripting.TextStream
    Dim ATxtLine As String
    Dim TheTag As String
    Dim ThePayload As String
    Dim XmlResultString As String
    Dim NumberOfRecords As Integer
    Dim RecordsDictionary As Dictionary
    Dim NewRecord As Boolean
    
    Let NewRecord = False
    Set RecordsDictionary = New Dictionary
    Set FsoObj = New Scripting.FileSystemObject

    Let NumberOfRecords = 0

   Set FileObj = FsoObj.OpenTextFile(Filename:=TheFileName, IOMode:=ForReading)
    
    ' Move to the first record divider
    If Not FileObj.AtEndOfStream Then
        Let ATxtLine = FileObj.ReadLine
    End If
    
    Do While Not FileObj.AtEndOfStream And InStr(ATxtLine, "_______________________") = 0
        Let ATxtLine = FileObj.ReadLine
    Loop
    
    If FileObj.AtEndOfStream Then
        Exit Function
    Else
        Let NewRecord = True
        Let XmlResultString = "<LOG>"
    End If
    
    ' Start assembling items for the dictionary
    Do While Not FileObj.AtEndOfStream    ' Loop until end of file.
        Let ATxtLine = FileObj.ReadLine

        ' Process line into XML
        ' Determine if this is a new log or part of the current one
        If InStr(ATxtLine, "_______________________") <> 0 Then
            Let XmlResultString = XmlResultString & vbCrLf & "</LOG>"
            
            If Trim(XmlResultString) <> "" Then
                Let NumberOfRecords = NumberOfRecords + 1
                Call RecordsDictionary.Add(Key:=NumberOfRecords, Item:=XmlResultString)
            End If
        
            Let NewRecord = True
            Let XmlResultString = "<LOG>"
        ElseIf ATxtLine <> "" Then
            Let TheTag = GetLineTag(ATxtLine)
            
            ' Handle the special case of the details field
            If TheTag <> "DETAILS" Then
                Let ThePayload = GetLinePayload(ATxtLine)
                Let XmlResultString = XmlResultString & vbCrLf & "<" & TheTag & ">" & ThePayload & "</" & TheTag & ">"
            Else
                Let XmlResultString = XmlResultString & vbCrLf & "<DETAILS>"
                If GetLinePayload(ATxtLine) <> "" Then
                    Let XmlResultString = XmlResultString & vbCrLf & GetLinePayload(ATxtLine)
                End If
            
                Do While Not FileObj.AtEndOfStream And NewRecord <> False
                    Let ATxtLine = FileObj.ReadLine
    
                    If InStr(ATxtLine, "_______________________") = 0 Then
                        If ATxtLine <> "" Then
                            Let XmlResultString = XmlResultString & vbCrLf & RemoveIllegalCharacters(ATxtLine)
                        End If
                    Else
                        Let NewRecord = False
                    End If
                Loop
                
                Let XmlResultString = XmlResultString & vbCrLf & "</DETAILS>" & vbCrLf & "</LOG>"
                
                If Trim(XmlResultString) <> "" Then
                    Let NumberOfRecords = NumberOfRecords + 1
                
                    Call RecordsDictionary.Add(Key:=NumberOfRecords, Item:=XmlResultString)
                End If
                
                If Not FileObj.AtEndOfStream Then
                    Let NewRecord = True
                    Let XmlResultString = "<LOG>"
                End If
            End If
        End If
    Loop

    Call FileObj.Close
    
    Set GetLogFileContents = RecordsDictionary
End Function

Private Function GetLineTag(ATxtLine As String) As String
    Let GetLineTag = UCase(RemoveIllegalTagCharacters(Trim(Split(ATxtLine, ":")(0))))
End Function

Private Function GetLinePayload(ATxtLine As String) As String
    If GetArrayLength(Split(ATxtLine, ":")) < 2 Then
        Let GetLinePayload = ""
    Else
        If Split(ATxtLine, ":")(1) = "" Then
            Let GetLinePayload = ""
        Else
            Let GetLinePayload = Right(ATxtLine, Len(ATxtLine) - InStr(ATxtLine, ":"))
            Let GetLinePayload = Trim(RemoveIllegalCharacters(GetLinePayload))
        End If
    End If
End Function

Private Function RemoveIllegalCharacters(arg As String) As String
    Let RemoveIllegalCharacters = Replace(Replace(Replace(Replace(Replace(arg, "/", "-"), "<", "&lt;"), ">", "&gt;"), "&", "&amp;"), """", "")
End Function

Private Function RemoveIllegalTagCharacters(arg As String) As String
    Let RemoveIllegalTagCharacters = Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(arg, "/", ""), "<", "&lt;"), ">", "&gt;"), "/", ""), ")", "&gt;"), "(", "&lt;"), " ", ""), """", "&quot;")
End Function

' Format underlying range appropriately
Public Sub FormatUnderlyingRange()
    Let Me.GetRange.Rows(1).Offset(-1, 0).Font.Bold = True
    Call Me.GetRange.Range("A1").Select
    Let Me.GetRange.Worksheet.Parent.Windows(1).FreezePanes = True
    
    Call Me.GetDateTimes.EntireColumn.AutoFit
    
    Let Me.GetXmls.EntireColumn.ColumnWidth = 80
    
    Let Me.GetRange.EntireColumn.VerticalAlignment = xlVAlignTop
    Let Me.GetRange.EntireColumn.HorizontalAlignment = xlHAlignLeft
    Let Me.GetDateTimes.EntireColumn.NumberFormat = "yyyy-mm-dd hh:mm:ss"
    
    Call Me.GetSourceFileNames.EntireColumn.AutoFit
End Sub

' This re-formats the date/time stamps after they are read into the workbook from text files (dump by SEIM)
Public Sub CorrectDateFormats()
    Dim AnAlert As SeimRecord
    Dim TempSplitString As Variant
    Dim DateTimeStamp As Date
    Dim i As Integer
    
    For i = 1 To Me.GetDbLength
        Set AnAlert = Me.GetIthRecord(i)
        Let TempSplitString = Split(AnAlert.GetDateTime.Value2, ":")
        Let AnAlert.SetDateTime = CDate(TempSplitString(0) & " " & TempSplitString(1) & ":" & TempSplitString(2) & ":" & (Application.Floor_Precise(Split(TempSplitString(3), " ")(0)) Mod 60))
    Next i
    
    Let Me.GetDateTimes.EntireColumn.NumberFormat = "yyyy-mm-dd hh:mm:ss"
End Sub

Public Sub InjectIntoMySql()
    Dim TheData As Variant
    Dim TheHeaders As Variant
    Dim TheDateStrings() As String
    Dim OriginalDates As Variant
    Dim i As Integer
    Dim NumberOfBlocks As Long
    Dim LeftOver As Long
    Const MaxBlockSize As Long = 1000
    
    Let TheHeaders = ConnectAndSelect("SELECT * FROM `monitoring`.`seimlogs` LIMIT 0,0;", _
                                      "monitoring", _
                                      DbServerAddress, _
                                      DbUserName, _
                                      DbPassword)
                                      
    Let TheHeaders = AddSingleBackQuotesToAllArrayElements(Most(ConvertTo1DArray(GetRow(TheHeaders, 1))))
    
    Let OriginalDates = Me.GetDateTimes.Value2
    ReDim TheDateStrings(Me.GetDbLength)
    For i = 1 To Me.GetDbLength
        Let TheDateStrings(i) = "'" & ConvertDateTimeToMySQLFormat(Me.GetIthRecord(i).GetDateTime.Text)
    Next i
    
    Let Me.GetDateTimes.NumberFormat = "@"
    Let Me.GetDateTimes.Value2 = Application.Transpose(TheDateStrings)
    
    Let TheData = GetSubMatrix(DoubleQuote2DArray(Me.GetRange.Value2), 1, Me.GetDbLength, 1, Me.GetNumberOfColumns - 1)
    
    Let Me.GetDateTimes.Value2 = OriginalDates
    
    Let NumberOfBlocks = Application.Floor_Precise(Me.GetDbLength / MaxBlockSize)
    For i = 1 To NumberOfBlocks
        Call ConnectAndExecuteInsertQuery(GetSubMatrix(TheData, 1 + MaxBlockSize * (i - 1), MaxBlockSize * i, 1, ArrayFormulas.GetNumberOfColumns(TheData)), _
                                          TheHeaders, "seimlogs", DbServerAddress, "monitoring", DbUserName, DbPassword)
    Next i
    
    If Me.GetDbLength > NumberOfBlocks * MaxBlockSize Then
        Call ConnectAndExecuteInsertQuery(GetSubMatrix(TheData, 1 + NumberOfBlocks * MaxBlockSize, Me.GetDbLength, 1, ArrayFormulas.GetNumberOfColumns(TheData)), _
                                          TheHeaders, "seimlogs", DbServerAddress, "monitoring", DbUserName, DbPassword)

    End If
    
    Call Me.FormatUnderlyingRange
End Sub


