VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BloombergEquityAlerts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' This class requires a folder named "Bloomberg ALTR" as a sub-folder of "Inbox."
Option Base 1
Option Explicit

Public UnderlyingWorbook As Workbook
Public RangeDb As Range

' This destructor closes (without saving changes) the workbook holding the underlying data
' This is the class destructor.
Public Sub Class_Terminate()
    If Not (Me.UnderlyingWorbook Is Nothing) Then
        If Not Me.UnderlyingWorbook Is ThisWorkbook Then
            Call Me.UnderlyingWorbook.Close(SaveChanges:=False)
        End If
    End If
End Sub

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

Public Function GetHeaders() As Variant
    Dim ARow As New BloombergEquityAlertRow

    Let GetHeaders = ARow.GetHeaders
End Function

Public Function GetNumberOfColumns() As Integer
    Dim ARow As New BloombergEquityAlertRow

    Let GetNumberOfColumns = ARow.GetNumberOfColumns
End Function

Property Get GetDates() As Range
    Dim AnAlert As New BloombergEquityAlertRow
    
    Call AnAlert.InitializeWithRange(Me.GetRange.Rows(1))
    
    Set GetDates = Me.GetRange.Columns(AnAlert.GetDate.Column)
End Property

Property Get GetBloombergTickers() As Range
    Dim AnAlert As New BloombergEquityAlertRow
    
    Call AnAlert.InitializeWithRange(Me.GetRange.Rows(1))
    
    Set GetBloombergTickers = Me.GetRange.Columns(AnAlert.GetBloombergTicker.Column)
End Property

Property Get GetReferences() As Range
    Dim AnAlert As New BloombergEquityAlertRow
    
    Call AnAlert.InitializeWithRange(Me.GetRange.Rows(1))
    
    Set GetReferences = Me.GetRange.Columns(AnAlert.GetReference.Column)
End Property

Property Get GetSubjects() As Range
    Dim AnAlert As New BloombergEquityAlertRow
    
    Call AnAlert.InitializeWithRange(Me.GetRange.Rows(1))
    
    Set GetSubjects = Me.GetRange.Columns(AnAlert.GetSubject.Column)
End Property

Property Get GetBodies() As Range
    Dim AnAlert As New BloombergEquityAlertRow
    
    Call AnAlert.InitializeWithRange(Me.GetRange.Rows(1))
    
    Set GetBodies = Me.GetRange.Columns(AnAlert.GetBody.Column)
End Property

Property Get GetTypes() As Range
    Dim AnAlert As New BloombergEquityAlertRow
    
    Call AnAlert.InitializeWithRange(Me.GetRange.Rows(1))
    
    Set GetTypes = Me.GetRange.Columns(AnAlert.GetType.Column)
End Property

Property Get GetQualifiers() As Range
    Dim AnAlert As New BloombergEquityAlertRow
    
    Call AnAlert.InitializeWithRange(Me.GetRange.Rows(1))
    
    Set GetQualifiers = Me.GetRange.Columns(AnAlert.GetQualifier.Column)
End Property

Property Get GetXmlBodies() As Range
    Dim AnAlert As New BloombergEquityAlertRow
    
    Call AnAlert.InitializeWithRange(Me.GetRange.Rows(1))
    
    Set GetXmlBodies = Me.GetRange.Columns(AnAlert.GetXmlBody.Column)
End Property

Property Get GetStartDates() As Range
    Dim AnAlert As New BloombergEquityAlertRow
    
    Call AnAlert.InitializeWithRange(Me.GetRange.Rows(1))
    
    Set GetStartDates = Me.GetRange.Columns(AnAlert.GetStartDate.Column)
End Property

Property Get GetEndDates() As Range
    Dim AnAlert As New BloombergEquityAlertRow
    
    Call AnAlert.InitializeWithRange(Me.GetRange.Rows(1))
    
    Set GetEndDates = Me.GetRange.Columns(AnAlert.GetEndDate.Column)
End Property

Property Get GetVerifiedQs() As Range
    Dim AnAlert As New BloombergEquityAlertRow
    
    Call AnAlert.InitializeWithRange(Me.GetRange.Rows(1))
    
    Set GetVerifiedQs = Me.GetRange.Columns(AnAlert.GetVerifiedQ.Column)
End Property

Property Get GetEqyFundTickers() As Range
    Dim AnAlert As New BloombergEquityAlertRow
    
    Call AnAlert.InitializeWithRange(Me.GetRange.Rows(1))
    
    Set GetEqyFundTickers = Me.GetRange.Columns(AnAlert.GetEqyFundTicker.Column)
End Property

' Must include a header row
Public Sub InitializeWithRange(ARange As Range)
    Set Me.UnderlyingWorbook = ARange.Worksheet.Parent
    
    If ARange.Rows.Count < 2 Then
        Exit Sub
    End If
    
    Set RangeDb = ARange.Offset(1, 0).Resize(ARange.Rows.Count - 1, Me.GetNumberOfColumns)
    
    Call FormatUnderlyingRange
End Sub

' This function returns a dictionary instance.
' The keys are the subject lines of the emails from Bloomberg
' The items are arrays of strings, where each array elt is one line
' in the email.
Public Sub InitializeWithOutlook()
    Dim AnAction As New BloombergEquityAlertRow
    Dim OutlookInboxFolder As Outlook.Folder
    Dim TheResults() As String
    Dim QueryResults As Variant
    Dim i As Integer
    Dim wbk As Workbook
    Dim DatePart As String
    Dim TimePart As String
    Dim AmVsPm As String
    Dim TheQuery As String

    Set OutlookInboxFolder = Me.GetBlombergAlertsFolder()

    ' Exit if there are not messages in the folder
    If OutlookInboxFolder.Items.Count = 0 Then
        Exit Sub
    End If
    
    ' Instantiate a workbook to hold the
    Set wbk = Application.Workbooks.Add
    Call RemoveAllOtherWorksheets(wbk.Worksheets(1))
    
    Call AnAction.InitializeWithRange(wbk.Worksheets(1).Range("A1"))

    ReDim TheResults(1 To 1 + Me.GetBlombergAlertsFolder.Items.Count, 1 To Me.GetNumberOfColumns)
    
    ' Set the headers
    For i = 1 To Me.GetNumberOfColumns
        Let TheResults(1, i) = Me.GetHeaders(i)
    Next i
    
    For i = 1 To GetBlombergAlertsFolder.Items.Count
        Let TheResults(i + 1, 1) = OutlookInboxFolder.Items(i).SentOn
        Let TheResults(i + 1, 2) = Me.GetReferenceFromBodyString(OutlookInboxFolder.Items(i).Body)
        Let TheResults(i + 1, 3) = Me.GetTickerFromBodyString(OutlookInboxFolder.Items(i).Body)
        Let TheResults(i + 1, 4) = OutlookInboxFolder.Items(i).Subject
        Let TheResults(i + 1, 5) = Me.RemoveBlankLinesFromBody(OutlookInboxFolder.Items(i).Body)
        Let TheResults(i + 1, 8) = Me.ConvertBodyToXml(TheResults(i + 1, 5))
        Let TheResults(i + 1, 11) = 0
    Next i

    ' Pre-format ranges
    Let wbk.Worksheets(1).Range("C2:H2").Resize(GetBlombergAlertsFolder.Items.Count + 1, 6).NumberFormat = "@"

    ' Dump messages into a range
    Call DumpInSheet(TheResults, wbk.Worksheets(1).Range("A1"))

    ' Initialize with range
    Call Me.InitializeWithRange(wbk.Worksheets(1).Range("A1").CurrentRegion)

    ' Determine and record alert types now that the object has been initialized
    For i = 1 To Me.GetDbLength
        Call Me.GetIthAlert(i).ComputeAlertStringType
        Call Me.GetIthAlert(i).ComputeLifeCycleDates
    Next i
    
    ' Post-process the date column so it becomes a real Excel date/time data type
    For i = 1 To Me.GetDbLength
        Let DatePart = CStr(ConvertDateToSerial(CDate(Split(Trim(Me.GetIthAlert(i).GetDate.Value2), " ")(0))))
        Let TimePart = Split(Trim(Me.GetIthAlert(i).GetDate.Value2), " ")(1)
        Let AmVsPm = Right(Trim(Me.GetIthAlert(i).GetDate.Value2), 2)

        Let Me.GetIthAlert(i).SetDate = Left(DatePart, 4) & "-" & Mid(DatePart, 5, 2) & "-" & Right(DatePart, 2) & " " & TimePart & " " & AmVsPm
    Next i
    
    ' Post-process the EQY_FUND_COLUMN.  Pull it from `equity250v1`.`bloomberg-eqs` if available
    ' Return the list of tickers with no EQY_FUND_TICKER as a 1D array so the user may pull them directly from
    ' Bloomberg
    
    ' Instantiate the latest Bloomberg EQS
'***HERE

    Call FormatUnderlyingRange
End Sub

' Format underlying range appropriately
Public Sub FormatUnderlyingRange()
    Let Me.GetDates.EntireColumn.NumberFormat = "yyyy-mm-dd hh:mm:ss"
    Call Me.GetDates.EntireColumn.AutoFit
    Call Me.GetReferences.EntireColumn.AutoFit
    Call Me.GetBloombergTickers.EntireColumn.AutoFit
    Let Me.GetSubjects.EntireColumn.ColumnWidth = 40
    Let Me.GetSubjects.EntireColumn.WrapText = True
    Let Me.GetBodies.EntireColumn.ColumnWidth = 60
    Call Me.GetTypes.EntireColumn.AutoFit
    Call Me.GetQualifiers.EntireColumn.AutoFit
    Let Me.GetXmlBodies.EntireColumn.ColumnWidth = 80
    Let Me.GetXmlBodies.EntireColumn.WrapText = True
    Call Me.GetStartDates.EntireColumn.AutoFit
    Call Me.GetEndDates.EntireColumn.AutoFit
    Call Me.GetVerifiedQs.EntireColumn.AutoFit
    Call Me.GetEqyFundTickers.EntireColumn.AutoFit
    Let Me.GetRange.VerticalAlignment = xlTop
    Let Me.GetRange.Rows(1).Offset(-1, 0).Font.Bold = True
    Call Me.GetRange.Range("A1").Select
    Let Me.GetRange.Worksheet.Parent.Windows(1).FreezePanes = True
End Sub

' Helper function for Public Function GetBloombergAlerts() in this module
Public Function GetBlombergAlertsFolder() As Outlook.Folder
    Dim OutlookApp As Outlook.Application
    Dim OutlookNamespace As Outlook.Namespace
    Dim OutlookFolder As Outlook.Folder
    Dim OutlookFolderSet As Outlook.Folders
    Dim obj As Object
    Dim i As Integer

    ' Instantiate Outlook application
    Set OutlookApp = New Outlook.Application

    ' Set reference to the MAPI namespace
    Set OutlookNamespace = OutlookApp.GetNamespace("MAPI")

    ' Identify top-level folder
    For i = 1 To OutlookNamespace.Folders.Count
        If InStr(UCase(OutlookNamespace.Folders.Item(i).Name), UCase("Public Folders - ")) = 0 Then
            Set OutlookFolderSet = OutlookNamespace.Folders.Item(i).Folders

            Exit For
        End If
    Next i

    ' Get the correct folder holding the Bloomberg alerts
    For i = 1 To OutlookFolderSet.Count
        If UCase(OutlookFolderSet.Item(i).Name) = UCase("Inbox") Then
            Set OutlookFolderSet = OutlookFolderSet.Item(i).Folders

            Exit For
        End If
    Next i

    For i = 1 To OutlookFolderSet.Count
        If UCase(OutlookFolderSet.Item(i).Name) = UCase("Bloomberg ALTR") Then
            Set GetBlombergAlertsFolder = OutlookFolderSet.Item(i)
        End If
    Next i
End Function

Property Get GetIthAlert(i As Integer) As BloombergEquityAlertRow
    Dim AnAlert As BloombergEquityAlertRow
    
    Set AnAlert = New BloombergEquityAlertRow
    
    Call AnAlert.InitializeWithRange(Me.GetRange.Rows(i))
    
    Set GetIthAlert = AnAlert
End Property

' -------------------------------------------------------
' Helper functions to process the body of alerts
' -------------------------------------------------------
Public Function RemoveBlankLinesFromBody(TheBody As String) As String
    Dim TheCorrectBody As String
    Dim obj As Variant

    For Each obj In Split(TheBody, vbCrLf)
        If obj <> "" And obj <> " " Then
            Let TheCorrectBody = TheCorrectBody & obj & vbCrLf
        End If
    Next
    Let RemoveBlankLinesFromBody = Trim(TheCorrectBody)
    
    If Len(TheCorrectBody) > 0 And Right(TheCorrectBody, 1) <> vbCrLf Then
        Let RemoveBlankLinesFromBody = Trim(Left(TheCorrectBody, Len(TheCorrectBody) - 1))
    End If
End Function

Public Function ConvertBodyToXml(TheBodyString As String) As String
    Dim BodyLinesDict As Dictionary
    Dim BodyLinesArray As Variant
    Dim OpeningTag As String
    Dim ClosingTag As String
    Dim i As Integer
    Dim c As Integer
    
    Set BodyLinesDict = New Dictionary
    
    ' Split the email's body into lines
    Let BodyLinesArray = Split(Trim(TheBodyString), vbCrLf)
    Let c = 1
    For i = LBound(BodyLinesArray) To UBound(BodyLinesArray)
        If BodyLinesArray(i) <> "" And BodyLinesArray(i) <> " " Then
        
            ' If this line has a tag, then it is not the one containing Bloomberg's reference for this corporate action
            If Me.BodyLineHasTagQ(CStr(BodyLinesArray(i))) Then
                If UCase(Me.GetLineTag(CStr(BodyLinesArray(i)))) <> "WARNING" Then
                    Let OpeningTag = "<" & Replace(Replace(UCase(Me.GetLineTag(CStr(BodyLinesArray(i)))), " ", ""), "%", "PERCENT") & ">"
                    Let ClosingTag = "</" & Replace(Replace(UCase(Me.GetLineTag(CStr(BodyLinesArray(i)))), " ", ""), "%", "PERCENT") & ">"
                    
                    Call BodyLinesDict.Add(Key:=c, Item:=OpeningTag & Replace(Me.GetLinePayload(CStr(BodyLinesArray(i))), "&", "and") & ClosingTag)
                End If
            ' This line contains Bloomberg's reference for this corporate action
            Else
                Let OpeningTag = "<REFERENCE>"
                Let ClosingTag = "</REFERENCE>"
                
                Call BodyLinesDict.Add(Key:=c, Item:=OpeningTag & Me.GetReferenceFromBodyString(TheBodyString) & ClosingTag)
            End If
                
                Let c = c + 1
        End If
    Next i
    
    ' Add starting and ending tags for the XML document
    Let ConvertBodyToXml = "<ALERT>" & StringConcatenate(BodyLinesDict.Items) & "</ALERT>"
End Function

' This function returns True if the body line has a tag and False otherwise
Public Function BodyLineHasTagQ(ALine As String) As Boolean
    If InStr(ALine, ":") = 0 Then
        Let BodyLineHasTagQ = False
    Else
        Let BodyLineHasTagQ = True
    End If
End Function

' This function returns True is the line has a non-empty body
Public Function BodyLineHasPayloadQ(ALine As String) As Boolean
    If InStr(ALine, ":") = Len(ALine) Then
        Let BodyLineHasPayloadQ = False
    Else
        Let BodyLineHasPayloadQ = True
    End If
End Function

' This function returns a line's tag (e.g. thetag: ...) from an alert's body.
' The tag is obviously the text before ":"
Public Function GetLineTag(ALine As String) As String
    If Not BodyLineHasTagQ(ALine) Then
        Let GetLineTag = ""
        
        Exit Function
    End If
    
    Let GetLineTag = Trim(Left(ALine, InStr(ALine, ":") - 1))
End Function

' This function returns a line's tag (e.g. thetag: ...) from an alert's body.
' The tag is obviously the text before ":"
Public Function GetLinePayload(ALine As String) As String
    If Not BodyLineHasPayloadQ(ALine) Then
        Let GetLinePayload = ""
        
        Exit Function
    End If
    
    Let GetLinePayload = IIf(Trim(Right(ALine, Len(ALine) - InStr(ALine, ":"))) = "-", "", Trim(Right(ALine, Len(ALine) - InStr(ALine, ":"))))
End Function

' This search the given string (the body of the email containing the BB alert)
Public Function GetReferenceFromBodyString(BodyString As String) As Long
    Dim StartPos As Integer
    Dim EndPos As Integer
    
    Let StartPos = InStr(Trim(BodyString), "CACX ")
    Let EndPos = InStr(Trim(BodyString), "<GO>")
    Let GetReferenceFromBodyString = CLng(Mid(Trim(BodyString), StartPos + 5, EndPos - StartPos - 5))
End Function

' This returns the bloomberg ticker in the given string (the body of the email containing the BB alert)
Public Function GetTickerFromBodyString(BodyString As String) As String
    Dim StartPos As Integer
    Dim EndPos As Integer
    Dim SubStr As String
    
    Let StartPos = InStr(Trim(BodyString), "Company Name")
    
    If StartPos = 0 Then
        Let GetTickerFromBodyString = ""
        Exit Function
    End If
    
    Let EndPos = InStr(StartPos, Trim(BodyString), vbCrLf)
    If StartPos <> 0 And EndPos <> 0 And StartPos < EndPos Then
        Let SubStr = Mid(Trim(BodyString), StartPos, EndPos - StartPos + 1)
        
        If Not Me.BodyLineHasTagQ(Trim(SubStr)) Then
            Let GetTickerFromBodyString = ""
        Else
            Let SubStr = Me.GetLinePayload(SubStr)
            
            Let StartPos = InStr(SubStr, "(")
            Let EndPos = InStr(SubStr, ")")
            
            If StartPos <> 0 And EndPos <> 0 And EndPos > StartPos Then
                Let GetTickerFromBodyString = Mid(SubStr, StartPos + 1, EndPos - StartPos - 1)
            Else
                Let GetTickerFromBodyString = ""
            End If
        End If
    Else
        Let GetTickerFromBodyString = ""
    End If
End Function

Public Sub InjectIntoMySql()
    Dim TheData As Variant
    Dim TheHeaders As Variant
    Dim TheDateStrings() As String
    Dim OriginalDates As Variant
    Dim i As Integer
    
    Let TheHeaders = ConnectAndSelect("SELECT * FROM `equity250v1`.`bloombergequityalerts` LIMIT 0,0;", _
                                      "equity250v1", _
                                      DbServerAddress, _
                                      DbUserName, _
                                      DbPassword)
                                      
    Let TheHeaders = AddSingleBackQuotesToAllArrayElements(Most(ConvertTo1DArray(GetRow(TheHeaders, 1))))
    
    Let OriginalDates = Me.GetDates.Value2
    ReDim TheDateStrings(Me.GetDbLength)
    For i = 1 To Me.GetDbLength
        Let TheDateStrings(i) = "'" & ConvertDateTimeToMySQLFormat(Me.GetIthAlert(i).GetDate.Text)
    Next i
    
    Let Me.GetDates.NumberFormat = "@"
    Let Me.GetDates.Value2 = Application.Transpose(TheDateStrings)
    
    Let TheData = DoubleQuote2DArray(Me.GetRange.Value2)
    
    Let Me.GetDates.Value2 = OriginalDates
    
    Call ConnectAndExecuteInsertQuery(TheData, TheHeaders, "bloombergequityalerts", DbServerAddress, "equity250v1", _
                                      DbUserName, DbPassword)
                                      
    Call Me.FormatUnderlyingRange
End Sub

Public Sub InitializeWithMySql(Optional BloombergTickerArray As Variant, Optional AnnouncementDate As Date = Empty, Optional StartDate As Date = Empty, Optional EndDate As Date = Empty)
    Dim TheData As Variant
    Dim TheHeaders As Variant
    Dim wbk As Workbook
    Dim AnAction As New BloombergEquityAlertRow
    Dim BloombergTickerSetString As String
    
    Let TheHeaders = Most(GetTableHeaders("bloombergequityalerts", "equity250v1", DbServerAddress, DbUserName, DbPassword))
    
    If IsMissing(BloombergTickerArray) And AnnouncementDate <> Empty Then
        Let TheData = ConnectAndSelect("SELECT * FROM `equity250v1`.`bloombergequityalerts` WHERE date(`date`) = Date('" & ConvertDateTimeToMySQLFormat(AnnouncementDate) & "')", _
                                        "equity250v1", _
                                        DbServerAddress, _
                                        DbUserName, _
                                        DbPassword)
    ElseIf IsMissing(BloombergTickerArray) And StartDate = Empty And EndDate = Empty Then
        Let TheData = ConnectAndSelect("SELECT * FROM `equity250v1`.`bloombergequityalerts`;", _
                                        "equity250v1", _
                                        DbServerAddress, _
                                        DbUserName, _
                                        DbPassword)
    ElseIf IsMissing(BloombergTickerArray) And StartDate = Empty And EndDate <> Empty Then
        Let TheData = ConnectAndSelect("SELECT * FROM `equity250v1`.`bloombergequityalerts` WHERE date(`enddate`) <= Date('" & ConvertDateTimeToMySQLFormat(EndDate) & "');", _
                                        "equity250v1", _
                                        DbServerAddress, _
                                        DbUserName, _
                                        DbPassword)
    ElseIf IsMissing(BloombergTickerArray) And StartDate <> Empty And EndDate <> Empty Then
        Let TheData = ConnectAndSelect("SELECT * FROM `equity250v1`.`bloombergequityalerts` WHERE date(`startdate`) >= Date('" & ConvertDateTimeToMySQLFormat(StartDate) & "') AND date(`enddate`) <= Date('" & ConvertDateTimeToMySQLFormat(EndDate) & "');", _
                                        "equity250v1", _
                                        DbServerAddress, _
                                        DbUserName, _
                                        DbPassword)
    ElseIf Not IsArray(BloombergTickerArray) And StartDate <> Empty And EndDate = Empty Then
        Let TheData = ConnectAndSelect("SELECT * FROM `equity250v1`.`bloombergequityalerts` WHERE date(`startdate`) >= Date('" & ConvertDateTimeToMySQLFormat(StartDate) & "');", _
                                        "equity250v1", _
                                        DbServerAddress, _
                                        DbUserName, _
                                        DbPassword)
    Else
        Let BloombergTickerSetString = Convert1DArrayIntoParentheticalExpression(AddSingleQuotesToAllArrayElements(BloombergTickerArray))
    
        Let TheData = ConnectAndSelect("SELECT * FROM `equity250v1`.`bloombergequityalerts` WHERE date(`startdate`) >= Date('" & ConvertDateTimeToMySQLFormat(StartDate) & "') AND date(`enddate`) <= Date('" & ConvertDateTimeToMySQLFormat(EndDate) & "') AND `bloombergticker` IN " & BloombergTickerSetString & ";", _
                                        "equity250v1", _
                                        DbServerAddress, _
                                        DbUserName, _
                                        DbPassword)
    End If
    
    If GetNumberOfRows(TheData) > 1 Then
        Let TheData = GetSubMatrix(TheData, 2, GetNumberOfRows(TheData), 1, ArrayFormulas.GetNumberOfColumns(TheData) - 1)
    Else
        Let TheData = Array()
    End If
                                    
    Set wbk = Application.Workbooks.Add
    Call RemoveAllOtherWorksheets(wbk.Worksheets(1))
    
    Call AnAction.InitializeWithRange(wbk.Worksheets(1).Range("A1").Resize(1, AnAction.GetNumberOfColumns))
    
    Call DumpInSheet(TheHeaders, wbk.Worksheets(1).Range("A1"), False)
    Call DumpInSheet(TheData, wbk.Worksheets(1).Range("A2"), True)
    Call wbk.Worksheets(1).Columns(AnAction.GetNumberOfColumns + 1).ClearContents
    
    Call Me.InitializeWithRange(wbk.Worksheets(1).Range("A1").CurrentRegion)
    
    If Me.GetDbLength = 0 Then Exit Sub
    Call Me.FormatUnderlyingRange
End Sub

' This sub moves all alerts to the
' Archive sub-folder of Inbox.Bloomberg ALTR
Public Sub ArchiveAlerts()
    Dim OutlookInboxFolder As Outlook.Folder
    Dim DestinationFolder As Outlook.Folder
    Dim AnItem As Outlook.MailItem
    Dim i As Integer
    
    Set OutlookInboxFolder = Me.GetBlombergAlertsFolder()

    ' Exit if there are not messages in the folder
    If OutlookInboxFolder.Items.Count = 0 Then
        Exit Sub
    End If
    
    Set DestinationFolder = OutlookInboxFolder.Folders.GetFirst
    For i = 1 To OutlookInboxFolder.Items.Count
        Set AnItem = OutlookInboxFolder.Items.GetFirst
        Call AnItem.Move(DestFldr:=DestinationFolder)
    Next i
End Sub

' This returns an instance of this class underlying the subset of in-memory alerts for the given Bloomberg ticker array.
' If StartDate and EndDate are provided, the subset is further filtered by the given dates.
Property Get GetAlertsForGivenBloombergTickerList(BloombergTickerArray As Variant, Optional StartDate As Date = Empty, _
                                                  Optional EndDate As Date = Empty) As BloombergEquityAlerts
    Dim BloombergAlertSet As BloombergEquityAlerts
    Dim SerialStartDate As Long
    Dim SerialEndDate As Long
    
    ' Instantiate return variable
    Set BloombergAlertSet = New BloombergEquityAlerts

    ' Exit if object is empty or the Bloomberg tickers were not passed as an array
    If Not IsArray(BloombergTickerList) Or Me.GetDbLength = 0 Then
        If EmptyArrayQ(BloombergTickerList) Then
            Exit Property
        End If
            
        Set GetAlertsForGivenBloombergTickerList = BloombergAlertSet
    
        Exit Property
    End If
    
    ' Turn off autofilter, if it was on
    Let Me.GetRange.Worksheet.AutoFilterMode = False
    
    ' Do autofilters to select alerts
    Call Me.GetRange.Offset(-1, 0).Resize(Me.GetDbLength + 1, Me.GetNumberOfColumns).AutoFilter(Field:=Me.GetBloombergTickers.Column, _
                                                                                                Criteria1:=BloombergTickerArray, _
                                                                                                Operator:=xlFilterValues)
    
    ' Further restrict the autofilter using a date range if
    If StartDate <> Empty And EndDate <> Empty Then
        Let SerialStartDate = ConvertDateToSerial(StartDate)
        Let SerialEndDate = ConvertDateToSerial(EndDate)
    
        Call Me.GetRange.Offset(-1, 0).Resize(Me.GetDbLength + 1, Me.GetNumberOfColumns).AutoFilter(Field:=Me.GetStartDates.Column, _
                                                                                                    Criteria1:=">=" & SerialStartDate, _
                                                                                                    Operator:=xlFilterValues)
        Call Me.GetRange.Offset(-1, 0).Resize(Me.GetDbLength + 1, Me.GetNumberOfColumns).AutoFilter(Field:=Me.GetEndDates.Column, _
                                                                                                    Criteria1:="<=" & SerialEndDate, _
                                                                                                    Operator:=xlFilterValues)
    End If
    
    ' Copy the filtered set to a new workbook and initialize the Bloomberg alerts instance
    Set BloombergAlertSet.UnderlyingWorbook = Application.Workbooks.Add
    Call RemoveAllOtherWorksheets(BloombergAlertSet.UnderlyingWorbook.Worksheets(1))
    Call Me.GetRange.Worksheet.AutoFilter.Range.SpecialCells(xlCellTypeVisible).Copy
    Call BloombergAlertSet.UnderlyingWorbook.Worksheets(1).Range("A1").PasteSpecial(Paste:=xlPasteAll)
    Call BloombergAlertSet.InitializeWithRange(BloombergAlertSet.UnderlyingWorbook.Worksheets(1).Range("A1").CurrentRegion)
    
    ' Turn off autofilter
    Let Me.GetRange.Worksheet.AutoFilterMode = False
    
    Set GetAlertsForGivenBloombergTickerList = BloombergAlertSet
End Property
