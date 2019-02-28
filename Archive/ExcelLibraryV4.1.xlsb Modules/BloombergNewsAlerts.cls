VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BloombergNewsAlerts"
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
    Dim AnAlert As New BloombergNewsAlertRow

    Let GetHeaders = AnAlert.GetHeaders
End Function

Public Function GetNumberOfColumns() As Integer
    Dim AnAlert As New BloombergNewsAlertRow

    Let GetNumberOfColumns = AnAlert.GetNumberOfColumns
End Function

Property Get GetDates() As Range
    Dim AnAlert As New BloombergNewsAlertRow
    
    Call AnAlert.InitializeWithRange(Me.GetRange.Rows(1))
    
    Set GetDates = Me.GetRange.Columns(AnAlert.GetDate.Column)
End Property

Property Get GetBloombergTickers() As Range
    Dim AnAlert As New BloombergNewsAlertRow
    
    Call AnAlert.InitializeWithRange(Me.GetRange.Rows(1))
    
    Set GetBloombergTickers = Me.GetRange.Columns(AnAlert.GetBloombergTicker.Column)
End Property

Property Get GetBodies() As Range
    Dim AnAlert As New BloombergNewsAlertRow
    
    Call AnAlert.InitializeWithRange(Me.GetRange.Rows(1))
    
    Set GetBodies = Me.GetRange.Columns(AnAlert.GetBody.Column)
End Property

Property Get GetEqyFundTickers() As Range
    Dim AnAlert As New BloombergNewsAlertRow
    
    Call AnAlert.InitializeWithRange(Me.GetRange.Rows(1))
    
    Set GetEqyFundTickers = Me.GetRange.Columns(AnAlert.GetEqyFundTicker.Column)
End Property

Property Get GetAlertTypes() As Range
    Dim AnAlert As New BloombergNewsAlertRow
    
    Call AnAlert.InitializeWithRange(Me.GetRange.Rows(1))
    
    Set GetAlertTypes = Me.GetRange.Columns(AnAlert.GetAlertType.Column)
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

Public Sub InitializeWithOutlook()
    Dim AnAlert As New BloombergNewsAlertRow
    Dim OutlookInboxFolder As Outlook.Folder
    Dim TheNewsAlertsDict As New Dictionary
    Dim i As Long
    Dim wsht As Worksheet
    Dim DatePart As String
    Dim TimePart As String
    Dim AmVsPm As String
    Dim SummariesArray As Variant
    Dim StoriesByTickerArray As Variant
    Dim var As Variant
    Dim TheDate As Date
    Dim TheBody As String
    Dim TempRange As Range

    Set OutlookInboxFolder = GetBlombergAlertsFolder()

    ' Exit if there are not messages in the folder
    If OutlookInboxFolder.Items.Count = 0 Then
        Exit Sub
    End If
    
    ' Instantiate a workbook to hold the
    Set wsht = Application.Workbooks.Add.Worksheets(1)
    Call RemoveAllOtherWorksheets(wsht)
    
    ' Point AnAlert to a random location so we may use some its column position properties
    Call AnAlert.InitializeWithRange(wsht.Range("A1"))

    ' Process each email in the Outlook folder.  Each email has multiple alerts
    For i = 1 To GetBlombergAlertsFolder.Items.Count
        ' Parse date/time information
        Let TheDate = OutlookInboxFolder.Items(i).SentOn
        Let DatePart = CStr(ConvertDateToSerial(CDate(Split(Trim(TheDate), " ")(0))))
        Let TimePart = Split(Trim(TheDate), " ")(1)
        Let AmVsPm = Right(Trim(TheDate), 2)
        Let TheDate = Left(DatePart, 4) & "-" & Mid(DatePart, 5, 2) & "-" & Right(DatePart, 2) & " " & TimePart & " " & AmVsPm
    
        Let var = Split(OutlookInboxFolder.Items(i).Body, vbCrLf & "Stories By Security" & vbCrLf)
        Let SummariesArray = Rest(Split(var(0), vbCrLf & vbCrLf))
        If GetArrayLength(var) > 1 Then
            Let StoriesByTickerArray = Most(Split(var(1), vbCrLf & vbCrLf))
        Else
            Let StoriesByTickerArray = Array()
        End If
        
        For Each var In SummariesArray
            Let TheBody = Trim(Replace(CStr(var), vbCrLf, " "))
            Let TheBody = Right(TheBody, Len(TheBody) - 3)
            Call TheNewsAlertsDict.Add(Key:=TheNewsAlertsDict.Count, _
                                       Item:=Array(TheDate, Empty, TheBody, Empty, "Mention"))
        Next
        
        For Each var In StoriesByTickerArray
            Let TheBody = Trim(CStr(var))
            Call TheNewsAlertsDict.Add(Key:=TheNewsAlertsDict.Count, _
                                       Item:=Array(TheDate, GetTickerFromBodyString(CStr(var)), TheBody, Empty, "Specific"))
        Next
    Next i
    
    ' Dump news alerts into a range
    Set TempRange = DumpInSheet(Me.GetHeaders, wsht.Range("A1"))
    For Each var In TheNewsAlertsDict.Items
        Set TempRange = DumpInSheet(var, TempRange.Cells(1, 1).Offset(1, 0))
    Next

    ' Initialize with range
    Call Me.InitializeWithRange(wsht.Range("A1").CurrentRegion)
    
    ' Post-process the EQY_FUND_COLUMN.  Pull it from `equity250v1`.`bloomberg-eqs` if available
    ' Return the list of tickers with no EQY_FUND_TICKER as a 1D array so the user may pull them directly from
    ' Bloomberg '***HERE

    Call FormatUnderlyingRange
End Sub

' Format underlying range appropriately
Private Sub FormatUnderlyingRange()
    Let Me.GetDates.EntireColumn.NumberFormat = "yyyy-mm-dd hh:mm:ss"
    Call Me.GetDates.EntireColumn.AutoFit
    Let Me.GetBloombergTickers.EntireColumn.HorizontalAlignment = xlCenter
    Call Me.GetBloombergTickers.EntireColumn.AutoFit
    Let Me.GetBodies.EntireColumn.ColumnWidth = 71
    Let Me.GetBodies.EntireColumn.WrapText = True
    Call Me.GetEqyFundTickers.EntireColumn.AutoFit
    Let Me.GetAlertTypes.EntireColumn.HorizontalAlignment = xlCenter
    Call Me.GetAlertTypes.EntireColumn.AutoFit
    Let Me.GetRange.VerticalAlignment = xlTop
    Let Me.GetRange.Rows(1).Offset(-1, 0).Font.Bold = True
    Call Me.GetRange.Range("A1").Select
    Let Me.GetRange.Worksheet.Parent.Windows(1).FreezePanes = True
End Sub

' Helper function for Public Function GetBloombergAlerts() in this module
Private Function GetBlombergAlertsFolder() As Outlook.Folder
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
        If UCase(OutlookFolderSet.Item(i).Name) = UCase("Bloomberg News") Then
            Set GetBlombergAlertsFolder = OutlookFolderSet.Item(i)
        End If
    Next i
End Function

Property Get GetIthAlert(i As Long) As BloombergNewsAlertRow
    Dim AnAlert As BloombergNewsAlertRow
    
    Set AnAlert = New BloombergNewsAlertRow
    
    Call AnAlert.InitializeWithRange(Me.GetRange.Rows(i))
    
    Set GetIthAlert = AnAlert
End Property

' This helper function returns the bloomberg ticker in the given string (the body of the email containing the BB alert)
Private Function GetTickerFromBodyString(BodyString As String) As String
    Dim StartPos As Integer
    Dim EndPos As Integer
    
    Let StartPos = InStr(Trim(BodyString), "(")
    Let EndPos = InStr(Trim(BodyString), ")")
    
    If EndPos <= StartPos Then
        Let GetTickerFromBodyString = ""
        Exit Function
    End If
    
    Let GetTickerFromBodyString = Mid(BodyString, StartPos + 1, EndPos - StartPos - 1)
End Function

Public Sub InjectIntoMySql()
    Dim TheHeaders As Variant
    Dim TheDateStrings() As String
    Dim OriginalDates As Variant
    Dim i As Long
    
    Let TheHeaders = GetTableHeaders("bloombergnewsalerts", "equity250v1", DbServerAddress, DbUserName, DbPassword)
                                      
    Let OriginalDates = Me.GetDates.Value2
    ReDim TheDateStrings(Me.GetDbLength)
    For i = 1 To Me.GetDbLength
        Let TheDateStrings(i) = "'" & ConvertDateTimeToMySQLFormat(Me.GetIthAlert(i).GetDate.Text)
    Next i
    
    Let Me.GetDates.NumberFormat = "@"
    Let Me.GetDates.Value2 = Application.Transpose(TheDateStrings)
    
    Call InjectMatrixIntoMySql(Me.GetRange.Value2, TheHeaders, "bloombergnewsalerts", DbServerAddress, "equity250v1", DbUserName, DbPassword)
    
    Let Me.GetDates.Value2 = OriginalDates
    Call FormatUnderlyingRange
End Sub

Public Sub InitializeWithMySql(Optional BloombergTickerArray As Variant, Optional StartDate As Date = Empty, Optional EndDate As Date = Empty)
    Dim TheData As Variant
    Dim TheHeaders As Variant
    Dim wbk As Workbook
    Dim AnAction As New BloombergNewsAlertRow
    Dim BloombergTickerSetString As String
    
    Let TheHeaders = GetTableHeaders("bloombergnewsalerts", "equity250v1", DbServerAddress, DbUserName, DbPassword)
    
    If IsMissing(BloombergTickerArray) And StartDate = Empty And EndDate = Empty Then
        Let TheData = ConnectAndSelect("SELECT * FROM `equity250v1`.`bloombergnewsalerts`;", _
                                        "equity250v1", _
                                        DbServerAddress, _
                                        DbUserName, _
                                        DbPassword)
    ElseIf IsMissing(BloombergTickerArray) And StartDate <> Empty And EndDate = Empty Then
        Let TheData = ConnectAndSelect("SELECT * FROM `equity250v1`.`bloombergnewsalerts` WHERE date(`datetime`) >= Date('" & ConvertDateTimeToMySQLFormat(StartDate) & "');", _
                                        "equity250v1", _
                                        DbServerAddress, _
                                        DbUserName, _
                                        DbPassword)
    ElseIf IsMissing(BloombergTickerArray) And StartDate = Empty And EndDate <> Empty Then
        Let TheData = ConnectAndSelect("SELECT * FROM `equity250v1`.`bloombergnewsalerts` WHERE date(`datetime`) <= Date('" & ConvertDateTimeToMySQLFormat(EndDate) & "');", _
                                        "equity250v1", _
                                        DbServerAddress, _
                                        DbUserName, _
                                        DbPassword)
    ElseIf Not IsArray(BloombergTickerArray) And StartDate <> Empty And EndDate <> Empty Then
        Let TheData = ConnectAndSelect("SELECT * FROM `equity250v1`.`bloombergnewsalerts` WHERE date(`datetime`) >= Date('" & ConvertDateTimeToMySQLFormat(StartDate) & "') AND date(`datetime`) <= Date('" & ConvertDateTimeToMySQLFormat(EndDate) & "');", _
                                        "equity250v1", _
                                        DbServerAddress, _
                                        DbUserName, _
                                        DbPassword)
    Else
        Let BloombergTickerSetString = Convert1DArrayIntoParentheticalExpression(AddSingleQuotesToAllArrayElements(BloombergTickerArray))
    
        Let TheData = ConnectAndSelect("SELECT * FROM `equity250v1`.`bloombergnewsalerts` WHERE date(`datetime`) >= Date('" & ConvertDateTimeToMySQLFormat(StartDate) & "') AND date(`datetime`) <= Date('" & ConvertDateTimeToMySQLFormat(EndDate) & "') AND `bloombergticker` IN " & BloombergTickerSetString & ";", _
                                        "equity250v1", _
                                        DbServerAddress, _
                                        DbUserName, _
                                        DbPassword)
    End If
    Let TheData = GetSubMatrix(TheData, 2, GetNumberOfRows(TheData), 1, ArrayFormulas.GetNumberOfColumns(TheData))
                                    
    Set wbk = Application.Workbooks.Add
    Call RemoveAllOtherWorksheets(wbk.Worksheets(1))
    
    Call AnAction.InitializeWithRange(wbk.Worksheets(1).Range("A1").Resize(1, AnAction.GetNumberOfColumns))
    
    Call DumpInSheet(TheHeaders, wbk.Worksheets(1).Range("A1"), False)
    Call DumpInSheet(TheData, wbk.Worksheets(1).Range("A2"), True)
    Call wbk.Worksheets(1).Columns(AnAction.GetNumberOfColumns + 1).ClearContents
    
    Call Me.InitializeWithRange(wbk.Worksheets(1).Range("A1").CurrentRegion)
    Call FormatUnderlyingRange
End Sub

' This sub moves all alerts to the
' Archive sub-folder of Inbox.Bloomberg ALTR
Public Sub ArchiveAlerts()
    Dim OutlookInboxFolder As Outlook.Folder
    Dim DestinationFolder As Outlook.Folder
    Dim AnItem As Outlook.MailItem
    Dim i As Integer
    
    Set OutlookInboxFolder = GetBlombergAlertsFolder()

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
