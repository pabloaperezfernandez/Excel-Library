VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BloombergEquityAlertRow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Option Base 1

Public RangeDb As Range

Property Get GetRange() As Range
    Set GetRange = RangeDb
End Property

Public Function GetHeaders() As Variant
    Let GetHeaders = Array("Date", "Reference", "BB Ticker", "Subject", "Body", "Action Type", "Type Qualifier", "XML Body", "Start Date", "End Date", "VerifiedQ", "EqyFundTicker")
End Function

Public Function GetNumberOfColumns() As Integer
    Let GetNumberOfColumns = GetArrayLength(Me.GetHeaders)
End Function

Public Sub InitializeWithRange(ARange As Range)
    Set Me.RangeDb = ARange.Resize(1, Me.GetNumberOfColumns)
End Sub

Property Get GetDate() As Range
    Set GetDate = Me.GetRange.Cells(1)
End Property

Property Get GetReference() As Range
    Set GetReference = Me.GetRange.Cells(2)
End Property

Property Get GetBloombergTicker() As Range
    Set GetBloombergTicker = Me.GetRange.Cells(3)
End Property

Property Get GetSubject() As Range
    Set GetSubject = Me.GetRange.Cells(4)
End Property

Property Get GetBody() As Range
    Set GetBody = Me.GetRange.Cells(5)
End Property

Property Get GetType() As Range
    Set GetType = Me.GetRange.Cells(6)
End Property

Property Get GetQualifier() As Range
    Set GetQualifier = Me.GetRange.Cells(7)
End Property

Property Get GetXmlBody() As Range
    Set GetXmlBody = Me.GetRange.Cells(8)
End Property

Property Get GetStartDate() As Range
    Set GetStartDate = Me.GetRange.Cells(9)
End Property

Property Get GetEndDate() As Range
    Set GetEndDate = Me.GetRange.Cells(10)
End Property

Property Get GetVerifiedQ() As Range
    Set GetVerifiedQ = Me.GetRange.Cells(11)
End Property

Property Get GetEqyFundTicker() As Range
    Set GetEqyFundTicker = Me.GetRange.Cells(12)
End Property

Property Get GetReferenceFromBody() As Long
    Dim TheBody As String
    Dim StartPos As Integer
    Dim EndPos As Integer
    
    Let TheBody = Me.GetBody.Value2
    Let StartPos = InStr(TheBody, "CACX ")
    Let EndPos = InStr(TheBody, "<GO>")
    Let GetReferenceFromBody = CLng(Mid(TheBody, StartPos + 5, EndPos - StartPos - 5))
End Property


' --------------------------------------------------------------------------------------------------
' Setters
' --------------------------------------------------------------------------------------------------
Property Let SetDate(arg As String)
    Let Me.GetDate.Value2 = arg
End Property

Property Let SetReference(arg As Long)
    Let Me.GetReference.Value2 = arg
End Property

Property Let SetBloombergTicker(arg As String)
    Let Me.GetBloombergTicker.Value2 = arg
End Property

Property Let SetSubject(arg As String)
    Let Me.GetSubject.Value2 = arg
End Property

Property Let SetBody(arg As String)
    Let Me.GetBody.Value2 = arg
End Property

Property Let SetXmlBody(arg As String)
    Let Me.GetXmlBody.Value2 = arg
End Property

Property Let SetStartDate(arg As Variant)
    Let Me.GetStartDate.Value2 = arg
End Property

Property Let SetEndDate(arg As Variant)
    Let Me.GetEndDate.Value2 = arg
End Property

Property Let SetVerifiedQ(arg As Boolean)
    Let Me.GetVerifiedQ.Value2 = arg
End Property

Property Let SetEqyFundTicker(arg As String)
    Let Me.GetEqyFundTicker.Value2 = arg
End Property

' This function computes the alert type based on the subject line of the message
' It also determines if the action is a new one or an update
Public Sub ComputeAlertStringType()
    Dim AnAction As New CorporateAction
    Dim TheType As String
    Dim TheQualifier As String
        
    ' Check if alert type is dividend
    If InStr(UCase(Me.GetSubject.Value2), "CASH DIVIDEND") > 0 Then
        ' Insert the string representation of the type in the underlying cell and set the instance variable to
        ' the appropriate enumerated data type
        Let AnAction.SetType = Dividend
        Let TheType = AnAction.GetTypeAsString
        
        ' Determine if this is an update
        Let TheQualifier = IIf(InStr(UCase(Me.GetSubject.Value2), "ANNOUNCE") > 0, "Announcement", "Update")
        
    ElseIf InStr(UCase(Me.GetSubject.Value2), "STOCK SPLIT") > 0 Then
        ' Insert the string representation of the type in the underlying cell and set the instance variable to
        ' the appropriate enumerated data type
        Let AnAction.SetType = StockSplit
        Let TheType = AnAction.GetTypeAsString
        
        ' Determine if this is an update
        Let TheQualifier = IIf(InStr(UCase(Me.GetSubject.Value2), "ANNOUNCE") > 0, "Announcement", "Update")
        
        
    ElseIf InStr(UCase(Me.GetSubject.Value2), "STOCK BUYBACK") > 0 Then
        ' Insert the string representation of the type in the underlying cell and set the instance variable to
        ' the appropriate enumerated data type
        Let AnAction.SetType = StockBuyBack
        Let TheType = AnAction.GetTypeAsString
        
        ' Determine if this is an update
        Let TheQualifier = IIf(InStr(UCase(Me.GetSubject.Value2), "ANNOUNCE") > 0, "Announcement", "Update")
        
        
    ElseIf InStr(UCase(Me.GetSubject.Value2), "ACQUISITION") > 0 Then
        ' Insert the string representation of the type in the underlying cell and set the instance variable to
        ' the appropriate enumerated data type
        Let AnAction.SetType = Acquisition
        Let TheType = AnAction.GetTypeAsString
        
        ' Determine if this is an update
        Let TheQualifier = IIf(InStr(UCase(Me.GetSubject.Value2), "ANNOUNCE") > 0, "Announcement", "Update")
        

    ElseIf InStr(UCase(Me.GetSubject.Value2), "DIVESTITURE") > 0 Then
        ' Insert the string representation of the type in the underlying cell and set the instance variable to
        ' the appropriate enumerated data type
        Let AnAction.SetType = Divestiture
        Let TheType = AnAction.GetTypeAsString
        
        ' Determine if this is an update
        Let TheQualifier = IIf(InStr(UCase(Me.GetSubject.Value2), "ANNOUNCE") > 0, "Announcement", "Update")
        

    ElseIf InStr(UCase(Me.GetSubject.Value2), "EQUITY OFFERING") > 0 Then
        ' Insert the string representation of the type in the underlying cell and set the instance variable to
        ' the appropriate enumerated data type
        Let AnAction.SetType = EquityOffering
        Let TheType = AnAction.GetTypeAsString
        
        ' Determine if this is an update
        Let TheQualifier = IIf(InStr(UCase(Me.GetSubject.Value2), "ANNOUNCE") > 0, "Announcement", "Update")
        
    ElseIf InStr(UCase(Me.GetSubject.Value2), "DELISTING") > 0 Then
        ' Insert the string representation of the type in the underlying cell and set the instance variable to
        ' the appropriate enumerated data type
        Let AnAction.SetType = EquityOffering
        Let TheType = AnAction.GetTypeAsString
        
        ' Determine if this is an update
        Let TheQualifier = IIf(InStr(UCase(Me.GetSubject.Value2), "ANNOUNCE") > 0, "Announcement", "Update")
    
    ElseIf InStr(UCase(Me.GetSubject.Value2), "ID NUMBER CHANGE") > 0 Then
        ' Insert the string representation of the type in the underlying cell and set the instance variable to
        ' the appropriate enumerated data type
        Let AnAction.SetType = IdNumberChange
        Let TheType = AnAction.GetTypeAsString
        
        ' Determine if this is an update
        Let TheQualifier = IIf(InStr(UCase(Me.GetSubject.Value2), "ANNOUNCE") > 0, "Announcement", "Update")
        
    ElseIf InStr(UCase(Me.GetSubject.Value2), "NAME CHANGE") > 0 Then
        ' Insert the string representation of the type in the underlying cell and set the instance variable to
        ' the appropriate enumerated data type
        Let AnAction.SetType = NameChange
        Let TheType = AnAction.GetTypeAsString
        
        ' Determine if this is an update
        Let TheQualifier = IIf(InStr(UCase(Me.GetSubject.Value2), "ANNOUNCE") > 0, "Announcement", "Update")

    ElseIf InStr(UCase(Me.GetSubject.Value2), "TICKER SYMBOL CHANGE") > 0 Then
        ' Insert the string representation of the type in the underlying cell and set the instance variable to
        ' the appropriate enumerated data type
        Let AnAction.SetType = TickerSymbolChange
        Let TheType = AnAction.GetTypeAsString
        
        ' Determine if this is an update
        Let TheQualifier = IIf(InStr(UCase(Me.GetSubject.Value2), "ANNOUNCE") > 0, "Announcement", "Update")
        
    ElseIf InStr(UCase(Me.GetSubject.Value2), "DEBT OFFERING") > 0 Then
        ' Insert the string representation of the type in the underlying cell and set the instance variable to
        ' the appropriate enumerated data type
        Let AnAction.SetType = TickerSymbolChange
        Let TheType = AnAction.GetTypeAsString
        
        ' Determine if this is an update
        Let TheQualifier = IIf(InStr(UCase(Me.GetSubject.Value2), "ANNOUNCE") > 0, "Announcement", "Update")

    ElseIf InStr(UCase(Me.GetSubject.Value2), "RIGHTS ISSUE") > 0 Then
        ' Insert the string representation of the type in the underlying cell and set the instance variable to
        ' the appropriate enumerated data type
        Let AnAction.SetType = TickerSymbolChange
        Let TheType = AnAction.GetTypeAsString
        
        ' Determine if this is an update
        Let TheQualifier = IIf(InStr(UCase(Me.GetSubject.Value2), "Chg") > 0, "Update", "Announcement")

    ElseIf InStr(UCase(Me.GetSubject.Value2), "STOCK DIVIDEND") > 0 Then
        ' Insert the string representation of the type in the underlying cell and set the instance variable to
        ' the appropriate enumerated data type
        Let AnAction.SetType = StockDividend
        Let TheType = AnAction.GetTypeAsString
        
        ' Determine if this is an update
        Let TheQualifier = IIf(InStr(UCase(Me.GetSubject.Value2), "ANNOUNCE") > 0, "Announcement", "Update")

    ElseIf InStr(UCase(Me.GetSubject.Value2), "DEBT REDEMPTION") > 0 Then
        ' Insert the string representation of the type in the underlying cell and set the instance variable to
        ' the appropriate enumerated data type
        Let AnAction.SetType = StockDividend
        Let TheType = AnAction.GetTypeAsString
        
        ' Determine if this is an update
        Let TheQualifier = IIf(InStr(UCase(Me.GetSubject.Value2), "ANNOUNCE") > 0, "Announcement", "Update")

    Else
    
        Let TheType = "Other"
        Let TheQualifier = "Announcement"
    End If
    
    ' Insert the action type and its qualifier in the underlying cells
    Let Me.GetType.Value2 = TheType
    Let Me.GetQualifier.Value2 = TheQualifier
End Sub

' This function computes the start and end date during which the alert is relevant
' All info here is dependent on the alert type and extracted from the XML Body
Public Sub ComputeLifeCycleDates()
    Dim AnAction As New CorporateAction
    Dim xDoc As MSXML2.DOMDocument60
    Dim StartDate As Variant
    Dim EndDate As Variant

    ' Instantiate a DOMDocument object
    Set xDoc = New MSXML2.DOMDocument60
    
    ' Initialize the DOMDocument with the XML data
    Call xDoc.LoadXML(Me.GetXmlBody.Value2)
        
    ' Initialize AnAction
    Let AnAction.SetTypeFromString = Me.GetType.Value2
    
    ' Process each action's dates differently
    Select Case AnAction.GetType
        Case Dividend
            If Me.GetQualifier.Value2 = "Announcement" Then
                Let StartDate = GetXmlBodyElementPayload("ANNOUNCEDDATE")
                Let EndDate = IIf(GetXmlBodyElementPayload("PAYDATE") <> "", _
                                  GetXmlBodyElementPayload("PAYDATE"), _
                                  GetXmlBodyElementPayload("RECORDDATE") _
                                 )
            Else
                Let StartDate = GetXmlBodyElementPayload("ANNOUNCEDDATE")
                
                If InStr(Me.GetSubject.Value2, "Date Chg") > 0 Then
                    Let EndDate = IIf(GetXmlBodyElementPayload("NEWPAYDATE") <> "", _
                                      GetXmlBodyElementPayload("NEWPAYDATE"), _
                                      GetXmlBodyElementPayload("NEWRECORDDATE") _
                                     )
                ElseIf InStr(Me.GetSubject.Value2, "Net Amount Chg") > 0 Then
                    Let EndDate = IIf(GetXmlBodyElementPayload("PAYDATE") <> "", _
                                      GetXmlBodyElementPayload("PAYDATE"), _
                                      GetXmlBodyElementPayload("RECORDDATE") _
                                     )
                ElseIf InStr(Me.GetSubject.Value2, "Ex Date Chg") > 0 Then
                    Let EndDate = IIf(GetXmlBodyElementPayload("PAYDATE") <> "", _
                                      GetXmlBodyElementPayload("PAYDATE"), _
                                      GetXmlBodyElementPayload("RECORDDATE") _
                                     )
                Else
                    Let EndDate = GetXmlBodyElementPayload("RECORDDATE")
                End If
            End If
    
        Case StockSplit
            If Me.GetQualifier.Value2 = "Announcement" Then
                Let StartDate = GetXmlBodyElementPayload("ANNOUNCEDDATE")
                Let EndDate = GetXmlBodyElementPayload("RECORDDATE")
            End If

        Case StockBuyBack
            If Me.GetQualifier.Value2 = "Announcement" Then
                Let StartDate = GetXmlBodyElementPayload("ANNOUNCEDDATE")
                Let EndDate = GetXmlBodyElementPayload("EXPIRATIONDATE")
            End If
    
        Case EquityOffering
            If Me.GetQualifier.Value2 = "Announcement" Then
                Let StartDate = GetXmlBodyElementPayload("ANNOUNCEDDATE")
                Let EndDate = GetXmlBodyElementPayload("EFFECTIVEDATE")
            End If

        Case Divestiture
            If Me.GetQualifier.Value2 = "Announcement" Then
                Let StartDate = GetXmlBodyElementPayload("ANNOUNCEDDATE")
                Let EndDate = GetXmlBodyElementPayload("EFFECTIVEDATE")
            Else
                Let StartDate = GetXmlBodyElementPayload("ANNOUNCEDDATE")
                Let EndDate = GetXmlBodyElementPayload("NEWCOMPLETIONDATE")
            End If
            
        Case Acquisition
            If Me.GetQualifier.Value2 = "Announcement" Then
                Let StartDate = GetXmlBodyElementPayload("ANNOUNCEDDATE")
                Let EndDate = GetXmlBodyElementPayload("COMPLETIONDATE")
            Else
                Let StartDate = GetXmlBodyElementPayload("ANNOUNCEDDATE")
                Let EndDate = GetXmlBodyElementPayload("NEWCOMPLETIONDATE")
            End If
            
        Case Delisting
            Let StartDate = GetXmlBodyElementPayload("ANNOUNCEDDATE")
            Let EndDate = GetXmlBodyElementPayload("EFFECTIVEDATE")

        Case NameChange
            If Me.GetQualifier.Value2 = "Announcement" Then
                Let StartDate = GetXmlBodyElementPayload("ANNOUNCEDDATE")
                Let EndDate = GetXmlBodyElementPayload("EFFECTIVEDATE")
            End If

        Case TickerSymbolChange
            If Me.GetQualifier.Value2 = "Announcement" Then
                Let StartDate = GetXmlBodyElementPayload("ANNOUNCEDDATE")
                Let EndDate = GetXmlBodyElementPayload("EFFECTIVEDATE")
            End If

        Case IdNumberChange
            If Me.GetQualifier.Value2 = "Announcement" Then
                Let StartDate = GetXmlBodyElementPayload("ANNOUNCEDDATE")
                Let EndDate = GetXmlBodyElementPayload("EFFECTIVEDATE")
            End If
            
        Case DebtOffering
            If Me.GetQualifier.Value2 = "Announcement" Then
                Let StartDate = GetXmlBodyElementPayload("ANNOUNCEDDATE")
                Let EndDate = GetXmlBodyElementPayload("EFFECTIVEDATE")
            End If

        Case RightsIssue
            If Me.GetQualifier.Value2 = "Announcement" Then
                Let StartDate = GetXmlBodyElementPayload("ANNOUNCEDDATE")
                Let EndDate = GetXmlBodyElementPayload("RECORDDATE")
            End If
            
        Case StockDividend
            If Me.GetQualifier.Value2 = "Announcement" Then
                Let StartDate = GetXmlBodyElementPayload("ANNOUNCEDDATE")
                Let EndDate = GetXmlBodyElementPayload("RECORDDATE")
            End If

        Case DebtRedemption
            If Me.GetQualifier.Value2 = "Announcement" Then
                Let StartDate = GetXmlBodyElementPayload("ANNOUNCEDDATE")
                Let EndDate = GetXmlBodyElementPayload("RECORDDATE")
            End If

        Case Else
            Let StartDate = GetXmlBodyElementPayload("ANNOUNCEDDATE")
            Let EndDate = ""
    End Select
    
    ' Insert the dates in the underlying range
    If StartDate = "" Then
        Let Me.SetStartDate = Empty
    ElseIf IsDate(StartDate) Then
        Let Me.SetStartDate = ConvertDateToSerial(CDate(StartDate))
    Else
        Let Me.SetStartDate = Empty
    End If
    
    If EndDate = "" Then
        Let Me.SetEndDate = Empty
    ElseIf IsDate(EndDate) Then
        Let Me.SetEndDate = ConvertDateToSerial(CDate(EndDate))
    Else
        Let Me.SetEndDate = Empty
    End If
End Sub

' Helper function for Sub ComputeLifeCycleDates
' If an element with the given tag is found, its payload is return.
' Otherwise, this function returns ""
Private Function GetXmlBodyElementPayload(TheTag As String) As String
    Dim xDoc As MSXML2.DOMDocument60
    Dim CorrectedTag As String
    
    Let CorrectedTag = Replace(Trim(TheTag), " ", "")

    ' Instantiate a DOMDocument object
    Set xDoc = New MSXML2.DOMDocument60
    
    ' Initialize the DOMDocument with the XML data
    Call xDoc.LoadXML(Me.GetXmlBody.Value2)

    If xDoc.getElementsByTagName(CorrectedTag).Length > 0 Then
        Let GetXmlBodyElementPayload = xDoc.getElementsByTagName(CorrectedTag).Item(0).nodeTypedValue
    Else
        Let GetXmlBodyElementPayload = ""
    End If
End Function
