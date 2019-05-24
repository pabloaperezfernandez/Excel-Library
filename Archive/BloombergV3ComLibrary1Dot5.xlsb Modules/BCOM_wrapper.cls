VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BCOM_wrapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'
' public enumerator for request type
Public Enum ENUM_REQUEST_TYPE
    REFERENCE_DATA = 1
    HISTORICAL_DATA = 2
    BULK_REFERENCE_DATA = 3
End Enum
'
' constants
Private Const CONST_SERVICE_TYPE As String = "//blp/refdata"
Private Const CONST_REQUEST_TYPE_REFERENCE As String = "ReferenceDataRequest"
Private Const CONST_REQUEST_TYPE_BULK_REFERENCE As String = "ReferenceDataRequest"
Private Const CONST_REQUEST_TYPE_HISTORICAL As String = "HistoricalDataRequest"
'
' private data structures
Private bInputSecurityArray() As String
Private bInputFieldArray() As String
Private bOutputArray As Variant
Private bOverrideFieldArray() As String
Private bOverrideValueArray() As String
'
' BCOM objects
Private bSession As blpapicomLib2.session
Private bService As blpapicomLib2.Service
Private bRequest As blpapicomLib2.REQUEST

Private bSecurityArray As blpapicomLib2.Element
Private bFieldArray As blpapicomLib2.Element
Private bEvent As blpapicomLib2.Event

Private bIterator As blpapicomLib2.MessageIterator
Private bIteratorData As blpapicomLib2.Message
Private bSecurities As blpapicomLib2.Element

Private bSecurity As blpapicomLib2.Element
Private bSecurityName As blpapicomLib2.Element
Private bSecurityField As blpapicomLib2.Element
Private bFieldValue As blpapicomLib2.Element

Private bSequenceNumber As blpapicomLib2.Element

Private bFields As blpapicomLib2.Element
Private bField As blpapicomLib2.Element

Private bDataPoint As blpapicomLib2.Element

Private bOverrides As blpapicomLib2.Element
Private bOverrideArray() As blpapicomLib2.Element
'
' class non-object data members
Private bStartDate As String
Private bEndDate As String
Private bRequestType As ENUM_REQUEST_TYPE
Private nSecurities As Long
Private nSecurity As Long
Private bCalendarCodeOverride As String
Private bCurrencyCode As String
Private bNonTradingDayFillOption As String
Private bNonTradingDayFillMethod As String
Private bPeriodicityAdjustment As String
Private bPeriodicitySelection As String
Private bMaxDataPoints As Integer
Private bPricingOption As String
Private bAdjustmentFollowDPDF As Boolean
Private bAdjustmentNormal As Boolean
Private bAdjustmentAbnormal As Boolean
Private badjustmentSplit As Boolean
'
Public Function referenceData(ByRef Securities As Variant, _
                              ByRef Fields As Variant, _
                              Optional ByRef OverrideFields As Variant, _
                              Optional ByRef OverrideValues As Variant) As Variant
    ' mandatory user input parameters
    Let bRequestType = REFERENCE_DATA
    Let bInputSecurityArray = Securities
    Let bInputFieldArray = Fields

    ' field names and values for overrides
    If Not (VBA.IsMissing(OverrideFields)) Then bOverrideFieldArray = OverrideFields
    If Not (VBA.IsMissing(OverrideValues)) Then bOverrideValueArray = OverrideValues

    Call processDataRequest
    Let referenceData = bOutputArray
End Function
'
Public Function bulkReferenceData(ByRef Securities As Variant, _
                                  ByRef Fields As Variant, _
                                  Optional ByRef OverrideFields As Variant, _
                                  Optional ByRef OverrideValues As Variant) As Variant
    ' mandatory user input parameters
    Let bRequestType = BULK_REFERENCE_DATA
    Let bInputSecurityArray = Securities
    Let bInputFieldArray = Fields

    ' field names and values for overrides
    If Not (VBA.IsMissing(OverrideFields)) Then bOverrideFieldArray = OverrideFields
    If Not (VBA.IsMissing(OverrideValues)) Then bOverrideValueArray = OverrideValues

    Call processDataRequest
    Let bulkReferenceData = bOutputArray
End Function

Public Function historicalData(ByRef Securities As Variant, _
                               ByRef Fields As Variant, _
                               ByVal StartDate As Date, _
                               ByVal EndDate As Date, _
                               Optional ByVal calendarCodeOverride As String, _
                               Optional ByVal currencyCode As String, _
                               Optional ByVal nonTradingDayFillOption As String, _
                               Optional ByVal nonTradingDayFillMethod As String, _
                               Optional ByVal periodicityAdjustment As String, _
                               Optional ByVal periodicitySelection As String, _
                               Optional ByVal maxDataPoints As Integer, _
                               Optional ByVal pricingOption As String, _
                               Optional ByVal adjustmentFollowDPDF As Boolean = True, _
                               Optional ByVal adjustmentNormal As Boolean = True, _
                               Optional ByVal adjustmentAbnormal As Boolean = True, _
                               Optional ByVal adjustmentSplit As Boolean = True, _
                               Optional ByRef OverrideFields As Variant, _
                               Optional ByRef OverrideValues As Variant) As Variant
    ' mandatory user input parameters
    Let bRequestType = HISTORICAL_DATA
    Let bInputSecurityArray = Securities
    Let bInputFieldArray = Fields
    Let bStartDate = StartDate
    Let bEndDate = EndDate

    ' checks and conversions for user-defined dates
    If StartDate = CDate(0) Or EndDate = CDate(0) Then
        Call Err.Raise(vbObjectError, "Bloomberg API", "Date parameters missing for historical data query")
    End If
    
    If StartDate > EndDate Then
        Call Err.Raise(vbObjectError, "Bloomberg API", "Incorrect date parameters for historical data query")
    End If
    
    Let bStartDate = Format(StartDate, "YYYYMMDD")
    Let bEndDate = Format(EndDate, "YYYYMMDD")

    ' optional user input parameters
    Let bCalendarCodeOverride = calendarCodeOverride
    Let bCurrencyCode = currencyCode
    Let bNonTradingDayFillOption = nonTradingDayFillOption
    Let bNonTradingDayFillMethod = nonTradingDayFillMethod
    Let bPeriodicityAdjustment = periodicityAdjustment
    Let bPeriodicitySelection = periodicitySelection
    Let bMaxDataPoints = maxDataPoints
    Let bPricingOption = pricingOption
    Let bAdjustmentFollowDPDF = adjustmentFollowDPDF
    Let bAdjustmentNormal = adjustmentNormal
    Let bAdjustmentAbnormal = adjustmentAbnormal
    Let badjustmentSplit = adjustmentSplit

    ' field names and values for overrides
    If Not (VBA.IsMissing(OverrideFields)) Then Let bOverrideFieldArray = OverrideFields
    If Not (VBA.IsMissing(OverrideValues)) Then Let bOverrideValueArray = OverrideValues
    
    Call processDataRequest
    Let historicalData = bOutputArray
End Function

Private Function processDataRequest()
    Call openSession
    Call sendRequest
    Call catchServerEvent
    Call releaseObjects
End Function

' STEP 1
Private Function openSession()
    Set bSession = New blpapicomLib2.session
    Call bSession.Start
    Call bSession.OpenService(CONST_SERVICE_TYPE)
    Set bService = bSession.GetService(CONST_SERVICE_TYPE)
End Function

' STEP 2
Private Function sendRequest()
    Select Case bRequestType
        Case ENUM_REQUEST_TYPE.HISTORICAL_DATA
            ReDim bOutputArray(0 To UBound(bInputSecurityArray, 1), 0 To 0)
            Set bRequest = bService.CreateRequest(CONST_REQUEST_TYPE_HISTORICAL)
            '
            ' set mandatory user input parameter
            bRequest.Set "startDate", bStartDate
            bRequest.Set "endDate", bEndDate
            '
            ' set optional user input parameter
            If (bNonTradingDayFillOption <> "") Then bRequest.Set "nonTradingDayFillOption", bNonTradingDayFillOption
            If (bNonTradingDayFillMethod <> "") Then bRequest.Set "nonTradingDayFillMethod", bNonTradingDayFillMethod
            If (bPeriodicityAdjustment <> "") Then bRequest.Set "periodicityAdjustment", bPeriodicityAdjustment
            If (bPeriodicitySelection <> "") Then bRequest.Set "periodicitySelection", bPeriodicitySelection
            If (bCalendarCodeOverride <> "") Then bRequest.Set "calendarCodeOverride", bCalendarCodeOverride
            If (bCurrencyCode <> "") Then bRequest.Set "currency", bCurrencyCode
            If (bMaxDataPoints <> 0) Then bRequest.Set "maxDataPoints", bMaxDataPoints
            If (bPricingOption <> "") Then bRequest.Set "pricingOption", bPricingOption
            bRequest.Set "adjustmentFollowDPDF", bAdjustmentFollowDPDF
            bRequest.Set "adjustmentNormal", bAdjustmentNormal
            bRequest.Set "adjustmentAbnormal", bAdjustmentAbnormal
            bRequest.Set "adjustmentSplit", badjustmentSplit
            '
        Case ENUM_REQUEST_TYPE.REFERENCE_DATA
            Dim nSecurities As Long: nSecurities = UBound(bInputSecurityArray)
            Dim nFields As Long: nFields = UBound(bInputFieldArray)
            ReDim bOutputArray(0 To nSecurities, 0 To nFields)
            Set bRequest = bService.CreateRequest(CONST_REQUEST_TYPE_REFERENCE)
            '
        Case ENUM_REQUEST_TYPE.BULK_REFERENCE_DATA
            ReDim bOutputArray(0 To UBound(bInputSecurityArray, 1), 0 To 0)
            Set bRequest = bService.CreateRequest(CONST_REQUEST_TYPE_BULK_REFERENCE)
            '
    End Select
    '
    Set bSecurityArray = bRequest.GetElement("securities")
    Set bFieldArray = bRequest.GetElement("fields")
    Call appendRequestItems
    Call setOverrides
    Call bSession.sendRequest(bRequest)
End Function
'
Private Function setOverrides()
    '
    On Error GoTo errorHandler
    '
    If (UBound(bOverrideFieldArray) <> UBound(bOverrideValueArray)) Then Exit Function
    Set bOverrides = bRequest.GetElement("overrides")
    '
    ReDim bOverrideArray(LBound(bOverrideFieldArray) To UBound(bOverrideFieldArray))
    Dim i As Integer
    For i = 0 To UBound(bOverrideFieldArray)
        '
        If ((Len(bOverrideFieldArray(i)) > 0) And (Len(bOverrideValueArray(i)) > 0)) Then
            '
            Set bOverrideArray(i) = bOverrides.AppendElment()
            bOverrideArray(i).SetElement "fieldId", bOverrideFieldArray(i)
            bOverrideArray(i).SetElement "value", bOverrideValueArray(i)
        End If
    Next i
    Exit Function
    '
errorHandler:
    Exit Function
End Function
'
Private Function appendRequestItems()
    '
    Dim nSecurities As Long: nSecurities = UBound(bInputSecurityArray)
    Dim nFields As Long: nFields = UBound(bInputFieldArray)
    Dim i As Long
    Dim nItems As Integer: nItems = getMax(nSecurities, nFields)
    For i = 0 To nItems
        If (i <= nSecurities) Then Call bSecurityArray.AppendValue(bInputSecurityArray(i))
        If (i <= nFields) Then Call bFieldArray.AppendValue(bInputFieldArray(i))
    Next i
End Function

' STEP 3
Private Function catchServerEvent()
    '
    Dim bExit As Boolean
    Do While (bExit = False)
        Set bEvent = bSession.NextEvent
        If (bEvent.EventType = PARTIAL_RESPONSE Or bEvent.EventType = RESPONSE) Then
            '
            Select Case bRequestType
                Case ENUM_REQUEST_TYPE.REFERENCE_DATA: getServerData_reference
                Case ENUM_REQUEST_TYPE.HISTORICAL_DATA: getServerData_historical
                Case ENUM_REQUEST_TYPE.BULK_REFERENCE_DATA: getServerData_bulkReference
            End Select
            '
            If (bEvent.EventType = RESPONSE) Then bExit = True
        End If
    Loop
End Function

' STEP 4
Private Function releaseObjects()
    Let nSecurity = 0
    Set bDataPoint = Nothing
    Set bFieldValue = Nothing
    Set bSequenceNumber = Nothing
    Set bSecurityField = Nothing
    Set bSecurityName = Nothing
    Set bSecurity = Nothing
    Set bOverrides = Nothing
    Set bSecurities = Nothing
    Set bIteratorData = Nothing
    Set bIterator = Nothing
    Set bEvent = Nothing
    Set bFieldArray = Nothing
    Set bSecurityArray = Nothing
    Set bRequest = Nothing
    Set bService = Nothing
    Call bSession.Stop
    Set bSession = Nothing
End Function

Private Function getServerData_reference()
    Set bIterator = bEvent.CreateMessageIterator
    
    Do While (bIterator.Next)
        Set bIteratorData = bIterator.Message
        Set bSecurities = bIteratorData.GetElement("securityData")
        
        Dim offsetNumber As Long, i As Long, j As Long
        nSecurities = bSecurities.Count
        '
        For i = 0 To (nSecurities - 1)
            Set bSecurity = bSecurities.GetValue(i)
            Set bSecurityName = bSecurity.GetElement("security")
            Set bSecurityField = bSecurity.GetElement("fieldData")
            Set bSequenceNumber = bSecurity.GetElement("sequenceNumber")
            offsetNumber = CInt(bSequenceNumber.Value)
            '
            For j = 0 To UBound(bInputFieldArray)
                If (bSecurityField.HasElement(bInputFieldArray(j))) Then
                    Set bFieldValue = bSecurityField.GetElement(bInputFieldArray(j))
                    bOutputArray(offsetNumber, j) = bFieldValue.Value
                End If
            Next j
        Next i
    Loop
End Function

Private Function getServerData_bulkReference()
    Set bIterator = bEvent.CreateMessageIterator
    nSecurity = nSecurity + 1
    
    Do While (bIterator.Next)
        Set bIteratorData = bIterator.Message
        Set bSecurities = bIteratorData.GetElement("securityData")
        Dim offsetNumber As Long, i As Long, j As Long
        Dim nSecurities As Long: nSecurities = bSecurities.Count
        '
        Set bSecurity = bSecurities.GetValue(0)
        Set bSecurityField = bSecurity.GetElement("fieldData")
        '
        If (bSecurityField.HasElement(bInputFieldArray(0))) Then
            Set bFieldValue = bSecurityField.GetElement(bInputFieldArray(0))
            '
            If ((bFieldValue.NumValues - 1) > UBound(bOutputArray, 2)) Then _
                ReDim Preserve bOutputArray(0 To UBound(bOutputArray, 1), 0 To bFieldValue.NumValues - 1)
            '
            For i = 0 To bFieldValue.NumValues - 1
                Set bDataPoint = bFieldValue.GetValue(i)
                bOutputArray(nSecurity - 1, i) = bDataPoint.GetElement(0).Value
            Next i
        End If
    Loop
End Function

Private Function getServerData_historical()
    Set bIterator = bEvent.CreateMessageIterator
    
    Do While (bIterator.Next)
        Set bIteratorData = bIterator.Message
        Set bSecurities = bIteratorData.GetElement("securityData")
        Dim nSecurities As Long: nSecurities = bSecurityArray.Count
        Set bSecurityField = bSecurities.GetElement("fieldData")
        
        Dim nItems As Long, offsetNumber As Long, nFields As Long, i As Long, j As Long
        
        Let nItems = bSecurityField.NumValues
        If (nItems = 0) Then Exit Function
        
        If ((nItems > UBound(bOutputArray, 2))) Then
            ReDim Preserve bOutputArray(0 To nSecurities - 1, 0 To nItems - 1)
        End If
        
        Set bSequenceNumber = bSecurities.GetElement("sequenceNumber")
        Let offsetNumber = CInt(bSequenceNumber.Value)
        
        If (bSecurityField.Count > 0) Then
            For i = 0 To (nItems - 1)
                
                If (bSecurityField.Count > i) Then
                    Set bFields = bSecurityField.GetValue(i)
                    If (bFields.HasElement(bFieldArray(0))) Then
                        
                        Dim d As Variant: ReDim d(0 To bFields.NumElements - 1)
                        For j = 0 To bFields.NumElements - 1
                            Let d(j) = bFields.GetElement(j).GetValue(0)
                        Next j
                        
                        Let bOutputArray(offsetNumber, i) = d
                    End If
                End If
            Next i
        End If
    Loop
End Function

Private Function getMax(ByVal a As Long, ByVal b As Long) As Long
    '
    getMax = a: If (b > a) Then getMax = b
End Function

