Attribute VB_Name = "Identiiers"
Option Explicit
Option Base 1

' DESCRIPTION
' Given a Bloomberg identifier, this function returns the ticker. For instance,
' 1. IBM US Equity -> IBM
' 2. US4592001014 -> US4592001014
' 3. US4592001014 US Equity -> US4592001014
' 4. IBM 2 1/2 01/27/22 US Equity -> IBM 2 1/2 01/27/22
' 5. IBM 2 1/2 01/27/22 US -> IBM 2 1/2 01/27/22
'
' PARAMETERS
' 1. AnIdentifier - An identifier
'
' RETURNED VALUE
' The ticker of the identifier
Public Function GetTicker(ByVal AnIdentifier As String) As String
    Dim AnArray As Variant
    Dim AnId As String
    
    Let GetTicker = vbNullString
    Let AnId = RemoveDuplicatedSpaces(Trim(AnIdentifier))
    
    If AnId = vbNullString Then Exit Function
    
    If GetExchangeCode(AnId) = vbNullString And GetMarketSectorDes(AnId) = vbNullString Then
        Let GetTicker = AnId
    ElseIf GetExchangeCode(AnId) <> vbNullString And GetMarketSectorDes(AnId) = vbNullString Then
        Let GetTicker = Join(Most(Split(AnId, " ")))
    ElseIf GetExchangeCode(AnId) = vbNullString And GetMarketSectorDes(AnId) <> vbNullString Then
        Let GetTicker = Join(Most(Split(AnId, " ")))
    Else
        Let GetTicker = Join(Most(Most(Split(AnId, " "))), " ")
    End If
End Function

' DESCRIPTION
' Given a Bloomberg identifier, this function returns the exchange code. For instance,
' 1. IBM US Equity -> US
' 2. US4592001014 ->
' 3. US4592001014 US Equity -> US
' 4. IBM 2 1/2 01/27/22 US Equity -> US
' 5. IBM 2 1/2 01/27/22 US -> US
'
' PARAMETERS
' 1. AnIdentifier - An identifier
'
' RETURNED VALUE
' The exchange code of the identifier
Public Function GetExchangeCode(ByVal AnIdentifier As String) As String
    Dim AnArray As Variant
    Dim AnId As String

    Let GetExchangeCode = vbNullString
    
    ' This eliminates repeated spaces and then splits into words
    Let AnArray = FunctionalProgramming.Filter(Split(Trim(AnIdentifier), " "), _
                                               Lambda("x", "", "Not StrComp(x,"" "",vbTextCompare)"))
    Let AnId = Join(AnArray, " ")
    
    ' If there is only one word, this identifier has no exchange code.
    If Length(AnArray) < 2 Then Exit Function
    
    If StrComp(GetMarketSectorDes(AnId), vbNullString) = 0 Then
        If Len(Last(AnArray)) <> 2 Then Exit Function
        Let GetExchangeCode = Last(AnArray)
    Else
        If Len(Last(Most(AnArray))) <> 2 Then Exit Function
        Let GetExchangeCode = Last(Most(AnArray))
    End If
End Function

' DESCRIPTION
' Given a Bloomberg identifier, this function returns the MARKET_SECTOR_DES. For instance,
' 1. IBM US Equity -> Equity
' 2. US4592001014 ->
' 3. US4592001014 US Equity -> Equity
' 4. IBM 2 1/2 01/27/22 US Equity -> Equity
' 5. IBM 2 1/2 01/27/22 US ->
'
' PARAMETERS
' 1. AnIdentifier - An identifier
'
' RETURNED VALUE
' The exchange code of the identifier
Public Function GetMarketSectorDes(AnIdentifier As String) As String
    Dim TheLastPart As String
    
    ' This eliminates repeated spaces and then splits into words
    Let TheLastPart = Last(FunctionalProgramming.Filter(Split(Trim(AnIdentifier), " "), _
                                                        Lambda("x", "", "Not StrComp(x,"" "",vbTextCompare)")))

    If MemberQ(Array("Comdty", "Corp", "Curncy", "Equity", "Govt", "Index", "M -Mkt", "Mtge", "Muni", "Pfd"), _
               TheLastPart) Then
        Let GetMarketSectorDes = TheLastPart
    Else
        Let GetMarketSectorDes = vbNullString
    End If
End Function
' This function takes a string expected to be a 7-character SEDOL that may have dropped leading zeroes.
' It returns a seven character SEDOL, with as many zeroes as necessary on the left.
Public Function AddLeadingZeroesTo7CharacterSedol(arg As Variant) As String
    Dim TempString As String

    If Len(CStr(arg)) < 7 Then
        Let TempString = Application.Rept("0", 7 - Len(CStr(arg))) & CStr(arg)
    Else
        Let TempString = CStr(arg)
    End If
        
        ' Pull the Bloomberg ticker corresponding to this 7-char SEDOL from the master file
    Let AddLeadingZeroesTo7CharacterSedol = TempString
End Function

' This one takes a string that may be interpreted as a 6-character SEDOL.  It pads it on the left
' with as many zeroes as needed, and adds the checksum character on the right, returning a valid
' character 7-character SEDOL.
Public Function Make7CharacterSedol(arg As Variant) As String
    Dim TempString As String

    If Len(CStr(arg)) < 6 Then
        Let TempString = Application.Rept("0", 6 - Len(CStr(arg))) & CStr(arg)
    Else
        Let TempString = CStr(arg)
    End If

    If Len(TempString) < 7 Then
        Let TempString = TempString & CStr(GetSedolCheckDigit(TempString))
    End If
        
        ' Pull the Bloomberg ticker corresponding to this 7-char SEDOL from the master file
    Let Make7CharacterSedol = TempString
End Function

' Given a 6-character SEDOL, this function returns the checksum digit of the 7-character SEDOL
Public Function GetSedolCheckDigit(str As Variant) As String
    Dim Weights As Variant
    Dim UpperCaseSedol As String
    Dim Total As Long
    Dim i As Integer
    Dim s As String
    
    Let Weights = Array(1, 3, 1, 7, 3, 9)

    Let UpperCaseSedol = UCase(str)
    
    Let Total = 0
    For i = 1 To 6
        Let s = Mid(UpperCaseSedol, i, 1)
        
        If Asc(s) >= 48 And Asc(s) <= 57 Then
            Let Total = Total + CInt(s) * Weights(i)
        Else
            Let Total = Total + (Asc(s) - 55) * Weights(i)
        End If
 
    Next i
    
    Let GetSedolCheckDigit = (10 - (Total Mod 10)) Mod 10
End Function

' Converts a 1D array or 1D range (either Nx1 or 1xN) of SEDOLs into strings and ensures that all have length 6.
Public Function Enforce6DigitSedols(TheSedols As Variant) As Variant
    Dim tmpSht As Worksheet
    Dim TmpRange As Range
    
    Set tmpSht = ThisWorkbook.Worksheets("TempComputation")
    
    If TypeName(TheSedols) = "Range" Then
        Call ToTemp(Application.Transpose(Flatten(TheSedols)))
    ElseIf StringQ(TheSedols) Or NumberQ(TheSedols) Then
        Let tmpSht.Range("A1").Value = TheSedols
    Else
        Call ToTemp(Application.Transpose(Flatten(TheSedols)))
    End If
    
    If TypeName(TheSedols) = "Range" Then
        Set TmpRange = tmpSht.Range("A1").Resize(Length(TheSedols), 1)
    ElseIf TypeName(TheSedols) = "String" Or NumberQ(TheSedols) Then
        Set TmpRange = tmpSht.Range("A1")
    Else
        Set TmpRange = tmpSht.Range("A1").Resize(Length(TheSedols))
    End If

    Let TmpRange.Offset(0, 1).NumberFormat = "General"
    Let TmpRange.Offset(0, 1).Formula = _
        "=IF(LEN(R[0]C[-1])<6,REPT(" & Chr(34) & "0" & Chr(34) & ",6-LEN(R[0]C[-1]))&R[0]C[-1],R[0]C[-1])"
    Let TmpRange.Offset(0, 1).NumberFormat = "@"
    Let TmpRange.Offset(0, 1).Value = TmpRange.Offset(0, 1).Value
    Call TmpRange.ClearContents
    Let TmpRange.NumberFormat = "@"
    Let TmpRange.Value = TmpRange.Offset(0, 1).Value
    Call TmpRange.Offset(0, 1).ClearContents
    
    ' Returned fixed array of SEDOLs
    Let Enforce6DigitSedols = Flatten(TmpRange.Value)
End Function

' 1. Converts a 1D range (either Nx1 or 1xN) of SEDOLs into strings and ensure that all have length 6 if the length is less than or equal to 6.
' 2. Computes the 7th digit of the SEDOL identifier on the basis of PIAM's 6 digit SEDOL and appends it to the right of the 6 digit SEDOL.
Public Sub ConvertRangeOfSedolsToArrayOf7DigitStringSedols(TheSedols As Variant)
    Dim tmpSht As Worksheet
    Dim TmpRange As Range
    Dim ArrayOf7DigitStringSedols As Variant
    Dim j As Integer
            
    Set TmpRange = TheSedols.Offset(0, 25)
    Let TmpRange.NumberFormat = "General"
    Let TmpRange.Formula = _
        "=IF(LEN(R[0]C[-25])<6,REPT(" & Chr(34) & "0" & Chr(34) & ",6-LEN(R[0]C[-25]))&R[0]C[-25],R[0]C[-25])"
    Let TmpRange.NumberFormat = "@"
    Let TmpRange.Value = TmpRange.Value
    
    Call TheSedols.ClearContents
    Let TheSedols.NumberFormat = "@"
    Let TheSedols.Value = TmpRange.Value
    Call TmpRange.ClearContents
          
    ' Pre-allocate a string array that is long TheSedols.Rows.Count elements.
    Let ArrayOf7DigitStringSedols = Null
    ReDim ArrayOf7DigitStringSedols(TheSedols.Rows.Count)
     
    ' Loop through this file's securities, compute the 7th CheckSum digit of the Sedols and concatenate this digit
    ' to the right hand side of the 6 digit string Sedol
    For j = 1 To TheSedols.Rows.Count
        ' Compute the 7th CheckSum digit of the Sedols and concatenate this digit to the right hand side of the 6 digit string Sedol
        If Len(CStr(TheSedols.Cells(j))) <= 6 Then
            Let ArrayOf7DigitStringSedols(j) = CStr(TheSedols.Cells(j)) & CStr(GetSedolCheckDigit(TheSedols.Cells(j)))
        Else
            Let ArrayOf7DigitStringSedols(j) = CStr(TheSedols.Cells(j))
        End If
    Next j
    
    Let TheSedols.Range("A1").Resize(UBound(ArrayOf7DigitStringSedols, 1)) = Application.Transpose(ArrayOf7DigitStringSedols)
End Sub

' The purpose of this function is:
' 1. turn a ticker+" "+exchange code into ticker+ " "+exchange code+" Equity".
'    If there no exchange code is included in parameter TheTicker, the last
'    two characters are assumed to represent the exchange code (as Bloomberg
'    sometimes screws up when it returns the EQY_FUND_TICKER).  If the sufix
'    " Equity" is present. No aditional work is done.
Public Function EnforceBloombergEquityTicker(TheTicker As String) As String
    Dim ATicker As String

    ' Transform what Bloomberg returns as EQY_FUND_TICKER into a valid ticker
    Let ATicker = UCase(Trim(TheTicker))
    
    If InStr(1, ATicker, " EQUITY") > 0 Then
        Let ATicker = Trim(Left(ATicker, Len(ATicker) - Len(" EQUITY")))
    End If
    
    If ATicker <> "NULL" Then
        If InStr(1, ATicker, " ") = 0 Then
            Let ATicker = Left(ATicker, Len(ATicker) - 2) & " " & Right(ATicker, 2)
        End If
        
        Let ATicker = ATicker & " Equity"
    Else
        ATicker = TheTicker
    End If
    
    Let EnforceBloombergEquityTicker = ATicker
End Function

