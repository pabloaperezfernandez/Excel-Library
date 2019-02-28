Attribute VB_Name = "StringFormulas"
Option Explicit
Option Base 1

' Scape single
Function esc(txt As String)
    esc = Trim(Replace(txt, "'", "\'"))
End Function

' This function returns its string after replacing all "\" with "/".
Public Function ReplaceForwardSlashWithBackSlash(txt As String) As String
    Let ReplaceForwardSlashWithBackSlash = Trim(Replace(txt, "\", "/"))
End Function

Public Function DeleteDoubleQuotes(txt As String) As String
    Let DeleteDoubleQuotes = Trim(Replace(txt, Chr(34), ""))
End Function

' This function splits the given string using the aSeparator string and returns the
' element theEltPos-1 of the resulting set of sub-strings.
Public Function StringSplit(aString As String, aSeparator As String, theEltPos As Integer) As String
    Dim s As Variant
    
    Let s = Split(aString, aSeparator)
    Let StringSplit = s(theEltPos - 1)
End Function

' DEPRECATED. Superseded by StringJoin
' This function concatenates an array of strings, and it returns the result
Public Function StringConcatenate(AStringArray As Variant) As String
    Dim NewString As String
    Dim CurrentString As Variant
    
    ' Initialize the return string to an empty string
    Let NewString = ""
    
    ' Concatenate each string in the array into the return string
    For Each CurrentString In AStringArray
        Let NewString = NewString & CurrentString
    Next CurrentString
    
    Let StringConcatenate = NewString
End Function

' Supersedes StringConcatenate
' Returns Null when the parameters makes no sense and either parameter is neither a string not a string array
' 1. An array is joined into a string
' 2. Two strings are not allowed
' 3. Two arrays of the same length result in an array with elementwise string joins
' 4. An array and a string results in the elementwise concatenation of the elements of the array with the string.
'    Concatenating from the left or the right or the left is determine by the lateral location of the string.  So,
'    if the string on the left, the string is concatenated from the left to every element in the array. We proceed
'    similarly if the string parameter in on the right, but the concatenation happens from the right for every element
'    in the array
Public Function StringJoin(StringOrStringArray1 As Variant, Optional StringOrStringArray2 As Variant) As Variant
    Dim NewString As String
    Dim AnArray As Variant
    Dim CurrentString As Variant
    Dim i As Long
    Dim IndexOffset As Integer
    
    ' Exit with Null if (StringOrStringArray1 fails String1DArrayQ and StringQ)
    If Not (StringArrayQ(StringOrStringArray1) Or StringArrayQ(StringOrStringArray2)) Then
        Let StringJoin = Null
        Exit Function
    End If
    
    ' Check StringOrStringArray2 for consistency
    If Not IsMissing(StringOrStringArray2) Then
         ' Exit with Null if StringOrStringArray2 fails IsArray and StringQ
        If Not (IsArray(StringOrStringArray2) Or StringQ(StringOrStringArray2)) Then
            Let StringJoin = Null
            Exit Function
        End If
        
        ' Exit with Null if StringOrStringArray1 and StringOrStringArray2 are both arrays
        ' but don't have the same length
        If IsArray(StringOrStringArray1) And IsArray(StringOrStringArray2) And _
            GetArrayLength(StringOrStringArray1) <> GetArrayLength(StringOrStringArray2) Then
            Let StringJoin = Null
            Exit Function
        End If
    End If
    
    ' Process the case of StringOrStringArray2 not passed
    If IsMissing(StringOrStringArray2) Then
        If IsArray(StringOrStringArray1) Then
            Let StringJoin = Join(StringOrStringArray1, "")
        Else
            Let StringJoin = StringOrStringArray1
        End If
        
        Exit Function
    End If
    
    ' If the code gets here, both parameters have been passed
    
    ' Process the case of both parameters being strings
    If StringQ(StringOrStringArray1) And StringQ(StringOrStringArray2) Then
        Let StringJoin = StringOrStringArray1 & StringOrStringArray2
        Exit Function
    End If
    
    ' Process case of StringOrStringArray1 being a string and StringOrStringArray2 an array
    If StringQ(StringOrStringArray1) Then
        Let AnArray = StringOrStringArray2
        For i = LBound(StringOrStringArray2) To UBound(StringOrStringArray2)
            Let AnArray(i) = StringOrStringArray1 & AnArray(i)
        Next
        Let StringJoin = AnArray
        
        Exit Function
    End If
    
    ' Process case of StringOrStringArray2 being an array and StringOrStringArray1 a string
    If StringQ(StringOrStringArray2) Then
        Let AnArray = StringOrStringArray1
        For i = LBound(StringOrStringArray1) To UBound(StringOrStringArray1)
            Let AnArray(i) = AnArray(i) & StringOrStringArray2
        Next
        Let StringJoin = AnArray
        
        Exit Function
    End If
    
    
    ' Process case of both parameters being strings
    Let AnArray = StringOrStringArray1
        
    If LBound(StringOrStringArray1) = 0 And LBound(StringOrStringArray2) = 1 Then
        Let IndexOffset = -1
    ElseIf LBound(StringOrStringArray1) = 1 And LBound(StringOrStringArray2) = 1 Then
        Let IndexOffset = 0
    Else
        Let IndexOffset = 1
    End If
    
    For i = LBound(AnArray) To UBound(AnArray)
        Let AnArray(i) = AnArray(i) & StringOrStringArray2(IndexOffset + i)
    Next i
    
    Let StringJoin = AnArray
End Function

' This function creates a functional parameter string from the array of parameters.
' This is useful when building functional calls dynamically.  The parameter array is 1D.
Public Function CreateFunctionalParameterArray(ParameterArray As Variant) As String
    Dim i As Integer
    
    Let CreateFunctionalParameterArray = "("
    
    For i = LBound(ParameterArray) To UBound(ParameterArray)
        Let CreateFunctionalParameterArray = CreateFunctionalParameterArray & ParameterArray(i)
        
        If ParameterArray(i) <> "" And i < UBound(ParameterArray) Then
            Let CreateFunctionalParameterArray = CreateFunctionalParameterArray & ", "
        End If
    Next i
    
    Let CreateFunctionalParameterArray = CreateFunctionalParameterArray & ")"
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
    Dim total As Long
    Dim i As Integer
    Dim s As String
    
    Let Weights = Array(1, 3, 1, 7, 3, 9)

    Let UpperCaseSedol = UCase(str)
    
    Let total = 0
    For i = 1 To 6
        Let s = Mid(UpperCaseSedol, i, 1)

        If Asc(s) >= 48 And Asc(s) <= 57 Then
                Let total = total + CInt(s) * Weights(i)
        Else
                total = total + (Asc(s) - 55) * Weights(i)
        End If
 
    Next i
    
    Let GetSedolCheckDigit = (10 - (total Mod 10)) Mod 10
End Function

' Converts a 1D array or 1D range (either Nx1 or 1xN) of SEDOLs into strings and ensures that all have length 6.
Public Function Enforce6DigitSedols(TheSedols As Variant) As Variant
    Dim tmpSht As Worksheet
    Dim TmpRange As Range
    
    Set tmpSht = ThisWorkbook.Worksheets("TempComputation")
    
    If TypeName(TheSedols) = "Range" Then
        Call ToTemp(Application.Transpose(ConvertTo1DArray(TheSedols.Value)))
    ElseIf TypeName(TheSedols) = "String" Or NumberQ(TheSedols) Then
        Let tmpSht.Range("A1").Value = TheSedols
    Else
        Call ToTemp(Application.Transpose(ConvertTo1DArray(TheSedols)))
    End If
    
    If TypeName(TheSedols) = "Range" Then
        Set TmpRange = tmpSht.Range("A1").Resize(TheSedols.Rows.Count, 1)
    ElseIf TypeName(TheSedols) = "String" Or NumberQ(TheSedols) Then
        Set TmpRange = tmpSht.Range("A1")
    Else
        Set TmpRange = tmpSht.Range("A1").Resize(UBound(TheSedols, 1))
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
    Let Enforce6DigitSedols = ConvertTo1DArray(TmpRange.Value)
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

' This function takes a 1D range and trims its contents after converting it to upper case
Public Function TrimAndConvertArrayToCaps(TheArray As Variant) As Variant
    Dim i As Long
    Dim j As Long
    Dim ResultsArray As Variant
    
    ' Exit with an empty array if TheArray is empty
    If EmptyArrayQ(TheArray) Then
        Let TrimAndConvertArrayToCaps = EmptyArray()
        Exit Function
    End If
    
    If NumberOfDimensions(TheArray) = 0 Then
        Let ResultsArray = UCase(Trim(TheArray))
    ElseIf NumberOfDimensions(TheArray) = 1 Then
        Let ResultsArray = TheArray
        
        For i = LBound(TheArray) To UBound(TheArray)
            Let ResultsArray(i) = UCase(Trim(TheArray(i)))
        Next i
    ElseIf NumberOfDimensions(TheArray) = 2 Then
        Let ResultsArray = TheArray
        
        For i = LBound(TheArray, 1) To UBound(TheArray, 1)
            For j = LBound(TheArray, 2) To UBound(TheArray, 2)
                Let ResultsArray(i, j) = UCase(Trim(TheArray(i, j)))
            Next j
        Next i
    Else
        Let ResultsArray = TheArray
    End If
    
    Let TrimAndConvertArrayToCaps = ResultsArray
End Function

Public Function Convert1DArrayIntoParentheticalExpression(TheArray As Variant) As String
    Let Convert1DArrayIntoParentheticalExpression = "(" & Join(TheArray, ",") & ")"
End Function


