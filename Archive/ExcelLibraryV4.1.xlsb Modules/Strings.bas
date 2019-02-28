Attribute VB_Name = "Strings"
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
    Set ArrayOf7DigitStringSedols = Nothing
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
