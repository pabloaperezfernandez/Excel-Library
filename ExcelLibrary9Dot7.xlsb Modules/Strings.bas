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
Public Function StringSplit(AString As String, aSeparator As String, theEltPos As Integer) As String
    Dim s As Variant
    
    Let s = Split(AString, aSeparator)
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
    
    ' Exit with Null if (StringOrStringArray1 fails StringArrayQ and StringQ)
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
            Length(StringOrStringArray1) <> Length(StringOrStringArray2) Then
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
    
    Let CreateFunctionalParameterArray = "(" & Join(Filter(ParameterArray, Lambda("x", "", "StrComp(x,vbnullstring)<>0")), ",") & ")"
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

Public Function ToParentheticalString(TheArray As Variant) As String
    Let ToParentheticalString = Convert1DArrayIntoParentheticalExpression(TheArray)
End Function

' DESCRIPTION
' Returns an array of numbered strings using the string provided by the user
' as the root.
'
' This function returns a sequence of numbered strings based on the user's
' specifications. It has two calling modalities.

' 1. When not passing the optional ToEndNumberQ set to True, the function interprets N
'    as the number of terms expected in the return sequence.
'
' 2. When passing the optional ToEndNumberQ set to True, the function returns a sequence
'    starting with StartNumber and with every other number obtained sequentially from
'    the prior adding TheStep (set to 1 if not passed) up to an including N.
'
' This function returns Null if N is negative when called in modality 1.  The function
' currently requires N to be larger than or equal to StartNumber when ToEndNumberQ=True.
'
' PARAMETERS
' 1. StartNumber - First number in the array
' 2. N - Number of elements in the sequence or the ending number, depending on the
'    calling modality
' 3. TheStep (optional) - To create a sequence using a sequential step different from 1
' 4. ToEndNumberQ (optional) - When passed explicitly as True, it activates
'    calling modality 2
'
' RETURNED VALUE
' The requested string sequence
Public Function GenerateStringSequence(TheStringRoot As String, _
                                       StartNumber As Variant, _
                                       n As Variant, _
                                       Optional TheStep As Variant, _
                                       Optional ToEndNumberQ As Boolean = False) As Variant
    Dim TheNumericSequence As Variant
    
    ' Set default return value in case of error
    Let GenerateStringSequence = Null
    
    Let TheNumericSequence = NumericalSequence(StartNumber, n, TheStep, ToEndNumberQ)
    
    ' ErrorCheck: Exit if the parameters caused an error
    If IsNull(TheNumericSequence) Then Exit Function
    
    ' Generate and return the requested string sequence
    Let GenerateStringSequence = StringJoin("x", ToStrings(TheNumericSequence))
End Function

' DESCRIPTION
' Returns a fully qualified sub/function name in the given module and workbook
'
' EXAMPLE
' RoutineName(wbk, "MyModule", "MyFunc") -> "'" & wbk.name & "'!MyModule.MyFunc"
'
' PARAMETERS
' 1. AWorkbook - A reference of type Workbook
' 2. ModuleName - Name of a module in AWorkbook
' 2. RoutineName - The string name of the sub/function
'
' RETURNED VALUE
' The fully quallified name of a function in a given workbook.
Public Function MakeRoutineName(AWorkbook As Workbook, _
                                ModuleName As String, _
                                RoutineName As String) As String
    Let MakeRoutineName = "'" & AWorkbook.Name & "'!" & _
                          IIf(RoutineName = "", _
                          RoutineName, _
                          ModuleName & "." & RoutineName)
End Function

' DESCRIPTION
' Returns the given string with sequences of spaces replaced by a single space
'
' PARAMETERS
' 1. AString - The string to process
'
' RETURNED VALUE
' the string with sequences of spaces replaced by a single space
Public Function RemoveDuplicatedSpaces(ByVal AString As String) As String
    Let RemoveDuplicatedSpaces = RemoveDuplicatedString(AString, " ")
End Function


' DESCRIPTION
' Returns the given string with sequences of spaces replaced by a single space
'
' PARAMETERS
' 1. AString - The string to process
'
' RETURNED VALUE
' the string with sequences of spaces replaced by a single space
Public Function RemoveDuplicatedString(ByVal AString As String, ByVal TargetChars As String) As String
    Let RemoveDuplicatedString = IIf(StrComp(Left(AString, Len(TargetChars)), TargetChars, vbTextCompare) = 0, TargetChars, vbNullString) & _
                                 Join(Filter(Split(AString, TargetChars), _
                                             Lambda("x", "", "StrComp(x, vbNullString,vbTextCompare)<>0")), _
                                      TargetChars) & _
                                 IIf(StrComp(Right(AString, Len(TargetChars)), TargetChars, vbTextCompare) = 0, TargetChars, vbNullString)
End Function

