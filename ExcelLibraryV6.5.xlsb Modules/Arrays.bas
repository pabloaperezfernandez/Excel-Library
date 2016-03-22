Attribute VB_Name = "Arrays"
Option Explicit
Option Base 1

' DESCRIPTION
' This function takes a ParamArray and copies its elements into a regular variant array.  If
' the ParamArray is missing, this function returns an empty array.
'
' PARAMETERS
' 1. arg - a variant ParamArray
'
' RETURNED VALUE
' A variant array with the same elements as the ParamArray argument
Public Function CopyParamArray(ParamArray Args() As Variant) As Variant()
    Dim ArrayCopy() As Variant
    Dim i As Long

    If Not IsMissing(Args(0)) Then
        ReDim ArrayCopy(LBound(Args(0)) To UBound(Args(0)))
        
        For i = LBound(Args(0)) To UBound(Args(0))
            Let ArrayCopy(i) = Args(0)(i)
        Next
    End If

    Let CopyParamArray = ArrayCopy
End Function

' DESCRIPTION
' This function returns the number of dimensions of any VBA expression.  Non-array expressions all have
' 0 dimensions.  Undimemsioned arrays have 0 dimensions.
'
' PARAMETERS
' 1. arg - any value or object reference
'
' RETURNED VALUE
' Number of dimensions for its argument.  Dimensioned arrays are the only expressions returning non-zero.
Public Function EmptyArray() As Variant
    Let EmptyArray = Array()
End Function

' DESCRIPTION
' This function returns the number of dimensions of any VBA expression.  Non-array expressions all have
' 0 dimensions.  Undimemsioned arrays have 0 dimensions.
'
' PARAMETERS
' 1. arg - any value or object reference
'
' RETURNED VALUE
' Number of dimensions for its argument.  Dimensioned arrays are the only expressions returning non-zero.
Public Function NumberOfDimensions(MyArray As Variant) As Long
    Dim temp As Long
    Dim i As Long
    
    On Error GoTo FinalDimension
    
    ' EmptyArrayQ returns false if MyArray is an undimensioned array
    If EmptyArrayQ(MyArray) Then
        Let NumberOfDimensions = 1
        Exit Function
    End If
    
    If Not IsArray(MyArray) Then
        Let NumberOfDimensions = 0
        Exit Function
    End If
    
    For i = 1 To 60000
        Let temp = LBound(MyArray, i)
    Next i

    Let NumberOfDimensions = i
    
    Exit Function
        
FinalDimension:
    Let NumberOfDimensions = i - 1
    Exit Function
End Function

' DESCRIPTION
' Alias for NumberOfDimensions().  This function returns the number of dimensions of any VBA expression.
' Non-array expressions all have 0 dimensions.  Undimemsioned arrays have 0 dimensions.
'
' PARAMETERS
' 1. arg - any value or object reference
'
' RETURNED VALUE
' Number of dimensions for its argument.  Dimensioned arrays are the only expressions returning non-zero.
Public Function GetNumberOfDimensions(MyArray As Variant) As Long
    Let GetNumberOfDimensions = NumberOfDimensions(MyArray)
End Function

' DESCRIPTION
' This function returns the length of the given 1D or 2D array.  A 2D array is interpreted as 1D array
' of its rows. Non-array expressions all have 0 dimensions.  Undimemsioned arrays have 0 dimensions.
'
' PARAMETERS
' 1. arg - any value or object reference
'
' RETURNED VALUE
' Length of the given array.  Dimensioned arrays are the only expressions returning non-zero.
Public Function GetArrayLength(AnArray As Variant) As Long
    If Not DimensionedQ(AnArray) Then
        Let GetArrayLength = 0
    Else
        Let GetArrayLength = UBound(AnArray, 1) - LBound(AnArray, 1) + 1
    End If
End Function

' DESCRIPTION
' Alias for GetArrayLength.  This function returns the length of the given 1D or 2D array.  A 2D
' array is interpreted as 1D array of its rows. Non-array expressions all have 0 dimensions.
' Undimemsioned arrays have 0 dimensions.
'
' PARAMETERS
' 1. arg - any value or object reference
'
' RETURNED VALUE
' Length of the given array.  Dimensioned arrays are the only expressions returning non-zero.
Public Function Length(AnArray As Variant) As Long
    Let Length = GetArrayLength(AnArray)
End Function

' DESCRIPTION
' This function returns the number of rows in the given 2D array.  An empty array is said to have
' 0 rows.  Only dimensioned arrays return non-zero values.  All other parameters return 0.
'
' PARAMETERS
' 1. arg - any value or object reference
'
' RETURNED VALUE
' Length of the given array.  Dimensioned arrays are the only expressions returning non-zero.
Public Function GetNumberOfRows(aMatrix As Variant) As Long
    If Not DimensionedQ(aMatrix) Then
        Let GetNumberOfRows = 0
    ElseIf NumberOfDimensions(aMatrix) = 1 Then
        Let GetNumberOfRows = 1
    Else
        Let GetNumberOfRows = UBound(aMatrix, 1) - LBound(aMatrix, 1) + 1
    End If
End Function

' DESCRIPTION
' Alias for GetNumberOfRows.  This function returns the number of rows in the given 2D array.
' An empty array is said to have 0 rows.  Only dimensioned arrays return non-zero values.  All
' other parameters return 0.
'
' PARAMETERS
' 1. arg - any value or object reference
'
' RETURNED VALUE
' Length of the given array.  Dimensioned arrays are the only expressions returning non-zero.
Public Function NumberOfRows(AnArray As Variant) As Long
    Let NumberOfRows = GetNumberOfRows(AnArray)
End Function

' DESCRIPTION
' This function returns the number of columns in the given 2D array.  An empty array is said to
' have 0 columns.  Only dimensioned arrays return non-zero values.  All other parameters return 0.
'
' PARAMETERS
' 1. arg - any value or object reference
'
' RETURNED VALUE
' Number of columns of the given array.  Dimensioned arrays are the only expressions returning non-zero.
Public Function GetNumberOfColumns(aMatrix As Variant) As Long
    If Not DimensionedQ(aMatrix) Then
        Let GetNumberOfColumns = 0
    ElseIf NumberOfDimensions(aMatrix) = 1 Then
        Let GetNumberOfColumns = UBound(aMatrix, 1) - LBound(aMatrix, 1) + 1
    Else
        Let GetNumberOfColumns = UBound(aMatrix, 2) - LBound(aMatrix, 2) + 1
    End If
End Function

' DESCRIPTION
' Alias for GetNumberOfColumns. This function returns the number of columns in the given 2D array.
' An empty array is said to have 0 columns.  Only dimensioned arrays return non-zero values.  All
' other parameters return 0.
'
' PARAMETERS
' 1. arg - any value or object reference
'
' RETURNED VALUE
' Number of columns of the given array.  Dimensioned arrays are the only expressions returning non-zero.
Public Function NumberOfColumns(AnArray As Variant) As Long
    Let NumberOfColumns = GetNumberOfColumns(AnArray)
End Function

' DESCRIPTION
' Returns the first element in the given array.  Returns Null if the array is empty or
' there is a problem with it.  2D arrays are treated as an array of arrays, with each
' row being one of the elements in the array.  In other words, this function would return
' the same thing for [{1,2; 3,4}] and Array(Array(1,2), Array(3,4)).
'
' PARAMETERS
' 1. arg - any value or object reference
'
' RETURNED VALUE
' First element in the given array
Public Function First(AnArray As Variant) As Variant
    If Not DimensionedQ(AnArray) Then
        Let First = Null
    ElseIf NumberOfDimensions(AnArray) = 1 Then
        If EmptyArrayQ(AnArray) Then
            Let First = Null
        Else
            Let First = AnArray(LBound(AnArray))
        End If
    ElseIf NumberOfDimensions(AnArray) = 2 Then
        Let First = Flatten(GetRow(AnArray, 1))
    Else
        Let First = AnArray
    End If
End Function

' DESCRIPTION
' Returns the last element in the given array.  Returns Null if the array is empty or
' there is a problem with it.  2D arrays are treated as an array of arrays, with each
' row being one of the elements in the array.  In other words, this function would return
' the same thing for [{1,2; 3,4}] and Array(Array(1,2), Array(3,4)).
'
' PARAMETERS
' 1. arg - any value or object reference
'
' RETURNED VALUE
' Last element in the given array
Public Function Last(AnArray As Variant) As Variant
    If Not DimensionedQ(AnArray) Then
        Let Last = Null
    ElseIf NumberOfDimensions(AnArray) = 1 Then
        If EmptyArrayQ(AnArray) Then
            Let Last = Null
        Else
            Let Last = AnArray(UBound(AnArray))
        End If
    ElseIf NumberOfDimensions(AnArray) = 2 Then
        Let Last = Flatten(GetRow(AnArray, GetNumberOfRows(AnArray)))
    Else
        Let Last = AnArray
    End If
End Function

' DESCRIPTION
' Returns the 1D array given by all but the last element or the 2D given by all but the last row.
' Returns Null if there is a problem with its argument.  Most of a one-element, 1D array or one-row,
' 2D array returns the empty array.  2D arrays are treated as an array of arrays, with each row being
' one of the elements in the array.  In other words, this function would return the same thing for
' [{1,2; 3,4}] and Array(Array(1,2), Array(3,4)).
'
' PARAMETERS
' 1. arg - any value or object reference
'
' RETURNED VALUE
' An array with all but the last element of the array, interpreting 2D arrays as 1D arrays of their rows
Public Function Most(AnArray As Variant) As Variant
    If NumberOfDimensions(AnArray) = 1 Then
        If EmptyArrayQ(AnArray) Then
            Let Most = EmptyArray()
        ElseIf UBound(AnArray) = LBound(AnArray) Then
            Let Most = EmptyArray()
        Else
            Let Most = GetSubArray(AnArray, _
                                   LBound(AnArray), _
                                   UBound(AnArray) - 1)
        End If
    ElseIf NumberOfDimensions(AnArray) = 2 Then
        If EmptyArrayQ(AnArray) Then
            Let Most = EmptyArray()
        ElseIf UBound(AnArray, 1) = LBound(AnArray, 1) Then
            Let Most = EmptyArray()
        Else
            Let Most = GetSubMatrix(AnArray, _
                                    LBound(AnArray, 1), _
                                    UBound(AnArray, 1) - 1, _
                                    LBound(AnArray, 2), _
                                    UBound(AnArray, 2))
        End If
    Else
        Let Most = Null
    End If
End Function

' DESCRIPTION
' Returns the 1D array given by all but the first element or the 2D given by all but the first row.
' Returns Null if there is a problem with its argument.  Rest of a one-element ,1D array or one-row 2D
' array returns the empty array.  2D arrays are treated as an array of arrays, with each row being one
' of the elements in the array.  In other words, this function would return the same thing for
' [{1,2; 3,4}] and Array(Array(1,2), Array(3,4)).
'
' PARAMETERS
' 1. arg - any value or object reference
'
' RETURNED VALUE
' An array with all but the first element of the array, interpreting 2D arrays as 1D arrays of their rows
Public Function Rest(AnArray As Variant) As Variant
    If Not IsArray(AnArray) Then
        Let Rest = AnArray
    ElseIf NumberOfDimensions(AnArray) = 1 Then
        If EmptyArrayQ(AnArray) Then
            Let Rest = EmptyArray()
        ElseIf UBound(AnArray) = LBound(AnArray) Then
            Let Rest = EmptyArray()
        Else
            Let Rest = GetSubArray(AnArray, _
                                   LBound(AnArray) + 1, _
                                   UBound(AnArray))
        End If
    ElseIf NumberOfDimensions(AnArray) = 2 Then
        If EmptyArrayQ(AnArray) Then
            Let Rest = EmptyArray()
        ElseIf UBound(AnArray, 1) = LBound(AnArray, 1) Then
            Let Rest = EmptyArray()
        Else
            Let Rest = GetSubMatrix(AnArray, _
                                    LBound(AnArray, 1) + 1, _
                                    UBound(AnArray, 1), _
                                    LBound(AnArray, 2), _
                                    UBound(AnArray, 2))
        End If
    Else
        Let Rest = Null
    End If
End Function

' DESCRIPTION
' Flattens a system of nested arrays, regardless of the complexity of array nesting.  The leaves of
' the tree represented by the nested array can have any Excel type. The leaves of the tree must be
' atomic elements (e.g. satisfy AtomicQ) or the function returns Null.
'
' PARAMETERS
' 1. arg - any value or object reference
'
' RETURNED VALUE
' A 1D array all the values in the leaves of the tree represented by the original, nested array.
Public Function Flatten(a As Variant) As Variant
    Dim var As Variant
    Dim var2 As Variant
    Dim TempVariant As Variant
    Dim ResultsDict As Dictionary
    
    If AtomicQ(a) Then
        Let Flatten = a
        Exit Function
    End If
    
    If Not IsArray(a) Then
        Let Flatten = Null
        Exit Function
    End If

    Set ResultsDict = New Dictionary
    
    For Each var In a
        If AtomicQ(var) Then
            Call ResultsDict.Add(Key:=ResultsDict.Count, Item:=var)
        ElseIf IsArray(var) Then
            Let TempVariant = Flatten(var)
            
            If IsNull(TempVariant) Then
                Let Flatten = Null
                Exit Function
            End If
        
            For Each var2 In Flatten(var)
                Call ResultsDict.Add(Key:=ResultsDict.Count, Item:=var2)
            Next
        Else
            Let Flatten = Null
            Exit Function
        End If
    Next
    
    Let Flatten = ResultsDict.Items
End Function

' DESCRIPTION
' This function translates the given non-zero index to the array's instrinsic covention.
' For instance, we want to convert TheIndexOrIndexArray=1 into AnArray's intrinsic convention
' that LBound=0 and UBound=2.  Then, this function maps 1 to 0.  Negative indices map -1 to the
' last element of the array. The function returns Null if Abs(TheIndex)>Length(AnArray) or
' Abs(TheIndex)<1.  When the optional parameter RelativeToColumnsQ is explicit set
' to True, this function performs its operations relative to AnArray's columns.
'
' PARAMETERS
' 1. AnArray - A non-Empty 1D array
' 2. TheIndex - a whole number between 1 and Length(AnArray) or between -Length(AnArray) and -1
'
' RETURNED VALUE
' Translated index from lbound = 1 to the array's intrinsic convention
Public Function NormalizeIndex(AnArray As Variant, _
                               TheIndex As Variant, _
                               Optional DimensionIndexRelativeTo As Long = 1, _
                               Optional ParameterCheckQ As Boolean = True) As Variant
    ' This is here for high-speed applications
    If Not ParameterCheckQ Then
        Let NormalizeIndex = TheIndex
        Exit Function
    End If
                               
    ' Exit with Null if AnArray is undimensioned
    If Not DimensionedQ(AnArray) Then
        Let NormalizeIndex = Null
        Exit Function
    End If
     
    ' Exit if AnArray is the empty 1D array
    If EmptyArrayQ(AnArray) Then
        Let NormalizeIndex = Null
        Exit Function
    End If
        
    ' Exit with Null if TheIndex is not a positive integer
    If Not NonzeroWholeNumberQ(TheIndex) Then
        Let NormalizeIndex = Null
        Exit Function
    End If
    
    ' Exit with Null if TheIndex is outside of acceptable bounds
    If DimensionIndexRelativeTo > NumberOfDimensions(AnArray) Then
        Let NormalizeIndex = Null
        Exit Function
    End If
    
    If Abs(TheIndex) < 1 Or Abs(TheIndex) > UBound(AnArray, DimensionIndexRelativeTo) - LBound(AnArray, DimensionIndexRelativeTo) + 1 Then
        Let NormalizeIndex = Null
        Exit Function
    End If
    
    ' Handles non-negative TheIndex case
    If TheIndex > 0 Then
        Let NormalizeIndex = TheIndex + LBound(AnArray, DimensionIndexRelativeTo) - 1
    ' Handles negative TheIndex case
    Else
        Let NormalizeIndex = TheIndex + 1 + UBound(AnArray, DimensionIndexRelativeTo)
    End If
End Function

' DESCRIPTION
' Applies NormalizeIndex to each element of the given array
'
' PARAMETERS
' 1. AnArray - A non-Empty 1D array
' 2. IndexArray - A dimensioned array, each of whose elements is a whole number between 1 and
'    Length(AnArray) or between -Length(AnArray) and -1
'
' RETURNED VALUE
' Translated indices from lbound = 1 to the array's intrinsic convention
Public Function NormalizeIndexArray(AnArray As Variant, _
                                    TheIndices As Variant, _
                                    Optional DimensionIndexRelativeTo As Long = 1, _
                                    Optional ParameterCheckQ As Boolean = True) As Variant
    Dim ResultArray() As Variant
    Dim i As Long
                       
    If ParameterCheckQ Then
        ' Exit with Null if AnArray is undimensioned
        If Not DimensionedQ(AnArray) Then
            Let NormalizeIndexArray = Null
            Exit Function
        End If
         
        ' Exit if AnArray is the empty 1D array
        If EmptyArrayQ(AnArray) Then
            Let NormalizeIndexArray = Null
            Exit Function
        End If
        
        ' Exit with Null if AnArray fails both AtomicArrayQ and AtomicTableQ.
        ' In other words, exit with an error code if AnArray is neither an accetable 1D array or 2D table
        If Not (AtomicArrayQ(AnArray) Or AtomicTableQ(AnArray)) Then
            Let NormalizeIndexArray = Null
            Exit Function
        End If
        
        ' Exit with Null if TheIndex is not a positive integer
        If Not NonzeroWholeNumberArrayQ(TheIndices) Then
            Let NormalizeIndexArray = Null
            Exit Function
        End If
        
        ' Exit with Null if DimensionIndexRelativeTo larger than the number of dimensions
        If DimensionIndexRelativeTo > NumberOfDimensions(AnArray) Then
            Let NormalizeIndexArray = Null
            Exit Function
        End If
        
        If EmptyArrayQ(TheIndices) Then
            Let NormalizeIndexArray = EmptyArray()
            Exit Function
        End If
    End If

    ReDim ResultArray(LBound(TheIndices) To UBound(TheIndices))
    For i = LBound(TheIndices) To UBound(TheIndices)
        Let ResultArray(i) = NormalizeIndex(AnArray, TheIndices(i), DimensionIndexRelativeTo)
    Next
    
    Let NormalizeIndexArray = ResultArray
End Function

' DESCRIPTION
' Returns a sequential array of N whole numbers starting with StartNumber and with sequential step
' TheStep.  This function returns Null if N is negative.
'
' PARAMETERS
' 1. AnArray - A non-Empty 1D array
' 2. IndexArray - A dimensioned array, each of whose elements is a whole number between 1 and
'    Length(AnArray) or between -Length(AnArray) and -1
'
' RETURNED VALUE
' Translated indices from lbound = 1 to the array's intrinsic convention
Public Function CreateSequentialArray(StartNumber As Long, N As Long, Optional TheStep As Long = 1)
    Dim TheArray As Variant
    Dim i As Long
    
    If N < 0 Then
        Let CreateSequentialArray = Null
        Exit Function
    End If
    
    ReDim TheArray(1 To N)
    
    For i = StartNumber To StartNumber + N - 1
        Let TheArray(i - StartNumber + 1) = StartNumber + (i - StartNumber) * TheStep
    Next i
    
    Let CreateSequentialArray = TheArray
End Function

' DESCRIPTION
' This function returns the requested slice of an array.  It works just like Mathematica's Part[].  The returned
' value depends on the form of parameter Indices.  Works on 1D and 2D arrays.  As opposed to Part[], which can be
' very slow due to stringent parameter consistency checks, this version of Part[] is extremely fast, but it is up
' to the programmer to ensure all parameters are consistent.
'
' PARAMETERS
' 1. AnArray - A dimensioned array
' 2. Indices - a sequence of indices (with at least one supplied) of the forms below, with each one
'    referring to a different dimension of the array. At the moment we process only 1D and 2D arrays.
'    So, Indices can only be one or two of the forms below.
'
' Indices can take any of the following forms:
' 1. n - equivalent to Array(n) to get element n
' 2. [{n_1, n_2}] - Equivalent to Array(n_1, n_2) to get elements n_1 through n_2
' 3. [{n_1, n_2, step}] - Equivalent to Array(n_1, n_2, step) to get elements n_1 through n_2 every step
'    elements
' 4. Array([{n_1, n_2, ..., n_n}]) - Equivalent to Array(Array(n_1, n_2, ..., n_k)) to get elements n_1, n_2,
'    ..., n_k.
'
' It is important to understand that [{...}] is used as a shorthand notation
' to specify both 1D and proper 2D arrays.  Hence, [{1,2,3}]=Array(1,2,3) but
' [{[{1,2}],[{3,4}]}] is not equal to Array(Array(1,2), Array(3,4)).  In fact,
' [{[{1,2}],[{3,4}]}] raises an error.  However Array([{...}], [{...}]) is a valid
' syntax to specify a 1D array of 1D arrays without having to use the more cumbersome
' Array(Array(...), ..., Array(...))
'
' RETURNED VALUE
' The requested slice or element of the array.
Public Function PartNoParamCheck(AnArray As Variant, ParamArray Indices() As Variant) As Variant
    Dim IndicesCopy1 As Variant
    Dim IndicesCopy As Variant
    Dim IndexIndex As Variant
    Dim AnIndex As Variant
    Dim ReturnArray() As Variant
    Dim ni As Variant
    Dim ir As Long ' for index r
    Dim ic As Long ' for index c
    Dim r As Long ' for matrix row
    Dim c As Long ' for matrix column
    Dim RowIndices As Variant
    Dim ColumnIndices As Variant
    Dim var As Variant
    
    ' We use our function to convert the ParamArray to a regular array
    Let IndicesCopy1 = CopyParamArray(Indices)
    ReDim IndicesCopy(1 To Length(IndicesCopy1))
    For r = 1 To Length(IndicesCopy1)
        Let IndicesCopy(r) = IndicesCopy1(r - 1)
    Next
    
    ' Loop over the dimensions converting each index specification into an array of individual
    ' element positions
    For ir = 1 To Length(IndicesCopy)
        ' Convert from out lbound=1 convention to IndicesCopy's intrinsic convention
        Let IndexIndex = NormalizeIndex(IndicesCopy, ir, ParameterCheckQ:=False)
        
        ' Get the current dimensional index
        Let AnIndex = IndicesCopy(IndexIndex)
    
        If PartIndexSingleElementQ(AnIndex) Then
            Let IndicesCopy(IndexIndex) = Array(NormalizeIndex(AnArray, AnIndex, ir, ParameterCheckQ:=False))
        ElseIf PartIndexSequenceOfElementsQ(AnIndex) Then
            If NormalizeIndex(AnArray, First(AnIndex), ir) = NormalizeIndex(AnArray, Last(AnIndex), ir, ParameterCheckQ:=False) Then
                Let IndicesCopy(IndexIndex) = Array(NormalizeIndex(AnArray, First(AnIndex), ir, ParameterCheckQ:=False))
            Else
                Let IndicesCopy(IndexIndex) = PartIntervalIndices(AnArray, AnIndex, ir, False)
            End If
        ElseIf PartIndexSteppedSequenceQ(AnIndex) Then
            Let IndicesCopy(IndexIndex) = PartSteppedIntervalIndices(AnArray, AnIndex, ir, False)
        Else
            Let IndicesCopy(IndexIndex) = NormalizeIndexArray(AnArray, First(AnIndex), ir, ParameterCheckQ:=False)
        End If
    Next
    
    ' Collect the chosen array's slice depending on its number of dimensions
    If NumberOfDimensions(AnArray) = 1 Then
        Let ColumnIndices = First(IndicesCopy)
    
        ' Pre-allocate a 1D array since AnArray is one-dimensional
        ReDim ReturnArray(1 To Length(ColumnIndices))
        
        ' Extract the requested elements from AnArray
        For c = 1 To Length(ColumnIndices)
            Let ReturnArray(c) = AnArray(ColumnIndices(NormalizeIndex(ColumnIndices, c, ParameterCheckQ:=False)))
        Next
    ElseIf NumberOfDimensions(AnArray) = 2 Then
        ' Get all columns from the requested rows
        If Length(IndicesCopy) = 1 Then
            Let RowIndices = First(IndicesCopy)
        
            ReDim ReturnArray(1 To Length(RowIndices), 1 To NumberOfColumns(AnArray))
            For ir = 1 To Length(RowIndices)
                For c = 1 To NumberOfColumns(AnArray)
                    Let ReturnArray(ir, c) = AnArray(RowIndices(NormalizeIndex(RowIndices, ir, ParameterCheckQ:=False)), _
                                                     NormalizeIndex(AnArray, c, 2, ParameterCheckQ:=False))
                Next
            Next
        ' Get all elements requested
        Else
            Let RowIndices = First(IndicesCopy)
            Let ColumnIndices = Last(IndicesCopy)
        
            ReDim ReturnArray(1 To Length(RowIndices), 1 To Length(ColumnIndices))
            For ir = 1 To Length(RowIndices)
                For ic = 1 To Length(ColumnIndices)
                    Let ReturnArray(ir, ic) = AnArray(RowIndices(NormalizeIndex(RowIndices, ir, ParameterCheckQ:=False)), _
                                                      ColumnIndices(NormalizeIndex(ColumnIndices, ic, ParameterCheckQ:=False)))
                Next
            Next
        End If
    Else
        Let PartNoParamCheck = Null
    End If
    
    Let PartNoParamCheck = ReturnArray
End Function

' DESCRIPTION
' This function returns the requester part of an array.  It works just like Mathematica's Part[].  The returned
' value depends on the form of parameter Indices.  Works on 1D and 2D arrays.
'
' PARAMETERS
' 1. AnArray - A dimensioned array
' 2. Indices - a sequence of indices (with at least one supplied) of the forms below, with each one
'    referring to a different dimension of the array. At the moment we process only 1D and 2D arrays.
'    So, Indices can only be one or two of the forms below.
'
' Indices can take any of the following forms:
' 1. n - equivalent to Array(n) to get element n
' 2. [{n_1, n_2}] - Equivalent to Array(n_1, n_2) to get elements n_1 through n_2
' 3. [{n_1, n_2, step}] - Equivalent to Array(n_1, n_2, step) to get elements n_1 through n_2 every step
'    elements
' 4. Array([{n_1, n_2, ..., n_n}]) - Equivalent to Array(Array(n_1, n_2, ..., n_k)) to get elements n_1, n_2,
'    ..., n_k.
'
' It is important to understand that [{...}] is used as a shorthand notation
' to specify both 1D and proper 2D arrays.  Hence, [{1,2,3}]=Array(1,2,3) but
' [{[{1,2}],[{3,4}]}] is not equal to Array(Array(1,2), Array(3,4)).  In fact,
' [{[{1,2}],[{3,4}]}] raises an error.  However Array([{...}], [{...}]) is a valid
' syntax to specify a 1D array of 1D arrays without having to use the more cumbersome
' Array(Array(...), ..., Array(...))
'
' RETURNED VALUE
' The requested slice or element of the array.
Public Function Part(AnArray As Variant, ParamArray Indices() As Variant) As Variant
    Dim IndicesCopy As Variant
    Dim IndexIndex As Variant
    Dim AnIndex As Variant
    Dim ReturnArray() As Variant
    Dim ni As Variant
    Dim ir As Long ' for index r
    Dim ic As Long ' for index c
    Dim r As Long ' for matrix row
    Dim c As Long ' for matrix column
    Dim RowIndices As Variant
    Dim ColumnIndices As Variant
    Dim var As Variant
    
    ' We use our function to convert the ParamArray to a regular array
    Let IndicesCopy = CopyParamArray(Indices)

    ' Exit with Null if AnArray is neither an atomic array nor an atomic table
    If Not (AtomicArrayQ(AnArray) Or AtomicTableQ(AnArray)) Then
        Let Part = Null
        Exit Function
    End If

    ' Exit with Null if not even a single index was passed
    If Not DimensionedQ(IndicesCopy) Then
        Let Part = Null
        Exit Function
    End If
    
    ' Exit if more indices were passed than the number of dimensions in AnArray
    If NumberOfDimensions(AnArray) < Length(IndicesCopy) Then
        Let Part = Null
        Exit Function
    End If
    
    ' Exit with Null if any of the indices fails PartIndexQ
    If Not PartIndexArrayQ(IndicesCopy) Then
        Let Part = Null
        Exit Function
    End If
    
    ' Exit with Null if AnArray is empty because any index would be out of bounds
    If EmptyArrayQ(AnArray) Then
        Let Part = Null
        Exit Function
    End If
    
    ' Loop over the dimensions converting each index specification into an array of individual
    ' element positions
    For ir = 1 To Length(IndicesCopy)
        ' Convert from out lbound=1 convention to IndicesCopy's intrinsic convention
        Let IndexIndex = NormalizeIndex(IndicesCopy, ir)
        
        ' Get the current dimensional index
        Let AnIndex = IndicesCopy(IndexIndex)
    
        If PartIndexSingleElementQ(AnIndex) Then
            Let IndicesCopy(IndexIndex) = Array(NormalizeIndex(AnArray, AnIndex, ir))
        ElseIf PartIndexSequenceOfElementsQ(AnIndex) Then
            If NormalizeIndex(AnArray, First(AnIndex), ir) = NormalizeIndex(AnArray, Last(AnIndex), ir) Then
                Let IndicesCopy(IndexIndex) = Array(NormalizeIndex(AnArray, First(AnIndex), ir))
            Else
                Let IndicesCopy(IndexIndex) = PartIntervalIndices(AnArray, AnIndex, ir)
            End If
        ElseIf PartIndexSteppedSequenceQ(AnIndex) Then
            Let IndicesCopy(IndexIndex) = PartSteppedIntervalIndices(AnArray, AnIndex, ir)
        Else
            Let IndicesCopy(IndexIndex) = NormalizeIndexArray(AnArray, First(AnIndex), ir)
        End If
        
        If NullQ(IndicesCopy(IndexIndex)) Then
            Let Part = Null
            Exit Function
        End If
        
        If AnyTrueQ(IndicesCopy(IndexIndex), ThisWorkbook, "NullQ") Then
            Let Part = Null
            Exit Function
        End If
    Next
    
    ' Collect the chosen array's slice depending on its number of dimensions
    If NumberOfDimensions(AnArray) = 1 Then
        Let ColumnIndices = First(IndicesCopy)
    
        ' Pre-allocate a 1D array since AnArray is one-dimensional
        ReDim ReturnArray(1 To Length(ColumnIndices))
        
        ' Extract the requested elements from AnArray
        For c = 1 To Length(ColumnIndices)
            Let ReturnArray(c) = AnArray(ColumnIndices(NormalizeIndex(ColumnIndices, c)))
        Next
    ElseIf NumberOfDimensions(AnArray) = 2 Then
        ' Get all columns from the requested rows
        If Length(IndicesCopy) = 1 Then
            Let RowIndices = First(IndicesCopy)
        
            ReDim ReturnArray(1 To Length(RowIndices), 1 To NumberOfColumns(AnArray))
            For ir = 1 To Length(RowIndices)
                For c = 1 To NumberOfColumns(AnArray)
                    Let ReturnArray(ir, c) = AnArray(RowIndices(NormalizeIndex(RowIndices, ir)), _
                                                     NormalizeIndex(AnArray, c, 2))
                Next
            Next
        ' Get all elements requested
        Else
            Let RowIndices = First(IndicesCopy)
            Let ColumnIndices = Last(IndicesCopy)
        
            ReDim ReturnArray(1 To Length(RowIndices), 1 To Length(ColumnIndices))
            For ir = 1 To Length(RowIndices)
                For ic = 1 To Length(ColumnIndices)
                    Let ReturnArray(ir, ic) = AnArray(RowIndices(NormalizeIndex(RowIndices, ir)), _
                                                      ColumnIndices(NormalizeIndex(ColumnIndices, ic)))
                Next
            Next
        End If
    Else
        Let Part = Null
    End If
    
    Let Part = ReturnArray
End Function

' DESCRIPTION
' This function converts a valid index range specification --as determined by Predicates.PartIndexSequenceOfElementsQ--
' for function Arrays.Part into a sequence of individual indices. Suppose, the index has the form
' Array(n_1, n_2) = Array(1,5) for 1D array with LBound = 1. This function returns Array(1, 2, 3, 4, 5).
' This function performs no typechecking.
'
' PARAMETERS
' 1. AnArray - A dimensioned array
' 2. Indices - a sequence of indices (with at least one supplied) of the forms below, with each one
'    referring to a different dimension of the array. At the moment we process only 1D and 2D arrays.
'    So, Indices can only be one or two of the forms below.
' 3. TheDimension - The array's dimension relative to which the operations are perform
'
' RETURNED VALUE
' The requested sequence of indices
Private Function PartIntervalIndices(AnArray As Variant, _
                                     TheIndices As Variant, _
                                     Optional TheDimension As Long = 1, _
                                     Optional ParameterCheckQ As Boolean = True) As Variant
    Dim ni As Variant
    
    Let ni = NormalizeIndexArray(AnArray, TheIndices, TheDimension, ParameterCheckQ)
    
    If NullQ(First(ni)) Or NullQ(Last(ni)) Then
        Let PartIntervalIndices = Null
    ElseIf First(ni) > Last(ni) Then
        Let PartIntervalIndices = Null
    ElseIf First(ni) = Last(ni) Then
        Let PartIntervalIndices = Array(First(ni))
    Else
        Let PartIntervalIndices = CreateSequentialArray(CLng(First(ni)), CLng(Last(ni) - First(ni) + 1))
    End If
End Function

' DESCRIPTION
' This function converts a valid index range specification --as determined by Predicates.PartIndexSteppedSequenceQ--
' for function Arrays.Part into a sequence of individual indices. Suppose, the index has the form
' Array(n_1, n_2, step) = Array(1,5,2) for 1D array with LBound = 1. This function returns Array(1, 3, 5).
' This function performs no typechecking.
'
' PARAMETERS
' 1. AnArray - A dimensioned array
' 2. Indices - a sequence of indices (with at least one supplied) of the forms below, with each one
'    referring to a different dimension of the array. At the moment we process only 1D and 2D arrays.
'    So, Indices can only be one or two of the forms below.
' 3. TheDimension - The array's dimension relative to which the operations are perform
'
' RETURNED VALUE
' The requested sequence of indices
Private Function PartSteppedIntervalIndices(AnArray As Variant, _
                                            TheIndices As Variant, _
                                            TheDimension As Long, _
                                            Optional ParameterCheckQ As Boolean = True) As Variant
    Dim ni As Variant
    Dim FirstPos As Long
    Dim LastPos As Long
    Dim NumPos As Long
    Dim StepSize As Long
    
    Let ni = NormalizeIndexArray(AnArray, Most(TheIndices), TheDimension, ParameterCheckQ)
    Let FirstPos = CLng(First(ni))
    Let LastPos = CLng(Last(ni))
    Let StepSize = CLng(Last(TheIndices))
    
    Let FirstPos = CLng(First(ni))
    Let NumPos = CLng((LastPos - FirstPos - (LastPos - FirstPos) Mod StepSize) / StepSize + 1)
                                   
    Let PartSteppedIntervalIndices = CreateSequentialArray(FirstPos, NumPos, StepSize)
End Function

' DESCRIPTION
' Returns the subset of the 1D or 2D array specified by the indices.  Most common uses are:
'
' a. Take(m, n) - with n>0 returns the first n elements or rows of m
' b. Take(m, -n) - with n>0 returns the last n elements or rows of m
'
' PARAMETERS
' 1. AnArray - A dimensioned array
' 2. Indices - a sequence of indices (with at least one supplied) of the forms below, with each one
'    referring to a different dimension of the array. At the moment we process only 1D and 2D arrays.
'    So, Indices can only be one or two of the forms below.
'
' Indices can take any of the following forms:
' 1. n - Get elements 1 through
' 2. -n - Elements from the end of the array indexed by -1 to -n from right to left
' 2. [{n_1, n_2}] - Elements n_1 through n_2
'
' RETURNED VALUE
' The requested slice or element of the array.
Public Function Take(AnArray As Variant, ParamArray Indices() As Variant) As Variant
    Dim IndicesCopy As Variant
    
    ' We use our function to convert the ParamArray to a regular array
    Let IndicesCopy = CopyParamArray(Indices)

    ' Exit with Null if AnArray is neither an atomic array nor an atomic table
    If Not (AtomicArrayQ(AnArray) Or AtomicTableQ(AnArray)) Then
        Let Take = Null
        Exit Function
    End If

    ' Exit with Null if not even a single index was passed
    If Not DimensionedQ(IndicesCopy) Then
        Let Take = Null
        Exit Function
    End If
    
    ' Exit if more indices were passed than the number of dimensions in AnArray
    If NumberOfDimensions(AnArray) < Length(IndicesCopy) Then
        Let Take = Null
        Exit Function
    End If
    
    ' Exit with Null if any of the indices fails TakeIndexQ
    If Not TakeIndexArrayQ(IndicesCopy) Then
        Let Take = Null
        Exit Function
    End If
    
    ' Exit with Null if AnArray is empty because any index would be out of bounds
    If EmptyArrayQ(AnArray) Then
        Let Take = Null
        Exit Function
    End If
    
    ' Collect the chosen array's slice depending on its number of dimensions
    If Length(IndicesCopy) = 1 Then
        ' Process case on a single whole number given as index
        If PositiveWholeNumberQ(First(IndicesCopy)) Then
            Let Take = Part(AnArray, Array(1, First(IndicesCopy)))
        ElseIf NegativeWholeNumberQ(First(IndicesCopy)) Then
            Let Take = Part(AnArray, Array(First(IndicesCopy), -1))
        ElseIf WholeNumberArrayQ(First(IndicesCopy)) Then
            Let Take = Part(AnArray, First(IndicesCopy))
        Else
            Let Take = Null
        End If
    Else
        ' Process first dimensional index
        If PositiveWholeNumberQ(First(IndicesCopy)) Then
            Let First(IndicesCopy) = Array(1, First(IndicesCopy))
        ElseIf NegativeWholeNumberQ(First(IndicesCopy)) Then
            Let First(IndicesCopy) = Array(First(IndicesCopy), -1)
        ElseIf WholeNumberArrayQ(First(IndicesCopy)) Then
            Let First(IndicesCopy) = First(IndicesCopy)
        Else
            Let First(IndicesCopy) = Null
        End If
        
        ' Process last dimensional index
        If PositiveWholeNumberQ(Last(IndicesCopy)) Then
            Let Last(IndicesCopy) = Array(1, Last(IndicesCopy))
        ElseIf NegativeWholeNumberQ(Last(IndicesCopy)) Then
            Let Last(IndicesCopy) = Array(Last(IndicesCopy), -1)
        ElseIf WholeNumberArrayQ(Last(IndicesCopy)) Then
            Let Last(IndicesCopy) = Last(IndicesCopy)
        Else
            Let Last(IndicesCopy) = Null
        End If
        
        ' Exit with null if either index set is null
        If NullQ(First(IndicesCopy)) Or NullQ(Last(IndicesCopy)) Then
            Let Take = Null
        Else
            ' Call Part with the given dimensional index sets
            Let Take = Part(AnArray, First(IndicesCopy), Last(IndicesCopy))
        End If
    End If
End Function

' DESCRIPTION
' Returns the row with row index RowNumber as a 1D matrix if aMatrix satisfies Predicates.AtomicTableQ.
' We use the convention that 1 refers to the first row and -1 to the last.  The function returns Null
' if either aMatrix or RowNumber are invalid.
'
' PARAMETERS
' 1. aMatrix - any value or object reference
' 2. RowNumber - a non-zero integer smaller than or equal to the number of rows in aMatrix
'
' RETURNED VALUE
' Returns as a 1D array the row numbered RowNumber from the given 2D table
Public Function GetRow(aMatrix As Variant, RowNumber As Long, Optional ParameterCheckQ As Boolean = True) As Variant
    Dim c As Long
    Dim ResultArray() As Variant
    Dim NormalizedIndex As Variant
    
    ' This is here for high-speed applications
    If Not ParameterCheckQ Then
        ' Extract the requested row element by element
        ReDim ResultArray(1 To NumberOfColumns(aMatrix))
        For c = 1 To NumberOfColumns(aMatrix)
            Let ResultArray(c) = aMatrix(RowNumber, c)
        Next
        
        Let GetRow = ResultArray
        
        Exit Function
    End If
    
    ' Set default return case
    Let GetRow = Null
    
    ' Exit with Null is aMatrix is a 2D matrix
    If Not AtomicTableQ(aMatrix) Then Exit Function

    ' Exit with Null if aMatrix is either not dimensioned or null
    If EmptyArrayQ(aMatrix) Then Exit Function
    
    ' Exit with Null if RowNumber is not a nonzero wholenumber
    If Not NonzeroWholeNumberQ(RowNumber) Then Exit Function
    
    ' Exit with Null Abs(RowNumber) > NumberOfRows(aMatrix)
    If Abs(RowNumber) > NumberOfRows(aMatrix) Then Exit Function
    
    ' Simultaneously normalize the index and check it for consistency, exciting
    ' with Null in case of an error
    Let NormalizedIndex = NormalizeIndex(aMatrix, RowNumber)
    
    If NullQ(NormalizedIndex) Then Exit Function
    
    ' Extract the requested row element by element
    ReDim ResultArray(1 To NumberOfColumns(aMatrix))
    For c = 1 To NumberOfColumns(aMatrix)
        Let ResultArray(c) = aMatrix(NormalizedIndex, NormalizeIndex(aMatrix, c, 2))
    Next
    
    Let GetRow = ResultArray
End Function

' DESCRIPTION
' Returns the column with column index ColumnNumber as a 1D matrix if aMatrix satisfies Predicates.AtomicTableQ.
' We use the convention that 1 refers to the first column and -1 to the last.  The function returns Null
' if either aMatrix or ColumnNumber are invalid.
'
' PARAMETERS
' 1. aMatrix - any value or object reference
' 2. RowNumber - a non-zero integer smaller than or equal to the number of rows in aMatrix
'
' RETURNED VALUE
' Returns as a 1D array the column numbered ColumnNumber from the given 2D table
Public Function GetColumn(aMatrix As Variant, ColumnNumber As Long, Optional ParameterCheckQ As Boolean = True) As Variant
    Dim r As Long
    Dim ResultArray() As Variant
    Dim NormalizedIndex As Variant
    
    ' This is here for high-speed applications
    If Not ParameterCheckQ Then
        ' Extract the requested row element by element
        ReDim ResultArray(1 To NumberOfRows(aMatrix))
        For r = 1 To NumberOfColumns(aMatrix)
            Let ResultArray(r) = aMatrix(r, ColumnNumber)
        Next
        
        Let GetColumn = ResultArray
        
        Exit Function
    End If
    
    ' Set default return case
    Let GetColumn = Null
    
    ' Exit with Null is aMatrix is a 2D matrix
    If Not AtomicTableQ(aMatrix) Then Exit Function

    ' Exit with Null if aMatrix is either not dimensioned or null
    If EmptyArrayQ(aMatrix) Then Exit Function
    
    ' Exit with Null if RowNumber is not a nonzero wholenumber
    If Not NonzeroWholeNumberQ(ColumnNumber) Then Exit Function
    
    ' Exit with Null Abs(RowNumber) > NumberOfRows(aMatrix)
    If Abs(ColumnNumber) > NumberOfColumns(aMatrix) Then Exit Function
    
    ' Simultaneously normalize the index and check it for consistency, exciting
    ' with Null in case of an error
    Let NormalizedIndex = NormalizeIndex(aMatrix, ColumnNumber, 2)
    
    If NullQ(NormalizedIndex) Then Exit Function
    
    ' Extract the requested row element by element
    ReDim ResultArray(1 To NumberOfRows(aMatrix))
    For r = 1 To NumberOfRows(aMatrix)
        Let ResultArray(r) = aMatrix(NormalizeIndex(aMatrix, r, 1), NormalizedIndex)
    Next
    
    Let GetColumn = ResultArray
End Function

' DESCRIPTION
' Gets the sub-array of the given 1D array between StartIndex and EndIndex under the assumption
' that the array's 1st row has index 1.  If StartIndex or EndIndex are outside of the array's
' bounds, this function returns Null.  It also returns Null for an undimensioned or empty
' array or nonarrays
'
' PARAMETERS
' 1. AnArray - Any Excel object or reference
' 2. StartIndex - First index of requested subarray
' 3. EndIndex - Last index of requested subarray
' 4. ParameterCheckQ - (optional) When explicitly set to False, no parameter checks are done
'
' RETURNED VALUE
' Returns the given requested 1D subarray
Public Function GetSubArray(AnArray As Variant, _
                            StartIndex As Long, _
                            EndIndex As Long, _
                            Optional ParameterCheckQ As Boolean = True) As Variant
    Dim i As Long
    Dim ReturnedArray As Variant
    
    ' Set default return value
    Let GetSubArray = Null
    
    If ParameterCheckQ Then
        ' Exit with Null if AnArray is not dimensioned or empty
        If Not DimensionedQ(AnArray) Then Exit Function
        
        ' Exit with Null if AnArray is empty
        If EmptyArrayQ(AnArray) Then Exit Function
        
        ' Exit with Null if AnArray is not atomic
        If Not AtomicArrayQ(AnArray) Then Exit Function
        
        ' Exit with null if StartIndex or EndIndex are outside of the array's bounds
        If StartIndex < 1 Or StartIndex > Length(AnArray) Then Exit Function
        If EndIndex < 1 Or EndIndex > Length(AnArray) Then Exit Function
        
        ' Exit with Null if StartIndex>EndIndex
        If StartIndex > EndIndex Then Exit Function
        
        ' Exit with Null if this is not a 1D array
        If NumberOfDimensions(AnArray) <> 1 Then Exit Function
        
        ' Exit with Null if AnArray has LBound<>1
        If LBound(AnArray, 1) <> 1 Then Exit Function
    End If

    ReDim ReturnedArray(1 To EndIndex - StartIndex + 1)
    For i = StartIndex To EndIndex
        Let ReturnedArray(i - StartIndex + 1) = AnArray(i)
    Next i
    
    Let GetSubArray = ReturnedArray
End Function

' DESCRIPTION
' Returns the submatrix specified by the given endpoints.  Row and column endpoints are optional,
' but must come in pairs that are either both missing or both present.  The function returns Null
' if the given endpoints fall outside of the array's bounds. The function returns Null for all
' values of aMatrix that are not a dimensioned, non-empty atomic table.  The result is always returned
' as a 2D array.  No parameter consistency checks are done when ParameterCheckQ is explicitly passed
' as False.  All indices are assumed to start at 1.  The result is indexed starting with 1.
'
' PARAMETERS
' 1. aMatrix - Any Excel object or reference
' 2. TopEndPoint - First row of requested submatrix
' 3. BottomEndPoint - Last row of requested submatrix
' 4. LeftEndPoint - First column of requested submatrix
' 5. RightEndPoint - Last column of requested submatrix
' 4. ParameterCheckQ - (optional) When explicitly set to False, no parameter checks are done
'
' RETURNED VALUE
' Returns the given requested 2D submatrix.
Public Function GetSubMatrix(aMatrix As Variant, _
                             Optional TopEndPoint As Variant, _
                             Optional BottomEndPoint As Variant, _
                             Optional LeftEndPoint As Variant, _
                             Optional RightEndPoint As Variant, _
                             Optional ParameterCheckQ As Boolean = True) As Variant
    Dim TopEndPoint As Long
    Dim BottomEndPoint As Long
    Dim LeftEndPoint As Long
    Dim RightEndPoint As Long
    Dim r As Long
    Dim c As Long
    Dim ReturnMatrix As Variant
    
    ' Set default return value
    Let ParameterCheckQ = Null
                             
    If ParameterCheckQ Then
        ' Exit with Null if aMatrix is not dimensioned or empty
        If Not DimensionedQ(aMatrix) Then Exit Function
        
        ' Exit with Null if aMatrix is empty
        If EmptyArrayQ(aMatrix) Then Exit Function
        
        ' Exit with Null if aMatrix is not atomic
        If Not AtomicTableQ(aMatrix) Then Exit Function
        
        ' Exit with Null if TopEndPoint/BottomEndPoint or LeftEndPoint/RightEndPoint are not
        ' present in pairs
        If (Not IsMissing(TopEndPoint) And IsMissing(BottomEndPoint)) Or _
           (IsMissing(BottomEndPoint) And Not IsMissing(TopEndPoint)) Or _
           (IsMissing(LeftEndPoint) And Not IsMissing(RightEndPoint)) Or _
           (IsMissing(RightEndPoint) And Not IsMissing(LeftEndPoint)) Then Exit Function

        ' Exit with Null if either pair of endpoints are present a<b for the wrong endpoint
        If Not IsMissing(TopEndPoint) Then
            If Not PositiveWholeNumberQ(TopEndPoint) Then Exit Function
            If Not PositiveWholeNumberQ(BottomEndPoint) Then Exit Function
            If TopEndPoint > BottomEndPoint Then Exit Function
            If BottomEndPoint > NumberOfRows(aMatrix) Then Exit Function
        End If
        
        If Not IsMissing(LeftEndPoint) Then
            If Not PositiveWholeNumberQ(LeftEndPoint) Then Exit Function
            If Not PositiveWholeNumberQ(RightEndPoint) Then Exit Function
            If LeftEndPoint > RightEndPoint Then Exit Function
            If RightEndPoint > NumberOfColumns(aMatrix) Then Exit Function
        End If
        
        ' Exit with Null if aMatrix does not have to dimensions but LeftEndPoint is present
        If NumberOfDimensions(aMatrix) = 1 And Not IsMissing(LeftEndPoint) Then Exit Function
        
        ' Exit with Null if aMatrix has either LBound<>1
        If LBound(aMatrix, 1) <> 1 Or LBound(aMatrix, 2) <> 1 Then Exit Function
    End If
    
    If IsMissing(TopEndPoint) Then
        Let TopEndPoint = 1
        Let BottomEndPoint = NumberOfRows(aMatrix)
    End If
    
    If IsMissing(LeftEndPoint) Then
        Let LeftEndPoint = 1
        Let RightEndPoint = NumberOfColumns(aMatrix)
    End If
    
    ReDim ReturnMatrix(TopEndPoint To BottomEndPoint, LeftEndPoint To RightEndPoint)
    For r = 1 To BottomEndPoint - TopEndPoint + 1
        For c = 1 To RightEndPoint - LeftEndPoint + 1
            Let ReturnMatrix(r, c) = aMatrix(r + TopEndPoint - 1, c + LeftEndPoint - 1)
        Next
    Next
    
    Let GetSubMatrix = ReturnMatrix
End Function

' Stacks two arrays (may be 1 or 2-dimensional) on top of each other, provided they
' have the same number of columns. 1D arrays are allowed and interpreted as 1-row,
' 2D arrays. If either a or b is not an array, have dimensions > 2, or do not have
' the same number of columns, then this function returns an empty array (e.g. EmptyArray())
' The resulting arrays are indexed starting with 1
Public Function Stack2DArrays(ByVal a As Variant, ByVal b As Variant) As Variant
    Dim r As Long
    Dim c As Long
    Dim aprime As Variant
    Dim bprime As Variant
    Dim TheResult() As Variant
    
    If EmptyArrayQ(a) Or EmptyArrayQ(b) Or GetNumberOfColumns(a) <> GetNumberOfColumns(b) Then
        Let Stack2DArrays = EmptyArray()
        Exit Function
    End If
    
    ' If we have to 1D arrays of the same length, the we stack a on top of b
    If NumberOfDimensions(a) = 1 And NumberOfDimensions(b) = 1 Then
        ReDim TheResult(1 To 2, 1 To GetNumberOfColumns(a))
        
        For c = 1 To GetNumberOfColumns(a)
            Let TheResult(1, c) = a(IIf(LBound(a) = 0, c - 1, c))
            Let TheResult(2, c) = b(IIf(LBound(b) = 0, c - 1, c))
        Next c
    ElseIf NumberOfDimensions(a) = 1 And NumberOfDimensions(b) > 1 Then
        ReDim TheResult(1 To GetNumberOfRows(b) + 1, 1 To GetNumberOfColumns(b))
        
        For c = 1 To GetNumberOfColumns(a)
            Let TheResult(1, c) = a(IIf(LBound(a) = 0, c - 1, c))
        Next c
        
        For r = 1 To GetNumberOfRows(b)
            For c = 1 To GetNumberOfColumns(b)
                Let TheResult(1 + r, c) = b(IIf(LBound(b, 1) = 0, r - 1, r), IIf(LBound(b) = 0, c - 1, c))
            Next c
        Next r
    ElseIf NumberOfDimensions(a) > 1 And NumberOfDimensions(b) = 1 Then
        ReDim TheResult(1 To GetNumberOfRows(a) + 1, 1 To GetNumberOfColumns(a))
        
        For r = 1 To GetNumberOfRows(a) + 1
            For c = 1 To GetNumberOfColumns(a)
                If r < GetNumberOfRows(a) + 1 Then
                    Let TheResult(r, c) = a(IIf(LBound(a, 1) = 0, r - 1, r), IIf(LBound(a, 2) = 0, c - 1, c))
                Else
                    Let TheResult(GetNumberOfRows(a) + 1, c) = b(IIf(LBound(b) = 0, c - 1, c))
                End If
            Next c
        Next r
    Else
        ReDim TheResult(1 To GetNumberOfRows(a) + GetNumberOfRows(b), 1 To GetNumberOfColumns(b))
        For r = 1 To GetNumberOfRows(a)
            For c = 1 To GetNumberOfColumns(a)
                Let TheResult(r, c) = a(IIf(LBound(a, 1) = 0, r - 1, r), IIf(LBound(a, 2) = 0, c - 1, c))
            Next c
        Next r
    
        For r = 1 To GetNumberOfRows(b)
            For c = 1 To GetNumberOfColumns(b)
                Let TheResult(GetNumberOfRows(a) + r, c) = b(IIf(LBound(b, 1) = 0, r - 1, r), IIf(LBound(b, 2) = 0, c - 1, c))
            Next c
        Next r
    End If
    
    Let Stack2DArrays = TheResult
End Function

' DESCRIPTION
' Prints the given 1D or 2D printable table to a string and to the debug window.
' If TheArray is not a printable array, the function returns "Not An Array"
' If TheArray an empty array, the function returns "Empty 1D Array"
'
' PARAMETERS
' 1. TheArray - Any Excel object or reference
'
' RETURNED VALUE
' Returns the given 1D or 2D table as a string and prints it in the debug window.
Public Function PrintArray(TheArray As Variant) As String
    Dim ReturnString As String
    Dim ARow As Variant
    Dim c As Long
    Dim r As Long
    
    If Not IsArray(TheArray) Then
        Let ReturnString = "Not an array"
        Debug.Print ReturnString
    ElseIf NumberOfDimensions(TheArray) = 0 Then
        Let ReturnString = TheArray
        Debug.Print ReturnString
    ElseIf NumberOfDimensions(TheArray) = 1 Then
        If EmptyArrayQ(TheArray) Then
            Let ReturnString = "Empty 1D Array"
            Debug.Print ReturnString
        Else
            Let ARow = TheArray(LBound(TheArray))
            Let ReturnString = ReturnString & vbCr & ARow
        
            If UBound(TheArray) - LBound(TheArray) >= 1 Then
                For c = LBound(TheArray) + 1 To UBound(TheArray)
                    Let ARow = ARow & vbTab & TheArray(c)
                Next c
                
                Let ReturnString = ReturnString & vbCr & ARow
            End If
        
            Debug.Print ARow
        End If
    Else
        For r = LBound(TheArray, 1) To UBound(TheArray, 1)
            Let ARow = TheArray(r, LBound(TheArray, 2))
        
            If UBound(TheArray, 2) - LBound(TheArray, 2) >= 1 Then
                For c = LBound(TheArray, 2) + 1 To UBound(TheArray, 2)
                    Let ARow = ARow & vbTab & TheArray(r, c)
                Next c
                
                Let ReturnString = ReturnString & vbCr & ARow
            End If
        
            Debug.Print ARow
        Next r
    End If
    
    Let PrintArray = ReturnString
End Function

' This function inserts either an atomic element into a 1D array satisfying RowVectorQ or
' a row into a 2D array filled exclusively with atomic elements
' AnArray is returned unevaluated if either AnArray, TheElt, or ThePos is not as expected.
' To insert in AnArray's first position, set ThePos = LBound(AnArray)
' To insert in AnArray's last position, set ThePos = UBound(AnArray)+1
' You should think of acceptable values for ThePos as insert TheElt at ThePos and
' shifting to the right anything in AnArray at and to the right of ThePos.
Public Function Insert(AnArray As Variant, TheElt As Variant, ThePos As Long) As Variant
    Dim FirstPart As Variant
    Dim LastPart As Variant
    
    If Not DimensionedQ(AnArray) Then
        Let Insert = Null
        Exit Function
    End If
    
    If IsNull(TheElt) Then
        Let Insert = Null
        Exit Function
    End If

    ' Exit returning AnArray unevaluated if neither AnArray nor ThePos make sense
    If Not RowVectorQ(AnArray) And Not MatrixQ(AnArray) Or Not IsNumeric(ThePos) Then
        Let Insert = Null
        Exit Function
    End If
    
    ' Exit if AnArray is a row vector TheElt is not an atomic expression
    If RowVectorQ(AnArray) And IsArray(TheElt) Then
        Let Insert = Null
        Exit Function
    End If
    
    ' Exit if AnArray is a matrix and TheElt is not a row vector with the same number of columns
    If MatrixQ(AnArray) And (Not RowVectorQ(TheElt) Or GetNumberOfColumns(AnArray) <> GetArrayLength(TheElt)) Then
        Let Insert = Null
        Exit Function
    End If

    ' Exit if ThePos has unaceptable values
    If ThePos < LBound(AnArray) Or ThePos > UBound(AnArray) + 1 Then
        Let Insert = Null
        Exit Function
    End If
    
    ' Based on ThePos, shift parts of AnArray and insert TheElt
    If ThePos = LBound(AnArray, 1) Then
        If NumberOfDimensions(AnArray) = 1 Then
            Let Insert = Prepend(AnArray, TheElt)
        Else
            Let Insert = Stack2DArrays(TheElt, AnArray)
        End If
    ElseIf ThePos = UBound(AnArray, 1) + 1 Then
        If NumberOfDimensions(AnArray) = 1 Then
            Let Insert = Append(AnArray, TheElt)
        Else
            Let Insert = Stack2DArrays(AnArray, TheElt)
        End If
    ElseIf ThePos > LBound(AnArray, 1) And ThePos <= UBound(AnArray, 1) Then
        Let FirstPart = Take(AnArray, ThePos - 1)
        Let LastPart = Take(AnArray, -(UBound(AnArray, 1) - ThePos + 1))
        
        If NumberOfDimensions(AnArray) = 1 Then
            Let FirstPart = Append(FirstPart, TheElt)
            Let Insert = ConcatenateArrays(FirstPart, LastPart)
        Else
            Let FirstPart = Stack2DArrays(FirstPart, TheElt)
            Let Insert = Stack2DArrays(FirstPart, LastPart)
        End If
    Else
        Let Insert = AnArray
    End If
End Function

' This function turns a 1D array of 1D arrays into a 2D array.
' Each of the elemements (inner arrays) of the outermost array satisfies
' RowVectorQ
'
' This is useful to quickly build a matrix from 1D arrays
' This function assumes that all elements of TheRowsAs1DArrays have the same lbound()
' arg is allowed to have different lbound from that of its elements
'
' The 2D array returned is indexed starting at 1
' If the optional parameter PackAsColumnsQ is set to True, the 1D arrays in TheRowsAs1DArrays become columns.
Public Function Pack2DArray(TheRowsAs1DArrays As Variant, Optional PackAsColumnsQ As Boolean = False) As Variant
    Dim var As Variant
    Dim r As Long
    Dim c As Long
    Dim Results() As Variant
    Dim TheLength As Long
    
    ' Exit if the argument is not the expected type
    If NumberOfDimensions(TheRowsAs1DArrays) <> 1 Or EmptyArrayQ(TheRowsAs1DArrays) Then
        Let Pack2DArray = Null
        Exit Function
    End If
    
    ' Exit if any of the elements in not an atomic array or
    '  if all the array elements do not the same length
    Let TheLength = GetArrayLength(First(TheRowsAs1DArrays))
    For Each var In TheRowsAs1DArrays
        If Not AtomicArrayQ(var) Or GetArrayLength(var) <> TheLength Then
            Let Pack2DArray = Null
            Exit Function
        End If
    Next

    ' Pre-allocate a 2D array filled with Empty
    ReDim Results(1 To GetArrayLength(TheRowsAs1DArrays), 1 To GetArrayLength(First(TheRowsAs1DArrays)))
    
    ' Pack the array
    For r = LBound(TheRowsAs1DArrays) To UBound(TheRowsAs1DArrays)
        For c = LBound(First(TheRowsAs1DArrays)) To UBound(First(TheRowsAs1DArrays))
            Let Results(IIf(LBound(TheRowsAs1DArrays) = 0, 1, 0) + r, IIf(LBound(First(TheRowsAs1DArrays)) = 0, 1, 0) + c) = TheRowsAs1DArrays(r)(c)
        Next c
    Next r
    
    If PackAsColumnsQ Then
        Let Pack2DArray = TransposeMatrix(Results)
        Exit Function
    End If
    
    Let Pack2DArray = Results
End Function

' This function is the exact opposite of Take.  It removes what take would return
Public Function Drop(AnArray As Variant, N As Variant) As Variant
    Dim DualIndex As Variant
    Dim RenormalizedIndices As Variant
    Dim var As Variant
    Dim c As Long

    ' Exit with argument unchanged if the array has fewer or more than 2 dimensions
    If NumberOfDimensions(AnArray) = 0 Or NumberOfDimensions(AnArray) > 2 Then
        Let Drop = AnArray

        Exit Function
    End If

    ' Exit with argument unchanged if either N is not an integer or an array of integers
    If Not (IsNumeric(N) Or IsNumericArrayQ(N)) Or EmptyArrayQ(N) Or EmptyArrayQ(AnArray) Then
        Let Drop = EmptyArray()
        
        Exit Function
    End If
    
    If IsNumericArrayQ(N) Then
        For Each var In N
            If CLng(var) <> var Then
                Let Drop = EmptyArray()
                
                Exit Function
            End If
        Next
    End If
    
    ' Case of N an integer
    If NumberOfDimensions(N) = 0 Then
        If NumberOfDimensions(AnArray) = 1 Then
            If N > 0 Then
                Let DualIndex = IIf(N > GetArrayLength(AnArray), EmptyArray(), -(GetArrayLength(AnArray) - N))
            Else
                Let DualIndex = IIf(Abs(N) > GetArrayLength(AnArray), EmptyArray(), GetArrayLength(AnArray) + N)
            End If
        Else
            If N > 0 Then
                Let DualIndex = IIf(N > GetNumberOfRows(AnArray), EmptyArray(), -(GetNumberOfRows(AnArray) - N))
            Else
                Let DualIndex = IIf(Abs(N) > GetNumberOfRows(AnArray), EmptyArray(), GetNumberOfRows(AnArray) + N)
            End If
        End If
    
        Let Drop = Take(AnArray, DualIndex)
    
        Exit Function
    End If
    
    ' Proceed with case of N a 1D array of integers
    
    ' Turn all indices in N into their positive equivalents
    Let c = 1
    Let RenormalizedIndices = ConstantArray(Empty, GetArrayLength(N))
    For Each var In N
        Let RenormalizedIndices(c) = IIf(var < 0, IIf(NumberOfDimensions(AnArray) = 1, GetNumberOfColumns(AnArray), GetNumberOfRows(AnArray)) + var + 1, var)
        Let c = c + 1
    Next
    
    Let Drop = Take(AnArray, ComplementOfSets(CreateSequentialArray(1, GetArrayLength(AnArray)), RenormalizedIndices))
End Function

Public Function Convert1DArrayIntoParentheticalExpression(TheArray As Variant) As String
    Let Convert1DArrayIntoParentheticalExpression = "(" & Join(TheArray, ",") & ")"
End Function

' This function sorts the given 2D matrix by the columns whose positions are given by
' ArrayOfColPos. The sorting orientation in each column are in ArrayOfColsSortOrder
' ArrayOfColsSortOrder is a variant array whose elements are all of enumerated type XLSortOrder
' (e.g. xlAscending, xlDescending)
Public Function Sort2DArray(MyArray As Variant, ArrayOfColPos As Variant, _
                     ArrayOfColsSortOrder As Variant, _
                     WithHeaders As XlYesNoGuess) As Variant
    Dim TheRange As Range
    Dim TmpSheet As Worksheet
    Dim i As Integer
    
    ' Set pointer to temp sheet and clear its used range
    Set TmpSheet = ThisWorkbook.Worksheets("TempComputation")
    TmpSheet.UsedRange.ClearContents
    
    ' Dump array in temp sheet
    'Let TmpSheet.Range("A1").Resize(UBound(MyArray, 1), UBound(MyArray, 2)).Value2 = MyArray
    Call DumpInSheet(MyArray, TempComputation.Range("A1"), True)

    ' Set range pointer to the data we just dumped in the temp sheet
    Set TheRange = TmpSheet.Range("A1").Resize(GetNumberOfRows(MyArray), GetNumberOfColumns(MyArray))

    ' Clear any previous sorting criteria
    TmpSheet.Sort.SortFields.Clear

    ' Add all the sorting criteria
    For i = LBound(ArrayOfColPos) To UBound(ArrayOfColPos)
        ' Add criteria to sort by date
        TheRange.Worksheet.Sort.SortFields.Add _
            Key:=TheRange.Columns(ArrayOfColPos(i)), _
            SortOn:=xlSortOnValues, _
            Order:=ArrayOfColsSortOrder(i), _
            DataOption:=xlSortNormal
    Next i
    
    ' Execute the sort
    With TheRange.Worksheet.Sort
        .SetRange TheRange
        .Header = WithHeaders
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    ' Extract horizontally reversed matrix and set them to return when the function exits
    Let Sort2DArray = TmpSheet.Range("A1").CurrentRegion.Value2
End Function

' This function reverses the horizontal order of an array (1D or 2D)
' It assumes that the first column must remain unchanged.
'***HERE Implement using loops
Public Function ReverseHorizontally(MyArray As Variant) As Variant
    Dim TheRange As Range
    Dim TmpSheet As Worksheet

    ' Set pointer to temp sheet and clear its used range
    Set TmpSheet = ThisWorkbook.Worksheets("TempComputation")
    TmpSheet.UsedRange.ClearContents
    
    ' Dump array in temp sheet
    Let TmpSheet.Range("A1").Resize(UBound(MyArray, 1), UBound(MyArray, 2)).Value2 = MyArray

    ' Insert an empty row at the top where we will place the horizontal column counter
    TmpSheet.Range("1:1").Insert
    
    ' Insert column counter in first row
    Let TmpSheet.Range("A1").Resize(1, TmpSheet.Range("A2").CurrentRegion.Columns.Count).Formula = "=column()"
    
    ' Change the column index of the first column to be larger than the number of columns so it does not change
    ' position
    Let TmpSheet.Range("A1").Value2 = Application.Max(TmpSheet.Range(Range("B1"), Range("B1").End(xlToRight))) + 1
    
    ' Copy and paste column Indices as values
    TmpSheet.Range(Range("A1"), Range("A1").End(xlToRight)).Value2 = TmpSheet.Range(Range("A1"), Range("A1").End(xlToRight)).Value2
    
    ' Set range pointer to the data we just dumped in the temp sheet
    Set TheRange = TmpSheet.Range("A1").CurrentRegion

    ' Clear any previous sorting criteria
    TmpSheet.Sort.SortFields.Clear

    ' Add criteria to sort by date
    TheRange.Worksheet.Sort.SortFields.Add _
        Key:=TheRange.Rows(1), _
        SortOn:=xlSortOnValues, _
        Order:=xlDescending, _
        DataOption:=xlSortNormal
        
    ' Execute the sort
    With TheRange.Worksheet.Sort
        .SetRange TheRange
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlLeftToRight
        .SortMethod = xlPinYin
        .Apply
    End With
    
    ' Delete the header row
    TmpSheet.Range("1:1").Delete
    
    ' Extract horizontally reversed matrix and set them to return when the function exits
    Let ReverseHorizontally = TmpSheet.Range("A1").CurrentRegion.Value2
End Function

' This function reverses the horizontal order of an array (1D or 2D)
' It assumes that the first column must remain unchanged.
'***HERE Implement using loops
Public Function ReverseVertically(MyArray As Variant) As Variant
    Dim TheRange As Range
    Dim TmpSheet As Worksheet
    
    ' Set pointer to temp sheet and clear its used range
    Set TmpSheet = ThisWorkbook.Worksheets("TempComputation")
    TmpSheet.UsedRange.ClearContents
    
    ' Dump array in temp sheet
    Let TmpSheet.Range("A1").Resize(UBound(MyArray, 1), UBound(MyArray, 2)).Value2 = MyArray

    ' Insert an empty row at the top where we will place the horizontal column counter
    TmpSheet.Range("A:A").Insert
    
    ' Insert column counter in first row
    Let TmpSheet.Range("A1").Resize(TmpSheet.Range("A2").CurrentRegion.Rows.Count, 1).Formula = "=row()"
    
    ' Change the row index of the first row to be larger than the number of row so it does not change
    ' position
    Let TmpSheet.Range("A1").Value2 = Application.Max(TmpSheet.Range(TmpSheet.Range("A2"), TmpSheet.Range("A2").End(xlDown))) + 1
    
    ' Copy and paste column Indices as values
    Let TmpSheet.Range("A1").CurrentRegion.Columns(1).Value2 = TmpSheet.Range("A1").CurrentRegion.Columns(1).Value2
    
    ' Set range pointer to the data we just dumped in the temp sheet
    Set TheRange = TmpSheet.Range("A1").CurrentRegion

    ' Clear any previous sorting criteria
    TmpSheet.Sort.SortFields.Clear

    ' Add criteria to sort by date
    TheRange.Worksheet.Sort.SortFields.Add _
        Key:=TheRange.Columns(1), _
        SortOn:=xlSortOnValues, _
        Order:=xlDescending, _
        DataOption:=xlSortNormal
        
    ' Execute the sort
    With TheRange.Worksheet.Sort
        .SetRange TheRange
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    ' Delete the header row
    TmpSheet.Range("A:A").Delete
    
    ' Extract horizontally reversed matrix and set them to return when the function exits
    Let ReverseVertically = TmpSheet.Range("A1").CurrentRegion.Value2
End Function

' Appends a new element to the given array handles 1D and 2D arrays. Returns Null
' if the parameters are inconsistent. In the case of AnArray being a 2D array, AnElt
' must be a 1D array with the same number of colummns as AnArray for Append to make sense.
'
' If AnArray is a 1D or 2D array and AnElt is Null, the funtion returns AnArray unchanged.
'
' This works differently from simply using Stack2DArrays(AnArray, AnElt) since it can return
' something that is not a matrix.  Stack2DArrays ALWAYS returns a matrix
Public Function Append(AnArray As Variant, AnElt As Variant) As Variant
    Dim NewArray As Variant
    Dim AnArrayNumberOfDims As Integer
    
    Let AnArrayNumberOfDims = NumberOfDimensions(AnArray)
    
    If Not IsArray(AnArray) Or AnArrayNumberOfDims > 2 Or _
       (AnArrayNumberOfDims = 2 And GetNumberOfColumns(AnArray) <> GetNumberOfColumns(AnElt)) Then
        Let Append = Null
    
        Exit Function
    End If
    
    If IsNull(AnElt) Then
        Let Append = AnArray
        Exit Function
    End If
    
    ' If AnArray is 1D, then put the new element (whatever it may be) as the last element of a
    ' 1D array 1 longer than the original one.
    If NumberOfDimensions(AnArray) = 1 Then
        Let NewArray = AnArray
        ReDim Preserve NewArray(LBound(AnArray) To UBound(AnArray) + 1)
        Let NewArray(UBound(NewArray)) = AnElt
        
        Let Append = NewArray
        
        Exit Function
    End If
    
    ' If AnArray has two dims and the same number of columns as AnElt, then stack AnElt as the bottom
    ' of AnArray
    Let Append = Stack2DArrays(AnArray, AnElt)
End Function

' Appends a new element to the given array. Handles arrays of dimension 1 and 2
' Returns Null if AnArray is not an array or it has more than two dims,
' or dim(AnArray) = 2 and AnArray and AnElt don't have the same number of columns.
' If AnElt is Null, the function returns AnArray unevaluated.
'
' This works differently from simply using Stack2DArrays(AnArray, AnElt)
' This one can give you something that is not a matrix.  Stack2DArrays ALWAYS returns
' a matrix
Public Function Prepend(AnArray As Variant, AnElt As Variant) As Variant
    Dim NewEmptyArray() As Variant
    Dim AnArrayNumberOfDims As Integer
    Dim i As Long
    
    Let AnArrayNumberOfDims = NumberOfDimensions(AnArray)
    
    If Not IsArray(AnArray) Or AnArrayNumberOfDims > 2 Or _
       (AnArrayNumberOfDims = 2 And GetNumberOfColumns(AnArray) <> GetNumberOfColumns(AnElt)) Then
        Let Prepend = Null
    
        Exit Function
    End If
    
    If IsNull(AnElt) Then
        Let Prepend = AnArray
        Exit Function
    End If
    
    ' If AnArray is 1D, then put the new element (whatever it may be) as the last element of a
    ' 1D array 1 longer than the original one.
    If NumberOfDimensions(AnArray) = 1 Then
        ReDim NewArray(LBound(AnArray) To UBound(AnArray) + 1)
        
        For i = LBound(AnArray) + 1 To UBound(AnArray) + 1
            Let NewArray(i) = AnArray(i - 1)
        Next i
        
        Let NewArray(LBound(AnArray)) = AnElt
        
        Let Prepend = NewArray
        
        Exit Function
    End If
    
    ' If AnArray has two dims and the same number of columns as AnElt, then stack AnElt as the bottom
    ' of AnArray
    Let Prepend = Stack2DArrays(AnElt, AnArray)
End Function

' aMatrix is a 1D or 2D array of values (not a range).
' Returns a 1-dimensional array with the unique subset of elements.
' If aMatrix has more than two dimensions, the function returns Null
Public Function UniqueSubset(aMatrix As Variant) As Variant
    Dim r As Long
    Dim c As Long
    Dim UniqueDict As Dictionary
    Dim Dimensionality As Integer
    
    ' Exit, returning an empty array if aMatrix is empty
    If EmptyArrayQ(aMatrix) Or IsEmpty(aMatrix) Then
        Let UniqueSubset = EmptyArray()
        Exit Function
    End If
    
    Set UniqueDict = New Dictionary
    Let Dimensionality = NumberOfDimensions(aMatrix)
    
    If Dimensionality = 0 Then
        Let UniqueSubset = aMatrix
        
        Exit Function
    ElseIf Dimensionality = 1 Then
        For r = LBound(aMatrix) To UBound(aMatrix)
            If Not UniqueDict.Exists(aMatrix(r)) Then
                Call UniqueDict.Add(Key:=aMatrix(r), Item:=1)
            End If
        Next r
    ElseIf Dimensionality = 2 Then
        For r = LBound(aMatrix, 1) To UBound(aMatrix, 1)
            For c = LBound(aMatrix, 2) To UBound(aMatrix, 2)
                If Not UniqueDict.Exists(aMatrix(r, c)) Then
                    Call UniqueDict.Add(Key:=aMatrix(r, c), Item:=1)
                End If
            Next c
        Next r
    Else
        Let UniqueSubset = Null
        
        Exit Function
    End If
    
    Let UniqueSubset = UniqueDict.Keys
End Function
' Returns the unique set of values contained in the two 1D arrays. The union of Set1 and Set2 is returned as a 1D array.
' Set1 and Set2 must have no empty rows of columns.
Public Function UnionOfSets(Set1 As Variant, Set2 As Variant) As Variant
    Dim FirstSet As Variant
    Dim SecondSet As Variant
    Dim CombinedSet As Dictionary
    Dim i As Long

    Set CombinedSet = New Dictionary
    
    If (EmptyArrayQ(Set1) Or IsEmpty(Set1)) And (EmptyArrayQ(Set2) Or IsEmpty(Set2)) Then
        Let UnionOfSets = EmptyArray()
        Exit Function
    ElseIf (EmptyArrayQ(Set1) Or IsEmpty(Set1)) And Not EmptyArrayQ(Set2) Then
        Let UnionOfSets = UniqueSubset(Set2)
        Exit Function
    ElseIf Not EmptyArrayQ(Set1) And (EmptyArrayQ(Set2) Or IsEmpty(Set2)) Then
        Let UnionOfSets = UniqueSubset(Set1)
        Exit Function
    End If

    ' Stack the two sets on top of each other in the worksheet TempComputation
    Let FirstSet = UniqueSubset(Set1)
    Let SecondSet = UniqueSubset(Set2)
    
    For i = LBound(FirstSet) To UBound(FirstSet)
        If Not CombinedSet.Exists(FirstSet(i)) Then
            Call CombinedSet.Add(Key:=FirstSet(i), Item:=i)
        End If
    Next i
    
    For i = LBound(SecondSet) To UBound(SecondSet)
        If Not CombinedSet.Exists(SecondSet(i)) Then
            Call CombinedSet.Add(Key:=SecondSet(i), Item:=i)
        End If
    Next i
    
    Let UnionOfSets = CombinedSet.Keys
End Function

' This function takes two parameters.  Each could be either a 1D array or a 2D array.
' This function returns the intersection of the two arrays.
Public Function IntersectionOfSets(Set1 As Variant, Set2 As Variant) As Variant
    Dim FirstDict As Dictionary
    Dim First1DSet As Variant
    Dim Second1DSet As Variant
    Dim IntersectionDict As Dictionary
    Dim i As Long
    
    ' Exit returning an empty array if either set is empty
    If EmptyArrayQ(Set1) Or IsEmpty(Set1) Or EmptyArrayQ(Set2) Or IsEmpty(Set2) Then
        Let IntersectionOfSets = EmptyArray()
        Exit Function
    End If
    
    ' Instantiate dictionaries
    Set FirstDict = New Dictionary
    Set IntersectionDict = New Dictionary
    
    ' Convert each set to a 1D array
    Let First1DSet = Stack2DArrayAs1DArray(Set1)
    Let Second1DSet = Stack2DArrayAs1DArray(Set2)
    
    ' Load a dictionary with the elements of the first set
    For i = LBound(First1DSet) To UBound(First1DSet)
        Call FirstDict.Add(Key:=First1DSet(i), Item:=1)
    Next i
    
    ' Store the elements of the second set that are in the first
    For i = LBound(Second1DSet) To UBound(Second1DSet)
        If FirstDict.Exists(Key:=Second1DSet(i)) Then
            Call IntersectionDict.Add(Key:=Second1DSet(i), Item:=1)
        End If
    Next i
    
    ' Return the intesection of the two sets
    Let IntersectionOfSets = IntersectionDict.Keys
End Function

' This function returns the complement of set B in A
' Both A and B are required to be 1D arrays
' If the complement is empty, this function returns an empty array (e.g. EmptyArray())
Public Function ComplementOfSets(a As Variant, b As Variant) As Variant
    Dim BDict As Dictionary
    Dim ComplementDict As Dictionary
    Dim obj As Variant
    
    ' If a is an empty array, exit returning an empty array
    If EmptyArrayQ(a) Or IsEmpty(a) Then
        Let ComplementOfSets = EmptyArray()
        Exit Function
    End If
    
    If EmptyArrayQ(b) Or IsEmpty(b) Then
        Let ComplementOfSets = a
        Exit Function
    End If
    
    If NumberOfDimensions(a) < 1 Or NumberOfDimensions(b) < 1 Then
        Let ComplementOfSets = EmptyArray()
        
        Exit Function
    End If
    
    ' Instantiate dictionaries
    Set BDict = New Dictionary
    Set ComplementDict = New Dictionary
    
    ' Initialize ADict to get unique subset of ADict
    If GetArrayLength(b) > 0 Then
        For Each obj In b
            If Not BDict.Exists(Key:=obj) Then
                Call BDict.Add(Key:=obj, Item:=obj)
            End If
        Next
    End If
    
    ' Populate ComplementDict
    If GetArrayLength(a) > 0 Then
        For Each obj In a
            If Not BDict.Exists(Key:=obj) And Not ComplementDict.Exists(Key:=obj) Then
                Call ComplementDict.Add(Key:=obj, Item:=obj)
            End If
        Next
    End If
    
    ' Return complement as 1D array
    If ComplementDict.Count = 0 Then
        Let ComplementOfSets = EmptyArray()
    Else
        Let ComplementOfSets = ComplementDict.Keys
    End If
End Function

' This function takes a 1D or 2D array (NOT a range) and turns it into a 1D array with the same list of
' elements. It returns a 1D array.  Row 2 is appended to row 1.  Row 3 is then appended to that, etc.
Public Function Stack2DArrayAs1DArray(aMatrix As Variant) As Variant
    Dim TheResults() As Variant
    Dim var As Variant
    Dim j As Long

    If EmptyArrayQ(aMatrix) Or Not IsArray(aMatrix) Or NumberOfDimensions(aMatrix) > 2 Then
        Let Stack2DArrayAs1DArray = EmptyArray()
        Exit Function
    ElseIf NumberOfDimensions(aMatrix) = 0 Then
        Let Stack2DArrayAs1DArray = aMatrix
        Exit Function
    End If
    
    If NumberOfDimensions(aMatrix) = 1 Then
        ReDim TheResults(GetArrayLength(aMatrix))
    Else
        ReDim TheResults(GetNumberOfRows(aMatrix) * GetNumberOfColumns(aMatrix))
    End If
    
    Let j = 1
    For Each var In aMatrix
        Let TheResults(j) = var
        Let j = j + 1
    Next
    
    Let Stack2DArrayAs1DArray = TheResults
End Function

' Alias for Stack2DArrayAs1DArray
Public Function StackArrayAs1DArray(aMatrix As Variant) As Variant
    Let StackArrayAs1DArray = Stack2DArrayAs1DArray(aMatrix)
End Function

' This function dumps an array (1D or 2D) into worksheet TempComputation and then returns a reference to
' the underlying range.  Worksheet TempComputation is cleared before dumping.  Dimensions are preserved.
' This means that an m x n array is dumped into an m x n range.  This function should not be used if
' leading single quotes (e.g "'") are part of the array's elements.
Public Function ToTemp(AnArray As Variant, Optional PreserveColumnTextFormats As Boolean = False) As Range
    Call TempComputation.UsedRange.ClearFormats
    Call TempComputation.UsedRange.ClearContents

    Set ToTemp = DumpInSheet(AnArray, TempComputation.Range("A1"), PreserveColumnTextFormats)
End Function

' This function dumps an array (1D or 2D) into the worksheet with the range referenced by TopLeftCorner as the cell in the upper-left
' corner. It then returns a reference to the underlying range. Dimensions are preserved.  This means that an m x n array is dumped into
' an m x n range.
'
' This function is a helper for Arrays.DumpInSheet()
Private Function DumpInSheetHelper(AnArray As Variant, TopLeftCorner As Range, Optional PreserveColumnTextFormats As Boolean = False) As Range
    Dim c As Integer
    Dim NumberOfRows As Long
    Dim NumberOfColumns As Integer
    Dim NumDimensions As Integer

    Let NumDimensions = NumberOfDimensions(AnArray)
    Let NumberOfRows = GetNumberOfRows(AnArray)
    Let NumberOfColumns = GetNumberOfColumns(AnArray)
    
    If PreserveColumnTextFormats Then
        If NumDimensions = 0 Then
            Let TempComputation.Range("A1").NumberFormat = IIf(TypeName(AnArray) = "String", "@", "0")
            Let TempComputation.Range("A1").Value2 = AnArray
        Else
            ' Loop over the target columns, applying the format of the array's first column element to
            ' the entire, corresponding target range column
            For c = 0 To NumberOfColumns - 1
                If NumDimensions = 1 Then
                    Let TopLeftCorner.Offset(0, c).Resize(NumberOfRows, 1).NumberFormat = IIf(TypeName(AnArray(c + LBound(AnArray))) = "String", "@", "0")
                Else
                    Let TopLeftCorner.Offset(0, c).Resize(NumberOfRows, 1).NumberFormat = IIf(TypeName(AnArray(1, c + LBound(AnArray, 2))) = "String", "@", "0")
                End If
            Next c
        End If
    End If

    If NumDimensions = 1 Then
        Let TopLeftCorner(1, 1).Resize(1, NumberOfColumns).Value2 = AnArray
        Set DumpInSheetHelper = TopLeftCorner(1, 1).Resize(1, NumberOfColumns)
    Else
        Let TopLeftCorner(1, 1).Resize(NumberOfRows, NumberOfColumns).Value2 = AnArray
        Set DumpInSheetHelper = TopLeftCorner(1, 1).Resize(NumberOfRows, NumberOfColumns)
    End If
End Function

' This is an alias for function DumpInSheetHelper above
' This one handles empty arrays correctly, by doing nothing and returning Null
Public Function DumpInSheet(AnArray As Variant, _
                            TopLeftCorner As Range, _
                            Optional PreserveColumnTextFormats As Boolean = False) As Range
    If IsNull(AnArray) Then
        Set DumpInSheet = Nothing
        Exit Function
    End If
    
    If TopLeftCorner Is Nothing Then
        Set DumpInSheet = Nothing
        Exit Function
    End If
    
    If StringQ(AnArray) Or IsNumeric(AnArray) Or IsDate(AnArray) Then
        Let TopLeftCorner.Value2 = AnArray
        Set DumpInSheet = TopLeftCorner
        Exit Function
    End If

    If Not DimensionedQ(AnArray) Then
        Set DumpInSheet = Nothing
        Exit Function
    End If

    If EmptyArrayQ(AnArray) Or IsEmpty(AnArray) Then
        Set DumpInSheet = Nothing
        Exit Function
    End If
    
    If GetArrayLength(AnArray) = 0 Then
        Set DumpInSheet = Nothing
        Exit Function
    End If

    If PreserveColumnTextFormats Then
        Set DumpInSheet = DumpInSheetHelper(AnArray, TopLeftCorner, PreserveColumnTextFormats:=PreserveColumnTextFormats)
    Else
        Set DumpInSheet = DumpInSheetHelper(AnArray, TopLeftCorner)
    End If
End Function

' Applies the LAG operator to the power N to TheArray. TheArray has to be a 1D or 2D Array containing timeseries data.
' Timedimension is a parameter indicating whether the timeseries are organised from left to right (horizontal) or from
' top to bottom (vertical)
Public Function LAGN(TheArray As Variant, N As Variant, TimeDimension As String)
    If TimeDimension = "Vertical" Then
        LAGN = LagNRange(GetSubMatrix(TheArray, , UBound(TheArray, 1) - N), N, TimeDimension).Value2
    ElseIf TimeDimension = "Horizontal" Then
        LAGN = LagNRange(GetSubMatrix(TheArray, , , , UBound(TheArray, 2) - N), N, TimeDimension).Value2
    Else
        Exit Function
    End If
End Function

' Calculates the LN of a 1D or 2D Array
Public Function LnOfArray(TheArray As Variant) As Variant
    Dim tmpSht As Worksheet
    Dim RangeOfTheArray As Range

    Set RangeOfTheArray = ToTemp(TheArray)
    
    ' Set reference to the worksheet where TheArray has been dumped in
    Set tmpSht = RangeOfTheArray.Worksheet
    
    RangeOfTheArray.Offset(0, UBound(TheArray, 2) + 1) = Application.Ln(RangeOfTheArray)
    
    LnOfArray = RangeOfTheArray.Offset(0, UBound(TheArray, 2) + 1).Value2
End Function

' This function performs matrix element-wise division on two 0D, 1D, or 2D arrays.  Clearly, the two arrays
' must have the same dimensions.  The result is returned as an array of the same dimensions as those of
' the input.  This function throws an error if there is a division by 0
Public Function ElementwiseAddition(matrix1 As Variant, matrix2 As Variant) As Variant
    Dim TmpSheet As Worksheet
    Dim numRows As Long
    Dim numColumns As Long
    Dim r As Long ' for number of rows
    Dim c As Long ' for number of columns
    Dim rOffset1 As Long
    Dim cOffset1 As Long
    Dim rOffset2 As Long
    Dim cOffset2 As Long
    Dim TheResults() As Double
    Dim var As Variant
    
    ' Check parameter consistency
    If EmptyArrayQ(matrix1) Or EmptyArrayQ(matrix2) Then
        Let ElementwiseAddition = Null
    End If
    
    If Not (IsNumeric(matrix1) Or VectorQ(matrix1) Or MatrixQ(matrix1)) Then
        Let ElementwiseAddition = Null
        Exit Function
    End If
    
    If Not (IsNumeric(matrix2) Or VectorQ(matrix2) Or MatrixQ(matrix2)) Then
        Let ElementwiseAddition = Null
        Exit Function
    End If
    
    If MatrixQ(matrix1) And MatrixQ(matrix2) Then
        If GetNumberOfRows(matrix1) <> GetNumberOfRows(matrix2) Or _
           GetNumberOfColumns(matrix1) <> GetNumberOfColumns(matrix2) Then
           Let ElementwiseAddition = Null
           Exit Function
        End If
    End If
    
    If (RowVectorQ(matrix1) And MatrixQ(matrix2)) Or _
       (MatrixQ(matrix1) And RowVectorQ(matrix2)) Then
        If GetNumberOfColumns(matrix1) <> GetNumberOfColumns(matrix2) Then
           Let ElementwiseAddition = Null
           Exit Function
        End If
    End If
    
    If (ColumnVectorQ(matrix1) And MatrixQ(matrix2)) Or _
       (MatrixQ(matrix1) And ColumnVectorQ(matrix2)) Then
        If GetNumberOfRows(matrix1) <> GetNumberOfRows(matrix2) Then
           Let ElementwiseAddition = Null
           Exit Function
        End If
    End If
    
    ' Perform the calculations
    If IsNumeric(matrix1) And IsNumeric(matrix2) Then
        If CDbl(matrix2) <> 0 Then
            Let ElementwiseAddition = CDbl(matrix1) + CDbl(matrix2)
            Exit Function
        Else
            Let ElementwiseAddition = Null
            Exit Function
        End If
    ElseIf IsNumeric(matrix1) And RowVectorQ(matrix2) Then
        Let numColumns = GetNumberOfColumns(matrix2)
        
        ReDim TheResults(1 To numColumns)

        ' Compute the r and c offsets due to differences in array starts
        Let cOffset2 = IIf(LBound(matrix2, 1) = 0, 1, 0)
        
        For c = 1 To numColumns
            Let TheResults(c) = CDbl(matrix1) + CDbl(matrix2(c + cOffset2))
        Next c
    ElseIf IsNumeric(matrix1) And ColumnVectorQ(matrix2) Then
        Let numRows = GetNumberOfRows(matrix2)
        
        ReDim TheResults(1 To numRows, 1 To 1)

        ' Compute the r and c offsets due to differences in array starts
        Let rOffset2 = IIf(LBound(matrix2, 1) = 0, 1, 0)
        
        For r = 1 To numRows
            Let TheResults(r, 1) = CDbl(matrix1) + CDbl(matrix2(r + rOffset2, 1))
        Next r
    ElseIf IsNumeric(matrix1) And MatrixQ(matrix2) Then
        Let numRows = GetNumberOfRows(matrix2)
        Let numColumns = GetNumberOfColumns(matrix2)
        
        ReDim TheResults(1 To numRows, 1 To numColumns)
    
        ' Compute the r and c offsets due to differences in array starts
        Let rOffset2 = IIf(LBound(matrix2, 1) = 0, 1, 0)
        Let cOffset2 = IIf(LBound(matrix2, 2) = 0, 1, 0)
        
        For r = 1 To numRows
            For c = 1 To numColumns
                Let TheResults(r, c) = CDbl(matrix1) + CDbl(matrix2(r + rOffset2, c + cOffset2))
            Next c
        Next r
    ElseIf RowVectorQ(matrix1) And MatrixQ(matrix2) Then
        ' If the code gets here, we are adding two 2D matrices of the same size
        Let numRows = GetNumberOfRows(matrix2)
        Let numColumns = GetNumberOfColumns(matrix2)
        
        ReDim TheResults(1 To numRows, 1 To numColumns)
    
        ' Compute the r and c offsets due to differences in array starts
        Let rOffset2 = IIf(LBound(matrix2, 1) = 0, 1, 0)
        Let cOffset1 = IIf(LBound(matrix1) = 0, 1, 0)
        Let cOffset2 = IIf(LBound(matrix2, 2) = 0, 1, 0)
        
        For r = 1 To numRows
            For c = 1 To numColumns
                Let TheResults(r, c) = CDbl(matrix1(c + cOffset1)) + CDbl(matrix2(r + rOffset2, c + cOffset2))
            Next c
        Next r
    ElseIf ColumnVectorQ(matrix1) And MatrixQ(matrix2) Then
        Let numRows = GetNumberOfRows(matrix2)
        Let numColumns = GetNumberOfColumns(matrix2)
        
        ReDim TheResults(1 To numRows, 1 To numColumns)
    
        ' Compute the r and c offsets due to differences in array starts
        Let rOffset1 = IIf(LBound(matrix1, 1) = 0, 1, 0)
        Let rOffset2 = IIf(LBound(matrix2, 1) = 0, 1, 0)
        Let cOffset1 = IIf(LBound(matrix1, 2) = 0, 1, 0)
        Let cOffset2 = IIf(LBound(matrix2, 2) = 0, 1, 0)
        
        For r = 1 To numRows
            For c = 1 To numColumns
                Let TheResults(r, c) = CDbl(matrix1(r + rOffset1, 1 + cOffset1)) + CDbl(matrix2(r + rOffset2, c + cOffset2))
            Next c
        Next r
    ElseIf RowVectorQ(matrix1) And IsNumeric(matrix2) Then
        Let numColumns = GetNumberOfColumns(matrix1)
        
        ReDim TheResults(1 To numColumns)

        ' Compute the r and c offsets due to differences in array starts
        Let cOffset1 = IIf(LBound(matrix1) = 0, 1, 0)
        
        For c = 1 To numColumns
            Let TheResults(c) = CDbl(matrix1(c + cOffset1)) + CDbl(matrix2)
        Next c
    ElseIf ColumnVectorQ(matrix1) And IsNumeric(matrix2) Then
        Let numRows = GetNumberOfRows(matrix1)
        
        ReDim TheResults(1 To numRows, 1 To 1)

        ' Compute the r and c offsets due to differences in array starts
        Let rOffset1 = IIf(LBound(matrix1, 1) = 0, 1, 0)
        
        For r = 1 To numRows
            Let TheResults(r, 1) = CDbl(matrix1(r + rOffset1, 1)) + CDbl(matrix2)
        Next r
    ElseIf MatrixQ(matrix1) And IsNumeric(matrix2) Then
        Let numRows = GetNumberOfRows(matrix1)
        Let numColumns = GetNumberOfColumns(matrix1)
        
        ReDim TheResults(1 To numRows, 1 To numColumns)
    
        ' Compute the r and c offsets due to differences in array starts
        Let rOffset1 = IIf(LBound(matrix1, 1) = 0, 1, 0)
        Let cOffset1 = IIf(LBound(matrix1, 2) = 0, 1, 0)
        
        For r = 1 To numRows
            For c = 1 To numColumns
                Let TheResults(r, c) = CDbl(matrix1(r + rOffset1, c + cOffset1)) + CDbl(matrix2)
            Next c
        Next r
    ElseIf MatrixQ(matrix1) And RowVectorQ(matrix2) Then
        ' If the code gets here, we are adding two 2D matrices of the same size
        Let numRows = GetNumberOfRows(matrix1)
        Let numColumns = GetNumberOfColumns(matrix1)
        
        ReDim TheResults(1 To numRows, 1 To numColumns)
    
        ' Compute the r and c offsets due to differences in array starts
        Let rOffset1 = IIf(LBound(matrix1, 1) = 0, 1, 0)
        Let cOffset1 = IIf(LBound(matrix1, 2) = 0, 1, 0)
        Let cOffset2 = IIf(LBound(matrix2) = 0, 1, 0)
        
        For r = 1 To numRows
            For c = 1 To numColumns
                Let TheResults(r, c) = CDbl(matrix1(r + rOffset1, c + cOffset1)) + CDbl(matrix2(c + cOffset2))
            Next c
        Next r
    ElseIf ColumnVectorQ(matrix2) And MatrixQ(matrix1) Then
        Let numRows = GetNumberOfRows(matrix1)
        Let numColumns = GetNumberOfColumns(matrix1)
        
        ReDim TheResults(1 To numRows, 1 To numColumns)
    
        ' Compute the r and c offsets due to differences in array starts
        Let rOffset1 = IIf(LBound(matrix1, 1) = 0, 1, 0)
        Let rOffset2 = IIf(LBound(matrix2, 1) = 0, 1, 0)
        Let cOffset1 = IIf(LBound(matrix1, 2) = 0, 1, 0)
        Let cOffset2 = IIf(LBound(matrix2, 2) = 0, 1, 0)
        
        For r = 1 To numRows
            For c = 1 To numColumns
                Let TheResults(r, c) = CDbl(matrix1(r + rOffset1, c + cOffset1)) + CDbl(matrix2(r + rOffset2, 1 + cOffset2))
            Next c
        Next r
    ElseIf MatrixQ(matrix1) And MatrixQ(matrix2) Then
        ' If the code gets here, we are adding two 2D matrices of the same size
        Let numRows = GetNumberOfRows(matrix1)
        Let numColumns = GetNumberOfColumns(matrix1)
        
        ReDim TheResults(1 To numRows, 1 To numColumns)
    
        ' Compute the r and c offsets due to differences in array starts
        Let rOffset1 = IIf(LBound(matrix1, 1) = 0, 1, 0)
        Let rOffset2 = IIf(LBound(matrix2, 1) = 0, 1, 0)
        Let cOffset1 = IIf(LBound(matrix1, 2) = 0, 1, 0)
        Let cOffset2 = IIf(LBound(matrix2, 2) = 0, 1, 0)
        
        For r = 1 To numRows
            For c = 1 To numColumns
                Let TheResults(r, c) = CDbl(matrix1(r + rOffset1, c + cOffset1)) + CDbl(matrix2(r + rOffset2, c + cOffset2))
            Next c
        Next r
    Else
        Let ElementwiseAddition = Null
    End If
    
    ' Return the result
    Let ElementwiseAddition = TheResults
End Function

' This function performs matrix element-wise division on two 0D, 1D, or 2D arrays.  Clearly, the two arrays
' must have the same dimensions.  The result is returned as an array of the same dimensions as those of
' the input.  This function throws an error if there is a division by 0
Public Function ElementwiseMultiplication(matrix1 As Variant, matrix2 As Variant) As Variant
    Dim TmpSheet As Worksheet
    Dim numRows As Long
    Dim numColumns As Long
    Dim r As Long ' for number of rows
    Dim c As Long ' for number of columns
    Dim rOffset1 As Long
    Dim cOffset1 As Long
    Dim rOffset2 As Long
    Dim cOffset2 As Long
    Dim TheResults() As Double
    Dim var As Variant
    
    ' Check parameter consistency
    If EmptyArrayQ(matrix1) Or EmptyArrayQ(matrix2) Then
        Let ElementwiseMultiplication = Null
    End If
    
    If Not (IsNumeric(matrix1) Or VectorQ(matrix1) Or MatrixQ(matrix1)) Then
        Let ElementwiseMultiplication = Null
        Exit Function
    End If
    
    If Not (IsNumeric(matrix2) Or VectorQ(matrix2) Or MatrixQ(matrix2)) Then
        Let ElementwiseMultiplication = Null
        Exit Function
    End If
    
    If MatrixQ(matrix1) And MatrixQ(matrix2) Then
        If GetNumberOfRows(matrix1) <> GetNumberOfRows(matrix2) Or _
           GetNumberOfColumns(matrix1) <> GetNumberOfColumns(matrix2) Then
           Let ElementwiseMultiplication = Null
           Exit Function
        End If
    End If
    
    If (RowVectorQ(matrix1) And MatrixQ(matrix2)) Or _
       (MatrixQ(matrix1) And RowVectorQ(matrix2)) Then
        If GetNumberOfColumns(matrix1) <> GetNumberOfColumns(matrix2) Then
           Let ElementwiseMultiplication = Null
           Exit Function
        End If
    End If
    
    If (ColumnVectorQ(matrix1) And MatrixQ(matrix2)) Or _
       (MatrixQ(matrix1) And ColumnVectorQ(matrix2)) Then
        If GetNumberOfRows(matrix1) <> GetNumberOfRows(matrix2) Then
           Let ElementwiseMultiplication = Null
           Exit Function
        End If
    End If
    
    ' Perform the calculations
    If IsNumeric(matrix1) And IsNumeric(matrix2) Then
        If CDbl(matrix2) <> 0 Then
            Let ElementwiseMultiplication = CDbl(matrix1) * CDbl(matrix2)
            Exit Function
        Else
            Let ElementwiseMultiplication = Null
            Exit Function
        End If
    ElseIf IsNumeric(matrix1) And RowVectorQ(matrix2) Then
        Let numColumns = GetNumberOfColumns(matrix2)
        
        ReDim TheResults(1 To numColumns)

        ' Compute the r and c offsets due to differences in array starts
        Let cOffset2 = IIf(LBound(matrix2, 1) = 0, 1, 0)
        
        For c = 1 To numColumns
            Let TheResults(c) = CDbl(matrix1) * CDbl(matrix2(c + cOffset2))
        Next c
    ElseIf IsNumeric(matrix1) And ColumnVectorQ(matrix2) Then
        Let numRows = GetNumberOfRows(matrix2)
        
        ReDim TheResults(1 To numRows, 1 To 1)

        ' Compute the r and c offsets due to differences in array starts
        Let rOffset2 = IIf(LBound(matrix2, 1) = 0, 1, 0)
        
        For r = 1 To numRows
            Let TheResults(r, 1) = CDbl(matrix1) * CDbl(matrix2(r + rOffset2, 1))
        Next r
    ElseIf IsNumeric(matrix1) And MatrixQ(matrix2) Then
        Let numRows = GetNumberOfRows(matrix2)
        Let numColumns = GetNumberOfColumns(matrix2)
        
        ReDim TheResults(1 To numRows, 1 To numColumns)
    
        ' Compute the r and c offsets due to differences in array starts
        Let rOffset2 = IIf(LBound(matrix2, 1) = 0, 1, 0)
        Let cOffset2 = IIf(LBound(matrix2, 2) = 0, 1, 0)
        
        For r = 1 To numRows
            For c = 1 To numColumns
                Let TheResults(r, c) = CDbl(matrix1) * CDbl(matrix2(r + rOffset2, c + cOffset2))
            Next c
        Next r
    ElseIf RowVectorQ(matrix1) And MatrixQ(matrix2) Then
        ' If the code gets here, we are adding two 2D matrices of the same size
        Let numRows = GetNumberOfRows(matrix2)
        Let numColumns = GetNumberOfColumns(matrix2)
        
        ReDim TheResults(1 To numRows, 1 To numColumns)
    
        ' Compute the r and c offsets due to differences in array starts
        Let rOffset2 = IIf(LBound(matrix2, 1) = 0, 1, 0)
        Let cOffset1 = IIf(LBound(matrix1) = 0, 1, 0)
        Let cOffset2 = IIf(LBound(matrix2, 2) = 0, 1, 0)
        
        For r = 1 To numRows
            For c = 1 To numColumns
                Let TheResults(r, c) = CDbl(matrix1(c + cOffset1)) * CDbl(matrix2(r + rOffset2, c + cOffset2))
            Next c
        Next r
    ElseIf ColumnVectorQ(matrix1) And MatrixQ(matrix2) Then
        Let numRows = GetNumberOfRows(matrix2)
        Let numColumns = GetNumberOfColumns(matrix2)
        
        ReDim TheResults(1 To numRows, 1 To numColumns)
    
        ' Compute the r and c offsets due to differences in array starts
        Let rOffset1 = IIf(LBound(matrix1, 1) = 0, 1, 0)
        Let rOffset2 = IIf(LBound(matrix2, 1) = 0, 1, 0)
        Let cOffset1 = IIf(LBound(matrix1, 2) = 0, 1, 0)
        Let cOffset2 = IIf(LBound(matrix2, 2) = 0, 1, 0)
        
        For r = 1 To numRows
            For c = 1 To numColumns
                Let TheResults(r, c) = CDbl(matrix1(r + rOffset1, 1 + cOffset1)) * CDbl(matrix2(r + rOffset2, c + cOffset2))
            Next c
        Next r
    ElseIf RowVectorQ(matrix1) And IsNumeric(matrix2) Then
        Let numColumns = GetNumberOfColumns(matrix1)
        
        ReDim TheResults(1 To numColumns)

        ' Compute the r and c offsets due to differences in array starts
        Let cOffset1 = IIf(LBound(matrix1) = 0, 1, 0)
        
        For c = 1 To numColumns
            Let TheResults(c) = CDbl(matrix1(c + cOffset1)) * CDbl(matrix2)
        Next c
    ElseIf ColumnVectorQ(matrix1) And IsNumeric(matrix2) Then
        Let numRows = GetNumberOfRows(matrix1)
        
        ReDim TheResults(1 To numRows, 1 To 1)

        ' Compute the r and c offsets due to differences in array starts
        Let rOffset1 = IIf(LBound(matrix1, 1) = 0, 1, 0)
        
        For r = 1 To numRows
            Let TheResults(r, 1) = CDbl(matrix1(r + rOffset1, 1)) * CDbl(matrix2)
        Next r
    ElseIf MatrixQ(matrix1) And IsNumeric(matrix2) Then
        Let numRows = GetNumberOfRows(matrix1)
        Let numColumns = GetNumberOfColumns(matrix1)
        
        ReDim TheResults(1 To numRows, 1 To numColumns)
    
        ' Compute the r and c offsets due to differences in array starts
        Let rOffset1 = IIf(LBound(matrix1, 1) = 0, 1, 0)
        Let cOffset1 = IIf(LBound(matrix1, 2) = 0, 1, 0)
        
        For r = 1 To numRows
            For c = 1 To numColumns
                Let TheResults(r, c) = CDbl(matrix1(r + rOffset1, c + cOffset1)) * CDbl(matrix2)
            Next c
        Next r
    ElseIf MatrixQ(matrix1) And RowVectorQ(matrix2) Then
        ' If the code gets here, we are adding two 2D matrices of the same size
        Let numRows = GetNumberOfRows(matrix1)
        Let numColumns = GetNumberOfColumns(matrix1)
        
        ReDim TheResults(1 To numRows, 1 To numColumns)
    
        ' Compute the r and c offsets due to differences in array starts
        Let rOffset1 = IIf(LBound(matrix1, 1) = 0, 1, 0)
        Let cOffset1 = IIf(LBound(matrix1, 2) = 0, 1, 0)
        Let cOffset2 = IIf(LBound(matrix2) = 0, 1, 0)
        
        For r = 1 To numRows
            For c = 1 To numColumns
                Let TheResults(r, c) = CDbl(matrix1(r + rOffset1, c + cOffset1)) * CDbl(matrix2(c + cOffset2))
            Next c
        Next r
    ElseIf ColumnVectorQ(matrix2) And MatrixQ(matrix1) Then
        Let numRows = GetNumberOfRows(matrix1)
        Let numColumns = GetNumberOfColumns(matrix1)
        
        ReDim TheResults(1 To numRows, 1 To numColumns)
    
        ' Compute the r and c offsets due to differences in array starts
        Let rOffset1 = IIf(LBound(matrix1, 1) = 0, 1, 0)
        Let rOffset2 = IIf(LBound(matrix2, 1) = 0, 1, 0)
        Let cOffset1 = IIf(LBound(matrix1, 2) = 0, 1, 0)
        Let cOffset2 = IIf(LBound(matrix2, 2) = 0, 1, 0)
        
        For r = 1 To numRows
            For c = 1 To numColumns
                Let TheResults(r, c) = CDbl(matrix1(r + rOffset1, c + cOffset1)) * CDbl(matrix2(r + rOffset2, 1 + cOffset2))
            Next c
        Next r
    ElseIf MatrixQ(matrix1) And MatrixQ(matrix2) Then
        ' If the code gets here, we are adding two 2D matrices of the same size
        Let numRows = GetNumberOfRows(matrix1)
        Let numColumns = GetNumberOfColumns(matrix1)
        
        ReDim TheResults(1 To numRows, 1 To numColumns)
    
        ' Compute the r and c offsets due to differences in array starts
        Let rOffset1 = IIf(LBound(matrix1, 1) = 0, 1, 0)
        Let rOffset2 = IIf(LBound(matrix2, 1) = 0, 1, 0)
        Let cOffset1 = IIf(LBound(matrix1, 2) = 0, 1, 0)
        Let cOffset2 = IIf(LBound(matrix2, 2) = 0, 1, 0)
        
        For r = 1 To numRows
            For c = 1 To numColumns
                Let TheResults(r, c) = CDbl(matrix1(r + rOffset1, c + cOffset1)) * CDbl(matrix2(r + rOffset2, c + cOffset2))
            Next c
        Next r
    Else
        Let ElementwiseMultiplication = Null
    End If
    
    ' Return the result
    Let ElementwiseMultiplication = TheResults
End Function

' This function performs matrix element-wise division on two 0D, 1D, or 2D arrays.  Clearly, the two arrays
' must have the same dimensions.  The result is returned as an array of the same dimensions as those of
' the input.  This function throws an error if there is a division by 0
Public Function ElementWiseDivision(matrix1 As Variant, matrix2 As Variant) As Variant
    Dim TmpSheet As Worksheet
    Dim numRows As Long
    Dim numColumns As Long
    Dim r As Long ' for number of rows
    Dim c As Long ' for number of columns
    Dim rOffset1 As Long
    Dim cOffset1 As Long
    Dim rOffset2 As Long
    Dim cOffset2 As Long
    Dim TheResults() As Double
    Dim var As Variant
    
    ' Check parameter consistency
    If EmptyArrayQ(matrix1) Or EmptyArrayQ(matrix2) Then
        Let ElementWiseDivision = Null
    End If
    
    If Not (IsNumeric(matrix1) Or VectorQ(matrix1) Or MatrixQ(matrix1)) Then
        Let ElementWiseDivision = Null
        Exit Function
    End If
    
    If Not (IsNumeric(matrix2) Or VectorQ(matrix2) Or MatrixQ(matrix2)) Then
        Let ElementWiseDivision = Null
        Exit Function
    End If
    
    If MatrixQ(matrix1) And MatrixQ(matrix2) Then
        If GetNumberOfRows(matrix1) <> GetNumberOfRows(matrix2) Or _
           GetNumberOfColumns(matrix1) <> GetNumberOfColumns(matrix2) Then
           Let ElementWiseDivision = Null
           Exit Function
        End If
    End If
    
    If (RowVectorQ(matrix1) And MatrixQ(matrix2)) Or _
       (MatrixQ(matrix1) And RowVectorQ(matrix2)) Then
        If GetNumberOfColumns(matrix1) <> GetNumberOfColumns(matrix2) Then
           Let ElementWiseDivision = Null
           Exit Function
        End If
    End If
    
    If (ColumnVectorQ(matrix1) And MatrixQ(matrix2)) Or _
       (MatrixQ(matrix1) And ColumnVectorQ(matrix2)) Then
        If GetNumberOfRows(matrix1) <> GetNumberOfRows(matrix2) Then
           Let ElementWiseDivision = Null
           Exit Function
        End If
    End If
    
    ' Perform the calculations
    If IsNumeric(matrix1) And IsNumeric(matrix2) Then
        If CDbl(matrix2) <> 0 Then
            Let ElementWiseDivision = CDbl(matrix1) / CDbl(matrix2)
            Exit Function
        Else
            Let ElementWiseDivision = Null
            Exit Function
        End If
    ElseIf IsNumeric(matrix1) And RowVectorQ(matrix2) Then
        Let numColumns = GetNumberOfColumns(matrix2)
        
        ReDim TheResults(1 To numColumns)

        ' Compute the r and c offsets due to differences in array starts
        Let cOffset2 = IIf(LBound(matrix2, 1) = 0, 1, 0)
        
        For c = 1 To numColumns
            Let TheResults(c) = CDbl(matrix1) / CDbl(matrix2(c + cOffset2))
        Next c
    ElseIf IsNumeric(matrix1) And ColumnVectorQ(matrix2) Then
        Let numRows = GetNumberOfRows(matrix2)
        
        ReDim TheResults(1 To numRows, 1 To 1)

        ' Compute the r and c offsets due to differences in array starts
        Let rOffset2 = IIf(LBound(matrix2, 1) = 0, 1, 0)
        
        For r = 1 To numRows
            Let TheResults(r, 1) = CDbl(matrix1) / CDbl(matrix2(r + rOffset2, 1))
        Next r
    ElseIf IsNumeric(matrix1) And MatrixQ(matrix2) Then
        Let numRows = GetNumberOfRows(matrix2)
        Let numColumns = GetNumberOfColumns(matrix2)
        
        ReDim TheResults(1 To numRows, 1 To numColumns)
    
        ' Compute the r and c offsets due to differences in array starts
        Let rOffset2 = IIf(LBound(matrix2, 1) = 0, 1, 0)
        Let cOffset2 = IIf(LBound(matrix2, 2) = 0, 1, 0)
        
        For r = 1 To numRows
            For c = 1 To numColumns
                Let TheResults(r, c) = CDbl(matrix1) / CDbl(matrix2(r + rOffset2, c + cOffset2))
            Next c
        Next r
    ElseIf RowVectorQ(matrix1) And MatrixQ(matrix2) Then
        ' If the code gets here, we are adding two 2D matrices of the same size
        Let numRows = GetNumberOfRows(matrix2)
        Let numColumns = GetNumberOfColumns(matrix2)
        
        ReDim TheResults(1 To numRows, 1 To numColumns)
    
        ' Compute the r and c offsets due to differences in array starts
        Let rOffset2 = IIf(LBound(matrix2, 1) = 0, 1, 0)
        Let cOffset1 = IIf(LBound(matrix1) = 0, 1, 0)
        Let cOffset2 = IIf(LBound(matrix2, 2) = 0, 1, 0)
        
        For r = 1 To numRows
            For c = 1 To numColumns
                Let TheResults(r, c) = CDbl(matrix1(c + cOffset1)) / CDbl(matrix2(r + rOffset2, c + cOffset2))
            Next c
        Next r
    ElseIf ColumnVectorQ(matrix1) And MatrixQ(matrix2) Then
        Let numRows = GetNumberOfRows(matrix2)
        Let numColumns = GetNumberOfColumns(matrix2)
        
        ReDim TheResults(1 To numRows, 1 To numColumns)
    
        ' Compute the r and c offsets due to differences in array starts
        Let rOffset1 = IIf(LBound(matrix1, 1) = 0, 1, 0)
        Let rOffset2 = IIf(LBound(matrix2, 1) = 0, 1, 0)
        Let cOffset1 = IIf(LBound(matrix1, 2) = 0, 1, 0)
        Let cOffset2 = IIf(LBound(matrix2, 2) = 0, 1, 0)
        
        For r = 1 To numRows
            For c = 1 To numColumns
                Let TheResults(r, c) = CDbl(matrix1(r + rOffset1, 1 + cOffset1)) / CDbl(matrix2(r + rOffset2, c + cOffset2))
            Next c
        Next r
    ElseIf RowVectorQ(matrix1) And IsNumeric(matrix2) Then
        Let numColumns = GetNumberOfColumns(matrix1)
        
        ReDim TheResults(1 To numColumns)

        ' Compute the r and c offsets due to differences in array starts
        Let cOffset1 = IIf(LBound(matrix1) = 0, 1, 0)
        
        For c = 1 To numColumns
            Let TheResults(c) = CDbl(matrix1(c + cOffset1)) / CDbl(matrix2)
        Next c
    ElseIf ColumnVectorQ(matrix1) And IsNumeric(matrix2) Then
        Let numRows = GetNumberOfRows(matrix1)
        
        ReDim TheResults(1 To numRows, 1 To 1)

        ' Compute the r and c offsets due to differences in array starts
        Let rOffset1 = IIf(LBound(matrix1, 1) = 0, 1, 0)
        
        For r = 1 To numRows
            Let TheResults(r, 1) = CDbl(matrix1(r + rOffset1, 1)) / CDbl(matrix2)
        Next r
    ElseIf MatrixQ(matrix1) And IsNumeric(matrix2) Then
        Let numRows = GetNumberOfRows(matrix1)
        Let numColumns = GetNumberOfColumns(matrix1)
        
        ReDim TheResults(1 To numRows, 1 To numColumns)
    
        ' Compute the r and c offsets due to differences in array starts
        Let rOffset1 = IIf(LBound(matrix1, 1) = 0, 1, 0)
        Let cOffset1 = IIf(LBound(matrix1, 2) = 0, 1, 0)
        
        For r = 1 To numRows
            For c = 1 To numColumns
                Let TheResults(r, c) = CDbl(matrix1(r + rOffset1, c + cOffset1)) / CDbl(matrix2)
            Next c
        Next r
    ElseIf MatrixQ(matrix1) And RowVectorQ(matrix2) Then
        ' If the code gets here, we are adding two 2D matrices of the same size
        Let numRows = GetNumberOfRows(matrix1)
        Let numColumns = GetNumberOfColumns(matrix1)
        
        ReDim TheResults(1 To numRows, 1 To numColumns)
    
        ' Compute the r and c offsets due to differences in array starts
        Let rOffset1 = IIf(LBound(matrix1, 1) = 0, 1, 0)
        Let cOffset1 = IIf(LBound(matrix1, 2) = 0, 1, 0)
        Let cOffset2 = IIf(LBound(matrix2) = 0, 1, 0)
        
        For r = 1 To numRows
            For c = 1 To numColumns
                Let TheResults(r, c) = CDbl(matrix1(r + rOffset1, c + cOffset1)) / CDbl(matrix2(c + cOffset2))
            Next c
        Next r
    ElseIf ColumnVectorQ(matrix2) And MatrixQ(matrix1) Then
        Let numRows = GetNumberOfRows(matrix1)
        Let numColumns = GetNumberOfColumns(matrix1)
        
        ReDim TheResults(1 To numRows, 1 To numColumns)
    
        ' Compute the r and c offsets due to differences in array starts
        Let rOffset1 = IIf(LBound(matrix1, 1) = 0, 1, 0)
        Let rOffset2 = IIf(LBound(matrix2, 1) = 0, 1, 0)
        Let cOffset1 = IIf(LBound(matrix1, 2) = 0, 1, 0)
        Let cOffset2 = IIf(LBound(matrix2, 2) = 0, 1, 0)
        
        For r = 1 To numRows
            For c = 1 To numColumns
                Let TheResults(r, c) = CDbl(matrix1(r + rOffset1, c + cOffset1)) / CDbl(matrix2(r + rOffset2, 1 + cOffset2))
            Next c
        Next r
    ElseIf MatrixQ(matrix1) And MatrixQ(matrix2) Then
        ' If the code gets here, we are adding two 2D matrices of the same size
        Let numRows = GetNumberOfRows(matrix1)
        Let numColumns = GetNumberOfColumns(matrix1)
        
        ReDim TheResults(1 To numRows, 1 To numColumns)
    
        ' Compute the r and c offsets due to differences in array starts
        Let rOffset1 = IIf(LBound(matrix1, 1) = 0, 1, 0)
        Let rOffset2 = IIf(LBound(matrix2, 1) = 0, 1, 0)
        Let cOffset1 = IIf(LBound(matrix1, 2) = 0, 1, 0)
        Let cOffset2 = IIf(LBound(matrix2, 2) = 0, 1, 0)
        
        For r = 1 To numRows
            For c = 1 To numColumns
                Let TheResults(r, c) = CDbl(matrix1(r + rOffset1, c + cOffset1)) / CDbl(matrix2(r + rOffset2, c + cOffset2))
            Next c
        Next r
    Else
        Let ElementWiseDivision = Null
    End If
    
    ' Return the result
    Let ElementWiseDivision = TheResults
End Function

'***HERE explain this a bit clearer
' This function dumps an array (1D or 2D) into worksheet TempComputation, shifts the array by N rows or N columns depending
' on the value of TimeDimension, and then returns a reference to the underlying range.  Dimensions are preserved.
' This means that an m x n array is dumped into an m x n range. A 1D array is dumped vertically to allow for bigger arrays since Excel 2010
' has a horizontal maximum under 17K columns.
Public Function LagNRange(AnArray As Variant, N As Variant, TimeDimension As String) As Range
    Dim tmpSht As Worksheet
    
    ' Set reference to worksheet TempComputation
    Set tmpSht = ActiveWorkbook.Worksheets("TempComputation")
    
    ' Clear used range
    tmpSht.UsedRange.ClearContents
    
    If TimeDimension = "Vertical" Then
        If NumberOfDimensions(AnArray) = 1 Then
            Let tmpSht.Range("A1").Offset(N, 0).Resize(UBound(AnArray)).Value2 = Application.Transpose(AnArray)
            Set LagNRange = tmpSht.Range("A1").Resize(UBound(AnArray) + N, 1)
        Else
            Let tmpSht.Range("A1").Offset(N, 0).Resize(UBound(AnArray, 1), UBound(AnArray, 2)).Value2 = AnArray
            Set LagNRange = tmpSht.Range("A1").Resize(UBound(AnArray, 1) + N, UBound(AnArray, 2))
        End If
    ElseIf TimeDimension = "Horizontal" Then
        If NumberOfDimensions(AnArray) = 1 Then
            Let tmpSht.Range("A1").Offset(0, N).Resize(UBound(AnArray)).Value2 = Application.Transpose(AnArray)
            Set LagNRange = tmpSht.Range("A1").Resize(UBound(AnArray) + N, 1)
        Else
            Let tmpSht.Range("A1").Offset(0, N).Resize(UBound(AnArray, 1), UBound(AnArray, 2)).Value2 = AnArray
            Set LagNRange = tmpSht.Range("A1").Resize(UBound(AnArray, 1), UBound(AnArray, 2) + N)
        End If
    Else
        Exit Function:
    End If
End Function

'This function returns an Array of NrOfRows rows and NrOfColumns and all elements are equal to -1
Public Function MatrixOfMinusOnes(NrOfRows As Long, NrOfColumns As Long) As Variant
    Dim tmpSht As Worksheet
    
    ' Set reference to worksheet TempComputation
    Set tmpSht = Worksheets("TempComputation")
    
    ' Clear used range
    tmpSht.UsedRange.ClearContents
    
    Let tmpSht.Range("A1").Resize(NrOfRows, NrOfColumns).Value2 = -1
    Let MatrixOfMinusOnes = tmpSht.Range("A1").Resize(NrOfRows, NrOfColumns)
End Function

'This function returns an Array of NrOfRows rows and NrOfColumns columns and all elements are equal to 1
Public Function Ones(NrOfRows As Long, NrOfColumns As Long) As Variant
    Dim tmpSht As Worksheet
    
    ' Set reference to worksheet TempComputation
    Set tmpSht = ThisWorkbook.Worksheets("TempComputation")
    
    ' Clear used range
    Call tmpSht.UsedRange.ClearContents
    
    Let tmpSht.Range("A1").Resize(NrOfRows, NrOfColumns).Value2 = 1
    Let Ones = tmpSht.Range("A1").Resize(NrOfRows, NrOfColumns)
End Function

'This function returns an Array of NrOfRows rows and NrOfColumns columns and all elements are equal to 1
Public Function Zeroes(NrOfRows As Long, NrOfColumns As Long) As Variant
    Dim tmpSht As Worksheet
    
    ' Set reference to worksheet TempComputation
    Set tmpSht = ThisWorkbook.Worksheets("TempComputation")
    
    ' Clear used range
    Call tmpSht.UsedRange.ClearContents
    
    Let tmpSht.Range("A1").Resize(NrOfRows, NrOfColumns).Value2 = 0
    Let Zeroes = tmpSht.Range("A1").Resize(NrOfRows, NrOfColumns)
End Function

' This function sorts a 1D array in the ascending or descending order, assuming no header.
' This returns a 1D array with the sorted values
Public Function Sort1DArray(MyArray As Variant) As Variant
    Dim TheRange As Range
    Dim TmpSheet As Worksheet
    Dim i As Integer
    Dim LastRowNumber As Long
    
    ' Set pointer to temp sheet and clear its used range
    Set TmpSheet = ThisWorkbook.Worksheets("TempComputation")
    Call TmpSheet.UsedRange.ClearContents
    
    ' Dump array in temp sheet
    Call ToTemp(Application.Transpose(ConvertTo1DArray(MyArray)))
    
    ' Find the last row used
    Let LastRowNumber = TmpSheet.Range("A1").Offset(TmpSheet.Rows.Count - 1, 0).End(xlUp).row

    ' Set range pointer to the data we just dumped in the temp sheet
    Set TheRange = TmpSheet.Range("A1").Resize(LastRowNumber, 1)

    ' Clear any previous sorting criteria
    TmpSheet.Sort.SortFields.Clear

    ' Add all the sorting criteria
    TheRange.Worksheet.Sort.SortFields.Add _
        Key:=TheRange, _
        SortOn:=xlSortOnValues, _
        Order:=xlAscending, _
        DataOption:=xlSortNormal
    
    ' Execute the sort
    With TheRange.Worksheet.Sort
        .SetRange TheRange
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    ' Extract horizontally reversed matrix and set them to return when the function exits
    Let Sort1DArray = ConvertTo1DArray(TheRange.Value2)
End Function

' Returns a 2D array of uniformly distributed random numbers between 0 and 1.
Public Function RandomMatrix(NRows As Long, NColumns As Long) As Variant
    Dim ReturnMatrix() As Double

    If NRows < 0 Or NColumns < 0 Then
        Let RandomMatrix = Empty
        Exit Function
    End If
    
    If NRows = 0 Or NColumns = 0 Then
        Let RandomMatrix = Empty
        Exit Function
    End If
    
    ' Allocate a return
    ReDim ReturnMatrix(NRows, NColumns)
    
    ' Clear TmpSht
    Call ThisWorkbook.Worksheets("TempComputation").UsedRange.ClearContents
    
    ' Create random matrix
    Let Application.Calculation = xlManual
    Let ThisWorkbook.Worksheets("TempComputation").Cells(1, 1).Resize(NRows, NColumns).Formula = "=Rand()"
    Call Application.Calculate
    Let Application.Calculation = xlAutomatic
    
    ' Return random matrix
    Let RandomMatrix = ThisWorkbook.Worksheets("TempComputation").Range("A1").CurrentRegion.Value2
    
    ' Clear TmpSht
    Call ThisWorkbook.Worksheets("TempComputation").UsedRange.ClearContents
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

' This function returns a constant array with the requested dimensions
' If NCols is not given, the result is a 1D array
Public Function ConstantArray(TheValue As Variant, N As Long, Optional NCols As Variant) As Variant
    Dim TheResult() As Variant
    Dim c As Long
    Dim r As Long
    
    If IsMissing(NCols) Then
        If Not IsNumeric(N) Then
            Let ConstantArray = EmptyArray()
            Exit Function
        ElseIf N < 0 Then
            Let ConstantArray = EmptyArray()
            Exit Function
        ElseIf N = 1 Then
            Let ConstantArray = Array(TheValue)
            Exit Function
        End If
    
        ReDim TheResult(1 To N)
        
        For c = 1 To N
            Let TheResult(c) = TheValue
        Next c
        
        Let ConstantArray = TheResult
    Else
        ReDim TheResult(1 To N, 1 To NCols)
        
        For r = 1 To N
            For c = 1 To NCols
                Let TheResult(r, c) = TheValue
            Next c
        Next r
        
        Let ConstantArray = TheResult
    End If
End Function

' This function returns the array resulting from concatenating B to the right of A.
' A and B must have the same dimensions (e.g. 1 or 2D)
' If dim(A)<>dim(B) or dim(A)>2 or dim(B)>2 or dim(A)<1 or dim(B)<1 then
' this function returns EmptyArray()
Public Function ConcatenateArrays(a As Variant, b As Variant) As Variant
    If NumberOfDimensions(a) = 1 And NumberOfDimensions(b) = 1 Then
        Let ConcatenateArrays = ConvertTo1DArray(GetRow(TransposeMatrix(Stack2DArrays(TransposeMatrix(a), TransposeMatrix(b))), 1))
    Else
        Let ConcatenateArrays = TransposeMatrix(Stack2DArrays(TransposeMatrix(a), TransposeMatrix(b)))
    End If
End Function

' This function sorts the given range by the columns whose positions are given by ' ArrayOfColPos.
' The sorting orientation in each column are in ArrayOfColsSortOrder
' ArrayOfColsSortOrder is a variant array whose elements are all of enumerated type XLSortOrder (e.g. xlAscending, xlDescending)
Public Sub SortRange(MyRange As Range, ArrayOfColPos As Variant, _
                     ArrayOfColsSortOrder As Variant, _
                     WithHeaders As XlYesNoGuess)
    Dim TmpSheet As Worksheet
    Dim i As Integer
    
    ' Clear any previous sorting criteria
    Call MyRange.Worksheet.Sort.SortFields.Clear

    ' Add all the sorting criteria
    For i = LBound(ArrayOfColPos) To UBound(ArrayOfColPos)
        ' Add criteria to sort by date
        MyRange.Worksheet.Sort.SortFields.Add _
            Key:=MyRange.Columns(ArrayOfColPos(i)), _
            SortOn:=xlSortOnValues, _
            Order:=ArrayOfColsSortOrder(i), _
            DataOption:=xlSortNormal
    Next i
    
    ' Execute the sort
    With MyRange.Worksheet.Sort
        .SetRange MyRange
        .Header = WithHeaders
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub

' This function swaps two columns from a range, returning True if the operation is successful or False if it is not
Public Function SwapRangeColumns(TheRange As Range, FirstColumnIndex As Long, SecondColumIndex As Long) As Boolean
    Dim col1 As Range
    
    If FirstColumnIndex < 1 Or FirstColumnIndex > TheRange.Columns.Count Or SecondColumIndex < 1 Or SecondColumIndex > TheRange.Columns.Count Or FirstColumnIndex = SecondColumIndex Then
        Let SwapRangeColumns = False

        Exit Function
    End If
    
    Call TheRange.Worksheet.Activate
    
    Set col1 = TheRange.Worksheet.UsedRange
    Set col1 = col1.Range("a1").Offset(col1.Rows.Count, col1.Columns.Count).Resize(TheRange.Rows.Count, 1)
    
    Call TheRange.Columns(FirstColumnIndex).Copy
    Call col1.Range("A1").Select
    Call TheRange.Worksheet.Paste
    
    Call TheRange.Columns(SecondColumIndex).Copy
    Call TheRange.Cells(1, FirstColumnIndex).Select
    Call TheRange.Worksheet.Paste
    
    Call col1.Copy
    Call TheRange.Columns(SecondColumIndex).Select
    Call TheRange.Worksheet.Paste
    
    Call col1.ClearContents
    Call col1.ClearFormats
    Call col1.ClearComments
    Call col1.ClearHyperlinks
    Call col1.ClearNotes
    Call col1.ClearOutline
    
    Let SwapRangeColumns = True
End Function

' This function swaps two columns from a range, returning True if the operation is successful or False if it is not
Public Function SwapRangeRows(TheRange As Range, FirstRowIndex As Long, SecondRowIndex As Long) As Boolean
    Dim Row1 As Range
    
    If FirstRowIndex < 1 Or FirstRowIndex > TheRange.Columns.Count Or SecondRowIndex < 1 Or SecondRowIndex > TheRange.Columns.Count Or FirstRowIndex = SecondRowIndex Then
        Let SwapRangeRows = False

        Exit Function
    End If
    
    Call TheRange.Worksheet.Activate
    
    Set Row1 = TheRange.Worksheet.UsedRange
    Set Row1 = Row1.Range("a1").Offset(Row1.Rows.Count, Row1.Columns.Count).Resize(1, TheRange.Columns.Count)
    
    Call TheRange.Rows(FirstRowIndex).Copy
    Call Row1.Range("A1").Select
    Call TheRange.Worksheet.Paste
    
    Call TheRange.Rows(SecondRowIndex).Copy
    Call TheRange.Rows(FirstRowIndex).Select
    Call TheRange.Worksheet.Paste
    
    Call Row1.Copy
    Call TheRange.Rows(SecondRowIndex).Select
    Call TheRange.Worksheet.Paste
    
    Call Row1.ClearContents
    Call Row1.ClearFormats
    Call Row1.ClearComments
    Call Row1.ClearHyperlinks
    Call Row1.ClearNotes
    Call Row1.ClearOutline
    
    Let SwapRangeRows = True
End Function

' This function swaps two columns from a range, returning the resulting matrix if the operation is successful or False if it is not
'***HERE
Public Function SwapMatrixColumns(TheMatrix As Variant, FirstColumnIndex As Long, SecondColumIndex As Long) As Variant
    Dim col1 As Variant
    
    If FirstColumnIndex < 1 Or FirstColumnIndex > GetNumberOfColumns(TheMatrix) Or SecondColumIndex < 1 Or SecondColumIndex > GetNumberOfColumns(TheMatrix) Or FirstColumnIndex = SecondColumIndex Then
        Let SwapMatrixColumns = False

        Exit Function
    End If
    
    Call ToTemp(TheMatrix)
    Let col1 = GetColumn(TheMatrix, FirstColumnIndex)
    Call DumpInSheet(GetColumn(TheMatrix, SecondColumIndex), TempComputation.Cells(1, FirstColumnIndex))
    Call DumpInSheet(col1, TempComputation.Cells(1, SecondColumIndex))
    
    Let SwapMatrixColumns = TempComputation.Range("A1").CurrentRegion.Value2
End Function

' This function swaps two columns from a range, returning  the resulting matrix if the operation is successful or False if it is not
'***HERE
Public Function SwapMatrixRows(TheMatrix As Variant, FirstRowIndex As Long, SecondRowIndex As Long) As Variant
    Dim Row1 As Variant

    If FirstRowIndex < 1 Or FirstRowIndex > GetNumberOfRows(TheMatrix) Or SecondRowIndex < 1 Or SecondRowIndex > GetNumberOfRows(TheMatrix) Or FirstRowIndex = SecondRowIndex Then
        Let SwapMatrixRows = False

        Exit Function
    End If

    Call ToTemp(TheMatrix)
    Let Row1 = GetRow(TheMatrix, FirstRowIndex)
    Call DumpInSheet(GetRow(TheMatrix, SecondRowIndex), TempComputation.Cells(FirstRowIndex, 1))
    Call DumpInSheet(Row1, TempComputation.Cells(SecondRowIndex, 1))
    
    Let SwapMatrixRows = TempComputation.Range("A1").CurrentRegion.Value2
End Function

' This function transposes 1D or 2D arrays
' This function uses the built-in transposition function unless the optional parameter
' UseBuiltInQ is passed with a value of False.
Public Function TransposeMatrix(aMatrix As Variant, Optional UseBuiltInQ As Boolean = True) As Variant
    Dim r As Long
    Dim c As Long
    Dim TheResult() As Variant
    
    If Not DimensionedQ(aMatrix) Then
        Let TransposeMatrix = Null
        Exit Function
    End If
    
    If NumberOfDimensions(aMatrix) = 0 Then
        Let TransposeMatrix = aMatrix
        Exit Function
    End If
    
    If EmptyArrayQ(aMatrix) Then
        Let TransposeMatrix = EmptyArray()
        Exit Function
    End If
    
    If UseBuiltInQ Then
        Let TransposeMatrix = Application.Transpose(aMatrix)
        Exit Function
    End If
    
    If NumberOfDimensions(aMatrix) = 1 Then
        ReDim TheResult(LBound(aMatrix) To UBound(aMatrix), 1)
    
        For c = LBound(aMatrix) To UBound(aMatrix)
            Let TheResult(c, 1) = aMatrix(c)
        Next c
        
        Let TransposeMatrix = TheResult
        
        Exit Function
    End If
    
    ReDim TheResult(LBound(aMatrix, 2) To UBound(aMatrix, 2), LBound(aMatrix, 1) To UBound(aMatrix, 1))

    For r = LBound(aMatrix, 1) To UBound(aMatrix, 1)
        For c = LBound(aMatrix, 2) To UBound(aMatrix, 2)
            Let TheResult(c, r) = aMatrix(r, c)
        Next c
    Next r
    
    Let TransposeMatrix = TheResult
End Function

' This function transposes a 1D array of 1D arrays into a 1D array of 1D arrays.
' For instance,
' Array(Array(1,2,3), Array(10,20,30)) => Array(Array(1,10), Array(2, 20), Array(3, 30))
Public Function TransposeRectangular1DArrayOf1DArrays(AnArray As Variant) As Variant
    Dim r As Long
    Dim c As Long
    Dim TheResult() As Variant
    Dim TheMatrix As Variant
    Dim var As Variant
    Dim TheLength As Long
    
    ' Exit with null if AnArray is not an array
    If Not IsArray(AnArray) Then
        Let TransposeRectangular1DArrayOf1DArrays = Null
        Exit Function
    End If
    
    ' Exith with Null if AnArray is an empty array
    If EmptyArrayQ(AnArray) Then
        Let TransposeRectangular1DArrayOf1DArrays = Null
        Exit Function
    End If
    
    Let TheLength = GetArrayLength(First(AnArray))
    
    For Each var In AnArray
        If Not AtomicArrayQ(var) Or GetArrayLength(var) <> TheLength Then
            Let TransposeRectangular1DArrayOf1DArrays = Null
            Exit Function
        End If
    Next
    
    Let TheMatrix = Pack2DArray(AnArray)
    ReDim TheResult(1 To GetNumberOfColumns(TheMatrix))
    For c = 1 To GetArrayLength(First(AnArray))
        Let TheResult(c) = ConvertTo1DArray(GetColumn(TheMatrix, c))
    Next
    
    Let TransposeRectangular1DArrayOf1DArrays = TheResult
End Function

' This function returns a 2D array consolidating the special visible cells of an autofiltered range
' The returned array includes the header row
Public Function GetConsolidatedVisibleCells(ARange As Range) As Variant
    Call TempComputation.Range("A1").Resize(ARange.Rows.Count + 1, ARange.Columns.Count + 1).ClearContents
    Call DumpVisibleCells(SourceRange:=ARange, TargetRange:=TempComputation.Range("A1"))
    
    Let GetConsolidatedVisibleCells = TempComputation.Range("A1").CurrentRegion.Value2
End Function

' This function copies the visible cells of a range to the range defined by the upper left corner
' of the target range
Public Sub DumpVisibleCells(SourceRange As Range, TargetRange As Range)
    Call SourceRange.SpecialCells(xlCellTypeVisible).Copy
    Call TargetRange.Range("A1").Resize().PasteSpecial(Paste:=xlPasteAll)
End Sub

' This function returns the autofiltered result of a matrix with a header row. The resulting set is returned
' as a matrix.
' Parameters:
' 1. ColumnsToFilter is an array of integers
' 2. Criteria1List is an array of variants
' 3. OperatorList is an array of xlAutoFilterOperator (e.g. xlFilterValues, xlAnd, xlOr, etc.)
' 4. Criteria2List is an array of variants
'
' If the optional parameters are present, you must make them as long as ColumnsToFilter.
' If a particular column uses Criteria1List exclusively (e.g. filtered on one single condition),
' then enter Empty for the corresponding array elements in OperatorList and Criteria2List.
'
' For example, assume the matrix has 5 columns.  Assume you are filter on columns 1 and 3. Then,
' sample sets of parameters are:
'
' Example 1
' 1. ColumnsToFilter = Array(1,3)
' 2. Criteria1List = Array("<=0.5", "=10")
'
' Example 2
' 1. ColumnsToFilter = Array(1,3)
' 2. Criteria1List = Array("<=0.5", "<=1,5")
' 3. OperatorList = Array(Empty, xlOr)
' 4. Criteria2List = Array(Empty, ">=0.5")
'
' An empty array (identifier via the EmptyArrayQ() predicate, is returned when there is an error.
Public Function AutofilterMatrix(TheMatrix As Variant, ColumnsToFilter As Variant, Criteria1List As Variant, _
                                 Optional OperatorList As Variant, Optional Criteria2List As Variant) As Variant
    Dim i As Integer
    Dim ListObjRef As ListObject
    Dim NCols As Long
    Dim NRows As Long

    If NumberOfDimensions(TheMatrix) <> 2 Or GetNumberOfRows(TheMatrix) < 2 Or _
       NumberOfDimensions(ColumnsToFilter) <> 1 Or NumberOfDimensions(Criteria1List) <> 1 Or _
       GetArrayLength(ColumnsToFilter) <> GetArrayLength(Criteria1List) Then
        Let AutofilterMatrix = EmptyArray()
        Exit Function
    End If
    
    If Not IsMissing(OperatorList) Or Not IsMissing(Criteria2List) Then
        ' At least one of the optional parameters is present, both must be present
        If IsMissing(OperatorList) Or IsMissing(Criteria2List) Then
            Let AutofilterMatrix = EmptyArray()
            Exit Function
        End If
        
        If NumberOfDimensions(OperatorList) <> 1 Or NumberOfDimensions(Criteria2List) <> 1 Or _
           GetArrayLength(OperatorList) <> GetArrayLength(Criteria1List) Or _
           GetArrayLength(Criteria2List) <> GetArrayLength(Criteria1List) Then
            Let AutofilterMatrix = EmptyArray()
            Exit Function
        End If
    End If
    
    ' If the code gets to this point, parameters are consistent.

    ' Get dimensions of source matrix
    Let NCols = GetNumberOfColumns(TheMatrix)
    Let NRows = GetNumberOfRows(TheMatrix)

    ' Clear any existing contents in worksheet TempComputation
    Call TempComputation.UsedRange.ClearContents

    ' Dump matrix in worksheet TempComputation
    Call DumpInSheet(TheMatrix, TempComputation.Range("A1").Offset(NRows + 1))

    ' Turn the range into a table because this makes it easier to do computations
    Set ListObjRef = TempComputation.ListObjects.Add(SourceType:=xlSrcRange, Source:=TempComputation.Range("A1").Offset(NRows + 1).CurrentRegion, XlListObjectHasHeaders:=xlYes)

    ' Add each of the filtering conditions
    For i = LBound(ColumnsToFilter) To UBound(ColumnsToFilter)
        If IsMissing(Criteria2List) Then
            Call ListObjRef.Range.AutoFilter(Field:=ColumnsToFilter(i), Criteria1:=Criteria1List(i))
        Else
            Call ListObjRef.Range.AutoFilter(Field:=ColumnsToFilter(i), Criteria1:=Criteria1List(i), Operator:=OperatorList(i), Criteria2:=Criteria2List(i))
        End If
    Next i

    Call DumpVisibleCells(ListObjRef.Range, TempComputation.Range("A1"))

    Let AutofilterMatrix = TempComputation.Range("A1").CurrentRegion.Value2

    Call TempComputation.UsedRange.EntireColumn.Delete
End Function


' This returns an array of the same dimensions as the original one but with quotes surrounding
' every element.  This is done for dimensions one or two.  Matrices of higher dimensionality
' are returned unchanged.
Public Function AddQuotesToAllArrayElements(TheArray As Variant, Optional EscapeSingleQuote As Boolean = True) As Variant
    Dim NumberOfDims As Integer
    Dim i As Long
    Dim j As Long
    Dim TheResults As Variant
    
    Let NumberOfDims = NumberOfDimensions(TheArray)
    Let TheResults = TheArray
    
    If NumberOfDims = 1 Then
        For i = LBound(TheResults) To UBound(TheResults)
            If IsError(TheResults(i)) Then
                Let TheResults(i) = "NULL"
            ElseIf Trim(TheResults(i)) = "" Or (TheResults(i) = Empty And Not IsNumeric(TheResults(i))) Or TheResults(i) = "NULL" Then
                Let TheResults(i) = "NULL"
            Else
                Let TheResults(i) = """" & IIf(EscapeSingleQuote, esc(CStr(TheResults(i))), TheResults(i)) & """"
            End If
        Next i
    ElseIf NumberOfDims = 2 Then
        For i = LBound(TheResults, 1) To UBound(TheResults, 1)
            For j = LBound(TheResults, 2) To UBound(TheResults, 2)
                If Trim(TheResults(i, j)) = "" Or (TheResults(i, j) = Empty And Not IsNumeric(TheResults(i, j))) Or TheResults(i, j) = "NULL" Then
                    Let TheResults(i, j) = "NULL"
                Else
                    Let TheResults(i, j) = """" & IIf(EscapeSingleQuote, esc(CStr(TheResults(i, j))), TheResults(i, j)) & """"
                End If
            Next j
        Next i
    End If
        
    Let AddQuotesToAllArrayElements = TheResults
End Function

' Returns an identical array, but with single quotes (e.g. ') surrounding every element
Public Function AddSingleQuotesToAllArrayElements(TheArray As Variant, Optional EscapeSingleQuote As Boolean = True) As Variant
    Dim NumberOfDims As Integer
    Dim i As Long
    Dim j As Long
    Dim TheResults As Variant
    
    Let NumberOfDims = NumberOfDimensions(TheArray)
    Let TheResults = TheArray
    
    If NumberOfDims = 1 Then
        For i = LBound(TheResults) To UBound(TheResults)
            If IsError(TheResults(i)) Then
                Let TheResults(i) = "NULL"
            ElseIf Trim(TheResults(i)) = "" Or TheResults(i) = Empty Or TheResults(i) = "NULL" Then
                Let TheResults(i) = "NULL"
            Else
                Let TheResults(i) = "'" & IIf(EscapeSingleQuote, esc(CStr(TheResults(i))), TheResults(i)) & "'"
            End If
        Next i
    ElseIf NumberOfDims = 2 Then
        For i = LBound(TheResults, 1) To UBound(TheResults, 1)
            For j = LBound(TheResults, 2) To UBound(TheResults, 2)
                If IsError(TheResults(i, j)) Then
                    Let TheResults(i, j) = "NULL"
                ElseIf Trim(TheResults(i, j)) = "" Or (TheResults(i, j) = Empty And Not IsNumeric(TheResults(i, j))) Or TheResults(i, j) = "NULL" Then
                    Let TheResults(i, j) = "NULL"
                Else
                    Let TheResults(i, j) = "'" & IIf(EscapeSingleQuote, esc(CStr(TheResults(i, j))), TheResults(i, j)) & "'"
                End If
            Next j
        Next i
    End If
        
    Let AddSingleQuotesToAllArrayElements = TheResults
End Function

Public Function AddSingleBackQuotesToAllArrayElements(TheArray As Variant, Optional EscapeSingleQuote As Boolean = True) As Variant
    Dim NumberOfDims As Integer
    Dim i As Long
    Dim j As Long
    Dim TheResults As Variant
    
    Let NumberOfDims = NumberOfDimensions(TheArray)
    Let TheResults = TheArray
    
    If NumberOfDims = 1 Then
        For i = LBound(TheResults) To UBound(TheResults)
            If IsError(TheResults(i)) Then
                Let TheResults(i) = "NULL"
            ElseIf Trim(TheResults(i)) = "" Or (TheResults(i) = Empty And Not IsNumeric(TheResults(i))) Or TheResults(i) = "NULL" Then
                Let TheResults(i) = "NULL"
            Else
                Let TheResults(i) = "`" & IIf(EscapeSingleQuote, esc(CStr(TheResults(i))), TheResults(i)) & "`"
            End If
        Next i
    ElseIf NumberOfDims = 2 Then
        For i = LBound(TheResults, 1) To UBound(TheResults, 1)
            For j = LBound(TheResults, 2) To UBound(TheResults, 2)
                If IsError(TheResults(i, j)) Then
                    Let TheResults(i, j) = "NULL"
                ElseIf Trim(TheResults(i, j)) = "" Or (TheResults(i, j) = Empty And Not IsNumeric(TheResults(i, j))) Or TheResults(i, j) = "NULL" Then
                    Let TheResults(i, j) = "NULL"
                Else
                    Let TheResults(i, j) = "`" & IIf(EscapeSingleQuote, esc(CStr(TheResults(i, j))), TheResults(i, j)) & "`"
                End If
            Next j
        Next i
    End If
        
    Let AddSingleBackQuotesToAllArrayElements = TheResults
End Function

' This function acts on a 2D array.  It adds quotes around every element of the given submatrix
Public Function AddQuotesToAllSubArrayElements(TheArray As Variant, LeftIndex As Integer, TopIndex As Integer, _
                                               NRows As Integer, NColumns As Integer, Optional EscapeSingleQuote As Boolean = True) As Variant
    Dim i As Long
    Dim j As Long
    
    Dim TheResults As Variant
    
    Let TheResults = TheArray
    
    For i = TopIndex To TopIndex + NRows - 1
        For j = LeftIndex To LeftIndex + NColumns - 1
            If IsError(TheResults(i, j)) Then
                Let TheResults(i, j) = "NULL"
            ElseIf Trim(TheResults(i, j)) = "" Or (TheResults(i, j) = Empty And Not IsNumeric(TheResults(i, j))) Or TheResults(i, j) = "NULL" Then
                Let TheResults(i, j) = "NULL"
            Else
                Let TheResults(i, j) = """" & IIf(EscapeSingleQuote, esc(CStr(TheResults(i, j))), TheResults(i, j)) & """"
            End If
        Next j
    Next i
        
    Let AddQuotesToAllSubArrayElements = TheResults
End Function

Public Function AddSingleQuotesToAllSubArrayElements(TheArray As Variant, LeftIndex As Integer, TopIndex As Integer, _
                                               NRows As Integer, NColumns As Integer, Optional EscapeSingleQuote As Boolean = True) As Variant
    Dim i As Long
    Dim j As Long
    Dim TheResults As Variant
    
    Let TheResults = TheArray
    
    For i = TopIndex To TopIndex + NRows - 1
        For j = LeftIndex To LeftIndex + NColumns - 1
            If IsError(TheResults(i, j)) Then
                Let TheResults(i, j) = "NULL"
            ElseIf Trim(TheResults(i, j)) = "" Or TheResults(i, j) = "NULL" Or TheResults(i, j) = Empty Then
                Let TheResults(i, j) = "NULL"
            Else
                Let TheResults(i, j) = "'" & IIf(EscapeSingleQuote, esc(CStr(TheResults(i, j))), TheResults(i, j)) & "'"
            End If
        Next j
    Next i
        
    Let AddSingleQuotesToAllSubArrayElements = TheResults
End Function

Public Function AddSingleBackQuotesToAllSubArrayElements(TheArray As Variant, LeftIndex As Integer, TopIndex As Integer, _
                                               NRows As Integer, NColumns As Integer, Optional EscapeSingleQuote As Boolean = True) As Variant
    Dim i As Long
    Dim j As Long
    Dim TheResults As Variant
    
    Let TheResults = TheArray
    
    For i = TopIndex To TopIndex + NRows - 1
        For j = LeftIndex To LeftIndex + NColumns - 1
            If IsError(TheResults(i, j)) Then
                Let TheResults(i, j) = "NULL"
            ElseIf Trim(TheResults(i, j)) = "" Or (TheResults(i, j) = Empty And Not IsNumeric(TheResults(i, j))) Or TheResults(i, j) = "NULL" Then
                Let TheResults(i, j) = "NULL"
            Else
                Let TheResults(i, j) = "`" & IIf(EscapeSingleQuote, esc(CStr(TheResults(i, j))), TheResults(i, j)) & "`"
            End If
        Next j
    Next i
        
    Let AddSingleBackQuotesToAllSubArrayElements = TheResults
End Function

' Converts a 2D array into properly quoted array.  This means that numeric columns are not quoted but others are.  The only
' exception are empty columns, which are converted to NULL.  This returns the properly quoted string.
' This function may be used on the data array before calling sub ConnectAndExecuteInsertQuery() in this module
Public Function DoubleQuote2DArray(TheArray As Variant) As Variant
    Dim r As Long
    Dim c As Long
    Dim ParamCopy As Variant
    
    Let ParamCopy = TheArray
    
    For r = 1 To GetNumberOfRows(ParamCopy)
        For c = 1 To GetNumberOfColumns(ParamCopy)
            If IsError(ParamCopy(r, c)) Then
                Let ParamCopy(r, c) = "NULL"
            ElseIf IsEmpty(ParamCopy(r, c)) Or Trim(ParamCopy(r, c)) = "" Then
                Let ParamCopy(r, c) = "NULL"
            ElseIf TypeName(ParamCopy(r, c)) = "String" And Not IsNumeric(ParamCopy(r, c)) Then
                Let ParamCopy(r, c) = """" & ParamCopy(r, c) & """"
            End If
        Next c
    Next r
    
    Let DoubleQuote2DArray = ParamCopy
End Function

' Given a 2D array like [{1,2,3; 4, 5, 6}], this function returns the following STRING
' (1,2,3), (4,5,6)
'
' This is very useful to create insert statements for databases.
Public Function Convert2DArrayIntoListOfParentheticalExpressions(TheArray As Variant) As String
    Dim TheList As String
    Dim i As Integer

    For i = 1 To GetNumberOfRows(TheArray)
        Let TheList = TheList & Convert1DArrayIntoParentheticalExpression(GetRow(TheArray, CLng(i)))
        
        If i < GetNumberOfRows(TheArray) Then
            Let TheList = TheList & ", "
        End If
    Next i
    
    Let Convert2DArrayIntoListOfParentheticalExpressions = TheList
End Function

' This function resizes a 2D array while preserving values.
' This function returns a variant. AnArray must be a 2D array.
' If either NRows and NCols is less than Ubound(AnArray,1) or Ubound(AnArray,2) then the resulting
' matrix truncates the input matrix accordingly
Public Function Redim2DArray(AnArray As Variant, NRows As Long, NCols As Long) As Variant
    Dim NewArray() As Variant
    Dim r As Long
    Dim c As Long
    
    If NumberOfDimensions(AnArray) <> 2 Then
        Let Redim2DArray = AnArray
        Exit Function
    End If
    
    If Not IsArray(AnArray) Then
        Let Redim2DArray = AnArray
        Exit Function
    End If
    
    If NRows < LBound(AnArray, 1) Or NCols < LBound(AnArray, 2) Then
        Let Redim2DArray = AnArray
        Exit Function
    End If
            
    ReDim NewArray(LBound(AnArray, 1) To IIf(LBound(AnArray, 1) = 0, NRows - 1, NRows), _
                   LBound(AnArray, 2) To IIf(LBound(AnArray, 2) = 0, NCols - 1, NCols))
                   
    For r = LBound(AnArray, 1) To Application.Min(NRows + IIf(LBound(AnArray, 1) = 0, NRows - 1, NRows), UBound(AnArray, 1))
        For c = LBound(AnArray, 2) To Application.Min(NCols + IIf(LBound(AnArray, 2) = 0, NCols - 1, NRows), UBound(AnArray, 2))
            Let NewArray(r, c) = AnArray(r, c)
        Next c
    Next r
    
    Let Redim2DArray = NewArray
End Function

' The purpose of this function is to extend a2Darray1 with data from a2Darray2 using
' equality on the given key columns.  The function returns the "left joined" 2D array.
' This means that all rows in array1 are included. The data from array2 is included
' only if its key in also in array 1.
'
' If a key if found more than once in a2Darray, the first ocurrance is used. The resulting
' 2D array uses data from the columns in a2DArray1 specified in ColsPosArrayFrom2DArray1
' and the columns in a2DArray2 specified in ColsPosArrayFrom2DArray2
'
' When the parameters are inconsistent, the function returns Null
Public Function LeftJoin2DArraysOnKeyEquality(a2DArray1 As Variant, _
                                              a2DArray1KeyColPos As Integer, _
                                              ColsPosArrayFrom2DArray1 As Variant, _
                                              a2DArray2 As Variant, _
                                              a2DArray2KeyColPos As Integer, _
                                              ColsPosArrayFrom2DArray2 As Variant, _
                                              Optional ArraysHaveHeadersQ As Boolean = True, _
                                              Optional IncludeHeadersQ As Boolean = True) As Variant
    Dim ResultsDict As Dictionary
    Dim Array2TrackingDict As Dictionary
    Dim r As Long
    Dim var As Variant
    Dim TheKey As Variant
    Dim TheItems As Variant
    Dim AppendedItems As Variant
    Dim NumColsArray1 As Integer
    Dim NumColsArray2 As Integer
    Dim JoinedHeadersRow As Variant
    Dim TheResults As Variant

    ' Parameter consistency checks
    
    ' Exit with Null if ArraysHaveHeadersQ is False and IncludeHeadersQ = True
    If Not ArraysHaveHeadersQ And IncludeHeadersQ Then
        Let LeftJoin2DArraysOnKeyEquality = Null
        Exit Function
    End If
    
    ' Exit with Null if either a2Array1, a2Array2, ColsPosArrayFrom2DArray1
    ' ColsPosArrayFrom2DArray2 is not an array
    If Not (IsArray(a2DArray1) And IsArray(a2DArray2) And _
            IsArray(ColsPosArrayFrom2DArray1) And IsArray(ColsPosArrayFrom2DArray2)) Then
        Let LeftJoin2DArraysOnKeyEquality = Null
        Exit Function
    End If
    
    ' Exit with Null if a2DArray1 and a2DArray2 don't have the number of dimensions
    If NumberOfDimensions(a2DArray1) <> NumberOfDimensions(a2DArray2) Or _
       NumberOfDimensions(a2DArray1) <> 2 Then
        Let LeftJoin2DArraysOnKeyEquality = Null
        Exit Function
    End If
    
    ' Exit if either a2DArray1KeyColPos or a2DArray2KeyColPos is not an integer
    If (TypeName(a2DArray1KeyColPos) <> TypeName(1) And TypeName(a2DArray1KeyColPos) <> TypeName(99999999)) Or _
       (TypeName(a2DArray2KeyColPos) <> TypeName(1) And TypeName(a2DArray2KeyColPos) <> TypeName(99999999)) Then
        Let LeftJoin2DArraysOnKeyEquality = Null
        Exit Function
    End If
    
    ' Exit with Null if the index of key1 is non-positive or larger than the number
    ' of columns in array1 or the index of key2 is non-positive or larger than the number
    ' of columns in array2
    If a2DArray1KeyColPos < 1 Or a2DArray1KeyColPos > GetNumberOfColumns(a2DArray1) Or _
       a2DArray2KeyColPos < 1 Or a2DArray2KeyColPos > GetNumberOfColumns(a2DArray2) Then
        Let LeftJoin2DArraysOnKeyEquality = Null
        Exit Function
    End If
    
    ' Exit with Null if with ColsPosArrayFrom2DArray1 and ColsPosArrayFrom2DArray2
    ' are not numeric, positive, integer 1D arrays
    If Not (IsPositiveIntegerArrayQ(ColsPosArrayFrom2DArray1) And _
            IsPositiveIntegerArrayQ(ColsPosArrayFrom2DArray2)) Then
        Let LeftJoin2DArraysOnKeyEquality = Null
        Exit Function
    End If

    ' Exit with Null if any of the indices in ColsPosArrayFrom2DArray1 is less than 1
    ' or larger than the number of columns in array1
    For Each var In ColsPosArrayFrom2DArray1
        If var < 1 Or var > GetNumberOfColumns(a2DArray1) Then
            Let LeftJoin2DArraysOnKeyEquality = Null
            Exit Function
        End If
    Next
    
    ' Exit with Null if any of the indices in ColsPosArrayFrom2DArray2 is less than 1
    ' or larger than the number of columns in array2
    For Each var In ColsPosArrayFrom2DArray2
        If var < 1 Or var > GetNumberOfColumns(a2DArray2) Then
            Let LeftJoin2DArraysOnKeyEquality = Null
            Exit Function
        End If
    Next
    
    ' Exit with Null if either a2DArray1 or a2DArray2 has only 1 row when ArraysHaveHeadersQ
    ' is True
    If ArraysHaveHeadersQ And _
       (GetNumberOfRows(a2DArray1) < 2 Or GetNumberOfRows(a2DArray1) < 2) Then
        Let LeftJoin2DArraysOnKeyEquality = Null
        Exit Function
    End If

    ' Determine the number of columns from array 1
    Let NumColsArray1 = GetArrayLength(ColsPosArrayFrom2DArray1)
    
    ' Determine the number of columns from array 2 to append to array 1
    Let NumColsArray2 = GetArrayLength(ColsPosArrayFrom2DArray2)

    ' Load all information from a2DArray1 into a dictionary
    Set ResultsDict = New Dictionary
    For r = IIf(ArraysHaveHeadersQ, LBound(a2DArray1, 1) + 1, LBound(a2DArray1, 1)) To UBound(a2DArray1, 1)
        ' Get the for the current row
        Let TheKey = a2DArray1(r, a2DArray1KeyColPos)

        If Not ResultsDict.Exists(Key:=TheKey) Then
            ' Extract the columns needed from this security's row
            Let TheItems = Take(GetRow(a2DArray1, r), _
                                Prepend(ColsPosArrayFrom2DArray1, a2DArray1KeyColPos))
            
            ' Pad TheItems with enough slots for the items appended from a2DArray2
            Let TheItems = ConcatenateArrays(TheItems, ConstantArray(Empty, CLng(NumColsArray2)))

            ' Add the array of values to this security's entry
            Call ResultsDict.Add(Key:=TheKey, Item:=TheItems)
        End If
    Next r
    
    ' Scan a2DArray2 appending to the array of elements of each element in a2DArray1 the
    ' elements in a2DArray2
    Set Array2TrackingDict = New Dictionary
    For r = IIf(ArraysHaveHeadersQ, LBound(a2DArray2, 1) + 1, LBound(a2DArray2, 1)) To UBound(a2DArray2, 1)
        ' Get the for the current row
        Let TheKey = a2DArray2(r, a2DArray1KeyColPos)
        
        ' Append to the items in the results dicts for the current key if the current key
        ' is in the results dictionary already, and they has not already been appended
        If ResultsDict.Exists(Key:=TheKey) And Not Array2TrackingDict.Exists(Key:=TheKey) Then
            ' Mark this row in a2DArray2 as having been processed
            Call Array2TrackingDict.Add(Key:=TheKey, Item:=Empty)
            
            ' Take the portion of the items corresponding to array 1
            Let TheItems = Take(ResultsDict.Item(Key:=TheKey), 1 + NumColsArray1)
            
            ' Get the required columns from this row to append to those already in the results
            ' dictionary
            Let AppendedItems = Take(GetRow(a2DArray2, r), ColsPosArrayFrom2DArray2)
            
            Let ResultsDict.Item(Key:=TheKey) = ConcatenateArrays(TheItems, AppendedItems)
        End If
    Next r
    
    ' Repack the results as a 2D array
    Let TheResults = Pack2DArray(ResultsDict.Items)
    
    If Not IncludeHeadersQ Then
        Let LeftJoin2DArraysOnKeyEquality = TheResults
    
        Exit Function
    End If
    
    ' Prepend the headers row if the user chose to
    Let JoinedHeadersRow = ConcatenateArrays(Take(GetRow(a2DArray1, 1), _
                                                  Prepend(ColsPosArrayFrom2DArray1, a2DArray1KeyColPos)), _
                                             Take(GetRow(a2DArray2, 1), _
                                                  ColsPosArrayFrom2DArray2))
    
    ' Prepend headers to return matrix
    Let LeftJoin2DArraysOnKeyEquality = Prepend(TheResults, JoinedHeadersRow)
End Function

' The purpose of this function is to extend a2Darray1 with data from a2Darray2 using
' equality on the given key columns.  The function returns the "left joined" 2D array.
' This means that all rows in array1 are included. The data from array2 is included
' only if its key in also in array 1.
'
' If a key if found more than once in a2Darray, the first ocurrance is used.  The resulting
' 2D array uses data from the columns in a2DArray1 specified in ColsPosArrayFrom2DArray1
' and the columns in a2DArray2 specified in ColsPosArrayFrom2DArray2
'
' When the parameters are inconsistent, the function returns Null
Public Function InnerJoin2DArraysOnKeyEquality(a2DArray1 As Variant, _
                                               a2DArray1KeyColPos As Integer, _
                                               ColsPosArrayFrom2DArray1 As Variant, _
                                               a2DArray2 As Variant, _
                                               a2DArray2KeyColPos As Integer, _
                                               ColsPosArrayFrom2DArray2 As Variant, _
                                               Optional ArraysHaveHeadersQ As Boolean = True, _
                                               Optional IncludeHeadersQ As Boolean = True) As Variant
    Dim Array2Dict As Dictionary
    Dim ResultsDict As Dictionary
    Dim r As Long
    Dim var As Variant
    Dim TheKey As Variant
    Dim TheItems As Variant
    Dim AppendedItems As Variant
    Dim NumColsArray1 As Integer
    Dim NumColsArray2 As Integer
    Dim JoinedHeadersRow As Variant
    Dim TheResults As Variant

    ' Parameter consistency checks

    ' Exit with Null if ArraysHaveHeadersQ is False and IncludeHeadersQ = True
    If Not ArraysHaveHeadersQ And IncludeHeadersQ Then
        Let InnerJoin2DArraysOnKeyEquality = Null
        Exit Function
    End If

    ' Exit with Null if either a2Array1, a2Array2, ColsPosArrayFrom2DArray1
    ' ColsPosArrayFrom2DArray2 is not an array
    If Not (IsArray(a2DArray1) And IsArray(a2DArray2) And _
            IsArray(ColsPosArrayFrom2DArray1) And IsArray(ColsPosArrayFrom2DArray2)) Then
        Let InnerJoin2DArraysOnKeyEquality = Null
        Exit Function
    End If

    ' Exit with Null if a2DArray1 and a2DArray2 don't have the number of dimensions
    If NumberOfDimensions(a2DArray1) <> NumberOfDimensions(a2DArray2) Or _
       NumberOfDimensions(a2DArray1) <> 2 Then
        Let InnerJoin2DArraysOnKeyEquality = Null
        Exit Function
    End If

    ' Exit if either a2DArray1KeyColPos or a2DArray2KeyColPos is not an integer
    If (TypeName(a2DArray1KeyColPos) <> TypeName(1) And TypeName(a2DArray1KeyColPos) <> TypeName(99999999)) Or _
       (TypeName(a2DArray2KeyColPos) <> TypeName(1) And TypeName(a2DArray2KeyColPos) <> TypeName(99999999)) Then
        Let InnerJoin2DArraysOnKeyEquality = Null
        Exit Function
    End If

    ' Exit with Null if the index of key1 is non-positive or larger than the number
    ' of columns in array1 or the index of key2 is non-positive or larger than the number
    ' of columns in array2
    If a2DArray1KeyColPos < 1 Or a2DArray1KeyColPos > GetNumberOfColumns(a2DArray1) Or _
       a2DArray2KeyColPos < 1 Or a2DArray2KeyColPos > GetNumberOfColumns(a2DArray2) Then
        Let InnerJoin2DArraysOnKeyEquality = Null
        Exit Function
    End If

    ' Exit with Null if with ColsPosArrayFrom2DArray1 and ColsPosArrayFrom2DArray2
    ' are not numeric, positive, integer 1D arrays
    If Not (IsPositiveIntegerArrayQ(ColsPosArrayFrom2DArray1) And _
            IsPositiveIntegerArrayQ(ColsPosArrayFrom2DArray2)) Then
        Let InnerJoin2DArraysOnKeyEquality = Null
        Exit Function
    End If

    ' Exit with Null if any of the indices in ColsPosArrayFrom2DArray1 is less than 1
    ' or larger than the number of columns in array1
    For Each var In ColsPosArrayFrom2DArray1
        If var < 1 Or var > GetNumberOfColumns(a2DArray1) Then
            Let InnerJoin2DArraysOnKeyEquality = Null
            Exit Function
        End If
    Next

    ' Exit with Null if any of the indices in ColsPosArrayFrom2DArray2 is less than 1
    ' or larger than the number of columns in array2
    For Each var In ColsPosArrayFrom2DArray2
        If var < 1 Or var > GetNumberOfColumns(a2DArray2) Then
            Let InnerJoin2DArraysOnKeyEquality = Null
            Exit Function
        End If
    Next

    ' Exit with Null if either a2DArray1 or a2DArray2 has only 1 row when ArraysHaveHeadersQ
    ' is True
    If ArraysHaveHeadersQ And _
       (GetNumberOfRows(a2DArray1) < 2 Or GetNumberOfRows(a2DArray1) < 2) Then
        Let InnerJoin2DArraysOnKeyEquality = Null
        Exit Function
    End If

    ' Determine the number of columns in arrays 1 and 2
    Let NumColsArray1 = GetArrayLength(ColsPosArrayFrom2DArray1)
    Let NumColsArray2 = GetArrayLength(ColsPosArrayFrom2DArray2)

    ' Index the contents of array2
    Set Array2Dict = New Dictionary
    For r = IIf(ArraysHaveHeadersQ, LBound(a2DArray2, 1) + 1, LBound(a2DArray2, 1)) To UBound(a2DArray2, 1)
        ' Get the for the current row
        Let TheKey = a2DArray2(r, a2DArray2KeyColPos)

        If Not Array2Dict.Exists(Key:=TheKey) Then
            ' Add the array of values to this security's entry
            Call Array2Dict.Add(Key:=TheKey, _
                                Item:=Take(GetRow(a2DArray2, r), ColsPosArrayFrom2DArray2))
        End If
    Next r

    ' Scan a2DArray2 appending to the array of elements of each element in a2DArray1 the
    ' elements in a2DArray2
    Set ResultsDict = New Dictionary
    For r = IIf(ArraysHaveHeadersQ, LBound(a2DArray1, 1) + 1, LBound(a2DArray1, 1)) To UBound(a2DArray1, 1)
        ' Get the for the current row
        Let TheKey = a2DArray1(r, a2DArray1KeyColPos)

        ' Create the join for this row in array1 if it is found in array2 based on the key.
        If Not ResultsDict.Exists(Key:=TheKey) And Array2Dict.Exists(Key:=TheKey) Then
            ' Extract the columns required from this row in array 1
            Let TheItems = Take(GetRow(a2DArray1, r), _
                                Prepend(ColsPosArrayFrom2DArray1, a2DArray1KeyColPos))

            ' Get the corresponding items from array 2
            Let AppendedItems = Array2Dict.Item(Key:=TheKey)

            ' Index the joined items for this row in array 1
            Call ResultsDict.Add(Key:=TheKey, Item:=ConcatenateArrays(TheItems, AppendedItems))
        End If
    Next r

    ' Repack the results as a 2D array
    Let TheResults = Pack2DArray(ResultsDict.Items)

    If Not IncludeHeadersQ Then
        Set InnerJoin2DArraysOnKeyEquality = TheResults

        Exit Function
    End If

    ' Prepend the headers row if the user chose to
    Let JoinedHeadersRow = ConcatenateArrays(Take(GetRow(a2DArray1, 1), _
                                                  Prepend(ColsPosArrayFrom2DArray1, a2DArray1KeyColPos)), _
                                             Take(GetRow(a2DArray2, 1), _
                                                  ColsPosArrayFrom2DArray2))

    ' Prepend headers to return matrix
    Let InnerJoin2DArraysOnKeyEquality = Prepend(TheResults, JoinedHeadersRow)
End Function

' Calculates the dot product of two vectors. Returns Null if the parameters are incompatible.
Public Function DotProduct(v1 As Variant, v2 As Variant) As Variant
    Dim v1prime As Variant
    Dim v2prime As Variant
    Dim i As Long
    Dim TheResult As Double

    If Not (VectorQ(v1) And VectorQ(v2)) Then
        Let DotProduct = Null
        Exit Function
    End If
    
    If Length(v1) <> Length(v2) Then
        Let DotProduct = Null
        Exit Function
    End If
    
    Let v1prime = ConvertTo1DArray(v1)
    Let v2prime = ConvertTo1DArray(v2)
    
    For i = 1 To Length(v1prime)
        Let TheResult = TheResult + v1prime(i) * v2prime(i)
    Next i
    
    Let DotProduct = TheResult
End Function

' Performs matrix multiplication. Both parameters must satisfy VectorQ() or MatrixQ() The function returns Null if the parameters are incompatible.
Public Function MatrixMultiply(m1 As Variant, m2 As Variant) As Variant
    Dim r As Long
    Dim c As Long
    Dim TheResult() As Double
    
    If Not (MatrixQ(m1) And MatrixQ(m2)) Then
        Let MatrixMultiply = Null
        Exit Function
    End If
    
    If NumberOfColumns(m1) <> NumberOfRows(m2) Then
        Let MatrixMultiply = Null
        Exit Function
    End If
    
    ReDim TheResult(1 To NumberOfRows(m1), 1 To NumberOfColumns(m2))
    For r = 1 To NumberOfRows(m1)
        For c = 1 To NumberOfColumns(m2)
            Let TheResult(r, c) = DotProduct(GetRow(m1, r), GetColumn(m2, c))
        Next c
    Next r
    
    Let MatrixMultiply = TheResult
End Function

' Fills in the banks in a 1D array.  Repeats the value in the first cell until it finds a different value.
' It then repeats that one until a new one is found.  And so forth.
Public Function FillArrayBlanks(ByVal AnArray As Variant) As Variant
    Dim CurrentValue As Variant
    Dim c As Long

    If Not DimensionedQ(AnArray) Then
        Let FillArrayBlanks = Null
        Exit Function
    End If
    
    If EmptyArrayQ(AnArray) Then
        Let FillArrayBlanks = EmptyArray()
        Exit Function
    End If

    If IsNull(First(AnArray)) Then
        Let CurrentValue = Empty
    Else
        Let CurrentValue = First(AnArray)
    End If
    
    For c = LBound(AnArray, 1) To UBound(AnArray, 1)
        If IsEmpty(AnArray(c)) Or IsNull(AnArray(c)) Then
            Let AnArray(c) = CurrentValue
        ElseIf AnArray(c) = Empty Then
            Let AnArray(c) = CurrentValue
        ElseIf CurrentValue <> AnArray(c) Then
            Let CurrentValue = AnArray(c)
        End If
    Next c
            
    Let FillArrayBlanks = AnArray
End Function

' DESCRIPTION
' This function return the given array with sequential repeatitions blanked out. It is
' useful to turn arrays into columns for Pivot Table-like arrangements.  For example,
' Array(Empty, Empty, 1, 1, 1, 2, 2, 3) turns into
' Array(Empty, Empty, 1, Empty, Empty, 2, Empty, 3)
' If AnArray fails Predicates.AtomicArrayQ, the fuction returns Null.  The same thing
' happens whenever the parameter makes no sense.
'
' PARAMETERS
' 1. AnArray       - An array satisfying Predicates.AtomicArrayQ
'
' RETURNED VALUE
' AnAtomicArrayOrTable after blanking out sequential repetitions of its elements.
Public Function BlankOutArraySequentialRepetitions(ByVal AnArray As Variant)
    Dim CurrentValue As Variant
    Dim c As Long

    If Not DimensionedQ(AnArray) Then
        Let BlankOutArraySequentialRepetitions = Null
        Exit Function
    End If
    
    If EmptyArrayQ(AnArray) Then
        Let BlankOutArraySequentialRepetitions = EmptyArray()
        Exit Function
    End If
    
    If Not AtomicArrayQ(AnArray) Then
        Let BlankOutArraySequentialRepetitions = Null
        Exit Function
    End If
    
    Let CurrentValue = First(AnArray)
    For c = LBound(AnArray, 1) + 1 To UBound(AnArray, 1)
        If IsEmpty(AnArray(c)) Or IsNull(AnArray(c)) Then
            Let AnArray(c) = Empty
        ElseIf CurrentValue = AnArray(c) Then
            Let AnArray(c) = Empty
        Else
            Let CurrentValue = AnArray(c)
        End If
    Next c
            
    Let BlankOutArraySequentialRepetitions = AnArray
End Function

' DESCRIPTION
' This function inserts new elements in a repeated fashion in between the elements of the
' given array. It can take several forms:
' 1. Riffle(Array(e1, e2, ...), elt)
' 2. Riffle(Array(e1, e2, ...), Array(elt1, elt2, ...))
' 3. Riffle(Array(e1, e2, ...), elt, n)
' 4. Riffle(Array(e1, e2, ...), elt, Array(imin, imax, n))
'    This case requires assumes that 1 instead of 0 is the first array position.
'    It also requires imin<=imax
'
' If there are fewer elements in Array(elt1, elt2, ...) than gaps between  in Riffle(Array(e1,e2,…),Array(x1,x2,…)),
' the Array(elt1, elt2, ...) are used cyclically. Riffle(Array(e),x) gives Array(e). The specification  is of the
' type used in Take. Negative indices count from the end of the list.
'
' In Riffle[list, xlist], if list and xlist are of the same length, then their elements are directly interleaved,
' so that the last element of the result is the last element of xlist.
'
' When A1DArray is an empty array, the empty array is returned unchanged.  When the parameters are inconsistent,
' the function returns Null.
Public Function Riffle(A1DArray As Variant, Arg2 As Variant, Optional StepInterval As Variant) As Variant
    Dim ResultsDict As Dictionary
    Dim var As Dictionary
    Dim r As Long
    Dim s As Long
    Dim c As Long

    If Not DimensionedQ(A1DArray) Then
        Let Riffle = Null
        Exit Function
    End If

    If EmptyArrayQ(A1DArray) Then
        Let Riffle = A1DArray
        Exit Function
    End If
    
    If Not AtomicArrayQ(A1DArray) Then
        Let Riffle = Null
        Exit Function
    End If
    
    If Not IsMissing(StepInterval) Then
        If Not (PositiveWholeNumberQ(StepInterval) Or IsArray(StepInterval)) Then
            Let Riffle = Null
            Exit Function
        End If
        
        If IsArray(StepInterval) Then
            If Not AtomicQ(Arg2) Then
                Let Riffle = Null
                Exit Function
            End If
            
            If Not PositiveWholeNumberArrayQ(StepInterval) And GetArrayLength(StepInterval) <> 3 Then
                Let Riffle = Null
                Exit Function
            End If
            
            If First(StepInterval) > First(Rest(StepInterval)) Then
                Let Riffle = Null
                Exit Function
            End If
        End If
    End If
    
    If Not (AtomicQ(Arg2) Or AtomicArrayQ(Arg2)) Then
        Let Riffle = Null
        Exit Function
    End If
    
    If IsArray(Arg2) Then
        If Not DimensionedQ(Arg2) Then
            Let Riffle = Null
            Exit Function
        End If

        If EmptyArrayQ(Arg2) Then
            Let Riffle = Null
            Exit Function
        End If
        
        If Not AtomicArrayQ(Arg2) Then
            Let Riffle = Null
            Exit Function
        End If
        
        If Not IsMissing(StepInterval) Then
            Let Riffle = Null
            Exit Function
        End If
    End If
    
    ' Case Riffle(Array(e1, e2, ...), elt) and Riffle(Array(e1, e2, ...), elt, n)
    If AtomicArrayQ(A1DArray) And AtomicQ(Arg2) Then
        If IsMissing(StepInterval) Then
            Let StepInterval = 1
        End If
        
        Let s = 1
        Set ResultsDict = New Dictionary
        For r = LBound(A1DArray, 1) To UBound(A1DArray, 1) - 1
            Call ResultsDict.Add(Key:=r, Item:=A1DArray(r))
            
            If s Mod StepInterval = 0 Then
                Call ResultsDict.Add(Key:=r & "-separator", Item:=Arg2)
                Let s = 1
            Else
                Let s = s + 1
            End If
        Next
        Call ResultsDict.Add(Key:=UBound(A1DArray, 1), Item:=A1DArray(UBound(A1DArray, 1)))
        
        Let Riffle = ResultsDict.Items
        Exit Function
    End If
    
    ' Case Riffle(Array(e1, e2, ...), Array(elt1, elt2, ...))
    If AtomicArrayQ(A1DArray) And AtomicArrayQ(Arg2) Then
        Let s = LBound(Arg2, 1)
        Let c = 1
        
        Set ResultsDict = New Dictionary
        For r = LBound(A1DArray, 1) To UBound(A1DArray, 1) - 1
            Call ResultsDict.Add(Key:=c, Item:=A1DArray(r))
            Call ResultsDict.Add(Key:=c + 1, Item:=Arg2(s))
                
            If s = UBound(Arg2, 1) Then
                Let s = LBound(Arg2, 1)
            Else
                Let s = s + 1
            End If
            
            Let c = c + 2
        Next
        Call ResultsDict.Add(Key:=c, Item:=A1DArray(UBound(A1DArray, 1)))
        If GetArrayLength(A1DArray) = GetArrayLength(Arg2) Then
            Call ResultsDict.Add(Key:=c + 1, _
                                 Item:=Arg2(UBound(Arg2, 1)))
        End If
        
        Let Riffle = ResultsDict.Items
        Exit Function
    End If
    
    ' Case Riffle(Array(e1, e2, ...), elt, Array(imin, imax, n))
    Let s = 1
    '***HERE1for r =
End Function




