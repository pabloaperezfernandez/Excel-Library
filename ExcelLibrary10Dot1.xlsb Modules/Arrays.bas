Attribute VB_Name = "Arrays"
' PURPOSE OF THIS MODULE
'
' The purpose of this module is to provide faciities to manipulate and process
' arrays and ranges.

Option Explicit
Option Base 1

' DESCRIPTION
' Returns an array after removing all entries equal to a given value.
' The function returns Null if AnArray is not a dimensioned array.
' The function returns Null if any value in AnArray or TheValue is an object.
'
' PARAMETERS
' 1. AnArray - An array of atomic entries (cannot be objects)
' 2. TheValue - The value to filter out
'
' RETURNED VALUE
' Returns the array with the given values filtered.
Public Function FilterValueFromArray(AnArray As Variant, TheValue As Variant) As Variant
    Dim ResultArray() As Variant
    Dim TheLength As Long
    Dim i As Long
    Dim ThePart As Variant
    
    ' Set default return value when errors encountered
    Let FilterValueFromArray = Null

    ' Exit with Null if AnArray is not an array
    If Not DimensionedQ(AnArray) Or EmptyArrayQ(AnArray) Then Exit Function
    
    ' Examing all array elements, keeping only those that do not equal TheValue
    For i = 1 To Length(AnArray)
        Let ThePart = Part(AnArray, i)
    
        If ThePart <> TheValue Then
            Let TheLength = Length(ResultArray)
            ReDim Preserve ResultArray(1 To TheLength + 1)
            Let ResultArray(TheLength + 1) = ThePart
        End If
    Next
    
    Let FilterValueFromArray = ResultArray
End Function

' DESCRIPTION
' This function takes a ParamArray and copies its elements into a regular variant array.
' If the ParamArray is missing, this function returns an empty array.
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
            If IsObject(Args(0)(i)) Then
                Set ArrayCopy(i) = Args(0)(i)
            Else
                Let ArrayCopy(i) = Args(0)(i)
            End If
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
' This function returns the number of dimensions of any VBA expression.
' Non-array expressions all have 0 dimensions.  Undimemsioned arrays
' have 0 dimensions.
'
' PARAMETERS
' 1. arg - any value or object reference
'
' RETURNED VALUE
' Number of dimensions for its argument.  Dimensioned arrays are the
' only expressions returning non-zero.
Public Function NumberOfDimensions(MyArray As Variant) As Long
    Dim temp As Long
    Dim i As Long
    
    ' Accessing a non-existing dimension throws an error
    ' We count until we raise an error
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
' Alias for Arrays.NumberOfDimensions().  This function returns the
' number of dimensions of any VBA expression. Non-array expressions
' all have 0 dimensions.  Undimemsioned arrays have 0 dimensions.
'
' PARAMETERS
' 1. arg - any value or object reference
'
' RETURNED VALUE
' Number of dimensions for its argument.  Dimensioned arrays are the
' only expressions returning non-zero.
Public Function Dimensions(arg As Variant) As Long
    Let Dimensions = NumberOfDimensions(arg)
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
Public Function Length(AnArray As Variant) As Long
    If Not DimensionedQ(AnArray) Then
        Let Length = 0
    Else
        Let Length = UBound(AnArray, 1) - LBound(AnArray, 1) + 1
    End If
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
Public Function NumberOfRows(AMatrix As Variant) As Long
    If Not DimensionedQ(AMatrix) Then
        Let NumberOfRows = 0
    ElseIf NumberOfDimensions(AMatrix) = 1 Then
        Let NumberOfRows = 1
    Else
        Let NumberOfRows = UBound(AMatrix, 1) - LBound(AMatrix, 1) + 1
    End If
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
Public Function NumberOfColumns(AMatrix As Variant) As Long
    If Not DimensionedQ(AMatrix) Then
        Let NumberOfColumns = 0
    ElseIf NumberOfDimensions(AMatrix) = 1 Then
        Let NumberOfColumns = UBound(AMatrix, 1) - LBound(AMatrix, 1) + 1
    Else
        Let NumberOfColumns = UBound(AMatrix, 2) - LBound(AMatrix, 2) + 1
    End If
End Function

' DESCRIPTION
' Returns the first element in the given array.  Returns Null if the array is empty,
' there is a problem with it, or it has more than two dimensions.  2D arrays are treated
' as an array of arrays, with each row being one of the elements in the array.  In other
' words, this function would return the same thing for [{1,2; 3,4}] and
' Array(Array(1,2), Array(3,4)).
'
' PARAMETERS
' 1. arg - a 1D or 2D array
'
' RETURNED VALUE
' First element in the given array
Public Function First(AnArray As Variant) As Variant
    Let First = Null

    If Not DimensionedQ(AnArray) Then
        Exit Function
    ElseIf NumberOfDimensions(AnArray) = 1 Then
        If EmptyArrayQ(AnArray) Then Exit Function
        
        If IsObject(AnArray(LBound(AnArray))) Then
            Set First = AnArray(LBound(AnArray))
        Else
            Let First = AnArray(LBound(AnArray))
        End If
    ElseIf NumberOfDimensions(AnArray) = 2 Then
        Let First = Part(AnArray, 1)
    End If
End Function

' DESCRIPTION
' Returns the last element in the given array.  Returns Null if the array is empty,
' there is a problem with it, or it has more than two dimensions.  2D arrays are treated
' as an array of arrays, with each row being one of the elements in the array.  In other
' words, this function would return the same thing for [{1,2; 3,4}] and
' Array(Array(1,2), Array(3,4)).
'
' PARAMETERS
' 1. arg - a 1D or 2D array
'
' RETURNED VALUE
' Last element in the given array
Public Function Last(AnArray As Variant) As Variant
    Let Last = Null

    If Not DimensionedQ(AnArray) Then
        Exit Function
    ElseIf NumberOfDimensions(AnArray) = 1 Then
        If EmptyArrayQ(AnArray) Then Exit Function
            
        If IsObject(AnArray(UBound(AnArray))) Then
            Set Last = AnArray(UBound(AnArray))
        Else
            Let Last = AnArray(UBound(AnArray))
        End If
    ElseIf NumberOfDimensions(AnArray) = 2 Then
        Let Last = Part(AnArray, -1)
    Else
        Let Last = Null
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
    Let Most = Null
    
    If Not DimensionedQ(AnArray) Then Exit Function

    If EmptyArrayQ(AnArray) Then
        Let Most = EmptyArray()
    ElseIf UBound(AnArray) = LBound(AnArray) Then
        Let Most = EmptyArray()
    Else
        Let Most = Part(AnArray, Span(1, -2))
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
    Let Rest = Null
    
    If Not DimensionedQ(AnArray) Then Exit Function

    If EmptyArrayQ(AnArray) Then
        Let Rest = EmptyArray()
    ElseIf UBound(AnArray) = LBound(AnArray) Then
        Let Rest = EmptyArray()
    Else
        Let Rest = Part(AnArray, Span(2, -1))
    End If
End Function

' DESCRIPTION
' Flattens a system of nested arrays, regardless of the complexity of array nesting.  The leaves of
' the tree represented by the nested array can have any Excel type. The leaves of the tree must
' be atomic elements (e.g. satisfy AtomicQ) or the function returns Null.  This function flattens
' [{1,2; 3,4}] the same way as Array(Array(1,2), Array(3,4))
'
' PARAMETERS
' 1. arg - any value or object reference
'
' RETURNED VALUE
' A 1D array all the values in the leaves of the tree represented by the original, nested array.
Public Function Flatten(a As Variant, Optional ParameterCheckQ As Boolean = False) As Variant
    Dim var As Variant
    Dim var2 As Variant
    Dim TempVariant As Variant
    Dim ResultsArray() As Variant
    Dim i As Long
    Dim j As Long

    If ParameterCheckQ Then
        If Not IsArray(a) Then
            Let Flatten = Null
            Exit Function
        End If
    End If

    If ParameterCheckQ Then
        If AtomicQ(a) Then
            Let Flatten = a
            Exit Function
        End If
    Else
        If Not IsArray(a) Then
            Let Flatten = a
            Exit Function
        End If
    End If

    If NumberOfDimensions(a) = 2 Then
        Let TempVariant = UnPack2DArray(a)
        Let Flatten = Flatten(TempVariant)
    Else
        Let i = 0
        For Each var In a
            If ParameterCheckQ Then
                If AtomicQ(var) Then
                    Let i = i + 1
                    ReDim Preserve ResultsArray(1 To i)
                    Let ResultsArray(i) = var
                ElseIf IsArray(var) Then
                    Let TempVariant = Flatten(var)
    
                    If ParameterCheckQ Then
                        If IsNull(TempVariant) Then
                            Let Flatten = Null
                            Exit Function
                        End If
                    End If
    
                    For Each var2 In Flatten(var)
                        Let i = i + 1
                        ReDim Preserve ResultsArray(1 To i)
                        Let ResultsArray(i) = var2
                    Next
                Else
                    Let Flatten = Null
                    Exit Function
                End If
            Else
                If Not IsArray(var) Then
                    Let i = i + 1
                    ReDim Preserve ResultsArray(1 To i)
                    Let ResultsArray(i) = var
                Else
                    Let TempVariant = Flatten(var)
    
                    If ParameterCheckQ Then
                        If IsNull(TempVariant) Then
                            Let Flatten = Null
                            Exit Function
                        End If
                    End If
    
                    For Each var2 In Flatten(var)
                        Let i = i + 1
                        ReDim Preserve ResultsArray(1 To i)
                        Let ResultsArray(i) = var2
                    Next
                End If
            End If
        Next

        Let Flatten = ResultsArray
    End If
End Function

' DESCRIPTION
' This function translates the given non-zero index to the array's instrinsic convention.
' For instance, we want to convert TheIndexOrIndexArray=1 into AnArray's intrinsic convention
' that LBound=0 and UBound=2.  Then, this function maps 1 to 0.  Negative indices map -1 to the
' last element of the array. The function returns Null if Abs(TheIndex)>Length(AnArray) or
' Abs(TheIndex)<1.  When the optional parameter RelativeToColumnsQ is explicit set
' to True, this function performs its operations relative to AnArray's columns.
'
' PARAMETERS
' 1. AnArray - An array of any number of dimensions
' 2. TheIndex - a whole number between 1 and Length(AnArray) or between -Length(AnArray) and -1
'
' RETURNED VALUE
' Translated index from lbound = 1 to the array's intrinsic convention
Public Function NormalizeIndex(AnArray As Variant, _
                               TheIndex As Variant, _
                               Optional DimensionIndexRelativeTo As Long = 1, _
                               Optional ParameterCheckQ As Boolean = True) As Variant
    ' Set default return value when encountering errors
    Let NormalizeIndex = Null
                               
    ' This is here for high-speed applications
    If ParameterCheckQ Then
        ' Exit with Null if AnArray is undimensioned
        If Not DimensionedQ(AnArray) Then Exit Function
         
        ' Exit if AnArray is the empty 1D array
        If EmptyArrayQ(AnArray) Then Exit Function
            
        ' Exit with Null if TheIndex is not a positive integer
        If Not NonZeroWholeNumberQ(TheIndex) Then Exit Function
        
        ' Exit with Null if TheIndex is outside of acceptable bounds
        If DimensionIndexRelativeTo > NumberOfDimensions(AnArray) Then Exit Function
        
        If Abs(TheIndex) < 1 Or _
           Abs(TheIndex) > UBound(AnArray, DimensionIndexRelativeTo) - _
                           LBound(AnArray, DimensionIndexRelativeTo) + 1 Then
            Exit Function
        End If
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
        ' In other words, exit with an error code if AnArray is neither an
        ' accetable 1D array or 2D table
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
' This function is meant to operate on 1D or 2D arrays, whose elements may also be 1D or
' 2D arrays, and so recursively. It takes a tuple of any dimension (e.g. an array of
' indices). It takes an array of indices, each one operating deeper recursively. These
' indices are based on LBound=1 and are mapped recursively to each level's intrinsic
' indexing convention.
'
' PARAMETERS
' 1. AnArray - An array of any number of dimensions
' 2. TheTuple - an array of positive, whole numbers between 1 and Length(AnArray)
'    or between -Length(AnArray) and -1
'
' RETURNED VALUE
' Translated index tuple from lbound = 1 to the array's intrinsic convention
Public Function NormalizeIndicesTuple(ByVal AnArray As Variant, _
                                      ByVal TheTuple As Variant) As Variant
    Dim TheArrayElement As Variant
    Dim TheArrayLength As Long
    Dim i As Integer
    Dim var As Variant
    
    ' Set default return value when encountering errors
    Let NormalizeIndicesTuple = Null
                               
    ' Exit with Null if AnArray is undimensioned
    If Not DimensionedQ(AnArray) Then Exit Function
     
    ' Exit if AnArray is the empty 1D array
    If EmptyArrayQ(AnArray) Then Exit Function
        
    ' Exit with Null if TheTuple is not a positive integer
    If Not NonzeroWholeNumberArrayQ(TheTuple) Then Exit Function
    
    ' Handles non-negative TheTuple case
    Let TheArrayElement = AnArray
    Let TheArrayLength = Length(TheArrayElement)
    For i = LBound(TheTuple) To UBound(TheTuple)
        Let var = Part(TheArrayElement, TheTuple(i))
    
        If Abs(TheTuple(i)) < 1 Or Abs(TheTuple(i)) > TheArrayLength Then
            Exit Function
        ElseIf TheTuple(i) > 0 Then
            Let TheTuple(i) = TheTuple(i) + LBound(TheArrayElement) - 1
        ' Handles negative TheTuple case
        Else
            Let TheTuple(i) = TheTuple(i) + UBound(TheArrayElement) + 1
        End If
        
        Let TheArrayElement = var
        Let TheArrayLength = Length(var)
    Next
    
    Let NormalizeIndicesTuple = TheTuple
End Function

' DESCRIPTION
' This function translates an array of index tuples, each satisfying the conditions of
' Public Function NormalizeIndicesTuple(...) in h
'
' PARAMETERS
' 1. AnArray - An array of any number of dimensions
' 2. TheTuple - an array of positive, whole numbers between 1 and Length(AnArray)
'    or between -Length(AnArray) and -1
'
' RETURNED VALUE
' Translated index tuple from lbound = 1 to the array's intrinsic convention
Public Function NormalizeIndicesTupleArray(ByVal AnArray As Variant, _
                                           ByVal ATupleArray As Variant) As Variant
    Dim i As Integer
    Dim var As Variant
    
    ' Set default return value when encountering errors
    Let NormalizeIndicesTupleArray = Null
                               
    ' Exit with Null if AnArray is undimensioned
    If Not DimensionedQ(AnArray) Then Exit Function
     
    ' Exit with Null if AnArray is the empty 1D array
    If EmptyArrayQ(AnArray) Then Exit Function
        
    ' Exit with Null if ATupleArray is not an array of tuples
    For i = LBound(ATupleArray) To UBound(ATupleArray)
        Let var = ATupleArray(i)
        If Not NonzeroWholeNumberArrayQ(var) Then Exit Function
        If EmptyArrayQ(var) Then Exit Function
    Next
    
    ' Normalize each tuple of indices
    For i = LBound(ATupleArray) To UBound(ATupleArray)
        Let ATupleArray(i) = NormalizeIndicesTuple(AnArray, ATupleArray(i))
    Next
    
    Let NormalizeIndicesTupleArray = ATupleArray
End Function

' DESCRIPTION
' This function returns a sequence of numbers based on the user's specifications.
' It has two calling modalities.

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
' The requested numerical sequence
Public Function NumericalSequence(StartNumber As Variant, _
                                  n As Variant, _
                                  Optional TheStep As Variant, _
                                  Optional ToEndNumberQ As Boolean = False) As Variant
    Dim TheStepCopy As Variant
    Dim ReturnArray As Variant
    Dim CurrentNumber As Variant
    Dim i As Long
    
    ' Set default return value for errors
    Let NumericalSequence = Null
    
    ' Set default value for the return array
    Let ReturnArray = Null
    
    ' Set default value for TheStep if missing from function call
    Let TheStepCopy = IIf(IsMissing(TheStep), 1, TheStep)
    
    ' Process calling modality 2
    If ToEndNumberQ Then
        ' Exit with NULL if any of the three parameters is non-numeric
        If Not NumberArrayQ(Array(StartNumber, n, TheStepCopy)) Then Exit Function
        
        ' Set first return value
        ReDim ReturnArray(1 To 1): Let ReturnArray(1) = StartNumber
        
        ' Complete the sequence of numbers
        Let i = 1
        Let CurrentNumber = StartNumber + i * TheStepCopy
        Do While CurrentNumber <= n
            ReDim Preserve ReturnArray(1 To i + 1)
            
            Let ReturnArray(i + 1) = CurrentNumber
            Let i = i + 1
            Let CurrentNumber = StartNumber + i * TheStepCopy
        Loop
    ' Process calling modality 1
    Else
        If n < 0 Then Exit Function
        
        ReDim ReturnArray(1 To n)
        For i = 1 To n
            Let ReturnArray(i) = StartNumber + (i - 1) * TheStepCopy
        Next i
    End If
    
    Let NumericalSequence = ReturnArray
End Function

' DESCRIPTION
' Generates a sequence of numbers. It replicates Mathematica's Range[].
' It has three calling modalities

' 1. Range(ANumber)
'    + When ANumber is positive, returns Array(1, 2, ..., Floor(ANumber))
'    + When ANumber is negative, returns Array()]
' 2. Range(ANumber, EndNumber) - Returns Null if EndNumber<ANumber. Otherwise,
'    it returns Array(ANumber, ANumber+1, ..., Floor(EndNumber))
' 3. Range(ANumber, EndNumber, TheStep)
'    + ANumber<=EndNumber and TheStep>0 returns
'      Array(ANumber, ANumber+TheStep, ANumber+2*TheStep, ..., Floor(EndNumber))
'    + ANumber>EndNumber and TheStep<0 returns
'      Array(ANumber, ANumber-TheStep, ANumber-2*TheStep, ..., Ceiling(EndNumber))
'
' PARAMETERS
' 1. ANumber - A Number
' 2. Optional EndNumber - The end number in the requested sequence when provided
' 3. TheStep (optional) - To create a sequence using a sequential step different
'    from 1. This is required when ANumber > EndNumber.
'
' RETURNED VALUE
' The requested numerical sequence
Public Function Range(ANumber As Variant, _
                      Optional EndNumber As Variant, _
                      Optional TheStep As Variant) As Variant
    ' Set default return value in case of error
    Let Range = Null
                      
    ' Error Check: EndNumber is neither missing nor a number
    If Not (IsMissing(EndNumber) Or NumberQ(EndNumber)) Then Exit Function
    
    ' Error Check: TheStep not missing while EndNumber is missing
    If Not IsMissing(TheStep) And IsMissing(EndNumber) Then Exit Function
    
    ' TheStep is missing and EndNumber < ANumber
    If Not IsMissing(EndNumber) And IsMissing(TheStep) Then
        If EndNumber < ANumber Then
            Let Range = EmptyArray()
        End If
    End If
    
    ' EndNumber is missing and ANumber is negative
    If IsMissing(EndNumber) And NegativeQ(ANumber) Then
        Let Range = EmptyArray()
        Exit Function
    End If
    
    ' ANumber is given.  Nothing else given
    If IsMissing(EndNumber) Then
        Let Range = NumericalSequence(1, ANumber, 1, True)
        Exit Function
    End If
    
    ' ANumber<=EndNumber
    If EndNumber >= ANumber Then
        Let Range = NumericalSequence(ANumber, EndNumber, IIf(IsMissing(TheStep), 1, TheStep), True)
        Exit Function
    End If
    
    ' EndNumber<=ANumber and TheStep is given
    Let Range = NumericalSequence(ANumber, EndNumber, TheStep, True)
End Function

' DESCRIPTION
' This function returns the sequence of indices specified by the given Span instance relative
' to the given array and dimensional index.  This function gives an easy way to translate
' indices from our base 1 convention to the array's intrinsic convertion. For instance, while
' we take 1 to refer to the array's first element, the array might employ base 0.  Hence, this
' function translates 1 to 0.  It takes an array and a span instance, and translates the indices
' specified by the span into the array's indexing convention.
'
' PARAMETERS
' 1. AnArray - A dimensioned, nonempty array
' 2. ASpan - An instance of class Span
' 3. TheDimension - The dimension relative to which generate the indices sequence
'
' RETURNED VALUE
' The requested indices sequence
Public Function CreateIndexSequenceFromSpan(AnArray As Variant, _
                                            ASpan As Variant, _
                                            Optional TheDimension As Long = 1) As Variant
    Dim TheStart As Variant
    Dim TheEnd As Variant
    Dim c As Long
    Dim ReturnArray() As Long

    ' Set default return value when encountering errors
    Let CreateIndexSequenceFromSpan = Null
    
    If TypeName(ASpan) <> TypeName(New Span) Then Exit Function
    
    ' Exit with Null if AnArray is undimensioned or empty
    If Not DimensionedQ(AnArray) Or EmptyArrayQ(AnArray) Then Exit Function
    
    ' Exit with Null if TheDimension>NumberOfDimensions(AnArray)
    If TheDimension < 1 Or TheDimension > NumberOfDimensions(AnArray) Then Exit Function

    ' Turn the spans start and end points into positive indices relative
    ' to the array's intrinsic convention
    Let TheStart = NormalizeIndex(AnArray, ASpan.TheStart, TheDimension)
    Let TheEnd = NormalizeIndex(AnArray, ASpan.TheEnd, TheDimension)
    
    If IsNull(TheStart) Or IsNull(TheEnd) Then Exit Function
    If TheStart > TheEnd Or ASpan.TheStep <= 0 Then Exit Function
    
    ' Complete the sequence of numbers
    Let c = 0
    Do While TheStart + c * ASpan.TheStep <= TheEnd
        ReDim Preserve ReturnArray(1 To c + 1)
        
        Let ReturnArray(c + 1) = TheStart + c * ASpan.TheStep
        Let c = c + 1
    Loop
    
    Let CreateIndexSequenceFromSpan = ReturnArray
End Function

' DESCRIPTION
' This function returns the requested part of an array.  It works just like Mathematica's
' Part[]. The returned value depends on the form of parameter Indices.  Works on 1D and 2D
' arrays.  Use the function ClassConstructors.Span() wherever you would use Mathematica Span
' such as All and 1;;2, ;;2, 3;;.  For example, to get every index from the last to the 2nd,
' use 2;;-1.
'
' PARAMETERS
' 1. AnArray - A dimensioned array
' 2. Indices - a sequence of indices (with at least one supplied) using the forms below, with
'    each one referring to a different dimension of the array. At the moment we process only
'    1D and 2D arrays. So, Indices can only be one or two of the forms below.
'
' Indices can take any of the following forms:
' 1. n - to get element n.  If given a 2D array, n refers to the row number
' 2. n_1, n_2 - to get element with index (n_1, n_2)
' 3. Array() - A dimensioned, non-empty array of indices
' 4. Array1, Array2 - Two arrays of the type in #3, one for each dimension
' 5. Span - An instance of class Span, which can be conveniently generated using
'           ClassConstructors.Span()
' 6. Span_1, Span_2 - Each Span is an in #3 above, with _1 and _2 applying to dimensions 1 and 2
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
'
' Examples:
'
' 1. Part(anArray, n) - Returns the nth element of the array
' 2. Part(anArray, n1, n2) - Returns element (n1,n2) from 2D array anArray
' 3. Part(anArray, Array(n1,n2, ..., n_m)) - Returns an array with elements
'    n1, n2, ..., n_m of anArray
' 4. Part(anArray, Span(1,-2)) - Returns array with elements 1 through first before the last
' 5. Part([{1,2,3;4,5,6;7,8,9}], Span(1,3), Array(1,3)) - Returns a matrix with columns 1 and 3
'    of all rows.
Public Function Part(AnArray As Variant, ParamArray Indices() As Variant) As Variant
    Dim IndicesCopy As Variant
    Dim IndexIndex As Variant
    Dim AnIndex As Variant
    Dim IndexArray As Variant
    Dim ReturnArray As Variant
    Dim ni As Variant
    Dim ir As Long ' for nornalized index r
    Dim ic As Long ' for normalized index c
    Dim r As Long ' for matrix row
    Dim c As Long ' for matrix column
    Dim RowIndices As Variant
    Dim ColumnIndices As Variant
    Dim var As Variant
    Dim ASpan As Span
    
    ' Set default return value when errors encountered
    Let Part = Null

    ' Exit with Null if AnArray is not an array
    If Not DimensionedQ(AnArray) Or EmptyArrayQ(AnArray) Then Exit Function
    
    ' Convert ParamArray to a regular array
    Let IndicesCopy = CopyParamArray(Indices)

    ' Exit if more indices were passed than the number of dimensions in AnArray
    If NumberOfDimensions(AnArray) < Length(IndicesCopy) Then Exit Function
    
    ' Exit with Null if any of the indices fails PartIndexQ
    If Not PartIndexArrayQ(IndicesCopy) Then Exit Function
    
    ' Loop over dimensions converting index specifications into an array of individual
    ' element positions or arrays of individual element positions
    For r = 1 To Length(IndicesCopy)
        ' Convert from out lbound=1 convention to IndicesCopy's intrinsic convention
        Let IndexIndex = NormalizeIndex(IndicesCopy, r)
        
        ' Get the current dimensional index
        If IsObject(IndicesCopy(IndexIndex)) Then
            Set AnIndex = IndicesCopy(IndexIndex)
            Set ASpan = AnIndex
        Else
            Let AnIndex = IndicesCopy(IndexIndex)
        End If
    
        ' Compute indices sequence for this dimension (e.g. r) from its specification
        If NonZeroWholeNumberQ(AnIndex) Then
            Let IndicesCopy(IndexIndex) = NormalizeIndex(AnArray, AnIndex, r)
        ElseIf NonzeroWholeNumberArrayQ(AnIndex) And NonEmptyArrayQ(AnIndex) Then
            Let IndicesCopy(IndexIndex) = NormalizeIndexArray(AnArray, AnIndex, r, False)
        ElseIf SpanQ(AnIndex) Then
            Let IndicesCopy(IndexIndex) = CreateIndexSequenceFromSpan(AnArray, ASpan, r)
        Else
            Let IndicesCopy(IndexIndex) = Null
        End If
        
        ' Check for invalid indices and exit if one is found
        If NullQ(IndicesCopy(IndexIndex)) Then Exit Function
        
        If IsArray(IndicesCopy) Then
            If AnyTrueQ(IndicesCopy(IndexIndex), "NullQ") Then Exit Function
        End If
    Next
    
    ' Collect the chosen array's slice depending on its number of dimensions
    Select Case NumberOfDimensions(AnArray)
        Case 1
            ' Process first index specification
            Let ColumnIndices = First(IndicesCopy)
        
            ' A single element was requested
            If WholeNumberQ(ColumnIndices) Then
                If IsObject(AnArray(ColumnIndices)) Then
                    Set ReturnArray = AnArray(ColumnIndices)
                Else
                    Let ReturnArray = AnArray(ColumnIndices)
                End If
            ' A sequence of elements was requested
            Else
                ' Pre-allocate a 1D array since AnArray is one-dimensional
                ReDim ReturnArray(1 To Length(ColumnIndices))
                
                ' Extract the requested elements from AnArray
                For c = 1 To Length(ColumnIndices)
                    Let ic = NormalizeIndex(ColumnIndices, c)
                    If IsObject(AnArray(ColumnIndices(ic))) Then
                        Set ReturnArray(c) = AnArray(ColumnIndices(ic))
                    Else
                        Let ReturnArray(c) = AnArray(ColumnIndices(ic))
                    End If
                Next
            End If
            
            ' You could now recurse on each element of ReturnArray if more indices were given
            ' For now, we have avoid doing this
        Case 2
            ' Part was called to extract complete rows from this 2D array
            If Length(IndicesCopy) = 1 Then
                Let RowIndices = First(IndicesCopy)
                
                If WholeNumberQ(RowIndices) Then
                    ReDim ReturnArray(1 To NumberOfColumns(AnArray))
                    For c = 1 To NumberOfColumns(AnArray)
                        Let ic = NormalizeIndex(AnArray, c, 2)
                        
                        If IsObject(AnArray(RowIndices, ic)) Then
                            Set ReturnArray(c) = AnArray(RowIndices, ic)
                        Else
                            Let ReturnArray(c) = AnArray(RowIndices, ic)
                        End If
                    Next
                Else
                    ReDim ReturnArray(1 To Length(RowIndices), 1 To NumberOfColumns(AnArray))
                    For r = 1 To Length(RowIndices)
                        For c = 1 To NumberOfColumns(AnArray)
                            Let ir = NormalizeIndex(RowIndices, r)
                            Let ic = NormalizeIndex(AnArray, c, 2)
                        
                            If IsObject(AnArray(RowIndices(ir), ic)) Then
                                Set ReturnArray(r, c) = AnArray(RowIndices(ir), ic)
                            Else
                                Let ReturnArray(r, c) = AnArray(RowIndices(ir), ic)
                            End If
                        Next
                    Next
                End If
            ' Get all elements requested
            Else
                Let RowIndices = First(IndicesCopy)
                Let ColumnIndices = Last(IndicesCopy)
                
                If WholeNumberQ(RowIndices) And WholeNumberQ(ColumnIndices) Then
                    Let ReturnArray = AnArray(RowIndices, ColumnIndices)
                ElseIf WholeNumberQ(RowIndices) And IsArray(ColumnIndices) Then
                    ReDim ReturnArray(1 To Length(ColumnIndices))
                    For c = 1 To Length(ColumnIndices)
                        Let ic = ColumnIndices(NormalizeIndex(ColumnIndices, c))
                        
                        If IsObject(AnArray(RowIndices, ic)) Then
                            Set ReturnArray(c) = AnArray(RowIndices, ic)
                        Else
                            Let ReturnArray(c) = AnArray(RowIndices, ic)
                        End If
                    Next
                ElseIf IsArray(RowIndices) And WholeNumberQ(ColumnIndices) Then
                    ReDim ReturnArray(1 To Length(RowIndices))
                    For r = 1 To Length(RowIndices)
                        Let ir = RowIndices(NormalizeIndex(RowIndices, r))
                        
                        If IsObject(AnArray(ir, ColumnIndices)) Then
                            Set ReturnArray(r) = AnArray(ir, ColumnIndices)
                        Else
                            Let ReturnArray(r) = AnArray(ir, ColumnIndices)
                        End If
                    Next
                Else
                    ReDim ReturnArray(1 To Length(RowIndices), 1 To Length(ColumnIndices))
                    For ir = 1 To Length(RowIndices)
                        For ic = 1 To Length(ColumnIndices)
                            If IsObject(AnArray(RowIndices(NormalizeIndex(RowIndices, ir)), _
                                                ColumnIndices(NormalizeIndex(ColumnIndices, ic)))) Then
                                Set ReturnArray(ir, ic) = AnArray(RowIndices(NormalizeIndex(RowIndices, ir)), _
                                                                  ColumnIndices(NormalizeIndex(ColumnIndices, ic)))
                            Else
                                Let ReturnArray(ir, ic) = AnArray(RowIndices(NormalizeIndex(RowIndices, ir)), _
                                                                  ColumnIndices(NormalizeIndex(ColumnIndices, ic)))
                            End If
                        Next
                    Next
                End If
            End If
        Case Else
            Exit Function
    End Select
    
    If IsObject(ReturnArray) Then
        Set Part = ReturnArray
    Else
        Let Part = ReturnArray
    End If
End Function

' DESCRIPTION
' Returns the subset of the 1D or 2D array specified by the indices.  Most common uses are:
'
' a. Take(m, n) - with n>0 returns the first n elements or rows of m
' b. Take(m, -n) - with n>0 returns the last n elements or rows of m
'
' A big difference with Arrays.Part() is that it always returns a set instead of single
' elements as in the case of Part(AnArray, n) or Part(AnArray, n, m), which return single
' elements when n and m are whole numbers.
'
' Where as Mathematica's Take can operate on rows and columns, our implementation operates
' one the first dimension exclusively.  Use Part with Span() if you want to operate on more
' than one dimension simultaneously,
'
' PARAMETERS
' 1. AnArray - A dimensioned array
' 2. Indices - a sequence of indices (with at least one supplied) of the forms below, with
'    each one referring to a different dimension of the array. At the moment we process only
'    1D and 2D arrays. So, Indices can only be one or two of the forms below.
'
' Indices can take any of the following forms:
' 1. n - Get elements 1 through
' 2. -n - Elements from the end of the array indexed by -1 to -n from right to left
' 3. [{n}] - Element n only.  Works just like Part(AnArray, n)
' 4. [{n_1, n_2}] - Elements n_1 through n_2
' 5. [{n_1, n_2, TheStep}] - Elements n_1 through n_2 every TheStep elements. Identical to
'    Part(AnArray, Span(n_1, n_2, TheStep)
'
' RETURNED VALUE
' The requested slice or element of the array.
Public Function Take(AnArray As Variant, Indices As Variant) As Variant
    Let Take = Null
    
    If Not DimensionedQ(AnArray) Then Exit Function
    
    If Not (NonZeroWholeNumberQ(Indices) Or NonzeroWholeNumberArrayQ(Indices)) Then
        Exit Function
    End If
    
    If WholeNumberArrayQ(Indices) And (Length(Indices) < 1 Or Length(Indices) > 3) Then
        Exit Function
    End If
    
    If WholeNumberQ(Indices) Then
        If Indices >= 1 Then
            Let Take = Part(AnArray, Span(1, CLng(Indices)))
        Else
            Let Take = Part(AnArray, Span(CLng(Indices), -1))
        End If
    ElseIf WholeNumberArrayQ(Indices) And Length(Indices) = 1 Then
        Let Take = Part(AnArray, First(Indices))
    ElseIf WholeNumberArrayQ(Indices) And Length(Indices) = 2 Then
        Let Take = Part(AnArray, Span(CLng(First(Indices)), CLng(Last(Indices))))
    Else
        Let Take = Part(AnArray, Span(CLng(First(Indices)), CLng(Part(Indices, 2)), CLng(Last(Indices))))
    End If
End Function

' DESCRIPTION
' Stacks two 1D or 2D arrays on top of each other.  1D arrays are allowed and interpreted
' as single-row 2D arrays.  The two arrays must have the same number of columns or the function
' returns Null.  If either is empty, the function returns the other array.  The resulting 2D
' array has lbounds = 1.
'
' PARAMETERS
' 1. A - a 1D or 2D array
' 2. B - a 1D or 2D array
' 3. ParameterCheckQ - (optional) When explicitly set to False, no parameter checks are done
'
' RETURNED VALUE
' Returns the stacked 2D array resulting from stacking the two given arrays.
Public Function StackArrays(a As Variant, _
                            b As Variant, _
                            Optional ParameterCheckQ As Boolean = True) As Variant
    Dim r As Long
    Dim c As Long
    Dim DimsA As Long
    Dim DimsB As Long
    Dim NumColsA As Long
    Dim NumRowsA As Long
    Dim NumColsB As Long
    Dim NumRowsB As Long
    Dim ReturnArray As Variant
    
    ' Set default return value
    Let StackArrays = Null
    
    ' Check parameter consistency if ParameterCheckQ is True
    If ParameterCheckQ Then
        If Not (DimensionedQ(a) Or DimensionedQ(b)) Then Exit Function
        
        If Not (AtomicArrayQ(a) Or AtomicTableQ(a)) Or _
           Not (AtomicArrayQ(a) Or AtomicTableQ(a)) Then Exit Function
        
        If NumberOfColumns(a) <> NumberOfColumns(b) Then Exit Function
        
        If EmptyArrayQ(a) Then
            Let StackArrays = b
            Exit Function
        End If
        
        If EmptyArrayQ(b) Then
            Let StackArrays = a
            Exit Function
        End If
    End If

    ' Get dimensions of a and b
    Let DimsA = NumberOfDimensions(a)
    Let DimsB = NumberOfDimensions(b)
    Let NumColsA = NumberOfColumns(a)
    Let NumColsB = NumberOfColumns(b)
    Let NumRowsA = NumberOfRows(a)
    Let NumRowsB = NumberOfRows(b)
    
    ' Stack two 1D arrays
    If DimsA = 1 And DimsB = 1 Then
        ' Pre-allocate big enough 2D array
        ReDim ReturnArray(1 To 2, 1 To NumColsA)
        
        ' Add A's lone row and B's lone row in rows 1 and 2 respectively of the stacked array
        For c = 1 To NumColsA
            Let ReturnArray(1, c) = a(c + LBound(a) - 1)
            Let ReturnArray(2, c) = b(c + LBound(a) - 1)
        Next c
    ' Stack 1D array A on top of 2D array B
    ElseIf DimsA = 1 And DimsB > 1 Then
        ' Pre-allocate big enough 2D array
        ReDim ReturnArray(1 To NumRowsB + 1, 1 To NumColsB)
        
        ' Set A's lone row as the first row of the stacked array
        For c = 1 To NumColsA
            Let ReturnArray(1, c) = a(c + LBound(a) - 1)
        Next c
        
        For r = 1 To NumRowsB
            For c = 1 To NumColsB
                Let ReturnArray(1 + r, c) = b(r + LBound(b, 1) - 1, c + LBound(b, 2) - 1)
            Next c
        Next r
    ' Stack 2D array A on top of 1D array B
    ElseIf DimsA > 1 And DimsB = 1 Then
        ' Pre-allocate big enough 2D array
        ReDim ReturnArray(1 To NumRowsA + 1, 1 To NumColsA)
        
        For r = 1 To NumRowsA + 1
            For c = 1 To NumColsA
                If r < NumRowsA + 1 Then
                    Let ReturnArray(r, c) = a(r + LBound(a, 1) - 1, c + LBound(a, 2) - 1)
                Else
                    Let ReturnArray(r, c) = b(c + LBound(b) - 1)
                End If
            Next c
        Next r
    ' Stack 2D array A on top of 2D array B
    Else
        ' Pre-allocate big enough 2D array
        ReDim ReturnArray(1 To NumRowsA + NumRowsB, 1 To NumColsB)
        For r = 1 To NumRowsA
            For c = 1 To NumColsA
                Let ReturnArray(r, c) = a(r + LBound(a, 1) - 1, c + LBound(a, 2) - 1)
            Next c
        Next r
    
        For r = 1 To NumRowsB
            For c = 1 To NumColsB
                Let ReturnArray(NumRowsA + r, c) = b(r + LBound(b, 1) - 1, c + LBound(b, 2) - 1)
            Next c
        Next r
    End If
    
    ' Return the stacked arrays
    Let StackArrays = ReturnArray
End Function

' DESCRIPTION
' Appends an element to the given 1D or 2D array.  To be appended to a 2D array, the element must
' must be 1D or 2D array with the same number of columns.  The function returns Null in all other
' cases.
'
' PARAMETERS
' 1. AnArray - A 1D or 2D array to which AnElt should be appended
' 2. AnElt - a 1D or 2D array to append to AnArray
' 3. ParameterCheckQ - (optional) When explicitly set to False, no parameter checks are done
'
' RETURNED VALUE
' Returns the array that results from appending an element to the an array
Public Function Append(AnArray As Variant, _
                       AnElt As Variant, _
                       Optional ParameterCheckQ As Boolean = True) As Variant
    Dim NewArray As Variant
    
    ' Set default return value
    Let Append = Null
    
    ' Check parameter consitency if ParameterCheckQ is True
    If ParameterCheckQ Then
        ' Exit if AnArray is not dimensioned
        If Not DimensionedQ(AnArray) Then Exit Function
            
        ' Exit with Null if AnArray is a 2D array and AnElt does not have the same number of columns
        If NumberOfDimensions(AnArray) = 2 And NumberOfColumns(AnArray) <> NumberOfColumns(AnElt) Then Exit Function
        
        ' Exit with AnElt if AnArray is empty
        If EmptyArrayQ(AnArray) Then
            Let Append = Array(AnElt)
            Exit Function
        End If
        
        ' Exit with AnArray if AnElt is empty
        If EmptyArrayQ(AnElt) Then
            Let Append = AnArray
            Exit Function
        End If
        
        ' Exit with Null if AnElt is Null
        If NullQ(AnElt) Then Exit Function
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
    Let Append = StackArrays(AnArray, AnElt, ParameterCheckQ)
End Function

' DESCRIPTION
' Prepends an element to the given 1D or 2D array.  To be prepended to a 2D array, the element
' must be 1D or 2D array with the same number of columns.  The function returns Null in all
' other cases.
'
' PARAMETERS
' 1. AnArray - A 1D or 2D array to which AnElt should be appended
' 2. AnElt - a 1D or 2D array to append to AnArray
' 3. ParameterCheckQ - (optional) When explicitly set to False, no parameter checks are done
'
' RETURNED VALUE
' Returns the array resylt from prepending an element to an array.
Public Function Prepend(AnArray As Variant, _
                        AnElt As Variant, _
                        Optional ParameterCheckQ As Boolean = True) As Variant
    Dim ReturnArray() As Variant
    Dim c As Long

    ' Check parameter consitency if ParameterCheckQ is True
    If ParameterCheckQ Then
        ' Exit if AnArray is not dimensioned
        If Not DimensionedQ(AnArray) Then Exit Function
            
        ' Exit with Null if AnArray is a 2D array and AnElt does not have the same number of
        ' columns
        If NumberOfDimensions(AnArray) = 2 And _
           NumberOfColumns(AnArray) <> NumberOfColumns(AnElt) Then
            Exit Function
        End If
        
        ' Exit with AnElt if AnArray is empty
        If EmptyArrayQ(AnArray) Then
            Let Prepend = Array(AnElt)
            Exit Function
        End If
        
        ' Exit with AnArray if AnElt is empty
        If EmptyArrayQ(AnElt) Then
            Let Prepend = AnArray
            Exit Function
        End If
        
        ' Exit with Null if AnElt is Null
        If NullQ(AnElt) Then Exit Function
    End If
    
    ' If AnArray is 1D, then put the new element (whatever it may be) as the last element of a
    ' 1D array 1 longer than the original one.
    If NumberOfDimensions(AnArray) = 1 Then
        ReDim ReturnArray(1 To Length(AnArray) + 1)
        
        ' Prepend the element to the return array
        Let ReturnArray(1) = AnElt
        
        ' Copy the array to prepend to
        For c = 1 To Length(AnArray)
            Let ReturnArray(c + 1) = AnArray(LBound(AnArray) + c - 1)
        Next c
        
        ' Set the function to return the appended array
        Let Prepend = ReturnArray
        
    Else
        ' AnArray and AnElt have the same number of columns.  Stack them in the right order to
        ' prepend
        Let Prepend = StackArrays(AnElt, AnArray)
    End If
End Function

' DESCRIPTION
' Prepends an element to the given 1D or 2D array.  To be prepended to a 2D array, the element
' must be 1D or 2D array with the same number of columns.  The function returns Null in all
' other cases.
'
' PARAMETERS
' 1. AnArray - A 1D or 2D array to which AnElt should be appended
' 2. AnElt - a 1D or 2D array to append to AnArray
' 3. ParameterCheckQ - (optional) When explicitly set to False, no parameter checks are done
'
' RETURNED VALUE
' Returns the array resylt from prepending an element to an array.
Public Function Reverse(AnArray As Variant, Optional DimensionalityIndex As Integer = 1)
    Dim ReturnArray As Variant
    Dim i As Long
    Dim j As Long

    Let Reverse = Null
    
    If Not DimensionedQ(AnArray) Then Exit Function
    
    If EmptyArrayQ(AnArray) Then
        Let Reverse = EmptyArray
        Exit Function
    End If
    
    If NumberOfDimensions(AnArray) > 2 Then Exit Function
    
    If DimensionalityIndex > NumberOfDimensions(AnArray) Then Exit Function
    
    ' AnArray may be 1D or 2D.  Reverse along the first dimensionality index in either case.
    If DimensionalityIndex = 1 Then
        ' Reverse the 1D array
        If NumberOfDimensions(AnArray) = 1 Then
            ReDim ReturnArray(LBound(AnArray) To UBound(AnArray))
            
            For i = LBound(AnArray) To UBound(AnArray)
                Let ReturnArray(LBound(AnArray) - i + UBound(AnArray)) = AnArray(i)
            Next
        ' Reverse the 2D array with respect to its rows
        Else
            ReDim ReturnArray(LBound(AnArray, 1) To UBound(AnArray, 1), _
                              LBound(AnArray, 2) To UBound(AnArray, 2))
            
            For i = LBound(AnArray, 1) To UBound(AnArray, 1)
                For j = LBound(AnArray, 2) To UBound(AnArray, 2)
                    Let ReturnArray(LBound(AnArray, 1) - i + UBound(AnArray, 1), j) = _
                        AnArray(i, j)
                Next j
            Next i
        End If
    ' Reverse the 2D array with respect to the second dimensionality index
    Else
        ReDim ReturnArray(LBound(AnArray, 1) To _
              UBound(AnArray, 1), LBound(AnArray, 2) To UBound(AnArray, 2))
        
        For i = LBound(AnArray, 1) To UBound(AnArray, 1)
            For j = LBound(AnArray, 2) To UBound(AnArray, 2)
                Let ReturnArray(i, LBound(AnArray, 2) - j + UBound(AnArray, 2)) = AnArray(i, j)
            Next j
        Next i
    End If
    
    ' Return the reversed array
    Let Reverse = ReturnArray
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
Public Function PrintArray(TheArray As Variant, _
                           Optional SupressOutputQ As Boolean = False) As String
    Dim ReturnString As String
    Dim ARow As Variant
    Dim c As Long
    Dim r As Long
    
    If IsNull(TheArray) Then
        If Not SupressOutputQ Then Debug.Print "Null"
        Let PrintArray = vbNullString
        Exit Function
    ElseIf Not IsArray(TheArray) Then
        Let ReturnString = "Not an array"
        If Not SupressOutputQ Then Debug.Print ReturnString
    ElseIf NumberOfDimensions(TheArray) = 0 Then
        Let ReturnString = TheArray
        If Not SupressOutputQ Then Debug.Print ReturnString
    ElseIf NumberOfDimensions(TheArray) = 1 Then
        If EmptyArrayQ(TheArray) Then
            Let ReturnString = "Empty 1D Array"
            If Not SupressOutputQ Then Debug.Print ReturnString
        Else
            If UBound(TheArray) >= LBound(TheArray) Then
                For c = LBound(TheArray) To UBound(TheArray)
                    Let ARow = ARow & _
                               IIf(ARow = vbNullString, vbNullString, vbTab) & TheArray(c)
                Next c
                
                Let ReturnString = ReturnString & _
                                   IIf(ReturnString = vbNullString, vbNullString, vbCr) & ARow
            End If
        
            If Not SupressOutputQ Then Debug.Print ARow
        End If
    Else
        For r = LBound(TheArray, 1) To UBound(TheArray, 1)
            Let ARow = TheArray(r, LBound(TheArray, 2))
        
            If UBound(TheArray, 2) - LBound(TheArray, 2) >= 1 Then
                For c = LBound(TheArray, 2) + 1 To UBound(TheArray, 2)
                    Let ARow = ARow & _
                               IIf(ARow = vbNullString, vbNullString, vbTab) & _
                               TheArray(r, c)
                Next c
                
                Let ReturnString = ReturnString & _
                                   IIf(ReturnString = vbNullString, vbNullString, vbCr) & _
                                   ARow
            End If
        
            If Not SupressOutputQ Then Debug.Print ARow
        Next r
    End If
    
    Let PrintArray = ReturnString
End Function

' DESCRIPTION
' Returns the array resulting from inserting the given element at the given position.  It may
' be invoked with any of the following forms:
'
' a. Insert(AnArray, TheElement, n) - Returns the array obtained from inserting TheElement in
'    position n of AnArray. n may range from 1 to Length(AnArray).  The list is shifted to the
'    right and TheElement is inserted in position 1 if n = 1.  If n < Length(AnArray), elements
'    from n to the end of the array are shifted to the right and TheElement is inserted in
'    position n.  TheElement is appened to the array if n = Length(AnArray)+1. If AnArray is a
'    2D array satisfying Predicates.AtomicTableQ
' b. Insert(AnArray, TheElement, Array(n_1, ..., n_k)) - Inserts TheElement in positions n_1,
'    n_2, ..., and n_k per (a) above.  The rows of a 2D array are considered its elements for
'    the purpose of Insert.
'
' PARAMETERS
' 1. AnArray - A dimensioned array
' 2. TheElement - The element to insert in AnArray
' 3. ThePositions - The index of array of indices where TheElement should be inserted
'
' RETURNED VALUE
' AnArray with TheElement inserted in the requested positions
Public Function Insert(AnArray As Variant, _
                       TheElement As Variant, _
                       ThePositions As Variant, _
                       Optional ParameterCheckQ As Boolean = True) As Variant
    Dim var As Variant
    Dim i As Long
    Dim ResultArray As Variant
    Dim ShiftedThePositions As Variant
    Dim NumDimensions  As Integer
    Dim ni As Long
    Dim CurrentIndex As Long
    
    ' Set default return value
    Let Insert = Null
    
    Let NumDimensions = NumberOfDimensions(AnArray)
    
    ' Process case of insertion in a 1D array
    If NumDimensions = 1 Then
        ' Process the case when ThePositions is a non-zero whole number
        If WholeNumberQ(ThePositions) Then
            ' Convert negative indices to positive
            Let ni = IIf(ThePositions < 0, Length(AnArray) + ni + 1, ThePositions)
            
            ' Handle cases of prepend and appending to the array
            If ThePositions = -1 And EmptyArrayQ(AnArray) Then
                Let Insert = Array(TheElement)
            ElseIf ni = 1 Then
                Let Insert = Prepend(AnArray, TheElement, ParameterCheckQ)
            ElseIf ni = Length(AnArray) + 1 Then
                Let Insert = Append(AnArray, TheElement, ParameterCheckQ)
            Else
                ' Compute the array to return for the case of insertion in the middle of the
                ' array
                Let ResultArray = ConcatenateArrays(Part(AnArray, Span(1, ni - 1)), _
                                                    Array(TheElement), _
                                                    ParameterCheckQ)
                Let ResultArray = ConcatenateArrays(ResultArray, _
                                                    Part(AnArray, Span(ni, -1)), _
                                                    ParameterCheckQ)
                
                Let Insert = ResultArray
            End If
        ' Process the case when ThePositions is not a non-zero whole integer
        ' This is the case when we want to insert the same element in multiple
        ' positions.
        ElseIf EmptyArrayQ(ThePositions) Then
            Let Insert = AnArray
        Else
            ' Perform recursion on the first index
            Let ResultArray = Insert(AnArray, TheElement, First(First(ThePositions)))
            
            ' Exit with Null if ResultArray came back as Null
            If NullQ(ResultArray) Then Exit Function
            
            ' Exit with ResultArray if Rest(ThePositions) is empty
            Let ShiftedThePositions = Rest(ThePositions)
            If EmptyArrayQ(ShiftedThePositions) Then
                Let Insert = ResultArray
                Exit Function
            End If
            
            ' Increment each of the remaining indices since they will be off by one
            ' after inserting the element at the first index
            Let i = 1
            Let ni = First(First(ThePositions))
            Let ni = IIf(ni < 0, Length(AnArray) + ni + 1, ni)
            For Each var In ShiftedThePositions
                Let CurrentIndex = First(var)
                Let CurrentIndex = IIf(CurrentIndex < 0, _
                                       Length(ResultArray) + CurrentIndex + 1, _
                                       CurrentIndex)
            
                If CurrentIndex > ni Then
                    Let ShiftedThePositions(LBound(ShiftedThePositions) + i - 1) = _
                        Array(CurrentIndex + 1)
                Else
                    Let ShiftedThePositions(LBound(ShiftedThePositions) + i - 1) = _
                        Array(CurrentIndex)
                End If
            
                Let i = i + 1
            Next
            
            ' Recurse of the rest of the indices
            Let Insert = Insert(ResultArray, TheElement, ShiftedThePositions)
        End If
    ' Process case of insertion in a 2D array
    ElseIf NumDimensions = 2 Then
        ' Process the case when ThePositions is a non-zero whole number
        If WholeNumberQ(ThePositions) Then
            ' Convert negative indices to positive
            Let ni = IIf(ThePositions < 0, Length(AnArray) + ni + 1, ThePositions)
            
            ' Handle cases of prepend and appending to the array
            If ni = 1 Then
                Let Insert = Prepend(AnArray, TheElement, ParameterCheckQ)
            ElseIf ni = Length(AnArray) + 1 Then
                Let Insert = Append(AnArray, TheElement, ParameterCheckQ)
            Else
                ' Compute the array to return for the case of insertion in the middle of the
                ' array
                Let ResultArray = StackArrays(Part(AnArray, Span(1, ni - 1)), _
                                              TheElement, _
                                              ParameterCheckQ)
                Let ResultArray = StackArrays(ResultArray, _
                                              Part(AnArray, Span(ni, -1)), _
                                              ParameterCheckQ)
                
                Let Insert = ResultArray
            End If
        ' Process the case when ThePositions is not a non-zero whole integer
        ' This is the case when we want to insert the same element in multiple
        ' positions.
        ElseIf EmptyArrayQ(ThePositions) Then
            Let Insert = AnArray
        Else
            ' Perform recursion on the first index
            Let ResultArray = Insert(AnArray, TheElement, First(First(ThePositions)))
            
            ' Exit with Null if ResultArray came back as Null
            If NullQ(ResultArray) Then Exit Function
            
            ' Exit with ResultArray if Rest(ThePositions) is empty
            Let ShiftedThePositions = Rest(ThePositions)
            
            If EmptyArrayQ(ShiftedThePositions) Then
                Let Insert = ResultArray
                Exit Function
            End If
            
            ' Increment each of the remaining indices since they will be off by one
            ' after inserting the element at the first index
            Let i = 1
            Let ni = First(First(ThePositions))
            Let ni = IIf(ni < 0, Length(AnArray) + ni + 1, ni)
            For Each var In ShiftedThePositions
                Let CurrentIndex = First(var)
                Let CurrentIndex = IIf(CurrentIndex < 0, _
                                       Length(ResultArray) + CurrentIndex + 1, _
                                       CurrentIndex)
                
                If CurrentIndex > ni Then
                    Let ShiftedThePositions(LBound(ShiftedThePositions) + i - 1) = _
                        Array(CurrentIndex + 1)
                Else
                    Let ShiftedThePositions(LBound(ShiftedThePositions) + i - 1) = _
                        Array(CurrentIndex)
                End If
            
                Let i = i + 1
            Next
            
            ' Recurse of the rest of the indices
            Let Insert = Insert(ResultArray, TheElement, ShiftedThePositions)
        End If
    End If
End Function

' DESCRIPTION
' Returns the array resulting from dropping the element at the given position.  It may be
' invoked with any of the following forms:
'
' a. Drop(AnArray, n) - With n>0, returns the array resulting from dropping the first n elements of AnArray.
'    n may range from 1 to Length(AnArray).  If AnArray is a 2D array satisfying
'    Predicates.AtomicTableQ, the rows are interpreted as the elements.
' b. Drop(AnArray, -n) - With n>0, returns the array resulting from dropping the last n elements of AnArray.
'    n may range from 1 to Length(AnArray).  If AnArray is a 2D array satisfying
'    Predicates.AtomicTableQ, the rows are interpreted as the elements.
' c. Drop(AnArray, Array(n)) - With n<>0, returns the array resulting from dropping the elements with normalized
'    index of n.
' d. Drop(AnArray, Array(n1,n2)) - With n1 and n2<>0 returns the array resulting from dropping
'    the elements between the normalized elements indices n1 and n2.  The normalized value of
'    n1 must be smaller than or equal to that of n2.
' e. Drop(AnArray, Array(n1,n2,step)) - With n1 and n2<>0 returns the array resulting from
'    dropping the elements between the normalized elements indices n1 and n2 with the given
'    step.  The normalized value of n1 must be smaller than or equal to that of n2.  step must
'    be a whole number equal to or larger of 1.
'
' PARAMETERS
' 1. AnArray - A dimensioned array or dimensions 1 or 2
' 2. ThePositions - One of the forms detailed above
'
' RETURNED VALUE
' The requested slice or element of the array.
Public Function Drop(AnArray As Variant, ThePositions As Variant) As Variant
    Dim var As Variant
    Dim i As Long
    Dim j As Long
    Dim ResultArray As Variant
    Dim n1 As Long
    Dim n2 As Long
    Dim TheStep As Long
    Dim IndicesToDrop As Variant
    
    ' Set default return value
    Let Drop = Null
    
    ' Process the case when ThePositions is an empty array
    If EmptyArrayQ(ThePositions) Then
        Let Drop = AnArray
        Exit Function
    End If
    
    ' Process the case when ThePositions is a non-zero whole number
    If WholeNumberQ(ThePositions) Then
        ' Exit with error code if ThePositions if not either
        ' between 1 and Length(AnArray) or -1 and -Length(AnArray)
        If ThePositions = 0 Or ThePositions > Length(AnArray) Or _
            ThePositions < -Length(AnArray) Then Exit Function
        
        ' Case ThePositions is a positive whole number
        If ThePositions = Length(AnArray) Then
            Let Drop = EmptyArray()
        ElseIf ThePositions > 0 Then
            Let Drop = Take(AnArray, Array(ThePositions + 1, Length(AnArray)))
        
        ' Case ThePositions is a negative whole number
        ElseIf ThePositions = -Length(AnArray) Then
            Let Drop = EmptyArray()
        Else
            Let Drop = Take(AnArray, Length(AnArray) + ThePositions)
        End If
    ElseIf WholeNumberArrayQ(ThePositions) Then
        Select Case Length(ThePositions)
            Case 1
                Let n1 = First(ThePositions)
                
                ' Exit with error code if ThePositions if not either
                ' between 1 and Length(AnArray) or -1 and -Length(AnArray)
                If n1 = 0 Or n1 > Length(AnArray) Or n1 < -Length(AnArray) Then
                    Exit Function
                End If
                
                If n1 = 1 Then
                    Let Drop = Rest(AnArray)
                ElseIf n1 = Length(AnArray) Then
                    Let Drop = Most(AnArray)
                ElseIf n1 = -1 Then
                    Let Drop = Most(AnArray)
                ElseIf n1 = -Length(AnArray) Then
                    Let Drop = Rest(AnArray)
                ElseIf n1 > 0 Then
                    Let Drop = JoinArrays(Take(AnArray, n1 - 1), _
                                          Take(AnArray, Array(n1 + 1, Length(AnArray))))
                Else
                    Let Drop = JoinArrays(Take(AnArray, Length(AnArray) + n1), _
                                          Take(AnArray, _
                                               Array(Length(AnArray) + n1 + 2, Length(AnArray))))
                End If
            Case 2
                Let n1 = First(ThePositions)
                Let n2 = Last(ThePositions)
                
                ' Exit with error code if ThePositions if not either
                ' between 1 and Length(AnArray) or -1 and -Length(AnArray)
                If n1 = 0 Or n1 > Length(AnArray) Or n1 < -Length(AnArray) Then
                    Exit Function
                End If
                
                If n2 = 0 Or n2 > Length(AnArray) Or n2 < -Length(AnArray) Then
                    Exit Function
                End If
                
                If n1 = 1 And n2 = Length(AnArray) Then
                    Let Drop = EmptyArray()
                ElseIf n1 > 1 And n2 = Length(AnArray) Then
                    Let Drop = Take(AnArray, n1 - 1)
                ElseIf n1 = 1 And n2 > 0 And n2 < Length(AnArray) Then
                    Let Drop = Take(AnArray, Array(n2 + 1, Length(AnArray)))
                ElseIf n1 > 1 And n2 > 0 And n2 < Length(AnArray) Then
                    Let Drop = JoinArrays(Take(AnArray, n1 - 1), _
                                          Take(AnArray, Array(n2 + 1, Length(AnArray))))
                ElseIf n1 > 0 And n1 = n2 Then
                    Let Drop = Drop(AnArray, Array(n1))
                ElseIf n1 = -Length(AnArray) And n2 = -1 Then
                    Let Drop = EmptyArray()
                ElseIf n1 = -Length(AnArray) And n2 < -1 And n1 < n2 Then
                    Let Drop = Take(AnArray, n2 + 1)
                ElseIf n1 < 0 And n1 > -Length(AnArray) And n2 < -1 And n1 < n2 Then
                    Let Drop = JoinArrays(Take(AnArray, Length(AnArray) + n1), _
                                          Take(AnArray, n2 + 1))
                ElseIf n1 < 0 And n1 = n2 Then
                    Let Drop = Drop(AnArray, Array(n1))
                End If
            Case 3
                ' Extract the index step
                Let TheStep = Last(ThePositions)
                
                If TheStep <= 0 Then
                    Let Drop = Null
                Else
                    ' Extract the start and stop indices to drop
                    Let var = NormalizeIndex(AnArray, First(ThePositions))
                    If IsNull(var) Then
                        Exit Function
                    Else
                        Let n1 = var
                    End If
                    
                    Let var = NormalizeIndex(AnArray, First(Rest(ThePositions)))
                    If IsNull(var) Then
                        Exit Function
                    Else
                        Let n2 = var
                    End If
                
                    ' Compute the set of indices to drop. Reverse them to
                    ' start dropping them from the end of the list.
                    Let IndicesToDrop = Reverse(NumericalSequence(n1, n2, TheStep, True))
                    
                    ' Recurse of each of the indices top drop all of the requested elements
                    Let ResultArray = AnArray
                    For Each var In IndicesToDrop
                        Let ResultArray = Drop(ResultArray, Array(var))
                    Next
                    
                    Let Drop = ResultArray
                End If
            Case Else
                Let Drop = Null
        End Select
    Else
        ' If the code gets here, ThePositions has an invalid form.
        Let Drop = Null
    End If
End Function

' DESCRIPTION
' Returns a 2D array whose rows are the elements of the given array.  This function returns Null
' if the elements of the given array don't all have the same length.  For instance,
' Array(Array(1,2,3), Array(4,5,6)) is transformed into [{1,2,3; 4,5,6}].  The function may be
' called with the optional parameter PackAsColumnsQ set to True to send
' Array(Array(1,2,3), Array(4,5,6)) is transformed into [{1, 4; 2, 5; 3, 6}]
'
' PARAMETERS
' 1. TheRowsAs1DArrays - A 1D array of equal-length 1D arrays
' 2. PackAsColumnsQ (optional) - When explicitly passed as True, returns the 2D transpose
' 3. ParameterCheckQ (optional) - When explicitly passed as True, performs NO parameter
'    consistency checks.
'
' RETURNED VALUE
' Returns the 2D arrays returning from setting the given 1D arrays as the rows of a 2D array
Public Function Pack2DArray(TheRowsAs1DArrays As Variant, _
                            Optional PackAsColumnsQ As Boolean = False, _
                            Optional ParameterCheckQ As Boolean = True) As Variant
    Dim var As Variant
    Dim r As Long
    Dim c As Long
    Dim RowOffset As Long
    Dim ColOffset As Long
    Dim ReturnArray As Variant
    Dim ElementLength As Long
    
    ' Set default return value
    Let Pack2DArray = Null

    ' Check parameter consitency only if ParameterCheckQ True
    If ParameterCheckQ Then
        ' Exit with Null if A2DArray is undimensioned
        If Not DimensionedQ(TheRowsAs1DArrays) Then Exit Function
        
        ' Exit with the empty array if TheRowsAs1DArrays is the empty array
        If EmptyArrayQ(TheRowsAs1DArrays) Then
            Let Pack2DArray = EmptyArray()
            Exit Function
        End If
        
        ' Exit if the argument is not the expected type
        If NumberOfDimensions(TheRowsAs1DArrays) <> 1 Then Exit Function
        
        ' Exit if any of the elements in not an atomic array or
        ' if all the array elements do not the same length
        If Not IsNull(First(TheRowsAs1DArrays)) Then
            Let ElementLength = Length(First(TheRowsAs1DArrays))
        Else
            Exit Function
        End If
    
        For Each var In TheRowsAs1DArrays
            If Length(var) <> ElementLength Then
                Let Pack2DArray = Null
                Exit Function
            End If
        Next
    End If
    
    ' Compute row and column offsets for ReturnArray
    Let RowOffset = IIf(LBound(TheRowsAs1DArrays) = 0, 1, 0)
    Let ColOffset = IIf(LBound(First(TheRowsAs1DArrays)) = 0, 1, 0)

    If PackAsColumnsQ Then
        ' Pre-allocate a 2D array
        ReDim ReturnArray(1 To Length(First(TheRowsAs1DArrays)), 1 To Length(TheRowsAs1DArrays))
    
        For r = LBound(First(TheRowsAs1DArrays)) To UBound(First(TheRowsAs1DArrays))
            For c = LBound(TheRowsAs1DArrays) To UBound(TheRowsAs1DArrays)
                Let ReturnArray(RowOffset + r, ColOffset + c) = TheRowsAs1DArrays(c)(r)
            Next c
        Next r
    Else
        ' Pre-allocate a 2D array
        ReDim ReturnArray(1 To Length(TheRowsAs1DArrays), 1 To Length(First(TheRowsAs1DArrays)))
        
        For r = LBound(TheRowsAs1DArrays) To UBound(TheRowsAs1DArrays)
            For c = LBound(First(TheRowsAs1DArrays)) To UBound(First(TheRowsAs1DArrays))
                Let ReturnArray(RowOffset + r, ColOffset + c) = TheRowsAs1DArrays(r)(c)
            Next c
        Next r
    End If
    
    Let Pack2DArray = ReturnArray
End Function

' Alias for Pack2DArray
Public Function Pack(TheRowsAs1DArrays As Variant, _
                     Optional PackAsColumnsQ As Boolean = False, _
                     Optional ParameterCheckQ As Boolean = True) As Variant
    Let Pack = Pack2DArray(TheRowsAs1DArrays, PackAsColumnsQ, ParameterCheckQ)
End Function

' DESCRIPTION
' Returns a 1D array of equal length arrays, where each 1D element of the returned array is one
' of the rows of the given 2D array.  This function returns Null if the given array is
' undimensioned. It returns the empty array if the given array is empty.  The function may be
' called with the optional parameter UnPackAsColumnsQ set to True to send to return the columns
' of the given 2D array as the elements of the returned array.
'
' Examples:
' 1. [{1,2,3; 4,5,6}] maps to Array(Array(1,2,3), Array(4,5,6))
' 2. [{1, 4; 2, 5; 3, 6}] maps to Array(Array(1,2,3), Array(4,5,6))
'
' PARAMETERS
' 1. TheRowsAs1DArrays - A 1D array of equal-length 1D arrays
' 2. UnPackAsColumnsQ (optional) - When explicitly passed as True, returns the 1D array whose
'    elements are the columns of TheRowsAs1DArrays
'
' RETURNED VALUE
' Returns the 1D whose elements are the rows or columns (if explicitly requested) of the given
' 2D array
Public Function UnPack2DArray(A2DArray As Variant, _
                              Optional UnPackAsColumnsQ As Boolean = False) As Variant
    Dim r As Long
    Dim c As Long
    Dim rOffset As Long
    Dim cOffset As Long
    Dim ReturnArray As Variant
    Dim ColumnArray As Variant
    Dim RowArray As Variant
    
    ' Set default return value
    Let UnPack2DArray = Null
    
    ' Exit with Null if A2DArray is undimensioned
    If Not DimensionedQ(A2DArray) Then Exit Function
    
    ' Exit with the empty array if A2DArray is the empty array
    If EmptyArrayQ(A2DArray) Then
        Let UnPack2DArray = EmptyArray()
        Exit Function
    End If
    
    ' Exit if the argument is not the expected type
    If NumberOfDimensions(A2DArray) <> 2 Then Exit Function
    
    ' Unpack the array, returning the columns as the elements of the return array
    If UnPackAsColumnsQ Then
        ' Pre-allocate a 2D array filled with Empty
        ReDim ReturnArray(1 To NumberOfColumns(A2DArray))
        ReDim ColumnArray(1 To Length(A2DArray))
        
        Let rOffset = LBound(A2DArray, 1) - 1
        Let cOffset = LBound(A2DArray, 2) - 1
        
        For c = 1 To NumberOfColumns(A2DArray)
            For r = 1 To Length(A2DArray)
                Let ColumnArray(r) = A2DArray(r + rOffset, c + cOffset)
            Next
            
            Let ReturnArray(c) = ColumnArray
        Next c
    ' Unpack the array, returning the rows as the elements of the return array
    Else
        ' Pre-allocate a 2D array filled with Empty
        ReDim ReturnArray(1 To Length(A2DArray))
        ReDim RowArray(1 To NumberOfColumns(A2DArray))
        
        Let rOffset = LBound(A2DArray, 1) - 1
        Let cOffset = LBound(A2DArray, 2) - 1
    
        For r = 1 To Length(A2DArray)
            For c = 1 To NumberOfColumns(A2DArray)
                Let RowArray(c) = A2DArray(r + rOffset, c + cOffset)
            Next c
            
            Let ReturnArray(r) = RowArray
        Next r
    End If
    
    Let UnPack2DArray = ReturnArray
End Function

' DESCRIPTION
' Returns a constant array for the given value and with the given dimensions.  If the parameter
' n (e.g. number of columns) is omitted, the function returns a 1D array.  Otherwise, a 2D array
' is returned. The function returns Null if m is nonpositive.
'
' PARAMETERS
' 1. TheValue - Constant entry for the returned array
' 2. m - Array length for 1D and number of rows for 2D
' 3. n (optional) - Number of  columns when requesting a 2D array
'
' RETURNED VALUE
' A constant 1D or 2D array with the given number of rows and columns.
Public Function ConstantArray(TheValue As Variant, m As Long, Optional n As Long = 0) As Variant
    Dim ReturnMatrix() As Variant
    Dim i As Long
    Dim j As Long
    
    Let ConstantArray = Null

    If NegativeWholeNumberQ(m) Or m = 0 Then Exit Function
    If NegativeWholeNumberQ(n) Then Exit Function
    
    If n = 0 Then
        ReDim ReturnMatrix(1 To m)
        
        For i = 1 To m
            ReturnMatrix(i) = TheValue
        Next
    Else
        ReDim ReturnMatrix(1 To m, 1 To n)
        
        For i = 1 To m
            For j = 1 To n
                Let ReturnMatrix(i, j) = TheValue
            Next
        Next
    End If
    
    Let ConstantArray = ReturnMatrix
End Function

' DESCRIPTION
' Returns a square identity matrix with the given dimensions.  If the parameter
' n (e.g. number of columns) is passed, the function returns a rectangular 2D
' matrix with ones along the diagonal and zeroes elsewhere.  The function returns
' Null if m is nonpositive.
'
' PARAMETERS
' 1. TheValue - Constant entry for the returned array
' 2. m - Array length for 1D and number of rows for 2D
' 3. n (optional) - Number of  columns when requesting a 2D array
'
' RETURNED VALUE
' The identity matrix with the requested dimensions
Public Function IdentityMatrix(m As Long, Optional n As Long = 0) As Variant
    Dim ReturnMatrix() As Long
    Dim i As Long
    
    Let IdentityMatrix = Null

    If NegativeWholeNumberQ(m) Or m = 0 Then Exit Function
    If NegativeWholeNumberQ(n) Then Exit Function
    
    If n = 0 Then
        ReDim ReturnMatrix(1 To m, 1 To m)
        
        For i = 1 To m
            Let ReturnMatrix(i, i) = 1
        Next
    Else
        ReDim ReturnMatrix(1 To m, 1 To n)
        
        For i = 1 To Application.WorksheetFunction.Min(m, n)
            Let ReturnMatrix(i, i) = 1
        Next
    End If
    
    Let IdentityMatrix = ReturnMatrix
End Function

' DESCRIPTION
' Returns a single pseudorandom real number, an array, or a matrix of them per the
' user's specifications.
'
' PARAMETERS
' 1. aRangeOfNumbers (optional) - An Array of parameter.  The sequence of parameters
'    could be any of the following:
'
'    a. IsMissing=True - returns a pseudorandom number between 0 and 1
'    b. Xmax returns a pseudorandom number between 0 and Xmax
'    c. Array(Xmin, Xmax) - returns a pseudorandom number between Xmin and Xmax
'
' 2. TheDimensions (optional) - Indicates the number of elements to return.  aRangeOfNumbers
'    must be provided if TheDimensions is.  TheDimensions must be either a non-negative whole
'    number or an array of them. It may take one of the following two forms:
'
'    a. N - a positive whole number
'    b. (N, M) - An array of two positive whole numbers
'    c. 0 - or omitted to request a scalar
'
' RETURNED VALUE
' Either a pseudorandom real, or an array/matrix of them
Public Function RandomReal(Optional ByVal aRangeOfNumbers As Variant = 1, _
                           Optional ByVal TheDimensions As Variant = 0, _
                           Optional ParamCheckQ As Boolean = False, _
                           Optional ReturnIntegersQ As Boolean = False) As Variant
    Dim ParamsCopy As Variant
    Dim LB As Double
    Dim UB As Double
    Dim NumRows As Long
    Dim NumCols As Long
    Dim ReturnArray As Variant
    Dim i As Long
    Dim j As Long

    ' Set default value when encountering errors
    Let RandomReal = Null
    
    ' Exit with Null if aRangeOfNumbers fails param consistency check
    If ParamCheckQ Then
        If Not (NumberQ(aRangeOfNumbers) Or _
                (NumberArrayQ(aRangeOfNumbers) And Length(aRangeOfNumbers) = 2)) Then
            Exit Function
        End If

        ' Exit with Null if TheDimensions is not either a positive, whole number or a
        ' 2-elt array of them
        If Not (PositiveWholeNumberQ(TheDimensions) Or _
                (PositiveWholeNumberArrayQ(TheDimensions) And Length(TheDimensions) = 2)) Then
            Exit Function
        End If
        
        ' Exit with Null if Length(aRangeOfNumbers)<>0 or 2
        If Length(aRangeOfNumbers) <> 0 Or Length(aRangeOfNumbers) <> 2 Then Exit Function
        
        ' Exit with Null if aRangeOfNumbers = (LB,UB) and LB>=UB
        If IsArray(aRangeOfNumbers) Then
            If First(aRangeOfNumbers) >= Last(aRangeOfNumbers) Then
                Exit Function
            End If
        End If
    End If
    
    ' Case when TheDimensions = Array(LB, UB)
    If IsArray(TheDimensions) Then
        Let NumRows = First(TheDimensions)
        Let NumCols = Last(TheDimensions)
    
        ReDim ReturnArray(1 To NumRows, 1 To NumCols)

        ' aRange = Array(LB, UB)
        If IsArray(aRangeOfNumbers) Then
            Let LB = First(aRangeOfNumbers)
            Let UB = Last(aRangeOfNumbers)
            
            For i = 1 To NumRows
                For j = 1 To NumCols
                    Let ReturnArray(i, j) = Rnd() * (UB - LB) + LB
                    If ReturnIntegersQ Then Let ReturnArray(i, j) = Round(CDbl(ReturnArray(i, j)))
                Next
            Next
        ' aRangeOfNumbers = a real number
        Else
            For i = 1 To NumRows
                For j = 1 To NumCols
                    Let ReturnArray(i, j) = aRangeOfNumbers * Rnd()
                    If ReturnIntegersQ Then Let ReturnArray(i, j) = Round(CDbl(ReturnArray(i, j)))
                Next
            Next
        End If
    Else
        If TheDimensions = 0 And IsNumeric(aRangeOfNumbers) Then
            Let ReturnArray = aRangeOfNumbers * Rnd()
        ElseIf TheDimensions = 0 And IsArray(aRangeOfNumbers) Then
            Let LB = First(aRangeOfNumbers)
            Let UB = Last(aRangeOfNumbers)
            
            Let ReturnArray = Rnd() * (UB - LB) + LB
                If ReturnIntegersQ Then Let ReturnArray = Round((ReturnArray))
        ElseIf IsArray(aRangeOfNumbers) Then
            ReDim ReturnArray(1 To TheDimensions)
            
            Let LB = First(aRangeOfNumbers)
            Let UB = Last(aRangeOfNumbers)

            For i = 1 To TheDimensions
                Let ReturnArray(i) = Rnd() * (UB - LB) + LB
                If ReturnIntegersQ Then Let ReturnArray(i) = Round(CDbl(ReturnArray(i)))
            Next
        Else
            ReDim ReturnArray(1 To TheDimensions)
            
            For i = 1 To TheDimensions
                Let ReturnArray(i) = aRangeOfNumbers * Rnd()
                If ReturnIntegersQ Then Let ReturnArray(i) = Round(CDbl(ReturnArray(i)))
            Next
        End If

    End If
    
    Let RandomReal = ReturnArray
End Function

' DESCRIPTION
' Returns a random permutation of a 1D array
'
' PARAMETERS
' 1. A1DArray - a 1D array
' 2. Optional ReturnAsColumnVectorQ = False = When invoked with value True, the result is returned
'    as a column vector
'
' RETURNED VALUE
' A random permutation of a 1D array or Null if the input is not valid
Public Function Randomize1DArray(A1DArray As Variant, Optional ReturnAsColumnVectorQ As Boolean = False) As Variant
    Dim DumpArray As Variant
    Dim i As Long
    Dim n As Long
    
    ' Set default return value denoting an error in the inputs
    Let Randomize1DArray = Null
    
    If NumberOfDimensions(A1DArray) <> 1 Then Exit Function
    
    Let n = Length(A1DArray)
    Let DumpArray = ConstantArray(Empty, n, 2)
    For i = 1 To n
        Let DumpArray(i, 1) = A1DArray(i)
        Let DumpArray(i, 2) = Rnd()
    Next
    
    Let DumpArray = Sort2DArray(DumpArray, Array(2), Array(xlAscending), WithHeaders:=xlNo)
    
    If ReturnAsColumnVectorQ Then
        Let Randomize1DArray = Part(DumpArray, Span(1, -1), Span(1, 1))
    Else
        Let Randomize1DArray = Part(DumpArray, Span(1, -1), 1)
    End If
End Function

' DESCRIPTION
' This function returns the array resulting from concatenating the given list of arrays and
' along the optional dimensionality index. If the dimensionality index is not provided, it
' defaults to the first dimension. This function allows for dimensionality indices 1 and 2.
' All elements to join must be arrays.  The function returns Null if the dimensionality index
' is larger than the number of dimensions in any of the arrays beingjoined. 2D arrays are
' interpreted as 1D arrays of 1D arrays.
'
' PARAMETERS
' 1. Params - a variant ParamArray, each of which except possibly the last one must be arrays.
'    If the
'    last element of Params is an integer, then the function joins the arrays along the
'    dimensionality index specified by such integer.  At the moment, the dimensionality can
'    only be 1 and 2.
'
' RETURNED VALUE
' An array joining the arrays passed as parameters
Public Function JoinArrays(ParamArray Params()) As Variant
    Dim ParamsCopy As Variant
    Dim var As Variant
    Dim var2 As Variant
    Dim DimensionalityIndex As Integer
    Dim ReturnArray As Variant
    Dim ReturnArrayLength As Long
    Dim MaxNumRows As Long
    Dim i As Long
    Dim j As Long
    Dim k() As Long
    Dim OneParam2DFlag As Boolean
    Dim SameLengthFlag As Boolean
    
    ' Set default value to return when encountering errors
    Let JoinArrays = Null
    
    ' Make copy of ParamArray
    Let ParamsCopy = CopyParamArray(Params)
    
    ' Exit with Null if no parameters are passed
    If Not DimensionedQ(ParamsCopy) Then Exit Function
    
    ' Return the empty array if called with no arrays to join
    If EmptyArrayQ(ParamsCopy) Then
        Let JoinArrays = EmptyArray()
        Exit Function
    End If
    
    ' Exit with Null if any of the parameters is not an array (with the exception of the last
    ' one, which could be 1 or 2.  If the last parameter is 1 or 2 and the only parameter, exit
    ' with Null.  If any of the parameters other than the last one is not an array, exit with
    ' Null.  If the last parameter is 2 and any of the other parameters is 1, exith with Null.
    If PositiveWholeNumberQ(Last(ParamsCopy)) Then
        Let DimensionalityIndex = Last(ParamsCopy)
    
        If DimensionalityIndex > 2 Then Exit Function
        If Length(ParamsCopy) < 2 Then Exit Function
        
        For i = LBound(ParamsCopy) To UBound(ParamsCopy) - 1
            If Not DimensionedQ(ParamsCopy(i)) Then Exit Function
            If NumberOfDimensions(ParamsCopy(i)) < DimensionalityIndex Then Exit Function
        Next
    Else
        For i = LBound(ParamsCopy) To UBound(ParamsCopy)
            If Not DimensionedQ(ParamsCopy(i)) Then Exit Function
        Next
    End If
    
    ' Process according to the dimensionality index
    If DimensionalityIndex = 0 Or DimensionalityIndex = 1 Then
        ' Turn on repackaging flag if all arrays have the same number of columns and
        ' dimensionality index is 1 and at least one is a 2D array
        Let OneParam2DFlag = False
        Let SameLengthFlag = True
        For i = LBound(ParamsCopy) To UBound(ParamsCopy)
            If NumberOfColumns(First(ParamsCopy)) <> _
               NumberOfColumns(ParamsCopy(i)) And IsArray(ParamsCopy(i)) Then
                Let SameLengthFlag = False
            End If
            
            If NumberOfDimensions(ParamsCopy(i)) = 2 Then Let OneParam2DFlag = True
        Next
    
        ' Compute array length required to hold joined arrays
        For i = LBound(ParamsCopy) To UBound(ParamsCopy)
            Let ReturnArrayLength = ReturnArrayLength + Length(ParamsCopy(i))
        Next
                
        ' Convert all 2D arrays to 1D arrays
        For i = LBound(ParamsCopy) To UBound(ParamsCopy)
            If IsArray(ParamsCopy(i)) And NumberOfDimensions(ParamsCopy(i)) = 2 Then
                Let ParamsCopy(i) = UnPack2DArray(ParamsCopy(i))
            End If
        Next
        
        ' Pre-allocate an array big enough to hold the joined arrays
        ReDim ReturnArray(1 To ReturnArrayLength)
        
        ' Create the joined array
        Let i = 1
        For Each var In ParamsCopy
            If IsArray(var) Then
                For j = LBound(var) To UBound(var)
                    Let ReturnArray(i) = Part(var, j + 1 - LBound(var))
                    Let i = i + 1
                Next
            End If
        Next
        
        ' Repack as 2D array if necessary
        If OneParam2DFlag And SameLengthFlag Then Let ReturnArray = Pack2DArray(ReturnArray)
    Else
        ' Compute the number of columns required for the return array
        For Each var In ParamsCopy
            Let ReturnArrayLength = ReturnArrayLength + NumberOfColumns(var)
            Let MaxNumRows = Application.Max(MaxNumRows, Length(var))
        Next
        
        ' Pre-allocate return array
        ReDim ReturnArray(1 To MaxNumRows, 1 To ReturnArrayLength)
        ReDim k(1 To MaxNumRows)
        
        For Each var In ParamsCopy
            If IsArray(var) Then
                For i = LBound(var) To UBound(var)
                    For j = LBound(var, 2) To UBound(var, 2)
                        Let k(i + 1 - LBound(var)) = k(i + 1 - LBound(var)) + 1
                        Let ReturnArray(i + 1 - LBound(var), k(i + 1 - LBound(var))) = var(i, j)
                    Next
                Next
            End If
        Next
    End If
    
    Let JoinArrays = ReturnArray
End Function

' DESCRIPTION
' This function returns the list of unique values or object references in the given array.
' This array could be nested as many levels deep as desired.  The function considers only the
' leaves of the tree represented by the array.  This function returns Null if AnArray is
' non-dimensioned.
'
' PARAMETERS
' 1. AnArray - A dimensioned array
'
' RETURNED VALUE
' An array with the unique values at the leaves of the tree represented by the array or the
' unchanged argument if it is atomic
Public Function UniqueSubset(AnArray As Variant) As Variant
    Dim var As Variant
    Dim UniqueDict As Dictionary
    
    ' Set default return value in case of errors
    Let UniqueSubset = Null
    
    ' Exit with Null if AnArray is not dimensioned
    If Not DimensionedQ(AnArray) Then Exit Function
    
    ' Exit, returning an empty array if AnArray is empty
    If EmptyArrayQ(AnArray) Then
        Let UniqueSubset = EmptyArray()
        Exit Function
    End If
    
    Set UniqueDict = New Dictionary
    For Each var In Flatten(AnArray)
        If Not UniqueDict.Exists(var) Then Call UniqueDict.Add(Key:=var, Item:=1)
    Next
    
    Let UniqueSubset = UniqueDict.Keys
End Function

' DESCRIPTION
' This function returns the set-theoretic union of the atomic values in the leaves of the trees
' represented by the two given arrays.
'
' PARAMETERS
' 1. Set1 - A dimensioned array
' 2. Set2 - A dimensioned array
'
' RETURNED VALUE
' The set theoretic union of the atomic values in the leaves of the trees represented by
' the arrays
Public Function UnionOfSets(Set1 As Variant, Set2 As Variant) As Variant
    Let UnionOfSets = UniqueSubset(JoinArrays(Flatten(Set1), Flatten(Set2)))
End Function

' DESCRIPTION
' This function returns the set-theoretic intersection of the atomic values in the leaves of
' the trees represented by the two given arrays.
'
' PARAMETERS
' 1. Set1 - A dimensioned array
' 2. Set1 -  A dimensioned array
'
' RETURNED VALUE
' The set theoretic intersection of the atomic values in the leaves of the trees represented
' by the arrays
Public Function IntersectionOfSets(Set1 As Variant, Set2 As Variant) As Variant
    Dim FirstDict As Dictionary
    Dim First1DSet As Variant
    Dim Second1DSet As Variant
    Dim IntersectionDict As Dictionary
    Dim i As Long
    
    ' Set default value to return if error encountered
    Let IntersectionOfSets = Null
    
    ' Exit if either set if not dimensioned
    If Not (DimensionedQ(Set1) And DimensionedQ(Set2)) Then Exit Function
    
    ' Exit returning an empty array if either set is empty
    If EmptyArrayQ(Set1) Or EmptyArrayQ(Set2) Then
        Let IntersectionOfSets = EmptyArray()
        Exit Function
    End If
    
    ' Instantiate dictionaries
    Set FirstDict = New Dictionary
    Set IntersectionDict = New Dictionary
    
    ' Convert each set to a 1D array
    Let First1DSet = Flatten(Set1)
    Let Second1DSet = Flatten(Set2)
    
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

' DESCRIPTION
' This function returns the set-theoretic complement of the atomic values in the leaves of
' first tree with respect to the leaves of the other tree.
'
' PARAMETERS
' 1. Set1 - A dimensioned array
' 2. Set1 -  A dimensioned array
'
' RETURNED VALUE
' The set complement of one set with respect to another
Public Function ComplementOfSets(a As Variant, b As Variant) As Variant
    Dim BDict As Dictionary
    Dim ComplementDict As Dictionary
    Dim var As Variant
    
    ' Set default value to return if error encountered
    Let ComplementOfSets = Null
    
    ' Exit with Null if either set is undimensioned
    If Not (DimensionedQ(a) And DimensionedQ(b)) Then Exit Function
    
    ' If a is an empty array, exit returning an empty array
    If EmptyArrayQ(a) Then
        Let ComplementOfSets = EmptyArray()
        Exit Function
    End If
    
    If EmptyArrayQ(b) Then
        Let ComplementOfSets = a
        Exit Function
    End If
    
    ' Instantiate dictionaries
    Set BDict = New Dictionary
    Set ComplementDict = New Dictionary
    
    ' Initialize BDict to get unique subset of ADict
    For Each var In UniqueSubset(b)
        If Not BDict.Exists(Key:=var) Then
            Call BDict.Add(Key:=var, Item:=var)
        End If
    Next
    
    ' Populate ComplementDict
    For Each var In UniqueSubset(a)
        If Not BDict.Exists(Key:=var) And Not ComplementDict.Exists(Key:=var) Then
            Call ComplementDict.Add(Key:=var, Item:=var)
        End If
    Next
    
    ' Return complement as 1D array
    If ComplementDict.Count = 0 Then
        Let ComplementOfSets = EmptyArray()
    Else
        Let ComplementOfSets = Flatten(ComplementDict.Keys)
    End If
End Function

' DESCRIPTION
' This function returns the transpose the given 1D or rectangular table. The elements of the
' array or table need not be atomic.  It currently processes 1D, 2D arrays, and 1D array of
' arrays.
'
' PARAMETERS
' 1. aMatrix - The array to transpose
' 2. UseBuiltInQ -  Uses Application.Transpose when explicitly set to True
' 3. ParamConsistencyChecksQ (optiomal) - If explicitly passed as False, this function performs
'    no parameter consistency checks
'
' RETURNED VALUE
' The set complement of one set with respect to another
Public Function TransposeMatrix(AMatrix As Variant, _
                                Optional UseBuiltInQ As Boolean = False, _
                                Optional ParameterCheckQ As Boolean = True) As Variant
    Dim r As Long
    Dim c As Long
    Dim TheResult() As Variant

    Let TransposeMatrix = Null

    If ParameterCheckQ Then
        If Not DimensionedQ(AMatrix) Then Exit Function
        
        If NumberOfDimensions(AMatrix) = 0 Then
            Let TransposeMatrix = AMatrix
            Exit Function
        End If
    End If
    
    If EmptyArrayQ(AMatrix) Then
        Let TransposeMatrix = EmptyArray()
        Exit Function
    End If
    
    If UseBuiltInQ Then
        Let TransposeMatrix = Application.Transpose(AMatrix)
        Exit Function
    End If
    
    If NumberOfDimensions(AMatrix) = 1 Then
        ReDim TheResult(LBound(AMatrix) To UBound(AMatrix), 1)
    
        For c = LBound(AMatrix) To UBound(AMatrix)
            Let TheResult(c, 1) = AMatrix(c)
        Next c
        
        Let TransposeMatrix = TheResult
        
        Exit Function
    End If
    
    ReDim TheResult(LBound(AMatrix, 2) To _
          UBound(AMatrix, 2), LBound(AMatrix, 1) To UBound(AMatrix, 1))

    For r = LBound(AMatrix, 1) To UBound(AMatrix, 1)
        For c = LBound(AMatrix, 2) To UBound(AMatrix, 2)
            Let TheResult(c, r) = AMatrix(r, c)
        Next c
    Next r
    
    Let TransposeMatrix = TheResult
End Function

' DESCRIPTION
' Alias for Arrays.MatrixTraspose().
' This function returns the transpose the given 1D or rectangular
' table. The elements of the array or table need not be atomic.
' It currently processes 1D, 2D arrays, and 1D array of arrays.
'
' PARAMETERS
' 1. aMatrix - The array to transpose
' 2. UseBuiltInQ -  Uses Application.Transpose when explicitly set
'    to True
' 3. ParamConsistencyChecksQ (optiomal) - If explicitly passed as
'    False, this function performs no parameter consistency checks
'
' RETURNED VALUE
' The set complement of one set with respect to another
Public Function Transpose(AMatrix As Variant, _
                          Optional UseBuiltInQ As Boolean = False, _
                          Optional ParameterCheckQ As Boolean = True) As Variant
    Let Transpose = TransposeMatrix(AMatrix, UseBuiltInQ, ParameterCheckQ)
End Function

' DESCRIPTION
' This function adds numbers, vectors or matrices elementwise, as long as it makes sense.
' Examples:

'    1. scalars - The are added normally and returned as a scalar
'    2. scalars and vectors/matrices - The scalar is added to each element of the other
'       argument
'    3. vectors - Two row or two column vectors of the same dimensions are added elementwise
'    4. a row vector and matrix - are added elementwise for each row of the matrix. In other
'       words, the row vector is added once to each row of the matrix. The result is returned
'       as matrix.
'    5. a column vector and matrix - are added elementwise for each column of the matrix. In
'       other words, the column vector is added once to each column of the matrix. The result
'       is returned as matrix.
'    6. two matrices - Addition is performed elementwise and returned as a matrix
'
' PARAMETERS
' 1. Matrix1 - a scalar, vector, or matrix
' 2. Matrix2 -  a scalar, vector, or matrix
' 3. ParamConsistencyChecksQ (optiomal) - If explicitly passed as False, this function performs
'    no parameter consistency checks
'
' RETURNED VALUE
' The result of performing elementwise addition on the last two arguments
Public Function Add(Matrix1 As Variant, _
                    Matrix2 As Variant, _
                    Optional ParamConsistencyChecksQ = True) As Variant
    Dim Arg1 As Variant
    Dim Arg2 As Variant
    Dim nd1 As Integer
    Dim nd2 As Integer
    Dim nr1 As Long
    Dim nr2 As Long
    Dim nc1 As Long
    Dim nc2 As Long
    Dim l1 As Long
    Dim l2 As Long
    Dim r As Long
    Dim c As Long
    Dim Result() As Variant

    Let Add = Null

    If ParamConsistencyChecksQ Then
        If Not ElementwiseArithmeticParameterConsistentQ(Matrix1, Matrix2) Then
            Let Add = Null
            Exit Function
        End If
            
        If EmptyArrayQ(Matrix1) And EmptyArrayQ(Matrix2) Then
            Let Add = EmptyArray()
            Exit Function
        End If
    End If
    
    Let nd1 = NumberOfDimensions(Matrix1)
    Let nd2 = NumberOfDimensions(Matrix2)
    
    Let nr1 = NumberOfRows(Matrix1)
    Let nr2 = NumberOfRows(Matrix2)
    Let nc1 = NumberOfColumns(Matrix1)
    Let nc2 = NumberOfColumns(Matrix2)
    Let l1 = Length(Matrix1)
    Let l2 = Length(Matrix2)
    
    If nd1 = 2 And nd2 = 2 Then
        If nr1 = nr2 And nc1 = nc2 Then
            ReDim Result(1 To nr1, 1 To nc1)
            
            For r = 1 To nr1
                For c = 1 To nc1
                    Let Result(r, c) = Part(Matrix1, r, c) + Part(Matrix2, r, c)
                Next
            Next
        ElseIf nr1 = 1 Then
            ReDim Result(1 To nr2, 1 To nc2)
   
            For r = 1 To nr2
                For c = 1 To l1
                    Let Result(r, c) = Part(Matrix1, c) + Part(Matrix2, r, c)
                Next
            Next
        ElseIf nr2 = 1 Then
            ReDim Result(1 To nr1, 1 To nc1)

            For r = 1 To nr1
                For c = 1 To l2
                    Let Result(r, c) = Part(Matrix1, r, c) + Part(Matrix2, c)
                Next
            Next
        ElseIf nc1 = 1 Then
            ReDim Result(1 To nr2, 1 To nc2)

            For r = 1 To nr1
                For c = 1 To nc2
                    Let Result(r, c) = Part(Matrix1, r, 1) + Part(Matrix2, r, c)
                Next
            Next
        ElseIf nc2 = 1 Then
            ReDim Result(1 To nr1, 1 To nc1)

            For r = 1 To nr1
                For c = 1 To nc1
                    Let Result(r, c) = Part(Matrix1, r, c) + Part(Matrix2, r, 1)
                Next
            Next
        End If
    ElseIf nd1 = 1 And nd2 = 2 Then
        ReDim Result(1 To nr2, 1 To nc2)
        
        For r = 1 To nr2
            For c = 1 To nc2
                Let Result(r, c) = Part(Matrix1, c) + Part(Matrix2, r, c)
            Next
        Next
    ElseIf nd1 = 2 And nd2 = 1 Then
        ReDim Result(1 To nr1, 1 To nc1)
        
        For r = 1 To nr1
            For c = 1 To nc1
                Let Result(r, c) = Part(Matrix1, r, c) + Part(Matrix2, c)
            Next
        Next
    ElseIf nd1 = 2 And nd2 = 0 Then
        ReDim Result(1 To nr1, 1 To nc1)
    
        For r = 1 To nr1
            For c = 1 To nc1
                Let Result(r, c) = Part(Matrix1, r, c) + Matrix2
            Next
        Next
    ElseIf nd1 = 0 And nd2 = 2 Then
        ReDim Result(1 To nr2, 1 To nc2)
    
        For r = 1 To nr2
            For c = 1 To nc2
                Let Result(r, c) = Matrix1 + Part(Matrix2, r, c)
            Next
        Next
    ElseIf nd1 = 1 And nd2 = 1 Then
        ReDim Result(1 To l1)
        
        For c = 1 To l1
            Let Result(c) = Part(Matrix1, c) + Part(Matrix2, c)
        Next
    ElseIf nd1 = 0 And nd2 = 1 Then
        ReDim Result(1 To l2)
        
        For c = 1 To l2
            Let Result(c) = Matrix1 + Part(Matrix2, c)
        Next
    ElseIf nd1 = 1 And nd2 = 0 Then
        ReDim Result(1 To l1)
        
        For c = 1 To l1
            Let Result(c) = Part(Matrix1, c) + Matrix2
        Next
    Else
        Let Add = Matrix1 + Matrix2
        Exit Function
    End If
    
    Let Add = Result
End Function

' DESCRIPTION
' This function adds numbers, vectors or matrices elementwise, as long as it makes sense.
' Examples:

'    1. scalars - The are added normally and returned as a scalar
'    2. scalars and vectors/matrices - The scalar is added to each element of the other
'       argument
'    3. vectors - Two row or two column vectors of the same dimensions are added elementwise
'    4. a row vector and matrix - are added elementwise for each row of the matrix. In other
'       words, the row vector is added once to each row of the matrix. The result is returned
'       as matrix.
'    5. a column vector and matrix - are added elementwise for each column of the matrix. In
'       other words, the column vector is added once to each column of the matrix. The result
'       is returned as matrix.
'    6. two matrices - Addition is performed elementwise and returned as a matrix
'
' PARAMETERS
' 1. Matrix1 - a scalar, vector, or matrix
' 2. Matrix2 -  a scalar, vector, or matrix
' 3. ParamConsistencyChecksQ (optiomal) - If explicitly passed as False, this function performs
'    no parameter consistency checks
'
' RETURNED VALUE
' The result of performing elementwise addition on the last two arguments
Public Function Subtract(Matrix1 As Variant, _
                         Matrix2 As Variant, _
                         Optional ParamConsistencyChecksQ = True) As Variant
    Dim Arg1 As Variant
    Dim Arg2 As Variant
    Dim nd1 As Integer
    Dim nd2 As Integer
    Dim nr1 As Long
    Dim nr2 As Long
    Dim nc1 As Long
    Dim nc2 As Long
    Dim l1 As Long
    Dim l2 As Long
    Dim r As Long
    Dim c As Long
    Dim Result() As Variant

    Let Subtract = Null

    If ParamConsistencyChecksQ Then
        If Not ElementwiseArithmeticParameterConsistentQ(Matrix1, Matrix2) Then
            Let Subtract = Null
            Exit Function
        End If
            
        If EmptyArrayQ(Matrix1) And EmptyArrayQ(Matrix2) Then
            Let Subtract = EmptyArray()
            Exit Function
        End If
    End If
    
    Let nd1 = NumberOfDimensions(Matrix1)
    Let nd2 = NumberOfDimensions(Matrix2)
    
    Let nr1 = NumberOfRows(Matrix1)
    Let nr2 = NumberOfRows(Matrix2)
    Let nc1 = NumberOfColumns(Matrix1)
    Let nc2 = NumberOfColumns(Matrix2)
    Let l1 = Length(Matrix1)
    Let l2 = Length(Matrix2)
    
    If nd1 = 2 And nd2 = 2 Then
        If nr1 = nr2 And nc1 = nc2 Then
            ReDim Result(1 To nr1, 1 To nc1)
            
            For r = 1 To nr1
                For c = 1 To nc1
                    Let Result(r, c) = Part(Matrix1, r, c) - Part(Matrix2, r, c)
                Next
            Next
        ElseIf nr1 = 1 Then
            ReDim Result(1 To nr2, 1 To nc2)
   
            For r = 1 To nr2
                For c = 1 To l1
                    Let Result(r, c) = Part(Matrix1, c) - Part(Matrix2, r, c)
                Next
            Next
        ElseIf nr2 = 1 Then
            ReDim Result(1 To nr1, 1 To nc1)

            For r = 1 To nr1
                For c = 1 To l2
                    Let Result(r, c) = Part(Matrix1, r, c) - Part(Matrix2, c)
                Next
            Next
        ElseIf nc1 = 1 Then
            ReDim Result(1 To nr2, 1 To nc2)

            For r = 1 To nr1
                For c = 1 To nc2
                    Let Result(r, c) = Part(Matrix1, r, 1) - Part(Matrix2, r, c)
                Next
            Next
        ElseIf nc2 = 1 Then
            ReDim Result(1 To nr1, 1 To nc1)

            For r = 1 To nr1
                For c = 1 To nc1
                    Let Result(r, c) = Part(Matrix1, r, c) - Part(Matrix2, r, 1)
                Next
            Next
        End If
    ElseIf nd1 = 1 And nd2 = 2 Then
        ReDim Result(1 To nr2, 1 To nc2)
        
        For r = 1 To nr2
            For c = 1 To nc2
                Let Result(r, c) = Part(Matrix1, c) - Part(Matrix2, r, c)
            Next
        Next
    ElseIf nd1 = 2 And nd2 = 1 Then
        ReDim Result(1 To nr1, 1 To nc1)
        
        For r = 1 To nr1
            For c = 1 To nc1
                Let Result(r, c) = Part(Matrix1, r, c) - Part(Matrix2, c)
            Next
        Next
    ElseIf nd1 = 2 And nd2 = 0 Then
        ReDim Result(1 To nr1, 1 To nc1)
    
        For r = 1 To nr1
            For c = 1 To nc1
                Let Result(r, c) = Part(Matrix1, r, c) - Matrix2
            Next
        Next
    ElseIf nd1 = 0 And nd2 = 2 Then
        ReDim Result(1 To nr2, 1 To nc2)
    
        For r = 1 To nr2
            For c = 1 To nc2
                Let Result(r, c) = Matrix1 - Part(Matrix2, r, c)
            Next
        Next
    ElseIf nd1 = 1 And nd2 = 1 Then
        ReDim Result(1 To l1)
        
        For c = 1 To l1
            Let Result(c) = Part(Matrix1, c) - Part(Matrix2, c)
        Next
    ElseIf nd1 = 0 And nd2 = 1 Then
        ReDim Result(1 To l2)
        
        For c = 1 To l2
            Let Result(c) = Matrix1 - Part(Matrix2, c)
        Next
    ElseIf nd1 = 1 And nd2 = 0 Then
        ReDim Result(1 To l1)
        
        For c = 1 To l1
            Let Result(c) = Part(Matrix1, c) - Matrix2
        Next
    Else
        Let Subtract = Matrix1 - Matrix2
        Exit Function
    End If
    
    Let Subtract = Result
End Function

' DESCRIPTION
' This function multiplies numbers, vectors or matrices elementwise, as long as it makes sense.
' Examples:

'    1. scalars - The are multiplied normally and returned as a scalar
'    2. scalars and vectors/matrices - The scalar is multiplied to each element of the other argument
'    3. vectors - Two row or two column vectors of the same dimensions are multiplied elementwise
'    4. a row vector and matrix - are multiplied elementwise for each row of the matrix. In other words,
'       the row vector is multiplied once to each row of the matrix. The result is returned as matrix.
'    5. a column vector and matrix - are multiplied elementwise for each column of the matrix. In other words,
'       the column vector is multiplied once to each column of the matrix. The result is returned as matrix.
'    6. two matrices - Addition is performed elementwise and returned as a matrix
'
' PARAMETERS
' 1. Matrix1 - a scalar, vector, or matrix
' 2. Matrix2 -  a scalar, vector, or matrix
' 3. ParamConsistencyChecksQ (optiomal) - If explicitly passed as False, this function performs no
'    parameter consistency checks
'
' RETURNED VALUE
' The result of performing elementwise multiplication on the last two arguments
Public Function Multiply(Matrix1 As Variant, Matrix2 As Variant, Optional ParamConsistencyChecksQ = False) As Variant
    Dim Arg1 As Variant
    Dim Arg2 As Variant
    Dim nd1 As Integer
    Dim nd2 As Integer
    Dim nr1 As Long
    Dim nr2 As Long
    Dim nc1 As Long
    Dim nc2 As Long
    Dim l1 As Long
    Dim l2 As Long
    Dim r As Long
    Dim c As Long
    Dim Result() As Variant

    Let Multiply = Null

    If ParamConsistencyChecksQ Then
        If Not ElementwiseArithmeticParameterConsistentQ(Matrix1, Matrix2) Then
            Let Multiply = Null
            Exit Function
        End If
            
        If EmptyArrayQ(Matrix1) And EmptyArrayQ(Matrix2) Then
            Let Multiply = EmptyArray()
            Exit Function
        End If
    End If
    
    Let nd1 = NumberOfDimensions(Matrix1)
    Let nd2 = NumberOfDimensions(Matrix2)
    
    Let nr1 = NumberOfRows(Matrix1)
    Let nr2 = NumberOfRows(Matrix2)
    Let nc1 = NumberOfColumns(Matrix1)
    Let nc2 = NumberOfColumns(Matrix2)
    Let l1 = Length(Matrix1)
    Let l2 = Length(Matrix2)
    
    If nd1 = 2 And nd2 = 2 Then
        If nr1 = nr2 And nc1 = nc2 Then
            ReDim Result(1 To nr1, 1 To nc1)
            
            For r = 1 To nr1
                For c = 1 To nc1
                    Let Result(r, c) = Part(Matrix1, r, c) * Part(Matrix2, r, c)
                Next
            Next
        ElseIf nr1 = 1 Then
            ReDim Result(1 To nr2, 1 To nc2)
   
            For r = 1 To nr2
                For c = 1 To l1
                    Let Result(r, c) = Part(Matrix1, c) * Part(Matrix2, r, c)
                Next
            Next
        ElseIf nr2 = 1 Then
            ReDim Result(1 To nr1, 1 To nc1)

            For r = 1 To nr1
                For c = 1 To l2
                    Let Result(r, c) = Part(Matrix1, r, c) * Part(Matrix2, c)
                Next
            Next
        ElseIf nc1 = 1 Then
            ReDim Result(1 To nr2, 1 To nc2)

            For r = 1 To nr1
                For c = 1 To nc2
                    Let Result(r, c) = Part(Matrix1, r, 1) * Part(Matrix2, r, c)
                Next
            Next
        ElseIf nc2 = 1 Then
            ReDim Result(1 To nr1, 1 To nc1)

            For r = 1 To nr1
                For c = 1 To nc1
                    Let Result(r, c) = Part(Matrix1, r, c) * Part(Matrix2, r, 1)
                Next
            Next
        End If
    ElseIf nd1 = 1 And nd2 = 2 Then
        ReDim Result(1 To nr2, 1 To nc2)
        
        For r = 1 To nr2
            For c = 1 To nc2
                Let Result(r, c) = Part(Matrix1, c) * Part(Matrix2, r, c)
            Next
        Next
    ElseIf nd1 = 2 And nd2 = 1 Then
        ReDim Result(1 To nr1, 1 To nc1)
        
        For r = 1 To nr1
            For c = 1 To nc1
                Let Result(r, c) = Part(Matrix1, r, c) * Part(Matrix2, c)
            Next
        Next
    ElseIf nd1 = 2 And nd2 = 0 Then
        ReDim Result(1 To nr1, 1 To nc1)
    
        For r = 1 To nr1
            For c = 1 To nc1
                Let Result(r, c) = Part(Matrix1, r, c) * Matrix2
            Next
        Next
    ElseIf nd1 = 0 And nd2 = 2 Then
        ReDim Result(1 To nr2, 1 To nc2)
    
        For r = 1 To nr2
            For c = 1 To nc2
                Let Result(r, c) = Matrix1 * Part(Matrix2, r, c)
            Next
        Next
    ElseIf nd1 = 1 And nd2 = 1 Then
        ReDim Result(1 To l1)
        
        For c = 1 To l1
            Let Result(c) = Part(Matrix1, c) * Part(Matrix2, c)
        Next
    ElseIf nd1 = 0 And nd2 = 1 Then
        ReDim Result(1 To l2)
        
        For c = 1 To l2
            Let Result(c) = Matrix1 * Part(Matrix2, c)
        Next
    ElseIf nd1 = 1 And nd2 = 0 Then
        ReDim Result(1 To l1)
        
        For c = 1 To l1
            Let Result(c) = Part(Matrix1, c) * Matrix2
        Next
    Else
        Let Multiply = Matrix1 * Matrix2
        Exit Function
    End If
    
    Let Multiply = Result
End Function

' DESCRIPTION
' This function divides numbers, vectors or matrices elementwise, as long as it makes sense.
' Examples:

'    1. scalars - The are divided normally and returned as a scalar
'    2. scalars and vectors/matrices - The scalar is divided to each element of the other argument
'    3. vectors - Two row or two column vectors of the same dimensions are divided elementwise
'    4. a row vector and matrix - are divided elementwise for each row of the matrix. In other words,
'       the row vector is divided once to each row of the matrix. The result is returned as matrix.
'    5. a column vector and matrix - are divided elementwise for each column of the matrix. In other words,
'       the column vector is divided once to each column of the matrix. The result is returned as matrix.
'    6. two matrices - Addition is performed elementwise and returned as a matrix
'
' PARAMETERS
' 1. Matrix1 - a scalar, vector, or matrix
' 2. Matrix2 -  a scalar, vector, or matrix
' 3. ParamConsistencyChecksQ (optiomal) - If explicitly passed as False, this function performs no
'    parameter consistency checks
'
' RETURNED VALUE
' The result of performing elementwise division on the last two arguments
Public Function Divide(Matrix1 As Variant, Matrix2 As Variant, Optional ParamConsistencyChecksQ = True) As Variant
    Dim Arg1 As Variant
    Dim Arg2 As Variant
    Dim nd1 As Integer
    Dim nd2 As Integer
    Dim nr1 As Long
    Dim nr2 As Long
    Dim nc1 As Long
    Dim nc2 As Long
    Dim l1 As Long
    Dim l2 As Long
    Dim r As Long
    Dim c As Long
    Dim Result() As Variant

    Let Divide = Null

    If ParamConsistencyChecksQ Then
        If Not ElementwiseArithmeticParameterConsistentQ(Matrix1, Matrix2) Then
            Let Divide = Null
            Exit Function
        End If
            
        If EmptyArrayQ(Matrix1) And EmptyArrayQ(Matrix2) Then
            Let Divide = EmptyArray()
            Exit Function
        End If
    End If
    
    Let nd1 = NumberOfDimensions(Matrix1)
    Let nd2 = NumberOfDimensions(Matrix2)
    
    Let nr1 = NumberOfRows(Matrix1)
    Let nr2 = NumberOfRows(Matrix2)
    Let nc1 = NumberOfColumns(Matrix1)
    Let nc2 = NumberOfColumns(Matrix2)
    Let l1 = Length(Matrix1)
    Let l2 = Length(Matrix2)
    
    If nd1 = 2 And nd2 = 2 Then
        If nr1 = nr2 And nc1 = nc2 Then
            ReDim Result(1 To nr1, 1 To nc1)
            
            For r = 1 To nr1
                For c = 1 To nc1
                    If Part(Matrix2, r, c) = 0 Then Exit Function
                    
                    Let Result(r, c) = Part(Matrix1, r, c) / Part(Matrix2, r, c)
                Next
            Next
        ElseIf nr1 = 1 Then
            ReDim Result(1 To nr2, 1 To nc2)
   
            For r = 1 To nr2
                For c = 1 To l1
                    If Part(Matrix2, r, c) = 0 Then Exit Function
                
                    Let Result(r, c) = Part(Matrix1, c) / Part(Matrix2, r, c)
                Next
            Next
        ElseIf nr2 = 1 Then
            ReDim Result(1 To nr1, 1 To nc1)

            For r = 1 To nr1
                For c = 1 To l2
                    If Part(Matrix2, c) = 0 Then Exit Function
                
                    Let Result(r, c) = Part(Matrix1, r, c) / Part(Matrix2, c)
                Next
            Next
        ElseIf nc1 = 1 Then
            ReDim Result(1 To nr2, 1 To nc2)

            For r = 1 To nr1
                For c = 1 To nc2
                    If Part(Matrix2, r, c) = 0 Then Exit Function
                
                    Let Result(r, c) = Part(Matrix1, r, 1) / Part(Matrix2, r, c)
                Next
            Next
        ElseIf nc2 = 1 Then
            ReDim Result(1 To nr1, 1 To nc1)

            For r = 1 To nr1
                For c = 1 To nc1
                    If Part(Matrix2, r, 1) = 0 Then Exit Function
                
                    Let Result(r, c) = Part(Matrix1, r, c) / Part(Matrix2, r, 1)
                Next
            Next
        End If
    ElseIf nd1 = 1 And nd2 = 2 Then
        ReDim Result(1 To nr2, 1 To nc2)
        
        For r = 1 To nr2
            For c = 1 To nc2
                If Part(Matrix2, r, c) = 0 Then Exit Function
            
                Let Result(r, c) = Part(Matrix1, c) / Part(Matrix2, r, c)
            Next
        Next
    ElseIf nd1 = 2 And nd2 = 1 Then
        ReDim Result(1 To nr1, 1 To nc1)
        
        For r = 1 To nr1
            For c = 1 To nc1
                If Part(Matrix2, c) = 0 Then Exit Function
            
                Let Result(r, c) = Part(Matrix1, r, c) / Part(Matrix2, c)
            Next
        Next
    ElseIf nd1 = 2 And nd2 = 0 Then
        ReDim Result(1 To nr1, 1 To nc1)
    
        For r = 1 To nr1
            For c = 1 To nc1
                If Matrix2 = 0 Then Exit Function
            
                Let Result(r, c) = Part(Matrix1, r, c) / Matrix2
            Next
        Next
    ElseIf nd1 = 0 And nd2 = 2 Then
        ReDim Result(1 To nr2, 1 To nc2)
    
        For r = 1 To nr2
            For c = 1 To nc2
                If Part(Matrix2, r, c) = 0 Then Exit Function
            
                Let Result(r, c) = Matrix1 / Part(Matrix2, r, c)
            Next
        Next
    ElseIf nd1 = 1 And nd2 = 1 Then
        ReDim Result(1 To l1)
        
        For c = 1 To l1
            If Part(Matrix2, c) = 0 Then Exit Function
        
            Let Result(c) = Part(Matrix1, c) / Part(Matrix2, c)
        Next
    ElseIf nd1 = 0 And nd2 = 1 Then
        ReDim Result(1 To l2)
        
        For c = 1 To l2
            If Part(Matrix2, c) Then Exit Function
            Let Result(c) = Matrix1 / Part(Matrix2, c)
        Next
    ElseIf nd1 = 1 And nd2 = 0 Then
        ReDim Result(1 To l1)
        
        For c = 1 To l1
            If Matrix2 = 0 Then Exit Function
            Let Result(c) = Part(Matrix1, c) / Matrix2
        Next
    Else
        If Matrix2 = 0 Then Exit Function
        Let Divide = Matrix1 / Matrix2
        Exit Function
    End If
    
    Let Divide = Result
End Function

' DESCRIPTION
' Calculates the dot product of two vectors. Returns Null if the parameters are incompatible.
'
' PARAMETERS
' 1. v1 - a vector
' 2. v2 -  a vector
' 3. ParamConsistencyChecksQ (optiomal) - If explicitly passed as False, this function performs no
'    parameter consistency checks
'
' RETURNED VALUE
' The dot product of given vectors
Public Function DotProduct(v1 As Variant, v2 As Variant, Optional ParamConsistencyChecksQ = True) As Variant
    Dim v1prime As Variant
    Dim v2prime As Variant
    Dim i As Long
    Dim TheResult As Double

    Let DotProduct = Null

    If ParamConsistencyChecksQ Then
        If Not (DimensionedQ(v1) And DimensionedQ(v2)) Then Exit Function
    
        If Length(v1) <> Length(v2) Then Exit Function
    
        If Not (VectorQ(v1) And VectorQ(v2)) Then Exit Function
    End If
    
    Let v1prime = Flatten(v1)
    Let v2prime = Flatten(v2)
    
    For i = 1 To Length(v1prime)
        Let TheResult = TheResult + v1prime(i) * v2prime(i)
    Next i
    
    Let DotProduct = TheResult
End Function

' DESCRIPTION
' Calculates the matix product of the two matrices. Returns Null if the parameters are incompatible.
'
' PARAMETERS
' 1. m1 - a matrix
' 2. m2 -  a matrix
' 3. ParamConsistencyChecksQ (optiomal) - If explicitly passed as False, this function performs no
'    parameter consistency checks
'
' RETURNED VALUE
' The matrix product of given vectors
Public Function MatrixMultiply(M1 As Variant, M2 As Variant, Optional ParamConsistencyChecksQ = True) As Variant
    Dim r As Long
    Dim c As Long
    Dim TheResult As Variant

    Let MatrixMultiply = Null

    If ParamConsistencyChecksQ Then
        If Not (DimensionedQ(M1) And DimensionedQ(M2)) Then Exit Function
        
        If NumberOfColumns(M1) <> NumberOfRows(M2) Then Exit Function
    
        If Not (MatrixQ(M1) And MatrixQ(M2)) Then Exit Function
    End If
    
    ReDim TheResult(1 To NumberOfRows(M1), 1 To NumberOfColumns(M2))
    For r = 1 To NumberOfRows(M1)
        For c = 1 To NumberOfColumns(M2)
            Let TheResult(r, c) = DotProduct(Part(M1, r), Part(M2, Span(1, -1), c))
        Next c
    Next r
    
    Let MatrixMultiply = TheResult
End Function

' DESCRIPTION
' Fills in the banks and replaced nulls in a 1D array with the last, non-empty and non-null value.
' Repeats the value in the first cell until it finds a different value. It then repeats that one
' until a new one is found.  And so forth.  Nulls appearing before non-empty, non-null values are
' replaced with Empty.
'
' If the optional parameter FillValue is not missing, Empty and Null is replaced with FillValue.
'
' PARAMETERS
' 1. AnArray - An array
' 2. FillValue (optional) - Value to use for replacing Empty and Null
'
' RETURNED VALUE
' The given array with Empty and Null replaced by either the last non-Empty, non-Null value or
' the given optional value
Public Function FillArrayBlanks(AnArray As Variant, Optional FillValue As Variant) As Variant
    Dim CurrentValue As Variant
    Dim c As Long

     FillArrayBlanks = Null

    If Not DimensionedQ(AnArray) Then Exit Function
    
    If EmptyArrayQ(AnArray) Then
        Let FillArrayBlanks = EmptyArray()
        Exit Function
    End If

    If IsNull(First(AnArray)) Or IsEmpty(First(AnArray)) Then
        Let CurrentValue = Empty
    Else
        Let CurrentValue = First(AnArray)
    End If
    
    For c = LBound(AnArray, 1) To UBound(AnArray, 1)
        If IsNull(AnArray(c)) Or IsEmpty(AnArray(c)) Then
            Let AnArray(c) = IIf(IsMissing(FillValue), CurrentValue, FillValue)
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

    Let BlankOutArraySequentialRepetitions = Null

    If Not DimensionedQ(AnArray) Then Exit Function
    
    If EmptyArrayQ(AnArray) Then
        Let BlankOutArraySequentialRepetitions = EmptyArray()
        Exit Function
    End If
    
    If Not AtomicArrayQ(AnArray) Then Exit Function
    
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
' This function replaces any of the given values by another value in each
' cell of a range.  The function returns Null if parameter ValuesToReplace
' is not an array of numbers and strings.  If ValuesToReplace is empty, the
' function leaves the range unchanged. The function returns True if the code
' run without errors.
'
'
' PARAMETERS
' 1. ValuesToReplace    - A a number, string, or 1D array of strings and numbers
' 2. ReplacementValue   - The replacement value
' 3. TheRange           - The range to modify
' 4. ReplaceErrorsQ     - Optionl Booelan flag to also replace errors. Set to
'                         True by default
'
' RETURNED VALUE
' The given range is modified, replacing any of the given values with the
' given replacement.
Public Function ReplaceRangeValues(ValuesToReplace As Variant, _
                                   ReplacementValue As Variant, _
                                   ARange As Range, _
                                   Optional ReplaceErrorsQ As Boolean = True) As Variant
    Dim ADict As Dictionary
    Dim r As Long
    Dim c As Integer
    Dim DumpArray As Variant
    
    ' Set default return value for errors
    Let ReplaceRangeValues = Null
    
    ' Exit with Null if ValuesToReplace is not as expected
    If Not (NumberOrStringQ(ValuesToReplace) Or NumberOrStringArrayQ(ValuesToReplace)) Then
        Exit Function
    End If
    
    ' Exit with Null if ReplacementValue is not as expected
    If Not NumberOrStringQ(ValuesToReplace) Then
        Exit Function
    End If
    
    If ARange Is Nothing Then Exit Function
    
    Set ADict = New Dictionary
    For r = 1 To Length(ValuesToReplace)
        If Not ADict.Exists(Key:=Part(ValuesToReplace, r)) Then
            Call ADict.Add(Key:=Part(ValuesToReplace, r), Item:=Empty)
        End If
    Next
    
    ' Replace empty or error cells with NULL in every single column (except the ticker column)
    ReDim DumpArray(1 To ARange.Rows.Count, 1 To 1)
    For c = 1 To ARange.Columns.Count
        For r = 1 To ARange.Rows.Count
            If ReplaceErrorsQ Then
                If IsError(ARange(r, c).Value2) Then
                    Let DumpArray(r, 1) = ReplacementValue
                ElseIf InStr(1, ARange(r, c).Value2, "#N/A") > 0 Then
                    Let DumpArray(r, 1) = ReplacementValue
                End If
            End If
            
            If ADict.Exists(Key:=ARange(r, c).Value2) Then
                Let DumpArray(r, 1) = ReplacementValue
            Else
                Let DumpArray(r, 1) = ARange(r, c).Value2
            End If
        Next
        
        Call DumpInSheet(DumpArray, ARange(1, c))
    Next
    
    Let ReplaceRangeValues = True
End Function

' DESCRIPTION
' This function sorts the given 2D matrix by the columns whose positions are given by
' ArrayOfColPos. The sorting orientation in each column are in ArrayOfColsSortOrder
' ArrayOfColsSortOrder is a variant array whose elements are all of enumerated type
' XLSortOrder (e.g. xlAscending, xlDescending)
'
' PARAMETERS
' 1. AnArray       - An array satisfying Predicates.AtomicArrayQ
'
' RETURNED VALUE
' AnAtomicArrayOrTable after blanking out sequential repetitions of its elements.
Public Function Sort2DArray(MyArray As Variant, _
                            ArrayOfColPos As Variant, _
                            ArrayOfColsSortOrder As Variant, _
                            WithHeaders As XlYesNoGuess) As Variant
    Dim TheRange As Range
    Dim i As Integer
    
    ' Clear any existing formats and content
    Call TempComputation.UsedRange.ClearFormats
    Call TempComputation.UsedRange.ClearContents
    
    ' Dump array in temp sheet
    Call DumpInSheet(MyArray, TempComputation.Range("A1"), True)

    ' Set range pointer to the data we just dumped in the temp sheet
    Set TheRange = TempComputation.Range("A1").Resize(NumberOfRows(MyArray), NumberOfColumns(MyArray))

    ' Clear any previous sorting criteria
    Call TempComputation.Sort.SortFields.Clear

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
    ' We cannot use .CurrentRegion.Value2 because some of the columns on the right may be Empty on purpose
    Let Sort2DArray = TempComputation.Range("A1").Resize(NumberOfRows(MyArray), NumberOfColumns(MyArray)).Value2
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
    
    If Length(AnArray) = 0 Then
        Set DumpInSheet = Nothing
        Exit Function
    End If

    If PreserveColumnTextFormats Then
        Set DumpInSheet = DumpInSheetHelper(AnArray, TopLeftCorner, PreserveColumnTextFormats:=PreserveColumnTextFormats)
    Else
        Set DumpInSheet = DumpInSheetHelper(AnArray, TopLeftCorner)
    End If
End Function

' This function dumps an array (1D or 2D) into the worksheet with the range referenced by TopLeftCorner as the cell in the upper-left
' corner. It then returns a reference to the underlying range. Dimensions are preserved.  This means that an m x n array is dumped into
' an m x n range.
'
' This function is a helper for Arrays.DumpInSheet()
Private Function DumpInSheetHelper(AnArray As Variant, TopLeftCorner As Range, Optional PreserveColumnTextFormats As Boolean = False) As Range
    Dim c As Integer
    Dim NumRows As Long
    Dim NumColumns As Long
    Dim NumDimensions As Integer

    Let NumDimensions = NumberOfDimensions(AnArray)
    Let NumRows = NumberOfRows(AnArray)
    Let NumColumns = NumberOfColumns(AnArray)
    
    If PreserveColumnTextFormats Then
        If NumDimensions = 0 Then
            Let TempComputation.Range("A1").NumberFormat = IIf(TypeName(AnArray) = "String", "@", "0")
            Let TempComputation.Range("A1").Value2 = AnArray
        Else
            ' Loop over the target columns, applying the format of the array's first column element to
            ' the entire, corresponding target range column
            For c = 0 To NumColumns - 1
                If NumDimensions = 1 Then
                    Let TopLeftCorner.Offset(0, c).Resize(NumRows, 1).NumberFormat = IIf(TypeName(AnArray(c + LBound(AnArray))) = "String", "@", "0")
                Else
                    Let TopLeftCorner.Offset(0, c).Resize(NumRows, 1).NumberFormat = IIf(TypeName(AnArray(1, c + LBound(AnArray, 2))) = "String", "@", "0")
                End If
            Next c
        End If
    End If

    If NumDimensions = 1 Then
        Let TopLeftCorner(1, 1).Resize(1, NumColumns).Value2 = AnArray
        Set DumpInSheetHelper = TopLeftCorner(1, 1).Resize(1, NumColumns)
    Else
        Let TopLeftCorner(1, 1).Resize(NumRows, NumColumns).Value2 = AnArray
        Set DumpInSheetHelper = TopLeftCorner(1, 1).Resize(NumRows, NumColumns)
    End If
End Function

' This function sorts a 1D array in the ascending or descending order, assuming no header.
' This returns a 1D array with the sorted values
Public Function Sort1DArray(MyArray As Variant, _
                            Optional SortOrder As XlSortOrder = xlAscending, _
                            Optional CheckParameterQ As Boolean = False) As Variant
    Dim TheRange As Range
    Dim TheFormats() As String
    Dim r As Integer
    Dim LastRowNumber As Long
    
    ' Default return value for error cases
    Let Sort1DArray = MyArray
    
    ' Exit with Null if MyArray is not dimensioned
    If Not DimensionedQ(MyArray) Then Exit Function
    
    ' Exit with EmptyArray() if MyArray is empty
    If EmptyArrayQ(MyArray) Then Exit Function
    
    ' Verify this is an alphanumeric, 1D array. Else exit with Null
    If CheckParameterQ Then If Not NumberOrStringArrayQ(MyArray) Then Exit Function
    
    ' Set pointer to temp sheet and clear its used range
    Call TempComputation.UsedRange.ClearFormats
    Call TempComputation.UsedRange.ClearContents
    
    ' Set range pointer to the data we just dumped in the temp sheet
    Set TheRange = TempComputation.Range("A1").Resize(Length(MyArray), 1)
    
    ' Get the formats of the array
    ReDim TheFormats(1 To Length(MyArray), 1 To 1)
    For r = 1 To Length(MyArray)
        Let TheFormats(r, 1) = IIf(StringQ(Part(MyArray, r)), "@", "General")
    Next
    
    ' Format the target range
    Let TempComputation.Range("A1").Resize(Length(MyArray), 1).NumberFormat = TheFormats
    
    ' Dump array in temp sheet
    Call ToTemp(TransposeMatrix(MyArray))
    
    ' Clear any previous sorting criteria
    TempComputation.Sort.SortFields.Clear

    ' Add all the sorting criteria
    Call TheRange.Worksheet.Sort.SortFields.Add(Key:=TheRange, _
                                                SortOn:=xlSortOnValues, _
                                                Order:=SortOrder, _
                                                DataOption:=xlSortNormal)
    
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
    Let Sort1DArray = Flatten(IIf(NumberOfDimensions(TheRange.Value2) = 0, Array(TheRange.Value2), TheRange.Value2))
End Function

' This function returns the array resulting from concatenating B to the right of A.
' A and B must have the same dimensions (e.g. 1 or 2D)
' If dim(A)<>dim(B) or dim(A)>2 or dim(B)>2 or dim(A)<1 or dim(B)<1 then
' this function returns EmptyArray()
Public Function ConcatenateArrays(a As Variant, _
                                  b As Variant, _
                                  Optional ParameterCheckQ As Boolean = True) As Variant
    Dim i As Long
    Dim ResultArray As Variant
    
    Let ConcatenateArrays = Null
    
    If Not (DimensionedQ(a) And DimensionedQ(b)) Then Exit Function
    
    If EmptyArrayQ(a) Then
        Let ConcatenateArrays = b
        Exit Function
    End If
    
    If EmptyArrayQ(b) Then
        Let ConcatenateArrays = a
        Exit Function
    End If
    
    If NumberOfDimensions(a) = 1 And NumberOfDimensions(b) = 1 Then
        ReDim ResultArray(1 To Length(a) + Length(b))
        
        For i = 1 To Length(a)
            Let ResultArray(i) = a(NormalizeIndex(a, i, 1, ParameterCheckQ))
        Next
        
        For i = 1 To Length(b)
            Let ResultArray(i + Length(a)) = b(NormalizeIndex(b, i, 1, ParameterCheckQ))
        Next
        
        Let ConcatenateArrays = ResultArray
        Exit Function
    End If

    Let ConcatenateArrays = TransposeMatrix(StackArrays(TransposeMatrix(a, False, ParameterCheckQ), _
                                                        TransposeMatrix(b, False, ParameterCheckQ), _
                                                        ParameterCheckQ), _
                                            False, _
                                            ParameterCheckQ)
End Function

' This function sorts the given range by the columns whose positions are given by ArrayOfColPos.
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
Public Function SwapMatrixColumns(TheMatrix As Variant, FirstColumnIndex As Long, SecondColumIndex As Long) As Variant
    Dim col1 As Variant
    
    If FirstColumnIndex < 1 Or FirstColumnIndex > NumberOfColumns(TheMatrix) Or SecondColumIndex < 1 Or SecondColumIndex > NumberOfColumns(TheMatrix) Or FirstColumnIndex = SecondColumIndex Then
        Let SwapMatrixColumns = False

        Exit Function
    End If
    
    Call ToTemp(TheMatrix)
    Let col1 = Part(TheMatrix, Span(1, -1), FirstColumnIndex)
    Call DumpInSheet(Part(TheMatrix, Span(1, -1), SecondColumIndex), TempComputation.Cells(1, FirstColumnIndex))
    Call DumpInSheet(col1, TempComputation.Cells(1, SecondColumIndex))
    
    Let SwapMatrixColumns = TempComputation.Range("A1").CurrentRegion.Value2
End Function

' This function swaps two columns from a range, returning  the resulting matrix if the operation is successful or False if it is not
Public Function SwapMatrixRows(TheMatrix As Variant, FirstRowIndex As Long, SecondRowIndex As Long) As Variant
    Dim Row1 As Variant

    If FirstRowIndex < 1 Or FirstRowIndex > NumberOfRows(TheMatrix) Or SecondRowIndex < 1 Or SecondRowIndex > NumberOfRows(TheMatrix) Or FirstRowIndex = SecondRowIndex Then
        Let SwapMatrixRows = False

        Exit Function
    End If

    Call ToTemp(TheMatrix)
    Let Row1 = Part(TheMatrix, FirstRowIndex)
    Call DumpInSheet(Part(TheMatrix, SecondRowIndex), TempComputation.Cells(FirstRowIndex, 1))
    Call DumpInSheet(Row1, TempComputation.Cells(SecondRowIndex, 1))
    
    Let SwapMatrixRows = TempComputation.Range("A1").CurrentRegion.Value2
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
    
    Let TheLength = Length(First(AnArray))
    
    For Each var In AnArray
        If Not AtomicArrayQ(var) Or Length(var) <> TheLength Then
            Let TransposeRectangular1DArrayOf1DArrays = Null
            Exit Function
        End If
    Next
    
    Let TheMatrix = Pack2DArray(AnArray)
    ReDim TheResult(1 To NumberOfColumns(TheMatrix))
    For c = 1 To Length(First(AnArray))
        Let TheResult(c) = Flatten(Part(TheMatrix, Span(1, -1), c))
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

    If NumberOfDimensions(TheMatrix) <> 2 Or NumberOfRows(TheMatrix) < 2 Or _
       NumberOfDimensions(ColumnsToFilter) <> 1 Or NumberOfDimensions(Criteria1List) <> 1 Or _
       Length(ColumnsToFilter) <> Length(Criteria1List) Then
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
           Length(OperatorList) <> Length(Criteria1List) Or _
           Length(Criteria2List) <> Length(Criteria1List) Then
            Let AutofilterMatrix = EmptyArray()
            Exit Function
        End If
    End If
    
    ' If the code gets to this point, parameters are consistent.

    ' Get dimensions of source matrix
    Let NCols = NumberOfColumns(TheMatrix)
    Let NRows = NumberOfRows(TheMatrix)

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
    
    For r = 1 To NumberOfRows(ParamCopy)
        For c = 1 To NumberOfColumns(ParamCopy)
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

    For i = 1 To NumberOfRows(TheArray)
        Let TheList = TheList & Convert1DArrayIntoParentheticalExpression(Part(TheArray, CLng(i)))
        
        If i < NumberOfRows(TheArray) Then
            Let TheList = TheList & ", "
        End If
    Next i
    
    Let Convert2DArrayIntoListOfParentheticalExpressions = TheList
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
    If a2DArray1KeyColPos < 1 Or a2DArray1KeyColPos > NumberOfColumns(a2DArray1) Or _
       a2DArray2KeyColPos < 1 Or a2DArray2KeyColPos > NumberOfColumns(a2DArray2) Then
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
        If var < 1 Or var > NumberOfColumns(a2DArray1) Then
            Let LeftJoin2DArraysOnKeyEquality = Null
            Exit Function
        End If
    Next
    
    ' Exit with Null if any of the indices in ColsPosArrayFrom2DArray2 is less than 1
    ' or larger than the number of columns in array2
    For Each var In ColsPosArrayFrom2DArray2
        If var < 1 Or var > NumberOfColumns(a2DArray2) Then
            Let LeftJoin2DArraysOnKeyEquality = Null
            Exit Function
        End If
    Next
    
    ' Exit with Null if either a2DArray1 or a2DArray2 has only 1 row when ArraysHaveHeadersQ
    ' is True
    If ArraysHaveHeadersQ And _
       (NumberOfRows(a2DArray1) < 2 Or NumberOfRows(a2DArray1) < 2) Then
        Let LeftJoin2DArraysOnKeyEquality = Null
        Exit Function
    End If

    ' Determine the number of columns from array 1
    Let NumColsArray1 = Length(ColsPosArrayFrom2DArray1)
    
    ' Determine the number of columns from array 2 to append to array 1
    Let NumColsArray2 = Length(ColsPosArrayFrom2DArray2)

    ' Load all information from a2DArray1 into a dictionary
    Set ResultsDict = New Dictionary
    For r = IIf(ArraysHaveHeadersQ, LBound(a2DArray1, 1) + 1, LBound(a2DArray1, 1)) To UBound(a2DArray1, 1)
        ' Get the for the current row
        Let TheKey = a2DArray1(r, a2DArray1KeyColPos)

        If Not ResultsDict.Exists(Key:=TheKey) Then
            ' Extract the columns needed from this security's row
            Let TheItems = Take(Part(a2DArray1, r), _
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
            Let AppendedItems = Take(Part(a2DArray2, r), ColsPosArrayFrom2DArray2)
            
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
    Let JoinedHeadersRow = ConcatenateArrays(Take(Part(a2DArray1, 1), _
                                                  Prepend(ColsPosArrayFrom2DArray1, a2DArray1KeyColPos)), _
                                             Take(Part(a2DArray2, 1), _
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
    If a2DArray1KeyColPos < 1 Or a2DArray1KeyColPos > NumberOfColumns(a2DArray1) Or _
       a2DArray2KeyColPos < 1 Or a2DArray2KeyColPos > NumberOfColumns(a2DArray2) Then
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
        If var < 1 Or var > NumberOfColumns(a2DArray1) Then
            Let InnerJoin2DArraysOnKeyEquality = Null
            Exit Function
        End If
    Next

    ' Exit with Null if any of the indices in ColsPosArrayFrom2DArray2 is less than 1
    ' or larger than the number of columns in array2
    For Each var In ColsPosArrayFrom2DArray2
        If var < 1 Or var > NumberOfColumns(a2DArray2) Then
            Let InnerJoin2DArraysOnKeyEquality = Null
            Exit Function
        End If
    Next

    ' Exit with Null if either a2DArray1 or a2DArray2 has only 1 row when ArraysHaveHeadersQ
    ' is True
    If ArraysHaveHeadersQ And _
       (NumberOfRows(a2DArray1) < 2 Or NumberOfRows(a2DArray1) < 2) Then
        Let InnerJoin2DArraysOnKeyEquality = Null
        Exit Function
    End If

    ' Determine the number of columns in arrays 1 and 2
    Let NumColsArray1 = Length(ColsPosArrayFrom2DArray1)
    Let NumColsArray2 = Length(ColsPosArrayFrom2DArray2)

    ' Index the contents of array2
    Set Array2Dict = New Dictionary
    For r = IIf(ArraysHaveHeadersQ, LBound(a2DArray2, 1) + 1, LBound(a2DArray2, 1)) To UBound(a2DArray2, 1)
        ' Get the for the current row
        Let TheKey = a2DArray2(r, a2DArray2KeyColPos)

        If Not Array2Dict.Exists(Key:=TheKey) Then
            ' Add the array of values to this security's entry
            Call Array2Dict.Add(Key:=TheKey, _
                                Item:=Take(Part(a2DArray2, r), ColsPosArrayFrom2DArray2))
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
            Let TheItems = Take(Part(a2DArray1, r), _
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
    Let JoinedHeadersRow = ConcatenateArrays(Take(Part(a2DArray1, 1), _
                                                  Prepend(ColsPosArrayFrom2DArray1, a2DArray1KeyColPos)), _
                                             Take(Part(a2DArray2, 1), _
                                                  ColsPosArrayFrom2DArray2))

    ' Prepend headers to return matrix
    Let InnerJoin2DArraysOnKeyEquality = Prepend(TheResults, JoinedHeadersRow)
End Function

' DESCRIPTION
' Function to replace specific parts of an array. It works on 1D and 2D arrays.  The function returns Null
' when it fails due to parameter incompatibility. It can operate on a number of modes.  In particular, it
' can replace single locations in an array or sets of locations or spans.  The set of locations to replace
' must have the same shape as the items that replace.
'
' PARAMETERS
' 1. ArrayToChange    - Array to change
' 2. ReplacementItems - Values to replace existing ones in ArrayToChange
' 3. TheIndices       - It can be:
'                       a. single positive integer
'                       b. a 1D array of positive integers
'                       c. a 1D array or ordered pairs of indices (positive integers)
'                       c. an instance of Span
' 4. SecondDimIndices    - an optional instance of Span to refer to indices along the second dimension
'
' RETURNED VALUE
' Returns ArrayToChange after making the desired substitutions
Public Function SetPart(ByVal ArrayToChange As Variant, _
                        ReplacementItems As Variant, _
                        Locations1stDim As Variant, _
                        Optional Locations2ndDim As Variant, _
                        Optional ParameterCheckQ As Boolean = False) As Variant
    ' Set default return value in case of error
    Let SetPart = Null
        
    If IsNull(ArrayToChange) Then Exit Function
    If IsNull(ReplacementItems) Then Exit Function
    If IsNull(Locations1stDim) Then Exit Function
    If IsNull(Locations1stDim) Then Exit Function
    If IsNull(Locations2ndDim) Then Exit Function

    If ParameterCheckQ Then
        ' Exit with Null if ArrayToChange is not an array
        If Not IsArray(ArrayToChange) Then Exit Function
    End If

    ' Return ArrayToChange unchanged if no items were requested changed
    If EmptyArrayQ(ReplacementItems) Then
        Let SetPart = ArrayToChange
        
        Exit Function
    End If
    
    ' If the code gets to this point, then either parameters are as expected
    ' or the user chose not to invoke parameter checking
    
    ' Optional SecondDimIndices not passed
    If IsMissing(Locations2ndDim) Then
        Let SetPart = SetPartsUsing1stDim(ArrayToChange, _
                                          ReplacementItems, _
                                          Locations1stDim)
    ' Optional SecondDimIndices passed
    Else
        Let SetPart = SetPartsUsing2ndDim(ArrayToChange, _
                                          ReplacementItems, _
                                          Locations1stDim, _
                                          Locations2ndDim)
    End If
End Function

' DESCRIPTION
' Helper function for Public Function SetParts() in this module.
'
' This is the implementation of set when optional parameter Locations2ndDim is not passed to
' function SetParts in this module.
'
' Function to replace specific parts of 1D array. The function returns Null when it fails due
' to parameter incompatibility. It can operate on a number of modes.  In particular, it can
' replace single locations in an array or sets of locations or spans.  The set of locations to
' replace must have the same shape as the items that replace.
'
' PARAMETERS
' 1. ArrayToChange    - Array to change
' 2. ReplacementItems - Values to replace existing ones in ArrayToChange
' 3. TheIndices       - Locations to replace
'
' RETURNED VALUE
' Returns ArrayToChange after making the desired substitutions
Private Function SetPartsUsing1stDim(ArrayToChange As Variant, _
                                               ReplacementItems As Variant, _
                                               TheIndices As Variant) As Variant
    ' Set default return value in case of error
    Let SetPartsUsing1stDim = Null
        
    If IsNull(ArrayToChange) Then Exit Function
    If IsNull(ReplacementItems) Then Exit Function
    If IsNull(TheIndices) Then Exit Function
                                               
    If NumberOfDimensions(ArrayToChange) = 1 Then
        Let SetPartsUsing1stDim = SetPartsUsing1stDimIn1D(ArrayToChange, _
                                                          ReplacementItems, _
                                                          TheIndices)
    ElseIf NumberOfDimensions(ArrayToChange) = 2 Then
        Let SetPartsUsing1stDim = SetPartsUsing1stDimIn2D(ArrayToChange, _
                                                          ReplacementItems, _
                                                          TheIndices)
    End If
End Function

' DESCRIPTION
' Function to replace specific parts of an array. It works on 1D.  The function returns Null when it fails
' due to parameter incompatibility. It can operate on a number of modes.  In particular, it can replace
' single locations in an array or sets of locations or spans.  The set of locations to replace must have
' the same shape as the items used for replacement.
'
' PARAMETERS
' 1. ArrayToChange    - Array to change
' 2. ReplacementItems - Values to replace existing ones in ArrayToChange
' 3. TheIndices       - Locations to replace
'
' This function can handle any of the following calling modalities acting along
' positions of a 1D array or rows of a 1D array:
'
' 1. SINGLE LOCATION IN 1D ARRAY
'
'    ArrayToChange    <- Array(a_1, a_2, ...)
'    ReplacementItems <- e
'    TheIndices       <- r
'
'    Replace a_r with e
'
' 2. ARRAY OF LOCATIONS IN 1D ARRAY
'
'    ArrayToChange    <- Array(a_1, a_2, ...)
'    ReplacementItems <- Array(e_1, e_2, ...)
'    TheIndices       <- Array(r_1, r_2, ...)
'
'    Replace element a_(r_i) with e_i for all i
'
' 3. SINGLE LOCATIONS IN 1D ARRAY REPLACED BY ROWS FROM 2D ARRAY
'
'    Array of Single Locations in 1D Array
'    ArrayToChange    <- Array(a_1, a_2, ...)
'    ReplacementItems <- [{e^1_1, e^1_2, ...; e^2_1, e^2_2, ...;...}]
'    TheIndices       <- Array(r_1, r_2, ...)
'
'    Replace a_(r_i) with row Array(e^i_1, e^i_2, ...) for all i
'
' 4. ARRAY OF SINGLE 2D LOCATIONS IN 1D ARRAY REPLACED BY ELEMENTS
'
'    ArrayToChange    <- Array(Array(a^1_1, a^1_2,...), Array(a^2_1, a^2_2,...), ...)
'    ReplacementItems <- Array(e_1, e_2, ...)
'    TheIndices       <- Array(Array(r_1,c_1), Array(r_2,c_2), ...)
'
'    Replace a^(r_i)_(c_i) with e_i for every i in 1 through k
'
' 5. SPAN OF SINGLE LOCATIONS IN 1D ARRAY REPLACED BY ELEMENTS
'
'    ArrayToChange    <- Array(a_1, a_2, ...)
'    ReplacementItems <- Array(e_1, e_2, ...)
'    TheIndices       <- Span(...)
'
'    Replace a_Span(i) with e_i for each i, where Span(i) is the ith
'    element in the span
'
' 6. SPAN SINGLE LOCATIONS IN 1D ARRAY REPLACED BY ROWS FROM 2D ARRAY
'
'    ArrayToChange    <- Array(a_1, a_2, ..., a_n)
'    ReplacementItems <- [{e^1_1, e^1_2, ...; e^2_1, e^2_2, ...;...}]
'    TheIndices       <- Span(...)
'
'    Replace a_Span(i) with row Array(e^i_1, e^i_2, ...) for each i, where Span(i) is the ith
'    element in the span
'
' RETURNED VALUE
' Returns ArrayToChange after making the desired substitutions
Private Function SetPartsUsing1stDimIn1D(ByVal ArrayToChange As Variant, _
                                         ByVal ReplacementItems As Variant, _
                                         ByVal TheIndices As Variant) As Variant
    Dim NormalizedIndices As Variant
    Dim IndexTuple As Variant
    Dim r As Long
    Dim c As Long
    
    ' Set default return value in case of error
    Let SetPartsUsing1stDimIn1D = Null
    
    If IsNull(ArrayToChange) Then Exit Function
    If IsNull(ReplacementItems) Then Exit Function
    If IsNull(TheIndices) Then Exit Function

    ' Case 1 - SINGLE LOCATION IN 1D ARRAY
    If IsNumeric(TheIndices) Then
        Let NormalizedIndices = NormalizeIndex(ArrayToChange, TheIndices)
        If IsNull(NormalizedIndices) Then Exit Function
    
        Let ArrayToChange(NormalizedIndices) = ReplacementItems

        Let SetPartsUsing1stDimIn1D = ArrayToChange
    ' Case 2 in the function description
    ElseIf NumberOfDimensions(ReplacementItems) = 1 And NonzeroWholeNumberArrayQ(TheIndices) And Not EmptyArrayQ(TheIndices) Then
        Let NormalizedIndices = NormalizeIndexArray(ArrayToChange, TheIndices)
        
        For r = 1 To Length(TheIndices)
            If IsNull(Part(NormalizedIndices, r)) Then Exit Function
            
            Let ArrayToChange(Part(NormalizedIndices, r)) = Part(ReplacementItems, r)
        Next

        Let SetPartsUsing1stDimIn1D = ArrayToChange
    ' Case 3 in the function description
    ElseIf NumberOfDimensions(ReplacementItems) = 2 And NonzeroWholeNumberArrayQ(TheIndices) Then
        Let NormalizedIndices = NormalizeIndexArray(ArrayToChange, TheIndices)

        For r = 1 To Length(TheIndices)
            If IsNull(Part(NormalizedIndices, r)) Then Exit Function
            
            Let ArrayToChange(Part(NormalizedIndices, r)) = Part(ReplacementItems, r)
        Next

        Let SetPartsUsing1stDimIn1D = ArrayToChange
    ' Case 4 in the function description
    ElseIf NumberOfDimensions(First(ArrayToChange)) = 1 And _
           IsArray(First(TheIndices)) Then
        Let NormalizedIndices = NormalizeIndicesTupleArray(ArrayToChange, TheIndices)
        
        For r = 1 To Length(TheIndices)
            Let IndexTuple = Part(NormalizedIndices, r)
            
            If Length(IndexTuple) <> 2 Then Exit Function
            
            Let ArrayToChange(First(IndexTuple))(Last(IndexTuple)) = Part(ReplacementItems, r)
        Next

        Let SetPartsUsing1stDimIn1D = ArrayToChange
        Let SetPartsUsing1stDimIn1D = ArrayToChange
    ' Case 5 in the function description
    ElseIf SpanQ(TheIndices) And NumberOfDimensions(ReplacementItems) = 1 Then
        Let SetPartsUsing1stDimIn1D = _
            SetPartsUsing1stDimIn1D(ArrayToChange, _
                                    ReplacementItems, _
                                    CreateIndexSequenceFromSpan(ArrayToChange, _
                                                                TheIndices))
    ' Case 6 in the function description
    ElseIf SpanQ(TheIndices) And NumberOfDimensions(ReplacementItems) = 2 Then
        Let SetPartsUsing1stDimIn1D = _
            SetPartsUsing1stDimIn1D(ArrayToChange, _
                                    ReplacementItems, _
                                    CreateIndexSequenceFromSpan(ArrayToChange, _
                                                                          TheIndices))
    End If
End Function

' DESCRIPTION
' Function to replace specific parts of an array. It works on 1D.  The function returns Null when it fails
' due to parameter incompatibility. It can operate on a number of modes.  In particular, it can replace
' single locations in an array or sets of locations or spans.  The set of locations to replace must have
' the same shape as the items that replace.
'
' PARAMETERS
' 1. ArrayToChange    - Array to change
' 2. ReplacementItems - Values to replace existing ones in ArrayToChange
' 3. TheIndices       - It can be:
'                       a. single positive integer
'                       b. a 1D array of positive integers
'                       c. an instance of Span
'                       d. an array of ordered pairs
'
' This function can handle any of the following calling modalities acting along
' positions of a 1D array or rows of a 1D array:
'
' 1. SINGLE ROW LOCATION IN 2D ARRAY
'
'    ArrayToChange    <- [{a^1_1,a^1_2,..., a^1_n; a^2_1, e^2_2, ..., e^2_n; ...}]
'    ReplacementItems <- e
'    TheIndices       <- i
'
'    Replace ith row (e.g. {a^i_1, a^i_2, ...}) with replacement row e
'
' 2. ARRAY OF ROWS IN 2D ARRAY
'
'    ArrayToChange    <- [{a^1_1, a^1_2, ...; a^2_1, a^2_2, ...; ...}]
'    ReplacementItems <- Array(e_1, e_2, ...)
'    TheIndices       <- Array(i_1, i_2, ...)
'
'    Replace kth row in a (e.g. {a^k_1, a^k_2, ...}) with replacement row e_k for all k
'
' 3. SINGLE LOCATIONS IN 1D ARRAY REPLACED BY ROWS FROM 2D ARRAY
'
'    Array of Single Locations in 1D Array
'    ArrayToChange    <- [{a^1_1, a^1_2, ...; a^2_1, a^2_2, ...; ...}]
'    ReplacementItems <- [{e^1_1, e^1_2, ...; e^2_1, e^2_2, ...;...}]
'    TheIndices       <- Array(r_1, r_2, ...)
'
'    Replace row Array(a^(r_i)_1, a^(r_i)_2, ...) with row Array(e^i_1, e^i_2, ...) for all i
'
' 4. ARRAY OF 2D LOCATIONS IN 1D ARRAY REPLACED BY ELEMENTS
'
'    ArrayToChange    <- [{a^1_1, a^1_2, ...; a^2_1, a^2_2, ...; ...}]
'    ReplacementItems <- Array(e_1, e_2, ...)
'    TheIndices       <- Array(Array(r_1,c_1), Array(r_2,c_2), ...)
'
'    Replace element a^(r_i)_(c_i) with e_i for for all i
'
' 5. SPAN OF ROW LOCATIONS IN 2D ARRAY REPLACED BY ROWS
'
'    ArrayToChange    <- [{a^1_1, a^1_2, ...; a^2_1, a^2_2,; ...}]
'    ReplacementItems <- Array(e_1, e_2, ...)
'    TheIndices       <- Span(...)
'
'    Replace row {a^Span(k)_1, a^Span(k)_2, ...} with row e_k for all k,
'    where Span(k) is the kth index in the sequence corresponding to the
'    span.
'
' 6. SPAN OF ROW LOCATIONS IN 2D ARRAY REPLACED BY ROWS FROM 2D ARRAY
'
'    ArrayToChange    <- [{a^1_1, a^1_2, ...; a^2_1, a^2_2,; ...}]
'    ReplacementItems <- [{e^1_1, e^1_2,...; e^2_1, e^2_2, ...; ...}]
'    TheIndices       <- Span(...)
'
'    Replace the rows {a^Span(k)_1, a^Span(k)_2, ...} with {e^k_1, e^k_2, ...} for
'    all k, where Span(k) is the kth element in the sequence of indices corresponding
'    to the span.
'
' RETURNED VALUE
' Returns ArrayToChange after making the desired substitutions
Private Function SetPartsUsing1stDimIn2D(ByVal ArrayToChange As Variant, _
                                         ByVal ReplacementItems As Variant, _
                                         ByVal TheIndices As Variant) As Variant
    Dim NormalizedIndices As Variant
    Dim IndexTuple As Variant
    Dim AColIndex As Long
    Dim ARowIndex As Long
    Dim NumColumns As Long
    Dim ReplacementRow As Variant
    Dim ARow As Variant
    Dim r As Long
    Dim c As Long
    
    ' Set default return value in case of error
    Let SetPartsUsing1stDimIn2D = Null
        
    If IsNull(ArrayToChange) Then Exit Function
    If IsNull(ReplacementItems) Then Exit Function
    If IsNull(TheIndices) Then Exit Function

    ' Case 1 in the function description
    If IsNumeric(TheIndices) Then
        Let NumColumns = NumberOfColumns(ArrayToChange)
        Let ARowIndex = NormalizeIndex(ArrayToChange, TheIndices)
        
        ' Exit with Null if replacement row does not have the same length as a row of
        ' ArrayToChange
        If Length(ReplacementItems) <> NumColumns Then Exit Function
        
        For c = 1 To NumColumns
            Let AColIndex = NormalizeIndex(ArrayToChange, c, DimensionIndexRelativeTo:=2)
            
            Let ArrayToChange(ARowIndex, AColIndex) = Part(ReplacementItems, c)
        Next
        
        Let SetPartsUsing1stDimIn2D = ArrayToChange
    ' Case 2 in the function description
    ElseIf NumberOfDimensions(ReplacementItems) = 1 And NonzeroWholeNumberArrayQ(TheIndices) Then
        Let NormalizedIndices = NormalizeIndexArray(ArrayToChange, TheIndices)
        Let NumColumns = NumberOfColumns(ArrayToChange)

        For r = 1 To Length(TheIndices)
            Let ARowIndex = Part(NormalizedIndices, r)
            Let ReplacementRow = Part(ReplacementItems, r)
            
            ' Exit with Null if replacement row does not have the same length as a row of
            ' ArrayToChange
            If Length(ReplacementRow) <> NumColumns Then Exit Function
            
            For c = 1 To NumColumns
                Let AColIndex = NormalizeIndex(ArrayToChange, c, DimensionIndexRelativeTo:=2)
                
                Let ArrayToChange(ARowIndex, AColIndex) = Part(ReplacementRow, c)
            Next
        Next

        Let SetPartsUsing1stDimIn2D = ArrayToChange
    ' Case 3 in the function description
    ElseIf NumberOfDimensions(ReplacementItems) = 2 And NonzeroWholeNumberArrayQ(TheIndices) Then
        Let NumColumns = NumberOfColumns(ArrayToChange)
    
        ' Exit with Null if replacement rows and rows to replace do not have the same number
        ' of columns
        If NumberOfColumns(ArrayToChange) <> NumberOfColumns(ReplacementItems) Then Exit Function
        
        ' Exit with Null if the number of indices does not equal to the number of replacement rows
        If Length(TheIndices) <> Length(ReplacementItems) Then Exit Function

        Let NormalizedIndices = NormalizeIndexArray(ArrayToChange, TheIndices)

        For r = 1 To Length(TheIndices)
            Let ARowIndex = Part(NormalizedIndices, r)
            Let ReplacementRow = Part(ReplacementItems, r)
            
            For c = 1 To NumColumns
                Let AColIndex = NormalizeIndex(ArrayToChange, c, DimensionIndexRelativeTo:=2)
                
                Let ArrayToChange(ARowIndex, AColIndex) = Part(ReplacementRow, c)
            Next
        Next

        Let SetPartsUsing1stDimIn2D = ArrayToChange
    ' Case 4 in the function description
    ElseIf IsArray(First(TheIndices)) And NumberOfDimensions(ReplacementItems) = 1 Then
        ' Exit with Null if number of replacements does not equal number of 2-tuples
        If Length(TheIndices) <> Length(ReplacementItems) Then Exit Function
        
        Let NormalizedIndices = NormalizeIndicesTupleArray(ArrayToChange, TheIndices)
    
        ' Exit if an index is not a 2-tuple
        For r = 1 To Length(TheIndices)
            If Length(Part(TheIndices, r)) <> 2 Then Exit Function
        Next
    
        For r = 1 To Length(TheIndices)
            Let IndexTuple = Part(NormalizedIndices, r)
            
            Let ArrayToChange(First(IndexTuple), Last(IndexTuple)) = _
                Part(ReplacementItems, r)
        Next

        Let SetPartsUsing1stDimIn2D = ArrayToChange
    ' Case 5 in the function description
    ElseIf SpanQ(TheIndices) And NumberOfDimensions(ReplacementItems) = 1 Then
        Let SetPartsUsing1stDimIn2D = _
            SetPartsUsing1stDimIn2D(ArrayToChange, _
                                              ReplacementItems, _
                                              CreateIndexSequenceFromSpan(ArrayToChange, _
                                                                          TheIndices))
    ' Case 6 in the function description
    ElseIf SpanQ(TheIndices) And NumberOfDimensions(ReplacementItems) = 2 Then
        Let SetPartsUsing1stDimIn2D = _
            SetPartsUsing1stDimIn2D(ArrayToChange, _
                                              ReplacementItems, _
                                              CreateIndexSequenceFromSpan(ArrayToChange, _
                                                                          TheIndices))
    End If
End Function

' DESCRIPTION
' Helper function for Public Function SetParts() in this module.
' This is the implementation of set when optional parameter Locations2ndDim is passed to function
' SetParts in this module.
'
' Function to replace specific parts of an array. It works on 1D.  The function returns Null when it fails
' due to parameter incompatibility. It can operate on a number of modes.  In particular, it can replace
' single locations in an array or sets of locations or spans.  The set of locations to replace must have
' the same shape as the items that replace.
'
' PARAMETERS
' 1. ArrayToChange    - Array to change
' 2. ReplacementItems - Values to replace existing ones in ArrayToChange
' 3. Locations1stDim  - Indices into 1st dimension
' 4. Locations2ndDim  - Indices into 2nd dimension
'
' RETURNED VALUE
' Returns ArrayToChange after making the desired substitutions
Private Function SetPartsUsing2ndDim(ArrayToChange As Variant, _
                                     ReplacementItems As Variant, _
                                     Locations1stDim As Variant, _
                                     Locations2ndDim As Variant) As Variant
    ' Set default return value in case of error
    Let SetPartsUsing2ndDim = Null

    If IsNull(ArrayToChange) Then Exit Function
    If IsNull(ReplacementItems) Then Exit Function
    If IsNull(Locations1stDim) Then Exit Function
    If IsNull(Locations2ndDim) Then Exit Function
                                           
    If NumberOfDimensions(ArrayToChange) = 1 Then
        Let SetPartsUsing2ndDim = _
            SetPartsUsing2ndDimIn1DArray(ArrayToChange, _
                                         ReplacementItems, _
                                         Locations1stDim, _
                                         Locations2ndDim)
    ElseIf NumberOfDimensions(ArrayToChange) = 2 Then
        Let SetPartsUsing2ndDim = _
            SetPartsUsing2ndDimIn2DArray(ArrayToChange, _
                                         ReplacementItems, _
                                         Locations1stDim, _
                                         Locations2ndDim)
    End If
End Function

' DESCRIPTION
' Function to replace specific parts of an array. It works on 1D.  The function returns Null when it fails
' due to parameter incompatibility. It can operate on a number of modes.  In particular, it can replace
' single locations in an array or sets of locations or spans.  The set of locations to replace must have
' the same shape as the items that replace.
'
' PARAMETERS
' 1. ArrayToChange    - Array to change
' 2. ReplacementItems - Values to replace existing ones in ArrayToChange
' 3. Locations1stDim  - Indices into 1st dimension
' 4. Locations2ndDim  - Indices into 2nd dimension
'
' This function can handle any of the following calling modalities acting along
' positions of a 1D array or rows of a 1D array:
'
' 1. SINGLE LOCATION IN 1D ARRAY OF 1D ARRAYS
'
'    ArrayToChange         <- Array(Array(a^1_1, a^1_2,...), Array(a^2_1, a^2_2,...), ...)
'    ReplacementItems      <- r
'    Locations1stDim       <- i
'    Locations2ndDim       <- j
'
'    Replace element the a^i_j  with r
'
' 2. ArrayToChange         <- Array(Array(a^1_1, a^1_2,...), Array(a^2_1, a^2_2,...), ...)
'    ReplacementItems      <- Array(Array(r^1_1, r^1_2,...,r^1_m), ..., Array(r^l_1, r^l_2,...,r^l_m))
'    Locations1stDim       <- Array(i_1,i_2,...,i_l)
'    Locations2ndDim       <- Array(j_1,j_2,...,j_m)
'
'    Replace elements a^(i_x)_(j_y) with r^x_y for each x=1,...,l and y=1,...,m
'
' 3. ArrayToChange         <- Array(Array(a^1_1, a^1_2,...), Array(a^2_1, a^2_2,...), ...)
'    ReplacementItems      <- [{r^1_1, r^1_2,...,r^1_m;...;r^l_1, r^l_2,...,r^l_m}]
'    Locations1stDim       <- Array(i_1,i_2,...,i_l)
'    Locations2ndDim       <- Array(j_1,j_2,...,j_m)
'
'    Replace elements a^(i_x)_(j_y) with r^x_y for each x=1,...,l and y=1,...,m
'
' 4. ArrayToChange         <- Array(Array(a^1_1, a^1_2,...), Array(a^2_1, a^2_2,...), ...)
'    ReplacementItems      <- Array(r^1, r^2,...,r^m)
'    Locations1stDim       <- i
'    Locations2ndDim       <- Array(j_1,j_2,...,j_m)
'
'    Replace elements a^i_(j_y) with r^y for each y=1,...,m
'
' 5. ArrayToChange         <- Array(Array(a^1_1, a^1_2,...), Array(a^2_1, a^2_2,...), ...)
'    ReplacementItems      <- Array(r^1, r^2,...,r^m)
'    Locations1stDim       <- Array(j_1,j_2,...,j_m)
'    Locations2ndDim       <- j
'
'    Replace elements a^(i_x)_j with r^x for each x=1,...,m
'
' 6. ArrayToChange         <- Array(Array(a^1_1, a^1_2,...), Array(a^2_1, a^2_2,...), ...)
'    ReplacementItems      <- Array(r^1, r^1,...,r^1)
'    Locations1stDim       <- i
'    Locations2ndDim       <- Span(...)
'
'    Replace elements a^i_j, where j comes from the Span(), with the given replacement elements
'
' 7. ArrayToChange         <- Array(Array(a^1_1, a^1_2,...), Array(a^2_1, a^2_2,...), ...)
'    ReplacementItems      <- Array(r^1, r^1,...,r^1)
'    Locations1stDim       <- Span(...)
'    Locations2ndDim       <- j
'
'    Replace elements a^i_j, where i comes from the Span(), with the given replacement elements
'
' 8. ArrayToChange         <- Array(Array(a^1_1, a^1_2,...), Array(a^2_1, a^2_2,...), ...)
'    ReplacementItems      <- Array(Array(r^1_1, r^1_2,...,r^1_m), ..., Array(r^l_1, r^l_2,...,r^l_m))
'    Locations1stDim       <- Span(...)
'    Locations2ndDim       <- Span(...)
'
'    Replace elements a^i_j with the index arrays from the two spans as in #3 above
'
' 7. ArrayToChange         <- Array(Array(a^1_1, a^1_2,...), Array(a^2_1, a^2_2,...), ...)
'    ReplacementItems      <- [{r^1_1, r^1_2,...,r^1_m;...;r^l_1, r^l_2,...,r^l_m}]
'    Locations1stDim       <- Span(...)
'    Locations2ndDim       <- Span(...)
'
'    Replace elements a^i_j with the index arrays from the two spans as in #3 above
'
' RETURNED VALUE
' Returns ArrayToChange after making the desired substitutions
Private Function SetPartsUsing2ndDimIn1DArray(ArrayToChange As Variant, _
                                              ReplacementItems As Variant, _
                                              Locations1stDim As Variant, _
                                              Locations2ndDim As Variant) As Variant
    Dim r As Long
    Dim c As Long
    Dim TheIndices As Variant
    
    ' Set default return value in case of error
    Let SetPartsUsing2ndDimIn1DArray = Null

    If IsNull(ArrayToChange) Then Exit Function
    If IsNull(ReplacementItems) Then Exit Function
    If IsNull(Locations1stDim) Then Exit Function
    If IsNull(Locations2ndDim) Then Exit Function
    
    ' Case 1 in function reference
    If IsNumeric(Locations1stDim) And IsNumeric(Locations2ndDim) Then
        If NumberOfDimensions(Part(ArrayToChange, Locations1stDim)) <> 1 Then Exit Function
    
        Let ArrayToChange(Locations1stDim)(Locations2ndDim) = ReplacementItems

        Let SetPartsUsing2ndDimIn1DArray = ArrayToChange
    ' Cases 2 and 3 in function reference
    ElseIf IsArray(Locations1stDim) And IsArray(Locations2ndDim) Then
        If NumberOfDimensions(ReplacementItems) = 1 Then
            For r = 1 To Length(Locations1stDim)
                For c = 1 To Length(Locations2ndDim)
                    Let ArrayToChange(Part(Locations1stDim, r))(Part(Locations2ndDim, c)) = _
                        Part(Part(ReplacementItems, r), c)
                Next
            Next
        ElseIf NumberOfDimensions(ReplacementItems) = 2 Then
            For r = 1 To Length(Locations1stDim)
                For c = 1 To Length(Locations2ndDim)
                    Let ArrayToChange(Part(Locations1stDim, r))(Part(Locations2ndDim, c)) = _
                        Part(ReplacementItems, r, c)
                Next
            Next
        End If

        Let SetPartsUsing2ndDimIn1DArray = ArrayToChange
    ' Case 4 in function reference
    ElseIf IsNumeric(Locations1stDim) And IsArray(Locations2ndDim) Then
        For c = 1 To Length(Locations2ndDim)
            Let ArrayToChange(Locations1stDim)(Part(Locations2ndDim, c)) = _
                Part(ReplacementItems, c)
        Next

        Let SetPartsUsing2ndDimIn1DArray = ArrayToChange
    ' Case 5 in function reference
    ElseIf IsArray(Locations1stDim) And IsNumeric(Locations2ndDim) Then
        For r = 1 To Length(Locations1stDim)
            Let ArrayToChange(Part(Locations1stDim, r))(Locations2ndDim) = _
                Part(ReplacementItems, r)
        Next

        Let SetPartsUsing2ndDimIn1DArray = ArrayToChange
    ' Case 6 in function reference
    ElseIf IsNumeric(Locations1stDim) And SpanQ(Locations2ndDim) Then
        Let SetPartsUsing2ndDimIn1DArray = _
            SetPartsUsing2ndDimIn1DArray(ArrayToChange, _
                                         ReplacementItems, _
                                         Locations1stDim, _
                                         CreateIndexSequenceFromSpan(Part(ArrayToChange, Locations1stDim), _
                                                                     Locations2ndDim))
    ' Case 7 in function reference
    ElseIf SpanQ(Locations1stDim) And IsNumeric(Locations2ndDim) Then
        Let SetPartsUsing2ndDimIn1DArray = _
            SetPartsUsing2ndDimIn1DArray(ArrayToChange, _
                                         ReplacementItems, _
                                         CreateIndexSequenceFromSpan(ArrayToChange, _
                                                                     Locations1stDim), _
                                         Locations2ndDim)
    ' Cases 8 and 9 in function reference
    ElseIf SpanQ(Locations1stDim) And _
           SpanQ(Locations2ndDim) Then
        Let SetPartsUsing2ndDimIn1DArray = _
            SetPartsUsing2ndDimIn1DArray(ArrayToChange, _
                                         ReplacementItems, _
                                         CreateIndexSequenceFromSpan(ArrayToChange, _
                                                                     Locations1stDim), _
                                         CreateIndexSequenceFromSpan(First(ArrayToChange), _
                                                                     Locations2ndDim))
    Else
        Let SetPartsUsing2ndDimIn1DArray = Null
    End If
End Function

' DESCRIPTION
' Function to replace specific parts of an array. It works on 1D.  The function returns Null when it fails
' due to parameter incompatibility. It can operate on a number of modes.  In particular, it can replace
' single locations in an array or sets of locations or spans.  The set of locations to replace must have
' the same shape as the items that replace.
'
' PARAMETERS
' 1. ArrayToChange    - Array to change
' 2. ReplacementItems - Values to replace existing ones in ArrayToChange
' 3. Locations1stDim  - Indices into 1st dimension
' 4. Locations2ndDim  - Indices into 2nd dimension
'
' This function can handle any of the following calling modalities acting along
' positions of a 1D array or rows of a 1D array:
'
' 1. ArrayToChange         <- [{a^1_1, a^1_2,...; a^2_1, a^2_2,...; ...}]
'    ReplacementItems      <- r
'    Locations1stDim       <- i
'    Locations2ndDim       <- j
'
'    Replace element the (i,j) element a^i_j  with r
'
' 2. ArrayToChange         <- [{a^1_1, a^1_2,...; a^2_1, a^2_2,...; ...}]
'    ReplacementItems      <- Array(Array(r^1_1, r^1_2,...,r^1_m), ..., Array(r^l_1, r^l_2,...,r^l_m))
'    Locations1stDim       <- Array(i_1,i_2,...,i_l)
'    Locations2ndDim       <- Array(j_1,j_2,...,j_m)
'
'    Replace elements a^(i_x)_(j_y) with r^x_y for each x=1,...,l and y=1,...,m
'
' 3. ArrayToChange         <- [{a^1_1, a^1_2,...; a^2_1, a^2_2,...; ...}]
'    ReplacementItems      <- [{r^1_1, r^1_2,...,r^1_m;...;r^l_1, r^l_2,...,r^l_m}]
'    Locations1stDim       <- Array(i_1,i_2,...,i_l)
'    Locations2ndDim       <- Array(j_1,j_2,...,j_m)
'
'    Replace elements a^(i_x)_(j_y) with r^x_y for each x=1,...,l and y=1,...,m
'
' 4. ArrayToChange         <- [{a^1_1, a^1_2,...; a^2_1, a^2_2,...; ...}]
'    ReplacementItems      <- Array(r^1, r^2,...,r^m)
'    Locations1stDim       <- i
'    Locations2ndDim       <- Array(j_1,j_2,...,j_m)
'
'    Replace elements a^i_(j_y) with r^y for each y=1,...,m
'
' 5. ArrayToChange         <- [{a^1_1, a^1_2,...; a^2_1, a^2_2,...; ...}]
'    ReplacementItems      <- Array(r^1, r^2,...,r^m)
'    Locations1stDim       <- Array(j_1,j_2,...,j_m)
'    Locations2ndDim       <- j
'
'    Replace elements a^(i_x)_j with r^x for each x=1,...,m
'
' 6. ArrayToChange         <- [{a^1_1, a^1_2,...; a^2_1, a^2_2,...; ...}]
'    ReplacementItems      <- Array(r^1, r^1,...,r^1)
'    Locations1stDim       <- i
'    Locations2ndDim       <- Span(...)
'
'    Replace elements a^i_j, where j comes from the Span(), with the given replacement elements
'
' 7. ArrayToChange         <- [{a^1_1, a^1_2,...; a^2_1, a^2_2,...; ...}]
'    ReplacementItems      <- Array(r^1, r^1,...,r^1)
'    Locations1stDim       <- Span(...)
'    Locations2ndDim       <- j
'
'    Replace elements a^i_j, where i comes from the Span(), with the given replacement elements
'
' 8. ArrayToChange         <- [{a^1_1, a^1_2,...; a^2_1, a^2_2,...; ...}]
'    ReplacementItems      <- Array(Array(r^1_1, r^1_2,...,r^1_m), ..., Array(r^l_1, r^l_2,...,r^l_m))
'    Locations1stDim       <- Span(...)
'    Locations2ndDim       <- Span(...)
'
'    Replace elements a^i_j with the index arrays from the two spans as in #3 above
'
' 7. ArrayToChange         <- [{a^1_1, a^1_2,...; a^2_1, a^2_2,...; ...}]
'    ReplacementItems      <- [{r^1_1, r^1_2,...,r^1_m;...;r^l_1, r^l_2,...,r^l_m}]
'    Locations1stDim       <- Span(...)
'    Locations2ndDim       <- Span(...)
'
'    Replace elements a^i_j with the index arrays from the two spans as in #3 above
'
' RETURNED VALUE
' Returns ArrayToChange after making the desired substitutions
Private Function SetPartsUsing2ndDimIn2DArray(ArrayToChange As Variant, _
                                              ReplacementItems As Variant, _
                                              Locations1stDim As Variant, _
                                              Locations2ndDim As Variant) As Variant
    Dim r As Long
    Dim c As Long
    
    ' Set default return value in case of error
    Let SetPartsUsing2ndDimIn2DArray = Null

    If IsNull(ArrayToChange) Then Exit Function
    If IsNull(ReplacementItems) Then Exit Function
    If IsNull(Locations1stDim) Then Exit Function
    If IsNull(Locations2ndDim) Then Exit Function
    
    ' Case 1 in function reference
    If IsNumeric(Locations1stDim) And IsNumeric(Locations2ndDim) Then
        Let ArrayToChange(Locations1stDim, Locations2ndDim) = ReplacementItems

        Let SetPartsUsing2ndDimIn2DArray = ArrayToChange
    ' Cases 2 & 3 in function reference
    ElseIf IsArray(Locations1stDim) And IsArray(Locations2ndDim) Then
        If NumberOfDimensions(ReplacementItems) = 1 Then
            For r = 1 To Length(Locations1stDim)
                For c = 1 To Length(Locations2ndDim)
                    Let ArrayToChange(Part(Locations1stDim, r), Part(Locations2ndDim, c)) = _
                        Part(Part(ReplacementItems, r), c)
                Next
            Next
        ElseIf NumberOfDimensions(ArrayToChange) = 2 Then
            For r = 1 To Length(Locations1stDim)
                For c = 1 To Length(Locations2ndDim)
                    Let ArrayToChange(Part(Locations1stDim, r), Part(Locations2ndDim, c)) = _
                        Part(ReplacementItems, r, c)
                Next
            Next
        End If
        
        Let SetPartsUsing2ndDimIn2DArray = ArrayToChange
    ' Case 4 in function reference
    ElseIf IsNumeric(Locations1stDim) And IsArray(Locations2ndDim) Then
        For c = 1 To Length(Locations2ndDim)
            Let ArrayToChange(Locations1stDim, Part(Locations2ndDim, c)) = _
                Part(ReplacementItems, c)
        Next
        
        Let SetPartsUsing2ndDimIn2DArray = ArrayToChange
    ' Case 5 in function reference
    ElseIf IsArray(Locations1stDim) And IsNumeric(Locations2ndDim) Then
        For r = 1 To Length(Locations1stDim)
            Let ArrayToChange(Part(Locations1stDim, r), Locations2ndDim) = _
                Part(ReplacementItems, r)
        Next
        
        Let SetPartsUsing2ndDimIn2DArray = ArrayToChange
    ' Case 6 in function reference
    ElseIf IsNumeric(Locations1stDim) And SpanQ(Locations2ndDim) Then
        Let SetPartsUsing2ndDimIn2DArray = _
            SetPartsUsing2ndDimIn2DArray(ArrayToChange, _
                                         ReplacementItems, _
                                         Locations1stDim, _
                                         CreateIndexSequenceFromSpan(ArrayToChange, _
                                                                     Locations2ndDim, _
                                                                     2))
    ' Case 7 in function reference
    ElseIf SpanQ(Locations1stDim) And IsNumeric(Locations2ndDim) Then
        Let SetPartsUsing2ndDimIn2DArray = _
            SetPartsUsing2ndDimIn2DArray(ArrayToChange, _
                                               ReplacementItems, _
                                               CreateIndexSequenceFromSpan(ArrayToChange, _
                                                                           Locations1stDim, _
                                                                           2), _
                                               Locations2ndDim)
    ' Case 8 in function reference
    ElseIf NumberOfDimensions(ArrayToChange) = 1 And SpanQ(Locations1stDim) And SpanQ(Locations2ndDim) Then
        Let SetPartsUsing2ndDimIn2DArray = _
            SetPartsUsing2ndDimIn2DArray(ArrayToChange, _
                                         ReplacementItems, _
                                         Locations1stDim, _
                                         CreateIndexSequenceFromSpan(First(ArrayToChange), _
                                                                     Locations2ndDim))
    ' Case 9 in function reference
    ElseIf NumberOfDimensions(ArrayToChange) = 2 And SpanQ(Locations1stDim) And SpanQ(Locations2ndDim) Then
        Let SetPartsUsing2ndDimIn2DArray = _
            SetPartsUsing2ndDimIn2DArray(ArrayToChange, _
                                         ReplacementItems, _
                                         CreateIndexSequenceFromSpan(ArrayToChange, _
                                                                     Locations1stDim), _
                                         CreateIndexSequenceFromSpan(First(ArrayToChange), _
                                                                     Locations2ndDim))
    Else
        Let SetPartsUsing2ndDimIn2DArray = Null
    End If
End Function

' DESCRIPTION
' This function returns the column specified by the given header of a "true" 2D array
'
' PARAMETERS
' 1. A2DArray     - "True" 2D array to reorder
' 2. ColumnHeader - a string representing a column header
'
' RETURNED VALUE
' The desired column, including its headers, as a 1D array
Public Function GetColumnByHeader(A2DArray As Variant, ColumnHeader As String) As Variant
    Dim var As Variant
    Dim ColIndex As Integer
    
    ' Exit with Null if A2DArray is not a "true" 2D array
    If NumberOfDimensions(A2DArray) <> 2 Then
        Let GetColumnByHeader = Null
        Exit Function
    End If
    
    ' Exit with Null if ColumnHeader is not a header in A2DArray
    If FreeQ(First(A2DArray), ColumnHeader) Then
        Let GetColumnByHeader = Null
        Exit Function
    End If
    
    ' Get column position of header in A2Array
    Let ColIndex = Application.Match(ColumnHeader, First(A2DArray), 0)
    
    ' Get desired column
    Let GetColumnByHeader = Part(A2DArray, Span(1, -1), ColIndex)
End Function

' DESCRIPTION
' This function rearranges the columns of a "true" 2D array according to
' the order of the headers in the 1D array of strings ColumnNamesInDesiredOrder.
'
' PARAMETERS
' 1. A2DArray                   - "True" 2D array to reorder
' 2. ColumnNamesInDesiredOrder - 1D array of strings that is a permutations of
'                                the header row in A2DArray
'
' RETURNED VALUE
' The rearranged 2D array
Public Function ReorderColumns(A2DArray As Variant, ColumnNamesInDesiredOrder As Variant) As Variant
    Dim var As Variant
    Dim ReturnArray As Variant
    Dim i As Integer
    
    ' Exit with Nothing if A2DArray is not a "true" 2D array
    If NumberOfDimensions(A2DArray) <> 2 Then
        Let ReorderColumns = Null
        Exit Function
    End If
    
    ' Exit with Nothing if ColumnNamesInDesiredOrder is not a 1D array
    If NumberOfDimensions(ColumnNamesInDesiredOrder) <> 1 Then
        Let ReorderColumns = Null
        Exit Function
    End If
    
    ' Exit if number A2DArray and ColumnNamesInDesiredOrder dimensions incompatible
    If NumberOfColumns(A2DArray) <> Length(ColumnNamesInDesiredOrder) Then
        Let ReorderColumns = Null
        Exit Function
    End If
    
    ' Exit if ColumnNamesInDesiredOrder is not an array of strings
    For Each var In ColumnNamesInDesiredOrder
        If Not StringQ(var) Then
            Let ReorderColumns = Null
            Exit Function
        End If
    Next
    
    ' Exit if the two sets of headers are not the same
    If Not EmptyArrayQ(ComplementOfSets(First(A2DArray), ColumnNamesInDesiredOrder)) Then
        Let ReorderColumns = Null
        Exit Function
    End If
    
    ' Pre-allocate ReturnArray
    Let ReturnArray = ConstantArray(Empty, NumberOfColumns(A2DArray))
    
    For i = 1 To NumberOfColumns(A2DArray)
        Let ReturnArray(i) = GetColumnByHeader(A2DArray, CStr(ColumnNamesInDesiredOrder(i)))
    Next i
    
    ' Return the transpose array
    Let ReorderColumns = TransposeMatrix(Pack2DArray(ReturnArray))
End Function

' DESCRIPTION
' The purpose of this function is to add additional quotes around the
' strings in a variant array so they may be printed correctly. Without
' it, the strings in the array print unquoted.
'
' PARAMETERS
' 1. AnArray - An array
'
' RETURNED VALUE
' The properly quoted array for printing purposes
Public Function QuoteStringsInArray(ByVal AnArray As Variant) As Variant
    Dim i As Long
    
    For i = LBound(AnArray) To UBound(AnArray)
        If StringQ(AnArray(i)) Then Let AnArray(i) = """" & AnArray(i) & """"
    Next
    
    Let QuoteStringsInArray = AnArray
End Function

' DESCRIPTION
' Completes a times series by inserting rows with missing dates. Missings gets
' get the values from the last available numerical value in the same column.
'
' PARAMETERS
' 1. TimeSeries2DArray - A 2D array as follows:
'    Date   Header1     Header2
'    D1     a1          b1
'    D2     a2          b2
'    .      .           .
'    .      .           .
'    .      .           .
'
' For example, the table
'
' Date          IBM US Equity   7003 JP Equity
' 31/12/2014    160.44          #N/A N.A.
' 02/01/2015    162.06          #N/A N.A.
' 05/01/2015    159.51          2150
'
' is transformed into
'
' Date          IBM US Equity   7003 JP Equity
' 31/12/2014    160.44
' 01/01/2015    160.44
' 02/01/2015    162.06
' 03/01/2015    162.06
' 04/01/2015    162.06
' 05/01/2015    159.51          2150
'
' RETURNED VALUE
' The completed timeseries.
Public Function CompleteTimeSeries(TimeSeries2DArray As Variant) As Variant
    Dim LatestValues() As Variant
    Dim n As Long ' to keep track of row in ReturnArray
    Dim tsdc As Long ' to loop over rows in TimeSeries2DArray
    Dim CurrentDate As Long ' The current date in TimeSeries2DArray
    Dim IntraDate As Long
    Dim ReturnArray() As Variant
    Dim i As Integer
    Dim NumCols As Integer
    
    If Length(TimeSeries2DArray) < 2 Then
        Let CompleteTimeSeries = TimeSeries2DArray
        Exit Function
    End If
    
    ' Pre-allocated ReturnArray and LatestData
    Let NumCols = NumberOfColumns(TimeSeries2DArray)
    ReDim ReturnArray(1 To Part(TimeSeries2DArray, -1, 1) - Part(TimeSeries2DArray, 2, 1) + 2, 1 To NumCols)
    ReDim LatestData(1 To NumCols)
    
    ' Set the headers row of the return array
    Let ReturnArray = SetPart(ReturnArray, First(TimeSeries2DArray), 1)
    
    Let tsdc = 2
    Let n = 2
    While tsdc <= Length(TimeSeries2DArray)
        ' Replace latest available data with current values if numeric
        For i = 2 To NumberOfColumns(TimeSeries2DArray)
            If NumberQ(Part(TimeSeries2DArray, tsdc, i)) Then
                Let LatestData(i) = Part(TimeSeries2DArray, tsdc, i)
            End If
        Next
        
        ' Extract current date from input time series and set IntraDate to it
        Let CurrentDate = Part(TimeSeries2DArray, tsdc, 1)
        Let IntraDate = CurrentDate
        
        ' Continue adding dates to the return array until hitting next input date
        While IntraDate < Part(TimeSeries2DArray, tsdc + 1, 1)
            Let ReturnArray = SetPart(ReturnArray, LatestData, n)
            Let ReturnArray(n, 1) = IntraDate
            
            Let IntraDate = IntraDate + 1
            Let n = n + 1
        Wend
        
        ' Move to next date in TimeSeries2DArray
        Let tsdc = tsdc + 1
    Wend
    
    Let ReturnArray = SetPart(ReturnArray, LatestData, -1)
    Let ReturnArray = SetPart(ReturnArray, CurrentDate, Length(ReturnArray), 1)
    
    ' Return the completed time series
    Let CompleteTimeSeries = ReturnArray
End Function

' DESCRIPTION
' This function restores the formulas stored in the comments of the first row of a range.
' Columns with no comments are ignored.  The formulas as taken from the comments in the
' first row of the current region. However, they are applied only to the region selected.
'
' It is important that the formulas properly anchor row and column reference so they make
' sense when copied down from the second row to the other rows of the current region.
'
' PARAMETERS
' 1. ARange - A reference to a range
'
' RETURNED VALUE
' No value is returned
Public Sub RestoreRangeFormulas(ARange As Range)
    Dim c As Integer
    Dim HeadersRange As Range
    Dim TheCurrentRegion As Range
     
    ' Set reference to range's current region
    Set TheCurrentRegion = ARange.CurrentRegion
  
    ' Set reference to header row
    Set HeadersRange = TheCurrentRegion.Rows(1)
    
    ' Proceed if the current region has more than one row. Exit otherwise
    If ARange.CurrentRegion.Rows.Count > 1 Then
        ' Loop over the columns, insert the formula stored in the comments
        ' in the cells of the second row (e.g. first row below the headers)
        For c = 1 To TheCurrentRegion.Columns.Count
            ' Proceed to insert the formula if the comment is not null
            If Not (HeadersRange.Cells(1, c).Comment Is Nothing) Then
                ' Insert in the first cell below the header cell the formula stored in the
                ' comment
                Let TheCurrentRegion.Cells(1, c).Formula = HeadersRange.Cells(1, c).Comment.Text
                
                ' Copy the formula to the entire column, excluding the header row
                Call TheCurrentRegion(1, c).Copy
                Call TheCurrentRegion.Columns(c).Offset(1, 0).Resize(TheCurrentRegion.Rows.Count - 1, 1).PasteSpecial(xlFormulas)
            End If
        Next
    End If
End Sub

' DESCRIPTION
' Stores in the comments of the headers row the formulas in the first listrow.
' Columns without formulas are ignored.
'
' Note that this function wipes all pre-exisiting comments in the headers row.
'
' PARAMETERS
' None
'
' RETURNED VALUE
' No value is returned
Public Sub SaveRangeFormulas(ARange As Range)
    Dim c As Integer
    Dim HeadersRange As Range
    Dim TheCurrentRegion As Range

    ' Set reference to range's current region
    Set TheCurrentRegion = ARange.CurrentRegion

    ' Set reference to header row
    Set HeadersRange = TheCurrentRegion.Rows(1)
    
    ' Exit if TheCurrentRegion has only one row
    If TheCurrentRegion.Rows.Count < 2 Then Exit Sub

    ' Store formulas
    For c = 1 To HeadersRange.Columns.Count
        ' Store the formulas of every column
        If TheCurrentRegion(2, c).HasFormula Then
            ' Clear any previously stored comments
            If Not (HeadersRange.Cells(1, c).Comment Is Nothing) Then
                Call HeadersRange(1, c).ClearComments
            End If

            If Left(Trim(TheCurrentRegion(2, c).Formula), 1) = "=" Then
                Call HeadersRange.Cells(1, c).AddComment(Trim(TheCurrentRegion(2, c).Formula))
            End If
        End If
    Next
End Sub

' DESCRIPTION
' This function restores the formulas stored in the comments of the first row of a range.
' Columns with no comments are ignored.  The formulas as taken from the comments in the
' first row of the current region. However, they are applied only to the region selected.
'
' It is important that the formulas properly anchor row and column reference so they make
' sense when copied down from the second row to the other rows of the current region.
'
' PARAMETERS
' 1. ARange - A reference to a range
'
' RETURNED VALUE
' No value is returned
Public Sub RestoreCurrentRegionFormulas()
    Call RestoreRangeFormulas(Selection(1, 1))
End Sub

' DESCRIPTION
' Stores in the comments of the headers row the formulas in the first listrow.
' Columns without formulas are ignored.
'
' Note that this function wipes all pre-exisiting comments in the headers row.
'
' PARAMETERS
' None
'
' RETURNED VALUE
' No value is returned
Public Sub SaveCurrentRegionFormulas()
    Call SaveRangeFormulas(Selection(1, 1))
End Sub

