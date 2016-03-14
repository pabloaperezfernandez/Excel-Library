Attribute VB_Name = "Predicates"
Option Explicit
Option Base 1

' DESCRIPTION
' Boolean function returning True if its argument is an array that has been dimensioned.
' Returns False otherwise.  In other words, it returns False if its arg is neither an
' an array nor dimensioned.
'
' PARAMETERS
' 1. arg - any value or object reference
'
' RETURNED VALUE
' True when arg is a dimensioned array. False otherwise.
Public Function DimensionedQ(arg As Variant) As Boolean
    Dim i As Long
    
    On Error Resume Next
    
    ' Exit with False
    If Not IsArray(arg) Then
        Let DimensionedQ = False
        Exit Function
    End If
    
    ' If arg has not been dimensioned, this following line will raise an error.
    ' Due to On Error Resume Next, the code will resume in the next line, which
    ' will then check if an error has been raised.
    Let i = UBound(arg, 1)
    Let DimensionedQ = Err.Number = 0
End Function

' DESCRIPTION
' Boolean function returning True if its argument is an Empty 1D array. A prerequisite is that
' the array be dimensioned and 1D.  Returns False otherwise.
'
' PARAMETERS
' 1. arg - any value or object reference
'
' RETURNED VALUE
' True when arg is a dimensioned, 1D, empty array.  Returns False otherwise
Public Function EmptyArrayQ(anArray As Variant) As Boolean
    ' Exit with False if AnArray is not an array or has not been dimensioned
    If Not DimensionedQ(anArray) Then
        Let EmptyArrayQ = False
    ' Return True if we have an array with lower Ubound than LBound
    Else
        Let EmptyArrayQ = UBound(anArray, 1) - LBound(anArray, 1) < 0
    End If
End Function

' DESCRIPTION
' Boolean function returning True if its argument is one of the following:
'
' 1. Number
' 2. String
' 3. Date
' 4. Boolean
' 5. Error
' 6. Worksheet
' 7. Workbook
' 8. ListObject
' 9. Null
' 10. Empty
'
' It returns False otherwise
'
' PARAMETERS
' 1. arg - any value or object reference
'
' RETURNED VALUE
' True when arg has one of the types detailed above. Returns False otherwise
Public Function AtomicQ(arg As Variant) As Boolean
    Let AtomicQ = False
    
    If IsArray(arg) Then Exit Function
    
    If Not (TypeName(arg) = TypeName("a") Or _
            TypeName(arg) = TypeName(True) Or _
            TypeName(arg) = TypeName(CVErr(1)) Or _
            TypeName(arg) = TypeName(TempComputation) Or _
            TypeName(arg) = TypeName(ThisWorkbook) Or _
            TypeName(arg) = "ListObject" Or _
            IsNumeric(arg) Or _
            IsDate(arg) Or _
            IsNull(arg) Or _
            IsEmpty(arg)) Then
        Exit Function
    End If
    
    Let AtomicQ = True
End Function

' DESCRIPTION
' Boolean function returning True if its argument is array all of whose elements satisfy Predicates.AtomicQ
' The function returns False if arg fails Predicates.DimensionedQ or satisfying Predicates.EmptyArrayQ.
' Returns False otherwise
'
' PARAMETERS
' 1. arg - any value or object reference
'
' RETURNED VALUE
' True when arg is a dimensioned, non-empty array all of whose elements satisfy Predicates.AtomicQ.
' Returns False otherwise
Public Function AtomicArrayQ(arg As Variant) As Boolean
    Let AtomicArrayQ = EveryQ(arg, "AtomicQ")
End Function

' DESCRIPTION
' Boolean function returning True if its argument is a number.  Returns False otherwise
'
' PARAMETERS
' 1. arg - any value or object reference
'
' RETURNED VALUE
' True when arg is byte, integer, long, single, double, or currency. LongLong
' is available on 64-bit systems exclusively, and Decimal is not found in
' my systems.  Returns False otherwise.
Public Function NumberQ(arg As Variant) As Boolean
    Dim AByte As Byte
    Dim anInteger As Integer
    Dim ALong As Long
    Dim ASingle As Single
    Dim aDouble As Double
    Dim ACurrency As Currency

    Select Case TypeName(arg)
        Case TypeName(AByte)
            Let NumberQ = True
        Case TypeName(anInteger)
            Let NumberQ = True
        Case TypeName(ALong)
            Let NumberQ = True
        Case TypeName(ASingle)
            Let NumberQ = True
        Case TypeName(aDouble)
            Let NumberQ = True
        Case TypeName(ACurrency)
            Let NumberQ = True
        Case Else
            Let NumberQ = False
    End Select
End Function

' DESCRIPTION
' Boolean function returning True if its argument is a dimensioned, non-empty array all of whose elements satisfy
' Predicates.NumberQ.  Returns False otherwise.
'
' PARAMETERS
' 1. arg - any value or object reference
'
' RETURNED VALUE
' True or False depending on whether arg is dimensioned, non-empty and all its elements satisfy Predicates.NumberQ
Public Function NumberArrayQ(arg As Variant) As Boolean
    Let NumberArrayQ = EveryQ(arg, "NumberQ")
End Function

' DESCRIPTION
' Boolean function returning True if its argument is a any whole number type. Returns False otherwise.
'
' PARAMETERS
' 1. arg - any value or object reference
'
' RETURNED VALUE
' True or False depending on whether or not its argument is a whole number.
Public Function WholeNumberQ(arg As Variant) As Boolean
    Dim AByte As Byte
    Dim anInteger As Integer
    Dim ALong As Long

    Select Case TypeName(arg)
        Case TypeName(AByte)
            Let WholeNumberQ = True
        Case TypeName(anInteger)
            Let WholeNumberQ = True
        Case TypeName(ALong)
            Let WholeNumberQ = True
        Case Else
            Let WholeNumberQ = False
    End Select
End Function

' DESCRIPTION
' Boolean function returning True if its argument is a dimensioned, non-empty array all of whose
' elements satisfy Predicates.WholeNumberQ.  Returns False otherwise.
'
' PARAMETERS
' 1. arg - any value or object reference
'
' RETURNED VALUE
' True or False depending on whether arg is dimensioned, non-empty and all its elements satisfy
' Predicates.WholeNumberQ
Public Function WholeNumberArrayQ(arg As Variant) As Boolean
    Let WholeNumberArrayQ = EveryQ(arg, "WholeNumberQ")
End Function

' DESCRIPTION
' Boolean function returning True if its argument is a any whole number type and is positive.
' Returns False otherwise.
'
' PARAMETERS
' 1. arg - any value or object reference
'
' RETURNED VALUE
' True or False depending on whether or not its argument is a positive whole number.
Public Function PositiveWholeNumberQ(arg As Variant) As Boolean
    If WholeNumberQ(arg) Then
        Let PositiveWholeNumberQ = arg > 0
    Else
        Let PositiveWholeNumberQ = False
    End If
End Function

' DESCRIPTION
' Boolean function returning True if its argument is a dimensioned, non-empty array all of
' whose elements satisfyPredicates.PositiveWholeNumberQ.  Returns False otherwise
'
' PARAMETERS
' 1. arg - any value or object reference
'
' RETURNED VALUE
' True or False depending on whether or not its argument is dimensioned, non-empty and all
' its elements satisfy Predicates.PositiveWholeNumberQ
Public Function PositiveWholeNumberArrayQ(arg As Variant) As Boolean
    Let PositiveWholeNumberArrayQ = EveryQ(arg, "PositiveWholeNumberQ")
End Function

' DESCRIPTION
' Boolean function returning True if its argument is a negative whole number. Returns False
' otherwise.
'
' PARAMETERS
' 1. arg - any value or object reference
'
' RETURNED VALUE
' True or False depending on whether or not its argument is a negative whole number.
Public Function NegativeWholeNumberQ(arg As Variant) As Boolean
    If WholeNumberQ(arg) Then
        Let NegativeWholeNumberQ = arg < 0
    Else
        Let NegativeWholeNumberQ = False
    End If
End Function

' DESCRIPTION
' Boolean function returning True if its argument is a dimensioned, non-empty array all of whose
' elements satisfy Predicates.NegativeWholeNumberQ.  Returns False otherwise
'
' PARAMETERS
' 1. arg - any value or object reference
'
' RETURNED VALUE
' True or False depending on whether arg is dimensioned, non-empty and all its elements satisfy
' Predicates.NegativeWholeNumberQ
Public Function NegativeWholeNumberArrayQ(arg As Variant) As Boolean
    Let NegativeWholeNumberArrayQ = EveryQ(arg, "NegativeWholeNumberQ")
End Function

' DESCRIPTION
' Boolean function returning True if its argument is a non-negative whole number. Returns False
' otherwise.
'
' PARAMETERS
' 1. arg - any value or object reference
'
' RETURNED VALUE
' True or False depending on whether or not its argument is a non-negative whole number.
Public Function NonNegativeWholeNumberQ(arg As Variant) As Boolean
    If WholeNumberQ(arg) Then
        Let NonNegativeWholeNumberQ = arg >= 0
    Else
        Let NonNegativeWholeNumberQ = False
    End If
End Function

' DESCRIPTION
' Boolean function returning True if its argument is a dimensioned, non-empty array all of
' whose elements satisfy Predicates.NonNegativeWholeNumberQ.  Returns False otherwise
'
' PARAMETERS
' 1. arg - any value or object reference
'
' RETURNED VALUE
' True or False depending on whether arg is dimensioned, non-empty and all its elements satisfy
' Predicates.NonNegativeWholeNumberQ
Public Function NonNegativeWholeNumberArrayQ(arg As Variant) As Boolean
    Let NonNegativeWholeNumberArrayQ = EveryQ(arg, "NonNegativeWholeNumberQ")
End Function

' DESCRIPTION
' Boolean function returning True if its argument is a string. Returns False otherwise.
'
' PARAMETERS
' 1. arg - any value or object reference
'
' RETURNED VALUE
' True or False depending on whether or not its argument is a string
Public Function StringQ(AnArg As Variant) As Boolean
    Let StringQ = TypeName(AnArg) = TypeName("a")
End Function

' DESCRIPTION
' Boolean function returning True if its argument is a dimensioned, non-empty array all of whose
' elements satisfy Predicates.StringQ.  Returns False otherwise.
'
' PARAMETERS
' 1. arg - any value or object reference
'
' RETURNED VALUE
' True or False depending on whether arg is dimensioned, non-empty and all its elements satisfy
' Predicates.StringQ
Public Function StringArrayQ(anArray As Variant) As Boolean
    Let StringArrayQ = EveryQ(anArray, "StringQ")
End Function

' DESCRIPTION
' Boolean function returning True if its argument is a dictionary. Returns False otherwise.
'
' PARAMETERS
' 1. arg - any value or object reference
'
' RETURNED VALUE
' True or False depending on whether or not its argument is a dictionary
Public Function DictionaryQ(vValue As Variant) As Boolean
    Dim obj As New Dictionary
    
    Let DictionaryQ = TypeName(vValue) = TypeName(obj)
End Function

' DESCRIPTION
' Boolean function returning True if its argument is a dimensioned, non-empty array all of whose
' elements satisfy Predicates.DictionaryQ.  Returns False otherwise.
'
' PARAMETERS
' 1. arg - any value or object reference
'
' RETURNED VALUE
' True or False depending on whether arg is dimensioned, non-empty and all its elements satisfy
' Predicates.DictionaryQ
Public Function DictionaryArrayQ(anArray As Variant) As Boolean
    Let DictionaryArrayQ = EveryQ(anArray, "DictionaryQ")
End Function

' DESCRIPTION
' Boolean function returning True if its argument is a string or a whole number. Returns False otherwise.
'
' PARAMETERS
' 1. arg - any value or object reference
'
' RETURNED VALUE
' True or False depending on whether or not its argument string or a whole number.
Public Function WholeNumberOrStringQ(arg As Variant) As Boolean
    Let WholeNumberOrStringQ = WholeNumberQ(arg) Or StringQ(arg)
End Function

' DESCRIPTION
' Boolean function returning True if its argument is a dimensioned array all of whose elements satisfy
' Predicates.WholeNumberOrStringQ.  Returns False otherwise.
'
' PARAMETERS
' 1. arg - any value or object reference
'
' RETURNED VALUE
' True or False depending on whether arg is dimensioned, non-empty and all its elements satisfy
' Predicates.WholeNumberOrStringQ
Public Function WholeNumberOrStringArrayQ(anArray As Variant) As Boolean
    Let WholeNumberOrStringArrayQ = EveryQ(anArray, "WholeNumberOrStringQ")
End Function

' DESCRIPTION
' Boolean function returning True if its argument is a string or a number. Returns False otherwise.
'
' PARAMETERS
' 1. arg - any value or object reference
'
' RETURNED VALUE
' True or False depending on whether or not its argument string or a number.
Public Function NumberOrStringQ(arg As Variant) As Boolean
    Let NumberOrStringQ = NumberQ(arg) Or StringQ(arg)
End Function

' DESCRIPTION
' Boolean function returning True if its argument is a dimensioned array all of whose elements
' satisfy Predicates.NumberOrStringQ.  Returns False otherwise.
'
' PARAMETERS
' 1. arg - any value or object reference
'
' RETURNED VALUE
' True or False depending on whether arg is dimensioned, non-empty and all its elements satisfy
' Predicates.NumberOrStringQ
Public Function NumberOrStringArrayQ(anArray As Variant) As Boolean
    Let NumberOrStringArrayQ = EveryQ(anArray, "NumberOrStringQ")
End Function

' DESCRIPTION
' Boolean function returning True if all of the elements in AnArray satisfy the predicate whose
' name is PredicateName.
'
' PARAMETERS
' 1. arg - any value or object reference
'
' RETURNED VALUE
' True or False depending on whether arg is dimensioned and all its elements satisfy the
' predicate with name PredicateName
Public Function EveryQ(anArray As Variant, PredicateName As String) As Boolean
    Dim var As Variant

    Let EveryQ = True

    If Not DimensionedQ(anArray) Then
        Let EveryQ = False
        Exit Function
    End If
    
    For Each var In anArray
        If Not Application.Run(ThisWorkbook.Name & "!" & PredicateName, var) Then
            Let EveryQ = False
            Exit Function
        End If
    Next
End Function

' DESCRIPTION
' Boolean function returning True if at least one of the elements in AnArray satisfy the predcate
' whose name is PredicateName.
'
' PARAMETERS
' 1. arg - any value or object reference
'
' RETURNED VALUE
' True or False depending on whether arg is dimensioned and at least one of its elements satisfies
' the predicate with name PredicateName
Public Function SomeQ(anArray As Variant, PredicateName As String) As Boolean
    Dim var As Variant

    Let SomeQ = False

    If Not DimensionedQ(anArray) Then
        Let SomeQ = False
        Exit Function
    End If
    
    For Each var In anArray
        If Application.Run(ThisWorkbook.Name & "!" & PredicateName, var) Then
            Let SomeQ = True
            Exit Function
        End If
    Next
End Function

' DESCRIPTION
' Boolean function returns True if the given directory path or file exists.  Returns False
' otherwise.  You may pass a directory path or a full filename.
'
' PARAMETERS
' 1. ThePath - A valid path
'
' RETURNED VALUE
' True or False depending on whether the given path exists
Function FileExistsQ(ThePath As String) As Boolean
    If Not dir(ThePath, vbDirectory) = vbNullString Then
        FileExistsQ = True
    Else
        Let FileExistsQ = False
    End If
End Function

' DESCRIPTION
' Boolean function returns True if there is a listobject with the given name in the given
' worksheet.  Returns False otherwise.
'
' PARAMETERS
' 1. WorkSheetReference - A worksheet reference
' 2. ListObjectName - A string representing the name of the listobject in question
'
' RETURNED VALUE
' True or False depending on whether the a listobject with the given name exists on the
' given worksheet
Public Function ListObjectExistsQ(WorkSheetReference As Worksheet, ListObjectName As String) As Boolean
    Dim AName As String
    
    Let ListObjectExistsQ = True
    
    On Error GoTo ErrorHandler
    
    Let AName = WorkSheetReference.ListObjects(ListObjectName).Name
    
    Exit Function
    
ErrorHandler:
    Let ListObjectExistsQ = False
End Function

' DESCRIPTION
' Boolean function returns True if Returns True is its argument is 2D matrix with numeric entries.
'
' PARAMETERS
' 1. arg - Any Excel value or reference
'
' RETURNED VALUE
' Returns True or False depending on whether or not its argument can be considered a numerical matrix.
Public Function MatrixQ(arg As Variant) As Boolean
    Dim var As Variant
    
    Let MatrixQ = False
    
    ' Not necessary to test for DimensionedQ since NumberOfDimensions returns 0 for none arrays
    If NumberOfDimensions(arg) <> 2 Then Exit Function

    For Each var In arg
        If Not NumberQ(var) Then Exit Function
    Next
    
    Let MatrixQ = True
End Function

' DESCRIPTION
' Boolean function returns True if Returns True is its argument is 2D matrix with numeric entries.
'
' PARAMETERS
' 1. arg - Any Excel value or reference
'
' RETURNED VALUE
' Returns True or False depending on whether or not its argument can be considered a numerical matrix.
Public Function BooleanQ(arg As Variant) As Boolean
    Let BooleanQ = TypeName(arg) = TypeName(True)
End Function

' DESCRIPTION
' Boolean function returning True if its argument is a dimensioned array all of whose elements
' satisfy Predicates.BooleanQ.  Returns False otherwise.
'
' PARAMETERS
' 1. arg - any value or object reference
'
' RETURNED VALUE
' True or False depending on whether arg is dimensioned, non-empty and all its elements satisfy
' Predicates.BooleanQ
Public Function BooleanArrayQ(anArray As Variant) As Boolean
    Let BooleanArrayQ = EveryQ(anArray, "BooleanQ")
End Function

' DESCRIPTION
' Boolean function returns True if Returns True is its argument is 2D matrix with Atomic entries.
'
' PARAMETERS
' 1. arg - Any Excel value or reference
'
' RETURNED VALUE
' Returns True or False depending on whether or not its argument can be considered a table.
Public Function AtomicTableQ(arg As Variant) As Boolean
    Dim var As Variant

    Let AtomicTableQ = False
    
    If NumberOfDimensions(arg) <> 2 Then Exit Function

    For Each var In arg
        If Not AtomicQ(var) Then Exit Function
    Next var
    
    Let AtomicTableQ = True
End Function

' DESCRIPTION
' Boolean function returning True if its argument is printable (e.g. numeric, string, date, Boolean,
' Empty or Null)
'
' PARAMETERS
' 1. arg - Any Excel value or reference
'
' RETURNED VALUE
' Returns True or False depending on whether or not its argument can be considered a printable
Public Function PrintableQ(arg As Variant) As Boolean
    Let PrintableQ = NumberOrStringQ(arg) Or BooleanQ(arg) Or IsDate(arg) Or IsEmpty(arg) Or IsNull(arg)
End Function

' DESCRIPTION
' Boolean function returns True if Returns True is its argument is either an empty array or a 1D all of
' whose elements satisfy PrintableQ
'
' PARAMETERS
' 1. AnArray - Any Excel value or reference
'
' RETURNED VALUE
' Returns True or False depending on whether or not its arguments is a printable array
Public Function PrintableArrayQ(anArray As Variant) As Boolean
    Let PrintableArrayQ = EveryQ(anArray, "PrintableQ")
End Function

' DESCRIPTION
' Boolean function returns True if Returns True is its argument is 2D matrix with printable (e.g.
' numeric, string, date, True, False, Empty or Null) entries.
'
' PARAMETERS
' 1. arg - Any Excel value or reference
'
' RETURNED VALUE
' Returns True or False depending on whether or not its argument can be considered a printable table.
Public Function PrintableTableQ(arg As Variant) As Boolean
    Dim var As Variant

    Let PrintableTableQ = False
    
    If NumberOfDimensions(arg) <> 2 Then Exit Function

    For Each var In arg
        If Not PrintableQ(var) Then Exit Function
    Next var
    
    Let PrintableTableQ = True
End Function

' DESCRIPTION
' Boolean function returning True if its argument is a row or column vector as defined by
' ColumVectorQ and RowVectorQ
'
' PARAMETERS
' 1. arg - Any Excel value or reference
'
' RETURNED VALUE
' Returns True or False depending on whether or not its argument can be considered a vector.
Function VectorQ(arg As Variant) As Boolean
    Let VectorQ = RowVectorQ(arg) Or ColumnVectorQ(arg)
End Function

' DESCRIPTION
' Boolean function returning True if its argument is a column vector (e.g. n x 1 2D array with
' numeric entries exclusively) comprised of numeric, atomic expressions exclusively.
' Returns True for the empty 1D array
'
' PARAMETERS
' 1. arg - Any Excel value or reference
'
' RETURNED VALUE
' Returns True or False depending on whether or not its argument is a column vector.
Public Function ColumnVectorQ(arg As Variant) As Boolean
    Dim var As Variant
    
    Let ColumnVectorQ = False
    
    ' Returning True for an empty array is necesary recursion to work properly on column vectors
    If EmptyArrayQ(arg) Then
        Let ColumnVectorQ = True
        Exit Function
    End If
    
    If NumberOfDimensions(arg) <> 2 Then Exit Function
    
    If LBound(arg, 2) <> UBound(arg, 2) Then Exit Function
    
    For Each var In arg
        If Not NumberQ(var) Then Exit Function
    Next

    Let ColumnVectorQ = True
End Function

' DESCRIPTION
' Boolean function returning True if its argument is a row vector (e.g. a 1D vector comprised
' exclusively of numeric entries).  Returns True for the empty 1D array
'
' PARAMETERS
' 1. arg - Any Excel value or reference
'
' RETURNED VALUE
' Returns True or False depending on whether or not its argument is a row vector.
Public Function RowVectorQ(arg As Variant) As Boolean
    Dim var As Variant
    
    Let RowVectorQ = False
    
    If Not DimensionedQ(arg) Then Exit Function
    
    ' Returning True for an empty array is necesary recursion to work properly on row vectors
    If EmptyArrayQ(arg) Then
        Let RowVectorQ = True
        Exit Function
    End If
    
    If NumberOfDimensions(arg) <> 1 Then Exit Function
    
    If Not NumberArrayQ(arg) Then Exit Function
    
    Let RowVectorQ = True
End Function

' DESCRIPTION
' Boolean function returning True if its argument is a column array (e.g. n x 1 2D array with
' atomic entries exclusively) comprised of numeric, atomic expressions exclusively.
' Returns True for the empty 1D array
'
' PARAMETERS
' 1. arg - Any Excel value or reference
'
' RETURNED VALUE
' Returns True or False depending on whether or not its argument is a column vector.
Public Function ColumnArrayQ(arg As Variant) As Boolean
    Dim var As Variant
    
    Let ColumnArrayQ = False
    
    ' Returning True for an empty array is necesary recursion to work properly on column vectors
    If EmptyArrayQ(arg) Then
        Let ColumnArrayQ = True
        Exit Function
    End If
    
    If NumberOfDimensions(arg) <> 2 Then Exit Function
    
    If LBound(arg, 2) <> UBound(arg, 2) Then Exit Function
    
    For Each var In arg
        If Not AtomicQ(var) Then Exit Function
    Next

    Let ColumnArrayQ = True
End Function

' DESCRIPTION
' Alias for AtomicArrayQ. Included to preserve parallel structure (e.g. having both ColumnArrayQ
' and RowArrayQ just like we have RowVectorQ and ColumnVectorQ).
'
' PARAMETERS
' 1. arg - Any Excel value or reference
'
' RETURNED VALUE
' Returns True or False depending on whether or not its argument is a row vector.
Public Function RowArrayQ(arg As Variant) As Boolean
    Let RowArrayQ = AtomicArrayQ(arg)
End Function

' DESCRIPTION
' Boolean function returning True if its argument if either RowArrayQ or ColumnArrayQ return True
'
' PARAMETERS
' 1. arg - Any Excel value or reference
'
' RETURNED VALUE
' Returns True or False depending on whether or not its argument can be considered a row or column
' array.
Function RowOrColumnArrayQ(arg As Variant) As Boolean
    Let RowOrColumnArrayQ = RowArrayQ(arg) Or ColumnArrayQ(arg)
End Function

' DESCRIPTION
' Boolean function returning True if its argument could be interpreseted as a 1D array of atomic
' elements. This means that this function returns True for each of the following examples:
'
' 1. Array(1,2,3)
' 2. Array(Array(1,2,3))
' 3. Array(Array(Array(1,2,3)))),
' 4. [{1,2,3}], [{Array(1,2,3)}], etc. will evaluate to True.
'
' PARAMETERS
' 1. arg - Any Excel value or reference
'
' RETURNED VALUE
' Returns True or False depending on whether or not its argument could be interpreted as a row
' vector.
Public Function InterpretableAsRowArrayQ(a As Variant) As Boolean
    Dim nd As Integer
    Dim i As Long
    
    Let InterpretableAsRowArrayQ = False
    Let nd = NumberOfDimensions(a)
    
    ' If a has more than two or fewer than one dimension, then exit with False.
    If nd > 2 Or nd < 1 Then
        Exit Function
    End If
    
    ' Process arg is it has one dimensions
    If nd = 1 Then
        ' If this is a 1-element, 1D array
        If LBound(a, 1) = UBound(a, 1) Then
            If IsArray(a) Then Let InterpretableAsRowArrayQ = InterpretableAsRowArrayQ(First(a))
            Exit Function
        ' If this is a multi-element 1D array
        Else
            For i = LBound(a, 1) To UBound(a, 1)
                If Not AtomicQ(a(i)) Then Exit Function
            Next i
        End If
        
        Let InterpretableAsRowArrayQ = True
        Exit Function
    End If
    
    ' If we get here the array is two dimensional
    ' This is the 2D, single-element case
    If UBound(a, 1) = LBound(a, 1) And UBound(a, 2) = LBound(a, 2) Then
        Let InterpretableAsRowArrayQ = InterpretableAsRowArrayQ(a(LBound(a, 1), LBound(a, 2)))
        Exit Function
    ' This is the case when we have a matrix that cannot be interpreted as a row
    ElseIf UBound(a, 1) > LBound(a, 1) And UBound(a, 2) > LBound(a, 2) Then
        Exit Function
    ' This is the case when there is just one row
    ElseIf (UBound(a, 1) = LBound(a, 1)) And (UBound(a, 2) > LBound(a, 2)) Then
        For i = LBound(a, 2) To UBound(a, 2)
            If Not AtomicQ(a(LBound(a, 1), i)) Then
                Exit Function
            End If
        Next i
    ' This is the case when there is just one column of length > 1
    Else
        Exit Function
    End If
    
    Let InterpretableAsRowArrayQ = True
End Function

' DESCRIPTION
' Boolean function returning True if its argument could be interpreseted as a 2D, one-column
' array of atomic elements. This means that this function returns True for each of the following
' examples:
'
' 1. Array(1,2,3)
' 2. Array(Array(1,2,3))
' 3. [{1;2;3}], [{Array(1;2;3)}]
'
' PARAMETERS
' 1. arg - Any Excel value or reference
'
' RETURNED VALUE
' Returns True or False depending on whether or not its argument could be interpreted as a column
' vector.
Public Function InterpretableAsColumnArrayQ(a As Variant) As Boolean
    Dim nd As Integer
    Dim i As Long
    
    If EmptyArrayQ(a) Then
        Let InterpretableAsColumnArrayQ = True
        Exit Function
    End If
    
    Let InterpretableAsColumnArrayQ = False
    Let nd = NumberOfDimensions(a)
    
    If nd < 1 Or nd > 2 Then
        Exit Function
    End If
    
    If nd = 1 Then
        If LBound(a) = UBound(a) Then
            If Not Not AtomicQ(First(a)) Then
                Let InterpretableAsColumnArrayQ = True
            Else
                Let InterpretableAsColumnArrayQ = InterpretableAsColumnArrayQ(a(LBound(a, 1)))
            End If
        End If
        
        Exit Function
    End If
    
    If (UBound(a, 1) > LBound(a, 1)) And (UBound(a, 2) > LBound(a, 2)) Then
        Exit Function
    ElseIf (UBound(a, 1) > LBound(a, 1)) And (UBound(a, 2) = LBound(a, 2)) Then
        For i = LBound(a, 1) To UBound(a, 1)
            If IsArray(a(i, UBound(a, 2))) Then Exit Function
        Next i
    
        Let InterpretableAsColumnArrayQ = True
    ElseIf (UBound(a, 1) = LBound(a, 1)) And (UBound(a, 2) = LBound(a, 2)) Then
        If Not AtomicQ(a(LBound(a, 1), LBound(a, 2))) Then
            Let InterpretableAsColumnArrayQ = InterpretableAsColumnArrayQ(a(LBound(a, 1), LBound(a, 2)))
        Else
            Let InterpretableAsColumnArrayQ = True
        End If
    End If
End Function

' DESCRIPTION
' Boolean function returning True if TheValue is in the given 1D array. TheValue must
' satisfy NumberOrStringQ. Every element in TheArray must also satisfy NumberOfStringQ
'
' PARAMETERS
' 1. TheArray - A 1D array satisfying PrintableArrayQ
' 2. TheValue - Any value satisfying Predicates.PrintableQ
'
' RETURNED VALUE
' Returns True or False depending on whether or not the given value is in the given array
Public Function MemberQ(TheArray As Variant, TheValue As Variant) As Boolean
    Dim i As Long
    Dim var As Variant
    
    ' Assume result is False and change TheValue is in any one column of TheArray
    Let MemberQ = False
    
    ' Exit if TheArray is not a 1D array
    If NumberOfDimensions(TheArray) <> 1 Then Exit Function
    
    ' Exit with False if TheValue fails PrintableQ TheArray fails PrintableArrayQ
    If Not (PrintableQ(TheValue) And PrintableArrayQ(TheArray)) Then Exit Function
    
    For Each var In TheArray
        If IsEmpty(var) And IsEmpty(TheValue) Then
            Let MemberQ = True
            Exit Function
        End If
        
        If IsNull(var) And IsNull(TheValue) Then
            Let MemberQ = True
            Exit Function
        End If
    
        If TypeName(var) = TypeName(TheValue) And var = TheValue Then
            Let MemberQ = True
            Exit Function
        End If
    Next
End Function

' DESCRIPTION
' Boolean function returning True if TheValue is not in the given 1D array. TheValue must
' satisfy NumberOrStringQ. Every element in TheArray must also satisfy NumberOfStringQ
'
' PARAMETERS
' 1. TheArray - A 1D array satisfying PrintableArrayQ
' 2. TheValue - Any value satisfying PrintableQ
'
' RETURNED VALUE
' Returns True or False depending on whether or not the given value is in the given array
Public Function FreeQ(TheArray As Variant, TheValue As Variant) As Boolean
    Let FreeQ = IsArray(TheArray) And Not MemberQ(TheArray, TheValue)
End Function

' DESCRIPTION
' Boolean function returning True if the given workbook has a worksheet with the given name.
'
' PARAMETERS
' 1. aWorkbook - A workbook reference
' 2. WorksheetName - A worksheet reference
'
' RETURNED VALUE
' Returns True or False depending on whether or not the given workbook has a worksheet with
' the given name
Public Function WorksheetExistsQ(aWorkbook As Workbook, WorksheetName As String) As Boolean
    Let WorksheetExistsQ = False
    
    On Error Resume Next
    
    Let WorksheetExistsQ = aWorkbook.Worksheets(WorksheetName).Name <> ""
    Exit Function
    
    On Error GoTo 0
End Function

' DESCRIPTION
' Boolean function returning True if the given workbook has a sheet with the given name.
'
' PARAMETERS
' 1. aWorkbook - A workbook reference
' 2. WorksheetName - A worksheet reference
'
' RETURNED VALUE
' Returns True or False depending on whether or not the given workbook has a sheet with
' the given name
Public Function SheetExistsQ(aWorkbook As Workbook, SheetName As String) As Boolean
    Let SheetExistsQ = False
    
    On Error GoTo NoSuchSheet
    If Len(aWorkbook.Sheets(SheetName).Name) > 0 Then
        Let SheetExistsQ = True
        Exit Function
    End If

NoSuchSheet:
End Function
