Attribute VB_Name = "ArrayPredicates"
Option Explicit
Option Base 1

' DESCRIPTION
' Boolean function returning True if its argument is a dimensioned array. Returns
' False otherwise.  In other words, it returns False if its arg is neither an
' an array nor dimensioned.tabb
'
' PARAMETERS
' 1. arg - any value or object reference
'
' RETURNED VALUE
' True when arg is a dimensioned array. False otherwise.
Public Function DimensionedQ(Arg As Variant) As Boolean
    Dim i As Long
    
    On Error Resume Next
    
    ' Exit with False
    If Not IsArray(Arg) Then
        Let DimensionedQ = False
        Exit Function
    End If
    
    ' If arg has not been dimensioned, this following line will raise an error.
    ' Due to On Error Resume Next, the code will resume in the next line, which
    ' will then check if an error has been raised.
    Let i = UBound(Arg, 1)
    
    Let DimensionedQ = Err.Number = 0
End Function

' DESCRIPTION
' Boolean function returning True if its argument is an Empty 1D array. A prerequisite is that
' the array be dimensioned.  Returns False otherwise.
'
' PARAMETERS
' 1. arg - any value or object reference
'
' RETURNED VALUE
' True when arg is a dimensioned, 1D, empty array.  Returns False otherwise
Public Function EmptyArrayQ(Arg As Variant) As Boolean
    ' Exit with False if AnArray is not an array or has not been dimensioned
    If Not DimensionedQ(Arg) Then
        Let EmptyArrayQ = False
    ' Return True if we have an array with lower Ubound than LBound
    Else
        Let EmptyArrayQ = UBound(Arg, 1) - LBound(Arg, 1) < 0
    End If
End Function

' DESCRIPTION
' Boolean function returning True if its argument is a non-empty array. Returns False otherwise.
'
' PARAMETERS
' 1. arg - any value or object reference
'
' RETURNED VALUE
' True when arg is a non-empty array.  Returns False otherwise
Public Function NonEmptyArrayQ(Arg As Variant) As Boolean
    Let NonEmptyArrayQ = IsArray(Arg) And Not EmptyArrayQ(Arg)
End Function

' DESCRIPTION
' Boolean function returns True if Returns True is its argument is 2D matrix with numeric entries.
'
' PARAMETERS
' 1. arg - Any Excel value or reference
'
' RETURNED VALUE
' Returns True or False depending on whether or not its argument can be considered a numerical matrix.
Public Function MatrixQ(Arg As Variant) As Boolean
    Dim var As Variant
    
    Let MatrixQ = False
    
    ' Not necessary to test for DimensionedQ since NumberOfDimensions returns 0 for none arrays
    If NumberOfDimensions(Arg) <> 2 Then Exit Function

    If Not NumberArrayQ(Arg) Then Exit Function
    
    Let MatrixQ = True
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
Public Function AtomicArrayQ(AnArray As Variant) As Boolean
    Let AtomicArrayQ = AllTrueQ(AnArray, ThisWorkbook, "AtomicQ")
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
Public Function NumberArrayQ(Arg As Variant) As Boolean
    Let NumberArrayQ = AllTrueQ(Arg, ThisWorkbook, "NumberQ")
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
Public Function NegativeWholeNumberArrayQ(Arg As Variant) As Boolean
    Let NegativeWholeNumberArrayQ = AllTrueQ(Arg, "NegativeWholeNumberQ")
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
Public Function NonNegativeWholeNumberArrayQ(Arg As Variant) As Boolean
    Let NonNegativeWholeNumberArrayQ = AllTrueQ(Arg, ThisWorkbook, "NonNegativeWholeNumberQ")
End Function

' DESCRIPTION
' Boolean function returning True if its argument is a dimensioned, non-empty array all of
' whose elements satisfyPredicates.NonzeroWholeNumberQ.
'
' PARAMETERS
' 1. arg - any value or object reference
'
' RETURNED VALUE
' True or False depending on whether or not its argument is dimensioned, non-empty and all
' its elements satisfy Predicates.NonzeroWholeNumberQ
Public Function NonzeroWholeNumberArrayQ(Arg As Variant) As Boolean
    Let NonzeroWholeNumberArrayQ = AllTrueQ(Arg, ThisWorkbook, "NonzeroWholeNumberQ")
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
Public Function WholeNumberOrStringArrayQ(AnArray As Variant) As Boolean
    Let WholeNumberOrStringArrayQ = AllTrueQ(AnArray, ThisWorkbook, "WholeNumberOrStringQ")
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
Public Function NumberOrStringArrayQ(AnArray As Variant) As Boolean
    Let NumberOrStringArrayQ = AllTrueQ(AnArray, ThisWorkbook, "NumberOrStringQ")
End Function

' DESCRIPTION
' Boolean function returning True if its argument is a dimensioned, non-empty array all of whose
' elements satisfy Predicates.WorkbookQ.  Returns False otherwise.
'
' PARAMETERS
' 1. arg - any value or object reference
'
' RETURNED VALUE
' True or False depending on whether arg is dimensioned, non-empty and all its elements satisfy
' Predicates.WorkbookQ
Public Function WorkbookArrayQ(AnArray As Variant) As Boolean
    Let WorkbookArrayQ = AllTrueQ(AnArray, ThisWorkbook, "WorkbookQ")
End Function

' DESCRIPTION
' Boolean function returning True if its argument is a dimensioned, non-empty array all of whose
' elements satisfy Predicates.WorksheetQ.  Returns False otherwise.
'
' PARAMETERS
' 1. arg - any value or object reference
'
' RETURNED VALUE
' True or False depending on whether arg is dimensioned, non-empty and all its elements satisfy
' Predicates.WorksheetQ
Public Function WorksheetArrayQ(AnArray As Variant) As Boolean
    Let WorksheetArrayQ = AllTrueQ(AnArray, ThisWorkbook, "WorksheetQ")
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
Public Function WholeNumberArrayQ(Arg As Variant) As Boolean
    Let WholeNumberArrayQ = AllTrueQ(Arg, ThisWorkbook, "WholeNumberQ")
End Function

' DESCRIPTION
' Boolean function returning True if its argument is a dimensioned, non-empty array all of whose
' elements satisfy Predicates.ListObjectQ.  Returns False otherwise.
'
' PARAMETERS
' 1. arg - any value or object reference
'
' RETURNED VALUE
' True or False depending on whether arg is dimensioned, non-empty and all its elements satisfy
' Predicates.ListObjectQ
Public Function ListObjectArrayQ(AnArray As Variant) As Boolean
    Let ListObjectArrayQ = AllTrueQ(AnArray, ThisWorkbook, "ListObjectQ")
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
Public Function DictionaryArrayQ(AnArray As Variant) As Boolean
    Let DictionaryArrayQ = AllTrueQ(AnArray, ThisWorkbook, "DictionaryQ")
End Function

' DESCRIPTION
' Boolean function returning True if its argument is a dimensioned, non-empty array all of whose
' elements satisfy Predicates.ErrorQ.  Returns False otherwise.
'
' PARAMETERS
' 1. arg - any value or object reference
'
' RETURNED VALUE
' True or False depending on whether arg is dimensioned, non-empty and all its elements satisfy
' Predicates.ErrorQ
Public Function ErrorArrayQ(AnArray As Variant) As Boolean
    Let ErrorArrayQ = AllTrueQ(AnArray, ThisWorkbook, "ErrorQ")
End Function

' DESCRIPTION
' Boolean function returning True if its argument is a dimensioned, non-empty array all of whose
' elements satisfy Predicates.DateQ.  Returns False otherwise.
'
' PARAMETERS
' 1. arg - any value or object reference
'
' RETURNED VALUE
' True or False depending on whether arg is dimensioned, non-empty and all its elements satisfy
' Predicates.DateQ
Public Function DateArrayQ(AnArray As Variant) As Boolean
    Let DateArrayQ = AllTrueQ(AnArray, ThisWorkbook, "ErrorQ")
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
Public Function BooleanArrayQ(AnArray As Variant) As Boolean
    Let BooleanArrayQ = AllTrueQ(AnArray, ThisWorkbook, "BooleanQ")
End Function

' DESCRIPTION
' Boolean function returns True if Returns True is its argument is 2D matrix with Atomic entries.
'
' PARAMETERS
' 1. arg - Any Excel value or reference
'
' RETURNED VALUE
' Returns True or False depending on whether or not its argument can be considered a table.
Public Function AtomicTableQ(Arg As Variant) As Boolean
    Let AtomicTableQ = NumberOfDimensions(Arg) = 2 And AtomicArrayQ(Flatten(Arg))
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
Public Function PrintableArrayQ(AnArray As Variant) As Boolean
    Let PrintableArrayQ = AllTrueQ(AnArray, ThisWorkbook, "PrintableQ")
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
Public Function PrintableTableQ(Arg As Variant) As Boolean
    Let PrintableTableQ = NumberOfDimensions(Arg) = 2 And PrintableArrayQ(Flatten(Arg))
End Function

' DESCRIPTION
' Boolean function returning True if the given parameter is a dimensioned, 1D array
' all of whose elements satisfy ZeroQ.
'
' PARAMETERS
' 1. arg - any Excel value or reference
'
' RETURNED VALUE
' Returns True or False depending on whether or not all the elements in the 1D, dimensioned
' array arg satisfy ZeroQ
Public Function ZeroArrayQ(Arg As Variant) As Boolean
    Let ZeroArrayQ = AllTrueQ(AnArray, ThisWorkbook, "ZeroQ")
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
Public Function PositiveWholeNumberArrayQ(Arg As Variant) As Boolean
    Let PositiveWholeNumberArrayQ = AllTrueQ(Arg, ThisWorkbook, "PositiveWholeNumberQ")
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
Public Function StringArrayQ(AnArray As Variant) As Boolean
    Let StringArrayQ = AllTrueQ(AnArray, ThisWorkbook, "StringQ")
End Function

' DESCRIPTION
' Boolean function returning True if the given parameter is a dimensioned, 1D array
' all of whose elements satisfy ZeroQ.
'
' PARAMETERS
' 1. arg - any Excel value or reference
'
' RETURNED VALUE
' Returns True or False depending on whether or not all the elements in the 1D, dimensioned
' array arg satisfy ZeroQ
Public Function OneArrayQ(Arg As Variant) As Boolean
    Let OneArrayQ = AllTrueQ(AnArray, ThisWorkbook, "OneQ")
End Function

' DESCRIPTION
' Boolean function returning True if the given parameter is a dimensioned, 1D array
' all of whose elements satisfy PartIndexQ.
'
' PARAMETERS
' 1. arg - any Excel value or reference
'
' RETURNED VALUE
' Returns True or False depending on whether or not all the elements in the 1D, dimensioned
' array arg satisfy PartIndexQ
Public Function PartIndexArrayQ(Arg As Variant) As Boolean
    Let PartIndexArrayQ = AllTrueQ(Arg, ThisWorkbook, "PartIndexQ")
End Function

' DESCRIPTION
' Boolean function returning True if its argument is an array of class Span instances
'
' PARAMETERS
' 1. arg - any Excel value or reference
'
' RETURNED VALUE
' Returns True or False depending on whether or not its argument is an array of class Soab instances
Public Function SpanArrayQ(Arg As Variant) As Boolean
    Let SpanArrayQ = AllTrueQ(Arg, ThisWorkbook, "SpanQ")
End Function

' DESCRIPTION
' Boolean function returning True if the given parameter is a dimensioned, 1D array
' all of whose elements satisfy TakeIndexQ.
'
' PARAMETERS
' 1. arg - any Excel value or reference
'
' RETURNED VALUE
' Returns True or False depending on whether or not all the elements in the 1D, dimensioned
' array arg satisfy TakeIndexQ
Public Function TakeIndexArrayQ(Arg As Variant) As Boolean
    Let TakeIndexArrayQ = AllTrueQ(Arg, ThisWorkbook, "TakeIndexQ")
End Function

' DESCRIPTION
' This function returns True if the two parameters are consistent so elementwise operations may be
' perform on it (e.g. addition, multiplication, and division).
'
' PARAMETERS
' 1. Matrix1 - a scalar, vector, or matrix
' 2. Matrix2 -  a scalar, vector, or matrix
'
' RETURNED VALUE
' True if the dimensions of the two arguments are consistent for elementwise operations. False
' otherwise.
Public Function ElementwiseArithmeticParameterConsistentQ(Arg1 As Variant, Arg2 As Variant)
    ' Set default return value when encountering erros
    Let ElementwiseArithmeticParameterConsistentQ = False
    
    ' Check parameter consistency
    If Not (IsNumeric(Arg1) Or VectorQ(Arg1) Or MatrixQ(Arg1)) Then Exit Function
    
    If Not (IsNumeric(Arg2) Or VectorQ(Arg2) Or MatrixQ(Arg2)) Then Exit Function
    
    If RowVectorQ(Arg1) And MatrixQ(Arg2) And _
       GetNumberOfColumns(Arg1) <> GetNumberOfColumns(Arg2) Then Exit Function
    
    If MatrixQ(Arg1) And RowVectorQ(Arg2) And _
       GetNumberOfColumns(Arg1) <> GetNumberOfColumns(Arg2) Then Exit Function
    
    If ColumnVectorQ(Arg1) And MatrixQ(Arg2) Then
        If GetNumberOfRows(Arg1) <> GetNumberOfRows(Arg2) Then
            Exit Function
        Else
            Let ElementwiseArithmeticParameterConsistentQ = True
            Exit Function
        End If
    End If

    If ColumnVectorQ(Arg2) And MatrixQ(Arg1) Then
        If GetNumberOfRows(Arg1) <> GetNumberOfRows(Arg2) Then
            Exit Function
        Else
            Let ElementwiseArithmeticParameterConsistentQ = True
            Exit Function
        End If
    End If
    
    If MatrixQ(Arg1) And ColumnVectorQ(Arg2) And _
       GetNumberOfRows(Arg1) <> GetNumberOfRows(Arg2) Then Exit Function
       
    If MatrixQ(Arg1) And MatrixQ(Arg2) Then
        If GetNumberOfRows(Arg1) <> GetNumberOfRows(Arg2) Or _
           GetNumberOfColumns(Arg1) <> GetNumberOfColumns(Arg2) Then Exit Function
    End If
        
    If EmptyArrayQ(Arg1) And Not EmptyArrayQ(Arg2) Then Exit Function
    
    If EmptyArrayQ(Arg2) And Not EmptyArrayQ(Arg1) Then Exit Function
    
    Let ElementwiseArithmeticParameterConsistentQ = True
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
Function VectorQ(Arg As Variant) As Boolean
    Let VectorQ = RowVectorQ(Arg) Or ColumnVectorQ(Arg)
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
Public Function ColumnVectorQ(Arg As Variant) As Boolean
    Dim var As Variant
    
    Let ColumnVectorQ = False
    
    If Not DimensionedQ(Arg) Then Exit Function
    
    ' Returning True for an empty array is necesary recursion to work properly on column vectors
    If EmptyArrayQ(Arg) Then
        Let ColumnVectorQ = True
        Exit Function
    End If
    
    If NumberOfDimensions(Arg) <> 2 Then Exit Function
    
    If LBound(Arg, 2) <> UBound(Arg, 2) Then Exit Function
    
    If Not NumberArrayQ(Arg) Then Exit Function

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
Public Function RowVectorQ(Arg As Variant) As Boolean
    Dim var As Variant
    
    Let RowVectorQ = False
    
    If Not DimensionedQ(Arg) Then Exit Function
    
    ' Returning True for an empty array is necesary recursion to work properly on row vectors
    If EmptyArrayQ(Arg) Then
        Let RowVectorQ = True
        Exit Function
    End If
    
    If NumberOfDimensions(Arg) <> 1 Then Exit Function
    
    If Not NumberArrayQ(Arg) Then Exit Function
    
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
Public Function ColumnArrayQ(Arg As Variant) As Boolean
    Dim var As Variant
    
    Let ColumnArrayQ = False
    
    ' Returning True for an empty array is necesary recursion to work properly on column vectors
    If EmptyArrayQ(Arg) Then
        Let ColumnArrayQ = True
        Exit Function
    End If
    
    If NumberOfDimensions(Arg) <> 2 Then Exit Function
    
    If LBound(Arg, 2) <> UBound(Arg, 2) Then Exit Function
    
    If Not AtomicArrayQ(Arg) Then Exit Function

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
Public Function RowArrayQ(Arg As Variant) As Boolean
    Let RowArrayQ = AtomicArrayQ(Arg)
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
Function RowOrColumnArrayQ(Arg As Variant) As Boolean
    Let RowOrColumnArrayQ = RowArrayQ(Arg) Or ColumnArrayQ(Arg)
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
Public Function InterpretableAsRowArrayQ(A As Variant) As Boolean
    Dim nd As Integer
    Dim i As Long
    
    Let InterpretableAsRowArrayQ = False
    Let nd = NumberOfDimensions(A)
    
    ' If a has more than two or fewer than one dimension, then exit with False.
    If nd > 2 Or nd < 1 Then
        Exit Function
    End If
    
    ' Process arg is it has one dimensions
    If nd = 1 Then
        ' If this is a 1-element, 1D array
        If LBound(A, 1) = UBound(A, 1) Then
            If IsArray(A) Then Let InterpretableAsRowArrayQ = InterpretableAsRowArrayQ(First(A))
            Exit Function
        ' If this is a multi-element 1D array
        Else
            For i = LBound(A, 1) To UBound(A, 1)
                If Not AtomicQ(A(i)) Then Exit Function
            Next i
        End If
        
        Let InterpretableAsRowArrayQ = True
        Exit Function
    End If
    
    ' If we get here the array is two dimensional
    ' This is the 2D, single-element case
    If UBound(A, 1) = LBound(A, 1) And UBound(A, 2) = LBound(A, 2) Then
        Let InterpretableAsRowArrayQ = InterpretableAsRowArrayQ(A(LBound(A, 1), LBound(A, 2)))
        Exit Function
    ' This is the case when we have a matrix that cannot be interpreted as a row
    ElseIf UBound(A, 1) > LBound(A, 1) And UBound(A, 2) > LBound(A, 2) Then
        Exit Function
    ' This is the case when there is just one row
    ElseIf (UBound(A, 1) = LBound(A, 1)) And (UBound(A, 2) > LBound(A, 2)) Then
        For i = LBound(A, 2) To UBound(A, 2)
            If Not AtomicQ(A(LBound(A, 1), i)) Then
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
Public Function InterpretableAsColumnArrayQ(A As Variant) As Boolean
    Dim nd As Integer
    Dim i As Long
    
    If EmptyArrayQ(A) Then
        Let InterpretableAsColumnArrayQ = True
        Exit Function
    End If
    
    Let InterpretableAsColumnArrayQ = False
    Let nd = NumberOfDimensions(A)
    
    If nd < 1 Or nd > 2 Then
        Exit Function
    End If
    
    If nd = 1 Then
        If LBound(A) = UBound(A) Then
            If Not Not AtomicQ(First(A)) Then
                Let InterpretableAsColumnArrayQ = True
            Else
                Let InterpretableAsColumnArrayQ = InterpretableAsColumnArrayQ(A(LBound(A, 1)))
            End If
        End If
        
        Exit Function
    End If
    
    If (UBound(A, 1) > LBound(A, 1)) And (UBound(A, 2) > LBound(A, 2)) Then
        Exit Function
    ElseIf (UBound(A, 1) > LBound(A, 1)) And (UBound(A, 2) = LBound(A, 2)) Then
        For i = LBound(A, 1) To UBound(A, 1)
            If IsArray(A(i, UBound(A, 2))) Then Exit Function
        Next i
    
        Let InterpretableAsColumnArrayQ = True
    ElseIf (UBound(A, 1) = LBound(A, 1)) And (UBound(A, 2) = LBound(A, 2)) Then
        If Not AtomicQ(A(LBound(A, 1), LBound(A, 2))) Then
            Let InterpretableAsColumnArrayQ = InterpretableAsColumnArrayQ(A(LBound(A, 1), LBound(A, 2)))
        Else
            Let InterpretableAsColumnArrayQ = True
        End If
    End If
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
Public Function PrintableQ(Arg As Variant) As Boolean
    Let PrintableQ = AnyTrueQ(Through(GetPrintableTypePredicateNames(), Arg))
End Function


