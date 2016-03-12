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
    
    If Not IsArray(arg) Then
        Let DimensionedQ = False
        Exit Function
    End If
    
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
Public Function EmptyArrayQ(AnArray As Variant) As Boolean
    ' Exit with False if AnArray is not an array or has not been dimensioned
    If Not DimensionedQ(AnArray) Then
        Let EmptyArrayQ = False
    ' Return True if we have an array with lower Ubound than LBound
    Else
        Let EmptyArrayQ = UBound(AnArray, 1) - LBound(AnArray, 1) < 0
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
'
' It returns False otherwise
'
' PARAMETERS
' 1. arg - any value or object reference
'
' RETURNED VALUE
' True when arg has one of the types detailed above. Returns False otherwise
' Returns True if the argument is one of the following:
' 1. Number
' 2. String
' 3. Date
' 4. Boolean
' 5. Error
' 6. Worksheet
' 7. Workbook
' 8. ListObject
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
    Let AtomicArrayQ = ArrayQHelper(arg, "AtomicQ")
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
    Dim AnInteger As Integer
    Dim ALong As Long
    Dim ASingle As Single
    Dim ADouble As Double
    Dim ACurrency As Currency

    Select Case TypeName(arg)
        Case TypeName(AByte)
            Let NumberQ = True
        Case TypeName(AnInteger)
            Let NumberQ = True
        Case TypeName(ALong)
            Let NumberQ = True
        Case TypeName(ASingle)
            Let NumberQ = True
        Case TypeName(ADouble)
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
    Let NumberArrayQ = ArrayQHelper(arg, "NumberQ")
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
    Dim AnInteger As Integer
    Dim ALong As Long

    Select Case TypeName(arg)
        Case TypeName(AByte)
            Let WholeNumberQ = True
        Case TypeName(AnInteger)
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
    Let WholeNumberArrayQ = ArrayQHelper(arg, "WholeNumberQ")
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
    Let PositiveWholeNumberArrayQ = ArrayQHelper(arg, "PositiveWholeNumberQ")
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
    Let NegativeWholeNumberArrayQ = ArrayQHelper(arg, "NegativeWholeNumberQ")
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
    Let NonNegativeWholeNumberArrayQ = ArrayQHelper(arg, "NonNegativeWholeNumberQ")
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
Public Function StringArrayQ(AnArray As Variant) As Boolean
    Let StringArrayQ = ArrayQHelper(arg, "StringQ")
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
Public Function DictionaryArrayQ(AnArray As Variant) As Boolean
    Let DictionaryArrayQ = ArrayQHelper(arg, "DictionaryQ")
End Function

' DESCRIPTION
' Boolean function returning True if its argument is a string or a whole number. Returns False otherwise.
'
' PARAMETERS
' 1. arg - any value or object reference
'
' RETURNED VALUE
' True or False depending on whether or not its argument string or a dictionary.
Public Function WholeNumberOrStringQ(arg As Variant) As Boolean
    Let WholeNumberOrStringQ = Not (WholeNumberQ(arg) Or StringQ(arg))
End Function

' DESCRIPTION
' Boolean function returning True if its argument is a dimensioned, non-empty array all of whose
' elements satisfy Predicates.WholeNumberOrStringQ.  Returns False otherwise.
'
' PARAMETERS
' 1. arg - any value or object reference
'
' RETURNED VALUE
' True or False depending on whether arg is dimensioned, non-empty and all its elements satisfy
' Predicates.WholeNumberOrStringQ
Public Function WholeNumberOrStringArrayQ(AnArray As Variant) As Boolean
    Let WholeNumberOrStringArrayQ = ArrayQHelper(arg, "WholeNumberOrStringQ")
End Function

' DESCRIPTION
' Boolean helper function returning True if its argument is a dimensioned, non-empty array all of
' whose elements satisfy var \mapsto Application.Run(ThisWorkbook.Name & "!" & PredicateName, var).
' Returns False otherwise.
'
' PARAMETERS
' 1. arg - any value or object reference
'
' RETURNED VALUE
' True or False depending on whether arg is dimensioned, non-empty and all its elements satisfy
' Predicates.WholeNumberOrStringQ
Private Function ArrayQHelper(AnArray As Variant, PredicateName As String) As Boolean
    Dim var As Variant

    Let ArrayQHelper = True

    If Not IsArray(AnArray) Then
        Let ArrayQHelper = False
        Exit Function
    End If
    
    For Each var In AnArray
        If Not Application.Run(ThisWorkbook.Name & "!" & PredicateName, var) Then
            Let ArrayQHelper = False
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
    
    If NumberOfDimensions(arg) <> 2 Then
        Let MatrixQ = False
        Exit Function
    End If

    For Each var In arg
        If Not IsNumeric(var) Then
            Let MatrixQ = False
            Exit Function
        End If
    Next
    
    Let MatrixQ = True
End Function

' DESCRIPTION
' Boolean function returns True if Returns True is its argument is 2D matrix with non-array entries.
'
' PARAMETERS
' 1. arg - Any Excel value or reference
'
' RETURNED VALUE
' Returns True or False depending on whether or not its argument can be considered a table matrix.
Public Function TableQ(arg As Variant) As Boolean
    Dim var As Variant

    Let TableQ = False
    
    If NumberOfDimensions(arg) <> 2 Then Exit Function

    For Each var In arg
        If IsArray(var) Then Exit Function
    Next var
    
    Let TableQ = True
End Function

' Returns True is the argument is a row or column vector as defined by ColumVectorQ and RowVectorQ
Function VectorQ(arg As Variant) As Boolean
    Let VectorQ = RowVectorQ(arg) Or ColumnVectorQ(arg)
End Function

' Returns True is the argument is a column vector (e.g. n x 1 2D array with atomic entries exclusively)
' comprised of numeric, atomic expressions excluively
Public Function ColumnVectorQ(arg As Variant) As Boolean
    Dim var As Variant
    
    Let ColumnVectorQ = False
    
    If EmptyArrayQ(arg) Then Exit Function
    
    If NumberOfDimensions(arg) <> 2 Then Exit Function
    
    If LBound(arg, 2) <> UBound(arg, 2) Then Exit Function
    
    For Each var In arg
        If Not IsNumeric(var) Then Exit Function
    Next

    Let ColumnVectorQ = True
End Function

' Returns True if the argument is a row vector (e.g. a 1D vector comprised exclusively of
' numeric entries).  Returns False for the empty 1D array
Public Function RowVectorQ(arg As Variant) As Boolean
    Dim var As Variant
    
    Let RowVectorQ = False
    
    If Not DimensionedQ(arg) Then Exit Function
    
    If EmptyArrayQ(arg) Then Exit Function
    
    If NumberOfDimensions(arg) <> 1 Then Exit Function
    
    If Not NumberArrayQ(arg) Then Exit Function
    
    Let RowVectorQ = True
End Function

' Returns true if the array could be interpreseted as a 1D array of atomic elements.
' This means that this function returns True for each of the following examples:
' 1. Array(1,2,3)
' 2. Array(Array(1,2,3))
' 3. Array(Array(Array(1,2,3)))),
' 4. [{1,2,3}], [{Array(1,2,3)}], etc. will evaluate to True.
Public Function RowArrayQ(a As Variant) As Boolean
    Dim nd As Integer
    Dim i As Long
    
    Let RowArrayQ = False
    Let nd = NumberOfDimensions(a)
    
    ' If a has more than two or fewer than one dimension, then exit with False.
    If nd > 2 Or nd < 1 Then
        Exit Function
    End If
    
    ' Process arg is it has one dimensions
    If nd = 1 Then
        ' If this is a 1-element, 1D array
        If LBound(a, 1) = UBound(a, 1) Then
            If IsArray(a) Then Let RowArrayQ = RowArrayQ(a(LBound(a)))
            Exit Function
        ' If this is a multi-element 1D array
        Else
            For i = LBound(a) To UBound(a)
                If IsArray(a(i), 1) Then Exit Function
            Next i
        End If
        
        Let RowArrayQ = True
        Exit Function
    End If
    
    ' If we get here the array is two dimensional
    ' This is the 2D, single-element case
    If UBound(a, 1) = LBound(a, 1) And UBound(a, 2) = LBound(a, 2) Then
        Let RowArrayQ = RowArrayQ(a(LBound(a, 1), LBound(a, 2)))
        Exit Function
    ' This is the case when we have a matrix that cannot be interpreted as a row
    ElseIf UBound(a, 1) > LBound(a, 1) And UBound(a, 2) > LBound(a, 2) Then
        Exit Function
    ' This is the case when there is just one row
    ElseIf (UBound(a, 1) = LBound(a, 1)) And (UBound(a, 2) > LBound(a, 2)) Then
        For i = LBound(a, 2) To UBound(a, 2)
            If IsArray(a(LBound(a, 1), i)) Then
                Exit Function
            End If
        Next i
    ' This is the case when there is just one column of length > 1
    Else
        Exit Function
    End If
    
    Let RowArrayQ = True
End Function

' Returns true if the array is a row array (must be a 2D array)
' This returns TRUE only for a 2D array with one column
Public Function ColumnArrayQ(a As Variant) As Boolean
    Dim nd As Integer
    Dim i As Long
    
    If EmptyArrayQ(a) Then
        Let ColumnArrayQ = True
        Exit Function
    End If
    
    Let ColumnArrayQ = False
    Let nd = NumberOfDimensions(a)
    
    If nd < 1 Or nd > 2 Then
        Exit Function
    End If
    
    If nd = 1 Then
        If LBound(a) = UBound(a) Then
            If Not IsArray(a(LBound(a))) Then
                Let ColumnArrayQ = True
            Else
                Let ColumnArrayQ = ColumnArrayQ(a(LBound(a, 1)))
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
    
        Let ColumnArrayQ = True
    ElseIf (UBound(a, 1) = LBound(a, 1)) And (UBound(a, 2) = LBound(a, 2)) Then
        If IsArray(a(LBound(a, 1), LBound(a, 2))) Then
            Let ColumnArrayQ = ColumnArrayQ(a(LBound(a, 1), LBound(a, 2)))
        Else
            Let ColumnArrayQ = True
        End If
    End If
End Function

' This function determines if the given value is in the given array (1D or 2D)
Public Function MemberQ(TheArray As Variant, TheValue As Variant) As Boolean
    Dim i As Long
    Dim TheResultFlag As Boolean
    
    ' Assume result is False and change TheValue is in any one column of TheArray
    Let TheResultFlag = False
    
    If NumberOfDimensions(TheArray) <= 1 Then
        Let MemberQ = IsValueIn1DArray(TheArray, TheValue)
        
        Exit Function
    End If
    
    ' Go through all the columns, checking if the result is true for one of the columns
    For i = 1 To UBound(TheArray, 2)
        Let TheResultFlag = TheResultFlag Or IsValueIn1DArray(ConvertTo1DArray(GetColumn(TheArray, i)), TheValue)
    Next i
    
    ' Set the value to return
    Let MemberQ = TheResultFlag
End Function

' This function determines if the given value is missig from the given array (1D or 2D).
' It is the complete opposite of MemberQ
Public Function FreeQ(TheArray As Variant, TheValue As Variant) As Boolean
    Let FreeQ = Not MemberQ(TheArray, TheValue)
End Function

' This is private, helper function used by MemberQ above. This function returns true only
' if TheArray has dimension 0 and is equal to TheValue or if TheArray has dimension 1 and
' TheValue is in TheArray.
Private Function IsValueIn1DArray(TheArray As Variant, TheValue As Variant) As Boolean
    If EmptyArrayQ(TheArray) Then
        Let IsValueIn1DArray = False
    
        Exit Function
    End If
    
    If NumberOfDimensions(TheArray) > 1 Then
        Let IsValueIn1DArray = False
    ElseIf NumberOfDimensions(TheArray) = 0 Then
        Let IsValueIn1DArray = (TheValue = TheArray)
    ElseIf IsNumeric(Application.Match(TheValue, TheArray, 0)) Then
        Let IsValueIn1DArray = True
    Else
        Let IsValueIn1DArray = False
    End If
End Function

' This predicate returns true if the given workbook has a worksheet with name WorksheetName.
' Otherwise, it returns false
Public Function WorksheetExistsQ(AWorkBook As Workbook, WorksheetName As String) As Boolean
    Let WorksheetExistsQ = False
    
    On Error Resume Next
    
    Let WorksheetExistsQ = AWorkBook.Worksheets(WorksheetName).Name <> ""
    Exit Function
    
    On Error GoTo 0
End Function

' This predicate returns true if the given workbook has a sheet with name WorksheetName.
' Otherwise, it returns false
Public Function SheetExistsQ(AWorkBook As Workbook, SheetName As String) As Boolean
    ' returns TRUE if the sheet exists in the active workbook
    Let SheetExistsQ = False
    
    On Error GoTo NoSuchSheet
    If Len(AWorkBook.Sheets(SheetName).Name) > 0 Then
        Let SheetExistsQ = True
        Exit Function
    End If

NoSuchSheet:

End Function







