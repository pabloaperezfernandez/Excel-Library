Attribute VB_Name = "Predicates"
Option Explicit
Option Base 1

' DESCRIPTION
' This function returns an array of strings with the names the predicates used to
' identify atomic values.
'
' PARAMETERS
'
' RETURNED VALUE
' array of strings with the names the predicates used to identify atomic values.
Public Function GetAtomicTypePredicateNames() As Variant
    Let GetAtomicTypePredicateNames = Array("BooleanQ", _
                                            "DateQ", _
                                            "DictionaryQ", _
                                            "EmptyQ", _
                                            "ErrorQ", _
                                            "ListObjectQ", _
                                            "NullQ", _
                                            "NumberQ", _
                                            "StringQ", _
                                            "WorkbookQ", _
                                            "WorksheetQ")
End Function

' DESCRIPTION
' This function returns an array of strings with the names the predicates used to
' identify printable values.
'
' PARAMETERS
'
' RETURNED VALUE
' array of strings with the names the predicates used to identify pritable values.
Public Function GetPrintableTypePredicateNames() As Variant
    Let GetPrintableTypePredicateNames = Array("BooleanQ", _
                                               "DateQ", _
                                               "EmptyQ", _
                                               "NullQ", _
                                               "NumberQ", _
                                               "StringQ")
End Function

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
' the array be dimensioned.  Returns False otherwise.
'
' PARAMETERS
' 1. arg - any value or object reference
'
' RETURNED VALUE
' True when arg is a dimensioned, 1D, empty array.  Returns False otherwise
Public Function EmptyArrayQ(arg As Variant) As Boolean
    ' Exit with False if AnArray is not an array or has not been dimensioned
    If Not DimensionedQ(arg) Then
        Let EmptyArrayQ = False
    ' Return True if we have an array with lower Ubound than LBound
    Else
        Let EmptyArrayQ = UBound(arg, 1) - LBound(arg, 1) < 0
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
Public Function NonEmptyArrayQ(arg As Variant) As Boolean
    Let NonEmptyArrayQ = IsArray(arg) And Not EmptyArrayQ(arg)
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
' 9. Dictionary
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
    Dim var As Variant
    
    Let AtomicQ = True
    
    For Each var In GetAtomicTypePredicateNames()
        If Run(var, arg) Then Exit Function
    Next
    
    Let AtomicQ = False
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
    Dim TheVarType As VbVarType

    Let TheVarType = VarType(arg)

    If TheVarType = vbByte Or TheVarType = vbCurrency Or TheVarType = vbDecimal Or _
       TheVarType = vbDouble Or TheVarType = vbInteger Or TheVarType = vbLong Or _
       TheVarType = vbSingle Then
        Let NumberQ = True
    Else
        Let NumberQ = False
    End If
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
    Let NumberArrayQ = AllTrueQ(arg, ThisWorkbook, "NumberQ")
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
    Dim TheVarType As VbVarType

    Let TheVarType = VarType(arg)

    If TheVarType = vbByte Or TheVarType = vbInteger Or TheVarType = vbLong Then
        Let WholeNumberQ = True
    ' The reason this next iff is necessary is that [{1,2}] passes 1 and 2 with type
    ' double.  Don't know why.  Makes no sense.  In any case, all we care about is that
    ' this have no nonzero decimal part.
    ElseIf TheVarType = vbDecimal Or TheVarType = vbDouble Or TheVarType = vbDouble Then
        Let WholeNumberQ = CLng(arg) = arg
    Else
        Let WholeNumberQ = False
    End If
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
    Let WholeNumberArrayQ = AllTrueQ(arg, ThisWorkbook, "WholeNumberQ")
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
    Let PositiveWholeNumberArrayQ = AllTrueQ(arg, ThisWorkbook, "PositiveWholeNumberQ")
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
    Let NegativeWholeNumberArrayQ = AllTrueQ(arg, "NegativeWholeNumberQ")
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
    Let NonNegativeWholeNumberArrayQ = AllTrueQ(arg, ThisWorkbook, "NonNegativeWholeNumberQ")
End Function

' DESCRIPTION
' Boolean function returning True if its argument is a nonzero whole number.
'
' PARAMETERS
' 1. arg - any value or object reference
'
' RETURNED VALUE
' True or False depending on whether or not its argument is a nonzero, whole number.
Public Function NonzeroWholeNumberQ(arg As Variant) As Boolean
    If WholeNumberQ(arg) Then
        Let NonzeroWholeNumberQ = arg <> 0
    Else
        Let NonzeroWholeNumberQ = False
    End If
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
Public Function NonzeroWholeNumberArrayQ(arg As Variant) As Boolean
    Let NonzeroWholeNumberArrayQ = AllTrueQ(arg, ThisWorkbook, "NonzeroWholeNumberQ")
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
    Let StringArrayQ = AllTrueQ(AnArray, ThisWorkbook, "StringQ")
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
Public Function WholeNumberOrStringArrayQ(AnArray As Variant) As Boolean
    Let WholeNumberOrStringArrayQ = AllTrueQ(AnArray, ThisWorkbook, "WholeNumberOrStringQ")
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
Public Function NumberOrStringArrayQ(AnArray As Variant) As Boolean
    Let NumberOrStringArrayQ = AllTrueQ(AnArray, ThisWorkbook, "NumberOrStringQ")
End Function

' DESCRIPTION
' Boolean function returning True if its argument is an initialized workbook reference.
'
' PARAMETERS
' 1. arg - an initialized workbook reference
'
' RETURNED VALUE
' True or False depending on whether arg its argument is an initialized workbook reference.
Public Function WorkbookQ(arg As Variant) As Boolean
    Let WorkbookQ = TypeName(arg) = TypeName(ThisWorkbook)
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
' Boolean function returning True if its argument is an initialized worksheet reference.
'
' PARAMETERS
' 1. arg - an initialized workbook reference
'
' RETURNED VALUE
' True or False depending on whether arg its argument is an initialized worksheet reference.
Public Function WorksheetQ(arg As Variant) As Boolean
    Let WorksheetQ = TypeName(arg) = TypeName(ThisWorkbook.Worksheets(1))
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
' Boolean function returning True if its argument is an initialized ListObject reference.
'
' PARAMETERS
' 1. arg - an initialized workbook reference
'
' RETURNED VALUE
' True or False depending on whether arg its argument is an initialized ListObject reference.
Public Function ListObjectQ(arg As Variant) As Boolean
    Let ListObjectQ = TypeName(arg) = "ListObject"
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
    Let DictionaryArrayQ = AllTrueQ(AnArray, ThisWorkbook, "DictionaryQ")
End Function

' DESCRIPTION
' Boolean function returning True if its argument satisfies IsEmpty(). Returns False otherwise.
'
' PARAMETERS
' 1. arg - any value or object reference
'
' RETURNED VALUE
' True or False depending on whether or not its argument satisfies IsEmpty()
Public Function EmptyQ(vValue As Variant) As Boolean
    Let EmptyQ = IsEmpty(vValue)
End Function

' DESCRIPTION
' Boolean function returning True if its argument satisfies IsNull(). Returns False otherwise.
'
' PARAMETERS
' 1. arg - any value or object reference
'
' RETURNED VALUE
' True or False depending on whether or not its argument satisfies IsNull()
Public Function NullQ(vValue As Variant) As Boolean
    Let NullQ = IsNull(vValue)
End Function

' DESCRIPTION
' Boolean function returning True if its argument satisfies IsError(). Returns False otherwise.
'
' PARAMETERS
' 1. arg - any value or object reference
'
' RETURNED VALUE
' True or False depending on whether or not its argument satisfies IsError()
Public Function ErrorQ(vValue As Variant) As Boolean
    Let ErrorQ = IsNull(vValue)
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
' Boolean function returning True if its argument satisfies IsDate(). Returns False otherwise.
'
' PARAMETERS
' 1. arg - any value or object reference
'
' RETURNED VALUE
' True or False depending on whether or not its argument satisfies IsDate()
Public Function DateQ(vValue As Variant) As Boolean
    Let DateQ = IsDate(vValue)
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
' Boolean function returning True if all of the elements in AnArray satisfy the predicate whose
' name is PredicateName.  If PredicateName is not provided, then the function returns True when
' all of the elements of AnArray are True.
'
' PARAMETERS
' 1. AnArray - A dimensioned 1D array
' 2. PredicateName (optional) - A string representing the predicates name
' 3. WorkbookReference (optional) - A workbook reference to the workbook holding the predicate
'
' RETURNED VALUE
' True or False depending on whether arg is dimensioned and all its elements satisfy the
' predicate with name PredicateName or are True when PredicateName is missing.
Public Function AllTrueQ(AnArray As Variant, _
                         Optional WorkbookReference As Variant, _
                         Optional PredicateName As Variant) As Boolean
    Dim var As Variant
    
    ' Set the default return value of True
    Let AllTrueQ = True
    
    ' Exit with False if AnArrat is not dimensioned
    If Not DimensionedQ(AnArray) Then
        Let AllTrueQ = False
        Exit Function
    End If
    
    ' Exit the True of AnArray is an empty array
    If EmptyArrayQ(AnArray) Then Exit Function
    
    ' Exit with Null if PredicateName not missing and PredicateName not a string
    ' Exit with Null if PredicateName missing and AnArray is not an array of Booleans
    ' Cannot use BooleanArrayQ here because it would cause a circular definitional reference
    If IsMissing(PredicateName) Then
        For Each var In AnArray
            If Not BooleanQ(var) Then
                Let AllTrueQ = False
                Exit Function
            End If
        Next
    Else
        If Not StringQ(PredicateName) Then
            Let AllTrueQ = False
            Exit Function
        End If
    End If
    
    ' Exit with Null if only one of WorkbookReference and PredicateName is missing
    If (IsMissing(WorkbookReference) And Not IsMissing(PredicateName)) Or _
       (Not IsMissing(WorkbookReference) And IsMissing(PredicateName)) Then
        Let AllTrueQ = False
        Exit Function
    End If
    
    If Not IsMissing(PredicateName) Then
        For Each var In AnArray
            If Not Application.Run(WorkbookReference.Name & "!" & PredicateName, var) Then
                Let AllTrueQ = False
                Exit Function
            End If
        Next
    Else
        For Each var In AnArray
            If Not var Then
                Let AllTrueQ = False
                Exit Function
            End If
        Next
    End If
End Function

' DESCRIPTION
' Boolean function returning True if all of the elements in AnArray fail the predicate whose
' name is PredicateName.  If PredicateName is missing, the function returns True when all
' of AnArray are False.
'
' PARAMETERS
' 1. AnArray - A dimensioned 1D or 2D array
' 2. PredicateName (optional) - A string representing the predicates name
' 3. WorkbookReference (optional) - A workbook reference to the workbook holding the predicate
'
' RETURNED VALUE
' True or False depending on whether arg is dimensioned and all its elements fail the
' predicate with name PredicateName or are False when PredicateName is missing.
Public Function NoneTrueQ(AnArray As Variant, _
                          Optional WorkbookReference As Variant, _
                          Optional PredicateName As Variant) As Boolean
    Dim var As Variant
    
    Let NoneTrueQ = True
                          
    ' Exit with Null if AnArray is not dimensioned
    If Not DimensionedQ(AnArray) Then
        Let NoneTrueQ = False
        Exit Function
    End If
    
    ' Exit with True if AnArray is empty. This case is necessary because NoneTrueQ is not logically
    ' the negation of AllTrueQ.  For NoneTrueQ to be true, all elements of AnArray must be False.
    ' However, all elements of AnArray are False if AnArray is an empty set.
    If EmptyArrayQ(AnArray) Then Exit Function

    ' Exit with Null if PredicateName not missing and PredicateName not a string
    ' Exit with Null if PredicateName missing and AnArray is not an array of Booleans
    ' Cannot use BooleanArrayQ here because it would cause a circular definitional reference
    If IsMissing(PredicateName) Then
        For Each var In AnArray
            If Not BooleanQ(var) Then
                Let NoneTrueQ = False
                Exit Function
            End If
        Next
    Else
        If Not StringQ(PredicateName) Then
            Let NoneTrueQ = False
            Exit Function
        End If
    End If
    
    For Each var In AnArray
        If Not IsMissing(PredicateName) Then
            If Application.Run(WorkbookReference.Name & "!" & PredicateName, var) Then
                Let NoneTrueQ = False
                Exit Function
            End If
        Else
            If var Then
                Let NoneTrueQ = False
                Exit Function
            End If
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
Public Function AnyTrueQ(AnArray As Variant, _
                         Optional WorkbookReference As Variant, _
                         Optional PredicateName As Variant) As Boolean
    Dim var As Variant
    
    ' Set the default return value of True
    Let AnyTrueQ = False

    ' Exit with False if AnArray is not dimensioned array
    If Not DimensionedQ(AnArray) Then
        Let AnyTrueQ = False
        Exit Function
    End If
    
    ' Exit the True of AnArray is an empty array
    If EmptyArrayQ(AnArray) Then
        Let AnyTrueQ = True
        Exit Function
    End If
    
    ' Exit with Null if PredicateName not missing and PredicateName not a string
    ' Exit with Null if PredicateName missing and AnArray is not an array of Booleans
    ' Cannot use BooleanArrayQ here because it would cause a circular definitional reference
    If IsMissing(PredicateName) Then
        For Each var In AnArray
            If Not BooleanQ(var) Then
                Let AnyTrueQ = False
                Exit Function
            End If
        Next
    Else
        If Not StringQ(PredicateName) Then
            Let AnyTrueQ = False
            Exit Function
        End If
    End If
    
    ' Exit with Null if only one of WorkbookReference and PredicateName is missing
    If (IsMissing(WorkbookReference) And Not IsMissing(PredicateName)) Or _
       (Not IsMissing(WorkbookReference) And IsMissing(PredicateName)) Then
        Let AnyTrueQ = False
        Exit Function
    End If
    
    If Not IsMissing(PredicateName) Then
        For Each var In AnArray
            If Application.Run(WorkbookReference.Name & "!" & PredicateName, var) Then
                Let AnyTrueQ = True
                Exit Function
            End If
        Next
    Else
        For Each var In AnArray
        If var Then
            Let AnyTrueQ = True
            Exit Function
        End If
        Next
    End If
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

    If Not NumberArrayQ(arg) Then Exit Function
    
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
    Let BooleanQ = VarType(arg) = vbBoolean
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
Public Function AtomicTableQ(arg As Variant) As Boolean
    Let AtomicTableQ = NumberOfDimensions(arg) = 2 And AtomicArrayQ(Flatten(arg))
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
    Let PrintableQ = AnyTrueQ(Through(GetPrintableTypePredicateNames(), arg))
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
Public Function PrintableTableQ(arg As Variant) As Boolean
    Let PrintableTableQ = NumberOfDimensions(arg) = 2 And PrintableArrayQ(Flatten(arg))
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
    
    If Not DimensionedQ(arg) Then Exit Function
    
    ' Returning True for an empty array is necesary recursion to work properly on column vectors
    If EmptyArrayQ(arg) Then
        Let ColumnVectorQ = True
        Exit Function
    End If
    
    If NumberOfDimensions(arg) <> 2 Then Exit Function
    
    If LBound(arg, 2) <> UBound(arg, 2) Then Exit Function
    
    If Not NumberArrayQ(arg) Then Exit Function

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
    
    If Not AtomicArrayQ(arg) Then Exit Function

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
' Boolean function returning True if TheValue is in the given 1D array.
'
' PARAMETERS
' 1. TheArray - A 1D array
' 2. TheValue - Any Excel value or reference
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
    
    For Each var In TheArray
        If IsError(var) Then
            Let MemberQ = False
            Exit Function
        End If
        
        If IsEmpty(var) And IsEmpty(TheValue) Then
            Let MemberQ = True
            Exit Function
        End If
        
        If IsNull(var) And IsNull(TheValue) Then
            Let MemberQ = True
            Exit Function
        End If
        
        If IsObject(var) Then
            Let MemberQ = TheValue Is var
            Exit Function
        End If

        If VarType(var) = VarType(TheValue) And var = TheValue Then
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

' DESCRIPTION
' Boolean function returning True if the given parameter is 0.
'
' PARAMETERS
' 1. arg - any Excel expression
'
' RETURNED VALUE
' Returns True or False depending on whether or not its argument is zero
Public Function ZeroQ(arg As Variant) As Boolean
    If NumberQ(arg) Then
        Let ZeroQ = arg = 0
    Else
        Let ZeroQ = False
    End If
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
Public Function ZeroArrayQ(arg As Variant) As Boolean
    Let ZeroArrayQ = AllTrueQ(AnArray, ThisWorkbook, "ZeroQ")
End Function

' DESCRIPTION
' Boolean function returning True if the given parameter is 1.
'
' PARAMETERS
' 1. arg - any Excel expression
'
' RETURNED VALUE
' Returns True or False depending on whether or not its argument is 1
Public Function OneQ(arg As Variant) As Boolean
    If NumberQ(arg) Then
        Let OneQ = arg = 1
    Else
        Let OneQ = False
    End If
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
Public Function OneArrayQ(arg As Variant) As Boolean
    Let OneArrayQ = AllTrueQ(AnArray, ThisWorkbook, "OneQ")
End Function

' DESCRIPTION
' Boolean function returning True if its parameter any of the following forms:
'
' PARAMETERS
' 1. arg - Expressions with of the following forms:
'
'        1. n - to get element n.  If given a 2D array, n refers to the row number
'        2. AnArray() - A non-empty array of valid indices
'        3. Span - An instance of class Span, which can be conveniently generated using
'                  ClassConstructors.Span()
'
' RETURNED VALUE
' Returns True or False depending on whether or not the given parameter has one of the acceptable forms
Public Function PartIndexQ(arg As Variant) As Boolean
    Let PartIndexQ = NonzeroWholeNumberQ(arg) Or _
                     (NonEmptyArrayQ(arg) And NonzeroWholeNumberArrayQ(arg)) Or _
                     SpanQ(arg)
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
Public Function PartIndexArrayQ(arg As Variant) As Boolean
    Let PartIndexArrayQ = AllTrueQ(arg, ThisWorkbook, "PartIndexQ")
End Function

' DESCRIPTION
' Boolean function returning True if its argument is an instance of class Span.
'
' PARAMETERS
' 1. arg - any Excel value or reference
'
' RETURNED VALUE
' Returns True or False depending on whether or not its argument is an instance of class Span
Public Function SpanQ(arg As Variant) As Boolean
    Let SpanQ = TypeName(arg) = TypeName(ClassConstructors.Span(1, 1))
End Function

' DESCRIPTION
' Boolean function returning True if its argument is an array of class Span instances
'
' PARAMETERS
' 1. arg - any Excel value or reference
'
' RETURNED VALUE
' Returns True or False depending on whether or not its argument is an array of class Soab instances
Public Function SpanArrayQ(arg As Variant) As Boolean
    Let SpanArrayQ = AllTrueQ(arg, ThisWorkbook, "SpanQ")
End Function

' DESCRIPTION
' Boolean function returning True if its parameter any of the following forms:
'
' PARAMETERS
' 1. TheIndex - Any Excel expression with of the following forms:
'
'        1. n - with n nonzero
'        2. [{n_1, n_2}] - with n_i nonzero
'
' RETURNED VALUE
' Returns True or False depending on whether or not the given parameter has one of the acceptable forms
Public Function TakeIndexQ(TheIndex As Variant) As Boolean
    Let TakeIndexQ = NonzeroWholeNumberQ(TheIndex) Or (NonzeroWholeNumberArrayQ(TheIndex) And Length(TheIndex) = 2)
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
Public Function TakeIndexArrayQ(arg As Variant) As Boolean
    Let TakeIndexArrayQ = AllTrueQ(arg, ThisWorkbook, "TakeIndexQ")
End Function
