Attribute VB_Name = "TestedLibrary"
Option Explicit
Option Base 1

'********************************************************************************************
' Miscellaneous VBA
'********************************************************************************************
Public Sub UsingErrors()
    Debug.Print Error(2001)
    Debug.Print CInt(CVErr(2001))
    Debug.Print IsError(CVErr(2001))
    Debug.Print Error(CInt(CVErr(2001)))
End Sub

Public Sub TestArrayParamUsage()
    Debug.Print "Call VarParamFunction()"
    Call VarParamFunction
    Debug.Print
    Debug.Print "Call VarParamFunction(0)"
    Call VarParamFunction(0)
    Debug.Print
    Debug.Print "Call VarParamFunction(0, 1)"
    Call VarParamFunction(0, 1)
    Debug.Print
    Debug.Print "Call VarParamFunction(0, 1, 2)"
    Call VarParamFunction(0, 1, 2)
End Sub

' Helper for TestArrayParamUsage() in this module
Private Sub VarParamFunction(ParamArray Args() As Variant)
    Debug.Print "Args() is missing is " & IsMissing(Args)
    Debug.Print "Args() is dimensioned is " & DimensionedQ(CopyParamArray(Args))
    Debug.Print "Was passed " & Length(CopyParamArray(Args)) & " arguments."
End Sub

Public Sub TestForEachOnUnDimensionedArray()
    Dim AnUndimensionedArray() As Variant
    Dim var As Variant
    
    On Error GoTo ErrorHandler
    
    Debug.Print "Testing for each on an empty array"
    For Each var In EmptyArray()
        Debug.Print "Did one iteration."
    Next
    
    ' This should raise anerror
    Debug.Print "Testing for each on an undimensioned array"
    For Each var In AnUndimensionedArray
        Debug.Print "Did one iteration."
    Next
    
    Exit Sub

ErrorHandler:
    Debug.Print "An error was raised."
End Sub

Public Sub TestNumberOfDimensionsForUnDimensionedArray()
    Dim AnArray1() As Variant
    Dim AnArray2(1 To 2) As Variant
    
    Debug.Print "The number of dimensions of AnArray1() is " & NumberOfDimensions(AnArray1)
    Debug.Print "The number of dimensions of AnArray2(1 to 2) is " & NumberOfDimensions(AnArray2)
End Sub

'********************************************************************************************
' Predicates
'********************************************************************************************

Public Sub TestPredicatesPrintableQ()
    Dim anInteger As Integer
    Dim aDouble As Double
    Dim aDate As Date
    Dim aBoolean As Boolean
    Dim aString As String
    Dim aWorksheet As Worksheet
    Dim aWorkbook As Workbook
    Dim aListObject As ListObject
    Dim aDictionary As Dictionary
    Dim ANumericArray(1 To 2) As Integer
    Dim var1 As Variant
    Dim var2 As Variant
    
    Set aWorksheet = ActiveSheet
    Set aWorkbook = ThisWorkbook
    Set aListObject = AddListObject(TempComputation.Range("A1"))
    Set aDictionary = New Dictionary
    
    Let ANumericArray(1) = 1
    Let ANumericArray(2) = 1
    
    For Each var2 In Array(Array("aDate", aDate), _
                           Array("1", 1), _
                           Array("-1", -1), _
                           Array("1.1", 1.1), _
                           Array("-1.1", -1.1), _
                           Array("0", 0), _
                           Array("aDouble", aDouble), _
                           Array("aBoolean", aBoolean), _
                           Array("aString", aString), _
                           Array("CVErr(1)", CVErr(1)), _
                           Array("aWorksheet", aWorksheet), _
                           Array("aWorkbook", aWorkbook), _
                           Array("aListObject", aListObject), _
                           Array("ANumericArray", ANumericArray), _
                           Array("aDictionary", aDictionary), _
                           Array("Empty", Empty), _
                           Array("Null", Null))
        Debug.Print "PrintableQ(" & First(var2) & ") = " & PrintableQ(Last(var2))
    Next

    Call aListObject.Delete
End Sub

Public Sub TestAtomicPredicates()
    Dim anInteger As Integer
    Dim aDouble As Double
    Dim aDate As Date
    Dim aBoolean As Boolean
    Dim aString As String
    Dim aWorksheet As Worksheet
    Dim aWorkbook As Workbook
    Dim aListObject As ListObject
    Dim aDictionary As Dictionary
    Dim aVariant As Variant
    Dim AnArray(1 To 2) As Integer
    Dim var1 As Variant
    Dim var2 As Variant
    
    Set aWorksheet = ActiveSheet
    Set aWorkbook = ThisWorkbook
    Set aListObject = AddListObject(TempComputation.Range("A1"))
    Set aDictionary = New Dictionary
    Let aVariant = 1
    
    For Each var1 In GetAtomicTypePredicateNames()
        Debug.Print "Testing " & var1
    
        For Each var2 In Array(Array("aDate", aDate), _
                               Array("1", 1), _
                               Array("-1", -1), _
                               Array("1.1", 1.1), _
                               Array("-1.1", -1.1), _
                               Array("0", 0), _
                               Array("aDouble", aDouble), _
                               Array("aBoolean", aBoolean), _
                               Array("aString", aString), _
                               Array("CVErr(1)", CVErr(1)), _
                               Array("aWorksheet", aWorksheet), _
                               Array("aWorkbook", aWorkbook), _
                               Array("aListObject", aListObject), _
                               Array("aVariant", aVariant), _
                               Array("AnArray", AnArray), _
                               Array("aDictionary", aDictionary), _
                               Array("Empty", Empty), _
                               Array("Null", Null))
            Debug.Print "    - " & var1 & "(" & First(var2) & ") = " & Run(var1, Last(var2))
        Next
        
        Debug.Assert (1 = 2) ' Using this to pause between tests
        
        Debug.Print
    Next
    
    Debug.Assert ZeroQ(0) = True
    Debug.Assert NonZeroQ(0) = False
    Debug.Assert OneQ(1) = True
    Debug.Assert NonOneQ("a") = False
    Debug.Assert NonOneQ(2) = True
    Debug.Assert NumberOrStringQ(1) = True
    Debug.Assert NumberOrStringQ("a") = True
    Debug.Assert NumberOrStringQ(ThisWorkbook) = False
    Debug.Assert NumberOrStringQ(#3/4/1980#) = False

    Call aListObject.Delete
End Sub

Public Sub TestPredicatesAtomicQ()
    Dim aListObject As ListObject
    Dim aDictionary As Dictionary
    Dim AnArray(1 To 2) As Integer
    Dim var As Variant
    
    Set aListObject = AddListObject(TempComputation.Range("A1"))
    Set aDictionary = New Dictionary
    
    For Each var In Array(Array("#3/18/1999#", #3/18/1999#), _
                          Array("1", 1), _
                          Array("-1", -1), _
                          Array("1.1", 1.1), _
                          Array("-1.1", -1.1), _
                          Array("0", 0), _
                          Array("2.5", 2.5), _
                          Array("True", True), _
                          Array("""Test""", "Test"), _
                          Array("CVErr(1)", CVErr(1)), _
                          Array("aWorksheet", ActiveSheet), _
                          Array("aWorkbook", ThisWorkbook), _
                          Array("aListObject", aListObject), _
                          Array("AnArray", AnArray), _
                          Array("aDictionary", aDictionary), _
                          Array("Empty", Empty), _
                          Array("Null", Null), _
                          Array("Array()", Array()))
        Debug.Print "AtomicQ(" & First(var) & ") = " & AtomicQ(Last(var))
    Next

    Call aListObject.Delete
End Sub

Public Sub TestPredicatesSpanQ()
    Dim aListObject As ListObject
    Dim aDictionary As Dictionary
    Dim AnArray(1 To 2) As Integer
    Dim var As Variant
    
    Set aListObject = AddListObject(TempComputation.Range("A1"))
    Set aDictionary = New Dictionary
    
    Debug.Assert SpanQ(ClassConstructors.Span(1, 1))
    
    For Each var In Array(Array("#3/18/1999#", #3/18/1999#), _
                          Array("1", 1), _
                          Array("-1", -1), _
                          Array("1.1", 1.1), _
                          Array("-1.1", -1.1), _
                          Array("0", 0), _
                          Array("2.5", 2.5), _
                          Array("True", True), _
                          Array("""Test""", "Test"), _
                          Array("CVErr(1)", CVErr(1)), _
                          Array("aWorksheet", ActiveSheet), _
                          Array("aWorkbook", ThisWorkbook), _
                          Array("aListObject", aListObject), _
                          Array("AnArray", AnArray), _
                          Array("aDictionary", aDictionary), _
                          Array("Empty", Empty), _
                          Array("Null", Null), _
                          Array("Array()", Array()))
        Debug.Assert Not SpanQ(Last(var))
    Next

    Call aListObject.Delete
End Sub

'********************************************************************************************
' Functional Predicates
'********************************************************************************************

Public Sub TestFunctionalPredicatesAllTrueQ()
    Dim UndimensionedArray() As Variant

    Debug.Assert AllTrueQ(Array(1, 2, 4), "WholeNumberQ")
    Debug.Assert AllTrueQ(Array(1, 2, 4#), "WholeNumberQ")
    Debug.Assert AllTrueQ(Array(1, 2, 4), "NumberQ")
    Debug.Assert AllTrueQ(Array(1, 2, 4#), "NumberQ")
    Debug.Assert Not AllTrueQ(Array(1, 2, 4#), "StringQ")
    Debug.Assert Not AllTrueQ(Array("a", "b", Empty), "StringQ")
    Debug.Assert AllTrueQ(Array("a", "b"), "StringQ")
    Debug.Assert AllTrueQ(EmptyArray(), "StringQ")
    Debug.Assert Not AllTrueQ(UndimensionedArray, "StringQ")
    Debug.Assert AllTrueQ(Array(1.1, 2#, 4.2), "NumberOrStringQ")
    Debug.Assert AllTrueQ(Array(#1/1/2000#), "DateQ")
    Debug.Assert Not AllTrueQ(Array(#1/1/2000#, 1), "DateQ")
    Debug.Assert AllTrueQ(Array(ThisWorkbook), "WorkbookQ")
    Debug.Assert Not AllTrueQ(Array(ThisWorkbook.Worksheets(1)), "WorkbookQ")
    Debug.Assert AllTrueQ(Array(ThisWorkbook.Worksheets(1)), "WorksheetQ")
    Debug.Assert AllTrueQ(Array([{1,2;3,4}], [{1;2}]), "MatrixQ")
    Debug.Assert Not AllTrueQ(Array([{1,2;3,4}], Array(Array(1, 2), Array(3, 5))), "MatrixQ")
    Debug.Assert AllTrueQ(Array(True, True, True))
    Debug.Assert Not AllTrueQ(Array(True, True, False))
End Sub

Public Sub TestFunctionalPredicatesAnyTrueQ()
    Dim UndimensionedArray() As Variant
    Dim var As Variant
    
    Let var = TransposeMatrix(Array(Array(1, 2), Array(3, 4)))

    Debug.Assert AnyTrueQ(Array(1, 2, 4), "WholeNumberQ")
    Debug.Assert AnyTrueQ(Array(1, 2, 4#), "WholeNumberQ")
    Debug.Assert AnyTrueQ(Array(1, 2, 4), "NumberQ")
    Debug.Assert AnyTrueQ(Array(1, 2, 4#), "NumberQ")
    Debug.Assert Not AnyTrueQ(Array(1, 2, 4#), "StringQ")
    Debug.Assert AnyTrueQ(Array("a", "b", Empty), "StringQ")
    Debug.Assert AnyTrueQ(Array("a", "b"), "StringQ")
    Debug.Assert Not AnyTrueQ(EmptyArray(), "StringQ")
    Debug.Assert AnyTrueQ(Array(EmptyArray(), "a"), "StringQ")
    Debug.Assert Not AnyTrueQ(UndimensionedArray, "StringQ")
    Debug.Assert AnyTrueQ(Array(True, True, True))
    Debug.Assert AnyTrueQ(Array(True, True, False))
End Sub

Public Sub TestFunctionalPredicatesNoneTrueQ()
    Dim UndimensionedArray() As Variant

    Debug.Assert Not NoneTrueQ(Array(1, 2, 4), "WholeNumberQ")
    Debug.Assert Not NoneTrueQ(Array(1, 2, 4#), "WholeNumberQ")
    Debug.Assert Not NoneTrueQ(Array(1, 2, 4), "NumberQ")
    Debug.Assert Not NoneTrueQ(Array(1, 2, 4#), "NumberQ")
    Debug.Assert NoneTrueQ(Array(1, 2, 4#), "StringQ")
    Debug.Assert Not NoneTrueQ(Array("a", "b", Empty), "StringQ")
    Debug.Assert Not NoneTrueQ(Array("a", "b"), "StringQ")
    Debug.Assert NoneTrueQ(EmptyArray(), "StringQ")
    Debug.Assert Not NoneTrueQ(UndimensionedArray, "StringQ")
    Debug.Assert Not NoneTrueQ(UndimensionedArray, "StringQ")
    Debug.Assert Not NoneTrueQ(Array(True, True, True))
    Debug.Assert Not NoneTrueQ(Array(True, True, False))
    Debug.Assert NoneTrueQ(Array(False, False, False))
End Sub

Public Sub TestFunctionalPredicatesAllFalseQ()
    Dim UndimensionedArray() As Variant

    Debug.Assert Not AllFalseQ(Array(1, 2, 4), "WholeNumberQ")
    Debug.Assert Not AllFalseQ(Array(1, 2, 4#), "WholeNumberQ")
    Debug.Assert Not AllFalseQ(Array(1, 2, 4), "NumberQ")
    Debug.Assert Not AllFalseQ(Array(1, 2, 4#), "NumberQ")
    Debug.Assert AllFalseQ(Array(1, 2, 4#), "StringQ")
    Debug.Assert Not AllFalseQ(Array("a", "b", Empty), "StringQ")
    Debug.Assert Not AllFalseQ(Array("a", "b"), "StringQ")
    Debug.Assert AllFalseQ(EmptyArray(), "StringQ")
    Debug.Assert Not AllFalseQ(UndimensionedArray, "StringQ")
    Debug.Assert Not AllFalseQ(Array(1.1, 2#, 4.2), "NumberOrStringQ")
    Debug.Assert Not AllFalseQ(Array(#1/1/2000#), "DateQ")
    Debug.Assert Not AllFalseQ(Array(#1/1/2000#, 1), "DateQ")
    Debug.Assert Not AllFalseQ(Array(ThisWorkbook), "WorkbookQ")
    Debug.Assert AllFalseQ(Array(ThisWorkbook.Worksheets(1)), "WorkbookQ")
    Debug.Assert Not AllFalseQ(Array(ThisWorkbook.Worksheets(1)), "WorksheetQ")
    Debug.Assert Not AllFalseQ(Array([{1,2;3,4}], [{1;2}]), "MatrixQ")
    Debug.Assert Not AllFalseQ(Array([{1,2;3,4}], Array(Array(1, 2), Array(3, 5))), "MatrixQ")
    Debug.Assert AllFalseQ(Array(Array(1, 2), Array(3, 5)), "MatrixQ")
    Debug.Assert Not AllFalseQ(Array(True, True, True))
    Debug.Assert Not AllFalseQ(Array(True, True, False))
    Debug.Assert AllFalseQ(Array(False, False, False))
    Debug.Assert Not AllFalseQ(Empty)
End Sub

Public Sub TestPredicatesAnyFalseQ()
    Dim UndimensionedArray() As Variant
    Dim var As Variant
    
    Let var = TransposeMatrix(Array(Array(1, 2), Array(3, 4)))

    Debug.Assert Not AnyFalseQ(Array(1, 2, 4), "WholeNumberQ")
    Debug.Assert Not AnyFalseQ(Array(1, 2, 4#), "WholeNumberQ")
    Debug.Assert Not AnyFalseQ(Array(1, 2, 4), "NumberQ")
    Debug.Assert Not AnyFalseQ(Array(1, 2, 4#), "NumberQ")
    Debug.Assert AnyFalseQ(Array(1, 2, 4#), "StringQ")
    Debug.Assert AnyFalseQ(Array("a", "b", Empty), "StringQ")
    Debug.Assert Not AnyFalseQ(Array("a", "b"), "StringQ")
    Debug.Assert Not AnyFalseQ(EmptyArray(), "StringQ")
    Debug.Assert AnyFalseQ(Array(EmptyArray(), "a"), "StringQ")
    Debug.Assert Not AnyFalseQ(UndimensionedArray, "StringQ")
    Debug.Assert Not AnyFalseQ(Array(True, True, True))
    Debug.Assert AnyFalseQ(Array(True, True, False))
End Sub

Public Sub TestPredicatesNoneFalseQ()
    Dim UndimensionedArray() As Variant

    Debug.Assert Not NoneFalseQ(Array(1, 2, 4), "WholeNumberQ")
    Debug.Assert Not NoneFalseQ(Array(1, 2, 4#), "WholeNumberQ")
    Debug.Assert Not NoneFalseQ(Array(1, 2, 4), "NumberQ")
    Debug.Assert Not NoneFalseQ(Array(1, 2, 4#), "NumberQ")
    Debug.Assert NoneFalseQ(Array(1, 2, 4#), "StringQ")
    Debug.Assert Not NoneFalseQ(Array("a", "b", Empty), "StringQ")
    Debug.Assert Not NoneFalseQ(Array("a", "b"), "StringQ")
    Debug.Assert NoneFalseQ(EmptyArray(), "StringQ")
    Debug.Assert Not NoneFalseQ(UndimensionedArray, "StringQ")
    Debug.Assert Not NoneFalseQ(UndimensionedArray, "StringQ")
    Debug.Assert Not NoneFalseQ(Array(True, True, True))
    Debug.Assert Not NoneFalseQ(Array(True, True, False))
    Debug.Assert NoneFalseQ(Array(False, False, False))
End Sub

Public Sub TestPredicatesDimensionedQ()
    Dim A() As Variant
    Dim B(1 To 2) As Variant
    Dim c As Integer
    Dim wbk As Workbook
    
    Debug.Assert DimensionedQ(EmptyArray())
    Debug.Assert Not DimensionedQ(A)
    Debug.Assert DimensionedQ(B)
    Debug.Assert Not DimensionedQ(c)
    Debug.Assert Not DimensionedQ(wbk)
End Sub

Public Sub TestPredicatesEmptyArrayQ()
    Dim A() As Variant
    Dim B(1 To 2) As Variant
    Dim c As Integer
    Dim wbk As Workbook
    
    Debug.Assert EmptyArrayQ(EmptyArray())
    Debug.Assert Not EmptyArrayQ(A)
    Debug.Assert Not EmptyArrayQ(B)
    Debug.Assert Not EmptyArrayQ(c)
    Debug.Assert Not EmptyArrayQ(wbk)
End Sub

Public Sub TestPredicatesAtomicArrayQ()
    Dim anInteger As Integer
    Dim aDouble As Double
    Dim aDate As Date
    Dim aBoolean As Boolean
    Dim aString As String
    Dim aWorksheet As Worksheet
    Dim aWorkbook As Workbook
    Dim aListObject As ListObject
    Dim aVariant As Variant
    Dim AnArray(1 To 2) As Integer
    
    Set aWorksheet = ActiveSheet
    Set aWorkbook = ThisWorkbook
    Set aListObject = AddListObject(TempComputation.Range("A1"))
    
    Debug.Assert AtomicArrayQ(Array(anInteger, aDouble))
    Debug.Assert AtomicArrayQ(Array(aDate, aString))
    Debug.Assert Not AtomicArrayQ(Array(EmptyArray(), 1))
    Debug.Assert AtomicArrayQ(EmptyArray())
    Debug.Assert AtomicArrayQ(Array(Null, Empty))
    Debug.Assert Not AtomicArrayQ(Array(Nothing, 1))
    Debug.Assert Not AtomicArrayQ(1)
    Debug.Assert AtomicArrayQ([{1,2;3,4}])
    Debug.Assert Not AtomicArrayQ(Array(Array(1, 2), 2))
    Debug.Assert Not AtomicArrayQ(Null)

    Call aListObject.Delete
End Sub

Public Sub TestPredicatesFileExistsQ()
    Debug.Assert FileExistsQ(ThisWorkbook.Path & "\ExcelLibraryV6.0.xlsb")
    Debug.Assert FileExistsQ(ThisWorkbook.Path & "\ExcelLibraryV6.2.xlsb")
    Debug.Assert Not FileExistsQ(ThisWorkbook.Path & "\ExcelLibraryV-1.xlsb")
End Sub

Public Sub TestPredicatesListObjectExistsQ()
    Dim aListObject As ListObject
    
    Set aListObject = AddListObject(TempComputation.Range("A1"), "MyListObject")
   
    Debug.Assert ListObjectExistsQ(TempComputation, "MyListObject")
    Debug.Assert Not ListObjectExistsQ(TempComputation, "ASecondListObject")
    
    Call aListObject.Delete
End Sub

Public Sub TestPredicatesWorksheetExistsQ()
    Debug.Assert WorksheetExistsQ(ThisWorkbook, "TempComputation")
    Debug.Assert Not WorksheetExistsQ(ThisWorkbook, "NoneExistingWorksheet")
End Sub

Public Sub TestPredicatesSheetExistsQ()
    Debug.Assert SheetExistsQ(ThisWorkbook, "TempComputation")
    Debug.Assert Not SheetExistsQ(ThisWorkbook, "NoneExistingWorksheet")
End Sub

'********************************************************************************************
' Array Predicates
'********************************************************************************************
Public Sub TestArrayPredicatesMatrixQ()
    Dim M1 As Variant
    Dim M2(1 To 2, 1 To 3) As Double
    Dim i As Integer, j As Integer
    Dim var As Variant
    
    Let M1 = [{1,2;3,4}]
    For i = LBound(M2, 1) To UBound(M2, 1)
        For j = LBound(M2, 2) To UBound(M2, 2)
            Let M2(i, j) = i * 10 + j
        Next
    Next
    
    Debug.Assert MatrixQ(M1)
    Debug.Assert MatrixQ(M2)
    Debug.Assert Not MatrixQ([{1,"a";2,3}])
    
    For Each var In Array(1, 2, "a", ThisWorkbook, TempComputation)
        Debug.Assert Not MatrixQ(var)
    Next
End Sub


Public Sub TestArrayPredicatesNumberArrayQ()
    Dim a1(1 To 2, 1 To 2, 1 To 2) As Variant
    Dim i As Integer, j As Integer, k As Integer
    Dim var As Variant
    
    For i = 1 To 2
        For j = 1 To 2
            For k = 1 To 2
                Let a1(i, j, k) = i * 10 ^ 2 + j * 10 + k
            Next k
        Next j
    Next i
    
    Debug.Assert NumberArrayQ(a1)
    
    Let a1(1, 1, 1) = "A"
    Debug.Assert Not NumberArrayQ(a1)
    
    Debug.Assert NumberArrayQ(EmptyArray)
    Debug.Assert NumberArrayQ(Array(1, 2, 3, 4))
    Debug.Assert Not NumberArrayQ(Array("a"))
    Debug.Assert Not NumberArrayQ(Array("a", 1))
    
    For Each var In Array(Empty, Null, CVErr(1), ThisWorkbook, TempComputation)
        Debug.Assert Not NumberArrayQ(var)
    Next
End Sub

Public Sub TestArrayPredicatesNegativeWholeNumberArrayQ()
    Dim a1(1 To 2, 1 To 2, 1 To 2) As Variant
    Dim i As Integer, j As Integer, k As Integer
    Dim var As Variant
    
    For i = 1 To 2
        For j = 1 To 2
            For k = 1 To 2
                Let a1(i, j, k) = -i * 10 ^ 2 - j * 10 - k
            Next k
        Next j
    Next i
    
    Debug.Assert NegativeWholeNumberArrayQ(a1)
    
    Let a1(1, 1, 1) = "A"
    Debug.Assert Not NegativeWholeNumberArrayQ(a1)
    
    Debug.Assert NegativeWholeNumberArrayQ(EmptyArray)
    Debug.Assert NegativeWholeNumberArrayQ(Array(-1, -2, -3, -4))
    Debug.Assert Not NegativeWholeNumberArrayQ(Array(-1, -2.5, -3, -4))
    Debug.Assert Not NegativeWholeNumberArrayQ(Array(1, -2, -3, -4))
    Debug.Assert Not NegativeWholeNumberArrayQ(Array(0, 1, -2, -3, -4))
    Debug.Assert Not NegativeWholeNumberArrayQ(Array("a"))
    Debug.Assert Not NegativeWholeNumberArrayQ(Array("a", 1))
    
    For Each var In Array(Empty, Null, CVErr(1), ThisWorkbook, TempComputation)
        Debug.Assert Not NegativeWholeNumberArrayQ(var)
    Next
End Sub

Public Sub TestArrayPredicatesNonNegativeWholeNumberArrayQ()
    Dim a1(1 To 2, 1 To 2, 1 To 2) As Variant
    Dim i As Integer, j As Integer, k As Integer
    Dim var As Variant
    
    For i = 1 To 2
        For j = 1 To 2
            For k = 1 To 2
                Let a1(i, j, k) = i * 10 ^ 2 + j * 10 + k - 1
            Next k
        Next j
    Next i
    
    Debug.Assert NonNegativeWholeNumberArrayQ(a1)
    
    Let a1(1, 1, 1) = "A"
    Debug.Assert Not NonNegativeWholeNumberArrayQ(a1)
    
    Debug.Assert NonNegativeWholeNumberArrayQ(EmptyArray)
    Debug.Assert NonNegativeWholeNumberArrayQ(Array(0, 1, 2, 3, 4))
    Debug.Assert Not NonNegativeWholeNumberArrayQ(Array(0, 1, 2.5, 3, 4))
    Debug.Assert Not NonNegativeWholeNumberArrayQ(Array(1, -2, -3, -4))
    Debug.Assert Not NonNegativeWholeNumberArrayQ(Array("a"))
    Debug.Assert Not NonNegativeWholeNumberArrayQ(Array("a", 1))
    
    For Each var In Array(Empty, Null, CVErr(1), ThisWorkbook, TempComputation)
        Debug.Assert Not NonNegativeWholeNumberArrayQ(var)
    Next
End Sub

Public Sub TestArrayPredicatesNonzeroWholeNumberArrayQ()
    Dim a1(1 To 2, 1 To 2, 1 To 2) As Variant
    Dim i As Integer, j As Integer, k As Integer
    Dim var As Variant
    
    For i = 1 To 2
        For j = 1 To 2
            For k = 1 To 2
                Let a1(i, j, k) = i * 10 ^ 2 + j * 10 + k - 1
            Next k
        Next j
    Next i
    
    Debug.Assert NonzeroWholeNumberArrayQ(a1)
    
    Let a1(1, 1, 1) = "A"
    Debug.Assert Not NonzeroWholeNumberArrayQ(a1)
    
    Debug.Assert NonzeroWholeNumberArrayQ(EmptyArray)
    Debug.Assert Not NonzeroWholeNumberArrayQ(Array(0, 1, 2, 3, 4))
    Debug.Assert NonzeroWholeNumberArrayQ(Array(1, 2, 3, 4))
    Debug.Assert Not NonzeroWholeNumberArrayQ(Array(0, 1, 2.5, 3, 4))
    Debug.Assert NonzeroWholeNumberArrayQ(Array(1, -2, -3, -4))
    Debug.Assert Not NonzeroWholeNumberArrayQ(Array("a"))
    Debug.Assert Not NonzeroWholeNumberArrayQ(Array("a", 1))
    
    For Each var In Array(Empty, Null, CVErr(1), ThisWorkbook, TempComputation)
        Debug.Assert Not NonzeroWholeNumberArrayQ(var)
    Next
End Sub

Public Sub TestArrayPredicatesWholeNumberOrStringArrayQ()
    Dim a1(1 To 2, 1 To 2, 1 To 2) As Variant
    Dim i As Integer, j As Integer, k As Integer
    Dim var As Variant
    
    For i = 1 To 2
        For j = 1 To 2
            For k = 1 To 2
                Let a1(i, j, k) = i * 10 ^ 2 + j * 10 + k - 1
            Next k
        Next j
    Next i
    
    Debug.Assert WholeNumberOrStringArrayQ(a1)
    
    Let a1(1, 1, 1) = "A"
    Debug.Assert WholeNumberOrStringArrayQ(a1)
    
    Debug.Assert WholeNumberOrStringArrayQ(EmptyArray)
    Debug.Assert WholeNumberOrStringArrayQ(Array(0, 1, 2, 3, 4))
    Debug.Assert WholeNumberOrStringArrayQ(Array("a", "b"))
    Debug.Assert WholeNumberOrStringArrayQ(Array(1, 2, 3, 4))
    Debug.Assert Not WholeNumberOrStringArrayQ(Array(0, 1, 2.5, 3, 4))
    Debug.Assert WholeNumberOrStringArrayQ(Array(1, -2, -3, -4))
    Debug.Assert WholeNumberOrStringArrayQ(Array("a"))
    Debug.Assert WholeNumberOrStringArrayQ(Array("a", 1))
    Debug.Assert Not WholeNumberOrStringArrayQ(Array("a", 1.1))
    Debug.Assert Not WholeNumberOrStringArrayQ(Array("a", -1.2))
    
    For Each var In Array(Empty, Null, CVErr(1), ThisWorkbook, TempComputation)
        Debug.Assert Not NonzeroWholeNumberArrayQ(var)
    Next
End Sub

Public Sub TestArrayPredicatesNumberOrStringArrayQ()
    Dim a1(1 To 2, 1 To 2, 1 To 2) As Variant
    Dim i As Integer, j As Integer, k As Integer
    Dim var As Variant
    
    For i = 1 To 2
        For j = 1 To 2
            For k = 1 To 2
                Let a1(i, j, k) = i * 10 ^ 2 + j * 10 + k - 1
            Next k
        Next j
    Next i
    
    Debug.Assert NumberOrStringArrayQ(a1)
    
    Let a1(1, 1, 1) = "A"
    Debug.Assert NumberOrStringArrayQ(a1)
    
    Debug.Assert NumberOrStringArrayQ(EmptyArray)
    Debug.Assert NumberOrStringArrayQ(Array(0, 1, 2, 3, 4))
    Debug.Assert NumberOrStringArrayQ(Array("a", "b"))
    Debug.Assert NumberOrStringArrayQ(Array(1, 2, 3, 4))
    Debug.Assert NumberOrStringArrayQ(Array(0, 1, 2.5, 3, 4))
    Debug.Assert NumberOrStringArrayQ(Array(1, -2, -3, -4))
    Debug.Assert NumberOrStringArrayQ(Array("a"))
    Debug.Assert NumberOrStringArrayQ(Array("a", 1))
    Debug.Assert NumberOrStringArrayQ(Array("a", 1.1))
    Debug.Assert NumberOrStringArrayQ(Array("a", -1.2))
    Debug.Assert Not NumberOrStringArrayQ(Array(ThisWorkbook, 1))
    
    For Each var In Array(Empty, Null, CVErr(1), ThisWorkbook, TempComputation)
        Debug.Assert Not NumberOrStringArrayQ(var)
    Next
End Sub

'********************************************************************************************
' FunctionalProgramming
'********************************************************************************************

