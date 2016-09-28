Attribute VB_Name = "LibraryTesting"
Option Explicit
Option Base 1

'********************************************************************************************
' Miscellaneous VBA
'********************************************************************************************
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

Public Sub TestArrayPredicatesWorkbookArrayQ()
    Debug.Assert WorkbookArrayQ(EmptyArray)
    Debug.Assert WorkbookArrayQ(Array(ThisWorkbook))
    Debug.Assert WorkbookArrayQ(Array(ThisWorkbook, ThisWorkbook))
    Debug.Assert Not WorkbookArrayQ(Array(0, 1, 2, 3, 4))
    Debug.Assert Not WorkbookArrayQ(Array("a", "b"))
    Debug.Assert Not WorkbookArrayQ(Array(1, 2, 3, 4))
    Debug.Assert Not WorkbookArrayQ(Array(0, 1, 2.5, 3, 4))
    Debug.Assert Not WorkbookArrayQ(Array(1, -2, -3, -4))
    Debug.Assert Not WorkbookArrayQ(Array("a"))
    Debug.Assert Not WorkbookArrayQ(Array("a", 1))
    Debug.Assert Not WorkbookArrayQ(Array("a", 1.1))
    Debug.Assert Not WorkbookArrayQ(Array("a", -1.2))
    Debug.Assert Not WorkbookArrayQ(Array(ThisWorkbook, 1))
End Sub

Public Sub TestArrayPredicatesWorksheetArrayQ()
    Dim var As Variant
    Dim aListObject As ListObject
    Dim aDictionary As Dictionary
    
    For Each var In TempComputation.ListObjects
        Call var.Delete
    Next
    
    Set aListObject = AddListObject(TempComputation.Range("A1"))
    Set aDictionary = New Dictionary
    
    For Each var In Array(#1/1/2000#, _
                          1, _
                          -1, _
                          1.1, _
                          -1.1, _
                          0, _
                          2.5, _
                          True, _
                          "asfggd", _
                          CVErr(1), _
                          TempComputation, _
                          ThisWorkbook, _
                          aListObject, _
                          aDictionary, _
                          Empty, _
                          Null)
        Debug.Assert Not WholeNumberArrayQ(var)
    Next

    Debug.Assert WorksheetArrayQ(EmptyArray)
    Debug.Assert WorksheetArrayQ(Array(TempComputation))
    Debug.Assert WorksheetArrayQ(Array(TempComputation, TempComputation))
    Debug.Assert Not WorksheetArrayQ(Array(0, 1, 2, 3, 4))
    Debug.Assert Not WorksheetArrayQ(Array("a", "b"))
    Debug.Assert Not WorksheetArrayQ(Array(1, 2, 3, 4))
    Debug.Assert Not WorksheetArrayQ(Array(0, 1, 2.5, 3, 4))
    Debug.Assert Not WorksheetArrayQ(Array(1, -2, -3, -4))
    Debug.Assert Not WorksheetArrayQ(Array("a"))
    Debug.Assert Not WorksheetArrayQ(Array("a", 1))
    Debug.Assert Not WorksheetArrayQ(Array("a", 1.1))
    Debug.Assert Not WorksheetArrayQ(Array("a", -1.2))
    Debug.Assert Not WorksheetArrayQ(Array(TempComputation, 1))
    
    Call aListObject.Delete
End Sub

Public Sub TestArrayPredicatesWholeNumberArrayQ()
    Dim var As Variant
    Dim aListObject As ListObject
    Dim aDictionary As Dictionary
    
    For Each var In TempComputation.ListObjects
        Call var.Delete
    Next
    
    Set aListObject = AddListObject(TempComputation.Range("A1"))
    Set aDictionary = New Dictionary
    
    For Each var In Array(#1/1/2000#, _
                          1, _
                          -1, _
                          1.1, _
                          -1.1, _
                          0, _
                          2.5, _
                          True, _
                          "asfggd", _
                          CVErr(1), _
                          TempComputation, _
                          ThisWorkbook, _
                          aListObject, _
                          aDictionary, _
                          Empty, _
                          Null)
        Debug.Assert Not WholeNumberArrayQ(var)
    Next

    Debug.Assert WholeNumberArrayQ(Array(1, 2, 3, 4))
    Debug.Assert WholeNumberArrayQ(Array(0, 1, 2, 3, 4))
    Debug.Assert WholeNumberArrayQ(Array(1, -2, -3, -4))
    Debug.Assert WholeNumberArrayQ(EmptyArray)
    Debug.Assert Not WholeNumberArrayQ(Array(0, 1, 2.5, 3, 4))
    Debug.Assert Not WholeNumberArrayQ(Array(TempComputation))
    Debug.Assert Not WholeNumberArrayQ(Array(TempComputation, TempComputation))
    Debug.Assert Not WholeNumberArrayQ(Array("a", "b"))
    Debug.Assert Not WholeNumberArrayQ(Array("a"))
    Debug.Assert Not WholeNumberArrayQ(Array("a", 1))
    Debug.Assert Not WholeNumberArrayQ(Array("a", 1.1))
    Debug.Assert Not WholeNumberArrayQ(Array("a", -1.2))
    Debug.Assert Not WholeNumberArrayQ(Array(TempComputation, 1))
    
    Call aListObject.Delete
End Sub

Public Sub TestArrayPredicatesListObjectArrayQ()
    Dim var As Variant
    Dim aListObject As ListObject
    Dim aDictionary As Dictionary
    
    For Each var In TempComputation.ListObjects
        Call var.Delete
    Next
    
    Set aListObject = AddListObject(TempComputation.Range("A1"))
    Set aDictionary = New Dictionary
    
    For Each var In Array(#1/1/2000#, _
                          1, _
                          -1, _
                          1.1, _
                          -1.1, _
                          0, _
                          2.5, _
                          True, _
                          "asfggd", _
                          CVErr(1), _
                          TempComputation, _
                          ThisWorkbook, _
                          aListObject, _
                          aDictionary, _
                          Empty, _
                          Null)
        Debug.Assert Not ListObjectArrayQ(var)
    Next

    Debug.Assert ListObjectArrayQ(Array(aListObject, aListObject))
    Debug.Assert ListObjectArrayQ(Array(aListObject))
    Debug.Assert ListObjectArrayQ(EmptyArray)
    Debug.Assert Not ListObjectArrayQ(Array(aListObject, ThisWorkbook))
    Debug.Assert Not ListObjectArrayQ(Array(0, 1, 2.5, 3, 4))
    Debug.Assert Not ListObjectArrayQ(Array(TempComputation))
    Debug.Assert Not ListObjectArrayQ(Array(TempComputation, TempComputation))
    Debug.Assert Not ListObjectArrayQ(Array("a", "b"))
    Debug.Assert Not ListObjectArrayQ(Array("a"))
    Debug.Assert Not ListObjectArrayQ(Array("a", 1))
    Debug.Assert Not ListObjectArrayQ(Array("a", 1.1))
    Debug.Assert Not ListObjectArrayQ(Array("a", -1.2))
    Debug.Assert Not ListObjectArrayQ(Array(TempComputation, 1))
    
    Call aListObject.Delete
End Sub

Public Sub TestArrayPredicatesDictionaryArrayQ()
    Dim var As Variant
    Dim aListObject As ListObject
    Dim aDictionary As Dictionary
    
    For Each var In TempComputation.ListObjects
        Call var.Delete
    Next
    
    Set aListObject = AddListObject(TempComputation.Range("A1"))
    Set aDictionary = New Dictionary
    
    For Each var In Array(#1/1/2000#, _
                          1, _
                          -1, _
                          1.1, _
                          -1.1, _
                          0, _
                          2.5, _
                          True, _
                          "asfggd", _
                          CVErr(1), _
                          TempComputation, _
                          ThisWorkbook, _
                          aListObject, _
                          aDictionary, _
                          Empty, _
                          Null)
        Debug.Assert Not DictionaryArrayQ(var)
    Next

    Debug.Assert DictionaryArrayQ(Array(aDictionary, aDictionary))
    Debug.Assert DictionaryArrayQ(Array(aDictionary))
    Debug.Assert DictionaryArrayQ(EmptyArray)
    Debug.Assert Not DictionaryArrayQ(Array(aListObject, ThisWorkbook))
    Debug.Assert Not DictionaryArrayQ(Array(0, 1, 2.5, 3, 4))
    Debug.Assert Not DictionaryArrayQ(Array(TempComputation))
    Debug.Assert Not DictionaryArrayQ(Array(TempComputation, TempComputation))
    Debug.Assert Not DictionaryArrayQ(Array("a", "b"))
    Debug.Assert Not DictionaryArrayQ(Array("a"))
    Debug.Assert Not DictionaryArrayQ(Array("a", 1))
    Debug.Assert Not DictionaryArrayQ(Array("a", 1.1))
    Debug.Assert Not DictionaryArrayQ(Array("a", -1.2))
    Debug.Assert Not DictionaryArrayQ(Array(TempComputation, 1))
    
    Call aListObject.Delete
End Sub

Public Sub TestArrayPredicatesErrorArrayQ()
    Dim var As Variant
    Dim aListObject As ListObject
    Dim aDictionary As Dictionary
    
    For Each var In TempComputation.ListObjects
        Call var.Delete
    Next
    
    Set aListObject = AddListObject(TempComputation.Range("A1"))
    Set aDictionary = New Dictionary
    
    For Each var In Array(#1/1/2000#, _
                          1, _
                          -1, _
                          1.1, _
                          -1.1, _
                          0, _
                          2.5, _
                          True, _
                          "asfggd", _
                          CVErr(1), _
                          TempComputation, _
                          ThisWorkbook, _
                          aListObject, _
                          aDictionary, _
                          Empty, _
                          Null)
        Debug.Assert Not ErrorArrayQ(var)
    Next

    Debug.Assert ErrorArrayQ(Array(Null, Null))
    Debug.Assert ErrorArrayQ(Array(Null))
    Debug.Assert ErrorArrayQ(EmptyArray)
    Debug.Assert Not ErrorArrayQ(Array(aDictionary, aDictionary))
    Debug.Assert Not ErrorArrayQ(Array(aDictionary))
    Debug.Assert Not ErrorArrayQ(Array(aListObject, ThisWorkbook))
    Debug.Assert Not ErrorArrayQ(Array(0, 1, 2.5, 3, 4))
    Debug.Assert Not ErrorArrayQ(Array(TempComputation))
    Debug.Assert Not ErrorArrayQ(Array(TempComputation, TempComputation))
    Debug.Assert Not ErrorArrayQ(Array("a", "b"))
    Debug.Assert Not ErrorArrayQ(Array("a"))
    Debug.Assert Not ErrorArrayQ(Array("a", 1))
    Debug.Assert Not ErrorArrayQ(Array("a", 1.1))
    Debug.Assert Not ErrorArrayQ(Array("a", -1.2))
    Debug.Assert Not ErrorArrayQ(Array(TempComputation, 1))
    
    Call aListObject.Delete
End Sub

Public Sub TestArrayPredicatesDateArrayQ()
    Dim var As Variant
    Dim aListObject As ListObject
    Dim aDictionary As Dictionary
    
    For Each var In TempComputation.ListObjects
        Call var.Delete
    Next
    
    Set aListObject = AddListObject(TempComputation.Range("A1"))
    Set aDictionary = New Dictionary
    
    For Each var In Array(#1/1/2000#, _
                          1, _
                          -1, _
                          1.1, _
                          -1.1, _
                          0, _
                          2.5, _
                          True, _
                          "asfggd", _
                          CVErr(1), _
                          TempComputation, _
                          ThisWorkbook, _
                          aListObject, _
                          aDictionary, _
                          Empty, _
                          Null)
        Debug.Assert Not DateArrayQ(var)
    Next

    Debug.Assert DateArrayQ(Array(#1/1/2000#, #1/1/2000#))
    Debug.Assert DateArrayQ(Array(#1/1/2000#))
    Debug.Assert DateArrayQ(EmptyArray)
    Debug.Assert Not DateArrayQ(Array(aDictionary, aDictionary))
    Debug.Assert Not DateArrayQ(Array(aDictionary))
    Debug.Assert Not DateArrayQ(Array(aListObject, ThisWorkbook))
    Debug.Assert Not DateArrayQ(Array(0, 1, 2.5, 3, 4))
    Debug.Assert Not DateArrayQ(Array(TempComputation))
    Debug.Assert Not DateArrayQ(Array(TempComputation, TempComputation))
    Debug.Assert Not DateArrayQ(Array("a", "b"))
    Debug.Assert Not DateArrayQ(Array("a"))
    Debug.Assert Not DateArrayQ(Array("a", 1))
    Debug.Assert Not DateArrayQ(Array("a", 1.1))
    Debug.Assert Not DateArrayQ(Array("a", -1.2))
    Debug.Assert Not DateArrayQ(Array(TempComputation, 1))
    
    Call aListObject.Delete
End Sub

Public Sub TestArrayPredicatesBooleanArrayQ()
    Dim var As Variant
    Dim aListObject As ListObject
    Dim aDictionary As Dictionary
    
    For Each var In TempComputation.ListObjects
        Call var.Delete
    Next
    
    Set aListObject = AddListObject(TempComputation.Range("A1"))
    Set aDictionary = New Dictionary
    
    For Each var In Array(#1/1/2000#, _
                          1, _
                          -1, _
                          1.1, _
                          -1.1, _
                          0, _
                          2.5, _
                          True, _
                          "asfggd", _
                          CVErr(1), _
                          TempComputation, _
                          ThisWorkbook, _
                          aListObject, _
                          aDictionary, _
                          Empty, _
                          Null)
        Debug.Assert Not BooleanArrayQ(var)
    Next

    Debug.Assert BooleanArrayQ(Array(True, True))
    Debug.Assert BooleanArrayQ(Array(False))
    Debug.Assert BooleanArrayQ(EmptyArray)
    Debug.Assert Not BooleanArrayQ(Array(aDictionary, aDictionary))
    Debug.Assert Not BooleanArrayQ(Array(aDictionary))
    Debug.Assert Not BooleanArrayQ(Array(aListObject, ThisWorkbook))
    Debug.Assert Not BooleanArrayQ(Array(0, 1, 2.5, 3, 4))
    Debug.Assert Not BooleanArrayQ(Array(TempComputation))
    Debug.Assert Not BooleanArrayQ(Array(TempComputation, TempComputation))
    Debug.Assert Not BooleanArrayQ(Array("a", "b"))
    Debug.Assert Not BooleanArrayQ(Array("a"))
    Debug.Assert Not BooleanArrayQ(Array("a", 1))
    Debug.Assert Not BooleanArrayQ(Array("a", 1.1))
    Debug.Assert Not BooleanArrayQ(Array("a", -1.2))
    Debug.Assert Not BooleanArrayQ(Array(TempComputation, 1))
    
    Call aListObject.Delete
End Sub

Public Sub TestArrayPredicatesAtomicTableQ()
    Dim var As Variant
    Dim aListObject As ListObject
    Dim aDictionary As Dictionary
    
    For Each var In TempComputation.ListObjects
        Call var.Delete
    Next
    
    Set aListObject = AddListObject(TempComputation.Range("A1"))
    Set aDictionary = New Dictionary
    
    For Each var In Array(#1/1/2000#, _
                          1, _
                          -1, _
                          1.1, _
                          -1.1, _
                          0, _
                          2.5, _
                          True, _
                          "asfggd", _
                          CVErr(1), _
                          TempComputation, _
                          ThisWorkbook, _
                          aListObject, _
                          aDictionary, _
                          Empty, _
                          Null)
        Debug.Assert Not AtomicTableQ(var)
    Next

    Debug.Assert AtomicTableQ([{1,"a"; True, False}])
    Debug.Assert AtomicTableQ([{1,2,3;4,5,6}])
    Debug.Assert Not AtomicTableQ(EmptyArray)
    Debug.Assert Not AtomicTableQ(Array(aDictionary, aDictionary))
    Debug.Assert Not AtomicTableQ(Array(aDictionary))
    Debug.Assert Not AtomicTableQ(Array(aListObject, ThisWorkbook))
    Debug.Assert Not AtomicTableQ(Array(0, 1, 2.5, 3, 4))
    Debug.Assert Not AtomicTableQ(Array(TempComputation))
    Debug.Assert Not AtomicTableQ(Array(TempComputation, TempComputation))
    Debug.Assert Not AtomicTableQ(Array("a", "b"))
    Debug.Assert Not AtomicTableQ(Array("a"))
    Debug.Assert Not AtomicTableQ(Array("a", 1))
    Debug.Assert Not AtomicTableQ(Array("a", 1.1))
    Debug.Assert Not AtomicTableQ(Array("a", -1.2))
    Debug.Assert Not AtomicTableQ(Array(TempComputation, 1))
    
    Call aListObject.Delete
End Sub

Public Sub TestArrayPredicatesPrintableArrayQ()
    Dim var As Variant
    Dim aListObject As ListObject
    Dim aDictionary As Dictionary
    
    For Each var In TempComputation.ListObjects
        Call var.Delete
    Next
    
    Set aListObject = AddListObject(TempComputation.Range("A1"))
    Set aDictionary = New Dictionary
    
    For Each var In Array(#1/1/2000#, _
                          1, _
                          -1, _
                          1.1, _
                          -1.1, _
                          0, _
                          2.5, _
                          True, _
                          "asfggd", _
                          CVErr(1), _
                          TempComputation, _
                          ThisWorkbook, _
                          aListObject, _
                          aDictionary, _
                          Empty, _
                          Null)
        Debug.Assert Not PrintableArrayQ(var)
    Next

    Debug.Assert PrintableArrayQ(Array(1, 2, "ASFF"))
    Debug.Assert Not PrintableArrayQ([{1,"a"; True, False}])
    Debug.Assert Not PrintableArrayQ([{1,2,3;4,5,6}])
    Debug.Assert PrintableArrayQ(EmptyArray)
    Debug.Assert Not PrintableArrayQ(Array(aDictionary, aDictionary))
    Debug.Assert Not PrintableArrayQ(Array(aDictionary))
    Debug.Assert Not PrintableArrayQ(Array(aListObject, ThisWorkbook))
    Debug.Assert PrintableArrayQ(Array(0, 1, 2.5, 3, 4))
    Debug.Assert Not PrintableArrayQ(Array(TempComputation))
    Debug.Assert Not PrintableArrayQ(Array(TempComputation, TempComputation))
    Debug.Assert PrintableArrayQ(Array("a", "b"))
    Debug.Assert PrintableArrayQ(Array("a"))
    Debug.Assert PrintableArrayQ(Array("a", 1))
    Debug.Assert PrintableArrayQ(Array("a", 1.1))
    Debug.Assert PrintableArrayQ(Array("a", -1.2))
    Debug.Assert Not PrintableArrayQ(Array(TempComputation, 1))
    
    Call aListObject.Delete
End Sub

Public Sub TestArrayPredicatesPrintableTableQ()
    Dim var As Variant
    Dim aListObject As ListObject
    Dim aDictionary As Dictionary
    
    For Each var In TempComputation.ListObjects
        Call var.Delete
    Next
    
    Set aListObject = AddListObject(TempComputation.Range("A1"))
    Set aDictionary = New Dictionary
    
    For Each var In Array(#1/1/2000#, _
                          1, _
                          -1, _
                          1.1, _
                          -1.1, _
                          0, _
                          2.5, _
                          True, _
                          "asfggd", _
                          CVErr(1), _
                          TempComputation, _
                          ThisWorkbook, _
                          aListObject, _
                          aDictionary, _
                          Empty, _
                          Null)
        Debug.Assert Not PrintableTableQ(var)
    Next

    Debug.Assert Not PrintableTableQ(Array(1, 2, "ASFF"))
    Debug.Assert PrintableTableQ([{1,"a"; True, False}])
    Debug.Assert PrintableTableQ([{1,2,3;4,5,6}])
    Debug.Assert Not PrintableTableQ(EmptyArray)
    Debug.Assert Not PrintableTableQ(Array(aDictionary, aDictionary))
    Debug.Assert Not PrintableTableQ(Array(aDictionary))
    Debug.Assert Not PrintableTableQ(Array(aListObject, ThisWorkbook))
    Debug.Assert Not PrintableTableQ(Array(0, 1, 2.5, 3, 4))
    Debug.Assert Not PrintableTableQ(Array(TempComputation))
    Debug.Assert Not PrintableTableQ(Array(TempComputation, TempComputation))
    Debug.Assert Not PrintableTableQ(Array("a", "b"))
    Debug.Assert Not PrintableTableQ(Array("a"))
    Debug.Assert Not PrintableTableQ(Array("a", 1))
    Debug.Assert Not PrintableTableQ(Array("a", 1.1))
    Debug.Assert Not PrintableTableQ(Array("a", -1.2))
    Debug.Assert Not PrintableTableQ(Array(TempComputation, 1))
    
    Call aListObject.Delete
End Sub

Public Sub TestArrayPredicatesZeroArrayQ()
    Dim var As Variant
    Dim aListObject As ListObject
    Dim aDictionary As Dictionary
    
    For Each var In TempComputation.ListObjects
        Call var.Delete
    Next
    
    Set aListObject = AddListObject(TempComputation.Range("A1"))
    Set aDictionary = New Dictionary
    
    For Each var In Array(#1/1/2000#, _
                          1, _
                          -1, _
                          1.1, _
                          -1.1, _
                          0, _
                          2.5, _
                          True, _
                          "asfggd", _
                          CVErr(1), _
                          TempComputation, _
                          ThisWorkbook, _
                          aListObject, _
                          aDictionary, _
                          Empty, _
                          Null)
        Debug.Assert Not ZeroArrayQ(var)
    Next

    Debug.Assert ZeroArrayQ(Array(0))
    Debug.Assert ZeroArrayQ(Array(0, 0, 0))
    Debug.Assert ZeroArrayQ([{0,0;0,0}])
    Debug.Assert Not ZeroArrayQ(Array(1, 2, "ASFF"))
    Debug.Assert Not ZeroArrayQ([{1,"a"; True, False}])
    Debug.Assert Not ZeroArrayQ([{1,2,3;4,5,6}])
    Debug.Assert ZeroArrayQ(EmptyArray)
    Debug.Assert Not ZeroArrayQ(Array(aDictionary, aDictionary))
    Debug.Assert Not ZeroArrayQ(Array(aDictionary))
    Debug.Assert Not ZeroArrayQ(Array(aListObject, ThisWorkbook))
    Debug.Assert Not ZeroArrayQ(Array(0, 1, 2.5, 3, 4))
    Debug.Assert Not ZeroArrayQ(Array(TempComputation))
    Debug.Assert Not ZeroArrayQ(Array(TempComputation, TempComputation))
    Debug.Assert Not ZeroArrayQ(Array("a", "b"))
    Debug.Assert Not ZeroArrayQ(Array("a"))
    Debug.Assert Not ZeroArrayQ(Array("a", 1))
    Debug.Assert Not ZeroArrayQ(Array("a", 1.1))
    Debug.Assert Not ZeroArrayQ(Array("a", -1.2))
    Debug.Assert Not ZeroArrayQ(Array(TempComputation, 1))
    
    Call aListObject.Delete
End Sub

Public Sub TestArrayPredicatesPositiveWholeNumberArrayQ()
    Dim var As Variant
    Dim aListObject As ListObject
    Dim aDictionary As Dictionary
    
    For Each var In TempComputation.ListObjects
        Call var.Delete
    Next
    
    Set aListObject = AddListObject(TempComputation.Range("A1"))
    Set aDictionary = New Dictionary
    
    For Each var In Array(#1/1/2000#, _
                          1, _
                          -1, _
                          1.1, _
                          -1.1, _
                          0, _
                          2.5, _
                          True, _
                          "asfggd", _
                          CVErr(1), _
                          TempComputation, _
                          ThisWorkbook, _
                          aListObject, _
                          aDictionary, _
                          Empty, _
                          Null)
        Debug.Assert Not PositiveWholeNumberArrayQ(var)
    Next

    Debug.Assert Not PositiveWholeNumberArrayQ(Array(0))
    Debug.Assert Not PositiveWholeNumberArrayQ(Array(0, 0, 0))
    Debug.Assert Not PositiveWholeNumberArrayQ([{0,0;0,0}])
    Debug.Assert Not PositiveWholeNumberArrayQ(Array(1, 2, "ASFF"))
    Debug.Assert Not PositiveWholeNumberArrayQ([{1,"a"; True, False}])
    Debug.Assert PositiveWholeNumberArrayQ([{1,2,3;4,5,6}])
    Debug.Assert PositiveWholeNumberArrayQ(Array(1, 2, 3))
    Debug.Assert PositiveWholeNumberArrayQ(EmptyArray)
    Debug.Assert Not PositiveWholeNumberArrayQ(Array(aDictionary, aDictionary))
    Debug.Assert Not PositiveWholeNumberArrayQ(Array(aDictionary))
    Debug.Assert Not PositiveWholeNumberArrayQ(Array(aListObject, ThisWorkbook))
    Debug.Assert Not PositiveWholeNumberArrayQ(Array(0, 1, 2.5, 3, 4))
    Debug.Assert Not PositiveWholeNumberArrayQ(Array(TempComputation))
    Debug.Assert Not PositiveWholeNumberArrayQ(Array(TempComputation, TempComputation))
    Debug.Assert Not PositiveWholeNumberArrayQ(Array("a", "b"))
    Debug.Assert Not PositiveWholeNumberArrayQ(Array("a"))
    Debug.Assert Not PositiveWholeNumberArrayQ(Array("a", 1))
    Debug.Assert Not PositiveWholeNumberArrayQ(Array("a", 1.1))
    Debug.Assert Not PositiveWholeNumberArrayQ(Array("a", -1.2))
    Debug.Assert Not PositiveWholeNumberArrayQ(Array(TempComputation, 1))
    
    Call aListObject.Delete
End Sub

Public Sub TestArrayPredicatesStringArrayQ()
    Dim var As Variant
    Dim aListObject As ListObject
    Dim aDictionary As Dictionary
    
    For Each var In TempComputation.ListObjects
        Call var.Delete
    Next
    
    Set aListObject = AddListObject(TempComputation.Range("A1"))
    Set aDictionary = New Dictionary
    
    For Each var In Array(#1/1/2000#, _
                          1, _
                          -1, _
                          1.1, _
                          -1.1, _
                          0, _
                          2.5, _
                          True, _
                          "asfggd", _
                          CVErr(1), _
                          TempComputation, _
                          ThisWorkbook, _
                          aListObject, _
                          aDictionary, _
                          Empty, _
                          Null)
        Debug.Assert Not StringArrayQ(var)
    Next

    Debug.Assert Not StringArrayQ(Array(0))
    Debug.Assert Not StringArrayQ(Array(0, 0, 0))
    Debug.Assert Not StringArrayQ([{0,0;0,0}])
    Debug.Assert Not StringArrayQ(Array(1, 2, "ASFF"))
    Debug.Assert Not StringArrayQ([{1,"a"; True, False}])
    Debug.Assert Not StringArrayQ([{1,2,3;4,5,6}])
    Debug.Assert Not StringArrayQ(Array(1, 2, 3))
    Debug.Assert StringArrayQ(EmptyArray)
    Debug.Assert Not StringArrayQ(Array(aDictionary, aDictionary))
    Debug.Assert Not StringArrayQ(Array(aDictionary))
    Debug.Assert Not StringArrayQ(Array(aListObject, ThisWorkbook))
    Debug.Assert Not StringArrayQ(Array(0, 1, 2.5, 3, 4))
    Debug.Assert Not StringArrayQ(Array(TempComputation))
    Debug.Assert Not StringArrayQ(Array(TempComputation, TempComputation))
    Debug.Assert StringArrayQ(Array("a", "b"))
    Debug.Assert StringArrayQ(Array("a"))
    Debug.Assert StringArrayQ([{"a","b";"c","d"}])
    Debug.Assert Not StringArrayQ(Array("a", 1))
    Debug.Assert Not StringArrayQ(Array("a", 1.1))
    Debug.Assert Not StringArrayQ(Array("a", -1.2))
    Debug.Assert Not StringArrayQ(Array(TempComputation, 1))
    
    Call aListObject.Delete
End Sub

Public Sub TestArrayPredicatesOneArrayQ()
    Dim var As Variant
    Dim aListObject As ListObject
    Dim aDictionary As Dictionary
    
    For Each var In TempComputation.ListObjects
        Call var.Delete
    Next
    
    Set aListObject = AddListObject(TempComputation.Range("A1"))
    Set aDictionary = New Dictionary
    
    For Each var In Array(#1/1/2000#, _
                          1, _
                          -1, _
                          1.1, _
                          -1.1, _
                          0, _
                          2.5, _
                          True, _
                          "asfggd", _
                          CVErr(1), _
                          TempComputation, _
                          ThisWorkbook, _
                          aListObject, _
                          aDictionary, _
                          Empty, _
                          Null)
        Debug.Assert Not OneArrayQ(var)
    Next

    Debug.Assert OneArrayQ(Array(1))
    Debug.Assert OneArrayQ(Array(1, 1, 1))
    Debug.Assert OneArrayQ([{1,1;1,1}])
    Debug.Assert Not OneArrayQ(Array(0))
    Debug.Assert Not OneArrayQ(Array(0, 0, 0))
    Debug.Assert Not OneArrayQ([{0,0;0,0}])
    Debug.Assert Not OneArrayQ(Array(1, 2, "ASFF"))
    Debug.Assert Not OneArrayQ([{1,"a"; True, False}])
    Debug.Assert Not OneArrayQ([{1,2,3;4,5,6}])
    Debug.Assert Not OneArrayQ(Array(1, 2, 3))
    Debug.Assert OneArrayQ(EmptyArray)
    Debug.Assert Not OneArrayQ(Array(aDictionary, aDictionary))
    Debug.Assert Not OneArrayQ(Array(aDictionary))
    Debug.Assert Not OneArrayQ(Array(aListObject, ThisWorkbook))
    Debug.Assert Not OneArrayQ(Array(0, 1, 2.5, 3, 4))
    Debug.Assert Not OneArrayQ(Array(TempComputation))
    Debug.Assert Not OneArrayQ(Array(TempComputation, TempComputation))
    Debug.Assert Not OneArrayQ(Array("a", "b"))
    Debug.Assert Not OneArrayQ(Array("a"))
    Debug.Assert Not OneArrayQ([{"a","b";"c","d"}])
    Debug.Assert Not OneArrayQ(Array("a", 1))
    Debug.Assert Not OneArrayQ(Array("a", 1.1))
    Debug.Assert Not OneArrayQ(Array("a", -1.2))
    Debug.Assert Not OneArrayQ(Array(TempComputation, 1))
    
    Call aListObject.Delete
End Sub

Public Sub TestArrayPredicatesElementwiseArithmeticParameterConsistentQ()
    Debug.Assert ElementwiseArithmeticParameterConsistentQ(Empty, Empty)
    Debug.Assert ElementwiseArithmeticParameterConsistentQ(1, 1)
    Debug.Assert ElementwiseArithmeticParameterConsistentQ(1, Array(1))
    Debug.Assert ElementwiseArithmeticParameterConsistentQ(1, Array(1, 2))
    Debug.Assert ElementwiseArithmeticParameterConsistentQ(1, [{1;2}])
    Debug.Assert ElementwiseArithmeticParameterConsistentQ(1, [{1,2;3,4}])
    Debug.Assert Not ElementwiseArithmeticParameterConsistentQ(Array(Array(1)), Array(1, 2))
    Debug.Assert ElementwiseArithmeticParameterConsistentQ(Array(1, 2), Array(1, 2))
    Debug.Assert ElementwiseArithmeticParameterConsistentQ(Array(1, 2, 3), Array(1, 2, 3))
    Debug.Assert Not ElementwiseArithmeticParameterConsistentQ(Array(1, 2, 3), [{1;2;3}])
    Debug.Assert ElementwiseArithmeticParameterConsistentQ([{1,2,3;4,5,6}], Array(1, 2, 3))
    Debug.Assert Not ElementwiseArithmeticParameterConsistentQ([{1,2,3, 44;4,5,6,66}], Array(1, 2, 3))
    Debug.Assert ElementwiseArithmeticParameterConsistentQ([{1,2,3, 44;4,5,6,66}], [{11;22}])
    Debug.Assert Not ElementwiseArithmeticParameterConsistentQ([{1,2,3, 44;4,5,6,66}], [{11;22;33}])
    Debug.Assert Not ElementwiseArithmeticParameterConsistentQ([{1,2,3, 44;4,5,6,66}], [{11,22,33;111,222,333}])
    Debug.Assert Not ElementwiseArithmeticParameterConsistentQ([{1,2,3, 44;4,5,6,66}], [{11,22;22,33;33,44}])
    Debug.Assert ElementwiseArithmeticParameterConsistentQ([{1,2,3, 44;4,5,6,66}], [{1,2,3,4;5,6,7,8}])
End Sub

Public Sub TestPredicatesRowVectorQ()
    Dim var As Variant
    Dim aListObject As ListObject
    Dim aDictionary As Dictionary
    
    For Each var In TempComputation.ListObjects
        Call var.Delete
    Next
    
    Set aListObject = AddListObject(TempComputation.Range("A1"))
    Set aDictionary = New Dictionary
    
    For Each var In Array(#1/1/2000#, _
                          1, _
                          -1, _
                          1.1, _
                          -1.1, _
                          0, _
                          2.5, _
                          True, _
                          "asfggd", _
                          CVErr(1), _
                          TempComputation, _
                          ThisWorkbook, _
                          aListObject, _
                          aDictionary, _
                          Empty, _
                          Null, _
                          [{1; 2; 3}], _
                          Array(Array(1, 2, 3)), _
                          Array("a", 1, 2), _
                          Array("a"))
        Debug.Assert Not RowVectorQ(var)
    Next

    For Each var In Array(Array(1, 2, 3), EmptyArray(), NumericalSequence(1, 10, 0.5))
        Debug.Assert RowVectorQ(var)
    Next

    Call aListObject.Delete
End Sub

Public Sub TestPredicatesColumnVectorQ()
    Dim var As Variant
    Dim aListObject As ListObject
    Dim aDictionary As Dictionary
    
    For Each var In TempComputation.ListObjects
        Call var.Delete
    Next
    
    Set aListObject = AddListObject(TempComputation.Range("A1"))
    Set aDictionary = New Dictionary
    
    For Each var In Array(#1/1/2000#, _
                          1, _
                          -1, _
                          1.1, _
                          -1.1, _
                          0, _
                          2.5, _
                          True, _
                          "asfggd", _
                          CVErr(1), _
                          TempComputation, _
                          ThisWorkbook, _
                          aListObject, _
                          aDictionary, _
                          Empty, _
                          Null, _
                          [{1; 2; 3; 4,5,6}], _
                          Array(Array(1, 2, 3)), _
                          Array("a", 1, 2), _
                          Array("a"))
        Debug.Assert Not ColumnVectorQ(var)
    Next
    
    For Each var In Array([{1;2;3}], _
                          EmptyArray(), _
                          TransposeMatrix(NumericalSequence(1, 10, 0.5)))
        Debug.Assert ColumnVectorQ(var)
    Next

    Call aListObject.Delete
End Sub

Public Sub TestPredicatesVectorQ()
    Dim var As Variant
    Dim aListObject As ListObject
    Dim aDictionary As Dictionary
    
    For Each var In TempComputation.ListObjects
        Call var.Delete
    Next
    
    Set aListObject = AddListObject(TempComputation.Range("A1"))
    Set aDictionary = New Dictionary
    
    For Each var In Array(#1/1/2000#, _
                          1, _
                          -1, _
                          1.1, _
                          -1.1, _
                          0, _
                          2.5, _
                          True, _
                          "asfggd", _
                          CVErr(1), _
                          TempComputation, _
                          ThisWorkbook, _
                          aListObject, _
                          aDictionary, _
                          Empty, _
                          Null, _
                          [{1; 2; 3; 4,5,6}], _
                          Array(Array(1, 2, 3)), _
                          Array("a", 1, 2), _
                          Array("a"))
        Debug.Assert Not VectorQ(var)
    Next
    
    For Each var In Array([{1;2;3}], EmptyArray(), _
                          TransposeMatrix(NumericalSequence(1, 10, 0.5)), _
                          [{1;2;3}], _
                          EmptyArray())
        Debug.Assert VectorQ(var)
    Next

    Call aListObject.Delete
End Sub

Public Sub TestPredicatesRowArrayQ()
    Dim var As Variant
    Dim aListObject As ListObject
    Dim aDictionary As Dictionary
    
    For Each var In TempComputation.ListObjects
        Call var.Delete
    Next
    
    Set aListObject = AddListObject(TempComputation.Range("A1"))
    Set aDictionary = New Dictionary
    
    For Each var In Array(#1/1/2000#, _
                          1, _
                          -1, _
                          1.1, _
                          -1.1, _
                          0, _
                          2.5, _
                          True, _
                          "asfggd", _
                          CVErr(1), _
                          TempComputation, _
                          ThisWorkbook, _
                          aListObject, _
                          aDictionary, _
                          Empty, _
                          Null, _
                          [{1; 2; 3}], _
                          Array(Array(1, 2, 3)))
        Debug.Assert Not RowArrayQ(var)
    Next

    For Each var In Array(Array(1, 2, 3), _
                          EmptyArray(), _
                          NumericalSequence(1, 10, 0.5), _
                          Array("a", 1, 2), _
                          Array("a"))
        Debug.Assert RowArrayQ(var)
    Next

    Call aListObject.Delete
End Sub

Public Sub TestPredicatesColumnArrayQ()
    Dim var As Variant
    Dim aListObject As ListObject
    Dim aDictionary As Dictionary
    
    For Each var In TempComputation.ListObjects
        Call var.Delete
    Next
    
    Set aListObject = AddListObject(TempComputation.Range("A1"))
    Set aDictionary = New Dictionary
    
    For Each var In Array(#1/1/2000#, _
                          1, _
                          -1, _
                          1.1, _
                          -1.1, _
                          0, _
                          2.5, _
                          True, _
                          "asfggd", _
                          CVErr(1), _
                          TempComputation, _
                          ThisWorkbook, _
                          aListObject, _
                          aDictionary, _
                          Empty, _
                          Null, _
                          Array(1, 2, 3), _
                          Array(Array(1, 2, 3)), _
                          Array("a"))
        Debug.Assert Not ColumnArrayQ(var)
    Next

    For Each var In Array([{1; 2; 3}], _
                          TransposeMatrix(NumericalSequence(1, 10, 0.5)), _
                          TransposeMatrix(Array("a", 1, 2)), _
                          EmptyArray())
        Debug.Assert ColumnArrayQ(var)
    Next

    Call aListObject.Delete
End Sub

Public Sub TestPredicatesRowOrColumnArrayQ()
    Dim var As Variant
    Dim aListObject As ListObject
    Dim aDictionary As Dictionary
    
    For Each var In TempComputation.ListObjects
        Call var.Delete
    Next
    
    Set aListObject = AddListObject(TempComputation.Range("A1"))
    Set aDictionary = New Dictionary
    
    For Each var In Array(#1/1/2000#, _
                          1, _
                          -1, _
                          1.1, _
                          -1.1, _
                          0, _
                          2.5, _
                          True, _
                          "asfggd", _
                          CVErr(1), _
                          TempComputation, _
                          ThisWorkbook, _
                          aListObject, _
                          aDictionary, _
                          Null, _
                          Array(Array(1, 2, 3)), _
                          Empty)
        Debug.Assert Not RowOrColumnArrayQ(var)
    Next

    For Each var In Array([{1; 2; 3}], _
                          TransposeMatrix(NumericalSequence(1, 10, 0.5)), _
                          TransposeMatrix(Array("a", 1, 2)), _
                          Array(1, 2, 3), _
                          EmptyArray(), _
                          Array("a"))
        Debug.Assert RowOrColumnArrayQ(var)
    Next

    Call aListObject.Delete
End Sub

Public Sub TestPredicatesInterpretableAsRowArrayQ()
    Dim var As Variant
    Dim aListObject As ListObject
    Dim aDictionary As Dictionary
    
    For Each var In TempComputation.ListObjects
        Call var.Delete
    Next
    
    Set aListObject = AddListObject(TempComputation.Range("A1"))
    Set aDictionary = New Dictionary
    
    For Each var In Array(#1/1/2000#, _
                          1, _
                          -1, _
                          1.1, _
                          -1.1, _
                          0, _
                          2.5, _
                          True, _
                          "asfggd", _
                          CVErr(1), _
                          TempComputation, _
                          ThisWorkbook, _
                          aListObject, _
                          aDictionary, _
                          Null, _
                          Empty, _
                          [{1; 2; 3}] _
                          )
        Debug.Assert Not InterpretableAsRowArrayQ(var)
    Next

    For Each var In Array(Array(Array(1, 2, 3)), _
                          Array(Array(Array(1, 2, 3))), _
                          [{1,2,3}], _
                          Array(1, 2, 3), _
                          EmptyArray())
        Debug.Assert InterpretableAsRowArrayQ(var)
    Next

    Call aListObject.Delete
End Sub

Public Sub TestPredicatesInterpretableAsColumnArrayQ()
    Dim var As Variant
    Dim aListObject As ListObject
    Dim aDictionary As Dictionary
    
    For Each var In TempComputation.ListObjects
        Call var.Delete
    Next
    
    Set aListObject = AddListObject(TempComputation.Range("A1"))
    Set aDictionary = New Dictionary
    
    For Each var In Array(#1/1/2000#, _
                          1, _
                          -1, _
                          1.1, _
                          -1.1, _
                          0, _
                          2.5, _
                          True, _
                          "asfggd", _
                          CVErr(1), _
                          TempComputation, _
                          ThisWorkbook, _
                          aListObject, _
                          aDictionary, _
                          Null, _
                          Empty, _
                          Array(Array(1, 2, 3)), _
                          Array(Array(Array(1, 2, 3))), _
                          Array([{1;2;3}], 2) _
                          )
        Debug.Assert Not InterpretableAsColumnArrayQ(var)
    Next

    For Each var In Array([{1; 2; 3}], _
                          Array([{1;2;3}]), _
                          EmptyArray())
        Debug.Assert InterpretableAsColumnArrayQ(var)
    Next

    Call aListObject.Delete
End Sub

Public Sub TestPredicatesSpanArrayQ()
    Dim var As Variant
    Dim aListObject As ListObject
    Dim aDictionary As Dictionary
    Dim aSpan As New Span
    
    For Each var In TempComputation.ListObjects
        Call var.Delete
    Next
    
    Set aListObject = AddListObject(TempComputation.Range("A1"))
    Set aDictionary = New Dictionary
    
    For Each var In Array(#1/1/2000#, _
                          1, _
                          -1, _
                          1.1, _
                          -1.1, _
                          0, _
                          2.5, _
                          True, _
                          "asfggd", _
                          CVErr(1), _
                          TempComputation, _
                          ThisWorkbook, _
                          aListObject, _
                          aDictionary, _
                          Null, _
                          Empty, _
                          Array(Array(1, 2, 3)), _
                          Array(Array(Array(1, 2, 3))), _
                          Array([{1;2;3}], 2) _
                          )
        Debug.Assert Not SpanArrayQ(var)
    Next

    For Each var In Array(Array(aSpan), _
                          Array(aSpan, aSpan), _
                          EmptyArray())
        Debug.Assert SpanArrayQ(var)
    Next

    Call aListObject.Delete
End Sub

Public Sub TestPredicatesPartIndexArrayQ()
    Dim var As Variant
    Dim aListObject As ListObject
    Dim aDictionary As Dictionary
    Dim aSpan As New Span
    
    For Each var In TempComputation.ListObjects
        Call var.Delete
    Next
    
    Set aListObject = AddListObject(TempComputation.Range("A1"))
    Set aDictionary = New Dictionary
    
    For Each var In Array(#1/1/2000#, _
                          1, _
                          -1, _
                          1.1, _
                          -1.1, _
                          0, _
                          2.5, _
                          True, _
                          "asfggd", _
                          CVErr(1), _
                          TempComputation, _
                          ThisWorkbook, _
                          aListObject, _
                          aDictionary, _
                          Null, _
                          Empty, _
                          Array(Array(1, 2), 2, 3), _
                          Array(Array(1, 2, 3)), _
                          Array(Array(Array(1, 2, 3))), _
                          Array([{1;2;3}], 2) _
                          )
        Debug.Assert Not PartIndexArrayQ(var)
    Next

    For Each var In Array(Array(aSpan), _
                          Array(aSpan, aSpan), _
                          Array(1, aSpan), _
                          Array(1, 2, 3))
        Debug.Assert PartIndexArrayQ(var)
    Next

    Call aListObject.Delete
End Sub

Public Sub TestPredicatesTakeIndexArrayQ()
    Dim var As Variant
    Dim aListObject As ListObject
    Dim aDictionary As Dictionary
    Dim aSpan As New Span
    
    For Each var In TempComputation.ListObjects
        Call var.Delete
    Next
    
    Set aListObject = AddListObject(TempComputation.Range("A1"))
    Set aDictionary = New Dictionary
    
    For Each var In Array(#1/1/2000#, _
                          1, _
                          -1, _
                          1.1, _
                          -1.1, _
                          0, _
                          2.5, _
                          True, _
                          "asfggd", _
                          CVErr(1), _
                          TempComputation, _
                          ThisWorkbook, _
                          aListObject, _
                          aDictionary, _
                          Null, _
                          Empty, _
                          Array(Array(Array(1, 2, 3))), _
                          Array(aSpan), _
                          Array(aSpan, aSpan), _
                          Array(1, aSpan) _
                          )
        Debug.Assert Not TakeIndexArrayQ(var)
    Next

    For Each var In Array(Array([{1;2;3}], 2), _
                          Array(Array(1, 2), 2, 3), _
                          Array(Array(1, 2, 3)), _
                          Array(1, 2, 3))
        Debug.Assert TakeIndexArrayQ(var)
    Next

    Call aListObject.Delete
End Sub

'********************************************************************************************
' FunctionalProgramming
'********************************************************************************************

