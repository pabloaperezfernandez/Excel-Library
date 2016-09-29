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
    Dim a() As Variant
    Dim B(1 To 2) As Variant
    Dim c As Integer
    Dim wbk As Workbook
    
    Debug.Assert DimensionedQ(EmptyArray())
    Debug.Assert Not DimensionedQ(a)
    Debug.Assert DimensionedQ(B)
    Debug.Assert Not DimensionedQ(c)
    Debug.Assert Not DimensionedQ(wbk)
End Sub

Public Sub TestPredicatesEmptyArrayQ()
    Dim a() As Variant
    Dim B(1 To 2) As Variant
    Dim c As Integer
    Dim wbk As Workbook
    
    Debug.Assert EmptyArrayQ(EmptyArray())
    Debug.Assert Not EmptyArrayQ(a)
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
Public Sub TestFunctionalProgrammingScan()
    Call Scan("ScanHelper1", NumericalSequence(1, 10))
End Sub

Private Sub ScanHelper1(arg As Variant)
    Debug.Print "The time for interation " & arg & " is " & Now()
End Sub

Public Sub TestFunctionalProgrammingThrough()
    PrintArray Through([{"NumberQ", "StringQ","BooleanQ"}], 123)
End Sub

Public Sub TestFunctionalProgrammingArrayMap()
    Dim a As Variant
    Dim VariantArrayOfStringArrays(1 To 2) As Variant
    
    Debug.Print "We are showing how to do MapThread using ArrayMap."
    Debug.Print
    Let a = Array(Array("1", "2", "3"), Array("4", "5", "6"))
    Debug.Print "a is:"
    PrintArray Pack2DArray(a)
    Debug.Print "All entries in a are strings."
    Debug.Print
    
    Let a = ArrayMap("StringConcatenate", a)
    Debug.Print "Let a = TransposeMatrix(Pack2DArray(Array(Array(""1"", ""2"", ""3""), Array(""4"", ""5"", ""6"")))"
    Debug.Print "Let a = ArrayMap(""StringConcatenate"", a)"
    PrintArray a
    
    Debug.Print
    Debug.Print "We are showing how to do MapThread using ArrayMap."
    Debug.Print
    Let VariantArrayOfStringArrays(1) = ToStrings(Array("1", "2", "3"))
    Let VariantArrayOfStringArrays(2) = ToStrings(Array("4", "5", "6"))
    Debug.Print "StringArray is:"
    PrintArray Pack2DArray(VariantArrayOfStringArrays)
    PrintArray ArrayMap("StringConcatenate", VariantArrayOfStringArrays)
    
    Debug.Print

    PrintArray ArrayMap("ArrayMapHelper1", NumericalSequence(1, 10))
End Sub

Private Function ArrayMapHelper1(arg As Variant) As Integer
    Let ArrayMapHelper1 = arg * 10
End Function

Public Sub TestFunctionalProgrammingArrayMapThread()
    PrintArray ArrayMapThread("StringJoin", _
                              Array("1", "2", "3"), _
                              Array("-11", "-22", "-33"))
    Debug.Print
    
    PrintArray ArrayMapThread("StringJoin", _
                              Array("1", "2", "3"), _
                              Array("-11", "-22", "-33"), _
                              Array("-111", "-222", "-332"))
    Debug.Print
    
    PrintArray ArrayMapThread("Total", _
                              [{1,2,3}], _
                              [{10,20,30}])
                              
    Debug.Print
    
    PrintArray ArrayMapThread("Total", _
                              [{1,2,3}], _
                              [{10,20,30}], _
                              [{100,200,300}], _
                              [{1000,2000,3000}])
End Sub

Public Sub TestFunctionalProgrammingArraySelect()
    PrintArray ArraySelect(NumericalSequence(1, 100), "ArraySelectHelper1")
    PrintArray ArraySelect(Array("aPAPa", "bbbPAP", "sdgasdkh"), "ArraySelectHelper2")
End Sub

Private Function ArraySelectHelper1(arg As Variant) As Boolean
    Let ArraySelectHelper1 = arg > 1 And arg < 6
End Function

Private Function ArraySelectHelper2(arg As Variant) As Boolean
    Let ArraySelectHelper2 = InStr(1, arg, "PAP")
End Function

Public Sub TestFunctionalProgrammingTotal()
    Dim a1 As Variant
    Dim a2DArray(1 To 1, 1 To 3) As Variant
    Dim i As Integer
    
    Let a1 = [{1,2,3,4,5,6,7,8,9,10}]
    Debug.Print "Adding array [{1,2,3,4,5,6,7,8,9,10}] using Total"
    Debug.Print Total(a1)
    Debug.Print
    
    Debug.Print "Mapping Total onto Array([{1,2}], [{10,20,30}], [{100,200,300,400}])"
    PrintArray ArrayMap("Total", Array([{1,2}], [{10,20,30}], [{100,200,300,400}]))
    Debug.Print
    
    Debug.Print "Testing total on [{1,2,3;10,20,30;100,200,300}]"
    PrintArray Total([{1,2,3;10,20,30;100,200,300}])
    Debug.Print
    
    Debug.Print "Testing total on [{1,2,3;10,20,30;100,200,300}] on dimension 2 (adding by columns)"
    PrintArray Total([{1,2,3;10,20,30;100,200,300}], 2)
    Debug.Print
    
    Debug.Print "Testing total on [{1,2,3}] (as true 2D matrix) on dimension 2 (adding by columns)"
    For i = 1 To 3: Let a2DArray(1, i) = i: Next
    PrintArray Total(a2DArray)
    Debug.Print
End Sub

Public Sub TestFunctionalProgrammingTimes()
    Dim a1 As Variant
    Dim a2DArray(1 To 1, 1 To 3) As Variant
    Dim i As Integer
    
    Let a1 = [{1,2,3,4,5,6,7,8,9,10}]
    Debug.Print "Adding array [{1,2,3,4,5,6,7,8,9,10}] using Total"
    Debug.Print Times(a1)
    Debug.Print
    
    Debug.Print "Mapping Times onto Array([{1,2}], [{10,20,30}], [{100,200,300,400}])"
    PrintArray ArrayMap("Times", Array([{1,2}], [{10,20,30}], [{100,200,300,400}]))
    Debug.Print
    
    Debug.Print "Testing Times on [{1,2,3;10,20,30;100,200,300}]"
    PrintArray Times([{1,2,3;10,20,30;100,200,300}])
    Debug.Print
    
    Debug.Print "Testing Times on [{1,2,3;10,20,30;100,200,300}] on dimension 2 (mult by columns)"
    PrintArray Times([{1,2,3;10,20,30;100,200,300}], 2)
    Debug.Print
    
    Debug.Print "Testing Times on [{1,2,3}] (as true 2D matrix) on dimension 2 (mult by columns)"
    For i = 1 To 3: Let a2DArray(1, i) = i: Next
    PrintArray Times(a2DArray)
    Debug.Print
End Sub

Public Sub TestFunctionalProgrammingAccumulate()
    Dim a As Variant
    Dim m(1 To 1, 1 To 3) As Variant
    Dim i As Integer
    Dim var As Variant
    
    Let a = [{1,2,3,4,5,6,7,8,9,10}]
    Debug.Print "Accumulating [{1,2,3,4,5,6,7,8,9,10}]"
    PrintArray Accumulate(a)
    Debug.Print
    
    Debug.Print "Mapping Accumulate onto Array([{1,2}], [{10,20,30}], [{100,200,300,400}])"
    For Each var In ArrayMap("Accumulate", Array([{1,2}], [{10,20,30}], [{100,200,300,400}]))
        PrintArray var
        Debug.Print
    Next
    
    Debug.Print "Testing Accumulate on [{1,2,3;10,20,30;100,200,300}]"
    For Each var In Accumulate([{1,2,3;10,20,30;100,200,300}])
        PrintArray var
    Next
    Debug.Print

    For i = 1 To 3
        Let m(1, i) = i
    Next
    Debug.Print "Testing Accumulate on [{1,2,3}] (as true 2D matrix)"
    For Each var In Accumulate(m)
        PrintArray var
        Debug.Print
    Next
    Debug.Print
End Sub

Public Sub TestFunctionalProgrammingNest()
    Dim i As Long
    
    For i = 1 To 4
        Debug.Print "Nesting (1+#)^2 & " & i & " times on 1: " & Nest("NestHelper1", 1, i)
    Next
    
    Debug.Print
    
    For i = 1 To 9
        Debug.Print "Nesting Most " & i & " times on [{1,2,3,4,5,6,7,8,9}]: " & _
                    PrintArray(Nest("Most", [{1,2,3,4,5,6,7,8,9}], i), True)
    Next
End Sub

Private Function NestHelper1(arg As Variant) As Long
    Let NestHelper1 = (arg + 1) ^ 2
End Function

Public Sub TestFunctionalProgrammingNestList()
    Dim i As Long
    Dim var As Variant
       
    For i = 1 To 4
        Debug.Print "NestList (1+#)^2 & " & i & " times on 1:"
        
        PrintArray NestList("NestHelper1", 1, i)
        Debug.Print
    Next
    
    Debug.Print
    
    For i = 1 To 9
        Debug.Print "NestList Rest " & i & " times on [{1,2,3,4,5,6,7,8,9}]:"
        For Each var In NestList("Rest", [{1,2,3,4,5,6,7,8,9}], i)
            PrintArray var
        Next
        
        Debug.Print
    Next
End Sub

Public Sub TestFunctionalProgrammingFold()
    Dim i As Long
    Dim var As Variant
       
    For i = 1 To 4
        Debug.Print "Fold Times on 1 and " & PrintArray(NumericalSequence(1, i), True)
        
        Debug.Print Fold("Multiply", 1, NumericalSequence(1, i))
        Debug.Print
    Next
    
    Debug.Print
    
    For i = 1 To 4
        Debug.Print "Fold Times on 10 and " & PrintArray(NumericalSequence(1, i), True)
        
        Debug.Print Fold("Multiply", 10, NumericalSequence(1, i))
        Debug.Print
    Next
    
    Debug.Print
    
    Debug.Print Fold("Multiply", 100, EmptyArray())
End Sub

Public Sub TestFunctionalProgrammingFoldList()
    Dim i As Long
    Dim var As Variant
       
    For i = 1 To 4
        Debug.Print "Fold Times on 1 and " & PrintArray(NumericalSequence(1, i), True)
        
        PrintArray FoldList("Multiply", 1, NumericalSequence(1, i))
    Next
    
    Debug.Print

    For i = 1 To 4
        Debug.Print "Fold Times on 10 and " & PrintArray(NumericalSequence(1, i), True)
        
        PrintArray FoldList("Multiply", 10, NumericalSequence(1, i))
    Next
    
    Debug.Print
    
    PrintArray FoldList("Multiply", 10, EmptyArray())
End Sub

'********************************************************************************************
' Arrays
'********************************************************************************************

Public Sub TestArraysPart()
    Dim i As Long
    Dim j As Long
    Dim m As Variant
    Dim a As Variant
    Dim r As Long
    Dim c As Long
    
    ' 1D Tests
    Let m = Array(1, 2, 3, 4, 5)
    Debug.Print "We have m = Array(1, 2, 3, 4, 5)"
    For i = -10 To 10
        Debug.Print "Part(m, " & i & ") = " & Part(m, i)
    Next
    
    Debug.Print
    Debug.Print "Getting Part(m, Array(1, 5))"
    PrintArray Part(m, Array(1, 5))
    
    Debug.Print
    Debug.Print "Getting Part(m, Span(1, -2))"
    PrintArray Part(m, Span(1, -2))

    Debug.Print
    Debug.Print "Getting Part(m, Span(1, -3))"
    PrintArray Part(m, Span(1, -3))
    
    Debug.Print
    Debug.Print "Getting Part(m, Span(2, -3))"
    PrintArray Part(m, Span(2, -3))
    
    Debug.Print
    ReDim m(0 To 5)
    For i = 1 To 6: Let m(i - 1) = i: Next
    Debug.Print "Setting m = Array(1, 2, 3, 4, 5, 6)"
    Debug.Print "LBound(m), UBound(m) = " & LBound(m) & ", " & UBound(m)
    For i = 1 To 6: Debug.Print "Part(Array(1, 2, 3, 4, 5, 6), " & i & ") = " & Part(Array(1, 2, 3, 4, 5, 6), i): Next
    
    Debug.Print
    Debug.Print "Cycling through elts Array(1, i)"
    For i = 1 To 6: PrintArray Part(Array(1, 2, 3, 4, 5, 6), Array(1, i)): Next
    Debug.Print
    Debug.Print "Cycling through elts Array(i, -1)"
    For i = 1 To 6: PrintArray Part(Array(1, 2, 3, 4, 5, 6), Array(i, -1)): Next
    Debug.Print
    Debug.Print "Cycling through elts Array(-6, i)"
    For i = 1 To 6: PrintArray Part(Array(1, 2, 3, 4, 5, 6), Array(-6, i)): Next
    
    Debug.Print
    Debug.Print "Get elts 3, 5, 1 from Array(1, 2, 3, 4, 5, 6)"
    PrintArray Part(Array(1, 2, 3, 4, 5, 6), Array(3, 5, 1))
    Debug.Print
    Debug.Print "Get stepped segment Array(3, 5, 1)"
    PrintArray Part(Array(1, 2, 3, 4, 5, 6), Span(3, 5, 1))
    Debug.Print
    Debug.Print "Get stepped segment Array(3, 5, 2)"
    PrintArray Part(Array(1, 2, 3, 4, 5, 6), Span(3, 5, 2))
    Debug.Print
    Debug.Print "Get stepped segment Array(3, 5, 6)"
    PrintArray Part(Array(1, 2, 3, 4, 5, 6), Span(3, 5, 6))
    
    ' 2D Test with one dimensional index set
    Debug.Print
    Debug.Print "2D Test with one dimensional index set"
    Debug.Print
    
    Let a = [{1,2,3;4,5,6;7,8,9;10,11,12;13,14,15}]
    Debug.Print "Testing on A:"
    PrintArray a
    Debug.Print
    For i = -10 To 10: Debug.Print "Row " & i & " is " & PrintArray(Part(a, i), True): Next
    Debug.Print
    Debug.Print "Getting set of rows."
    ReDim m(1 To 3)
    For i = 1 To 10
        For j = 1 To 3
            Let m(j) = Application.WorksheetFunction.RandBetween(1, 5)
        Next
        
        Debug.Print "Rows (" & PrintArray(m, True) & ") is " & vbCr & PrintArray(Part(a, m), True)
        Debug.Print
    Next
    Debug.Print
    
    Debug.Print "A is"
    PrintArray a
    Debug.Print
    
    Debug.Print "Trying Part(A, Array(2, 5))"
    PrintArray Part(a, Array(2, 5))
    Debug.Print
    
    Debug.Print "Trying Part(A, Array(2, -2))"
    PrintArray Part(a, Array(2, -2))
    Debug.Print
        
    Debug.Print "Part(A, 2, Array(1, 2))"
    PrintArray Part(a, 2, Array(1, 2))
    Debug.Print

    Debug.Print "Part(A, 2, Array(2, 3))"
    PrintArray Part(a, 2, Array(2, 3))
    Debug.Print
    
    Debug.Print "Part(A, array(2,4), Array(2, 3))"
    PrintArray Part(a, Array(2, 4), Array(2, 3))
    Debug.Print
    
    Debug.Print "Testing spans"
    
    Debug.Print "Part(A, Span(1, 2))"
    PrintArray Part(a, Span(1, 2))
    Debug.Print
    
    Debug.Print "Part(A, Span(1, 3))"
    PrintArray Part(a, Span(1, 3))
    Debug.Print
    
    Debug.Print "Part(A, Span(1, 4))"
    PrintArray Part(a, Span(1, 4))
    Debug.Print
    
    Debug.Print "Part(A, Span(2, -1))"
    PrintArray Part(a, Span(2, -1))
    Debug.Print
    
    Debug.Print "Part(A, Span(2, -2))"
    PrintArray Part(a, Span(2, -2))
    Debug.Print
    
    Debug.Print "Part(A, Span(1, -1, 2))"
    PrintArray Part(a, Span(1, -1, 2))
    Debug.Print

    Debug.Print "Part(A, Span(2, -1, 2))"
    PrintArray Part(a, Span(2, -1, 2))
    Debug.Print

    Debug.Print "Part(A, 2, Span(2, -1))"
    PrintArray Part(a, 2, Span(2, -1))
    Debug.Print
    
    Debug.Print "Part(A, 2, 3)"
    Debug.Print Part(a, 2, 3)
    Debug.Print
    
    Debug.Print "Part(A, Span(1, -1), 2)"
    PrintArray Part(a, Span(1, -1), 2)
    Debug.Print
    
    Debug.Print "Part(A, Span(1, -1), Span(2, -1))"
    PrintArray Part(a, Span(1, -1), Span(2, -1))
    Debug.Print
    
    Debug.Print "Part(A, Span(1, -1, 2), Span(1, 3, 2))"
    PrintArray Part(a, Span(1, -1, 2), Span(1, 3, 2))
    Debug.Print
    
    Let m = ConstantArray(Empty, 7, 8)
    For i = 1 To NumberOfRows(m)
        For j = 1 To NumberOfColumns(m)
            If j > 1 Then
                Let m(i, j) = 10 ^ (j - 1) * i + m(i, j - 1)
            Else
                Let m(i, j) = i
            End If
        Next
    Next
    Debug.Print "Cycling through elements of:"
    PrintArray m
    Debug.Print
    Debug.Print "Getting individual rows 1 through 6"
    For i = 1 To 6: PrintArray Part(m, i): Next
    Debug.Print
    Debug.Print "Cycling through segments Array(1, i)"
    For i = 1 To 6: PrintArray Part(m, Array(1, i)): Debug.Print: Next
    Debug.Print
    Debug.Print "Cycling through segments Array(i, -1)"
    For i = 1 To 6: PrintArray Part(m, Array(i, -1)): Debug.Print: Next
    Debug.Print
    Debug.Print "Cycling through segments Array(-6, i)"
    For i = 1 To 6: PrintArray Part(m, Array(-6, i)): Debug.Print: Next
    Debug.Print
    
    Debug.Print "Get elts 3, 5, 1"
    PrintArray Part(m, Array(3, 5, 1))
    Debug.Print
    Debug.Print "Get stepped segment Span(3, 5, 1)"
    PrintArray Part(m, Span(3, 5, 1))
    Debug.Print
    Debug.Print "Get stepped segment Span(3, 5, 2)"
    PrintArray Part(m, Span(3, 5, 2))
    Debug.Print
    Debug.Print "Get stepped segment Span(3, 5, 6)"
    PrintArray Part(m, Span(3, 5, 6))
    Debug.Print
    Debug.Print "Get segment Span(-4,-2))"
    PrintArray Part(m, Span(-4, -2))
    Debug.Print
    Debug.Print "Get segment Span(-4,-2)"
    PrintArray Part(m, Span(-4, -2))
    
    ' 2D Test with two dimensional index sets
    Let m = ConstantArray(Empty, 7, 8)
    For i = 1 To NumberOfRows(m)
        For j = 1 To NumberOfColumns(m)
            If j > 1 Then
                Let m(i, j) = 10 ^ (j - 1) * i + m(i, j - 1)
            Else
                Let m(i, j) = i
            End If
        Next
    Next
    Debug.Print "Cycling through elements of:"
    PrintArray m
    Debug.Print
    Debug.Print "Getting individual columns 1 through 6"
    For i = 1 To 6: PrintArray Part(m, Span(1, -1), i): Debug.Print: Next
    Debug.Print
    Debug.Print "Cycling through segments Span(1, -1), Span(1, i)"
    For i = 1 To 6: PrintArray Part(m, Span(1, -1), Span(1, i)): Debug.Print: Next
    Debug.Print
    Debug.Print "Cycling through segments Array(1, -1), Array(i, -1)"
    For i = 1 To 6: PrintArray Part(m, Array(1, -1), Array(i, -1)): Debug.Print: Next
    Debug.Print
    Debug.Print "Cycling through segments Array(1, -1), Array(-6, i)"
    For i = 1 To 6: PrintArray Part(m, Array(1, -1), Array(-6, i)): Debug.Print: Next
    Debug.Print
    Debug.Print "Get Array(1, -1) with Array(Array(3, 5, 1))"
    PrintArray Part(m, Array(1, -1), Array(Array(3, 5, 1)))
    Debug.Print
    Debug.Print "Get stepped segment Array(1, -1), Span(3, 5, 1)"
    PrintArray Part(m, Array(1, -1), Span(3, 5, 1))
    Debug.Print
    Debug.Print "Get stepped segment Array(1, -1),Span(3, 5, 2)"
    PrintArray Part(m, Array(1, -1), Span(3, 5, 2))
    Debug.Print
    Debug.Print "Get stepped segment Array(1, -1), Span(3, 5, 6)"
    PrintArray Part(m, Array(1, -1), Span(3, 5, 6))
    Debug.Print
    Debug.Print "Get segment Array(1, -1), Array(-4,-2))"
    PrintArray Part(m, Array(1, -1), Array(-4, -2))
    Debug.Print
    Debug.Print "Get segment Array(1, -1), Array(Array(-4,-2))"
    PrintArray Part(m, Array(1, -1), Array(Array(-4, -2)))
    Debug.Print
    Debug.Print "Getting a rectangular submatrix Span(3,4), Span(5,6,7)"
    PrintArray Part(m, Span(3, 4), Span(5, 6, 7))
    Debug.Print
    
    Debug.Print "SPEED TESTS"
    Debug.Print "Creating a large 10000 by 1000 2D array"
    Let m = ConstantArray(Empty, 10000, 1000)
    For i = 1 To 10000: For j = 1 To 1000: Let m(i, j) = i * 1000 + j: Next: Next
    Debug.Print "Accessing element 500 by 200"
    Debug.Print m(500, 200)
    Debug.Print "Accessing row 7000"
    Let a = Part(m, 7000)
    PrintArray a
    Debug.Print "Accessig column 700"
    Let a = Part(m, Span(1, -1), 700)
    Debug.Print "The arrays dimensions are: ", LBound(a), UBound(a)
    Debug.Print "The array is"
    PrintArray a
End Sub

Public Sub TestArraysTake()
    Dim a() As Integer
    Dim r As Long
    Dim c As Long
    
    ReDim a(1 To 7)
    For r = 1 To 7
        Let a(r) = r
    Next
    
    Debug.Print
    Debug.Print "Set a equal to "
    PrintArray a
    Debug.Print "LBound(a,1), UBound(a,1) = " & LBound(a, 1) & ", " & UBound(a, 1)
    
    Debug.Print
    For r = -10 To 10
        Debug.Print "Testing Take(a, " & r; ")"
        PrintArray Take(a, r)
    Next
    
    ReDim a(0 To 6)
    For r = 0 To 6
        Let a(r) = r
    Next
    
    Debug.Print
    Debug.Print "Set a equal to "
    PrintArray a
    Debug.Print "LBound(a,1), UBound(a,1) = " & LBound(a, 1) & ", " & UBound(a, 1)
    
    Debug.Print
    For r = -10 To 10
        Debug.Print "Testing Take(a, " & r & ")"
        PrintArray Take(a, r)
    Next

    ReDim a(1 To 9, 1 To 3)
    For r = 1 To 9
        For c = 1 To 3
            Let a(r, c) = Application.WorksheetFunction.Rept(r, c)
        Next
    Next
    Debug.Print "Set a to:"
    PrintArray a
    Debug.Print
    Debug.Print "Bounds LBound(a,1), UBound(a,1), LBound(a,2), UBound(a,2): "
    Debug.Print LBound(a, 1) & ", " & UBound(a, 1) & ", " & LBound(a, 2) & ", " & UBound(a, 2)
    Debug.Print
    
    For r = -10 To 10
        Debug.Print "Testing Take(a," & r & ")"
        PrintArray Take(a, r)
        Debug.Print
    Next
    
    Debug.Print "Testing Take(EmptyArray(),1)"
    PrintArray Take(EmptyArray(), 1)
    Debug.Print
     
    Debug.Print "Testing Take(a, Array(-2,-4,-5))"
    PrintArray Take(a, Array(-2, -4, -5))
    Debug.Print
    
    Debug.Print "Testing Take(a, Array(-4,-2))"
    PrintArray Take(a, Array(-4, -2))
    Debug.Print
    
    Debug.Print "Testing Take(a, Array(-4,-1,2))"
    PrintArray Take(a, Array(-4, -1, 2))
    Debug.Print
    
    Debug.Print "Testing Take(a, Array(-6,-1,2))"
    PrintArray Take(a, Array(-6, -1, 2))
    Debug.Print
    
    Debug.Print "Testing Take(a, Array(-6,-1,10))"
    PrintArray Take(a, Array(-6, -1, 10))
    Debug.Print
    
    Debug.Print "Testing Take(a,EmptyArray())"
    PrintArray Take(a, EmptyArray())
    Debug.Print
    
    ReDim a(0 To 8, 0 To 3)
    For r = 0 To 8
        For c = 0 To 3
            Let a(r, c) = r
        Next
    Next
    
    Debug.Print "Set a to:"
    PrintArray a
    Debug.Print
    Debug.Print "Bounds LBound(a,1), UBound(a,1), LBound(a,2), UBound(a,2): " & ", " & LBound(a, 1) _
                & ", " & UBound(a, 1) & ", " & LBound(a, 2) & ", " & UBound(a, 2)
    Debug.Print
    
    For r = -10 To 10
        Debug.Print "Testing Take(a," & r & ")"
        PrintArray Take(a, r)
        Debug.Print
    Next
End Sub

Public Sub TestArraysFlatten()
    PrintArray Flatten(Array(Array(1, 2, 3), Array(4, 5, 6)))
    PrintArray Flatten([{1, 2, 3; 4, 5, 6}])
End Sub
