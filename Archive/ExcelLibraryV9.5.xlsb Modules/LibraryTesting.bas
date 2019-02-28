Attribute VB_Name = "LibraryTesting"
' PURPOSE OF THIS MODULE
'
' The purpose of this module is to record unit testing for this library. it
' also serves as examples of how to use the library.

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

' This function shows it is faster to dump an array than to change the screen
' directly even if screen updating is off
Public Sub TestChangeLargeRangeDirectlyInWorkSheetInsteadOfVba()
    Dim DumpArray() As Variant
    Dim r As Long
    
    Let Application.ScreenUpdating = False
    
    
    Call ToTemp(RandomReal(TheDimensions:=Array(100000, 1)))
    Debug.Print "DIRECT SCREEN METHOD"
    Debug.Print Now()
    For r = 1 To 100000 Step 2
        Let TempComputation.Cells(r, 1).Value2 = "temp"
    Next
    Debug.Print Now()
    
    Let Application.ScreenUpdating = True
    
    Call MsgBox("Direct Screen Updating Done")
    
    Let Application.ScreenUpdating = False
    
    Debug.Print
    Debug.Print "DUMPING ARRAY METHOD"
    Call ToTemp(RandomReal(TheDimensions:=Array(100000, 1)))
    ReDim DumpArray(1 To 100000, 1)
    Debug.Print Now()
    For r = 1 To 100000 Step 2
        Let DumpArray(r, 1) = "temp"
    Next
    For r = 2 To 100000 Step 2
        Let DumpArray(r, 1) = TempComputation.Cells(r, 1).Value2
    Next
    Call DumpInSheet(DumpArray, TempComputation.Range("A1"))
    Debug.Print Now()
    
    Let Application.ScreenUpdating = True
End Sub

' This shows that 1D array of 1D arrays matrices are only a little slower (in absolute time)
' than 2D arrays.  Hence, they are fast enough to our purposes.
Public Sub TestSpeedDumping1DOf1DMatrixVs2DMatrix()
    Dim DumpArray() As Variant
    Dim AnArray() As Variant
    Dim r As Long
    Dim c As Integer
    Const n As Long = 100000
    Const M As Integer = 20
    
    Debug.Print "1D Array of Arrays"
    Debug.Print "   - Array Creation: " & Now()
    Let AnArray = ConstantArray(1, M)
    ReDim DumpArray(1 To n)
    For r = 1 To n
        Let DumpArray(r) = AnArray
    Next
    Debug.Print "   - End Array Creation: " & Now()
    Debug.Print "   - Start Packing and Dumping: " & Now()
    Call ToTemp(Pack2DArray(DumpArray))
    Debug.Print "   - End Packing and Dumping: " & Now()
    
    Debug.Print
    
    Debug.Print "2D Array "
    Debug.Print "   - Array Creation: " & Now()
    ReDim DumpArray(1 To n, 1 To M)
    For r = 1 To n
        For c = 1 To M
            Let DumpArray(r, c) = 1
        Next
    Next
    Debug.Print "   - End Array Creation: " & Now()
    Debug.Print "   - Start Packing and Dumping: " & Now()
    Call ToTemp(DumpArray)
    Debug.Print "   - End Packing and Dumping: " & Now()
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
    Dim AWorkbook As Workbook
    Dim aListObject As ListObject
    Dim aDictionary As Dictionary
    Dim ANumericArray(1 To 2) As Integer
    Dim var2 As Variant
    
    Set aWorksheet = ActiveSheet
    Set AWorkbook = ThisWorkbook
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
                           Array("aWorkbook", AWorkbook), _
                           Array("aListObject", aListObject), _
                           Array("ANumericArray", ANumericArray), _
                           Array("aDictionary", aDictionary), _
                           Array("Empty", Empty), _
                           Array("Null", Null))
        Debug.Print "PrintableQ(" & First(var2) & ") = " & PrintableQ(Last(var2))
    Next

    Call aListObject.Delete
End Sub

Public Sub TestPredicatesAtomicPredicates()
    Dim anInteger As Integer
    Dim aDouble As Double
    Dim aDate As Date
    Dim aBoolean As Boolean
    Dim aString As String
    Dim aWorksheet As Worksheet
    Dim AWorkbook As Workbook
    Dim aListObject As ListObject
    Dim aDictionary As Dictionary
    Dim aVariant As Variant
    Dim AnArray(1 To 2) As Integer
    Dim var1 As Variant
    Dim var2 As Variant
    
    Set aWorksheet = ActiveSheet
    Set AWorkbook = ThisWorkbook
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
                               Array("aWorkbook", AWorkbook), _
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

Public Sub TestPredicatesLambdaQ()
    Dim l1 As New Lambda
    
    Debug.Assert LambdaQ(l1)
    Debug.Assert Not LambdaQ(1)
End Sub

Public Sub TestNumberOfDimensionsForUnDimensionedArray()
    Dim AnArray1() As Variant
    Dim AnArray2(1 To 2) As Variant
    
    Debug.Print "The number of dimensions of AnArray1() is " & NumberOfDimensions(AnArray1)
    Debug.Print "The number of dimensions of AnArray2(1 to 2) is " & NumberOfDimensions(AnArray2)
End Sub

Public Sub TestPredicatesEqualQ()
    Debug.Assert EqualQ(Error(1), Error(2))
    Debug.Assert Not EqualQ(Error(1), Error(3))
    Debug.Assert Not EqualQ(1, 1.2)
    Debug.Assert EqualQ(Array(1, 2), Array(1, 2))
    Debug.Assert Not EqualQ([{1, 2}], Array(1, 2))
    Debug.Assert EqualQ([{1,2;3,4}], [{1,2;3,4}])
    Debug.Assert Not EqualQ([{1,2;3,4}], [{1,2;3,5}])
    Debug.Assert EqualQ(Array(ThisWorkbook, TempComputation), _
                        Array(ThisWorkbook, TempComputation))
    Debug.Assert Not EqualQ(Array(ThisWorkbook, ThisWorkbook), _
                            Array(ThisWorkbook, TempComputation))
    Debug.Assert EqualQ(Array(Array(Array(1, 2), 3)), _
                        Array(Array(Array(1, 2), 3)))
    Debug.Assert Not EqualQ(Array(Array(Array(1, 2), 3)), _
                            Array(Array(Array(1, 2), 3#)))
    Debug.Assert EqualQ(Null, Null)
    Debug.Assert EqualQ(Null, Empty)
End Sub

Public Sub TestPredicatesMemberQ()
    Debug.Assert MemberQ(Array(1, 2, 3), 2)
    Debug.Assert MemberQ(Array(Empty, ThisWorkbook), ThisWorkbook)
    Debug.Assert Not MemberQ(Array(Empty, Null), ThisWorkbook)
    Debug.Assert MemberQ(Array(Empty, Null), Null)
    Debug.Assert MemberQ(Array(CVErr(1), 1), CVErr(1))
    Debug.Assert Not MemberQ(Array(CVErr(2), CVErr(3)), CVErr(1))
    Debug.Assert MemberQ(Array(Array(1, 2), 3), Array(1, 2))
    Debug.Assert Not MemberQ(Array(Array(1, 2), 3), 1)
End Sub

Public Sub TestPredicatesFreeQ()
    Debug.Assert Not FreeQ(Array(1, 2, 3), 2)
    Debug.Assert Not FreeQ(Array(Empty, ThisWorkbook), ThisWorkbook)
    Debug.Assert FreeQ(Array(Empty, Null), ThisWorkbook)
    Debug.Assert Not FreeQ(Array(Empty, Null), Null)
    Debug.Assert Not FreeQ(Array(CVErr(1), 1), CVErr(1))
    Debug.Assert FreeQ(Array(CVErr(2), CVErr(3)), CVErr(1))
End Sub

Public Sub TestDirectoryExistsQ()
    Debug.Print "C:\ exists is " & DirectoryExistsQ("C:\")
    Debug.Print "C:\Windows exists is " & DirectoryExistsQ("C:\Windows")
    Debug.Print "C:\Windows\Nahhhh exists is " & DirectoryExistsQ("C:\Windows\Nahhhh")
    Debug.Print "C:\Windows\Syst?m32 exists is " & DirectoryExistsQ("C:\Windows\Syst?m32")
End Sub

'********************************************************************************************
' FunctionalPredicates
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
    Debug.Assert AllTrueQ([{2,3,4}], Lambda("x", "", "x>1"))
    Debug.Assert Not AllTrueQ([{0,2,3,4}], Lambda("x", "", "x>1"))

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
    Debug.Assert AnyTrueQ(Array(2, 4, 6), Lambda("x", "", "x>1"))
    Debug.Assert AnyTrueQ(Array(2, 4, 6), Lambda("x", "", "x=4"))
    Debug.Assert Not AnyTrueQ(Array(2, 4, 6), Lambda("x", "", "x<0"))
    Debug.Assert Not AnyTrueQ(Array(2, 4, 6), Lambda("x", "", "StringQ(x)"))
    Debug.Assert AnyTrueQ(Array("a", 4, 6), Lambda("x", "", "StringQ(x)"))
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
    Debug.Assert Not NoneTrueQ(Array("a", "b", Empty), Lambda("x", "", "StringQ(x)"))
    Debug.Assert Not NoneTrueQ(Array("a", "b"), Lambda("x", "", "StringQ(x)"))
    Debug.Assert NoneTrueQ(Array(1, 2, Empty), Lambda("x", "", "StringQ(x)"))
    Debug.Assert NoneTrueQ(Array(1, 2), Lambda("x", "", "StringQ(x)"))
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
    Debug.Assert AllFalseQ([{1,2,3,4}], Lambda("x", "", "x<0"))
End Sub

Public Sub TestFunctionalPredicatesAnyFalseQ()
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
    Debug.Assert AnyFalseQ(Array(True, True), Lambda("x", "", "FalseQ(x)"))
End Sub

Public Sub TestFunctionalPredicatesNoneFalseQ()
    Dim UndimensionedArray() As Variant

    Debug.Assert NoneFalseQ(Array(1, 2, 4), "WholeNumberQ")
    Debug.Assert NoneFalseQ(Array(1, 2, 4#), "WholeNumberQ")
    Debug.Assert NoneFalseQ(Array(1, 2, 4), "NumberQ")
    Debug.Assert NoneFalseQ(Array(1, 2, 4#), "NumberQ")
    Debug.Assert Not NoneFalseQ(Array(1, 2, 4#), "StringQ")
    Debug.Assert Not NoneFalseQ(Array("a", "b", Empty), "StringQ")
    Debug.Assert NoneFalseQ(Array("a", "b"), "StringQ")
    Debug.Assert NoneFalseQ(EmptyArray(), "StringQ")
    Debug.Assert NoneFalseQ(UndimensionedArray, "StringQ")
    Debug.Assert NoneFalseQ(UndimensionedArray, "StringQ")
    Debug.Assert NoneFalseQ(Array(True, True, True))
    Debug.Assert Not NoneFalseQ(Array(True, True, False))
    Debug.Assert Not NoneFalseQ(Array(False, False, False))
    Debug.Assert NoneFalseQ([{1,2,3}], Lambda("x", "", "x<5"))
    Debug.Assert Not NoneFalseQ([{1,2,3}], Lambda("x", "", "x<3"))
End Sub

'********************************************************************************************
' Predicates
'********************************************************************************************
Public Sub TestPredicatesDimensionedQ()
    Dim a() As Variant
    Dim b(1 To 2) As Variant
    Dim c As Integer
    Dim Wbk As Workbook
    
    Debug.Assert DimensionedQ(EmptyArray())
    Debug.Assert Not DimensionedQ(a)
    Debug.Assert DimensionedQ(b)
    Debug.Assert Not DimensionedQ(c)
    Debug.Assert Not DimensionedQ(Wbk)
End Sub

Public Sub TestPredicatesEmptyArrayQ()
    Dim a() As Variant
    Dim b(1 To 2) As Variant
    Dim c As Integer
    Dim Wbk As Workbook
    
    Debug.Assert EmptyArrayQ(EmptyArray())
    Debug.Assert Not EmptyArrayQ(a)
    Debug.Assert Not EmptyArrayQ(b)
    Debug.Assert Not EmptyArrayQ(c)
    Debug.Assert Not EmptyArrayQ(Wbk)
End Sub

Public Sub TestPredicatesAtomicArrayQ()
    Dim anInteger As Integer
    Dim aDouble As Double
    Dim aDate As Date
    Dim aBoolean As Boolean
    Dim aString As String
    Dim aWorksheet As Worksheet
    Dim AWorkbook As Workbook
    Dim aListObject As ListObject
    Dim aVariant As Variant
    Dim AnArray(1 To 2) As Integer
    
    Set aWorksheet = ActiveSheet
    Set AWorkbook = ThisWorkbook
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
    Debug.Assert FileExistsQ(WorkbookFullPath(ThisWorkbook))
    Debug.Assert Not FileExistsQ(ThisWorkbook.Path & "\NotHere.xlsb")
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
Public Sub TestArrayPredicatesDimensionedNonEmptyArrayQ()
    Dim AnArray() As Integer
    
    Debug.Assert Not DimensionedNonEmptyArrayQ(AnArray)
    Debug.Assert Not DimensionedNonEmptyArrayQ(EmptyArray())
    
    ReDim AnArray(1 To 2)
    Debug.Assert DimensionedNonEmptyArrayQ(AnArray)
    
    Debug.Assert Not DimensionedNonEmptyArrayQ(1)
    Debug.Assert Not DimensionedNonEmptyArrayQ(ThisWorkbook)
End Sub

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

Public Sub TestArrayPredicatesLambdaArrayQ()
    Dim l1 As Lambda, l2 As Lambda
    
    Set l1 = Lambda("", "", """""")
    Set l2 = Lambda("", "", """""")
    
    Debug.Print "l1 -> " & LambdaQ(l1)
    Debug.Print "l2 -> " & LambdaQ(l2)
    
    Debug.Print "LambdaArrayQ(Array(l1, l2) should be True. It is: " & LambdaArrayQ(Array(l1, l2))
    Debug.Print
    Debug.Print "LambdaArrayQ(EmptyArray()) is " & LambdaArrayQ(EmptyArray())
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
    Dim ASpan As New Span
    
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

    For Each var In Array(Array(ASpan), _
                          Array(ASpan, ASpan), _
                          EmptyArray())
        Debug.Assert SpanArrayQ(var)
    Next

    Call aListObject.Delete
End Sub

Public Sub TestPredicatesPartIndexArrayQ()
    Dim var As Variant
    Dim aListObject As ListObject
    Dim aDictionary As Dictionary
    Dim ASpan As New Span
    
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

    For Each var In Array(Array(ASpan), _
                          Array(ASpan, ASpan), _
                          Array(1, ASpan), _
                          Array(1, 2, 3))
        Debug.Assert PartIndexArrayQ(var)
    Next

    Call aListObject.Delete
End Sub

Public Sub TestPredicatesTakeIndexArrayQ()
    Dim var As Variant
    Dim aListObject As ListObject
    Dim aDictionary As Dictionary
    Dim ASpan As New Span
    
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
                          Array(ASpan), _
                          Array(ASpan, ASpan), _
                          Array(1, ASpan) _
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

Public Sub TestPredicatesFormControlButtonQ()
    Dim AShape1 As Shape
    Dim AShape2 As Shape
    
    Set AShape1 = TempComputation.Shapes.AddFormControl(xlButtonControl, _
                                                        Range("B2").Left, _
                                                        Range("B2").Top, _
                                                        40, _
                                                        40)
                                                        
    Set AShape2 = TempComputation.Shapes.AddFormControl(xlDropDown, _
                                                        Range("B2").Left, _
                                                        Range("B2").Top, _
                                                        40, _
                                                        40)
                                                        
    Debug.Assert FormControlButtonQ(AShape1)
    Debug.Assert Not FormControlButtonQ(AShape2)
    Debug.Assert Not FormControlButtonQ(1)
    Debug.Assert Not FormControlButtonQ(TempComputation)
    Debug.Assert Not FormControlButtonQ(ThisWorkbook)
    Debug.Assert Not FormControlButtonQ("not")
    
    Call DeleteShapes(Array(AShape1, AShape2))
End Sub

'********************************************************************************************
' FunctionalProgramming
'********************************************************************************************
Public Sub TestFunctionalProgrammingEval()
    Dim f As Lambda

    Debug.Assert Eval(Lambda("x", "", "2*x"), 20) = 40

    Set f = Lambda("x", "", "TypeName(First(x))")
    Debug.Assert TypeName(Array(1, 2)(1)) = "Integer"
    Debug.Assert Eval(f, [{1}]) = "Double"
    Debug.Assert Eval(f, ToIntegers([{1}])) = "Integer"
End Sub

Public Sub TestFunctionalProgrammingApply()
    Debug.Assert Apply(Lambda([{"x","y"}], "", "2*x*y"), [{20,30}]) = 1200
    Debug.Assert Apply("add", [{20,30}]) = 50
End Sub

Public Sub TestFunctionalProgrammingScan()
    Call Scan("ScanHelper1", NumericalSequence(1, 10))
End Sub

Private Sub ScanHelper1(arg As Variant)
    Debug.Print "The time for interation " & arg & " is " & Now()
End Sub

Public Sub TestFunctionalProgrammingThrough()
    PrintArray Through([{"NumberQ", "StringQ","BooleanQ"}], 123)
End Sub

Public Sub TestFunctionalProgrammingMap()
    Dim a As Variant
    Dim VariantArrayOfStringArrays(1 To 2) As Variant
    
    Debug.Print "We are showing how to do MapThread using ArrayMap."
    Debug.Print
    Let a = Array(Array("1", "2", "3"), Array("4", "5", "6"))
    Debug.Print "a is:"
    PrintArray Pack2DArray(a)
    Debug.Print "All entries in a are strings."
    Debug.Print
    
    Let a = Map("StringConcatenate", a)
    Debug.Print "Let a = TransposeMatrix(Pack2DArray(Array(Array(""1"", ""2"", ""3""), Array(""4"", ""5"", ""6"")))"
    Debug.Print "Let a = Map(""StringConcatenate"", a)"
    PrintArray a
    
    Debug.Print
    Debug.Print "We are showing how to do MapThread using Map."
    Debug.Print
    Let VariantArrayOfStringArrays(1) = ToStrings(Array("1", "2", "3"))
    Let VariantArrayOfStringArrays(2) = ToStrings(Array("4", "5", "6"))
    Debug.Print "StringArray is:"
    PrintArray Pack2DArray(VariantArrayOfStringArrays)
    Debug.Print "The concatenation of each row is:"
    PrintArray Map("StringConcatenate", VariantArrayOfStringArrays)
    
    Debug.Print

    Debug.Print "We are not showing to multiply each element in " & PrintArray(NumericalSequence(1, 10), True)
    PrintArray Map("ArrayMapHelper1", NumericalSequence(1, 10))
    
    Debug.Print
    Debug.Print "We are evaluating: Map(Lambda(""x"", """", ""10*x""), NumericalSequence(1, 10))"
    PrintArray Map(Lambda("x", "", "10*x"), NumericalSequence(1, 10))
End Sub

Private Function ArrayMapHelper1(arg As Variant) As Integer
    Let ArrayMapHelper1 = arg * 10
End Function

Public Sub TestFunctionalProgrammingMapThread()
    PrintArray MapThread(Lambda([{"x1","x2", "x3", "x4"}], "", "x1+x2+x3+X4"), _
                         [{1,2,3}], _
                         [{10,20,30}], _
                         [{100,200,300}], _
                         [{1000,2000,3000}])

    Debug.Print "Test 2"
    
    PrintArray MapThread("Add", [{1,2,3,4,5}], [{10, 20, 30, 40, 50}])
End Sub

Public Sub TestFunctionalProgrammingFilter()
    PrintArray Filter(NumericalSequence(1, 100), "FilterHelper1")
    PrintArray Filter(Array("aPAPa", "bbbPAP", "sdgasdkh"), "FilterHelper2")
    PrintArray Filter(Array("aPAPa", "bbbPAP", "sdgasdkh"), Lambda("x", "", "InStr(1, x, ""PAP"")"))
End Sub

Private Function FilterHelper1(arg As Variant) As Boolean
    Let FilterHelper1 = arg > 1 And arg < 6
End Function

Private Function FilterHelper2(arg As Variant) As Boolean
    Let FilterHelper2 = InStr(1, arg, "PAP")
End Function

Public Sub TestFunctionalProgrammingTotal()
    Dim a1 As Variant
    Dim A2DArray(1 To 1, 1 To 3) As Variant
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
    For i = 1 To 3: Let A2DArray(1, i) = i: Next
    PrintArray Total(A2DArray)
    Debug.Print
End Sub

Public Sub TestFunctionalProgrammingTimes()
    Dim a1 As Variant
    Dim A2DArray(1 To 1, 1 To 3) As Variant
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
    For i = 1 To 3: Let A2DArray(1, i) = i: Next
    PrintArray Times(A2DArray)
    Debug.Print
End Sub

Public Sub TestFunctionalProgrammingAccumulate()
    Dim a As Variant
    Dim M(1 To 1, 1 To 3) As Variant
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
        Let M(1, i) = i
    Next
    Debug.Print "Testing Accumulate on [{1,2,3}] (as true 2D matrix)"
    For Each var In Accumulate(M)
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
    Dim M As Variant
    Dim a As Variant
    Dim r As Long
    Dim c As Long
    
    ' 1D Tests
    Let M = Array(1, 2, 3, 4, 5)
    Debug.Print "We have m = Array(1, 2, 3, 4, 5)"
    For i = -10 To 10
        Debug.Print "Part(m, " & i & ") = " & Part(M, i)
    Next
    
    Debug.Print
    Debug.Print "Getting Part(m, Array(1, 5))"
    PrintArray Part(M, Array(1, 5))
    
    Debug.Print
    Debug.Print "Getting Part(m, Span(1, -2))"
    PrintArray Part(M, Span(1, -2))

    Debug.Print
    Debug.Print "Getting Part(m, Span(1, -3))"
    PrintArray Part(M, Span(1, -3))
    
    Debug.Print
    Debug.Print "Getting Part(m, Span(2, -3))"
    PrintArray Part(M, Span(2, -3))
    
    Debug.Print
    ReDim M(0 To 5)
    For i = 1 To 6: Let M(i - 1) = i: Next
    Debug.Print "Setting m = Array(1, 2, 3, 4, 5, 6)"
    Debug.Print "LBound(m), UBound(m) = " & LBound(M) & ", " & UBound(M)
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
    ReDim M(1 To 3)
    For i = 1 To 10
        For j = 1 To 3
            Let M(j) = Application.WorksheetFunction.RandBetween(1, 5)
        Next
        
        Debug.Print "Rows (" & PrintArray(M, True) & ") is " & vbCr & PrintArray(Part(a, M), True)
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
    
    Let M = ConstantArray(Empty, 7, 8)
    For i = 1 To NumberOfRows(M)
        For j = 1 To NumberOfColumns(M)
            If j > 1 Then
                Let M(i, j) = 10 ^ (j - 1) * i + M(i, j - 1)
            Else
                Let M(i, j) = i
            End If
        Next
    Next
    Debug.Print "Cycling through elements of:"
    PrintArray M
    Debug.Print
    Debug.Print "Getting individual rows 1 through 6"
    For i = 1 To 6: PrintArray Part(M, i): Next
    Debug.Print
    Debug.Print "Cycling through segments Array(1, i)"
    For i = 1 To 6: PrintArray Part(M, Array(1, i)): Debug.Print: Next
    Debug.Print
    Debug.Print "Cycling through segments Array(i, -1)"
    For i = 1 To 6: PrintArray Part(M, Array(i, -1)): Debug.Print: Next
    Debug.Print
    Debug.Print "Cycling through segments Array(-6, i)"
    For i = 1 To 6: PrintArray Part(M, Array(-6, i)): Debug.Print: Next
    Debug.Print
    
    Debug.Print "Get elts 3, 5, 1"
    PrintArray Part(M, Array(3, 5, 1))
    Debug.Print
    Debug.Print "Get stepped segment Span(3, 5, 1)"
    PrintArray Part(M, Span(3, 5, 1))
    Debug.Print
    Debug.Print "Get stepped segment Span(3, 5, 2)"
    PrintArray Part(M, Span(3, 5, 2))
    Debug.Print
    Debug.Print "Get stepped segment Span(3, 5, 6)"
    PrintArray Part(M, Span(3, 5, 6))
    Debug.Print
    Debug.Print "Get segment Span(-4,-2))"
    PrintArray Part(M, Span(-4, -2))
    Debug.Print
    Debug.Print "Get segment Span(-4,-2)"
    PrintArray Part(M, Span(-4, -2))
    
    ' 2D Test with two dimensional index sets
    Let M = ConstantArray(Empty, 7, 8)
    For i = 1 To NumberOfRows(M)
        For j = 1 To NumberOfColumns(M)
            If j > 1 Then
                Let M(i, j) = 10 ^ (j - 1) * i + M(i, j - 1)
            Else
                Let M(i, j) = i
            End If
        Next
    Next
    Debug.Print "Cycling through elements of:"
    PrintArray M
    Debug.Print
    Debug.Print "Getting individual columns 1 through 6"
    For i = 1 To 6: PrintArray Part(M, Span(1, -1), i): Debug.Print: Next
    Debug.Print
    Debug.Print "Cycling through segments Span(1, -1), Span(1, i)"
    For i = 1 To 6: PrintArray Part(M, Span(1, -1), Span(1, i)): Debug.Print: Next
    Debug.Print
    Debug.Print "Cycling through segments Array(1, -1), Array(i, -1)"
    For i = 1 To 6: PrintArray Part(M, Array(1, -1), Array(i, -1)): Debug.Print: Next
    Debug.Print
    Debug.Print "Cycling through segments Array(1, -1), Array(-6, i)"
    For i = 1 To 6: PrintArray Part(M, Array(1, -1), Array(-6, i)): Debug.Print: Next
    Debug.Print
    Debug.Print "Get Array(1, -1) with Array(Array(3, 5, 1))"
    PrintArray Part(M, Array(1, -1), Array(Array(3, 5, 1)))
    Debug.Print
    Debug.Print "Get stepped segment Array(1, -1), Span(3, 5, 1)"
    PrintArray Part(M, Array(1, -1), Span(3, 5, 1))
    Debug.Print
    Debug.Print "Get stepped segment Array(1, -1),Span(3, 5, 2)"
    PrintArray Part(M, Array(1, -1), Span(3, 5, 2))
    Debug.Print
    Debug.Print "Get stepped segment Array(1, -1), Span(3, 5, 6)"
    PrintArray Part(M, Array(1, -1), Span(3, 5, 6))
    Debug.Print
    Debug.Print "Get segment Array(1, -1), Array(-4,-2))"
    PrintArray Part(M, Array(1, -1), Array(-4, -2))
    Debug.Print
    Debug.Print "Get segment Array(1, -1), Array(Array(-4,-2))"
    PrintArray Part(M, Array(1, -1), Array(Array(-4, -2)))
    Debug.Print
    Debug.Print "Getting a rectangular submatrix Span(3,4), Span(5,6,7)"
    PrintArray Part(M, Span(3, 4), Span(5, 6, 7))
    Debug.Print
    
    Debug.Print "SPEED TESTS"
    Debug.Print "Creating a large 10000 by 1000 2D array"
    Let M = ConstantArray(Empty, 10000, 1000)
    For i = 1 To 10000: For j = 1 To 1000: Let M(i, j) = i * 1000 + j: Next: Next
    Debug.Print "Accessing element 500 by 200"
    Debug.Print M(500, 200)
    Debug.Print "Accessing row 7000"
    Let a = Part(M, 7000)
    PrintArray a
    Debug.Print "Accessig column 700"
    Let a = Part(M, Span(1, -1), 700)
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

Public Sub TestArraysArraysFlatten()
    PrintArray Flatten(Array(Array(1, 2, 3), Array(4, 5, 6)))
    PrintArray Flatten([{1, 2, 3; 4, 5, 6}])
    PrintArray Flatten(Array(Array(Array(1, 2, 3), Array(Array(4, 5)), 6), 7))
End Sub

Public Sub TestArraysNumericalSequence()
    Debug.Print "Sequential array 1...10"
    PrintArray NumericalSequence(1, 10)
    Debug.Print
    
    Debug.Print "Sequential array 2...6"
    PrintArray NumericalSequence(2, 5)
    Debug.Print
    
    Debug.Print "Sequential array 2...12 step 2"
    PrintArray NumericalSequence(2, 6, 2)
    Debug.Print

    Debug.Print "Sequential array 2...2 repeated 10 times at step 0"
    PrintArray NumericalSequence(2, 10, 0)
    Debug.Print
    
    Debug.Print "Sequential array from 1 to 11"
    PrintArray NumericalSequence(1, 11, ToEndNumberQ:=True)
    
    Debug.Print "Sequential array from 1 to 11 step 3"
    PrintArray NumericalSequence(1, 11, TheStep:=3, ToEndNumberQ:=True)
    
    Debug.Print "Sequential array from 0.5, -20 numbers"
    PrintArray NumericalSequence(0.5, -20)
    
    Debug.Print "Sequential array from 0.5 to -10 step -3.1"
    PrintArray NumericalSequence(0.5, -10, TheStep:=-3.1)
    
    Debug.Print "Sequential array from 0.5 to -10 step -3.1, to end number"
    PrintArray NumericalSequence(0.5, -10, TheStep:=-3.1, ToEndNumberQ:=True)
End Sub

Public Sub TestArraysRange()
    Debug.Assert EqualQ(Range(3), Array(1, 2, 3))
End Sub

Public Sub TestArraysCreateIndexSequenceFromSpan()
    Dim ASpan As Span
    Dim AnArray() As Integer
    Dim i As Integer
    
    PrintArray CreateIndexSequenceFromSpan(NumericalSequence(1, 10), Span(1, 10))
    Debug.Print
    PrintArray CreateIndexSequenceFromSpan(NumericalSequence(1, 5), Span(1, -1))
    Debug.Print
    PrintArray CreateIndexSequenceFromSpan(NumericalSequence(1, 5), Span(2, -2))
    Debug.Print
    PrintArray CreateIndexSequenceFromSpan(NumericalSequence(1, 5), Span(2, -2, 2))
    Debug.Print
    PrintArray CreateIndexSequenceFromSpan(NumericalSequence(1, 20), Span(1, -1, 2))
    Debug.Print
    PrintArray CreateIndexSequenceFromSpan(NumericalSequence(100, 10), Span(1, 10))
    Debug.Print
    PrintArray CreateIndexSequenceFromSpan(NumericalSequence(100, 5), Span(1, -1))
    Debug.Print
    PrintArray CreateIndexSequenceFromSpan(NumericalSequence(100, 5), Span(2, -2))
    Debug.Print
    PrintArray CreateIndexSequenceFromSpan(NumericalSequence(100, 5), Span(2, -2, 2))
    Debug.Print
    PrintArray CreateIndexSequenceFromSpan(NumericalSequence(100, 20), Span(1, -1, 2))
    Debug.Print
    
    ReDim AnArray(0 To 10)
    For i = 1 To 10
        Let AnArray(i - 1) = i
    Next
    
    PrintArray CreateIndexSequenceFromSpan(AnArray, Span(1, 10))
    Debug.Print
End Sub

Public Sub TestArraysNormalizeIndicesAndNormalizeArrayIndices()
    Dim AnArray() As Integer
    Dim TheIndices As Variant
    Dim i As Integer
    
    ReDim AnArray(0 To 2)
    Let TheIndices = Array(1, 2, 3)
    Debug.Print "Array's LBound, UBound are: " & LBound(AnArray) & ", " & UBound(AnArray)
    Debug.Print "The indices are: (" & Join(TheIndices, ",") & ")"
    For i = 1 To Length(TheIndices)
        Debug.Print "   - Index = " & Part(TheIndices, i) & ", Normalized = " & NormalizeIndex(AnArray, Part(TheIndices, i))
    Next
    Debug.Print "   - NormalizedAsAnArray = (" & Join(NormalizeIndexArray(AnArray, TheIndices), ", ") & ")"
    
    Debug.Print
    
    ReDim AnArray(2 To 5)
    Let TheIndices = Array(1, 2, 3)
    Debug.Print "Array's LBound, UBound are: " & LBound(AnArray) & ", " & UBound(AnArray)
    Debug.Print "The indices are: (" & Join(TheIndices, ",") & ")"
    For i = 1 To Length(TheIndices)
        Debug.Print "   - Index " & i & " = " & Part(TheIndices, i) & ", Normalized = " & NormalizeIndex(AnArray, Part(TheIndices, i))
    Next
    Debug.Print "   - NormalizedAsAnArray = (" & Join(NormalizeIndexArray(AnArray, TheIndices), ", ") & ")"
    
    Debug.Print
End Sub

Public Sub TestArraysIndicesTuple()
    Dim a(0 To 2) As Variant
    Dim a1(0 To 1) As Variant
    Dim a2(0 To 2) As Variant
    Dim b(0 To 3, 0 To 3) As Variant
    Dim i As Integer
    Dim j As Integer
    Dim TheCount As Integer
    
    For i = 0 To 2
        Let a2(i) = i + 2
    Next
    Let a1(0) = 1
    Let a1(1) = a2
    Let a(0) = a1
    Let a(1) = 2
    PrintArray NormalizeIndicesTuple(a, Array(1, 2, 3))
    PrintArray NormalizeIndicesTuple(a, Array(1, 2, 2))
    PrintArray NormalizeIndicesTuple(a, Array(2, 2, 3))
    
    Let TheCount = 1
    For i = 1 To 3
        For j = 1 To 3
            Let b(i, j) = TheCount
            Let TheCount = TheCount + 1
        Next
    Next
    Let b(2, 1) = a2
    Let b(2, 1)(2) = [{10,20,30;40,50,60}]
    PrintArray NormalizeIndicesTuple(b, Array(3, 2, 1))
End Sub

Public Sub TestArraysSetPart()
    Dim AnArray As Variant
    Dim r As Long
    Dim c As Long

    ' EXAMPLE 1 - Acting on 1D, Individual part, 1st Dim
    Debug.Print "EXAMPLE 1 - Acting on 1D, Individual Location"
    
    Let AnArray = ConstantArray("Empty", 10)
    Debug.Print "Set individual locations in this array:"
    PrintArray AnArray
    For r = 1 To Length(AnArray)
        Let AnArray = SetPart(AnArray, r, r)
    Next
    
    Debug.Print
    Debug.Print "Every location should now have its index"
    PrintArray AnArray
    
    ' EXAMPLE 2 - Acting on 1D array of 1D arrays, 1st Dim
    Debug.Print "--------------------------------------------------------------------------"
    Debug.Print "EXAMPLE 2 - Acting on 1D, Array of Individual Locations"
    
    Let AnArray = ConstantArray("Empty", 10)
    Debug.Print "Set individual locations in this array:"
    PrintArray AnArray
    Let AnArray = SetPart(AnArray, Array(40, 60, 80), Array(4, 6, 8))
    
    Debug.Print
    Debug.Print "Every location should now have its index"
    PrintArray AnArray
    
    ' EXAMPLE 3 - Acting on 1D, Replacing Array of Individual Locations Using a Span
    Debug.Print "--------------------------------------------------------------------------"
    Debug.Print "EXAMPLE 3 - Acting on 1D, Replacing Array of Individual Locations Using a Span"
    
    Let AnArray = ConstantArray("Empty", 10)
    Debug.Print "Set individual locations in this array using Span(1, 10, 2):"
    PrintArray AnArray
    Let AnArray = SetPart(AnArray, Array(10, 30, 50, 70, 90), Span(1, 10, 2))
    
    Debug.Print
    Debug.Print "Every location should now have its index"
    PrintArray AnArray
    
    ' EXAMPLE 4 - Acting on 1D, Using a Span that is too long, 1st Dim
    Debug.Print "--------------------------------------------------------------------------"
    Debug.Print "EXAMPLE 4 - Acting on 1D, Using a Span"
    
    Let AnArray = ConstantArray("Empty", 10)
    Debug.Print "Set individual locations in this array using Span(1, 12, 2):"
    PrintArray AnArray
    Let AnArray = SetPart(AnArray, Array(10, 30, 50, 70, 90), Span(1, 12, 2))
    
    Debug.Print
    Debug.Print "The array should have turned into Null"
    PrintArray AnArray
    
    ' EXAMPLE 5 - Acting on 2D, Individual Location (for whole row)
    Debug.Print "EXAMPLE 5 - Acting on 2D, Individual Location (for whole row)"
    
    Let AnArray = ConstantArray("Empty", 10, 10)
    Debug.Print "Setting each row to a constant array of the row number"
    PrintArray AnArray
    For r = 1 To Length(AnArray)
        Let AnArray = SetPart(AnArray, ConstantArray(r, 10), r)
    Next
    
    Debug.Print
    Debug.Print "Every location should now have its index"
    PrintArray AnArray

    ' EXAMPLE 6 - Acting on 2D and replacing rows with array of 1D arrays
    Debug.Print "--------------------------------------------------------------------------"
    Debug.Print "EXAMPLE 6 - Acting on 2D, Array of Rows to Array of Rows"
    
    Let AnArray = ConstantArray("Empty", 10, 10)
    Debug.Print "Set individual locations in this array:"
    PrintArray AnArray
    Let AnArray = SetPart(AnArray, _
                          Array(ConstantArray(4, 10), _
                                ConstantArray(6, 10), _
                                ConstantArray(8, 10)), _
                          Array(4, 6, 8))
    
    Debug.Print
    Debug.Print "Every location should now have its index"
    PrintArray AnArray
    
    ' EXAMPLE 7 - Acting on 1D of 1D Matrix, replacing array of individual locations using array of ordered pairs
    Debug.Print "--------------------------------------------------------------------------"
    Debug.Print "EXAMPLE 7 - Acting on 1D of 1D Matrix, replacing array of individual locations using array of ordered pairs"
    
    Let AnArray = ConstantArray(ConstantArray("Empty", 10), 10)
    For r = 1 To 10
        For c = 1 To 10
            Let AnArray(r)(c) = 10 * r + c - 1
        Next
    Next
    Debug.Print "Set individual locations using ordered pairs:"
    PrintArray Pack2DArray(AnArray)
    Let AnArray = SetPart(AnArray, _
                          Array("4-4", "6-6", "8-8"), _
                          Array(Array(4, 4), Array(6, 6), Array(8, 8)))
    
    Debug.Print
    Debug.Print "Every location should now have its index"
    PrintArray Pack2DArray(AnArray)
    
    ' EXAMPLE 8 - Acting on 2D and replacing rows with 2D array
    Debug.Print "--------------------------------------------------------------------------"
    Debug.Print "EXAMPLE 8 - Acting on 2D and replacing rows with 2D array"
    
    Let AnArray = ConstantArray("Empty", 10, 10)
    Debug.Print "Set individual locations in this array:"
    PrintArray AnArray
    Let AnArray = SetPart(AnArray, _
                          [{4,4,4,4,4,4,4,4,4,4;6,6,6,6,6,6,6,6,6,6; 8,8,8,8,8,8,8,8,8,8}], _
                          Array(4, 6, 8))
    
    Debug.Print
    Debug.Print "Every location should now have its index"
    PrintArray AnArray
    
    ' EXAMPLE 9 - Acting on 2D, Chaging Array of Specific Locations
    Debug.Print "--------------------------------------------------------------------------"
    Debug.Print "EXAMPLE 9 - Acting on 2D, Chaging Array of Specific Locations"
    
    Let AnArray = ConstantArray("Empty", 10, 10)
    Debug.Print "Set individual locations in this array:"
    PrintArray AnArray
    Let AnArray = SetPart(AnArray, _
                          [{4,4,4,4,4,4,4,4,4,4;6,6,6,6,6,6,6,6,6,6; 8,8,8,8,8,8,8,8,8,8}], _
                          Array(4, 6, 8))
    PrintArray AnArray
    
    Let AnArray = SetPart(AnArray, _
                          Array("4-4", "6-6", "8-8"), _
                          Array(Array(4, 4), Array(6, 6), Array(8, 8)))
    Debug.Print
    Debug.Print "The array should have turned into Null"
    PrintArray AnArray
    
    ' EXAMPLE 10 - Acting on 2D and replacing rows using a Span
    Debug.Print "--------------------------------------------------------------------------"
    Debug.Print "EXAMPLE 10 - Acting on 2D and replacing rows using a Span"
    
    Let AnArray = ConstantArray("Empty", 10, 10)
    Debug.Print "Set individual locations in this array:"
    PrintArray AnArray
    Let AnArray = SetPart(AnArray, _
                          [{4,4,4,4,4,4,4,4,4,4;6,6,6,6,6,6,6,6,6,6; 8,8,8,8,8,8,8,8,8,8}], _
                          Span(4, 8, 2))
    
    Debug.Print
    Debug.Print "Every location should now have its index"
    PrintArray AnArray
    
    ' EXAMPLE 11 - Acting on 2D and replacing rows using a Span that is too long
    Debug.Print "--------------------------------------------------------------------------"
    Debug.Print "EXAMPLE 11 - Acting on 2D and replacing rows using a Span that is too long"
    
    Let AnArray = ConstantArray("Empty", 10, 10)
    Debug.Print "Set individual locations in this array:"
    PrintArray AnArray
    Let AnArray = SetPart(AnArray, _
                          [{4,4,4,4,4,4,4,4,4,4;6,6,6,6,6,6,6,6,6,6; 8,8,8,8,8,8,8,8,8,8}], _
                          Span(4, 10, 2))
    
    Debug.Print
    Debug.Print "Every location should now have its index"
    PrintArray AnArray
    
    ' EXAMPLE 12 - Acting on 1D of 1D Matrix, a single row using Span
    Debug.Print "--------------------------------------------------------------------------"
    Debug.Print "EXAMPLE 12 - Acting on 1D of 1D Matrix, a single row using Span"
    
    Let AnArray = ConstantArray(ConstantArray("Empty", 10), 10)
    For r = 1 To 10
        For c = 1 To 10
            Let AnArray(r)(c) = 10 * r + c - 1
        Next
    Next
    Debug.Print "Set individual locations using ordered pairs:"
    PrintArray Pack2DArray(AnArray)
    
    Let AnArray = SetPart(AnArray, _
                          NumericalSequence(1, 10), _
                          2, _
                          Span(1, -1))
    
    Debug.Print
    Debug.Print "Every location should now have its index"
    PrintArray Pack2DArray(AnArray)

    ' EXAMPLE 13 - Acting on 1D of 1D Matrix, a single column using Span
    Debug.Print "--------------------------------------------------------------------------"
    Debug.Print "EXAMPLE 13 - Acting on 1D of 1D Matrix, a single column using Span"
    
    Let AnArray = ConstantArray(ConstantArray("Empty", 10), 10)
    For r = 1 To 10
        For c = 1 To 10
            Let AnArray(r)(c) = 10 * r + c - 1
        Next
    Next
    Debug.Print "Set individual locations using ordered pairs:"
    PrintArray Pack2DArray(AnArray)
    
    Let AnArray = SetPart(AnArray, _
                          NumericalSequence(1, 10), _
                          Span(1, -1), _
                          2)
    
    Debug.Print
    Debug.Print "Every location should now have its index"
    PrintArray Pack2DArray(AnArray)
    
    ' EXAMPLE 14 - Acting on 1D of 1D Matrix, A Rectangle Specified by two 1D arrays of indices, Replaced by a 2D Array
    Debug.Print "--------------------------------------------------------------------------"
    Debug.Print "EXAMPLE 14 - Acting on 1D of 1D Matrix, A Rectangle Specified by two 1D arrays of indices, Replaced by a 2D Array"
    
    Let AnArray = ConstantArray(ConstantArray("Empty", 10), 10)
    For r = 1 To 10
        For c = 1 To 10
            Let AnArray(r)(c) = 10 * r + c - 1
        Next
    Next
    Debug.Print "Set individual locations using ordered pairs:"
    PrintArray Pack2DArray(AnArray)
    
    Let AnArray = SetPart(AnArray, _
                          [{"1-*", "2-*", "3-*"; "10-*", "20-*", "30-*"}], _
                          Array(2, 3), _
                          Array(4, 5, 6))
    
    Debug.Print
    Debug.Print "Every location should now have its index"
    PrintArray Pack2DArray(AnArray)

    ' EXAMPLE 15 - Acting on 1D of 1D Matrix, A Rectangle Specified by two Spans, Replaced by a 2D Array
    Debug.Print "--------------------------------------------------------------------------"
    Debug.Print "EXAMPLE 15 - Acting on 1D of 1D Matrix, A Rectangle Specified by two Spanss, Replaced by a 2D Array"
    
    Let AnArray = ConstantArray(ConstantArray("Empty", 10), 10)
    For r = 1 To 10
        For c = 1 To 10
            Let AnArray(r)(c) = 10 * r + c - 1
        Next
    Next
    Debug.Print "Set individual locations using ordered pairs:"
    PrintArray Pack2DArray(AnArray)
    
    Let AnArray = SetPart(AnArray, _
                          [{"1-*", "2-*", "3-*"; "10-*", "20-*", "30-*"}], _
                          Span(2, 3), _
                          Span(4, 5, 6))
    
    Debug.Print
    Debug.Print "Every location should now have its index"
    PrintArray Pack2DArray(AnArray)
    
    ' EXAMPLE 16 - Acting on 1D of 1D Matrix, A Rectangle Specified by two 1D arrays of indices, Replaced by a 1D of 1D Array
    Debug.Print "--------------------------------------------------------------------------"
    Debug.Print "EXAMPLE 16 - Acting on 1D of 1D Matrix, A Rectangle Specified by two 1D arrays of indices, Replaced by a 2D Array"
    
    Let AnArray = ConstantArray(ConstantArray("Empty", 10), 10)
    For r = 1 To 10
        For c = 1 To 10
            Let AnArray(r)(c) = 10 * r + c - 1
        Next
    Next
    Debug.Print "Set individual locations using ordered pairs:"
    PrintArray Pack2DArray(AnArray)
    
    Let AnArray = SetPart(AnArray, _
                          Array(Array("1-*", "2-*", "3-*"), Array("10-*", "20-*", "30-*")), _
                          Array(2, 3), _
                          Array(4, 6))
    
    Debug.Print
    Debug.Print "Every location should now have its index"
    PrintArray Pack2DArray(AnArray)
    
    ' EXAMPLE 17 - Acting on 1D of 1D Matrix, A Rectangle Specified by two Span, Replaced by a 1D of 1D Array
    Debug.Print "--------------------------------------------------------------------------"
    Debug.Print "EXAMPLE 17 - Acting on 1D of 1D Matrix, A Rectangle Specified by two 1D arrays of indices, Replaced by a 2D Array"
    
    Let AnArray = ConstantArray(ConstantArray("Empty", 10), 10)
    For r = 1 To 10
        For c = 1 To 10
            Let AnArray(r)(c) = 10 * r + c - 1
        Next
    Next
    Debug.Print "Set individual locations using ordered pairs:"
    PrintArray Pack2DArray(AnArray)
    
    Let AnArray = SetPart(AnArray, _
                          Array(Array("1-*", "2-*", "3-*"), Array("10-*", "20-*", "30-*")), _
                          Span(2, 3), _
                          Span(4, 6))
    
    Debug.Print
    Debug.Print "Every location should now have its index"
    PrintArray Pack2DArray(AnArray)
End Sub

Public Sub TestReorderColumns()
    Dim A2DArray As Variant
    
    Let A2DArray = [{"Col1", "Col2", "Col3", "Col4"; 1,2,3,4;10,20,30,40;100,200,300,400}]
    
    PrintArray ReorderColumns(A2DArray, Array("Col4", "Col3", "Col2", "Col1"))
    Debug.Print
    PrintArray ReorderColumns(A2DArray, Array("Col3", "Col1", "Col2", "Col4"))
End Sub

'********************************************************************************************
' JSON
'*******************************************************************************************

Public Sub TestJsonConverter()
    Dim JSON As Object
    Set JSON = JsonConverter.ParseJson("{""a"":123,""b"":[1,2,3,4],""c"":{""d"":456}}")
    
    ' Json("a") -> 123
    ' Json("b")(2) -> 2
    ' Json("c")("d") -> 456
    
    Let JSON("c")("e") = 789
    
    Debug.Print JsonConverter.ConvertToJson(JSON)
    ' -> "{"a":123,"b":[1,2,3,4],"c":{"d":456,"e":789}}"
    
    Debug.Print JsonConverter.ConvertToJson(JSON, Whitespace:=2)
End Sub

'********************************************************************************************
' Dictionaries
'*******************************************************************************************

Public Sub TestDictionariesTranslateUsingDictionary()
    Dim ADict As Dictionary
    Dim i As Integer
    Dim AnArray As Variant
    
    Set ADict = New Dictionary
    
    For i = 1 To 10
        Call ADict.Add(Key:=i, Item:=i ^ 2)
    Next
    
    Let AnArray = NumericalSequence(1, 10)
    Debug.Print "The array is:"
    PrintArray AnArray
    
    Let AnArray = TranslateUsingDictionary(AnArray, ADict)
    
    PrintArray AnArray
End Sub

'********************************************************************************************
' ExcelSql
'********************************************************************************************

Public Sub TestSelectUsingSql()
    Dim wsht1 As Worksheet
    Dim wsht2 As Worksheet
    Dim headers1 As Variant
    Dim headers2 As Variant
    Dim data1 As Variant
    Dim data2 As Variant
    Dim SqlQuery As String
    
    Let Application.ScreenUpdating = False
    Let Application.DisplayAlerts = False
    
    Set wsht1 = ThisWorkbook.Worksheets.Add
    Set wsht2 = ThisWorkbook.Worksheets.Add
    
    Let headers1 = Array("ID", "COL1")
    Let headers2 = Array("ID", "COL2")
    
    Let data1 = [{"ID1", "a1"; "ID2", "a2"; "ID3", "a3"; "ID4", "a4"}]
    Let data2 = [{"ID0", "b0"; "ID1", "b1"; "ID2", "b2"; "ID3", "b3"}]
    
    Call DumpInSheet(headers1, wsht1.Range("A1"))
    Call DumpInSheet(data1, wsht1.Range("A2"))
    
    Call DumpInSheet(headers2, wsht2.Range("A1"))
    Call DumpInSheet(data2, wsht2.Range("A2"))
    
    Let SqlQuery = "SELECT * FROM [" & wsht1.Name & "$];"
    Debug.Print SqlQuery
    PrintArray ExcelSql.SelectUsingSql(SqlQuery, ThisWorkbook.Path & "\" & ThisWorkbook.Name)
    
    Debug.Print
    Debug.Print
    
    Let SqlQuery = "SELECT * FROM [" & wsht2.Name & "$];"
    Debug.Print SqlQuery
    PrintArray ExcelSql.SelectUsingSql(SqlQuery, ThisWorkbook.Path & "\" & ThisWorkbook.Name)
    
    Debug.Print
    Debug.Print

    Let SqlQuery = "SELECT * FROM [" & wsht1.Name & "$] as wsht1,[" & wsht2.Name & "$] as wsht2;"
    Debug.Print SqlQuery
    PrintArray ExcelSql.SelectUsingSql(SqlQuery, ThisWorkbook.Path & "\" & ThisWorkbook.Name)
    
    Debug.Print
    Debug.Print
    
    Let SqlQuery = "SELECT t1.ID, t2.COL2 FROM [" & wsht1.Name & "$] AS t1 LEFT OUTER JOIN [" & wsht2.Name & "$] AS t2 ON t1.ID=t2.ID;"
    Debug.Print SqlQuery
    PrintArray ExcelSql.SelectUsingSql(SqlQuery, ThisWorkbook.Path & "\" & ThisWorkbook.Name)

    Debug.Print
    Debug.Print

    Let SqlQuery = "SELECT t1.ID, t2.COL2 FROM [" & wsht1.Name & "$] AS t1 RIGHT OUTER JOIN [" & wsht2.Name & "$] AS t2 ON t1.ID=t2.ID;"
    Debug.Print SqlQuery
    PrintArray ExcelSql.SelectUsingSql(SqlQuery, ThisWorkbook.Path & "\" & ThisWorkbook.Name)

    Debug.Print
    
    Let SqlQuery = "SELECT wsht1.ID, wsht1.COL1, wsht2.COL2 FROM [" & wsht1.Name & "$] as wsht1,[" & wsht2.Name & "$] as wsht2 WHERE wsht1.ID = wsht2.ID;"
    Debug.Print SqlQuery
    PrintArray ExcelSql.SelectUsingSql(SqlQuery, ThisWorkbook.Path & "\" & ThisWorkbook.Name)
    
    Debug.Print
    Debug.Print
    
    Let SqlQuery = "SELECT t1.ID, t1.COL1, t2.COL2 FROM [" & wsht1.Name & "$] AS t1 INNER JOIN [" & wsht2.Name & "$] AS t2 ON t1.ID=t2.ID;"
    Debug.Print SqlQuery
    PrintArray ExcelSql.SelectUsingSql(SqlQuery, ThisWorkbook.Path & "\" & ThisWorkbook.Name)
    
    Debug.Print
    Debug.Print
    
    Let SqlQuery = "SELECT t1.ID, t2.COL2 FROM [" & wsht1.Name & "$] AS t1 LEFT JOIN [" & wsht2.Name & "$] AS t2 ON t1.ID=t2.ID;"
    Debug.Print SqlQuery
    PrintArray ExcelSql.SelectUsingSql(SqlQuery, ThisWorkbook.Path & "\" & ThisWorkbook.Name)
    
    Debug.Print
    Debug.Print
    
    Let SqlQuery = "SELECT t1.ID, t2.COL2 FROM [" & wsht1.Name & "$] AS t1 RIGHT JOIN [" & wsht2.Name & "$] AS t2 ON t1.ID=t2.ID;"
    Debug.Print SqlQuery
    PrintArray ExcelSql.SelectUsingSql(SqlQuery, ThisWorkbook.Path & "\" & ThisWorkbook.Name)
    
    Call wsht1.Delete
    Call wsht2.Delete
    
    Let Application.ScreenUpdating = True
    Let Application.DisplayAlerts = True
End Sub

'********************************************************************************************
' ListObjectsModule
'********************************************************************************************

Public Sub TestAddListObject()
    Dim lo As ListObject
    
    Call ToTemp([{"Col1", "Col2"; 11,12;21,22;31,32;41,42}])
    
    Set lo = AddListObject(TempComputation.Range("A1"))
    
    Call MsgBox("Inspect TempComputation. Entire CurrentRange should be a listobject")
    
    Call ToTemp([{"Col1", "Col2"; 11,12;21,22;31,32;41,42}])
    
    Set lo = AddListObject(TempComputation.Range("A1").Resize(3, 1), UseCurrentRegionQ:=False)
    Call MsgBox("Inspect TempComputation. Only 1st three rows and column 1 should be the listobject")
End Sub

'********************************************************************************************
' StringFormulas
'********************************************************************************************
Public Sub TestVbaCodeManipulationMakeRoutineName()
    Debug.Assert "'ExcelLibraryV9.0.xlsb'!MyMod.MyFunc" = _
                 MakeRoutineName(ThisWorkbook, "MyMod", "MyFunc")
End Sub

'********************************************************************************************
' Documentation
'********************************************************************************************
Public Sub TestDocumentationGetReferences()
    Dim RefsDict As Dictionary
    Dim RefDict As Dictionary
    Dim i As Integer
    
    Set RefsDict = GetReferences(ThisWorkbook)
    
    Debug.Print "We got " & RefsDict.Count & " non-builtin references."
    For i = 0 To RefsDict.Count - 1
        Set RefDict = RefsDict.Items(i)
        
        Debug.Print
        Debug.Print "Name: " & RefDict.Keys(0)
        Debug.Print "Description: " & RefDict.Item(Key:="Description")
        Debug.Print "Description: " & RefDict.Item(Key:="FullPath")
    Next
End Sub

'********************************************************************************************
' TypeConversions
'********************************************************************************************
Public Sub TestTypeConversionsCastShapesToArray()
    Call CreateButtonsGrid(TempComputation, Range("B2"), _
                           3, 10, 10, 40, 40, _
                           Array("B1", "B2", "B3", "B4", "B5"), _
                           Array("Backup", "Backup", "Backup", "Backup", "Backup"))
                           
    Call MsgBox("We will now check that 5 buttons were created.")

    Debug.Assert TempComputation.Shapes.Count = 5

    Call UI.DeleteShapes(CastShapesToArray(TempComputation.Shapes))
    
    Call MsgBox("We will now check that 5 buttons were deleted.")
    
    Debug.Assert TempComputation.Shapes.Count = 0
End Sub

'********************************************************************************************
' UI
'********************************************************************************************
Public Sub TestsUiEqualizeFormControlButtonSizes()
    Dim AShape As Shape

    Call CreateButtonsGrid(TempComputation, Range("B2"), _
                           3, 10, 10, 40, 40, _
                           Array("B1", "B2", "B3", "B4", "B5"), _
                           Array("Backup", "Backup", "Backup", "Backup", "Backup"))
  
    Call EqualizeFormControlButtonSizes(CastShapesToArray(TempComputation.Shapes), 50, 60)
    
    For Each AShape In TempComputation.Shapes
        Debug.Assert AShape.Width = 50
        Debug.Assert AShape.Height = 60
    Next
    
    Call UI.DeleteShapes(CastShapesToArray(TempComputation.Shapes))
End Sub

Public Sub TestUiDistributeButtonsHorizontally()
    Call DeleteShapes(CastShapesToArray(TempComputation.Shapes))

    Call CreateButtonsGrid(TempComputation, Range("B2"), _
                           1, 10, 10, 40, 40, _
                           Array("B1", "B2", "B3"), _
                           Array("Backup", "Backup", "Backup"))

    Call DistributeButtonsHorizontally(CastShapesToArray(TempComputation.Shapes), TempComputation.Range("B5:E10"))
End Sub


Public Sub TestUiDistributeButtonsHorizontally2()
    Call DeleteShapes(CastShapesToArray(TempComputation.Shapes))
    
    Call CreateButtonsGrid(TempComputation, Range("B2"), _
                           1, 10, 10, 40, 40, _
                           Array("B1", "B2", "B3"), _
                           Array("Backup", "Backup", "Backup"))

    Let TempComputation.Shapes(1).Width = 70
    Call DistributeButtonsHorizontally(CastShapesToArray(TempComputation.Shapes), TempComputation.Range("B5:E10"))
End Sub

Public Sub TestUiDistributeButtonsHorizontally3()
    Call DeleteShapes(CastShapesToArray(TempComputation.Shapes))

    Call CreateButtonsGrid(TempComputation, Range("B2"), _
                           1, 10, 10, 40, 40, _
                           Array("B1", "B2", "B3"), _
                           Array("Backup", "Backup", "Backup"))

    Let TempComputation.Shapes(1).Width = 200
    Call DistributeButtonsHorizontally(CastShapesToArray(TempComputation.Shapes), TempComputation.Range("B5:E10"))
End Sub

'********************************************************************************************
' Statistics
'********************************************************************************************
Public Sub TestStatisticsMaximum()
    Debug.Assert 4 = Maximum([{1,2,3,4}])
    Debug.Assert IsNull(Maximum([{1,2,3,"4"}]))
End Sub

Public Sub TestStatisticsMinimum()
    Debug.Assert 1 = Minimum([{1,2,3,4}])
    Debug.Assert IsNull(Minimum([{1,2,3,"4"}]))
End Sub

Public Sub TestStatisticsMedian()
    Debug.Assert Application.Median([{1,2,3,4}]) = Median([{1,2,3,4}])
    Debug.Assert IsNull(Median([{1,2,3,"4"}]))
End Sub

Public Sub TestStatisticsAverage()
    Debug.Assert Application.Average([{1,2,3,4}]) = Average([{1,2,3,4}])
    Debug.Assert IsNull(Average([{1,2,3,"4"}]))
End Sub

Public Sub TestStatisticsRound()
    Dim i As Integer

    Dim Dividends As Variant
    Dim Divisors As Variant
    Dim MathematicaResults As Variant
    Dim Rounds As Variant
    
    Let Dividends = Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10)
    Let Divisors = Array(1, 5, 0.1, 0.2, 3, 1.5, 0.7, 0.8, 1.2, 2.1)
    ' Following are the results from running this function on Mathematica 11.0
    Let MathematicaResults = Array(1, 0, 3#, 4#, 6, 6#, 7#, 8#, 9.6, 10.5)
    
    Let Rounds = MapThread("Round", Dividends, Divisors)

    For i = 1 To Length(MathematicaResults)
        Debug.Assert Part(MathematicaResults, i) = Part(Rounds, i)
    Next
End Sub

Public Sub TestStatisticsStemAndLeafDiagram()
    Dim DataSet As Variant
    Dim DiagDict As Dictionary
    Dim var As Variant
    
    Let DataSet = [{11,12,13,14,21,51,32,43,72,73,74}]
    Set DiagDict = StemAndLeafDiagram(DataSet, 1)
    For Each var In DiagDict.Keys
        Debug.Print var & " -> " & DiagDict.Item(Key:=var)
    Next
End Sub


'********************************************************************************************
' Files
'********************************************************************************************
Public Sub TestGetDirectories()
    Dim aFileName As Variant
    
    Debug.Print "Without path"
    For Each aFileName In GetDirectories("c:\")
        Debug.Print aFileName
    Next
    
    Debug.Print

    Debug.Print "With path"
    For Each aFileName In GetDirectories("c:\", True)
        Debug.Print aFileName
    Next
    
    Debug.Print
    Debug.Print "Looking for a directory that does not exists"
    If EmptyArrayQ(GetDirectories("c:\TestDir\", True)) Then
        Debug.Print "There are no directories inside that one."
    ElseIf NullQ(GetDirectories("c:\TestDir\", True)) Then
        Debug.Print "The directory does not exists."
    Else
        For Each aFileName In GetDirectories("c:\TestDir\", True)
            Debug.Print aFileName
        Next
    End If
    
    Debug.Print
    Debug.Print "Looking for a directory that exists but is empty"
    If EmptyArrayQ(GetDirectories("c:\TestMe\", True)) Then
        Debug.Print "There are no directories inside that one."
    ElseIf NullQ(GetDirectories("c:\TestMe\", True)) Then
        Debug.Print "The directory does not exists."
    Else
        For Each aFileName In GetDirectories("c:\TestMe\", True)
            Debug.Print aFileName
        Next
    End If
    
    Debug.Print
    Debug.Print "Pull a subset of subdirectories using wildchars."
    For Each aFileName In GetDirectories("c:\Windows\Sy?tem*\", True)
        Debug.Print aFileName
    Next
End Sub

Public Sub TestGetFileNames()
    Dim aFileName As Variant
    
    Debug.Print "Get all filenames in C:\, prepending path"
    For Each aFileName In GetFileNames("c:\", True)
        Debug.Print aFileName
    Next
    
    Debug.Print
    Debug.Print "Get all filenames in C:\h*, do not prepend path"
    For Each aFileName In GetFileNames("c:\h*")
        Debug.Print aFileName
    Next
    
    Debug.Print
    Debug.Print "Get all filenames in C:\h*, prepend path"
    For Each aFileName In GetFileNames("c:\h*", True)
        Debug.Print aFileName
    Next
    
    Debug.Print
    Debug.Print "Get all filenames in C:\Windows\*.ini, prepend path"
    For Each aFileName In GetFileNames("C:\Windows\*.ini", True)
        Debug.Print aFileName
    Next
    
    Debug.Print
    Debug.Print "Get all filenames in C:\Windows\*.ini, do not prepend path"
    For Each aFileName In GetFileNames("C:\Windows\*.ini", False)
        Debug.Print aFileName
    Next
End Sub

