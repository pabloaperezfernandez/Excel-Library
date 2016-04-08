Attribute VB_Name = "TestedLibrary"
Option Explicit
Option Base 1

'********************************************************************************************
' Miscellaneous VBA
'********************************************************************************************
Public Sub TestArrayParamUsage()
    Call VarParamFunction
    Debug.Print
    Call VarParamFunction(0)
    Debug.Print
    Call VarParamFunction(0, 1)
    Debug.Print
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
Public Sub TestPredicatesAllTrueQ()
    Dim UndimensionedArray() As Variant

    Debug.Print "AllTrueQ(Array(1,2,4), ""WholeNumberQ"") is " & AllTrueQ(Array(1, 2, 4), ThisWorkbook, "WholeNumberQ")
    Debug.Print "AllTrueQ(Array(1,2,4.0), ""WholeNumberQ"") is " & AllTrueQ(Array(1, 2, 4#), ThisWorkbook, "WholeNumberQ")
    Debug.Print "AllTrueQ(Array(1,2,4), ""NumberQ"") is " & AllTrueQ(Array(1, 2, 4), ThisWorkbook, "NumberQ")
    Debug.Print "AllTrueQ(Array(1,2,4.0), ""NumberQ"") is " & AllTrueQ(Array(1, 2, 4#), ThisWorkbook, "NumberQ")
    Debug.Print "AllTrueQ(Array(1,2,4.0), ""StringQ"") is " & AllTrueQ(Array(1, 2, 4#), ThisWorkbook, "StringQ")
    Debug.Print "AllTrueQ(Array(""a"", ""b"",Empty), ""StringQ"") is " & AllTrueQ(Array("a", "b", Empty), ThisWorkbook, "StringQ")
    Debug.Print "AllTrueQ(Array(""a"", ""b""), ""StringQ"") is " & AllTrueQ(Array("a", "b"), ThisWorkbook, "StringQ")
    Debug.Print "AllTrueQ(EmptyArray(), ""StringQ"") is " & AllTrueQ(EmptyArray(), ThisWorkbook, "StringQ")
    Debug.Print "AllTrueQ(UndimensionedArray, ""StringQ"") is " & AllTrueQ(UndimensionedArray, ThisWorkbook, "StringQ")
End Sub

Public Sub TestPredicatesAnyTrueQ()
    Dim UndimensionedArray() As Variant

    Debug.Print "AnyTrueQ(Array(1,2,4), ThisWorkbook, ""WholeNumberQ"") is " & AnyTrueQ(Array(1, 2, 4), ThisWorkbook, "WholeNumberQ")
    Debug.Print "AnyTrueQ(Array(1,2,4.0), ThisWorkbook, ""WholeNumberQ"") is " & AnyTrueQ(Array(1, 2, 4#), ThisWorkbook, "WholeNumberQ")
    Debug.Print "AnyTrueQ(Array(1,2,4), ThisWorkbook, ""NumberQ"") is " & AnyTrueQ(Array(1, 2, 4), ThisWorkbook, "NumberQ")
    Debug.Print "AnyTrueQ(Array(1,2,4.0), ThisWorkbook, ""NumberQ"") is " & AnyTrueQ(Array(1, 2, 4#), ThisWorkbook, "NumberQ")
    Debug.Print "AnyTrueQ(Array(1,2,4.0), ThisWorkbook, ""StringQ"") is " & AnyTrueQ(Array(1, 2, 4#), ThisWorkbook, "StringQ")
    Debug.Print "AnyTrueQ(Array(""a"", ""b"",Empty), ThisWorkbook, ""StringQ"") is " & AnyTrueQ(Array("a", "b", Empty), ThisWorkbook, "StringQ")
    Debug.Print "AnyTrueQ(Array(""a"", ""b""), ThisWorkbook, ""StringQ"") is " & AnyTrueQ(Array("a", "b"), ThisWorkbook, "StringQ")
    Debug.Print "AnyTrueQ(EmptyArray(), ThisWorkbook, ""StringQ"") is " & AnyTrueQ(EmptyArray(), ThisWorkbook, "StringQ")
    Debug.Print "AnyTrueQ(UndimensionedArray, ThisWorkbook, ""StringQ"") is " & AnyTrueQ(UndimensionedArray, ThisWorkbook, "StringQ")
End Sub

Public Sub TestPredicatesNoneTrueQ()
    Dim UndimensionedArray() As Variant

    Debug.Print "NoneTrueQ(Array(1,2,4), ThisWorkbook, ""WholeNumberQ"") is " & NoneTrueQ(Array(1, 2, 4), ThisWorkbook, "WholeNumberQ")
    Debug.Print "NoneTrueQ(Array(1,2,4.0), ThisWorkbook, ""WholeNumberQ"") is " & NoneTrueQ(Array(1, 2, 4#), ThisWorkbook, "WholeNumberQ")
    Debug.Print "NoneTrueQ(Array(1,2,4), ThisWorkbook, ""NumberQ"") is " & NoneTrueQ(Array(1, 2, 4), ThisWorkbook, "NumberQ")
    Debug.Print "NoneTrueQ(Array(1,2,4.0), ThisWorkbook, ""NumberQ"") is " & NoneTrueQ(Array(1, 2, 4#), ThisWorkbook, "NumberQ")
    Debug.Print "NoneTrueQ(Array(1,2,4.0), ThisWorkbook, ""StringQ"") is " & NoneTrueQ(Array(1, 2, 4#), ThisWorkbook, "StringQ")
    Debug.Print "NoneTrueQ(Array(""a"", ""b"",Empty), ThisWorkbook, ""StringQ"") is " & NoneTrueQ(Array("a", "b", Empty), ThisWorkbook, "StringQ")
    Debug.Print "NoneTrueQ(Array(""a"", ""b""), ThisWorkbook, ""StringQ"") is " & NoneTrueQ(Array("a", "b"), ThisWorkbook, "StringQ")
    Debug.Print "NoneTrueQ(EmptyArray(), ThisWorkbook, ""StringQ"") is " & NoneTrueQ(EmptyArray(), ThisWorkbook, "StringQ")
    Debug.Print "NoneTrueQ(UndimensionedArray, ThisWorkbook, ""StringQ"") is " & NoneTrueQ(UndimensionedArray, ThisWorkbook, "StringQ")
End Sub

Public Sub TestPredicatesDimensionedQ()
    Dim A() As Variant
    Dim B(1 To 2) As Variant
    Dim c As Integer
    Dim wbk As Workbook
    
    Debug.Print "EmptyArray() is dimensioned is " & DimensionedQ(EmptyArray())
    Debug.Print "a() is dimensioned is " & DimensionedQ(A)
    Debug.Print "b(1 To 2) is dimensioned is " & DimensionedQ(B)
    Debug.Print "c is an integer is dimensioned is " & DimensionedQ(c)
    Debug.Print "wbk is dimensioned is " & DimensionedQ(wbk)
End Sub

Public Sub TestPredicatesEmptyArrayQ()
    Dim A() As Variant
    Dim B(1 To 2) As Variant
    Dim c As Integer
    Dim wbk As Workbook
    
    Debug.Print "EmptyArray() is EmptyArrayQ is " & EmptyArrayQ(EmptyArray())
    Debug.Print "a() is EmptyArrayQ is " & EmptyArrayQ(A)
    Debug.Print "b(1 To 2) is EmptyArrayQ is " & EmptyArrayQ(B)
    Debug.Print "c is an integer is EmptyArrayQ is " & EmptyArrayQ(c)
    Debug.Print "wbk is EmptyArrayQ is " & EmptyArrayQ(wbk)
End Sub

Public Sub TestPredicatesAtomicQ()
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
    
    Let aVariant = 1
    
    Debug.Print "anInteger is AtomicQ is " & AtomicQ(anInteger)
    Debug.Print "aDouble is AtomicQ is " & AtomicQ(aDouble)
    Debug.Print "aDate is AtomicQ is " & AtomicQ(aDate)
    Debug.Print "aBoolean is AtomicQ is " & AtomicQ(aBoolean)
    Debug.Print "aString is AtomicQ is " & AtomicQ(aString)
    Debug.Print "CVErr(1) is AtomicQ is " & AtomicQ(CVErr(1))
    Debug.Print "aWorksheet is AtomicQ is " & AtomicQ(aWorksheet)
    Debug.Print "aWorkbook is AtomicQ is " & AtomicQ(aWorkbook)
    Debug.Print "aListObject is AtomicQ is " & AtomicQ(aListObject)
    Debug.Print "aVariant is AtomicQ is " & AtomicQ(aVariant)
    Debug.Print "anArray is AtomicQ is " & AtomicQ(AnArray)

    Call aListObject.Delete
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
    
    For Each aVariant In Array(Array(anInteger, aDouble), _
                               Array(aDate, aString), _
                               Array(EmptyArray(), 1), _
                               EmptyArray(), _
                               Array(Null, Empty), _
                               Array(Nothing, 1))
        Debug.Print "Test is " & AtomicArrayQ(aVariant)
    Next

    Call aListObject.Delete
End Sub

Public Sub TestPredicatesNumberQ()
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
    
    Let aVariant = 1
    
    Debug.Print "anInteger is NumberQ is " & NumberQ(anInteger)
    Debug.Print "aDouble is NumberQ is " & NumberQ(aDouble)
    Debug.Print "aDate is NumberQ is " & NumberQ(aDate)
    Debug.Print "aBoolean is NumberQ is " & NumberQ(aBoolean)
    Debug.Print "aString is NumberQ is " & NumberQ(aString)
    Debug.Print "CVErr(1) is NumberQ is " & NumberQ(CVErr(1))
    Debug.Print "aWorksheet is NumberQ is " & NumberQ(aWorksheet)
    Debug.Print "aWorkbook is NumberQ is " & NumberQ(aWorkbook)
    Debug.Print "aListObject is NumberQ is " & NumberQ(aListObject)
    Debug.Print "aVariant is NumberQ is " & NumberQ(aVariant)
    Debug.Print "anArray is NumberQ is " & NumberQ(AnArray)

    Call aListObject.Delete
End Sub

Public Sub TestPredicatesNumberArrayQ()
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
    
    For Each aVariant In Array(Array(anInteger, aDouble), _
                               Array(aDate, aString), _
                               Array(EmptyArray(), 1), _
                               EmptyArray(), _
                               Array(Null, Empty), _
                               Array(Nothing, 1))
        Debug.Print "Test is " & NumberArrayQ(aVariant)
    Next

    Call aListObject.Delete
End Sub

Public Sub TestPredicatesWholeNumberQ()
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
    
    Let aVariant = 1
    
    Debug.Print "anInteger is WholeNumberQ is " & WholeNumberQ(anInteger)
    Debug.Print "aDouble is WholeNumberQ is " & WholeNumberQ(aDouble)
    Debug.Print "aDate is WholeNumberQ is " & WholeNumberQ(aDate)
    Debug.Print "aBoolean is WholeNumberQ is " & WholeNumberQ(aBoolean)
    Debug.Print "aString is WholeNumberQ is " & WholeNumberQ(aString)
    Debug.Print "CVErr(1) is WholeNumberQ is " & WholeNumberQ(CVErr(1))
    Debug.Print "aWorksheet is WholeNumberQ is " & WholeNumberQ(aWorksheet)
    Debug.Print "aWorkbook is WholeNumberQ is " & WholeNumberQ(aWorkbook)
    Debug.Print "aListObject is WholeNumberQ is " & WholeNumberQ(aListObject)
    Debug.Print "aVariant is WholeNumberQ is " & WholeNumberQ(aVariant)
    Debug.Print "anArray is WholeNumberQ is " & WholeNumberQ(AnArray)

    Call aListObject.Delete
End Sub

Public Sub TestPredicatesWholeNumberArrayQ()
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
    
    For Each aVariant In Array(Array(1, 2), _
                               Array(anInteger, aDouble), _
                               Array(aDate, aString), _
                               Array(EmptyArray(), 1), _
                               EmptyArray(), _
                               Array(Null, Empty), _
                               Array(Nothing, 1))
        Debug.Print "Test is " & WholeNumberArrayQ(aVariant)
    Next

    Call aListObject.Delete
End Sub

Public Sub TestPredicatesPositiveWholeNumberQ()
    Dim anInteger As Integer
    Dim aNegativeInteger As Integer
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
    
    Let aVariant = 1
    Let aNegativeInteger = -1
    
    Debug.Print "anInteger is PositiveWholeNumberQ is " & PositiveWholeNumberQ(anInteger)
    Debug.Print "aDouble is PositiveWholeNumberQ is " & PositiveWholeNumberQ(aDouble)
    Debug.Print "aDate is PositiveWholeNumberQ is " & PositiveWholeNumberQ(aDate)
    Debug.Print "aBoolean is PositiveWholeNumberQ is " & PositiveWholeNumberQ(aBoolean)
    Debug.Print "aString is PositiveWholeNumberQ is " & PositiveWholeNumberQ(aString)
    Debug.Print "CVErr(1) is PositiveWholeNumberQ is " & PositiveWholeNumberQ(CVErr(1))
    Debug.Print "aWorksheet is PositiveWholeNumberQ is " & PositiveWholeNumberQ(aWorksheet)
    Debug.Print "aWorkbook is PositiveWholeNumberQ is " & PositiveWholeNumberQ(aWorkbook)
    Debug.Print "aListObject is PositiveWholeNumberQ is " & PositiveWholeNumberQ(aListObject)
    Debug.Print "aVariant is PositiveWholeNumberQ is " & PositiveWholeNumberQ(aVariant)
    Debug.Print "anArray is PositiveWholeNumberQ is " & PositiveWholeNumberQ(AnArray)
    Debug.Print "aNegativeInteger is PositiveWholeNumberQ is " & PositiveWholeNumberQ(aNegativeInteger)

    Call aListObject.Delete
End Sub

Public Sub TestPredicatesPositiveWholeNumberArrayQ()
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
    
    For Each aVariant In Array(Array(1, 2), _
                               Array(anInteger, aDouble), _
                               Array(aDate, aString), _
                               Array(EmptyArray(), 1), _
                               EmptyArray(), _
                               Array(Null, Empty), _
                               Array(Nothing, 1), _
                               Array(-1, 2, 1))
        Debug.Print "Test is " & PositiveWholeNumberArrayQ(aVariant)
    Next

    Call aListObject.Delete
End Sub

Public Sub TestPredicatesNegativeWholeNumberQ()
    Dim anInteger As Integer
    Dim aNegativeInteger As Integer
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
    
    Let aVariant = 1
    Let aNegativeInteger = -1
    
    Debug.Print "anInteger is NegativeWholeNumberQ is " & NegativeWholeNumberQ(anInteger)
    Debug.Print "aDouble is NegativeWholeNumberQ is " & NegativeWholeNumberQ(aDouble)
    Debug.Print "aDate is NegativeWholeNumberQ is " & NegativeWholeNumberQ(aDate)
    Debug.Print "aBoolean is NegativeWholeNumberQ is " & NegativeWholeNumberQ(aBoolean)
    Debug.Print "aString is NegativeWholeNumberQ is " & NegativeWholeNumberQ(aString)
    Debug.Print "CVErr(1) is NegativeWholeNumberQ is " & NegativeWholeNumberQ(CVErr(1))
    Debug.Print "aWorksheet is NegativeWholeNumberQ is " & NegativeWholeNumberQ(aWorksheet)
    Debug.Print "aWorkbook is NegativeWholeNumberQ is " & NegativeWholeNumberQ(aWorkbook)
    Debug.Print "aListObject is NegativeWholeNumberQ is " & NegativeWholeNumberQ(aListObject)
    Debug.Print "aVariant is NegativeWholeNumberQ is " & NegativeWholeNumberQ(aVariant)
    Debug.Print "anArray is NegativeWholeNumberQ is " & NegativeWholeNumberQ(AnArray)
    Debug.Print "aNegativeInteger is NegativeWholeNumberQ is " & NegativeWholeNumberQ(aNegativeInteger)

    Call aListObject.Delete
End Sub

Public Sub TestPredicatesNegativeWholeNumberArrayQ()
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
    
    For Each aVariant In Array(Array(-1, -2), _
                               Array(anInteger, aDouble), _
                               Array(aDate, aString), _
                               Array(EmptyArray(), 1), _
                               EmptyArray(), _
                               Array(Null, Empty), _
                               Array(Nothing, 1), _
                               Array(-1, 2, 1), _
                               Array(0, -1))
        Debug.Print "Test is " & NegativeWholeNumberArrayQ(aVariant)
    Next

    Call aListObject.Delete
End Sub

Public Sub TestPredicatesNonNonNegativeWholeNumberQ()
    Dim anInteger As Integer
    Dim aNegativeInteger As Integer
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
    
    Let aVariant = 1
    Let aNegativeInteger = -1
    
    Debug.Print "anInteger is NonNegativeWholeNumberQ is " & NonNegativeWholeNumberQ(anInteger)
    Debug.Print "aDouble is NonNegativeWholeNumberQ is " & NonNegativeWholeNumberQ(aDouble)
    Debug.Print "aDate is NonNegativeWholeNumberQ is " & NonNegativeWholeNumberQ(aDate)
    Debug.Print "aBoolean is NonNegativeWholeNumberQ is " & NonNegativeWholeNumberQ(aBoolean)
    Debug.Print "aString is NonNegativeWholeNumberQ is " & NonNegativeWholeNumberQ(aString)
    Debug.Print "CVErr(1) is NonNegativeWholeNumberQ is " & NonNegativeWholeNumberQ(CVErr(1))
    Debug.Print "aWorksheet is NonNegativeWholeNumberQ is " & NonNegativeWholeNumberQ(aWorksheet)
    Debug.Print "aWorkbook is NonNegativeWholeNumberQ is " & NonNegativeWholeNumberQ(aWorkbook)
    Debug.Print "aListObject is NonNegativeWholeNumberQ is " & NonNegativeWholeNumberQ(aListObject)
    Debug.Print "aVariant is NonNegativeWholeNumberQ is " & NonNegativeWholeNumberQ(aVariant)
    Debug.Print "anArray is NonNegativeWholeNumberQ is " & NonNegativeWholeNumberQ(AnArray)
    Debug.Print "aNegativeInteger is NonNegativeWholeNumberQ is " & NonNegativeWholeNumberQ(aNegativeInteger)

    Call aListObject.Delete
End Sub

Public Sub TestPredicatesNonNegativeWholeNumberArrayQ()
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
    
    For Each aVariant In Array(Array(-1, -2), _
                               Array(anInteger, aDouble), _
                               Array(aDate, aString), _
                               Array(EmptyArray(), 1), _
                               EmptyArray(), _
                               Array(Null, Empty), _
                               Array(Nothing, 1), _
                               Array(-1, 2, 1), _
                               Array(0, -1), _
                               Array(1, 2, 3))
        Debug.Print "Test is " & NonNegativeWholeNumberArrayQ(aVariant)
    Next

    Call aListObject.Delete
End Sub

Public Sub TestPredicatesWholeNumberOrStringQ()
    Dim anInteger As Integer
    Dim aNegativeInteger As Integer
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
    
    Let aVariant = 1
    Let aNegativeInteger = -1
    
    Debug.Print "anInteger is WholeNumberOrStringQ is " & WholeNumberOrStringQ(anInteger)
    Debug.Print "aDouble is WholeNumberOrStringQ is " & WholeNumberOrStringQ(aDouble)
    Debug.Print "aDate is WholeNumberOrStringQ is " & WholeNumberOrStringQ(aDate)
    Debug.Print "aBoolean is WholeNumberOrStringQ is " & WholeNumberOrStringQ(aBoolean)
    Debug.Print "aString is WholeNumberOrStringQ is " & WholeNumberOrStringQ(aString)
    Debug.Print "CVErr(1) is WholeNumberOrStringQ is " & WholeNumberOrStringQ(CVErr(1))
    Debug.Print "aWorksheet is WholeNumberOrStringQ is " & WholeNumberOrStringQ(aWorksheet)
    Debug.Print "aWorkbook is WholeNumberOrStringQ is " & WholeNumberOrStringQ(aWorkbook)
    Debug.Print "aListObject is WholeNumberOrStringQ is " & WholeNumberOrStringQ(aListObject)
    Debug.Print "aVariant is WholeNumberOrStringQ is " & WholeNumberOrStringQ(aVariant)
    Debug.Print "anArray is WholeNumberOrStringQ is " & WholeNumberOrStringQ(AnArray)
    Debug.Print "aNegativeInteger is WholeNumberOrStringQ is " & WholeNumberOrStringQ(aNegativeInteger)

    Call aListObject.Delete
End Sub

Public Sub TestPredicatesWholeNumberOrStringArrayQ()
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
    
    For Each aVariant In Array(Array(-1, -2), _
                               Array(anInteger, aDouble), _
                               Array(aDate, aString), _
                               Array(EmptyArray(), 1), _
                               EmptyArray(), _
                               Array(Null, Empty), _
                               Array(Nothing, 1), _
                               Array(-1, 2, 1), _
                               Array(0, -1), _
                               Array(1, 2, 3))
        Debug.Print "Test is " & WholeNumberOrStringArrayQ(aVariant)
    Next

    Call aListObject.Delete
End Sub

Public Sub TestPredicatesNumberOrStringQ()
    Dim anInteger As Integer
    Dim aNegativeInteger As Integer
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
    
    Let aVariant = 1
    Let aNegativeInteger = -1
    
    Debug.Print "anInteger is NumberOrStringQ is " & NumberOrStringQ(anInteger)
    Debug.Print "aDouble is NumberOrStringQ is " & NumberOrStringQ(aDouble)
    Debug.Print "aDate is NumberOrStringQ is " & NumberOrStringQ(aDate)
    Debug.Print "aBoolean is NumberOrStringQ is " & NumberOrStringQ(aBoolean)
    Debug.Print "aString is NumberOrStringQ is " & NumberOrStringQ(aString)
    Debug.Print "CVErr(1) is NumberOrStringQ is " & NumberOrStringQ(CVErr(1))
    Debug.Print "aWorksheet is NumberOrStringQ is " & NumberOrStringQ(aWorksheet)
    Debug.Print "aWorkbook is NumberOrStringQ is " & NumberOrStringQ(aWorkbook)
    Debug.Print "aListObject is NumberOrStringQ is " & NumberOrStringQ(aListObject)
    Debug.Print "aVariant is NumberOrStringQ is " & NumberOrStringQ(aVariant)
    Debug.Print "anArray is NumberOrStringQ is " & NumberOrStringQ(AnArray)
    Debug.Print "aNegativeInteger is NumberOrStringQ is " & NumberOrStringQ(aNegativeInteger)

    Call aListObject.Delete
End Sub

Public Sub TestPredicatesNumberOrStringArrayQ()
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
    
    For Each aVariant In Array(Array(-1, -2), _
                               Array(anInteger, aDouble), _
                               Array(aDate, aString), _
                               Array(EmptyArray(), 1), _
                               EmptyArray(), _
                               Array(Null, Empty), _
                               Array(Nothing, 1), _
                               Array(-1, 2, 1), _
                               Array(0, -1), _
                               Array(1, 2, 3))
        Debug.Print "Test is " & NumberOrStringArrayQ(aVariant)
    Next

    Call aListObject.Delete
End Sub

Public Sub TestPredicatesStringArrayQ()
    Debug.Print "StringArrayQ(Array(1,""a"")) is " & StringArrayQ(Array(1, "a"))
    Debug.Print "StringArrayQ(Array(""a"",""b"")) is " & StringArrayQ(Array("a", "b"))
End Sub

Public Sub TestPredicatesDictionaryQ()
    Dim aDict As New Dictionary
    
    Debug.Print "DictionaryQ(aDict) is " & DictionaryQ(aDict)
    Debug.Print "DictionaryQ(1) is " & DictionaryQ(1)
End Sub

Public Sub TestPredicatesDictionaryArrayQ()
    Dim aDict As New Dictionary
    
    Debug.Print "DictionaryArrayQ(Array(aDict, aDict)) is " & DictionaryArrayQ(Array(aDict, aDict))
    Debug.Print "DictionaryArrayQ(Array(aDict, 1)) is " & DictionaryArrayQ(Array(aDict, 1))
End Sub

Public Sub TestPredicatesMatrixQ()
    Dim AnArray(1 To 2, 1 To 2) As Variant

    Debug.Print MatrixQ(Array(1, 2, 3))
    Debug.Print MatrixQ(1)
    Debug.Print MatrixQ(Empty)
    Debug.Print MatrixQ([{1,2; 3,4}])
    Debug.Print MatrixQ([{1,2; "A",4}])
    
    Let AnArray(1, 1) = 1
    Let AnArray(1, 2) = 2
    Let AnArray(2, 1) = Empty
    Let AnArray(2, 2) = Null
    Debug.Print MatrixQ(AnArray)
End Sub

Public Sub TestPredicatesAtomicTableQ()
    Dim AnArray(1 To 2, 1 To 2) As Variant
    Dim wsht As Worksheet
    Dim wbk As Workbook
    
    Set wsht = TempComputation
    Set wbk = ThisWorkbook
    
    Debug.Print "AtomicTableQ(Array(1, 2, 3)) is " & AtomicTableQ(Array(1, 2, 3))
    Debug.Print "AtomicTableQ(1) is " & AtomicTableQ(1)
    Debug.Print "AtomicTableQ(Empty) is " & AtomicTableQ(Empty)
    Debug.Print "AtomicTableQ([{1,2; 3,4}]) is " & AtomicTableQ([{1,2; 3,4}]); ""
    Debug.Print "AtomicTableQ([{1,2; ""A"",4}]) is " & AtomicTableQ([{1,2; "A",4}])
    
    Let AnArray(1, 1) = 1
    Let AnArray(1, 2) = 2
    Let AnArray(2, 1) = Empty
    Set AnArray(2, 2) = wsht
    Debug.Print "AtomicTableQ([{1,2; wsht,4}]) is " & AtomicTableQ(AnArray)
    
    Let AnArray(1, 1) = 1
    Let AnArray(1, 2) = 2
    Let AnArray(2, 1) = Empty
    Set AnArray(2, 2) = wbk
    Debug.Print "AtomicTableQ([{1,2; wbk,4}]) is " & AtomicTableQ(AnArray)

    Let AnArray(1, 1) = 1
    Let AnArray(1, 2) = 2
    Let AnArray(2, 1) = Empty
    Let AnArray(2, 2) = Null
    Debug.Print "AtomicTableQ([{1,2; Empty, Null}] is " & AtomicTableQ(AnArray)
    
    Let AnArray(1, 1) = 1
    Let AnArray(1, 2) = 2
    Let AnArray(2, 1) = #1/2/2001#
    Let AnArray(2, 2) = True
    Debug.Print "PrintableTableQ([{1,2; #1/2/2001#, True}] is " & AtomicTableQ(AnArray)
End Sub

Public Sub TestPredicatesPrintableTableQ()
    Dim AnArray(1 To 2, 1 To 2) As Variant
    Dim wsht As Worksheet
    Dim wbk As Workbook
    
    Set wsht = TempComputation
    Set wbk = ThisWorkbook
    
    Debug.Print "PrintableTableQ(Array(1, 2, 3)) is " & PrintableTableQ(Array(1, 2, 3))
    Debug.Print "PrintableTableQ(1) is " & PrintableTableQ(1)
    Debug.Print "PrintableTableQ(Empty) is " & PrintableTableQ(Empty)
    Debug.Print "PrintableTableQ([{1,2; 3,4}]) is " & PrintableTableQ([{1,2; 3,4}]); ""
    Debug.Print "PrintableTableQ([{1,2; ""A"",4}]) is " & PrintableTableQ([{1,2; "A",4}])

    Let AnArray(1, 1) = 1
    Let AnArray(1, 2) = 2
    Let AnArray(2, 1) = Empty
    Set AnArray(2, 2) = wsht
    Debug.Print "AtomicPrintableTableQ([{1,2; wsht,4}]) is " & PrintableTableQ(AnArray)
    
    Let AnArray(1, 1) = 1
    Let AnArray(1, 2) = 2
    Let AnArray(2, 1) = Empty
    Set AnArray(2, 2) = wbk
    Debug.Print "AtomicPrintableTableQ([{1,2; Empty, wbk}]) is " & PrintableTableQ(AnArray)

    Let AnArray(1, 1) = 1
    Let AnArray(1, 2) = 2
    Let AnArray(2, 1) = Empty
    Let AnArray(2, 2) = Null
    Debug.Print "PrintableTableQ([{1,2; Empty, Null}]) is " & PrintableTableQ(AnArray)
    
    Let AnArray(1, 1) = 1
    Let AnArray(1, 2) = 2
    Let AnArray(2, 1) = #1/2/2001#
    Let AnArray(2, 2) = True
    Debug.Print "PrintableTableQ([{1,2; #1/2/2001#, True}]) is " & PrintableTableQ(AnArray)
    
    Let AnArray(1, 1) = 1
    Let AnArray(1, 2) = 2
    Let AnArray(2, 1) = #1/2/2001#
    Set AnArray(2, 2) = ThisWorkbook
    Debug.Print "PrintableTableQ([{1,2; #1/2/2001#, ThisWorkbook}]) is " & PrintableTableQ(AnArray)
End Sub

Public Sub TestPredicatesColumnVectorQ()
    Dim v As Variant
    
    Let v = Array(1, 2, 3)
    Debug.Print "Testing: Array(1, 2, 3)"
    Debug.Print ColumnVectorQ(v)
    Debug.Print
    
    Let v = [{1; 2; 3}]
    Debug.Print "Testing: [{1; 2; 3}]"
    Debug.Print ColumnVectorQ(v)
    Debug.Print
    
    Let v = Array(Array(1, 2, 3))
    Debug.Print "Testing: Array(Array(1, 2, 3))"
    Debug.Print ColumnVectorQ(v)
    Debug.Print
    
    Let v = EmptyArray()
    Debug.Print "Testing: EmptyArray()"
    Debug.Print ColumnVectorQ(v)
    Debug.Print
    
    Let v = 1
    Debug.Print "Testing: 1"
    Debug.Print ColumnVectorQ(v)
    Debug.Print
End Sub

Public Sub TestPredicatesRowVectorQ()
    Dim v As Variant
    
    Let v = Array(1, 2, 3)
    Debug.Print "Testing: Array(1, 2, 3)"
    Debug.Print RowVectorQ(v)
    Debug.Print
    
    Let v = [{1; 2; 3}]
    Debug.Print "Testing: [{1; 2; 3}]"
    Debug.Print RowVectorQ(v)
    Debug.Print
    
    Let v = Array(Array(1, 2, 3))
    Debug.Print "Testing: Array(Array(1, 2, 3))"
    Debug.Print RowVectorQ(v)
    Debug.Print
    
    Let v = EmptyArray()
    Debug.Print "Testing: EmptyArray()"
    Debug.Print RowVectorQ(v)
    Debug.Print
    
    Let v = 1
    Debug.Print "Testing: 1"
    Debug.Print RowVectorQ(v)
    Debug.Print
End Sub

Public Sub TestInterpretableAsColumnArrayQ()
    Dim v As Variant
    Dim A() As Variant
    Dim B() As Variant
    
    Let v = [{1, 2, 3; 4,5,6; 7,8,9}]
    Debug.Print "Testing: [{1, 2, 3; 4,5,6; 7,8,9}]"
    Debug.Print InterpretableAsColumnArrayQ(v)
    Debug.Print

    ReDim A(1, 3)
    Let A(1, 1) = 1
    Let A(1, 2) = 2
    Let A(1, 3) = 3
    Debug.Print "Testing: [{1, 2, 3}]"
    Debug.Print InterpretableAsColumnArrayQ(A)
    Debug.Print

    Let v = [{1; 2; 3}]
    Debug.Print "Testing: [{1; 2; 3}]"
    Debug.Print InterpretableAsColumnArrayQ(v)
    Debug.Print

    Let v = Array([{1; 2; 3}])
    Debug.Print "Testing: Array([{1; 2; 3}])"
    Debug.Print InterpretableAsColumnArrayQ(v)
    Debug.Print
    
    Let v = [{[{1;2;3}]; 4; 5}]
    Debug.Print "Testing: [{[{1;2;3}]; 4; 5}]"
    Debug.Print InterpretableAsColumnArrayQ(v)
    Debug.Print

    ReDim A(1, 1)
    Let A(1, 1) = [{1;2;3}]
    Debug.Print "Testing: [{[{1;2;3}]}]"
    Debug.Print InterpretableAsColumnArrayQ(A)
    Debug.Print

    Let v = Array(1, 2, 3)
    Debug.Print "Testing: Array(1, 2, 3)"
    Debug.Print InterpretableAsColumnArrayQ(v)
    Debug.Print
    
    Let v = Array(Array(1, 2, 3))
    Debug.Print "Testing: Array(Array(1, 2, 3))"
    Debug.Print InterpretableAsColumnArrayQ(v)
    Debug.Print
    
    Let v = Array(Array(Array(1, 2, 3)))
    Debug.Print "Testing: Array(Array(Array(1, 2, 3)))"
    Debug.Print InterpretableAsColumnArrayQ(v)
    Debug.Print
    
    Let v = Array([{1, 2, 3}])
    Debug.Print "Testing: Array([{1, 2, 3}])"
    Debug.Print InterpretableAsColumnArrayQ(v)
    Debug.Print
    
    ReDim A(1 To 1, 1 To 1)
    Let A(1, 1) = Array(1, 2, 3)
    Debug.Print "Testing: [{Array(1, 2, 3)}]"
    Debug.Print InterpretableAsColumnArrayQ(A)
    Debug.Print
    
    ReDim A(1 To 1, 1 To 1)
    ReDim B(1 To 1, 1 To 1)
    Let v = [{1;2;3}]
    Let A(1, 1) = v
    Let B(1, 1) = A
    Debug.Print "Testing: [{[{[{1; 2; 3}]}]}]"
    Debug.Print InterpretableAsColumnArrayQ(v)
    Debug.Print
    
    Let v = EmptyArray()
    Debug.Print "Testing: EmptyArray()"
    Debug.Print InterpretableAsColumnArrayQ(v)
    Debug.Print
    
    Let v = Array(Array(Array(1), Array(2), Array(3)))
    Debug.Print "Testing: Array(Array(Array(1), Array(2), Array(3)))"
    Debug.Print InterpretableAsColumnArrayQ(v)
    Debug.Print
End Sub

Public Sub TestPredicatesInterpretableAsRowArrayQ()
    Dim v As Variant
    Dim A() As Variant
    
    Let v = [{1, 2, 3; 4,5,6; 7,8,9}]
    Debug.Print "Testing: [{1, 2, 3; 4,5,6; 7,8,9}]"
    Debug.Print InterpretableAsRowArrayQ(v)
    Debug.Print

    Let v = [{1, 2, 3}]
    Debug.Print "Testing: [{1, 2, 3}]"
    Debug.Print InterpretableAsRowArrayQ(v)
    Debug.Print

    ReDim A(1 To 1, 1 To 1)
    Let A(1, 1) = [{1, 2, 3}]
    Debug.Print "Testing: [{[{1, 2, 3}]}]"
    Debug.Print InterpretableAsRowArrayQ(A)
    Debug.Print

    Let v = [{1; 2; 3}]
    Debug.Print "Testing: [{1; 2; 3}]"
    Debug.Print InterpretableAsRowArrayQ(v)
    Debug.Print

    Let v = Array(1, 2, 3)
    Debug.Print "Testing: Array(1, 2, 3)"
    Debug.Print InterpretableAsRowArrayQ(v)
    Debug.Print
    
    Let v = Array(Array(1, 2, 3))
    Debug.Print "Testing: Array(Array(1, 2, 3))"
    Debug.Print InterpretableAsRowArrayQ(v)
    Debug.Print
    
    Let v = Array(Array(Array(1, 2, 3)))
    Debug.Print "Testing: Array(Array(Array(1, 2, 3)))"
    Debug.Print InterpretableAsRowArrayQ(v)
    Debug.Print
    
    Let v = Array([{1, 2, 3}])
    Debug.Print "Testing: Array([{1, 2, 3}])"
    Debug.Print InterpretableAsRowArrayQ(v)
    Debug.Print
    
    ReDim A(1 To 1, 1 To 1)
    Let A(1, 1) = Array(1, 2, 3)
    Debug.Print "Testing: [{Array(1, 2, 3)}]"
    Debug.Print InterpretableAsRowArrayQ(A)
    Debug.Print
    
    Let v = EmptyArray()
    Debug.Print "Testing: EmptyArray()"
    Debug.Print InterpretableAsRowArrayQ(v)
    Debug.Print
End Sub

Public Sub TestPredicatesMemberQ()
    Debug.Print "MemberQ(1,1) is " & MemberQ(1, 1)
    Debug.Print "MemberQ(Array(1,1),1) is " & MemberQ(Array(1, 1), 1)
    Debug.Print "MemberQ(Array(1,1,Empty),1) is " & MemberQ(Array(1, 1, Empty), 1)
    Debug.Print "MemberQ(Array(1,1,Empty),Empty) is " & MemberQ(Array(1, 1, Empty), Empty)
    Debug.Print "MemberQ(Array(1,1,Empty, Null),Null) is " & MemberQ(Array(1, 1, Null), Null)
    Debug.Print "MemberQ(Array(1,1,Empty),Null) is " & MemberQ(Array(1, 1, Empty), Null)
    Debug.Print "MemberQ(Array(1,1,Null), Empty) is " & MemberQ(Array(1, 1, Null), Empty)
    Debug.Print "MemberQ(Array(1,2,#1/1/2001#), 3) is " & MemberQ(Array(1, 2, #1/1/2001#), 3)
    Debug.Print "MemberQ(Array(1,2,#1/1/2001#), #1/1/2001#) is " & MemberQ(Array(1, 2, #1/1/2001#), #1/1/2001#)
End Sub

Public Sub TestPredicatesFreeQ()
    Debug.Print "FreeQ(1,1) is " & FreeQ(1, 1)
    Debug.Print "FreeQ(Array(1,1),1) is " & FreeQ(Array(1, 1), 1)
    Debug.Print "FreeQ(Array(1,1,Empty),1) is " & FreeQ(Array(1, 1, Empty), 1)
    Debug.Print "FreeQ(Array(1,1,Empty),Empty) is " & FreeQ(Array(1, 1, Empty), Empty)
    Debug.Print "FreeQ(Array(1,1,Empty, Null),Null) is " & FreeQ(Array(1, 1, Null), Null)
    Debug.Print "FreeQ(Array(1,1,Empty),Null) is " & FreeQ(Array(1, 1, Empty), Null)
    Debug.Print "FreeQ(Array(1,1,Null), Empty) is " & FreeQ(Array(1, 1, Null), Empty)
    Debug.Print "FreeQ(Array(1,2,#1/1/2001#), 3) is " & FreeQ(Array(1, 2, #1/1/2001#), 3)
    Debug.Print "FreeQ(Array(1,2,#1/1/2001#), #1/1/2001#) is " & FreeQ(Array(1, 2, #1/1/2001#), #1/1/2001#)
End Sub

Public Sub TestCreateIndexSequenceFromSpan()
    Dim ASpan As Span
    Dim AnArray As Variant
    
    PrintArray CreateIndexSequenceFromSpan(NumericalSequence(1, 10), Span(1, 10))
    Debug.Print
    PrintArray CreateIndexSequenceFromSpan(NumericalSequence(1, 5), Span(1, -1))
    Debug.Print
    PrintArray CreateIndexSequenceFromSpan(NumericalSequence(1, 5), Span(2, -2))
    Debug.Print
    PrintArray CreateIndexSequenceFromSpan(NumericalSequence(1, 5), Span(2, -2, 2))
    Debug.Print
    PrintArray CreateIndexSequenceFromSpan(NumericalSequence(1, 20), Span(1, -1, 2))
End Sub

Public Sub TestArraysPart()
    Dim i As Long
    Dim j As Long
    Dim M As Variant
    Dim A As Variant
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
    Debug.Print "LBound(m), UBound(m) = " & LBound(M), "&", UBound(M)
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
    
    Let A = [{1,2,3;4,5,6;7,8,9;10,11,12;13,14,15}]
    Debug.Print "Testing on A:"
    PrintArray A
    Debug.Print
    For i = -10 To 10: Debug.Print "Row " & i & " is " & PrintArray(Part(A, i), True): Next
    Debug.Print
    Debug.Print "Getting set of rows."
    ReDim M(1 To 3)
    For i = 1 To 10
        For j = 1 To 3
            Let M(j) = Application.WorksheetFunction.RandBetween(1, 5)
        Next
        
        Debug.Print "Rows (" & PrintArray(M, True) & ") is " & vbCr & PrintArray(Part(A, M), True)
        Debug.Print
    Next
    Debug.Print
    
    Debug.Print "A is"
    PrintArray A
    Debug.Print
    
    Debug.Print "Trying Part(A, Array(2, 5))"
    PrintArray Part(A, Array(2, 5))
    Debug.Print
    
    Debug.Print "Trying Part(A, Array(2, -2))"
    PrintArray Part(A, Array(2, -2))
    Debug.Print
        
    Debug.Print "Part(A, 2, Array(1, 2))"
    PrintArray Part(A, 2, Array(1, 2))
    Debug.Print

    Debug.Print "Part(A, 2, Array(2, 3))"
    PrintArray Part(A, 2, Array(2, 3))
    Debug.Print
    
    Debug.Print "Part(A, array(2,4), Array(2, 3))"
    PrintArray Part(A, Array(2, 4), Array(2, 3))
    Debug.Print
    
    Debug.Print "Testing spans"
    
    Debug.Print "Part(A, Span(1, 2))"
    PrintArray Part(A, Span(1, 2))
    Debug.Print
    
    Debug.Print "Part(A, Span(1, 3))"
    PrintArray Part(A, Span(1, 3))
    Debug.Print
    
    Debug.Print "Part(A, Span(1, 4))"
    PrintArray Part(A, Span(2, 4))
    Debug.Print
    
    Debug.Print "Part(A, Span(2, -1))"
    PrintArray Part(A, Span(2, -1))
    Debug.Print
    
    Debug.Print "Part(A, Span(2, -2))"
    PrintArray Part(A, Span(2, -2))
    Debug.Print
    
    Debug.Print "Part(A, Span(1, -1, 2))"
    PrintArray Part(A, Span(1, -1, 2))
    Debug.Print

    Debug.Print "Part(A, Span(2, -1, 2))"
    PrintArray Part(A, Span(2, -1, 2))
    Debug.Print

    Debug.Print "Part(A, Span(2, -1))"
    PrintArray Part(A, 2, Span(2, -1))
    Debug.Print
    
    Debug.Print "Part(A, 2, 3)"
    Debug.Print Part(A, 2, 3)
    Debug.Print
    
    Debug.Print "Part(A, Span(1, -1), 2)"
    PrintArray Part(A, Span(1, -1), 2)
    Debug.Print
    
    Debug.Print "Part(A, Span(1, -1), Span(2, -1))"
    PrintArray Part(A, Span(1, -1), Span(2, -1))
    Debug.Print
    
    Debug.Print "Part(A, Span(1, -1, 2), Span(1, 3, 2))"
    PrintArray Part(A, Span(1, -1, 2), Span(1, 3, 2))
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
    Debug.Print "Get Array(1, -1) with elts 3, 5, 1"
    PrintArray Part(M, Array(1, -1), Array(Array(3, 5, 1)))
    Debug.Print
    Debug.Print "Get stepped segment Array(1, -1), Array(3, 5, 1)"
    PrintArray Part(M, Array(1, -1), Array(3, 5, 1))
    Debug.Print
    Debug.Print "Get stepped segment Array(1, -1),Array(3, 5, 2)"
    PrintArray Part(M, Array(1, -1), Array(3, 5, 2))
    Debug.Print
    Debug.Print "Get stepped segment Array(1, -1), Array(3, 5, 6)"
    PrintArray Part(M, Array(1, -1), Array(3, 5, 6))
    Debug.Print
    Debug.Print "Get segment Array(1, -1), Array(-4,-2))"
    PrintArray Part(M, Array(1, -1), Array(-4, -2))
    Debug.Print
    Debug.Print "Get segment Array(1, -1), Array(Array(-4,-2))"
    PrintArray Part(M, Array(1, -1), Array(Array(-4, -2)))
    Debug.Print
    Debug.Print "Getting a rectangular submatrix Array(Array(3,4)), Array(Array(5,6,7))"
    PrintArray Part(M, Array(Array(3, 4)), Array(Array(5, 6, 7)))
    Debug.Print
    
    Debug.Print "SPEED TESTS"
    Debug.Print "Creating a large 10000 by 1000 2D array"
    Let M = ConstantArray(Empty, 10000, 1000)
    For i = 1 To 10000: For j = 1 To 1000: Let M(i, j) = i * 1000 + j: Next: Next
    Debug.Print "Accessing element 500 by 200"
    Debug.Print M(500, 200)
    Debug.Print "Accessing row 7000"
    Let A = Part(M, 7000)
    PrintArray A
    Debug.Print "Accessig column 700"
    Let A = Part(M, Span(1, -1), 700)
    Debug.Print "The arrays dimensions are: ", LBound(A), UBound(A)
    Debug.Print "The array is"
    PrintArray A
End Sub

Public Sub TestArraysConcatenateArrays()
    Dim A As Variant
    Dim B As Variant
    Dim c As Variant
    
    Let A = [{1,2,3; 4,5,6}]
    Let B = [{7;8}]
    Let c = ConcatenateArrays(A, B)
    Debug.Print "a is:"
    PrintArray A
    Debug.Print
    Debug.Print "b is:"
    PrintArray B
    Debug.Print
    If EmptyArrayQ(c) Then
        Debug.Print "a and b have incompatible dimensions."
    Else
        Debug.Print "The concatenation is:"
    
        PrintArray c
    End If
    Debug.Print "--------------------------" & vbCrLf
    
    Let A = [{1,2,3; 4,5,6}]
    Let B = [{7;8;9}]
    Let c = ConcatenateArrays(A, B)
    Debug.Print "a is:"
    PrintArray A
    Debug.Print
    Debug.Print "b is:"
    PrintArray B
    Debug.Print
    If EmptyArrayQ(c) Then
        Debug.Print "a and b have incompatible dimensions."
    Else
        Debug.Print "The concatenation is:"
    
        PrintArray c
    End If
    Debug.Print "--------------------------" & vbCrLf

    Let A = [{1,2,3; 4,5,6}]
    Let B = 23
    Let c = ConcatenateArrays(A, B)
    Debug.Print "a is:"
    PrintArray A
    Debug.Print
    Debug.Print "b is:"
    PrintArray B
    Debug.Print
    If EmptyArrayQ(c) Then
        Debug.Print "a and b have incompatible dimensions."
    Else
        Debug.Print "The concatenation is:"
    
        PrintArray c
    End If
    Debug.Print "--------------------------" & vbCrLf
    
    Let A = [{1,"02",3; 4,"05",6}]
    Let B = [{"07";"08"}]
    Let c = ConcatenateArrays(A, B)
    Debug.Print "a is:"
    PrintArray A
    Debug.Print
    Debug.Print "b is:"
    PrintArray B
    Debug.Print
    If EmptyArrayQ(c) Then
        Debug.Print "a and b have incompatible dimensions."
    Else
        Debug.Print "The concatenation is:"
    
        PrintArray c
    End If
    Call ToTemp(c)
    Debug.Print "--------------------------" & vbCrLf
    
    Let A = Array(1, 2, 3)
    Let B = Array(4, 5, 6)
    Let c = ConcatenateArrays(A, B)
    Debug.Print "a is:"
    PrintArray A
    Debug.Print
    Debug.Print "b is:"
    PrintArray B
    Debug.Print
    If EmptyArrayQ(c) Then
        Debug.Print "a and b have incompatible dimensions."
    Else
        Debug.Print "The concatenation is:"
    
        PrintArray c
    End If
    Call ToTemp(c)
    Debug.Print "--------------------------" & vbCrLf
    
    Let A = EmptyArray()
    Let B = Array(4, 5, 6)
    Let c = ConcatenateArrays(A, B)
    Debug.Print "a is:"
    PrintArray A
    Debug.Print
    Debug.Print "b is:"
    PrintArray B
    Debug.Print
    If EmptyArrayQ(c) Then
        Debug.Print "a and b have incompatible dimensions."
    Else
        Debug.Print "The concatenation is:"
    
        PrintArray c
    End If
    Call ToTemp(c)
    Debug.Print "--------------------------" & vbCrLf
End Sub

Public Sub TestArraysTake()
    Dim A() As Integer
    Dim r As Long
    Dim c As Long
    
    ReDim A(1 To 7)
    For r = 1 To 7
        Let A(r) = r
    Next
    
    Debug.Print
    Debug.Print "Set a equal to "
    PrintArray A
    Debug.Print "LBound(a,1), UBound(a,1) = " & LBound(A, 1), UBound(A, 1)
    
    Debug.Print
    For r = -10 To 10
        Debug.Print "Testing Take(a, " & r; ")"
        PrintArray Take(A, r)
    Next
    
    ReDim A(0 To 6)
    For r = 0 To 6
        Let A(r) = r
    Next
    
    Debug.Print
    Debug.Print "Set a equal to "
    PrintArray A
    Debug.Print "LBound(a,1), UBound(a,1) = " & LBound(A, 1), UBound(A, 1)
    
    Debug.Print
    For r = -10 To 10
        Debug.Print "Testing Take(a, " & r & ")"
        PrintArray Take(A, r)
    Next

    ReDim A(1 To 9, 1 To 3)
    For r = 1 To 9
        For c = 1 To 3
            Let A(r, c) = r
        Next
    Next
    Debug.Print "Set a to:"
    PrintArray A
    Debug.Print
    Debug.Print "Bounds LBound(a,1), UBound(a,1), LBound(a,2), UBound(a,2): "
    Debug.Print LBound(A, 1), UBound(A, 1), LBound(A, 2), UBound(A, 2)
    Debug.Print
    
    For r = -10 To 10
        Debug.Print "Testing Take(a," & r & ")"
        PrintArray Take(A, r)
        Debug.Print
    Next
    
    Debug.Print "Testing Take(EmptyArray(),1)"
    PrintArray Take(EmptyArray(), 1)
    Debug.Print
     
    Debug.Print "Testing Take(a,Array(-2,-4,-5))"
    PrintArray Take(A, Array(-2, -4, -5))
    Debug.Print
    
    Debug.Print "Testing Take(a,EmptyArray())"
    PrintArray Take(A, EmptyArray())
    Debug.Print
    
    ReDim A(0 To 8, 0 To 3)
    For r = 0 To 8
        For c = 0 To 3
            Let A(r, c) = r
        Next
    Next
    
    Debug.Print "Set a to:"
    PrintArray A
    Debug.Print
    Debug.Print "Bounds LBound(a,1), UBound(a,1), LBound(a,2), UBound(a,2): ", LBound(A, 1), UBound(A, 1), LBound(A, 2), UBound(A, 2)
    Debug.Print
    
    For r = -10 To 10
        Debug.Print "Testing Take(a," & r & ")"
        PrintArray Take(A, r)
        Debug.Print
    Next
End Sub

Public Sub TestArraysStackArrays()
    Dim A As Variant
    
    Debug.Print "Testing StackArrays(Array(1, 2, 3), Array(4, 5, 6)):"
    PrintArray StackArrays(Array(1, 2, 3), Array(4, 5, 6))
    Debug.Print
    
    Debug.Print "Testing StackArrays([{1,2,3,4; 5,6,7,8}], Array(9,10,11,12)):"
    PrintArray StackArrays([{1,2,3,4; 5,6,7,8}], Array(9, 10, 11, 12))
    Debug.Print
    
    Debug.Print "Testing StackArrays(Array(9,10,11,12), [{1,2,3,4; 5,6,7,8}]):"
    PrintArray StackArrays(Array(9, 10, 11, 12), [{1,2,3,4; 5,6,7,8}])
    Debug.Print
    
    Debug.Print "Testing StackArrays([{1,2,3,4; 5,6,7,8}], [{10,20,30,40; 50,60,70,80}]):"
    PrintArray StackArrays([{1,2,3,4; 5,6,7,8}], [{10,20,30,40; 50,60,70,80}])
    Debug.Print

    Debug.Print "Testing StackArrays([{1,""02"", 3 ,4; 5, ""06"",7,8}], [{10,""020"",30,40; 50,""060"",70,80}]):"
    Let A = StackArrays([{1,"02", 3 ,4; 5,"06",7,8}], [{10,"020",30,40; 50,"060",70,80}])
    PrintArray A
    Debug.Print "The dimensions of the stacked array are #rows = " & GetNumberOfRows(A) & ", #cols = " & GetNumberOfColumns(A)
    Debug.Print "Look at worksheet TempComputation to see if formats were preserved."
    Call ToTemp(A, True)
End Sub

Public Sub TestArraysAppend()
    Dim A As Variant
    Dim B As Variant
    
    Debug.Print "Testing Append(Array(1,2,3), 4)"
    PrintArray Append(Array(1, 2, 3), 4)
    Debug.Print
    
    Let A = Append(Array(1, 2, 3), Array(1, 4))
    Debug.Print "Testing Append(Array(1,2,3), array(1,4))"
    Debug.Print "Use watch on variable a"
    Debug.Print
    
    Let A = [{1,2,3; 4,5,6}]
    Let B = Array(7, 8, 9)
    Debug.Print "Testing Append(a, b) on a = [{1,2,3; 4,5,6}] and b = Array(7, 8, 9)"
    PrintArray Append(A, B)
    Debug.Print

    Let A = [{7,8,9; 10,11,12}]
    Let B = [{1,2,3; 4,5,6}]
    Debug.Print "Testing Append(a, b) on a = [{7,8,9; 10,11,12}] and b = [{1,2,3; 4,5,6}]"
    PrintArray Append(A, B)
    Debug.Print
    
    Let A = Array(1, 2, 3)
    Let B = Null
    Debug.Print "Let a = Array(1, 2, 3)"
    Debug.Print "Let b = Null"
    PrintArray Append(A, B)
    Debug.Print "The result has length " & GetArrayLength(Append(A, B))
End Sub

Public Sub TestArraysPrepend()
    Dim A As Variant
    Dim B As Variant
    
    Debug.Print "Testing Prepend(Array(1,2,3), 4)"
    PrintArray Prepend(Array(1, 2, 3), 4)
    Debug.Print
    
    Let A = Prepend(Array(1, 2, 3), Array(1, 4))
    Debug.Print "Testing Prepend(Array(1,2,3), array(1,4))"
    Debug.Print "Use watch on variable a"
    Debug.Print
    
    Let A = [{1,2,3; 4,5,6}]
    Let B = Array(7, 8, 9)
    Debug.Print "Testing Prepend(a, b) on a = [{1,2,3; 4,5,6}] and b = Array(7, 8, 9)"
    PrintArray Prepend(A, B)
    Debug.Print

    Let A = Array(7, 8, 9)
    Let B = [{1,2,3; 4,5,6}]
    Debug.Print "Testing Prepend(a, b) on a = Array(7, 8, 9) and b = [{1,2,3; 4,5,6}]"
    Debug.Print "Use a watch to see the output. Cannot be printed."
    Call Prepend(A, B)
    Debug.Print
    
    Let A = [{7,8,9; 10,11,12}]
    Let B = [{1,2,3; 4,5,6}]
    Debug.Print "Testing Prepend(a, b) on a = [{7,8,9; 10,11,12}] and b = [{1,2,3; 4,5,6}]"
    PrintArray Prepend(A, B)
End Sub

Public Sub TestArraysInsert()
    Dim A As Variant
    Dim i As Long
    
    Debug.Print "Testing insertion into 1D array:"
    Let A = NumericalSequence(1, 15)
    Debug.Print "Insert in a = NumericalSequence(1, 15) i for 1 to 16"
    For i = 1 To 16
        PrintArray Insert(A, "*", i)
    Next
    
    Debug.Print
    Debug.Print "Insert using negative indices."
    Debug.Print "Insert in a = NumericalSequence(1, 15) i for 16 to 1 step -1"
    For i = 16 To 1 Step -1
        PrintArray Insert(A, "*", i)
    Next
    Debug.Print
    
    ReDim A(0 To 14)
    For i = 0 To 14: Let A(i) = i: Next
    Debug.Print "A now has indices LBound(A) = " & LBound(A) & " and UBound(A) = " & UBound(A)
    Debug.Print "Insert in a = NumericalSequence(1, 15) i for 0 to 15"
    For i = 1 To 16
        PrintArray Insert(A, "*", i)
    Next
    
    Debug.Print
    Debug.Print "Insert using negative indices."
    Debug.Print "Insert in a = NumericalSequence(1, 15) i for 16 to 1 step -1"
    For i = 16 To 1 Step -1
        PrintArray Insert(A, "*", i)
    Next
    
    Debug.Print
    Let A = ConstantArray(Empty, 3, 3)
    For i = 1 To 3: Let A(i, i) = i: Next
    Debug.Print "Testing insertion into 2D array:"
    PrintArray A
    Debug.Print
    For i = 1 To 4: Debug.Print "Insert into row " & i: PrintArray Insert(A, Array("*", "*", "*"), i): Next
    
    Debug.Print
    Let A = NumericalSequence(1, 15)
    Debug.Print "Inserting into multiple places at once."
    Debug.Print "Inserting * at Array(Array(1),Array(3)) in"
    Debug.Print "A = (" & PrintArray(A, True) & ")"
    PrintArray Insert(A, "*", Array(Array(1), Array(3)))
    Debug.Print
    Debug.Print "Inserting * at Array(Array(2),Array(4)) in"
    Debug.Print "A = (" & PrintArray(A, True) & ")"
    PrintArray Insert(A, "*", Array(Array(2), Array(4)))
    Debug.Print
    Debug.Print "Inserting * at Array(Array(5),Array(7),Array(9)) in"
    Debug.Print "A = (" & PrintArray(A, True) & ")"
    PrintArray Insert(A, "*", Array(Array(5), Array(7), Array(9)))
    Debug.Print
    Debug.Print "Inserting * at Array(Array(1),Array(16)) in"
    Debug.Print "A = (" & PrintArray(A, True) & ")"
    PrintArray Insert(A, "*", Array(Array(1), Array(16)))
    Debug.Print
    Debug.Print "Inserting * at Array(Array(1),Array(15)) in"
    Debug.Print "A = (" & PrintArray(A, True) & ")"
    PrintArray Insert(A, "*", Array(Array(1), Array(15)))
    Debug.Print
    Debug.Print "Inserting * at Array(Array(1),Array(2)) in"
    Debug.Print "A = (" & PrintArray(Array(), True) & ")"
    PrintArray Insert(Array(), "*", Array(Array(1), Array(2)))
    Debug.Print
    Debug.Print "Inserting * at Array(Array(1)) in"
    Debug.Print "A = (" & PrintArray(Array(), True) & ")"
    PrintArray Insert(Array(), "*", Array(Array(1)))
    Debug.Print
    Debug.Print "Inserting * at Array(Array(1), Array(-1)) in"
    Debug.Print "A = (" & PrintArray(Array(), True) & ")"
    PrintArray Insert(Array(), "*", Array(Array(1), Array(-1)))
    Debug.Print
    Debug.Print "Inserting * at Array(Array(1), Array(-2)) in"
    Debug.Print "A = (" & PrintArray(Array(1), True) & ")"
    PrintArray Insert(Array(), "*", Array(Array(1), Array(-2)))
    
    Debug.Print
    Let A = ConstantArray(Empty, 3, 3)
    For i = 1 To 3: Let A(i, i) = i: Next
    Debug.Print "Testing insertion in multiple indices into 2D array:"
    PrintArray A
    Debug.Print
    
    Debug.Print "Inserting Array(""*"", ""*"", ""*"") at Array(Array(2), Array(3)):"
    PrintArray Insert(A, Array("*", "*", "*"), Array(Array(2), Array(3)))
    Debug.Print
    Debug.Print "Inserting Array(""*"", ""*"", ""*"") at Array(Array(1), Array(4)):"
    PrintArray Insert(A, Array("*", "*", "*"), Array(Array(1), Array(4)))
    Debug.Print
    Debug.Print "Inserting Array(""*"", ""*"", ""*"") at Array(Array(1), Array(-1)):"
    PrintArray Insert(A, Array("*", "*", "*"), Array(Array(1), Array(-1)))
    Debug.Print
    Debug.Print "Inserting Array(""*"", ""*"", ""*"") at Array(Array(3), Array(2)):"
    PrintArray Insert(A, Array("*", "*", "*"), Array(Array(3), Array(2)))
    Debug.Print
    Debug.Print "Inserting Array(""*"", ""*"", ""*"") at Array(Array(3), Array(-1)):"
    PrintArray Insert(A, Array("*", "*", "*"), Array(Array(3), Array(-1)))
End Sub

Public Sub TestArraysJoinArrays()
    Dim m1 As Variant
    Dim m2 As Variant
    Dim c1 As Variant
    Dim c2 As Variant
    Dim r As Variant

    Let m1 = [{1,2,3; 4,5,6}]
    Let c1 = [{"one";"two";"three"}]
    Debug.Print "Setting m to"
    PrintArray m1
    Debug.Print
    Debug.Print "Setting c1 to"
    PrintArray c1
    Debug.Print "Watch r = JoinArrays(m1, c1)"
    Let r = JoinArrays(m1, c1)
    Debug.Print "Watch r = JoinArrays(m1, c1, 2)"
    Let r = JoinArrays(m1, c1, 2)
    Debug.Print
    Debug.Print
    
    Let m1 = Array(1, 2, 3, 4, 5, 6)
    Let c1 = Array("one", "two", "three")
    PrintArray JoinArrays(m1, c1)
    Debug.Print
    Debug.Print
    
    Let m1 = Array(1, 2, 3, 4, 5, 6)
    Let c1 = Array("one", "two", "three")
    Let c2 = Array("four", "five", "six")
    PrintArray JoinArrays(m1, c1, c2)
    Debug.Print
    Debug.Print
    
    Let m1 = [{1, 2, 3; 4, 5, 6}]
    ReDim c1(1 To 1, 1 To 3)
    Let c1(1, 1) = "One"
    Let c1(1, 2) = "Two"
    Let c1(1, 3) = "Three"
    ReDim c2(1 To 1, 1 To 3)
    Let c2(1, 1) = "four"
    Let c2(1, 2) = "five"
    Let c2(1, 3) = "six"
    PrintArray JoinArrays(m1, c1, c2, 2)
    Debug.Print
    Debug.Print
    
    Let m1 = [{1, 2, 3; 4, 5, 6}]
    ReDim c1(1 To 1, 1 To 3)
    Let c1(1, 1) = "One"
    Let c1(1, 2) = "Two"
    Let c1(1, 3) = "Three"
    ReDim c2(1 To 1, 1 To 3)
    Let c2(1, 1) = "four"
    Let c2(1, 2) = "five"
    Let c2(1, 3) = "six"
    PrintArray JoinArrays(m1, c1, c2)
    Debug.Print
    Debug.Print
    
    Debug.Print "Testing JoinArrays([{1,2;3,4}], [{5,6;7,8}])"
    PrintArray JoinArrays([{1,2;3,4}], [{5,6;7,8}])
    Debug.Print
    Debug.Print
    Debug.Print "Testing JoinArrays([{1,2;3,4}], [{5,6;7,8}], 2)"
    PrintArray JoinArrays([{1,2;3,4}], [{5,6;7,8}], 2)
End Sub

Public Sub TestArraysReverse()
    Dim A As Variant
    Dim M As Variant
    
    Let A = NumericalSequence(1, 10)
    Debug.Print "A is (" & Join(A, ",") & ")"
    Debug.Print "Reverse(A) = (" & Join(Reverse(A), ",") & ")"
    
    Debug.Print
    Debug.Print "Reverse(A,2) is"
    PrintArray Reverse(A, 2)

    Debug.Print
    Let M = [{1,2,3;4,5,6;7,8,9}]
    Debug.Print "M is"
    PrintArray M
    
    Debug.Print
    Debug.Print "Reverse(M) is"
    PrintArray Reverse(M)

    Debug.Print
    Debug.Print "Reverse(M,2) is"
    PrintArray Reverse(M, 2)
    
    Debug.Print
    Let M = [{1;2;3}]
    Debug.Print "M is"
    PrintArray M
    
    Debug.Print
    Debug.Print "Reverse(M) is"
    PrintArray Reverse(M)

    Debug.Print
    Debug.Print "Reverse(M,2) is"
    PrintArray Reverse(M, 2)
End Sub

Public Sub TestArraysDrop()
    Dim A As Variant
    Dim M As Variant
    Dim i As Long
    Dim var As Variant
    
    Let A = NumericalSequence(1, 15)
    Debug.Print "A = " & CreateFunctionalParameterArray(A)
    Debug.Print
    
    For Each var In NumericalSequence(-17, 17, 1, True)
        If IsNull(Drop(A, var)) Then
            Debug.Print "Drop(A, [{" & var & "}]) = Null"
        Else
            Debug.Print "Drop(A, [{" & var & "}]) = " & CreateFunctionalParameterArray(Drop(A, [{var}]))
        End If
    Next
    
    Debug.Print
    For Each var In NumericalSequence(-17, 17, 1, True)
        If IsNull(Drop(A, var)) Then
            Debug.Print "Drop(A, " & var & ") = Null"
        Else
            Debug.Print "Drop(A, " & var & ") = " & CreateFunctionalParameterArray(Drop(A, var))
        End If
    Next
End Sub

Public Sub TestConnectAndSelect()
    PrintArray ConnectAndSelect("SELECT * FROM `documentation`.`wp_posts`;", "documentation", "localhost", "root", "")
End Sub

Public Sub TestGetTableHeaders()
    PrintArray MySql.GetTableHeaders("wp_posts", "documentation", "localhost", "root", "")
End Sub

Public Sub TestFileNameJoin()
    Dim AnArray As Variant
    
    Let AnArray = Array("c:", "dir1", "dir2")
    Debug.Print "The result is " & FileNameJoin(AnArray)
    
    Let AnArray = Array("c:", "dir1", "file.txt")
    Debug.Print "The result is " & FileNameJoin(AnArray)
    
    Let AnArray = Array("c:", "dir1", "file.txt")
    Debug.Print "The result is " & FileNameJoin(AnArray)
    
    Let AnArray = EmptyArray()
    Debug.Print "The result is " & FileNameJoin(AnArray)
End Sub

Public Sub TestFileNameSplit()
    Dim AnArray As Variant
    
    Let AnArray = Array("c:", "dir1", "dir2")
    Debug.Print "The result is"
    PrintArray FileNameSplit(FileNameJoin(AnArray))
    Debug.Print
    
    Let AnArray = Array("c:", "dir1", "file.txt")
    Debug.Print "The result is"
    PrintArray FileNameSplit(FileNameJoin(AnArray))
    Debug.Print
    
    Let AnArray = Empty
    Debug.Print "The result is"
    PrintArray FileNameSplit(Empty)
    Debug.Print "IsNull(FileNameSplit(Empty)) = " & IsNull(FileNameSplit(Empty))
    Debug.Print
End Sub

Public Sub TestFileBaseName()
    Dim var As Variant

    For Each var In Array("c:\dir1\dir2\base.ext1.ext2", _
                          "c:\dir\base.txt", _
                          "c:\dir\base", _
                          "base.txt", _
                          "base")
        Debug.Print "For " & var & " the base name is -" & FileBaseName(CStr(var)) & "-"
    Next
End Sub

Public Sub TestFileExtension()
    Dim var As Variant

    For Each var In Array("c:\dir1\dir2\base.ext1.ext2", _
                          "c:\dir\base.txt", _
                          "c:\dir\base", _
                          "base.txt", _
                          "base")
        Debug.Print "For " & var & " the extension is -" & FileExtension(CStr(var)) & "-"
    Next
End Sub
