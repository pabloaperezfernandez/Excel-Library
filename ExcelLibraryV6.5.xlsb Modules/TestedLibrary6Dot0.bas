Attribute VB_Name = "TestedLibrary6Dot0"
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
    Dim a() As Variant
    Dim b(1 To 2) As Variant
    Dim c As Integer
    Dim wbk As Workbook
    
    Debug.Print "EmptyArray() is dimensioned is " & DimensionedQ(EmptyArray())
    Debug.Print "a() is dimensioned is " & DimensionedQ(a)
    Debug.Print "b(1 To 2) is dimensioned is " & DimensionedQ(b)
    Debug.Print "c is an integer is dimensioned is " & DimensionedQ(c)
    Debug.Print "wbk is dimensioned is " & DimensionedQ(wbk)
End Sub

Public Sub TestPredicatesEmptyArrayQ()
    Dim a() As Variant
    Dim b(1 To 2) As Variant
    Dim c As Integer
    Dim wbk As Workbook
    
    Debug.Print "EmptyArray() is EmptyArrayQ is " & EmptyArrayQ(EmptyArray())
    Debug.Print "a() is EmptyArrayQ is " & EmptyArrayQ(a)
    Debug.Print "b(1 To 2) is EmptyArrayQ is " & EmptyArrayQ(b)
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
    Dim a() As Variant
    Dim b() As Variant
    
    Let v = [{1, 2, 3; 4,5,6; 7,8,9}]
    Debug.Print "Testing: [{1, 2, 3; 4,5,6; 7,8,9}]"
    Debug.Print InterpretableAsColumnArrayQ(v)
    Debug.Print

    ReDim a(1, 3)
    Let a(1, 1) = 1
    Let a(1, 2) = 2
    Let a(1, 3) = 3
    Debug.Print "Testing: [{1, 2, 3}]"
    Debug.Print InterpretableAsColumnArrayQ(a)
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

    ReDim a(1, 1)
    Let a(1, 1) = [{1;2;3}]
    Debug.Print "Testing: [{[{1;2;3}]}]"
    Debug.Print InterpretableAsColumnArrayQ(a)
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
    
    ReDim a(1 To 1, 1 To 1)
    Let a(1, 1) = Array(1, 2, 3)
    Debug.Print "Testing: [{Array(1, 2, 3)}]"
    Debug.Print InterpretableAsColumnArrayQ(a)
    Debug.Print
    
    ReDim a(1 To 1, 1 To 1)
    ReDim b(1 To 1, 1 To 1)
    Let v = [{1;2;3}]
    Let a(1, 1) = v
    Let b(1, 1) = a
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
    Dim a() As Variant
    
    Let v = [{1, 2, 3; 4,5,6; 7,8,9}]
    Debug.Print "Testing: [{1, 2, 3; 4,5,6; 7,8,9}]"
    Debug.Print InterpretableAsRowArrayQ(v)
    Debug.Print

    Let v = [{1, 2, 3}]
    Debug.Print "Testing: [{1, 2, 3}]"
    Debug.Print InterpretableAsRowArrayQ(v)
    Debug.Print

    ReDim a(1 To 1, 1 To 1)
    Let a(1, 1) = [{1, 2, 3}]
    Debug.Print "Testing: [{[{1, 2, 3}]}]"
    Debug.Print InterpretableAsRowArrayQ(a)
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
    
    ReDim a(1 To 1, 1 To 1)
    Let a(1, 1) = Array(1, 2, 3)
    Debug.Print "Testing: [{Array(1, 2, 3)}]"
    Debug.Print InterpretableAsRowArrayQ(a)
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

Public Sub TestArraysPart()
    Dim i As Integer
    Dim j As Integer
    Dim m As Variant
    
    ' 1D Tests
    Debug.Print "Cycling through elements of Array(1, ..., 6)"
    For i = 1 To 6: PrintArray Part(Array(1, 2, 3, 4, 5, 6), i): Next
    Debug.Print
    Debug.Print "Cycling through segments Array(1, i)"
    For i = 1 To 6: PrintArray Part(Array(1, 2, 3, 4, 5, 6), Array(1, i)): Next
    Debug.Print
    Debug.Print "Cycling through segments Array(i, -1)"
    For i = 1 To 6: PrintArray Part(Array(1, 2, 3, 4, 5, 6), Array(i, -1)): Next
    Debug.Print
    Debug.Print "Cycling through segments Array(-6, i)"
    For i = 1 To 6: PrintArray Part(Array(1, 2, 3, 4, 5, 6), Array(-6, i)): Next
    Debug.Print
    Debug.Print "Get elts 3, 5, 1 from Array(1, 2, 3, 4, 5, 6)"
    PrintArray Part(Array(1, 2, 3, 4, 5, 6), Array(Array(3, 5, 1)))
    Debug.Print
    Debug.Print "Get stepped segment Array(3, 5, 1)"
    PrintArray Part(Array(1, 2, 3, 4, 5, 6), Array(3, 5, 1))
    Debug.Print
    Debug.Print "Get stepped segment Array(3, 5, 2)"
    PrintArray Part(Array(1, 2, 3, 4, 5, 6), Array(3, 5, 2))
    Debug.Print
    Debug.Print "Get stepped segment Array(3, 5, 6)"
    PrintArray Part(Array(1, 2, 3, 4, 5, 6), Array(3, 5, 6))
    
    ' 2D Test with one dimensional index set
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
    PrintArray Part(m, Array(Array(3, 5, 1)))
    Debug.Print
    Debug.Print "Get stepped segment Array(3, 5, 1)"
    PrintArray Part(m, Array(3, 5, 1))
    Debug.Print
    Debug.Print "Get stepped segment Array(3, 5, 2)"
    PrintArray Part(m, Array(3, 5, 2))
    Debug.Print
    Debug.Print "Get stepped segment Array(3, 5, 6)"
    PrintArray Part(m, Array(3, 5, 6))
    Debug.Print
    Debug.Print "Get segment Array(-4,-2))"
    PrintArray Part(m, Array(-4, -2))
    Debug.Print
    Debug.Print "Get segment Array(Array(-4,-2))"
    PrintArray Part(m, Array(Array(-4, -2)))
    
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
    For i = 1 To 6: PrintArray Part(m, Array(1, -1), i): Debug.Print: Next
    Debug.Print
    Debug.Print "Cycling through segments Array(1, -1), Array(1, i)"
    For i = 1 To 6: PrintArray Part(m, Array(1, -1), Array(1, i)): Debug.Print: Next
    Debug.Print
    Debug.Print "Cycling through segments Array(1, -1), Array(i, -1)"
    For i = 1 To 6: PrintArray Part(m, Array(1, -1), Array(i, -1)): Debug.Print: Next
    Debug.Print
    Debug.Print "Cycling through segments Array(1, -1), Array(-6, i)"
    For i = 1 To 6: PrintArray Part(m, Array(1, -1), Array(-6, i)): Debug.Print: Next
    Debug.Print
    Debug.Print "Get Array(1, -1) with elts 3, 5, 1"
    PrintArray Part(m, Array(1, -1), Array(Array(3, 5, 1)))
    Debug.Print
    Debug.Print "Get stepped segment Array(1, -1), Array(3, 5, 1)"
    PrintArray Part(m, Array(1, -1), Array(3, 5, 1))
    Debug.Print
    Debug.Print "Get stepped segment Array(1, -1),Array(3, 5, 2)"
    PrintArray Part(m, Array(1, -1), Array(3, 5, 2))
    Debug.Print
    Debug.Print "Get stepped segment Array(1, -1), Array(3, 5, 6)"
    PrintArray Part(m, Array(1, -1), Array(3, 5, 6))
    Debug.Print
    Debug.Print "Get segment Array(1, -1), Array(-4,-2))"
    PrintArray Part(m, Array(1, -1), Array(-4, -2))
    Debug.Print
    Debug.Print "Get segment Array(1, -1), Array(Array(-4,-2))"
    PrintArray Part(m, Array(1, -1), Array(Array(-4, -2)))
    Debug.Print
    Debug.Print "Getting a rectangular submatrix Array(Array(3,4)), Array(Array(5,6,7))"
    PrintArray Part(m, Array(Array(3, 4)), Array(Array(5, 6, 7)))
End Sub

' This tests Arrays.Take
Public Sub TestTake()
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
    Debug.Print "LBound(a,1), UBound(a,1) = " & LBound(a, 1), UBound(a, 1)
    
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
    Debug.Print "LBound(a,1), UBound(a,1) = " & LBound(a, 1), UBound(a, 1)
    
    Debug.Print
    For r = -10 To 10
        Debug.Print "Testing Take(a, " & r & ")"
        PrintArray Take(a, r)
    Next

    ReDim a(1 To 9, 1 To 3)
    For r = 1 To 9
        For c = 1 To 3
            Let a(r, c) = r
        Next
    Next
    Debug.Print "Set a to:"
    PrintArray a
    Debug.Print
    Debug.Print "Bounds LBound(a,1), UBound(a,1), LBound(a,2), UBound(a,2): ", LBound(a, 1), UBound(a, 1), LBound(a, 2), UBound(a, 2)
    Debug.Print
    
    For r = -10 To 10
        Debug.Print "Testing Take(a," & r & ")"
        PrintArray Take(a, r)
        Debug.Print
    Next
    
    Debug.Print "Testing Take(EmptyArray(),1)"
    PrintArray Take(EmptyArray(), 1)
    Debug.Print
     
    Debug.Print "Testing Take(a,Array(-2,-4,-5))"
    PrintArray Take(a, Array(-2, -4, -5))
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
    Debug.Print "Bounds LBound(a,1), UBound(a,1), LBound(a,2), UBound(a,2): ", LBound(a, 1), UBound(a, 1), LBound(a, 2), UBound(a, 2)
    Debug.Print
    
    For r = -10 To 10
        Debug.Print "Testing Take(a," & r & ")"
        PrintArray Take(a, r)
        Debug.Print
    Next
End Sub

Public Sub TestGetRow()
    Dim m As Variant
    Dim a(0 To 2, 0 To 2) As Integer
    Dim b(1 To 3, 1 To 3) As Integer
    Dim i As Long
    Dim j As Long
    
    For i = 0 To 2
        For j = 0 To 2
            Let a(i, j) = i + j * 3
        Next j
    Next i
    
    Debug.Print "The Array is:"
    PrintArray a
    Debug.Print "Rows 1, 2, and 3 are:"
    
    For i = 1 To 3
        PrintArray GetRow(a, i)
        Debug.Print
    Next
    
    For i = 1 To 3
        For j = 1 To 3
            Let b(i, j) = i + j * 3
        Next j
    Next i
    
    Debug.Print "The Array is:"
    PrintArray b
    Debug.Print
    Debug.Print "Rows 1, 2, and 3 are:"
    
    For i = 1 To 3
        PrintArray GetRow(b, i)
        Debug.Print
    Next
    
    Let m = [{1,2,3;4,5,6;7,8,9}]
    Debug.Print "Testing getting row 1 from [{1,2,3;4,5,6;7,8,9}]"
    PrintArray GetRow(m, 1)
    Debug.Print
    
    Let m = [{1,2,3;4,5,6;7,8,9}]
    Debug.Print "Testing getting row 2 from [{1,2,3;4,5,6;7,8,9}]"
    PrintArray GetRow(m, 2)
    Debug.Print

    Let m = [{1,2,3;4,5,6;7,8,9}]
    Debug.Print "Testing getting row 3 from [{1,2,3;4,5,6;7,8,9}]"
    PrintArray GetRow(m, 3)
    Debug.Print
    
    Let m = [{1;2;3}]
    Debug.Print "Testing getting row 3 from [{1;2;3}]"
    PrintArray GetRow(m, 3)
    Debug.Print
    
    Let m = [{1,2,3;4,5,6;7,8,9}]
    Debug.Print "Testing getting row 10 from [{1,2,3;4,5,6;7,8,9}]"
    PrintArray GetRow(m, 10)
    Debug.Print
    
    Let m = [{1,2,3;4,5,6;7,8,9}]
    Debug.Print "Testing getting row 0 from [{1,2,3;4,5,6;7,8,9}]"
    PrintArray GetRow(m, 0)
    Debug.Print

    Let m = [{1,2,3}]
    Debug.Print "Testing getting row 2 from [{1,2,3}]"
    PrintArray GetRow(m, 2)
    Debug.Print
    
    Let m = Array(1, 2, 3)
    Debug.Print "Testing getting row 0 from Array(1, 2, 3)"
    PrintArray GetRow(m, 0)
    Debug.Print
    
    Let m = Array(1, 2, 3)
    Debug.Print "Testing getting row 3 from Array(1, 2, 3)"
    PrintArray GetRow(m, 3)
    Debug.Print
    
    Let m = Array(1, 2, 3)
    Debug.Print "Testing getting row 1 from Array(1, 2, 3)"
    PrintArray GetRow(m, 1)
    Debug.Print
    
    Let m = EmptyArray()
    Debug.Print "Testing  EmptyArray()"
    PrintArray GetRow(m, 1)
    Debug.Print
End Sub

Public Sub TestGetColumn()
    Dim m As Variant
    Dim c As Integer
    
    Let m = [{1,2,3;4,5,6;7,8,9}]
    PrintArray m
    Debug.Print
    For c = -4 To 4
        Debug.Print "Getting column " & c
        PrintArray GetColumn(m, CLng(c))
    Next
    Debug.Print
    
    Let m = [{1,2,3;4,5,6;7,8,9}]
    Debug.Print "Testing getting column 10 from [{1,2,3;4,5,6;7,8,9}]"
    PrintArray GetColumn(m, 10)
    Debug.Print
    
    Let m = [{1,2,3;4,5,6;7,8,9}]
    Debug.Print "Testing getting column 0 from [{1,2,3;4,5,6;7,8,9}]"
    PrintArray GetColumn(m, 0)
    Debug.Print

    Let m = [{1,2,3; 4,5,6}]
    Debug.Print "Testing getting column 2 from [{1,2,3; 4,5,6}]"
    PrintArray GetColumn(m, 2)
    Debug.Print
    
    Let m = Array(1, 2, 3)
    Debug.Print "Testing getting column 2 from Array(1, 2, 3)"
    PrintArray GetColumn(m, 2)
    Debug.Print
    
    Let m = EmptyArray()
    Debug.Print "Testing  EmptyArray()"
    PrintArray GetColumn(m, 1)
    Debug.Print
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
