Attribute VB_Name = "TestedLibrary6Dot0"
Option Explicit
Option Base 1

'********************************************************************************************
' Miscellaneous VBA
'********************************************************************************************
Public Sub TestForEachOnUnDimensionedArray()
    Dim anArray() As Variant
    Dim var As Variant
    
    On Error GoTo ErrorHandler
    
    Debug.Print "Testing for each on an empty array"
    For Each var In Array()
        Debug.Print "Did one iteration."
    Next
    
    ' This should raise anerror
    Debug.Print "Testing for each on an undimensioned array"
    For Each var In anArray
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

Public Sub TestPredicatesEveryQ()
    Dim UndimensionedArray() As Variant

    Debug.Print "EveryQ(Array(1,2,4), ""WholeNumberQ"") is " & EveryQ(Array(1, 2, 4), "WholeNumberQ")
    Debug.Print "EveryQ(Array(1,2,4.0), ""WholeNumberQ"") is " & EveryQ(Array(1, 2, 4#), "WholeNumberQ")
    Debug.Print "EveryQ(Array(1,2,4), ""NumberQ"") is " & EveryQ(Array(1, 2, 4), "NumberQ")
    Debug.Print "EveryQ(Array(1,2,4.0), ""NumberQ"") is " & EveryQ(Array(1, 2, 4#), "NumberQ")
    Debug.Print "EveryQ(Array(1,2,4.0), ""StringQ"") is " & EveryQ(Array(1, 2, 4#), "StringQ")
    Debug.Print "EveryQ(Array(""a"", ""b"",Empty), ""StringQ"") is " & EveryQ(Array("a", "b", Empty), "StringQ")
    Debug.Print "EveryQ(Array(""a"", ""b""), ""StringQ"") is " & EveryQ(Array("a", "b"), "StringQ")
    Debug.Print "EveryQ(Array(), ""StringQ"") is " & EveryQ(Array(), "StringQ")
    Debug.Print "EveryQ(UndimensionedArray, ""StringQ"") is " & EveryQ(UndimensionedArray, "StringQ")
End Sub

Public Sub TestPredicatesSomeQ()
    Dim UndimensionedArray() As Variant

    Debug.Print "SomeQ(Array(1,2,4), ""WholeNumberQ"") is " & SomeQ(Array(1, 2, 4), "WholeNumberQ")
    Debug.Print "SomeQ(Array(1,2,4.0), ""WholeNumberQ"") is " & SomeQ(Array(1, 2, 4#), "WholeNumberQ")
    Debug.Print "SomeQ(Array(1,2,4), ""NumberQ"") is " & SomeQ(Array(1, 2, 4), "NumberQ")
    Debug.Print "SomeQ(Array(1,2,4.0), ""NumberQ"") is " & SomeQ(Array(1, 2, 4#), "NumberQ")
    Debug.Print "SomeQ(Array(1,2,4.0), ""StringQ"") is " & SomeQ(Array(1, 2, 4#), "StringQ")
    Debug.Print "SomeQ(Array(""a"", ""b"",Empty), ""StringQ"") is " & SomeQ(Array("a", "b", Empty), "StringQ")
    Debug.Print "SomeQ(Array(""a"", ""b""), ""StringQ"") is " & SomeQ(Array("a", "b"), "StringQ")
    Debug.Print "SomeQ(Array(), ""StringQ"") is " & SomeQ(Array(), "StringQ")
    Debug.Print "SomeQ(UndimensionedArray, ""StringQ"") is " & SomeQ(UndimensionedArray, "StringQ")
End Sub

Public Sub TestPredicatesDimensionedQ()
    Dim a() As Variant
    Dim b(1 To 2) As Variant
    Dim c As Integer
    Dim wbk As Workbook
    
    Debug.Print "Array() is dimensioned is " & DimensionedQ(Array())
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
    
    Debug.Print "Array() is EmptyArrayQ is " & EmptyArrayQ(Array())
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
    Dim anArray(1 To 2) As Integer
    
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
    Debug.Print "anArray is AtomicQ is " & AtomicQ(anArray)

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
    Dim anArray(1 To 2) As Integer
    
    Set aWorksheet = ActiveSheet
    Set aWorkbook = ThisWorkbook
    Set aListObject = AddListObject(TempComputation.Range("A1"))
    
    For Each aVariant In Array(Array(anInteger, aDouble), _
                               Array(aDate, aString), _
                               Array(Array(), 1), _
                               Array(), _
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
    Dim anArray(1 To 2) As Integer
    
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
    Debug.Print "anArray is NumberQ is " & NumberQ(anArray)

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
    Dim anArray(1 To 2) As Integer
    
    Set aWorksheet = ActiveSheet
    Set aWorkbook = ThisWorkbook
    Set aListObject = AddListObject(TempComputation.Range("A1"))
    
    For Each aVariant In Array(Array(anInteger, aDouble), _
                               Array(aDate, aString), _
                               Array(Array(), 1), _
                               Array(), _
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
    Dim anArray(1 To 2) As Integer
    
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
    Debug.Print "anArray is WholeNumberQ is " & WholeNumberQ(anArray)

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
    Dim anArray(1 To 2) As Integer
    
    Set aWorksheet = ActiveSheet
    Set aWorkbook = ThisWorkbook
    Set aListObject = AddListObject(TempComputation.Range("A1"))
    
    For Each aVariant In Array(Array(1, 2), _
                               Array(anInteger, aDouble), _
                               Array(aDate, aString), _
                               Array(Array(), 1), _
                               Array(), _
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
    Dim anArray(1 To 2) As Integer
    
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
    Debug.Print "anArray is PositiveWholeNumberQ is " & PositiveWholeNumberQ(anArray)
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
    Dim anArray(1 To 2) As Integer
    
    Set aWorksheet = ActiveSheet
    Set aWorkbook = ThisWorkbook
    Set aListObject = AddListObject(TempComputation.Range("A1"))
    
    For Each aVariant In Array(Array(1, 2), _
                               Array(anInteger, aDouble), _
                               Array(aDate, aString), _
                               Array(Array(), 1), _
                               Array(), _
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
    Dim anArray(1 To 2) As Integer
    
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
    Debug.Print "anArray is NegativeWholeNumberQ is " & NegativeWholeNumberQ(anArray)
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
    Dim anArray(1 To 2) As Integer
    
    Set aWorksheet = ActiveSheet
    Set aWorkbook = ThisWorkbook
    Set aListObject = AddListObject(TempComputation.Range("A1"))
    
    For Each aVariant In Array(Array(-1, -2), _
                               Array(anInteger, aDouble), _
                               Array(aDate, aString), _
                               Array(Array(), 1), _
                               Array(), _
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
    Dim anArray(1 To 2) As Integer
    
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
    Debug.Print "anArray is NonNegativeWholeNumberQ is " & NonNegativeWholeNumberQ(anArray)
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
    Dim anArray(1 To 2) As Integer
    
    Set aWorksheet = ActiveSheet
    Set aWorkbook = ThisWorkbook
    Set aListObject = AddListObject(TempComputation.Range("A1"))
    
    For Each aVariant In Array(Array(-1, -2), _
                               Array(anInteger, aDouble), _
                               Array(aDate, aString), _
                               Array(Array(), 1), _
                               Array(), _
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
    Dim anArray(1 To 2) As Integer
    
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
    Debug.Print "anArray is WholeNumberOrStringQ is " & WholeNumberOrStringQ(anArray)
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
    Dim anArray(1 To 2) As Integer
    
    Set aWorksheet = ActiveSheet
    Set aWorkbook = ThisWorkbook
    Set aListObject = AddListObject(TempComputation.Range("A1"))
    
    For Each aVariant In Array(Array(-1, -2), _
                               Array(anInteger, aDouble), _
                               Array(aDate, aString), _
                               Array(Array(), 1), _
                               Array(), _
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
    Dim anArray(1 To 2) As Integer
    
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
    Debug.Print "anArray is NumberOrStringQ is " & NumberOrStringQ(anArray)
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
    Dim anArray(1 To 2) As Integer
    
    Set aWorksheet = ActiveSheet
    Set aWorkbook = ThisWorkbook
    Set aListObject = AddListObject(TempComputation.Range("A1"))
    
    For Each aVariant In Array(Array(-1, -2), _
                               Array(anInteger, aDouble), _
                               Array(aDate, aString), _
                               Array(Array(), 1), _
                               Array(), _
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
    Dim anArray(1 To 2, 1 To 2) As Variant

    Debug.Print MatrixQ(Array(1, 2, 3))
    Debug.Print MatrixQ(1)
    Debug.Print MatrixQ(Empty)
    Debug.Print MatrixQ([{1,2; 3,4}])
    Debug.Print MatrixQ([{1,2; "A",4}])
    
    Let anArray(1, 1) = 1
    Let anArray(1, 2) = 2
    Let anArray(2, 1) = Empty
    Let anArray(2, 2) = Null
    Debug.Print MatrixQ(anArray)
End Sub

Public Sub TestPredicatesAtomicTableQ()
    Dim anArray(1 To 2, 1 To 2) As Variant
    Dim wsht As Worksheet
    Dim wbk As Workbook
    
    Set wsht = TempComputation
    Set wbk = ThisWorkbook
    
    Debug.Print "AtomicTableQ(Array(1, 2, 3)) is " & AtomicTableQ(Array(1, 2, 3))
    Debug.Print "AtomicTableQ(1) is " & AtomicTableQ(1)
    Debug.Print "AtomicTableQ(Empty) is " & AtomicTableQ(Empty)
    Debug.Print "AtomicTableQ([{1,2; 3,4}]) is " & AtomicTableQ([{1,2; 3,4}]); ""
    Debug.Print "AtomicTableQ([{1,2; ""A"",4}]) is " & AtomicTableQ([{1,2; "A",4}])
    
    Let anArray(1, 1) = 1
    Let anArray(1, 2) = 2
    Let anArray(2, 1) = Empty
    Set anArray(2, 2) = wsht
    Debug.Print "AtomicTableQ([{1,2; wsht,4}]) is " & AtomicTableQ(anArray)
    
    Let anArray(1, 1) = 1
    Let anArray(1, 2) = 2
    Let anArray(2, 1) = Empty
    Set anArray(2, 2) = wbk
    Debug.Print "AtomicTableQ([{1,2; wbk,4}]) is " & AtomicTableQ(anArray)

    Let anArray(1, 1) = 1
    Let anArray(1, 2) = 2
    Let anArray(2, 1) = Empty
    Let anArray(2, 2) = Null
    Debug.Print "AtomicTableQ([{1,2; Empty, Null}] is " & AtomicTableQ(anArray)
    
    Let anArray(1, 1) = 1
    Let anArray(1, 2) = 2
    Let anArray(2, 1) = #1/2/2001#
    Let anArray(2, 2) = True
    Debug.Print "PrintableTableQ([{1,2; #1/2/2001#, True}] is " & AtomicTableQ(anArray)
End Sub

Public Sub TestPredicatesPrintableTableQ()
    Dim anArray(1 To 2, 1 To 2) As Variant
    Dim wsht As Worksheet
    Dim wbk As Workbook
    
    Set wsht = TempComputation
    Set wbk = ThisWorkbook
    
    Debug.Print "PrintableTableQ(Array(1, 2, 3)) is " & PrintableTableQ(Array(1, 2, 3))
    Debug.Print "PrintableTableQ(1) is " & PrintableTableQ(1)
    Debug.Print "PrintableTableQ(Empty) is " & PrintableTableQ(Empty)
    Debug.Print "PrintableTableQ([{1,2; 3,4}]) is " & PrintableTableQ([{1,2; 3,4}]); ""
    Debug.Print "PrintableTableQ([{1,2; ""A"",4}]) is " & PrintableTableQ([{1,2; "A",4}])

    Let anArray(1, 1) = 1
    Let anArray(1, 2) = 2
    Let anArray(2, 1) = Empty
    Set anArray(2, 2) = wsht
    Debug.Print "AtomicPrintableTableQ([{1,2; wsht,4}]) is " & PrintableTableQ(anArray)
    
    Let anArray(1, 1) = 1
    Let anArray(1, 2) = 2
    Let anArray(2, 1) = Empty
    Set anArray(2, 2) = wbk
    Debug.Print "AtomicPrintableTableQ([{1,2; Empty, wbk}]) is " & PrintableTableQ(anArray)

    Let anArray(1, 1) = 1
    Let anArray(1, 2) = 2
    Let anArray(2, 1) = Empty
    Let anArray(2, 2) = Null
    Debug.Print "PrintableTableQ([{1,2; Empty, Null}]) is " & PrintableTableQ(anArray)
    
    Let anArray(1, 1) = 1
    Let anArray(1, 2) = 2
    Let anArray(2, 1) = #1/2/2001#
    Let anArray(2, 2) = True
    Debug.Print "PrintableTableQ([{1,2; #1/2/2001#, True}]) is " & PrintableTableQ(anArray)
    
    Let anArray(1, 1) = 1
    Let anArray(1, 2) = 2
    Let anArray(2, 1) = #1/2/2001#
    Set anArray(2, 2) = ThisWorkbook
    Debug.Print "PrintableTableQ([{1,2; #1/2/2001#, ThisWorkbook}]) is " & PrintableTableQ(anArray)
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
    
    Let v = Array()
    Debug.Print "Testing: Array()"
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
    
    Let v = Array()
    Debug.Print "Testing: Array()"
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
    
    Let v = Array()
    Debug.Print "Testing: Array()"
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
    
    Let v = Array()
    Debug.Print "Testing: Array()"
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

Public Sub TestFileNameJoin()
    Dim anArray As Variant
    
    Let anArray = Array("c:", "dir1", "dir2")
    Debug.Print "The result is " & FileNameJoin(anArray)
    
    Let anArray = Array("c:", "dir1", "file.txt")
    Debug.Print "The result is " & FileNameJoin(anArray)
    
    Let anArray = Array("c:", "dir1", "file.txt")
    Debug.Print "The result is " & FileNameJoin(anArray)
    
    Let anArray = Array()
    Debug.Print "The result is " & FileNameJoin(anArray)
End Sub

Public Sub TestFileNameSplit()
    Dim anArray As Variant
    
    Let anArray = Array("c:", "dir1", "dir2")
    Debug.Print "The result is"
    PrintArray FileNameSplit(FileNameJoin(anArray))
    Debug.Print
    
    Let anArray = Array("c:", "dir1", "file.txt")
    Debug.Print "The result is"
    PrintArray FileNameSplit(FileNameJoin(anArray))
    Debug.Print
    
    Let anArray = Empty
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
