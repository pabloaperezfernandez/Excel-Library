Attribute VB_Name = "TestedLibrary6Dot0"
Option Explicit
Option Base 1

'********************************************************************************************
' Predicates
'********************************************************************************************

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

Public Sub TestPredicatesMatrixQ()
    Debug.Print MatrixQ(Array(1, 2, 3))
    Debug.Print MatrixQ(1)
    Debug.Print MatrixQ(Empty)
    Debug.Print MatrixQ([{1,2; 3,4}])
    Debug.Print MatrixQ([{1,2; "A",4}])
End Sub

Public Sub TestPredicatesTableQ()
    Debug.Print TableQ(Array(1, 2, 3))
    Debug.Print TableQ(1)
    Debug.Print TableQ(Empty)
    Debug.Print TableQ([{1,2; 3,4}])
    Debug.Print TableQ([{1,2; "A",4}])
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
