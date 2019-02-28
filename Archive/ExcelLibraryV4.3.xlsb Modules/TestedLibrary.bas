Attribute VB_Name = "TestedLibrary"
Option Explicit
Option Base 1

Public Sub TestConvertDateToSerial()
    Dim AnArray As Variant
    
    Let AnArray = Array(#1/1/2011#, #1/2/2012#)
    PrintArray MapFunctionOverArray("ConvertDateToSerial", AnArray)
End Sub

Public Sub TestCenterShapeInRange()
    Dim AShape As Shape
    Dim var As Variant
    
    For Each var In TempComputation.Shapes: var.Delete: Next
    
    Set AShape = TempComputation.Shapes.AddFormControl(xlButtonControl, TempComputation.Range("B2").Left, TempComputation.Range("B2").Top, 100, 50)
    Let AShape.Name = "MyShape"
    
    Call CenterShapeInRange(AShape, TempComputation.Range("A1:D6"))
End Sub

Public Sub TestConstantArray()
    Dim a() As Variant
    Dim b As Variant
    
    ReDim a(3, 4)
    Debug.Print "The number of dims is " & NumberOfDimensions(a)
    Debug.Print "dim1 has bounds " & LBound(a, 1) & " to " & UBound(a, 1)
    Debug.Print "dim2 has bounds " & LBound(a, 2) & " to " & UBound(a, 2)
    
    ReDim a(100, 100000)
    Debug.Print "The number of dims is " & NumberOfDimensions(a)
    Debug.Print "dim1 has bounds " & LBound(a, 1) & " to " & UBound(a, 1)
    Debug.Print "dim2 has bounds " & LBound(a, 2) & " to " & UBound(a, 2)
    
    Let a = ConstantArray(Empty, 100000)
    Debug.Print "The number of dims is " & NumberOfDimensions(a)
    Debug.Print "dim1 has bounds " & LBound(a, 1) & " to " & UBound(a, 1)
    
    Let a = ConstantArray(Empty, 100, 50000)
    Debug.Print "The number of dims is " & NumberOfDimensions(a)
    Debug.Print "dim1 has bounds " & LBound(a, 1) & " to " & UBound(a, 1)
    Debug.Print "dim2 has bounds " & LBound(a, 2) & " to " & UBound(a, 2)
    
    Let a = ConstantArray("Test", 5, 10)
    Debug.Print "The number of dims is " & NumberOfDimensions(a)
    Debug.Print "dim1 has bounds " & LBound(a, 1) & " to " & UBound(a, 1)
    Debug.Print "dim2 has bounds " & LBound(a, 2) & " to " & UBound(a, 2)
    PrintArray a
End Sub

Public Sub TestEmsxTradeList()
    Dim ATradeList As EmsxTradeList
    
    Let Application.EnableEvents = False
    Let Application.DisplayAlerts = False
    Let Application.ScreenUpdating = False
    
    Debug.Print "Before opening a workbook, values are:"
    Debug.Print "Application.ScreenUpdating = " & Application.ScreenUpdating
    Debug.Print "Application.DisplayAlerts = " & Application.DisplayAlerts
    
    Set ATradeList = NewEmsxTradeList
    
    Call ATradeList.InitializeWithMySql(DbServerAddress, "etwip2dot0", "emsxtradinglists", "100969727705", "Deutsche Bank", DbUserName, DbPassword, "Equity", "Equity", "20140403")
    
    Debug.Print "After opening a workbook, values are:"
    Debug.Print "Application.ScreenUpdating = " & Application.ScreenUpdating
    Debug.Print "Application.DisplayAlerts = " & Application.DisplayAlerts

    Let Application.EnableEvents = True
    Let Application.DisplayAlerts = True
    Let Application.ScreenUpdating = True
End Sub

Public Sub TestHoldingsFromAa1()
    Dim aHoldingsReport As HoldingsFromAa
    
    Let Application.EnableEvents = False
    Let Application.DisplayAlerts = False
    Let Application.ScreenUpdating = False
    
    Debug.Print "Before opening a workbook, values are:"
    Debug.Print "Application.ScreenUpdating = " & Application.ScreenUpdating
    Debug.Print "Application.DisplayAlerts = " & Application.DisplayAlerts
    
    
    ' Instantiate the HoldingsFromAa object
    Set aHoldingsReport = NewHoldingsFromAa
    
    ' Initialize HoldingsFromAa instance
    Call aHoldingsReport.InitializeWithMySql(DbServerAddress, "etwip2dot0", "holdingsfromaa", "100969727705", "Deutsche Bank", DbUserName, DbPassword, "Equity", "Equity", "20140403")

    Debug.Print "After opening a workbook, values are:"
    Debug.Print "Application.ScreenUpdating = " & Application.ScreenUpdating
    Debug.Print "Application.DisplayAlerts = " & Application.DisplayAlerts

    Let Application.EnableEvents = True
    Let Application.DisplayAlerts = True
    Let Application.ScreenUpdating = True
End Sub

Public Sub TestHoldingsFromAa2()
    Dim aHoldingsReport As HoldingsFromAa
    Dim TheFileName As String
    
    ' Set the input file's name
    Let TheFileName = "X:\WorkProduct\SourceCode\VBA\TestingFiles\ExcelLibraryV2.0-NoFactSetAddIn\HoldingsFromAa-20140227.xlsb"
    
    ' Instantiate the HoldingsFromAa object
    Set aHoldingsReport = NewHoldingsFromAa
    
    ' Initialize HoldingsFromAa instance
    Call aHoldingsReport.InitializeWithTradarView(TheFileName)
    
    ' Inject the instance into the DB
    Call aHoldingsReport.InjectIntoMySql("localhost", "etwip2dot0", "holdingsfromaa", "root", "")
End Sub

Public Sub TestInsert()
    Dim a As Variant
    
    Let a = Array(1, 2, 3, 4)
    Debug.Print "Let a = Array(1, 2, 3, 4)"
    Debug.Print
    
    Debug.Print "Insert(a, 5, 5):"
    PrintArray Insert(a, 5, 5)
    Debug.Print
    
    Debug.Print "Insert(a, 0, 1):"
    PrintArray Insert(a, 0, 1)
    Debug.Print
    
    Debug.Print "Insert(a, 3.5, 4):"
    PrintArray Insert(a, 3.5, 4)
    Debug.Print
    
    Let a = [{1,1,1; 2,2,2; 3,3,3; 4,4,4}]
    Debug.Print "Let a = [{1,1,1; 2,2,2; 3,3,3; 4,4,4}]"
    Debug.Print
    
    Debug.Print "Insert(a, Array(5,5,5), 5):"
    PrintArray Insert(a, Array(5, 5, 5), 5)
    Debug.Print
    
    Debug.Print "Insert(a, Array(0,0,0), 1):"
    PrintArray Insert(a, Array(0, 0, 0), 1)
    Debug.Print
    
    Debug.Print "Insert(a, Array(3.5, 3.5,3.5), 4):"
    PrintArray Insert(a, Array(3.5, 3.5, 3.5), 4)
    Debug.Print
    
    Debug.Print "Insert(a, Array(3.5, 3.5,3.5), 6):"
    PrintArray Insert(a, Array(3.5, 3.5, 3.5), 6)
    Debug.Print
    
    Debug.Print "Insert(a, Array(3.5, 3.5,3.5), 0):"
    PrintArray Insert(a, Array(3.5, 3.5, 3.5), 0)
    Debug.Print
End Sub

Public Sub TestPack2DArray()
    Dim a As Variant
    
    Let a = Array(Array(1, 2, 3), Array(4, 5, 6))
    
    Debug.Print "Let a = Array(Array(1, 2, 3), Array(4, 5, 6))"
    Debug.Print "Pack2DArray(a) is:"
    PrintArray Pack2DArray(a)
    
    Debug.Print
    Let a = TransposeMatrix(Pack2DArray(Array(Array(1, 2, 3), Array(4, 5, 6))))
    Debug.Print "Let a = TransposeMatrix(Pack2DArray(Array(Array(1, 2, 3), Array(4, 5, 6))))"
    PrintArray a

End Sub

' This tests ArrayFormulas.Predicates
Public Sub TestIsNumericArrayQ()
    Dim a As Variant
    
    Let a = 3
    Debug.Print "Let a = 3"
    Debug.Print "IsNumericArrayQ(a) is " & IsNumericArrayQ(a)
    Debug.Print
    
    Let a = "asf"
    Debug.Print "Let a = ""asf"""
    Debug.Print "IsNumericArrayQ(a) is " & IsNumericArrayQ(a)
    Debug.Print
    
    Let a = Array()
    Debug.Print "Let a = Array()"
    Debug.Print "IsNumericArrayQ(a) is " & IsNumericArrayQ(a)
    Debug.Print
    
    Let a = Array(1, 2, 3, Array())
    Debug.Print "Let a = Array(1, 2, 3, Array())"
    Debug.Print "IsNumericArrayQ(a) is " & IsNumericArrayQ(a)
    Debug.Print
    
    Let a = Array(1, 2, 3)
    Debug.Print "Let a = Array(1, 2, 3)"
    Debug.Print "IsNumericArrayQ(a) is " & IsNumericArrayQ(a)
    Debug.Print
    
    Let a = [{1,3,5; 2,4,6}]
    Debug.Print "Let a = [{1,3,5; 2,4,6}]"
    Debug.Print "IsNumericArrayQ(a) is " & IsNumericArrayQ(a)
    Debug.Print
End Sub

' This tests ArrayFormulas.Take
Public Sub TestTake()
    Dim a As Variant
    
    Let a = Array(ConstantArray(1, 20000), ConstantArray(1, 20000), ConstantArray(3, 20000), ConstantArray(4, 20000))
    Debug.Print "a = Array(ConstantArray(1, 20000), ConstantArray(2, 20000), ConstantArray(3, 20000), ConstantArray(4, 20000))"
    Let a = Pack2DArray(a)
    Let a = Take(a, Array(2, 3))
    Debug.Print "Let a = Pack2DArray(a)"
    Debug.Print "Let a = Take(a, Array(2, 3))"
    Debug.Print "a has # dims = " & NumberOfDimensions(a)
    Debug.Print "dim1 is " & GetNumberOfRows(a) & " and dim2 is " & GetNumberOfColumns(a)
    Debug.Print
    
    Let a = Array(1, 2, 3, 4, 5, 6, 7)
    Debug.Print "Set a = Array(1, 2, 3, 4, 5, 6, 7)"
    Debug.Print "Testing Take(a,1)"
    PrintArray Take(a, 1)
    Debug.Print
    
    Debug.Print "Testing Take(a,0)"
    PrintArray Take(a, 0)
    Debug.Print
    
    Debug.Print "Testing Take(a,4)"
    PrintArray Take(a, 4)
    Debug.Print
    
    Debug.Print "Testing Take(a,10)"
    PrintArray Take(a, 10)
    Debug.Print
    
    Debug.Print "Testing Take(a,-1)"
    PrintArray Take(a, -1)
    Debug.Print

    Debug.Print "Testing Take(a,array(-1))"
    PrintArray Take(a, Array(-1))
    Debug.Print

    Debug.Print "Testing Take(a,-3)"
    PrintArray Take(a, -3)
    Debug.Print
    
    Debug.Print "Testing Take(a,array(-3))"
    PrintArray Take(a, Array(-3))
    Debug.Print
    
    Debug.Print "Testing Take(a,-10)"
    PrintArray Take(a, -10)
    Debug.Print
    
    Debug.Print "Testing Take(a,array(-10))"
    PrintArray Take(a, Array(-10))
    Debug.Print
    
    Debug.Print "Testing Take(a, Array(1))"
    PrintArray Take(a, Array(1))
    Debug.Print

    Debug.Print "Testing Take(a, Array(2))"
    PrintArray Take(a, Array(2))
    Debug.Print

    Debug.Print "Testing Take(a, Array(7))"
    PrintArray Take(a, Array(7))
    Debug.Print
    
    Debug.Print "Testing Take(a, Array(0))"
    PrintArray Take(a, Array(0))
    Debug.Print
    
    Debug.Print "Testing Take(a, Array(8))"
    PrintArray Take(a, Array(8))
    Debug.Print

    Debug.Print "Testing Take(a, Array(2,4))"
    PrintArray Take(a, Array(2, 4))
    Debug.Print
    
    Debug.Print "Testing Take(Array(1,2,3,4), Array(2,4))"
    PrintArray Take(Array(1, 2, 3, 4), Array(2, 4))
    
    Debug.Print "Testing Take(a, Array(-2,-4))"
    PrintArray Take(a, Array(-2, -4))
    Debug.Print

    Let a = [{1,1,1; 2,2,2; 3,3,3; 4,4,4; 5,5,5; 6,6,6; 7,7,7; 8,8,8; 9,9,9}]
    Debug.Print "Set a = [{1,1,1; 2,2,2; 3,3,3; 4,4,4; 5,5,5; 6,6,6; 7,7,7; 8,8,8; 9,9,9}]"
    Debug.Print "Testing Take(a,1)"
    PrintArray Take(a, 1)
    Debug.Print
    
    Debug.Print "Testing Take(a,0)"
    PrintArray Take(a, 0)
    Debug.Print
    
    Debug.Print "Testing Take(a,4)"
    PrintArray Take(a, 4)
    Debug.Print
    
    Debug.Print "Testing Take(a,10)"
    PrintArray Take(a, 10)
    Debug.Print
    
    Debug.Print "Testing Take(a,-1)"
    PrintArray Take(a, -1)
    Debug.Print
    
    Debug.Print "Testing Take(a,Array(-1))"
    PrintArray Take(a, Array(-1))
    Debug.Print
    
    Debug.Print "Testing Take(a,-3)"
    PrintArray Take(a, -3)
    Debug.Print
    
    Debug.Print "Testing Take(a,Array(-3))"
    PrintArray Take(a, Array(-3))
    Debug.Print
    
    Debug.Print "Testing Take(a,-10)"
    PrintArray Take(a, -10)
    Debug.Print
    
    Debug.Print "Testing Take(a,Array(-10))"
    PrintArray Take(a, Array(-10))
    Debug.Print
    
    Debug.Print "Testing Take(a,-20)"
    PrintArray Take(a, -20)
    Debug.Print
    
    Debug.Print "Testing Take(Array(),1)"
    PrintArray Take(Array(), 1)
    Debug.Print
    
    Debug.Print "Testing Take(a,Array(1))"
    PrintArray Take(a, Array(1))
    Debug.Print
    
    Debug.Print "Testing Take(a,Array(-1))"
    PrintArray Take(a, Array(-1))
    Debug.Print
    
    Debug.Print "Testing Take(a,Array(9))"
    PrintArray Take(a, Array(9))
    Debug.Print

    Debug.Print "Testing Take(a,Array(-9))"
    PrintArray Take(a, Array(-9))
    Debug.Print

    Debug.Print "Testing Take(a,Array(10))"
    PrintArray Take(a, Array(10))
    Debug.Print

    Debug.Print "Testing Take(a,Array(0))"
    PrintArray Take(a, Array(0))
    Debug.Print
    
    Debug.Print "Testing Take(a,Array(2,4,5))"
    PrintArray Take(a, Array(2, 4, 5))
    Debug.Print
    
    Debug.Print "Testing Take(a,Array(-2,-4,-5))"
    PrintArray Take(a, Array(-2, -4, -5))
    Debug.Print
    
    Debug.Print "Testing Take(a,Array())"
    PrintArray Take(a, Array())
    Debug.Print
End Sub

' This tests ArrayFormulas.Drop
Public Sub TestDrop()
    Dim a As Variant
    
    Let a = Array(1, 2, 3, 4, 5, 6, 7)
    Debug.Print "Set a = Array(1, 2, 3, 4, 5, 6, 7)"
    Debug.Print
    
    Debug.Print "Testing Drop(a,1)"
    PrintArray Drop(a, 1)
    Debug.Print
    
    Debug.Print "Testing Drop(a,0)"
    PrintArray Drop(a, 0)
    Debug.Print
    
    Debug.Print "Testing Drop(a,4)"
    PrintArray Drop(a, 4)
    Debug.Print
    
    Debug.Print "Testing Drop(a,10)"
    PrintArray Drop(a, 10)
    Debug.Print
    
    Debug.Print "Testing Drop(a,-1)"
    PrintArray Drop(a, -1)
    Debug.Print

    Debug.Print "Testing Drop(a,array(-1))"
    PrintArray Drop(a, Array(-1))
    Debug.Print

    Debug.Print "Testing Drop(a,-3)"
    PrintArray Drop(a, -3)
    Debug.Print
    
    Debug.Print "Testing Drop(a,array(-3))"
    PrintArray Drop(a, Array(-3))
    Debug.Print
    
    Debug.Print "Testing Drop(a,-10)"
    PrintArray Drop(a, -10)
    Debug.Print
    
    Debug.Print "Testing Drop(a,array(-10))"
    PrintArray Drop(a, Array(-10))
    Debug.Print
    
    Debug.Print "Testing Drop(a, Array(1))"
    PrintArray Drop(a, Array(1))
    Debug.Print

    Debug.Print "Testing Drop(a, Array(2))"
    PrintArray Drop(a, Array(2))
    Debug.Print

    Debug.Print "Testing Drop(a, Array(7))"
    PrintArray Drop(a, Array(7))
    Debug.Print
    
    Debug.Print "Testing Drop(a, Array(0))"
    PrintArray Drop(a, Array(0))
    Debug.Print
    
    Debug.Print "Testing Drop(a, Array(8))"
    PrintArray Drop(a, Array(8))
    Debug.Print

    Debug.Print "Testing Drop(a, Array(2,4))"
    PrintArray Drop(a, Array(2, 4))
    Debug.Print
    
    Debug.Print "Testing Drop(a, Array(-2,-4))"
    PrintArray Drop(a, Array(-2, -4))
    Debug.Print

    Let a = [{1,1,1; 2,2,2; 3,3,3; 4,4,4; 5,5,5; 6,6,6; 7,7,7; 8,8,8; 9,9,9}]
    Debug.Print "Set a = [{1,1,1; 2,2,2; 3,3,3; 4,4,4; 5,5,5; 6,6,6; 7,7,7; 8,8,8; 9,9,9}]"
    Debug.Print "Testing Drop(a,1)"
    PrintArray Drop(a, 1)
    Debug.Print
    
    Debug.Print "Testing Drop(a,0)"
    PrintArray Drop(a, 0)
    Debug.Print
    
    Debug.Print "Testing Drop(a,4)"
    PrintArray Drop(a, 4)
    Debug.Print
    
    Debug.Print "Testing Drop(a,10)"
    PrintArray Drop(a, 10)
    Debug.Print
    
    Debug.Print "Testing Drop(a,-1)"
    PrintArray Drop(a, -1)
    Debug.Print
    
    Debug.Print "Testing Drop(a,Array(-1))"
    PrintArray Drop(a, Array(-1))
    Debug.Print
    
    Debug.Print "Testing Drop(a,-3)"
    PrintArray Drop(a, -3)
    Debug.Print
    
    Debug.Print "Testing Drop(a,Array(-3))"
    PrintArray Drop(a, Array(-3))
    Debug.Print
    
    Debug.Print "Testing Drop(a,-10)"
    PrintArray Drop(a, -10)
    Debug.Print
    
    Debug.Print "Testing Drop(a,Array(-10))"
    PrintArray Drop(a, Array(-10))
    Debug.Print
    
    Debug.Print "Testing Drop(a,-20)"
    PrintArray Drop(a, -20)
    Debug.Print
    
    Debug.Print "Testing Drop(Array(),1)"
    PrintArray Drop(Array(), 1)
    Debug.Print
    
    Debug.Print "Testing Drop(a,Array(1))"
    PrintArray Drop(a, Array(1))
    Debug.Print
    
    Debug.Print "Testing Drop(a,Array(-1))"
    PrintArray Drop(a, Array(-1))
    Debug.Print
    
    Debug.Print "Testing Drop(a,Array(9))"
    PrintArray Drop(a, Array(9))
    Debug.Print

    Debug.Print "Testing Drop(a,Array(-9))"
    PrintArray Drop(a, Array(-9))
    Debug.Print

    Debug.Print "Testing Drop(a,Array(10))"
    PrintArray Drop(a, Array(10))
    Debug.Print

    Debug.Print "Testing Drop(a,Array(0))"
    PrintArray Drop(a, Array(0))
    Debug.Print
    
    Debug.Print "Testing Drop(a,Array(2,4,5))"
    PrintArray Drop(a, Array(2, 4, 5))
    Debug.Print
    
    Debug.Print "Testing Drop(a,Array(-2,-4,-5))"
    PrintArray Drop(a, Array(-2, -4, -5))
    Debug.Print
    
    Debug.Print "Testing Drop(a,Array())"
    PrintArray Drop(a, Array())
    Debug.Print
End Sub

' This tests processing Bloomberg's new alerts, injecting them into the DB and moving the source emails to the archive directory
Public Sub TestBloombergNewsAlerts1()
    Dim TheBloombergAlerts As BloombergNewsAlerts
        
    Let Application.DisplayAlerts = False
    
    Set TheBloombergAlerts = New BloombergNewsAlerts
    
    Call TheBloombergAlerts.InitializeWithOutlook
    
    Call TheBloombergAlerts.InjectIntoMySql

    Call TheBloombergAlerts.ArchiveAlerts

    Let Application.DisplayAlerts = True
End Sub

' This tests poulating a spreadsheet with today's news alerts
Public Sub TestBloombergNewsAlerts2()
    Dim TheBloombergAlerts As BloombergNewsAlerts
        
    Let Application.DisplayAlerts = False

    Set TheBloombergAlerts = New BloombergNewsAlerts

    Call TheBloombergAlerts.InitializeWithMySql(StartDate:=Date, EndDate:=Date)

    Let Application.DisplayAlerts = True
End Sub

Public Sub TestGetColumn()
    Dim m As Variant
    
    Let m = [{1,2,3;4,5,6;7,8,9}]
    Debug.Print "Testing getting column 1 from [{1,2,3;4,5,6;7,8,9}]"
    PrintArray GetColumn(m, 1)
    Debug.Print
    
    Let m = [{1,2,3;4,5,6;7,8,9}]
    Debug.Print "Testing getting column 10 from [{1,2,3;4,5,6;7,8,9}]"
    PrintArray GetColumn(m, 10)
    Debug.Print
    
    Let m = [{1,2,3;4,5,6;7,8,9}]
    Debug.Print "Testing getting column 0 from [{1,2,3;4,5,6;7,8,9}]"
    PrintArray GetColumn(m, 0)
    Debug.Print

    Let m = [{1,2,3}]
    Debug.Print "Testing getting column 2 from [{1,2,3}]"
    PrintArray GetColumn(m, 2)
    Debug.Print
    
    Let m = Array(1, 2, 3)
    Debug.Print "Testing getting column 2 from Array(1, 2, 3)"
    PrintArray GetColumn(m, 2)
    Debug.Print
    
    Let m = Array()
    Debug.Print "Testing  Array()"
    PrintArray GetColumn(m, 1)
    Debug.Print
End Sub

Public Sub TestGetRow()
    Dim m As Variant
    Dim a(0 To 2, 0 To 2) As Integer
    Dim i As Integer, j As Integer
    
    For i = 0 To 2
        For j = 0 To 2
            Let a(i, j) = i + j * 3
        Next j
    Next i
    
    Debug.Print "The Array is:"
    PrintArray a
    Debug.Print
    
    Debug.Print "First row is:"
    PrintArray GetRow(a, 1)
    Debug.Print
        
    Debug.Print "Second row is:"
    PrintArray GetRow(a, 2)
    Debug.Print
    
    Debug.Print "Third row is:"
    PrintArray GetRow(a, 3)
    Debug.Print
    
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
    
    Let m = Array()
    Debug.Print "Testing  Array()"
    PrintArray GetRow(m, 1)
    Debug.Print
End Sub

Public Sub TestApplicationTransposeApplicationTranspose()
    Dim a As Variant
    
    Let a = Application.Transpose(Application.Transpose(Array(Array(1, 2), Array(3, 4))))
    
    PrintArray a
    
    Debug.Print
    Debug.Print "a has " & NumberOfDimensions(a) & " dimensions."
End Sub

' This example shows how Application.Transpose(Application.Transpose(a)) changes indices to 1 to n from either 0/1 to n-1/n
Public Sub TestDictionaryItemsArray()
    Dim aDict As New Dictionary
    Dim i As Integer
    Dim a As Variant
    Dim b As Variant
    
    For i = 1 To 10
        Call aDict.Add(Key:="Key" & i, Item:="Item" & i)
    Next i
    
    Let a = aDict.Items
    Let b = Application.Transpose(Application.Transpose(a))
End Sub

Public Sub TestConvertTo1DArray()
    Dim v As Variant
    Dim a() As Variant
    Dim b() As Variant
    
    Let v = [{1, 2, 3; 4,5,6; 7,8,9}]
    Debug.Print "Testing: [{1, 2, 3; 4,5,6; 7,8,9}]"
    PrintArray ConvertTo1DArray(v)
    Debug.Print

    ReDim a(1, 3)
    Let a(1, 1) = 1
    Let a(1, 2) = 2
    Let a(1, 3) = 3
    Debug.Print "Testing: [{1, 2, 3}]"
    PrintArray ConvertTo1DArray(a)
    Debug.Print

    Let v = [{1; 2; 3}]
    Debug.Print "Testing: [{1; 2; 3}]"
    PrintArray ConvertTo1DArray(v)
    Debug.Print

    Let v = Array([{1; 2; 3}])
    Debug.Print "Testing: Array([{1; 2; 3}])"
    PrintArray ConvertTo1DArray(v)
    Debug.Print
    
    Let v = [{[{1;2;3}]; 4; 5}]
    Debug.Print "Testing: [{[{1;2;3}]; 4; 5}]"
    PrintArray ConvertTo1DArray(v)
    Debug.Print

    ReDim a(1, 1)
    Let a(1, 1) = [{1;2;3}]
    Debug.Print "Testing: [{[{1;2;3}]}]"
    PrintArray ConvertTo1DArray(a)
    Debug.Print

    Let v = Array(1, 2, 3)
    Debug.Print "Testing: Array(1, 2, 3)"
    PrintArray ConvertTo1DArray(v)
    Debug.Print
    
    Let v = Array(Array(1, 2, 3))
    Debug.Print "Testing: Array(Array(1, 2, 3))"
    PrintArray ConvertTo1DArray(v)
    Debug.Print
    
    Let v = Array(Array(Array(1, 2, 3)))
    Debug.Print "Testing: Array(Array(Array(1, 2, 3)))"
    PrintArray ConvertTo1DArray(v)
    Debug.Print
    
    Let v = Array([{1, 2, 3}])
    Debug.Print "Testing: Array([{1, 2, 3}])"
    PrintArray ConvertTo1DArray(v)
    Debug.Print
    
    ReDim a(1 To 1, 1 To 1)
    Let a(1, 1) = Array(1, 2, 3)
    Debug.Print "Testing: [{Array(1, 2, 3)}]"
    PrintArray ConvertTo1DArray(a)
    Debug.Print
    
    ReDim a(1 To 1, 1 To 1)
    ReDim b(1 To 1, 1 To 1)
    Let v = [{1;2;3}]
    Let a(1, 1) = v
    Let b(1, 1) = a
    Debug.Print "Testing: [{[{[{1; 2; 3}]}]}]"
    PrintArray ConvertTo1DArray(v)
    Debug.Print
    
    Let v = Array()
    Debug.Print "Testing: Array()"
    PrintArray ConvertTo1DArray(v)
    Debug.Print
End Sub

Public Sub TestColumnArrayQ()
    Dim v As Variant
    Dim a() As Variant
    Dim b() As Variant
    
    Let v = [{1, 2, 3; 4,5,6; 7,8,9}]
    Debug.Print "Testing: [{1, 2, 3; 4,5,6; 7,8,9}]"
    Debug.Print ColumnArrayQ(v)
    Debug.Print

    ReDim a(1, 3)
    Let a(1, 1) = 1
    Let a(1, 2) = 2
    Let a(1, 3) = 3
    Debug.Print "Testing: [{1, 2, 3}]"
    Debug.Print ColumnArrayQ(a)
    Debug.Print

    Let v = [{1; 2; 3}]
    Debug.Print "Testing: [{1; 2; 3}]"
    Debug.Print ColumnArrayQ(v)
    Debug.Print

    Let v = Array([{1; 2; 3}])
    Debug.Print "Testing: Array([{1; 2; 3}])"
    Debug.Print ColumnArrayQ(v)
    Debug.Print
    
    Let v = [{[{1;2;3}]; 4; 5}]
    Debug.Print "Testing: [{[{1;2;3}]; 4; 5}]"
    Debug.Print ColumnArrayQ(v)
    Debug.Print

    ReDim a(1, 1)
    Let a(1, 1) = [{1;2;3}]
    Debug.Print "Testing: [{[{1;2;3}]}]"
    Debug.Print ColumnArrayQ(a)
    Debug.Print

    Let v = Array(1, 2, 3)
    Debug.Print "Testing: Array(1, 2, 3)"
    Debug.Print ColumnArrayQ(v)
    Debug.Print
    
    Let v = Array(Array(1, 2, 3))
    Debug.Print "Testing: Array(Array(1, 2, 3))"
    Debug.Print ColumnArrayQ(v)
    Debug.Print
    
    Let v = Array(Array(Array(1, 2, 3)))
    Debug.Print "Testing: Array(Array(Array(1, 2, 3)))"
    Debug.Print ColumnArrayQ(v)
    Debug.Print
    
    Let v = Array([{1, 2, 3}])
    Debug.Print "Testing: Array([{1, 2, 3}])"
    Debug.Print ColumnArrayQ(v)
    Debug.Print
    
    ReDim a(1 To 1, 1 To 1)
    Let a(1, 1) = Array(1, 2, 3)
    Debug.Print "Testing: [{Array(1, 2, 3)}]"
    Debug.Print ColumnArrayQ(a)
    Debug.Print
    
    ReDim a(1 To 1, 1 To 1)
    ReDim b(1 To 1, 1 To 1)
    Let v = [{1;2;3}]
    Let a(1, 1) = v
    Let b(1, 1) = a
    Debug.Print "Testing: [{[{[{1; 2; 3}]}]}]"
    Debug.Print ColumnArrayQ(v)
    Debug.Print
    
    Let v = Array()
    Debug.Print "Testing: Array()"
    Debug.Print ColumnArrayQ(v)
    Debug.Print
End Sub

Public Sub TestRowArrayQ()
    Dim v As Variant
    Dim a() As Variant
    
    Let v = [{1, 2, 3; 4,5,6; 7,8,9}]
    Debug.Print "Testing: [{1, 2, 3; 4,5,6; 7,8,9}]"
    Debug.Print RowArrayQ(v)
    Debug.Print

    Let v = [{1, 2, 3}]
    Debug.Print "Testing: [{1, 2, 3}]"
    Debug.Print RowArrayQ(v)
    Debug.Print

    ReDim a(1 To 1, 1 To 1)
    Let a(1, 1) = [{1, 2, 3}]
    Debug.Print "Testing: [{[{1, 2, 3}]}]"
    Debug.Print RowArrayQ(a)
    Debug.Print

    Let v = [{1; 2; 3}]
    Debug.Print "Testing: [{1; 2; 3}]"
    Debug.Print RowArrayQ(v)
    Debug.Print

    Let v = Array(1, 2, 3)
    Debug.Print "Testing: Array(1, 2, 3)"
    Debug.Print RowArrayQ(v)
    Debug.Print
    
    Let v = Array(Array(1, 2, 3))
    Debug.Print "Testing: Array(Array(1, 2, 3))"
    Debug.Print RowArrayQ(v)
    Debug.Print
    
    Let v = Array(Array(Array(1, 2, 3)))
    Debug.Print "Testing: Array(Array(Array(1, 2, 3)))"
    Debug.Print RowArrayQ(v)
    Debug.Print
    
    Let v = Array([{1, 2, 3}])
    Debug.Print "Testing: Array([{1, 2, 3}])"
    Debug.Print RowArrayQ(v)
    Debug.Print
    
    ReDim a(1 To 1, 1 To 1)
    Let a(1, 1) = Array(1, 2, 3)
    Debug.Print "Testing: [{Array(1, 2, 3)}]"
    Debug.Print RowArrayQ(a)
    Debug.Print
    
    Let v = Array()
    Debug.Print "Testing: Array()"
    Debug.Print RowArrayQ(v)
    Debug.Print
End Sub

Public Sub TestMatrixQ()
    Dim v As Variant
    
    Let v = [{1, 2, 3; 4,5,6; 7,8,9}]
    Debug.Print "Testing: [{1, 2, 3; 4,5,6; 7,8,9}]"
    Debug.Print MatrixQ(v)
    Debug.Print
    
    Let v = [{1; 2; 3}]
    Debug.Print "Testing: [{1; 2; 3}]"
    Debug.Print MatrixQ(v)
    Debug.Print

    Let v = Array(1, 2, 3)
    Debug.Print "Testing: Array(1, 2, 3)"
    Debug.Print MatrixQ(v)
    Debug.Print
    
    Let v = Array(Array(1, 2, 3))
    Debug.Print "Testing: Array(Array(1, 2, 3))"
    Debug.Print MatrixQ(v)
    Debug.Print
    
    Let v = Array()
    Debug.Print "Testing: Array()"
    Debug.Print MatrixQ(v)
    Debug.Print
    
    Let v = Array(Array(1, 2, 3), Array(4, 5, 6), Array(7, 8, 9))
    Debug.Print "Testing: Array(Array(1, 2, 3), Array(4, 5, 6), Array(7, 8, 9))"
    Debug.Print MatrixQ(v)
    Debug.Print
End Sub

Public Sub TestVectorQ()
    Dim v As Variant
    
    Let v = Array(1, 2, 3)
    Debug.Print "Testing: Array(1, 2, 3)"
    Debug.Print VectorQ(v)
    Debug.Print
    
    Let v = [{1; 2; 3}]
    Debug.Print "Testing: [{1; 2; 3}]"
    Debug.Print VectorQ(v)
    Debug.Print
    
    Let v = Array(Array(1, 2, 3))
    Debug.Print "Testing: Array(Array(1, 2, 3))"
    Debug.Print VectorQ(v)
    Debug.Print
    
    Let v = Array()
    Debug.Print "Testing: Array()"
    Debug.Print VectorQ(v)
    Debug.Print
End Sub

Public Sub TestColumnVectorQ()
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

Public Sub TestRowVectorQ()
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

Public Sub TestStack2DArrays()
    Dim a As Variant
    
    Debug.Print "Testing Stack2DArrays(Array(1, 2, 3), Array(4, 5, 6)):"
    PrintArray Stack2DArrays(Array(1, 2, 3), Array(4, 5, 6))
    Debug.Print
    
    Debug.Print "Testing Stack2DArrays([{1,2,3,4; 5,6,7,8}], Array(9,10,11,12)):"
    PrintArray Stack2DArrays([{1,2,3,4; 5,6,7,8}], Array(9, 10, 11, 12))
    Debug.Print
    
    Debug.Print "Testing Stack2DArrays(Array(9,10,11,12), [{1,2,3,4; 5,6,7,8}]):"
    PrintArray Stack2DArrays(Array(9, 10, 11, 12), [{1,2,3,4; 5,6,7,8}])
    Debug.Print
    
    Debug.Print "Testing Stack2DArrays([{1,2,3,4; 5,6,7,8}], [{10,20,30,40; 50,60,70,80}]):"
    PrintArray Stack2DArrays([{1,2,3,4; 5,6,7,8}], [{10,20,30,40; 50,60,70,80}])
    Debug.Print

    Debug.Print "Testing Stack2DArrays([{1,2,3,4; 5,6,7,8}], [{10,20,30,40; 50,60,70,80}]):"
    PrintArray Stack2DArrays([{1,"02", 3 ,4; 5,"06",7,8}], [{10,"020",30,40; 50,"060",70,80}])
    Debug.Print "Look at worksheet TempComputation to see if formats were preserved."
    Call ToTemp(Stack2DArrays([{1,"02", 3 ,4; 5,"06",7,8}], [{10,20,"030",40; 50,60,"070",80}]))

    Debug.Print "Testing Stack2DArrays([{1,2,3,4; 5,6,7,8}], [{10,20,30,40; 50,60,70,80}]):"
    Let a = Stack2DArrays([{1,"02", 3 ,4; 5,"06",7,8}], [{10,"020",30,40; 50,"060",70,80}])
    PrintArray a
    Debug.Print "The dimensions of the stacked array are #rows = " & GetNumberOfRows(a) & ", #cols = " & GetNumberOfColumns(a)
    Debug.Print "Look at worksheet TempComputation to see if formats were preserved."
    Call ToTemp(a, True)
End Sub

Public Sub TestStack2DArrayAs1DArray()
    Dim m As Variant
    Dim m2(0 To 3) As Variant
    Dim m3(0 To 3, 0 To 1) As Variant
    
    Let m = Array(1, 2, 3, 4)
    Debug.Print "Testing:"
    PrintArray m
    Debug.Print "We get:"
    PrintArray Stack2DArrayAs1DArray(m)
    Debug.Print

    Let m = [{1; 2; 3; 4}]
    Debug.Print "Testing:"
    PrintArray m
    Debug.Print "We get:"
    PrintArray Stack2DArrayAs1DArray(m)
    Debug.Print

    Let m = [{1,2; 3,4; 5,6; 7,8}]
    Debug.Print "Testing:"
    PrintArray m
    Debug.Print "We get:"
    PrintArray Stack2DArrayAs1DArray(m)
    Debug.Print

    Let m2(0) = 0
    Let m2(1) = 1
    Let m2(2) = 2
    Let m2(3) = 3
    Debug.Print "Testing a 1D array indexed at 0:"
    PrintArray m2
    Debug.Print "We get:"
    PrintArray Stack2DArrayAs1DArray(m2)
    Debug.Print

    Let m3(0, 0) = 0
    Let m3(1, 0) = 1
    Let m3(2, 0) = 2
    Let m3(3, 0) = 3
    Let m3(0, 1) = 10
    Let m3(1, 1) = 11
    Let m3(2, 1) = 12
    Let m3(3, 1) = 13
    Debug.Print "Testing a 1D array indexed at 0:"
    PrintArray m3
    Debug.Print "We get:"
    PrintArray Stack2DArrayAs1DArray(m3)
    Debug.Print
End Sub

Public Sub TestGetSubMatrix()
    Dim m As Variant
    
    Let m = [{1,2,3; 4,5,6; 7,8,9}]
    PrintArray GetSubMatrix(m, 2, 3, 1, 2)
    Debug.Print
    
    Let m = Array(1, 2, 3, 4, 5, 6, 7)
    PrintArray GetSubMatrix(m, BottomColumnNumber:=3, TopColumnNumber:=5)
    Debug.Print
    
    Let m = [{1;2;3;4;5;6;7;8;9;10}]
    PrintArray GetSubMatrix(m, 2, 6)
    Debug.Print

    Let m = Array()
    PrintArray GetSubMatrix(m, 2, 6)
    Debug.Print

    Let m = Array(2)
    PrintArray GetSubMatrix(m, 1, 1)
    Debug.Print
End Sub

' This function was used to determine the relative speed of various methods to extract a sub-matrix from a 2D array
' Clearly, for loops work the fastests in VBA.
Public Sub NewTestGetSubMatrix()
    Dim m As Variant
    Dim result0 As Variant
    Dim result As Variant
    Dim r As Variant
    Dim c As Variant
    Dim SubMatrix() As Variant

    Let TempComputation.Range("A1").Resize(20000, 300).Formula = "=rand()"
    Let m = TempComputation.Range("A1").CurrentRegion.Value2
    
    Debug.Print "Testing Looping:"
    Debug.Print "Starting time: " & Now
    Let result0 = GetSubMatrix(m, 2000, 20000, 2, 300)
    Debug.Print "Ending time: " & Now
    Debug.Print
    
    Debug.Print "Testing Looping:"
    Debug.Print "Starting time: " & Now
    ReDim SubMatrix(18001, 299)
    For r = 2000 To 20000
        For c = 2 To 300
            Let SubMatrix(r + 1 - 2000, c + 1 - 2) = m(r, c)
        Next c
    Next r
    Debug.Print "Ending time: " & Now
    Debug.Print
    
    Debug.Print "Testing Slicing:"
    Debug.Print "Starting time: " & Now
    Let result = NewGetSubMatrix(m, 2000, 8000, 2, 100)
    Debug.Print "Ending time: " & Now
    Debug.Print
End Sub

' This performs an HTTP post call
Public Sub TestHTTP()
    With ActiveSheet.QueryTables.Add(Connection:="URL;http://perezhortinelafamily.us/WebServiceSamples/SimplePhpQuerry/SimplePhpQuerry.php ", Destination:=Range("A2"))
        .PostText = "name=Arno&place=Amsterdam"
        .RefreshStyle = xlOverwriteCells
        .SaveData = True
        .Refresh
    End With
End Sub

Public Sub TestRest()
    PrintArray Rest(Array(1, 2, 3))
    PrintArray Rest(Array(1, 2))
    PrintArray Rest(Array(1))
    PrintArray Rest(Array())
End Sub

Public Sub TestIifAndOneLineIf()
    Debug.Print "IIF test " & IIf(3 < 4, 3, 4)
    
    If 3 < 4 Then Debug.Print "3 is less than 4"
    
    Debug.Print "We are out of the if"
End Sub

Public Sub TestTextFileReading()
    Dim FsoObj As Scripting.FileSystemObject
    Dim FileObj As Scripting.TextStream
    Dim ATxtLine As String
    Dim TheFileName As String
    
    Let TheFileName = "X:\TestProductionEnvironment\SeimAudit\Inputs\DailyAuditLogs\TestLog.txt"
    
    Set FsoObj = New Scripting.FileSystemObject
    Set FileObj = FsoObj.OpenTextFile(Filename:=TheFileName, IOMode:=ForReading)
    
    Do Until FileObj.AtEndOfLine
        Let ATxtLine = FileObj.ReadLine
        Debug.Print ATxtLine
    Loop
    
    Call FileObj.Close

End Sub

Public Sub TestStringConcatenate()
    Debug.Print StringConcatenate(Array("Pablo is", " awesome."))
    Debug.Print StringConcatenate(Array("Paulina is", "", " awesome."))
    Debug.Print StringConcatenate(Array())
End Sub

Public Sub TestGetFileNames()
    Dim ThePath As String

    Let ThePath = "X:\TestProductionEnvironment\ET Wip Production Directory\ET Wip Input Directory\HoldingsFiles\"
    
    Debug.Print "Test 1"
    Debug.Print "Pattern: ""Holdings Old*"""
    PrintArray GetFileNames(ThePath, "Holdings Old*")
    Debug.Print
    
    Debug.Print "Test 1"
    Debug.Print "Pattern: ""Holdings-Equity-Equity-????????.*"""
    PrintArray GetFileNames(ThePath, "Holdings-Equity-Equity-????????.*")
    Debug.Print

    Debug.Print "Test 1"
    Debug.Print "Pattern: ""Holdings-Alternatives-Alternatives-????????.*"""
    PrintArray GetFileNames(ThePath, "Holdings-Alternatives-Alternatives-????????.*")
    Debug.Print
End Sub

Public Sub TestSlicingArrays()
    Dim array1 As Variant
    Dim array2 As Variant
    
    Let array1 = Array(1, 2, 3, 4, 5)
    Debug.Print "Let array1 = Array(1,2,3,4,5)"
    Debug.Print "Slicing using Application.Index(array1, array(2,4))."
    PrintArray Application.Index(array1, Array(2, 4))
    Debug.Print
    
    Let array2 = [{1,2,3; 4,5,6; 7,8,9; 3,11,12}]
    Debug.Print "Let array2 = [{1,2,3; 4,5,6; 7,8,9; 10,11,12}]"
    Debug.Print "The array looks like:"
    PrintArray array2
    Debug.Print "Slicing using Application.Index(array2, 0, 2)."
    PrintArray Application.Index(array2, 0, 2)
    Debug.Print "Slicing using Application.Index(array2, 2, 0)."
    PrintArray Application.Index(array2, 2, 0)
    Debug.Print "Slicing using Application.Index(array2, Array(1,3), 1)."
    PrintArray Application.Index(array2, Array(1, 3), 1)
    Debug.Print "Slicing using Application.Index(array2, Array(1,3), 2)."
    PrintArray Application.Index(array2, Array(1, 3), 2)
    Debug.Print "Slicing using Application.Index(array2, Array(1, 3), Array(2,3))."
    PrintArray Application.Index(array2, Array(1, 3), Array(2, 3))
    
    Debug.Print
    Debug.Print "Testing assigning entire columns and rows to an array."
    Let array1 = [{1,2,3; 4,5,6; 7,8,9; 3,11,12}]
    Debug.Print "Before injecting new values in column 1"
    PrintArray array1
    Let Application.Index(array1, 0, 1) = [{10; 40; 70;30}]
    Debug.Print "After injecting new values in columns 1"
    PrintArray array1
    
    Call DumpInSheet(array1, TempComputation.Range("A1"))
    Let Application.Index(TempComputation.Range("A1").CurrentRegion, 0, 1) = [{10; 40; 70;30}]
End Sub

Public Sub TestConcatenateArrays()
    Dim a As Variant
    Dim b As Variant
    Dim c As Variant
    
    Let a = [{1,2,3; 4,5,6}]
    Let b = [{7;8}]
    Let c = ConcatenateArrays(a, b)
    Debug.Print "a is:"
    PrintArray a
    Debug.Print
    Debug.Print "b is:"
    PrintArray b
    Debug.Print
    If EmptyArrayQ(c) Then
        Debug.Print "a and b have incompatible dimensions."
    Else
        Debug.Print "The concatenation is:"
    
        PrintArray c
    End If
    Debug.Print "--------------------------" & vbCrLf
    
    Let a = [{1,2,3; 4,5,6}]
    Let b = [{7;8;9}]
    Let c = ConcatenateArrays(a, b)
    Debug.Print "a is:"
    PrintArray a
    Debug.Print
    Debug.Print "b is:"
    PrintArray b
    Debug.Print
    If EmptyArrayQ(c) Then
        Debug.Print "a and b have incompatible dimensions."
    Else
        Debug.Print "The concatenation is:"
    
        PrintArray c
    End If
    Debug.Print "--------------------------" & vbCrLf

    Let a = [{1,2,3; 4,5,6}]
    Let b = 23
    Let c = ConcatenateArrays(a, b)
    Debug.Print "a is:"
    PrintArray a
    Debug.Print
    Debug.Print "b is:"
    PrintArray b
    Debug.Print
    If EmptyArrayQ(c) Then
        Debug.Print "a and b have incompatible dimensions."
    Else
        Debug.Print "The concatenation is:"
    
        PrintArray c
    End If
    Debug.Print "--------------------------" & vbCrLf
    
    Let a = [{1,"02",3; 4,"05",6}]
    Let b = [{"07";"08"}]
    Let c = ConcatenateArrays(a, b)
    Debug.Print "a is:"
    PrintArray a
    Debug.Print
    Debug.Print "b is:"
    PrintArray b
    Debug.Print
    If EmptyArrayQ(c) Then
        Debug.Print "a and b have incompatible dimensions."
    Else
        Debug.Print "The concatenation is:"
    
        PrintArray c
    End If
    Call ToTemp(c)
    Debug.Print "--------------------------" & vbCrLf
    
    Let a = Array(1, 2, 3)
    Let b = Array(4, 5, 6)
    Let c = ConcatenateArrays(a, b)
    Debug.Print "a is:"
    PrintArray a
    Debug.Print
    Debug.Print "b is:"
    PrintArray b
    Debug.Print
    If EmptyArrayQ(c) Then
        Debug.Print "a and b have incompatible dimensions."
    Else
        Debug.Print "The concatenation is:"
    
        PrintArray c
    End If
    Call ToTemp(c)
    Debug.Print "--------------------------" & vbCrLf
    
    Let a = Array()
    Let b = Array(4, 5, 6)
    Let c = ConcatenateArrays(a, b)
    Debug.Print "a is:"
    PrintArray a
    Debug.Print
    Debug.Print "b is:"
    PrintArray b
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

Public Sub TestPrepend()
    Dim a As Variant
    Dim b As Variant
    
    Debug.Print "Testing Prepend(Array(1,2,3), 4)"
    PrintArray Prepend(Array(1, 2, 3), 4)
    Debug.Print
    
    Let a = Prepend(Array(1, 2, 3), Array(1, 4))
    Debug.Print "Testing Prepend(Array(1,2,3), array(1,4))"
    Debug.Print "Use watch on variable a"
    Debug.Print
    
    Let a = [{1,2,3; 4,5,6}]
    Let b = Array(7, 8, 9)
    Debug.Print "Testing Prepend(a, b) on a = [{1,2,3; 4,5,6}] and b = Array(7, 8, 9)"
    PrintArray Prepend(a, b)
    Debug.Print

    Let a = Array(7, 8, 9)
    Let b = [{1,2,3; 4,5,6}]
    Debug.Print "Testing Prepend(a, b) on a = Array(7, 8, 9) and b = [{1,2,3; 4,5,6}]"
    Debug.Print "Use a watch to see the output. Cannot be printed."
    Call Prepend(a, b)
    Debug.Print
    
    Let a = [{7,8,9; 10,11,12}]
    Let b = [{1,2,3; 4,5,6}]
    Debug.Print "Testing Prepend(a, b) on a = [{7,8,9; 10,11,12}] and b = [{1,2,3; 4,5,6}]"
    PrintArray Prepend(a, b)
End Sub

Public Sub TestAppend()
    Dim a As Variant
    Dim b As Variant
    
    Debug.Print "Testing Append(Array(1,2,3), 4)"
    PrintArray Append(Array(1, 2, 3), 4)
    Debug.Print
    
    Let a = Append(Array(1, 2, 3), Array(1, 4))
    Debug.Print "Testing Append(Array(1,2,3), array(1,4))"
    Debug.Print "Use watch on variable a"
    Debug.Print
    
    Let a = [{1,2,3; 4,5,6}]
    Let b = Array(7, 8, 9)
    Debug.Print "Testing Append(a, b) on a = [{1,2,3; 4,5,6}] and b = Array(7, 8, 9)"
    PrintArray Append(a, b)
    Debug.Print

    
    Let a = [{7,8,9; 10,11,12}]
    Let b = [{1,2,3; 4,5,6}]
    Debug.Print "Testing Append(a, b) on a = [{7,8,9; 10,11,12}] and b = [{1,2,3; 4,5,6}]"
    PrintArray Append(a, b)
End Sub

Public Sub TestCreateSequentialArray()
    Debug.Print "Sequential array 1...10"
    PrintArray CreateSequentialArray(1, 10)
    Debug.Print
    
    Debug.Print "Sequential array 2...6"
    PrintArray CreateSequentialArray(2, 5)
    Debug.Print
    
    Debug.Print "Sequential array 2...12 step 2"
    PrintArray CreateSequentialArray(2, 6, 2)
    Debug.Print

    Debug.Print "Sequential array 2...2 repeated 10 times at step 0"
    PrintArray CreateSequentialArray(2, 10, 0)
    Debug.Print
End Sub

Public Sub TestDumpInSheet()
    Dim AnArray As Variant
    
    Let AnArray = [{1,"02", 3, "04", 5; 10, "020", 30, "040", 50}]
    
    Call TempComputation.UsedRange.ClearContents
    Call TempComputation.UsedRange.ClearContents
    
    Call DumpInSheet(AnArray, TempComputation.Range("A1"), PreserveColumnTextFormats:=True)
    Call DumpInSheet(AnArray, TempComputation.Range("G1"))
End Sub

Public Sub TestSetRangeObjectToNothing()
    Dim ARange As Range
    
    Set ARange = Nothing
    Debug.Print "aRange is Nothing is " & (ARange Is Nothing)

    Set ARange = Range("A1")
    Debug.Print "aRange is Nothing is " & (ARange Is Nothing)
End Sub

Public Sub TestConnectAndSelect()
    PrintArray ConnectAndSelect("SELECT * FROM `ipreo`.`primarycompanytypemap`;", "ipreo", "PC-RiteshOLD", "Pablo", "whitetricornio")
End Sub

Public Sub Test2ConnectAndSelect()
    PrintArray ConnectAndSelect("SELECT * FROM `rndsecuritydescriptivedb`.`gicstopiamsectormapping`;", "rndsecuritydescriptivedb", "srv-rnd-sbx", "superuser", "tricornio")
End Sub

Public Sub Test3ConnectAndSelect()
    Dim TheResults As Variant
    
    Let TheResults = ConnectAndSelect("SELECT * FROM `piammaindb`.`employeedata`;", "piammaindb", "localhost", "root", "")
    PrintArray TheResults
    
    Debug.Print
    Debug.Print "Now, returning the first column with no header."
    
    PrintArray Rest(ConvertTo1DArray(GetColumn(TheResults, 1)))
    
    Debug.Print
    Debug.Print "The largest number is " & Application.Max(Rest(ConvertTo1DArray(GetColumn(TheResults, 1))))
End Sub

Public Sub Test4ConnectAndSelect()
    Dim TheResults As Variant
    
    Let TheResults = ConnectAndSelect("SELECT * FROM `piammaindb`.`employeedata`;", "piammaindb", "localhost", "root", "")
    PrintArray TheResults
    
    Debug.Print
    Debug.Print "Printing all but header row"
    
    PrintArray GetSubMatrix(TheResults, 2, GetNumberOfRows(TheResults), 1, GetNumberOfColumns(TheResults))
End Sub

Public Sub Test5ConnectAndSelect()
    Dim TheResults As Variant
    
    Let TheResults = ConnectAndSelect("SELECT * FROM `piammaindb`.`employeedata` WHERE `employmentstatuscode` = 10;", "piammaindb", "localhost", "root", "")
    PrintArray TheResults
    
    Debug.Print "The length of the array is " & GetNumberOfRows(TheResults) - 1
End Sub

Public Sub TestExecuteQuery()

    Dim cnt As ADODB.Connection

    'Instantiate the ADO-objects.
    Set cnt = New ADODB.Connection

    ' Open the database connection
    cnt.Open "Driver={MySQL ODBC 5.2 ANSI Driver};" & _
            "Server=PC-RiteshOLD;" & _
            "Database=ipreo;" & _
            "Uid=Pablo;" & _
            "Pwd=whitetricornio;"
    
    PrintArray ExecuteQuery(cnt, "SELECT * FROM `ipreo`.`primarycompanytypemap`;")
    
    'Release objects from the memory.
    Call cnt.Close
    Set cnt = Nothing
End Sub

Public Sub TestConnectToMsSqlAndExecuteSelectQuery()
    PrintArray ConnectToMsSqlAndExecuteSelectQuery("SELECT * FROM [Trade];", "TradarBE", "srv-db02\instance01")
End Sub

Public Sub TestExecuteMsSqlQuery()

    Dim cnt As ADODB.Connection

    'Instantiate the ADO-objects.
    Set cnt = New ADODB.Connection

    ' Open the database connection
    cnt.Open "Driver={SQL Server};" & _
            "Server=srv-db02\instance01;" & _
            "Database=TradarBE;"
    
    PrintArray ExecuteQuery(cnt, "SELECT TOP 10 * FROM [Trade];")
    
    'Release objects from the memory.
    Call cnt.Close
    Set cnt = Nothing
End Sub

Public Sub TestConnectAndExecuteInsertQuery()
    Dim DataArray As Variant
    Dim FieldNames As Variant
    Dim TableName As String
    Dim DatabaseName As String
    
    Let TableName = "table1"
    Let DatabaseName = "testdb"
    Let DataArray = DoubleQuote2DArray([{"'name1'", "'id1'", 1; "'name3'", "'id3'", Empty; "'name4'", "'id4'", 40}])
    Let FieldNames = AddSingleBackQuotesToAllArrayElements(Array("`name`", "`id`", "`units`"))

    Call ConnectAndExecuteInsertQuery(ValuesMatrix:=DataArray, _
                                      FieldNames:=FieldNames, _
                                      TableName:=TableName, _
                                      ServerAddress:=DbServerAddress, _
                                      DatabaseName:=DatabaseName, _
                                      UserName:=DbUserName, _
                                      ThePassword:=DbPassword)
End Sub

Public Sub TestTargetAssetAllocations()
    Dim aTargetAssetAllocations As TargetAssetAllocations
    Dim TheFileName As String
    
    ' Set the input file's name
    Let TheFileName = "X:\WorkProduct\SourceCode\VBA\TestingFiles\ExcelLibraryV2.0-NoFactSetAddIn\AssetAllocations-20140227.xlsb"
    
    ' Instantiate the HoldingsFromAa object
    Set aTargetAssetAllocations = NewTargetAssetAllocations
    
    ' Initialize HoldingsFromAa instance
    Call aTargetAssetAllocations.InitializeWithFile(TheFileName)
    
    ' Inject the instance into the DB
    Call aTargetAssetAllocations.InjectIntoMySql("localhost", "etwip2dot0", "targetassetallocations", "root", "")
End Sub

' Testing what a dictionary returns when it does not find something.
Public Sub TestDictionary()
    Dim aDict As New Dictionary
        
    Call aDict.Add(Key:=1, Item:=10)
    Debug.Print "We found 1: " & aDict.Item(Key:=1) & "."
    If aDict.Item(Key:=2) = Empty Then
        Debug.Print "We did not find 2."
    Else
        Debug.Print "We found 2."
    End If
End Sub

Public Sub TestPrivateFile()
    Dim pf As PrivateFile
    Dim aPfFileName As String
    Dim a1Dot4FileName As String
    
    Let Application.DisplayAlerts = False
    
    Let aPfFileName = "X:\PiamSoftwareInfrastructure\Interfaces\EquityModelImprovements\Sample-PrivateFile-20140321.xlsb"
    Let a1Dot4FileName = "X:\PiamSoftwareInfrastructure\Interfaces\EquityModelImprovements\EquityModel-1Dot4_Entire_Universe.xlsm"
    
    ' Instatiate private file class
    Set pf = New PrivateFile
    
    ' Initialize private file instance with the constents of an unclassified private file
    Call pf.InitializeWithPrivateFile(aPfFileName)
    
    ' Classify pf
    Call pf.ClassifyWithEquity1Dot4(a1Dot4FileName)
    
    ' Inject classified private file instance into DB
    Call pf.InjectIntoMySql(DbServerAddress, "equity250v1", "weeklyprivatefiles", DbUserName, DbPassword)
    
    Let Application.DisplayAlerts = True
End Sub

Public Sub TestPrivateFile2()
    Dim pf As PrivateFile
    Dim aPfFileName As String
    
    Let Application.DisplayAlerts = False
    
    Let aPfFileName = "X:\PiamSoftwareInfrastructure\Interfaces\EquityModelImprovements\Sample-PrivateFile-20140321.xlsb"
    
    ' Instatiate private file class
    Set pf = New PrivateFile
    
    ' Initialize private file instance with the constents of an unclassified private file
    Call pf.LoadAndProcessPrivateFile(aPfFileName)
    
    Let Application.DisplayAlerts = True
End Sub

Public Sub TestGetPallStudiedFactors()
    'Instantiate an object of type EquityDbQuerier
    Dim aEquityHandlerEngine As EquityDbHandler
    Set aEquityHandlerEngine = NewEquityDbHandler
    
    ' Define a holding object
    Dim TheResults As Variant

    ' Query Quintile Studied Factors
    Let TheResults = aEquityHandlerEngine.GetPallFactorsStudied(Quintile)
    PrintArray TheResults
    
    ' Query Decile Studied Factors
    Let TheResults = aEquityHandlerEngine.GetPallFactorsStudied(Decile)
    PrintArray TheResults
    
End Sub

Public Sub TestGetMimickingPortfolioReturns()
    'Instantiate an object of type EquityDbQuerier
    Dim aEquityHandlerEngine As EquityDbHandler
    Set aEquityHandlerEngine = NewEquityDbHandler

    'Define a holding object
    Dim TheResults As Variant

    'Query Mimicking Quintile Returns (LOW, HIGH, SPREAD)
    Let TheResults = aEquityHandlerEngine.GetTimeSeriesOfMimickingPortfolioReturns("PB_SER", "REG1", "EQW", str(20010101), str(20030101), Quintile, False)
    PrintArray TheResults
    
    'Query Mimicking Quintile Returns (LOW, HIGH, SPREAD) and revert the sorting
    Let TheResults = aEquityHandlerEngine.GetTimeSeriesOfMimickingPortfolioReturns("PB_SER", "REG1", "EQW", str(20010101), str(20030101), Quintile, True)
    PrintArray TheResults
    
    'Query Mimicking Decile Returns (LOW, HIGH, SPREAD)
    Let TheResults = aEquityHandlerEngine.GetTimeSeriesOfMimickingPortfolioReturns("PB_SER", "REG1", "EQW", str(20010101), str(20030101), Decile, False)
    PrintArray TheResults
End Sub

Public Sub TestGetMimimickingReturnsAndStats()
    'Instantiate an object of type EquityDbQuerier
    Dim aEquityHandlerEngine As EquityDbHandler
    Set aEquityHandlerEngine = NewEquityDbHandler

    'Define a holding object
    Dim TheResults As Variant

    'Get 3 things: 1.- Returns, 2.- cumReturns, 3.- histOfReturns
    Let TheResults = aEquityHandlerEngine.GetMimickingReturnsAndStats("PB_SER", "REG1", "EQW", 20010101, 20030101, 28, Quintile, False)
        
    'Print Results
    PrintArray TheResults(0)
    PrintArray TheResults(1)
    PrintArray TheResults(2)
End Sub


' Testing that the the private file objects behave correctly
' This program classifies the private file according to Ismael's requirements.
' It throws away all securities that are filtered out due to unacceptable PEs or
' missing country codes.  This is also the way that Ismael's requested private files
' to be indexed.  A bunch of these steps are computationally unnecessary, but if he wants
' green eggs and ham, he can have them since he is paying for breakfast.
Public Sub TestPrivateFileComputeRegionalStylesSizesMsciCodesAndWeights()
    Dim PalladyneAlphaToNumericSectorCodeMap As Dictionary
    Dim pf As PrivateFile
    Dim Weights As Variant
    Dim MarketCapWeightedCellularRanks As Variant
    Dim MarketCapWeightedSectoralRegionalRanks As Variant
    Dim MarketCapWeightedSizeStyleRanks As Variant
    Dim RegionalMap As Dictionary
    Dim SizeMap As Dictionary
    Dim StyleMap As Dictionary
    Dim SectoralMap As Dictionary
    Dim TheMsciSubIndustryMap As Dictionary
    Dim NumericalRegionalCodes() As Integer
    Dim NumericalSizeCodes() As Integer
    Dim NumericalStyleCodes() As Integer
    Dim NumericalSectoralCodes() As Integer
    Dim Counts As Variant
    Dim i As Long
    
    Let Application.ScreenUpdating = False
    Let Application.DisplayAlerts = False
    
    ' Instantiate and initialize PrivateFile variable
    Set pf = New PrivateFile
    Call pf.InitializeWithPrivateFile(ThisWorkbook.Path() & "\..\ET Wip Production Directory\ET Wip Input Directory\InputFromR&D\PrivateFiles\20140328.xls")

    ' Pre-allocate arrays for numerical code arrays
    ReDim NumericalRegionalCodes(pf.GetDbLength)
    ReDim NumericalSizeCodes(pf.GetDbLength)
    ReDim NumericalStyleCodes(pf.GetDbLength)
    ReDim NumericalSectoralCodes(pf.GetDbLength)
    ReDim MsciSubIndustryNames(pf.GetDbLength)
    
    ' Load translation mappings
    Set RegionalMap = RegionalNumericalMap()
    Set SizeMap = SizeNumericalMap()
    Set StyleMap = StyleNumericalMap()
    Set SectoralMap = SectoralNumericalMap()
    Set TheMsciSubIndustryMap = MsciSubIndustryMap()

    ' Filter out non-positive and extreme PEs.  Doing this step here deletes approximately 2,500 securities, which
    ' speeds up all subsequent steps.
    'Call pf.FilterOutUnacceptablePes
    
    ' Filter out securities with missing country codes.
    Call pf.FilterOutMissingCountries

    ' Load and initialize country-to-region mapping.
    Call pf.InitializeCountryToRegionMap(ThisWorkbook.Path() & "\Mappings\Country Region Relation.xlsb")
    
    ' Compute and load regional codes.  These must be done before initializing size and style threshold tables
    Call pf.ComputeRegionalCodes
    
    ' Compute threshold and size tables.  These must be done after filtering unacceptable PEs. Initialize GICS-to-MSCI mapping
    Call pf.InitializeSizeThresholds
    Call pf.InitializeStyleThresholds
    Call pf.InitializeGicsToMsciSubIndustryCodeMappingWithFile(ThisWorkbook.Path() & "\Mappings\GicsCodeToMsciSubIndustryCodeMap.xlsb")
    
    ' Compute and populate regions, sizes, and styles for all securities.
    Call pf.ComputesSizes
    Call pf.ComputesStyles
    Call pf.ComputeClassifications
    
    ' Compute and export the weights to the right of the underlying range
    ' Set headers for weights table
    Let Weights = pf.ComputeCellularWeights
    Let pf.GetRange.Worksheet.Range("AJ1").Resize(1, 2).Value2 = Array("Cell", "Weight")
    Let pf.GetRange.Worksheet.Range("AJ2").Resize(UBound(Weights, 1), 2).Value2 = Weights
    
    ' Compute and export market cap-weighted, cellular ranks
    Let MarketCapWeightedCellularRanks = pf.ComputeCellularMarketCapWeightedRanks
    Let pf.GetRange.Worksheet.Range("AM1").Resize(1, 2).Value2 = Array("Cell", "Market Cap-Weighted Rank")
    Let pf.GetRange.Worksheet.Range("AM2").Resize(UBound(MarketCapWeightedCellularRanks, 1), 2).Value2 = MarketCapWeightedCellularRanks
    
    ' Compute and export market cap-weight, sectoral/regional ranks
    Let MarketCapWeightedSectoralRegionalRanks = pf.ComputeSectoralRegionalMarketCapWeightedRanks
    Let pf.GetRange.Worksheet.Range("AP1").Resize(1, 2).Value2 = Array("Cell", "Market Cap-Weighted Rank")
    Let pf.GetRange.Worksheet.Range("AP2").Resize(UBound(MarketCapWeightedSectoralRegionalRanks, 1), 2).Value2 = MarketCapWeightedSectoralRegionalRanks
    
    ' Compute and export market cap-weight, size/style ranks
    Let MarketCapWeightedSizeStyleRanks = pf.ComputeSizeStyleMarketCapWeightedRanks
    Let pf.GetRange.Worksheet.Range("AS1").Resize(1, 2).Value2 = Array("Cell", "Size/Style Ranks")
    Let pf.GetRange.Worksheet.Range("AS2").Resize(UBound(MarketCapWeightedSizeStyleRanks, 1), 2).Value2 = MarketCapWeightedSizeStyleRanks

    ' Compute and export sectoral weight
    Let Weights = pf.ComputeSectoralWeights
    Let pf.GetRange.Worksheet.Range("AV1").Resize(1, 2).Value2 = Array("Cell", "Sectoral Weight")
    Let pf.GetRange.Worksheet.Range("AV2").Resize(UBound(Weights, 1), 2).Value2 = Weights
    
    ' Compute and export regional weight
    Let Weights = pf.ComputeRegionalWeights
    Let pf.GetRange.Worksheet.Range("AY1").Resize(1, 2).Value2 = Array("Cell", "Regional Weight")
    Let pf.GetRange.Worksheet.Range("AY2").Resize(UBound(Weights, 1), 2).Value2 = Weights
    
    ' Compute and export sectoral/regional weight
    Let Weights = pf.ComputeSectoralRegionalWeights
    Let pf.GetRange.Worksheet.Range("BB1").Resize(1, 2).Value2 = Array("Cell", "Sectoral/Regional Weight")
    Let pf.GetRange.Worksheet.Range("BB2").Resize(UBound(Weights, 1), 2).Value2 = Weights
    
    ' Compute and export size/style weight
    Let Weights = pf.ComputeSizeStyleWeights
    Let pf.GetRange.Worksheet.Range("BE1").Resize(1, 2).Value2 = Array("Cell", "Size/Style Weight")
    Let pf.GetRange.Worksheet.Range("BE2").Resize(UBound(Weights, 1), 2).Value2 = Weights
    
    ' Compute and export cellular counts
    Let Counts = pf.ComputeCellularSecurityCounts
    Let pf.GetRange.Worksheet.Range("BH1").Resize(1, 2).Value2 = Array("Cell", "Cellular Counts")
    Let pf.GetRange.Worksheet.Range("BH2").Resize(UBound(Counts, 1), 2).Value2 = Counts
    
    ' Compute and export regional counts
    Let Counts = pf.ComputeRegionalSecurityCounts
    Let pf.GetRange.Worksheet.Range("BK1").Resize(1, 2).Value2 = Array("Cell", "Regional Counts")
    Let pf.GetRange.Worksheet.Range("BK2").Resize(UBound(Counts, 1), 2).Value2 = Counts
    
    ' Compute and export Sector counts
    Let Counts = pf.ComputeSectoralSecurityCounts
    Let pf.GetRange.Worksheet.Range("BN1").Resize(1, 2).Value2 = Array("Cell", "Sectoral Counts")
    Let pf.GetRange.Worksheet.Range("BN2").Resize(UBound(Counts, 1), 2).Value2 = Counts
    
    ' Compute and export the portfolios, market cap-weighted rank
    Let pf.GetRange.Worksheet.Range("BQ1").Resize(1, 2).Value2 = Array("Portfolio's rank", pf.ComputePortfolioMarketCapWeightedRank)
    
    Call pf.GetRange.Worksheet.Range("AJ:BR").EntireColumn.AutoFit

    Let Application.ScreenUpdating = False
    Let Application.DisplayAlerts = False
End Sub

Public Sub TestLoadAndProcessPrivateFile()
    Dim PalladyneAlphaToNumericSectorCodeMap As Dictionary
    Dim pf As PrivateFile
    Dim Weights As Variant
    Dim MarketCapWeightedCellularRanks As Variant
    Dim MarketCapWeightedSectoralRegionalRanks As Variant
    Dim MarketCapWeightedSizeStyleRanks As Variant
    Dim RegionalMap As Dictionary
    Dim SizeMap As Dictionary
    Dim StyleMap As Dictionary
    Dim SectoralMap As Dictionary
    Dim TheMsciSubIndustryMap As Dictionary
    Dim NumericalRegionalCodes() As Integer
    Dim NumericalSizeCodes() As Integer
    Dim NumericalStyleCodes() As Integer
    Dim NumericalSectoralCodes() As Integer
    Dim Counts As Variant
    Dim i As Long
    
    Let Application.ScreenUpdating = False
    Let Application.DisplayAlerts = False
    
    ' Instantiate and initialize PrivateFile variable
    Set pf = New PrivateFile
    Call pf.LoadAndProcessPrivateFile(PrivateFileFileName:=ThisWorkbook.Path() & "\..\ET Wip Production Directory\ET Wip Input Directory\InputFromR&D\PrivateFiles\20140328.xls", _
                                      FilteredOutUnacceptablePesAndCountries:=False)
                                      
    ' Compute and export the weights to the right of the underlying range
    ' Set headers for weights table
    Let Weights = pf.ComputeCellularWeights
    Let pf.GetRange.Worksheet.Range("AJ1").Resize(1, 2).Value2 = Array("Cell", "Weight")
    Let pf.GetRange.Worksheet.Range("AJ2").Resize(UBound(Weights, 1), 2).Value2 = Weights
    
    ' Compute and export market cap-weighted, cellular ranks
    Let MarketCapWeightedCellularRanks = pf.ComputeCellularMarketCapWeightedRanks
    Let pf.GetRange.Worksheet.Range("AM1").Resize(1, 2).Value2 = Array("Cell", "Market Cap-Weighted Rank")
    Let pf.GetRange.Worksheet.Range("AM2").Resize(UBound(MarketCapWeightedCellularRanks, 1), 2).Value2 = MarketCapWeightedCellularRanks
    
    ' Compute and export market cap-weight, sectoral/regional ranks
    Let MarketCapWeightedSectoralRegionalRanks = pf.ComputeSectoralRegionalMarketCapWeightedRanks
    Let pf.GetRange.Worksheet.Range("AP1").Resize(1, 2).Value2 = Array("Cell", "Market Cap-Weighted Rank")
    Let pf.GetRange.Worksheet.Range("AP2").Resize(UBound(MarketCapWeightedSectoralRegionalRanks, 1), 2).Value2 = MarketCapWeightedSectoralRegionalRanks
    
    ' Compute and export market cap-weight, size/style ranks
    Let MarketCapWeightedSizeStyleRanks = pf.ComputeSizeStyleMarketCapWeightedRanks
    Let pf.GetRange.Worksheet.Range("AS1").Resize(1, 2).Value2 = Array("Cell", "Size/Style Ranks")
    Let pf.GetRange.Worksheet.Range("AS2").Resize(UBound(MarketCapWeightedSizeStyleRanks, 1), 2).Value2 = MarketCapWeightedSizeStyleRanks

    ' Compute and export sectoral weight
    Let Weights = pf.ComputeSectoralWeights
    Let pf.GetRange.Worksheet.Range("AV1").Resize(1, 2).Value2 = Array("Cell", "Sectoral Weight")
    Let pf.GetRange.Worksheet.Range("AV2").Resize(UBound(Weights, 1), 2).Value2 = Weights
    
    ' Compute and export regional weight
    Let Weights = pf.ComputeRegionalWeights
    Let pf.GetRange.Worksheet.Range("AY1").Resize(1, 2).Value2 = Array("Cell", "Regional Weight")
    Let pf.GetRange.Worksheet.Range("AY2").Resize(UBound(Weights, 1), 2).Value2 = Weights
    
    ' Compute and export sectoral/regional weight
    Let Weights = pf.ComputeSectoralRegionalWeights
    Let pf.GetRange.Worksheet.Range("BB1").Resize(1, 2).Value2 = Array("Cell", "Sectoral/Regional Weight")
    Let pf.GetRange.Worksheet.Range("BB2").Resize(UBound(Weights, 1), 2).Value2 = Weights
    
    ' Compute and export size/style weight
    Let Weights` = pf.ComputeSizeStyleWeights
    Let pf.GetRange.Worksheet.Range("BE1").Resize(1, 2).Value2 = Array("Cell", "Size/Style Weight")
    Let pf.GetRange.Worksheet.Range("BE2").Resize(UBound(Weights, 1), 2).Value2 = Weights
    
    ' Compute and export cellular counts
    Let Counts = pf.ComputeCellularSecurityCounts
    Let pf.GetRange.Worksheet.Range("BH1").Resize(1, 2).Value2 = Array("Cell", "Cellular Counts")
    Let pf.GetRange.Worksheet.Range("BH2").Resize(UBound(Counts, 1), 2).Value2 = Counts
    
    ' Compute and export regional counts
    Let Counts = pf.ComputeRegionalSecurityCounts
    Let pf.GetRange.Worksheet.Range("BK1").Resize(1, 2).Value2 = Array("Cell", "Regional Counts")
    Let pf.GetRange.Worksheet.Range("BK2").Resize(UBound(Counts, 1), 2).Value2 = Counts
    
    ' Compute and export Sector counts
    Let Counts = pf.ComputeSectoralSecurityCounts
    Let pf.GetRange.Worksheet.Range("BN1").Resize(1, 2).Value2 = Array("Cell", "Sectoral Counts")
    Let pf.GetRange.Worksheet.Range("BN2").Resize(UBound(Counts, 1), 2).Value2 = Counts
    
    ' Compute and export the portfolios, market cap-weighted rank
    Let pf.GetRange.Worksheet.Range("BQ1").Resize(1, 2).Value2 = Array("Portfolio's rank", pf.ComputePortfolioMarketCapWeightedRank)
    
    Call pf.GetRange.Worksheet.Range("A:BR").EntireColumn.AutoFit
    
    ' Inject into DB
    Call pf.InjectIntoMySql(DbServerAddress, "equity250v1", "weeklyprivatefiles", DbUserName, DbPassword)

    Let Application.ScreenUpdating = False
    Let Application.DisplayAlerts = False
End Sub

Public Sub TestUnionOfSets()
    Dim a As Variant
    Dim b As Variant
    
    Let a = Array(1, 2, 3, 4, 5)
    Let b = Array(3, 5)
    
    Debug.Print "With A = Array(1, 2, 3, 4, 5) and B = Array(3, 5)"
    PrintArray UnionOfSets(a, b)
    
    Let a = Array(1, 2, 3, 4, 5)
    Let b = Array()
    
    Debug.Print "With A = Array(1, 2, 3, 4, 5) and B = Array()"
    PrintArray UnionOfSets(a, b)

    Let a = Array()
    Let b = Array(1)
    
    Debug.Print "With A = Array() and B = Array(1)"
    PrintArray UnionOfSets(a, b)
    
    Let a = Array()
    Let b = Array()
    
    Debug.Print "With A = Array() and B = Array()"
    PrintArray UnionOfSets(a, b)
    
    Let a = Array(1, 2, 2)
    Let b = Array(2)
    
    Debug.Print "With A = Array(1, 2, 2) and B = Array(2)"
    PrintArray UnionOfSets(a, b)
    
    Let a = Array(1, 4, 2, 2)
    Let b = Array(2, 2, 3)
    
    Debug.Print "With A = Array(1, 4, 2, 2) and B = Array(2, 2, 3)"
    PrintArray UnionOfSets(a, b)
End Sub

Public Sub TestComplementOfSets()
    Dim a As Variant
    Dim b As Variant
    
    Let a = Array(1, 2, 3, 4, 5)
    Let b = Array(3, 5)
    
    Debug.Print "With A = Array(1, 2, 3, 4, 5) and B = Array(3, 5)"
    PrintArray ComplementOfSets(a, b)
    
    Let a = Array(1, 2, 3, 4, 5)
    Let b = Array()
    
    Debug.Print "With A = Array(1, 2, 3, 4, 5) and B = Array()"
    PrintArray ComplementOfSets(a, b)

    Let a = Array()
    Let b = Array(1)
    
    Debug.Print "With A = Array() and B = Array(1)"
    PrintArray ComplementOfSets(a, b)
    
    Let a = Array()
    Let b = Array()
    
    Debug.Print "With A = Array() and B = Array()"
    PrintArray ComplementOfSets(a, b)
    
    Let a = Array(1, 2, 2)
    Let b = Array(2)
    
    Debug.Print "With A = Array(1, 2, 2) and B = Array(2)"
    PrintArray ComplementOfSets(a, b)
    
    Let a = Array(1, 4, 2, 2)
    Let b = Array(2, 2, 3)
    
    Debug.Print "With A = Array(1, 4, 2, 2) and B = Array(2, 2, 3)"
    PrintArray ComplementOfSets(a, b)
End Sub

Public Sub TestOptimalPortfolio()
    Dim anOptimalPortfolio As OptimalPortfolio
    Dim TheFileName As String

    ' Set the input file's name
    Let TheFileName = "X:\TestProductionEnvironment\ET Wip Production Directory\ET Wip Input Directory\InputFromR&D\FinalOptimalPortfolios\New Optimal-Equity-Equity-20140318.xlsx"

    ' Instantiate the HoldingsFromAa object
    Set anOptimalPortfolio = NewOptimalPortfolio

    ' Initialize OptimalPortfolio instance
    Call anOptimalPortfolio.InitializeWithFile(TheFileName)
    
    ' Inject portfolio into database
    Call anOptimalPortfolio.InjectIntoMySql("localhost", "etwip2dot1", "finaloptimalportfolios", "root", "")
    
    ' Initialize with database
    Call anOptimalPortfolio.InitializeWithMySql("localhost", "etwip2dot1", "finaloptimalportfolios", "root", "", _
                                                "Equity", "Equity", 20140227)
    PrintArray anOptimalPortfolio.GetHeaders
    PrintArray anOptimalPortfolio.GetRange.Value2
End Sub

Public Sub Test1ConsolidateWorksheets()
    Dim WshtArray(5) As Worksheet
    Dim i As Integer
    Dim NumberOfRows As Long
    Dim Headers As Variant
    Dim RandomArray As Variant
    Dim TheConsolidatedResults As Variant
    Dim ResultsWsht As Worksheet
    
    Let Application.DisplayAlerts = False
    
    ' Set the headers row
    Let Headers = Array("Col1", "Col2")
    
    ' Fill each of the newly instantiated worksheets with a small headers row and random data
    For i = 1 To 5
        ' Instantiate test worksheet
        Set WshtArray(i) = ThisWorkbook.Worksheets.Add
        
        ' Add headers row
        Call DumpInTempPositionWithoutFirstClearing(Headers, WshtArray(i).Range("A1"))
        
        ' Choose a random number of rows between 1 and 100
        Let NumberOfRows = CInt(Application.WorksheetFunction.RandBetween(1, 100))
        
        ' Instantiate a random 2D array of doubles
        Call DumpInTempPositionWithoutFirstClearing(RandomMatrix(NumberOfRows, 2), WshtArray(i).Range("A2"))
    Next i
    
    ' Consolidate the worksheets
    Let TheConsolidatedResults = ConsolidateWorksheets(WshtArray, 2)
    
    ' Instantiate worksheet to hold results
    Set ResultsWsht = ThisWorkbook.Worksheets.Add
    Let ResultsWsht.Name = "ConsolidationWorksheet"

    ' Dump consolidated results in the consolidation worksheet
    Call DumpInTempPositionWithoutFirstClearing(TheConsolidatedResults, ResultsWsht.Range("A1"))
    
    Let Application.DisplayAlerts = True
End Sub

Public Sub Test2ConsolidateWorksheets()
    Dim WshtArray(5) As Worksheet
    Dim i As Integer
    Dim NumberOfRows As Long
    Dim Headers As Variant
    Dim RandomArray As Variant
    Dim TheConsolidatedResults As Variant
    Dim ResultsWsht As Worksheet
    
    Let Application.DisplayAlerts = False
    
    ' Set the headers row
    Let Headers = [{"Col1", "Col2"; "Col3", "Col4"}]
    
    ' Fill each of the newly instantiated worksheets with a small headers row and random data
    For i = 1 To 5
        ' Instantiate test worksheet
        Set WshtArray(i) = ThisWorkbook.Worksheets.Add
        
        ' Add headers row
        Call DumpInTempPositionWithoutFirstClearing(Headers, WshtArray(i).Range("A1"))
        
        ' Choose a random number of rows between 1 and 100
        Let NumberOfRows = CInt(Application.WorksheetFunction.RandBetween(1, 100))
        
        ' Instantiate a random 2D array of doubles
        Call DumpInTempPositionWithoutFirstClearing(RandomMatrix(NumberOfRows, 2), WshtArray(i).Range("A3"))
    Next i
    
    ' Consolidate the worksheets
    Let TheConsolidatedResults = ConsolidateWorksheets(WshtArray, 3)
    
    ' Instantiate worksheet to hold results
    Set ResultsWsht = ThisWorkbook.Worksheets.Add
    Let ResultsWsht.Name = "ConsolidationWorksheet"

    ' Dump consolidated results in the consolidation worksheet
    Call DumpInTempPositionWithoutFirstClearing(TheConsolidatedResults, ResultsWsht.Range("A1"))
    
    Let Application.DisplayAlerts = True
End Sub

Public Sub Test3ConsolidateWorksheets()
    Dim WshtArray(5) As Worksheet
    Dim i As Integer
    Dim NumberOfRows As Long
    Dim RandomArray As Variant
    Dim TheConsolidatedResults As Variant
    Dim ResultsWsht As Worksheet
    
    Let Application.DisplayAlerts = False
    
    ' Fill each of the newly instantiated worksheets with a small headers row and random data
    For i = 1 To 5
        ' Instantiate test worksheet
        Set WshtArray(i) = ThisWorkbook.Worksheets.Add
        
        ' Choose a random number of rows between 1 and 100
        Let NumberOfRows = CInt(Application.WorksheetFunction.RandBetween(1, 100))
        
        ' Instantiate a random 2D array of doubles
        Call DumpInTempPositionWithoutFirstClearing(RandomMatrix(NumberOfRows, 2), WshtArray(i).Range("A1"))
    Next i
    
    ' Consolidate the worksheets
    Let TheConsolidatedResults = ConsolidateWorksheets(WshtArray)
    
    ' Instantiate worksheet to hold results
    Set ResultsWsht = ThisWorkbook.Worksheets.Add
    Let ResultsWsht.Name = "ConsolidationWorksheet"

    ' Dump consolidated results in the consolidation worksheet
    Call DumpInTempPositionWithoutFirstClearing(TheConsolidatedResults, ResultsWsht.Range("A1"))
    
    Let Application.DisplayAlerts = True
End Sub

Public Sub TestConsolidateWorkbooks()
    Dim WbkArray(5) As Workbook
    Dim i As Integer
    Dim NumberOfRows As Long
    Dim RandomArray As Variant
    Dim TheConsolidatedResults As Variant
    Dim ResultsWsht As Worksheet
    
    Let Application.DisplayAlerts = False
    
    ' Fill each of the newly instantiated worksheets with a small headers row and random data
    For i = 1 To 5
        ' Instantiate test worksheet
        Set WbkArray(i) = Application.Workbooks.Add
        Call RemoveAllOtherWorksheets(WbkArray(i).Worksheets(1))
        
        ' Choose a random number of rows between 1 and 100
        Let NumberOfRows = CInt(Application.WorksheetFunction.RandBetween(1, 100))
        
        ' Instantiate a random 2D array of doubles
        Call DumpInTempPositionWithoutFirstClearing(RandomMatrix(NumberOfRows, 2), WbkArray(i).Worksheets(1).Range("A1"))
    Next i
    
    ' Consolidate the worksheets
    Let TheConsolidatedResults = ConsolidateWorkbooks(WbkArray)
    
    ' Instantiate worksheet to hold results
    Set ResultsWsht = ThisWorkbook.Worksheets.Add
    Let ResultsWsht.Name = "ConsolidationWorksheet"

    ' Dump consolidated results in the consolidation worksheet
    Call DumpInTempPositionWithoutFirstClearing(TheConsolidatedResults, ResultsWsht.Range("A1"))
    
    Let Application.DisplayAlerts = True
End Sub

Public Sub Test2ConsolidateWorkbooks()
    Dim Headers As Variant
    Dim WbkArray(5) As Workbook
    Dim i As Integer
    Dim NumberOfRows As Long
    Dim RandomArray As Variant
    Dim TheConsolidatedResults As Variant
    Dim ResultsWsht As Worksheet
    
    Let Application.DisplayAlerts = False
    
    ' Set the headers row
    Let Headers = Array("Col1", "Col2")
    
    ' Fill each of the newly instantiated worksheets with a small headers row and random data
    For i = 1 To 5
        ' Instantiate test worksheet
        Set WbkArray(i) = Application.Workbooks.Add
        Call RemoveAllOtherWorksheets(WbkArray(i).Worksheets(1))
        
        ' Add headers row
        Call DumpInTempPositionWithoutFirstClearing(Headers, WbkArray(i).Worksheets(1).Range("A1"))

        ' Choose a random number of rows between 1 and 100
        Let NumberOfRows = CInt(Application.WorksheetFunction.RandBetween(1, 100))
        
        ' Instantiate a random 2D array of doubles
        Call DumpInTempPositionWithoutFirstClearing(RandomMatrix(NumberOfRows, 2), WbkArray(i).Worksheets(1).Range("A2"))
    Next i
    
    ' Consolidate the worksheets
    Let TheConsolidatedResults = ConsolidateWorkbooks(WbkArray, StartingRow:=2)
    
    ' Instantiate worksheet to hold results
    Set ResultsWsht = ThisWorkbook.Worksheets.Add
    Let ResultsWsht.Name = "ConsolidationWorksheet"

    ' Dump consolidated results in the consolidation worksheet
    Call DumpInTempPositionWithoutFirstClearing(TheConsolidatedResults, ResultsWsht.Range("A1"))
    
    Let Application.DisplayAlerts = True
End Sub

Public Sub Test3ConsolidateWorkbooks()
    Dim Headers As Variant
    Dim WbkArray(5) As Workbook
    Dim i As Integer
    Dim NumberOfRows As Long
    Dim RandomArray As Variant
    Dim TheConsolidatedResults As Variant
    Dim ResultsWsht As Worksheet
    
    Let Application.DisplayAlerts = False
    
    ' Set the headers row
    Let Headers = [{"Col1", "Col2"; "Col3", "Col4"}]
    
    ' Fill each of the newly instantiated worksheets with a small headers row and random data
    For i = 1 To 5
        ' Instantiate test worksheet
        Set WbkArray(i) = Application.Workbooks.Add
        Call RemoveAllOtherWorksheets(WbkArray(i).Worksheets(1))
        
        ' Add headers
        Call DumpInTempPositionWithoutFirstClearing(Headers, WbkArray(i).Worksheets(1).Range("A1"))

        ' Choose a random number of rows between 1 and 100
        Let NumberOfRows = CInt(Application.WorksheetFunction.RandBetween(1, 100))
        
        ' Instantiate a random 2D array of doubles
        Call DumpInTempPositionWithoutFirstClearing(RandomMatrix(NumberOfRows, 2), WbkArray(i).Worksheets(1).Range("A3"))
    Next i
    
    ' Consolidate the worksheets
    Let TheConsolidatedResults = ConsolidateWorkbooks(WbkArray, StartingRow:=3)
    
    ' Instantiate worksheet to hold results
    Set ResultsWsht = ThisWorkbook.Worksheets.Add
    Let ResultsWsht.Name = "ConsolidationWorksheet"

    ' Dump consolidated results in the consolidation worksheet
    Call DumpInTempPositionWithoutFirstClearing(TheConsolidatedResults, ResultsWsht.Range("A1"))
    
    Let Application.DisplayAlerts = True
End Sub

Public Sub TestSerialDateAndSerialTime()
    Dim T As Long
    Dim d As Long
    
    Let d = ConvertDateToSerial(Date)
    Debug.Print "The date is " & d & "."
    Debug.Print "-" & GetYearFromSerialDate(d) & "-"
    Debug.Print "-" & GetMonthFromSerialDate(d) & "-"
    Debug.Print "-" & GetDayFromSerialDate(d) & "-" & vbCrLf
    
    Let T = ConvertTimeToSerial(Time)
    Debug.Print "The time is " & T & "."
    Debug.Print "-" & GetHourFromSerialTime(T) & "-"
    Debug.Print "-" & GetMinuteFromSerialTime(T) & "-"
    Debug.Print "-" & GetSecondFromSerialTime(T) & "-" & vbCrLf
    
    Let T = 12
    Debug.Print "The time is " & T & "."
    Debug.Print "-" & GetHourFromSerialTime(T) & "-"
    Debug.Print "-" & GetMinuteFromSerialTime(T) & "-"
    Debug.Print "-" & GetSecondFromSerialTime(T) & "-"
End Sub

Public Sub TestPostTradingPortfolio()
    Dim ThePath As String
    Dim aPostTradingPortfolio As PostTradingPortfolio
    Dim ConsolidatedPostTradingPortfolio As PostTradingPortfolio
    Dim AFileName As String
    
    Let Application.DisplayAlerts = False
    
    Let ThePath = "X:\TestProductionEnvironment\ET Wip Production Directory\ET Wip Output Directory\TradeLists\"
    
    Set aPostTradingPortfolio = NewPostTradingPortfolio
    
    ' Test if filenames are valid trade list filenames
    Let AFileName = "TradeList-100969727707-Equity-Equity-20140403-140410.xlsx"
    Debug.Print "Filename " & AFileName & " is valid: " & aPostTradingPortfolio.ValidTradeListFileNameQ(AFileName) & "." & vbCrLf
    
    Let AFileName = "TradeList-100969727707-Equity-Convertible-20140403-140410.xlsx"
    Debug.Print "Filename " & AFileName & " is valid: " & aPostTradingPortfolio.ValidTradeListFileNameQ(AFileName) & "." & vbCrLf

    Let AFileName = "TradeList-asdgg-Equity-Equity-20140403-140410.xlsx"
    Debug.Print "Filename " & AFileName & " is valid: " & aPostTradingPortfolio.ValidTradeListFileNameQ(AFileName) & "." & vbCrLf

    Let AFileName = "TradeLists-100969727707-Equity-Equity-20140403-140410.xlsx"
    Debug.Print "Filename " & AFileName & " is valid: " & aPostTradingPortfolio.ValidTradeListFileNameQ(AFileName) & "." & vbCrLf

    Let AFileName = "TradeLists-100969727707-Equity-Equity-20140403-1404100.xlsx"
    Debug.Print "Filename " & AFileName & " is valid: " & aPostTradingPortfolio.ValidTradeListFileNameQ(AFileName) & "." & vbCrLf

    Debug.Print "The list of files names matching Equity, Equity, and 20140403 are:"
    PrintArray aPostTradingPortfolio.GetFileList(ThePath, "Equity", "Equity", 20140403)
    
    ' Test initializing an instance with files
    Call aPostTradingPortfolio.InitializeWithFiles(aPostTradingPortfolio.GetFileList(ThePath, "Equity", "Equity", 20140410))
    Call aPostTradingPortfolio.GetRange.Worksheet.Activate
    
    ' Destroy the object and re-instantiate it to try initializing from the database
    Call aPostTradingPortfolio.InjectIntoMySql(DbServerAddress, "etwip2dot1", "posttradingportfolio", DbUserName, DbPassword)
    
    Set aPostTradingPortfolio = New PostTradingPortfolio
    Call aPostTradingPortfolio.InitializeWithMySql(DbServerAddress, "etwip2dot1", "posttradingportfolio", DbUserName, DbPassword, "Equity", "Equity", 20140410)

    Call MsgBox("About to test consolidation")
    
    Set ConsolidatedPostTradingPortfolio = aPostTradingPortfolio.GetConsolidatePostTradingPortfolio

    Let Application.DisplayAlerts = True
End Sub

' This function is used to test class MasterFile
Public Sub TestMasterFileClass()
    Dim mf As MasterFile
    Dim MasterFileName As String
    Dim ASedol As String
    Dim TheResult As Variant
    Dim msg As String
    
    ' Set the name of the file holding the master file
    Let MasterFileName = "X:\TestProductionEnvironment\Common Software Directory\Mappings\UniverseMasterFile.xlsb"
    
    ' Instantiate MasterFile and MasterFileRow classes
    Set mf = New MasterFile

    ' Initialize MasterInstance with the master file
    Call mf.InitializeWithFile(MasterFileName)
    
    ' Print some data on master file
    Debug.Print "There are " & mf.GetDbLength & " securities in the master file."
    
    Let ASedol = "B6X2H81"
    Debug.Print "Getting info on security with SEDOL " & ASedol
    
    Set TheResult = mf.GetMasterFileRowWith7DigitSedol(ASedol)
    Let msg = "Its ISIN is "
    If TheResult Is Nothing Then
        Let msg = msg & " NOT FOUND"
    Else
        Let msg = msg & TheResult.GetIsin.Value2
    End If
    Debug.Print msg & "."
    
    Set TheResult = mf.GetMasterFileRowWith7DigitSedol(ASedol)
    Let msg = "Its 6-digit SEDOL is "
    If TheResult Is Nothing Then
        Let msg = msg & " NOT FOUND"
    Else
        Let msg = msg & TheResult.Get6DigitSedol.Value2
    End If
    Debug.Print msg & "."

    Set TheResult = mf.GetMasterFileRowWith7DigitSedol(ASedol)
    Let msg = "Its Bloomberg ticker is "
    If TheResult Is Nothing Then
        Let msg = msg & " NOT FOUND"
    Else
        Let msg = msg & TheResult.GetBloombergTicker.Value2
    End If
    Debug.Print msg & "."
    
    Set TheResult = mf.GetMasterFileRowWith7DigitSedol(ASedol)
    Let msg = "Its EQY_FUND_TICKER is "
    If TheResult Is Nothing Then
        Let msg = msg & " NOT FOUND"
    Else
        Let msg = msg & TheResult.GetEqyFundTicker.Value2
    End If
    Debug.Print msg & "."
    
    Set TheResult = mf.GetMasterFileRowWith7DigitSedol(ASedol)
    Let msg = "The ISINs of the securities with the same EQY_FUND_TICKER are:"
    If TheResult Is Nothing Then
        Let msg = msg & " NOT FOUND"
    Else
        Let msg = msg & TheResult.GetBloombergTicker.Value2
    End If
    Debug.Print msg & "."
    PrintArray mf.GetAllIsinsWithGivenEqyFundTicker(mf.GetMasterFileRowWith7DigitSedol(ASedol).GetEqyFundTicker.Value2)
    
    Debug.Print "The 7Char SEDOLs of the securities with the same EQY_FUND_TICKER are:"
    PrintArray mf.GetAll7CharSedolsWithGivenEqyFundTicker(mf.GetMasterFileRowWith7DigitSedol(ASedol).GetEqyFundTicker.Value2)
    
    Debug.Print "The Bloomberg tickers of the securities with the same EQY_FUND_TICKER are:"
    PrintArray mf.GetAllBloombergTickersWithGivenEqyFundTicker(mf.GetMasterFileRowWith7DigitSedol(ASedol).GetEqyFundTicker.Value2)
    Debug.Print
    
    Let ASedol = "B7RL3L6"
    Debug.Print "Getting info on security with SEDOL " & ASedol
    
    Set TheResult = mf.GetMasterFileRowWith7DigitSedol(ASedol)
    Let msg = "Its ISIN is "
    If TheResult Is Nothing Then
        Let msg = msg & " NOT FOUND"
    Else
        Let msg = msg & TheResult.GetIsin.Value2
    End If
    Debug.Print msg & "."
    
    Set TheResult = mf.GetMasterFileRowWith7DigitSedol(ASedol)
    Let msg = "Its 6-digit SEDOL is "
    If TheResult Is Nothing Then
        Let msg = msg & " NOT FOUND"
    Else
        Let msg = msg & TheResult.Get6DigitSedol.Value2
    End If
    Debug.Print msg & "."

    Set TheResult = mf.GetMasterFileRowWith7DigitSedol(ASedol)
    Let msg = "Its Bloomberg ticker is "
    If TheResult Is Nothing Then
        Let msg = msg & " NOT FOUND"
    Else
        Let msg = msg & TheResult.GetBloombergTicker.Value2
    End If
    Debug.Print msg & "."
    
    Set TheResult = mf.GetMasterFileRowWith7DigitSedol(ASedol)
    Let msg = "Its EQY_FUND_TICKER is "
    If TheResult Is Nothing Then
        Let msg = msg & " NOT FOUND"
    Else
        Let msg = msg & TheResult.GetEqyFundTicker.Value2
    End If
    Debug.Print msg & "."
    
    Set TheResult = mf.GetMasterFileRowWith7DigitSedol(ASedol)
    Let msg = "The ISINs of the securities with the same EQY_FUND_TICKER are:"
    If TheResult Is Nothing Then
        Let msg = msg & " NOT FOUND"
    Else
        Let msg = msg & TheResult.GetBloombergTicker.Value2
    End If
    Debug.Print msg & "."
    PrintArray mf.GetAllIsinsWithGivenEqyFundTicker(mf.GetMasterFileRowWith7DigitSedol(ASedol).GetEqyFundTicker.Value2)
    
    Debug.Print "The 7Char SEDOLs of the securities with the same EQY_FUND_TICKER are:"
    PrintArray mf.GetAll7CharSedolsWithGivenEqyFundTicker(mf.GetMasterFileRowWith7DigitSedol(ASedol).GetEqyFundTicker.Value2)
    
    Debug.Print "The Bloomberg tickers of the securities with the same EQY_FUND_TICKER are:"
    PrintArray mf.GetAllBloombergTickersWithGivenEqyFundTicker(mf.GetMasterFileRowWith7DigitSedol(ASedol).GetEqyFundTicker.Value2)
    Debug.Print
    
    
    
    Let ASedol = "CSDS432"
    Debug.Print "Getting info on security with SEDOL " & ASedol
    
    Set TheResult = mf.GetMasterFileRowWith7DigitSedol(ASedol)
    Let msg = "Its ISIN is "
    If TheResult Is Nothing Then
        Let msg = msg & " NOT FOUND"
    Else
        Let msg = msg & TheResult.GetIsin.Value2
    End If
    Debug.Print msg & "."
    
    Set TheResult = mf.GetMasterFileRowWith7DigitSedol(ASedol)
    Let msg = "Its 6-digit SEDOL is "
    If TheResult Is Nothing Then
        Let msg = msg & " NOT FOUND"
    Else
        Let msg = msg & TheResult.Get6DigitSedol.Value2
    End If
    Debug.Print msg & "."

    Set TheResult = mf.GetMasterFileRowWith7DigitSedol(ASedol)
    Let msg = "Its Bloomberg ticker is "
    If TheResult Is Nothing Then
        Let msg = msg & " NOT FOUND"
    Else
        Let msg = msg & TheResult.GetBloombergTicker.Value2
    End If
    Debug.Print msg & "."
    
    Set TheResult = mf.GetMasterFileRowWith7DigitSedol(ASedol)
    Let msg = "Its EQY_FUND_TICKER is "
    If TheResult Is Nothing Then
        Let msg = msg & " NOT FOUND"
    Else
        Let msg = msg & TheResult.GetEqyFundTicker.Value2
    End If
    Debug.Print msg & "."
    
    Set TheResult = mf.GetMasterFileRowWith7DigitSedol(ASedol)
    Let msg = "The ISINs of the securities with the same EQY_FUND_TICKER are:"
    If TheResult Is Nothing Then
        Let msg = msg & " NOT FOUND"
    Else
        Let msg = msg & TheResult.GetBloombergTicker.Value2
    End If
    Debug.Print msg & "."
    
    If mf.GetMasterFileRowWith7DigitSedol(ASedol) Is Nothing Then
        Let TheResult = Array()
    Else
        Let TheResult = mf.GetAllIsinsWithGivenEqyFundTicker(mf.GetMasterFileRowWith7DigitSedol(ASedol).GetEqyFundTicker.Value2)
    End If
    
    If EmptyArrayQ(TheResult) Then
        Debug.Print "The ISINs of the securities with the same EQY_FUND_TICKER are:"
        Debug.Print "There are no securities with the given 7-char SEDOL."
    Else
        Debug.Print "The ISINs of the securities with the same EQY_FUND_TICKER are:"
        PrintArray TheResult
    End If
    
    If mf.GetMasterFileRowWith7DigitSedol(ASedol) Is Nothing Then
        Let TheResult = Array()
    Else
        Let TheResult = mf.GetAll7CharSedolsWithGivenEqyFundTicker(mf.GetMasterFileRowWith7DigitSedol(ASedol).GetEqyFundTicker.Value2)
    End If
    If EmptyArrayQ(TheResult) Then
        Debug.Print "The 7Char SEDOLs of the securities with the same EQY_FUND_TICKER are:"
        Debug.Print "There are no securities with the given 7-char SEDOL."
    Else
        Debug.Print "The 7Char SEDOLs of the securities with the same EQY_FUND_TICKER are:"
        PrintArray TheResult
    End If
    
        
    Call TempComputation.UsedRange.ClearContents
    Call TempComputation.UsedRange.ClearFormats
End Sub

' Legacy test cost
'Public Sub TestGetXmlElementIncludingTags()
'    Dim SomeXml As String
'
'    Let SomeXml = "<ALERT><TAG1 DSFSGSAG><TAG1>Payload2</TAG1>This is the payload</TAG1></ALERT>"
'    Debug.Print "Example 1"
'    Debug.Print "XML: " & SomeXml
'    Debug.Print "Tag = TAG1"
'    Debug.Print "With the tags: " & GetXmlElement(SomeXml, "TAG1", True)
'    Debug.Print "Without the tags: " & GetXmlElement(SomeXml, "TAG1", False)
'    Debug.Print
'
'    Let SomeXml = "<ALERT></ALERT>"
'    Debug.Print "Example 2"
'    Debug.Print "XML: " & SomeXml
'    Debug.Print "Tag = ALERT"
'    Debug.Print "With the tags: " & GetXmlElement(SomeXml, "ALERT", True)
'    Debug.Print "Without the tags: " & GetXmlElement(SomeXml, "ALERT", False)
'    Debug.Print
'
'    Let SomeXml = "<ALERT></ALERT>"
'    Debug.Print "Example 3"
'    Debug.Print "XML: " & SomeXml
'    Debug.Print "Tag = TAG1"
'    Debug.Print "With the tags: " & GetXmlElement(SomeXml, "TAG1", True)
'    Debug.Print "Without the tags: " & GetXmlElement(SomeXml, "TAG1", False)
'    Debug.Print
'
'    Let SomeXml = "<ALERT><TAG1 DSFSGSAG><TAG1>Payload2</TAG1>This is the payload</TAG1></ALERT>"
'    Debug.Print "Example 4: Getting the second TAG1"
'    Debug.Print "XML: " & SomeXml
'    Debug.Print "Tag = TAG1"
'    Debug.Print "With the tags: " & GetXmlElement(GetXmlElement(SomeXml, "TAG1", False), "TAG1", True)
'    Debug.Print "With the tags: " & GetXmlElement(GetXmlElement(SomeXml, "TAG1", False), "TAG1", False)
'End Sub


Public Sub TestCorporateAction()
    Dim AnAction As CorporateAction
    
    Set AnAction = New CorporateAction
    
    Let AnAction.SetType = Dividend
    Debug.Print "The type is " & AnAction.GetTypeAsString & "." & vbCrLf
    
    Let AnAction.SetType = StockSplit
    Debug.Print "The type is " & AnAction.GetTypeAsString & "." & vbCrLf

    Let AnAction.SetType = Sale
    Debug.Print "The type is " & AnAction.GetTypeAsString & "." & vbCrLf
End Sub

' This test processing Bloomberg's corporate action alerts, injecting them into the DB, and moving the source emails to the archive directory
Public Sub TestBloombergEquityAlerts1()
    Dim TheBloombergAlerts As BloombergEquityAlerts
        
    Let Application.DisplayAlerts = False
    
    Set TheBloombergAlerts = New BloombergEquityAlerts
    
    Call TheBloombergAlerts.InitializeWithOutlook
    
    Call TheBloombergAlerts.InjectIntoMySql

    Call TheBloombergAlerts.ArchiveAlerts

    Let Application.DisplayAlerts = True
End Sub

' This tests populating a worksheet with today's corporate action alerts
Public Sub TestBloombergEquityAlerts2()
    Dim TheBloombergAlerts As BloombergEquityAlerts
        
    Let Application.DisplayAlerts = False
    
    Set TheBloombergAlerts = New BloombergEquityAlerts
    Call TheBloombergAlerts.InitializeWithMySql(AnnouncementDate:=Date)

    Let Application.DisplayAlerts = True
End Sub

Public Sub TestBloombergEquityAlerts3()
    Dim TheBloombergAlerts As BloombergEquityAlerts
    Dim AlertSubset As BloombergEquityAlerts
        
    Let Application.DisplayAlerts = False
    
    Set TheBloombergAlerts = New BloombergEquityAlerts
    
    Call TheBloombergAlerts.InitializeWithOutlook
    
    Set AlertSubset = TheBloombergAlerts.GetAlertsForGivenBloombergTickerList(BloombergTickerArray:=Array("8130 JP", "9810 JP"), _
                                                                              StartDate:=#4/8/2014#, EndDate:=#12/31/2015#)
    
    Let Application.DisplayAlerts = True
End Sub

Public Sub TestBloombergEquityAlerts4()
    Dim TheBloombergAlerts As BloombergEquityAlerts
    Dim AlertSubset As New BloombergEquityAlerts
    
    Let Application.DisplayAlerts = False
    
    Set TheBloombergAlerts = New BloombergEquityAlerts
    
    Call TheBloombergAlerts.InitializeWithOutlook
    
    Call TheBloombergAlerts.InjectIntoMySql
    
    Call AlertSubset.InitializeWithMySql(BloombergTickerArray:=Array("8130 JP", "9810 JP"), _
                                         StartDate:=#4/8/2014#, EndDate:=#12/31/2015#)
End Sub

Public Sub TestGetLogFileContents()
    Dim aDict As Dictionary
    Dim ASeimAlertSet As New SeimRecordSet
    Dim TheResults As Variant
    Dim AFileName As String
    
    Let Application.DisplayAlerts = False
    
    Let AFileName = "X:\TestProductionEnvironment\SeimAudit\Inputs\DailyAuditLogs\audit1405270000-192.168.1.71.log"
    
    Set aDict = ASeimAlertSet.GetLogFileContents(AFileName)

    Debug.Print "We got " & aDict.Count & " items."
    Debug.Print
    Debug.Print "The 1st is:"
    Debug.Print aDict.Item(Key:=1)
    Debug.Print
    Debug.Print "The 2nd is:"
    Debug.Print aDict.Item(Key:=2)
    Debug.Print
    Debug.Print "The 3rd is:"
    Debug.Print aDict.Item(Key:=3)
    
    Call ASeimAlertSet.InitializedWithLogFile(AFileName)
    Call ASeimAlertSet.InjectIntoMySql

    Let Application.DisplayAlerts = True
End Sub

Public Sub TestSelectFromArrayWithFunction()
    Dim AnArray As Variant
    
    Let AnArray = Array(1, 2, 3, 4, -5, 6, -7, 8)
    
    PrintArray SelectFromArrayWithFunction(AnArray, "HelperForTestSelectFromArrayWithFunction")
End Sub

Private Function HelperForTestSelectFromArrayWithFunction(Arg As Integer) As Boolean
    If Arg <> Abs(Arg) Then
        Let HelperForTestSelectFromArrayWithFunction = False
    Else
        Let HelperForTestSelectFromArrayWithFunction = True
    End If
End Function

Public Sub TestSwapMatrixColumns()
    Dim a As Variant
    
    Let a = [{1,2,3; 4,5,6; 7,8,9}]
    Debug.Print "a is:"
    PrintArray a
    Debug.Print "Swaping columns 1 and 3. The new matrix is:"
    PrintArray SwapMatrixColumns(a, 1, 3)
    
    Debug.Print "Swaping columns 0 and 3. The new matrix is:"
    PrintArray SwapMatrixColumns(a, 0, 3)
End Sub

Public Sub TestSwapMatrixRows()
    Dim a As Variant
    
    Let a = [{1,2,3; 4,5,6; 7,8,9}]
    Debug.Print "a is:"
    PrintArray a
    Debug.Print "Swaping rows 1 and 3. The new matrix is:"
    PrintArray SwapMatrixRows(a, 1, 3)
    
    Debug.Print "Swaping rows 0 and 3. The new matrix is:"
    PrintArray SwapMatrixRows(a, 0, 3)
End Sub

Public Sub TestSwapRangeColumns()
    Dim a As Range
    
    Call TempComputation.UsedRange.ClearContents
    Set a = TempComputation.Range("A1").Resize(3, 3)
    Let a.Value2 = [{1,2,3; 4,5,6; 7,8,9}]
    If Not SwapRangeColumns(a, 1, 3) Then
        MsgBox "You screwed up."
    End If
    
    If Not SwapRangeColumns(a, 0, 3) Then
        MsgBox "You screwed up."
    End If
End Sub

Public Sub TestSwapRangeRows()
    Dim a As Range

    Call TempComputation.UsedRange.ClearContents
    Set a = TempComputation.Range("A1").Resize(3, 3)
    Let a.Value2 = [{1,2,3; 4,5,6; 7,8,9}]
    If Not SwapRangeRows(a, 1, 3) Then
        MsgBox "You screwed up."
    End If
    
    If Not SwapRangeRows(a, 0, 3) Then
        MsgBox "You screwed up."
    End If
End Sub

Public Sub TestBrokerAllocationClass()
    Dim ba As BrokerAllocation
    
    Let Application.ScreenUpdating = False
    Let Application.DisplayAlerts = False
    
    Set ba = New BrokerAllocation
    Call ba.InitializeWithMySql(DbServerAddress, "etwip2dot0", "brokerallocation", DbUserName, DbPassword, "Equity", "Equity")
    Debug.Print "For the latest equity run, the allocation percentages are:"
    PrintArray ba.GetAllocationPercentages.Value2
    
    Set ba = New BrokerAllocation
    Call ba.InitializeWithMySql(DbServerAddress, "etwip2dot0", "brokerallocation", DbUserName, DbPassword, "Equity", "Equity", 1)
    Debug.Print "For runnumber 1 of equity, the allocation percentages are:"
    PrintArray ba.GetAllocationPercentages.Value2
    
    Call ba.InjectIntoMySql(DbServerAddress, "etwip2dot0", "brokerallocation", DbUserName, DbPassword)

    Let Application.ScreenUpdating = True
    Let Application.DisplayAlerts = True
End Sub

Public Sub TestStringQ()
    Dim Expressions As Variant
    Dim var As Variant
    
    Let Expressions = Array(1, 2, 3, "four", "f", True, False, Null)
    
    For Each var In Expressions
        Debug.Print IIf(IsNull(var), "Null", var) & " is a string is " & StringQ(var)
    Next
End Sub

Public Sub TestString1DArrayQ()
    Dim Expressions As Variant
    Dim var As Variant
    
    Let Expressions = Array(1, 2, 3, "four", "f", True, False, Null)
    Debug.Print "All are strings in array is " & String1DArrayQ(Expressions)

    Let Expressions = Array("four", "f")
    Debug.Print "All are strings in array is " & String1DArrayQ(Expressions)

    Let Expressions = Array()
    Debug.Print "All are strings in array is " & String1DArrayQ(Expressions)
End Sub

Public Sub TestArrayMap()
    Dim a As Variant
    
    Debug.Print "We are showing how to do MapThread using ArrayMap."
    Debug.Print
    Let a = TransposeMatrix(Pack2DArray(Array(Array(1, 2, 3), Array(4, 5, 6))))
    Debug.Print "a is:"
    PrintArray a
    Let a = ArrayMap("StringConcatenate", a)
    Debug.Print "Let a = TransposeMatrix(Pack2DArray(Array(Array(1, 2, 3), Array(4, 5, 6))))"
    Debug.Print "Let a = ArrayMap(""StringConcatenate"", a)"
    PrintArray a
End Sub

Public Sub TestCreateDictionary1()
    Dim TheKeys As Variant
    Dim TheItems As Variant
    Dim aDict As Dictionary
    
    Set aDict = New Dictionary
    
    Let TheKeys = Array(1, 2, 3, 4, 5)
    Let TheItems = Array(10, 20, 30, 40, 50)
    
    Set aDict = CreateDictionary(TheKeys, TheItems)
    
    Debug.Print "The keys are:"
    PrintArray aDict.Keys
    
    Debug.Print "The items are:"
    PrintArray aDict.Items
End Sub

Public Sub TestCreateDictionary2()
    Dim TheKeys As Variant
    Dim TheItems As Variant
    Dim aDict As Dictionary
    
    Set aDict = New Dictionary
    
    Let TheKeys = Array(1, 2, 3)
    Let TheItems = Array(Array(10, 100, 1000), Array(20, 200, 2000), Array(30, 300, 3000))
    
    Set aDict = CreateDictionary(TheKeys, TheItems)
End Sub

Public Sub TestCreateTableDictionary()
    Dim AListObject As ListObject
    Dim TheHeaders As Variant
    Dim TheData As Variant
    Dim r As Integer
    Dim var As Variant
    Dim aDict As Dictionary
    
    Let TheHeaders = Array("Col1", "Col2", "Col3", "Col4", "Col5")
    Let TheData = RandomMatrix(20, 5)
    
    For r = LBound(TheData, 1) To UBound(TheData, 1)
        Let TheData(r, LBound(TheData, 1)) = "Col" & r
    Next r
    
    Let TheData = Prepend(TheData, TheHeaders)
    Call ToTemp(TheData)
    
    Set AListObject = TempComputation.ListObjects.Add(SourceType:=xlSrcRange, _
                                                      Source:=[TempComputation!A1].CurrentRegion, _
                                                      XlListObjectHasHeaders:=xlYes)
    
    Set aDict = CreateTableDictionary(AListObject, "Col1", Array("Col2", "Col5"))
    
End Sub

Public Sub TestStringJoin()
    Dim s As Variant
    Dim s2 As Variant
    
    Let s = Array("1", "2", "3")
    Debug.Print "Let s = Array(""1"", ""2"", ""3"")"
    Debug.Print "StringJoin(s) = " & StringJoin(s)
    
    Debug.Print
    Let s = Array()
    Debug.Print "Let s = Array()"
    Debug.Print "StringJoin(s) = " & StringJoin(s)
    
    Debug.Print
    Let s = Array(1, 2)
    Debug.Print "Let s = Array(1,2)"
    If IsNull(StringJoin(s)) Then
        Debug.Print "StringJoin(s) is Null"
    Else
        Debug.Print "StringJoin(s) = " & StringJoin(s)
    End If
    
    Debug.Print
    Let s = Array("a", "b", "c")
    Let s2 = "z"
    Debug.Print "Let s = Array(""a"", ""b"", ""c"")"
    Debug.Print "Let s2 = ""z"""
    PrintArray StringJoin(s, s2)
    
    Debug.Print
    Let s = Array("a", "b", "c")
    Let s2 = Array("aa", "bb", "cc")
    Debug.Print "Let s = Array(""a"", ""b"", ""c"")"
    Debug.Print "Let s2 = Array(""aa"", ""bb"", ""cc"""
    PrintArray StringJoin(s, s2)
End Sub

Public Sub TestArrayMapThread()
    PrintArray ArrayMapThread("StringJoin", Array(Array("1", "2", "3"), Array("10", "20", "30")))
End Sub

Public Sub TestLeftJoin2DArraysOnKeyEquality()
    Dim t1 As Variant
    Dim t2 As Variant
    Dim T As Variant
    Dim key1 As Integer
    Dim key2 As Integer
    Dim cols1 As Variant
    Dim cols2 As Variant

    Debug.Print "Test 1"
    
    Let t1 = [{1, 10, 100, 1000, 10000; 2, 20, 200, 2000, 20000; 3, 30, 300, 3000, 30000}]
    Let t2 = [{1, 11, 111, 1111, 11111; 3, 33, 333, 3333, 33333; 4, 44, 444, 4444, 44444}]
    
    Let T = LeftJoin2DArraysOnKeyEquality(t1, 1, Array(2, 4), t2, 1, Array(3, 5))

    Debug.Print "t1 is:"
    PrintArray t1
    Debug.Print
    
    Debug.Print "t2 is:"
    PrintArray t2
    Debug.Print
    
    If IsNull(T) Then
        Debug.Print "There was an error in the parameters"
    Else
        Debug.Print "The left join is for t1 with columns 2 and 4 and t2 columns 3 and 5:"
        PrintArray T
    End If
    
    Debug.Print
    Debug.Print



    Debug.Print "Test 2"
    
    Let t1 = [{1, 10, 100, 1000, 10000; 2, 20, 200, 2000, 20000; 3, 30, 300, 3000, 30000}]
    Let t2 = [{1, 11, 111, 1111, 11111; 3, 33, 333, 3333, 33333; 4, 44, 444, 4444, 44444}]
    
    Let T = LeftJoin2DArraysOnKeyEquality(t1, 1, Array(-2, 4), t2, 1, Array(3, 5))

    Debug.Print "t1 is:"
    PrintArray t1
    Debug.Print
    
    Debug.Print "t2 is:"
    PrintArray t2
    Debug.Print
    
    If IsNull(T) Then
        Debug.Print "There was an error in the parameters"
    Else
        Debug.Print "The left join is for t1 with columns 2 and 4 and t2 columns 3 and 5:"
        PrintArray T
    End If



    Debug.Print "Test 3"
    
    Let t1 = [{"Col1", "Col2", "Col3", "Col4", "Col5"; 1, 10, 100, 1000, 10000; 2, 20, 200, 2000, 20000; 3, 30, 300, 3000, 30000}]
    Let t2 = [{"Col1", "Col2", "Col3", "Col4", "Col5";1, 11, 111, 1111, 11111; 3, 33, 333, 3333, 33333; 4, 44, 444, 4444, 44444}]
    
    Let T = LeftJoin2DArraysOnKeyEquality(t1, 1, Array(2, 4), t2, 1, Array(3, 5), True, True)

    Debug.Print "t1 is:"
    PrintArray t1
    Debug.Print
    
    Debug.Print "t2 is:"
    PrintArray t2
    Debug.Print
    
    If IsNull(T) Then
        Debug.Print "There was an error in the parameters"
    Else
        Debug.Print "The left join is for t1 with columns 2 and 4 and t2 columns 3 and 5:"
        PrintArray T
    End If
End Sub

Public Sub TestInnerJoin2DArraysOnKeyEquality()
    Dim t1 As Variant
    Dim t2 As Variant
    Dim T As Variant
    Dim key1 As Integer
    Dim key2 As Integer
    Dim cols1 As Variant
    Dim cols2 As Variant

    Debug.Print "Test 1"

    Let t1 = [{1, 10, 100, 1000, 10000; 2, 20, 200, 2000, 20000; 3, 30, 300, 3000, 30000}]
    Let t2 = [{1, 11, 111, 1111, 11111; 3, 33, 333, 3333, 33333; 4, 44, 444, 4444, 44444}]
    
    Let T = InnerJoin2DArraysOnKeyEquality(t1, 1, Array(2, 4), t2, 1, Array(3, 5))
    
    Debug.Print "t1 is:"
    PrintArray t1
    Debug.Print
    
    Debug.Print "t2 is:"
    PrintArray t2
    Debug.Print
    
    If IsNull(T) Then
        Debug.Print "There was an error in the parameters"
    Else
        Debug.Print "The left join is for t1 with columns 2 and 4 and t2 columns 3 and 5:"
        PrintArray T
    End If
    
    Debug.Print
    Debug.Print



    Debug.Print "Test 2"

    Let t1 = [{1, 10, 100, 1000, 10000; 2, 20, 200, 2000, 20000; 3, 30, 300, 3000, 30000}]
    Let t2 = [{1, 11, 111, 1111, 11111; 3, 33, 333, 3333, 33333; 4, 44, 444, 4444, 44444}]
    
    Let T = InnerJoin2DArraysOnKeyEquality(t1, 1, Array(-2, 4), t2, 1, Array(3, 5))
    
    Debug.Print "t1 is:"
    PrintArray t1
    Debug.Print
    
    Debug.Print "t2 is:"
    PrintArray t2
    Debug.Print
    
    If IsNull(T) Then
        Debug.Print "There was an error in the parameters"
    Else
        Debug.Print "The left join is for t1 with columns 2 and 4 and t2 columns 3 and 5:"
        PrintArray T
    End If



    Debug.Print "Test 3"
    
    Let t1 = [{"Col1", "Col2", "Col3", "Col4", "Col5"; 1, 10, 100, 1000, 10000; 2, 20, 200, 2000, 20000; 3, 30, 300, 3000, 30000}]
    Let t2 = [{"Col1", "Col2", "Col3", "Col4", "Col5";1, 11, 111, 1111, 11111; 3, 33, 333, 3333, 33333; 4, 44, 444, 4444, 44444}]
    
    Let T = InnerJoin2DArraysOnKeyEquality(t1, 1, Array(2, 4), t2, 1, Array(3, 5), True, True)

    Debug.Print "t1 is:"
    PrintArray t1
    Debug.Print
    
    Debug.Print "t2 is:"
    PrintArray t2
    Debug.Print
    
    If IsNull(T) Then
        Debug.Print "There was an error in the parameters"
    Else
        Debug.Print "The left join is for t1 with columns 2 and 4 and t2 columns 3 and 5:"
        PrintArray T
    End If
End Sub

Public Sub TestCast()
    Debug.Print TypeName(Cast(Array(1, 2, 3), xlParamTypeChar))
    Debug.Print TypeName(Cast(Array(1, 2, 3), xlParamTypeInteger))
    Debug.Print TypeName(Cast(Array(1, 2, 3), xlParamTypeDouble))
    Debug.Print TypeName(Cast(Array(1, 2, 3), xlParamTypeBinary))
    Debug.Print TypeName(Cast(Array(1, 2, 3), xlParamTypeBit))
End Sub

Public Sub TestLeftJoinListObjectsOnKeyEquality()
    Dim t1 As Variant
    Dim t2 As Variant
    Dim l1 As ListObject
    Dim l2 As ListObject
    Dim r1 As Range
    Dim r2 As Range
    Dim h1 As Variant
    Dim h2 As Variant
    Dim cl1() As String
    Dim cl2() As String
    Dim l As Variant
    Dim d As Dictionary
    Dim key1 As Integer
    Dim key2 As Integer
    Dim cols1 As Variant
    Dim cols2 As Variant

    Debug.Print "Test 1"
    
    Let h1 = Array("key1", "A1Col1", "A1Col2", "A1Col3", "A1Col4")
    Let h2 = Array("key2", "A2Col1", "A2Col2", "A2Col3", "A2Col4")
    
    Let t1 = [{1, 10, 100, 1000, 10000; 2, 20, 200, 2000, 20000; 3, 30, 300, 3000, 30000}]
    Let t2 = [{1, 11, 111, 1111, 11111; 3, 33, 333, 3333, 33333; 4, 44, 444, 4444, 44444}]
    
    Let cl1 = Cast(Array("A1Col2", "A1Col4"), xlParamTypeChar)
    Let cl2 = Cast(Array("A2Col2", "A2Col4"), xlParamTypeChar)
    
    Call TempComputation.UsedRange.ClearFormats
    Call TempComputation.UsedRange.ClearContents
    
    Set r1 = ToTemp(Prepend(t1, h1))
    Set r2 = DumpInSheet(Prepend(t2, h2), r1.Cells(1, 1).Offset(r1.Rows.Count + 1, 0))
    
    Set l1 = TempComputation.ListObjects.Add(SourceType:=xlSrcRange, _
                                             Source:=r1, _
                                             XlListObjectHasHeaders:=xlYes)
    Set l2 = TempComputation.ListObjects.Add(SourceType:=xlSrcRange, _
                                             Source:=r2, _
                                             XlListObjectHasHeaders:=xlYes)
        
    Let l = LeftJoinListObjectsOnKeyEquality(l1, "key1", cl1, l2, "key2", cl2)

    Debug.Print "l1 is:"
    PrintArray l1.Range.Value2
    Debug.Print

    Debug.Print "l2 is:"
    PrintArray l2.Range.Value2
    Debug.Print

    If IsNull(l) Then
        Debug.Print "There was an error in the parameters"
    Else
        Debug.Print "The left join is for t1 with columns 2 and 4 and t2 columns 3 and 5:"
        PrintArray l
    End If
    
    
    Debug.Print
    Debug.Print
    Debug.Print "Test 2"
    
    Let h1 = Array("key1", "A1Col1", "A1Col2", "A1Col3", "A1Col4")
    Let h2 = Array("key2", "A2Col1", "A2Col2", "A2Col3", "A2Col4", "A2Col5")
    
    Let t1 = [{1, 10, 100, 1000, 10000; 2, 20, 200, 2000, 20000; 3, 30, 300, 3000, 30000}]
    Let t2 = [{1, 11, 111, 1111, 11111, 111111; 3, 33, 333, 3333, 33333, 333333; 4, 44, 444, 4444, 44444, 444444}]
    
    Let cl1 = Cast(Array("A1Col2", "A1Col4"), xlParamTypeChar)
    Let cl2 = Cast(Array("A2Col2", "A2Col5"), xlParamTypeChar)
    
    Call TempComputation.UsedRange.ClearFormats
    Call TempComputation.UsedRange.ClearContents
    
    Set r1 = ToTemp(Prepend(t1, h1))
    Set r2 = DumpInSheet(Prepend(t2, h2), r1.Cells(1, 1).Offset(r1.Rows.Count + 1, 0))
    
    Set l1 = TempComputation.ListObjects.Add(SourceType:=xlSrcRange, _
                                             Source:=r1, _
                                             XlListObjectHasHeaders:=xlYes)
    Set l2 = TempComputation.ListObjects.Add(SourceType:=xlSrcRange, _
                                             Source:=r2, _
                                             XlListObjectHasHeaders:=xlYes)
        
    Set d = LeftJoinListObjectsOnKeyEquality(l1, "key1", cl1, l2, "key2", cl2, True)

    Debug.Print "l1 is:"
    PrintArray l1.Range.Value2
    Debug.Print

    Debug.Print "l2 is:"
    PrintArray l2.Range.Value2
    Debug.Print

    If d Is Nothing Then
        Debug.Print "There was an error in the parameters"
    Else
        Debug.Print "The left join is for t1 with columns 2 and 4 and t2 columns 3 and 5:"
        PrintArray Pack2DArray(d.Items)
    End If
End Sub


Public Sub TestInnerJoinListObjectsOnKeyEquality()
    Dim t1 As Variant
    Dim t2 As Variant
    Dim l1 As ListObject
    Dim l2 As ListObject
    Dim r1 As Range
    Dim r2 As Range
    Dim h1 As Variant
    Dim h2 As Variant
    Dim cl1() As String
    Dim cl2() As String
    Dim l As Variant
    Dim d As Dictionary
    Dim key1 As Integer
    Dim key2 As Integer
    Dim cols1 As Variant
    Dim cols2 As Variant

    Debug.Print "Test 1"
    
    Let h1 = Array("key1", "A1Col1", "A1Col2", "A1Col3", "A1Col4")
    Let h2 = Array("key2", "A2Col1", "A2Col2", "A2Col3", "A2Col4")
    
    Let t1 = [{1, 10, 100, 1000, 10000; 2, 20, 200, 2000, 20000; 3, 30, 300, 3000, 30000}]
    Let t2 = [{1, 11, 111, 1111, 11111; 3, 33, 333, 3333, 33333; 4, 44, 444, 4444, 44444}]
    
    Let cl1 = Cast(Array("A1Col2", "A1Col4"), xlParamTypeChar)
    Let cl2 = Cast(Array("A2Col2", "A2Col4"), xlParamTypeChar)
    
    Call TempComputation.UsedRange.ClearFormats
    Call TempComputation.UsedRange.ClearContents
    
    Set r1 = ToTemp(Prepend(t1, h1))
    Set r2 = DumpInSheet(Prepend(t2, h2), r1.Cells(1, 1).Offset(r1.Rows.Count + 1, 0))
    
    Set l1 = TempComputation.ListObjects.Add(SourceType:=xlSrcRange, _
                                             Source:=r1, _
                                             XlListObjectHasHeaders:=xlYes)
    Set l2 = TempComputation.ListObjects.Add(SourceType:=xlSrcRange, _
                                             Source:=r2, _
                                             XlListObjectHasHeaders:=xlYes)
        
    Let l = InnerJoinListObjectsOnKeyEquality(l1, "key1", cl1, l2, "key2", cl2)

    Debug.Print "l1 is:"
    PrintArray l1.Range.Value2
    Debug.Print

    Debug.Print "l2 is:"
    PrintArray l2.Range.Value2
    Debug.Print

    If IsNull(l) Then
        Debug.Print "There was an error in the parameters"
    Else
        Debug.Print "The left join is for t1 with columns 2 and 4 and t2 columns 3 and 5:"
        PrintArray l
    End If
    
    Debug.Print
    Debug.Print
    Debug.Print "Test 2"
    
    Let h1 = Array("key1", "A1Col1", "A1Col2", "A1Col3", "A1Col4")
    Let h2 = Array("key2", "A2Col1", "A2Col2", "A2Col3", "A2Col4", "A2Col5")
    
    Let t1 = [{1, 10, 100, 1000, 10000; 2, 20, 200, 2000, 20000; 3, 30, 300, 3000, 30000}]
    Let t2 = [{1, 11, 111, 1111, 11111, 111111; 3, 33, 333, 3333, 33333, 333333; 4, 44, 444, 4444, 44444, 444444}]
    
    Let cl1 = Cast(Array("A1Col2", "A1Col4"), xlParamTypeChar)
    Let cl2 = Cast(Array("A2Col2", "A2Col5"), xlParamTypeChar)
    
    Call TempComputation.UsedRange.ClearFormats
    Call TempComputation.UsedRange.ClearContents
    
    Set r1 = ToTemp(Prepend(t1, h1))
    Set r2 = DumpInSheet(Prepend(t2, h2), r1.Cells(1, 1).Offset(r1.Rows.Count + 1, 0))
    
    Set l1 = TempComputation.ListObjects.Add(SourceType:=xlSrcRange, _
                                             Source:=r1, _
                                             XlListObjectHasHeaders:=xlYes)
    Set l2 = TempComputation.ListObjects.Add(SourceType:=xlSrcRange, _
                                             Source:=r2, _
                                             XlListObjectHasHeaders:=xlYes)
        
    Set d = InnerJoinListObjectsOnKeyEquality(l1, "key1", cl1, l2, "key2", cl2, True)

    Debug.Print "l1 is:"
    PrintArray l1.Range.Value2
    Debug.Print

    Debug.Print "l2 is:"
    PrintArray l2.Range.Value2
    Debug.Print

    If d Is Nothing Then
        Debug.Print "There was an error in the parameters"
    Else
        Debug.Print "The left join is for t1 with columns 2 and 4 and t2 columns 3 and 5:"
        PrintArray Pack2DArray(d.Items)
    End If
End Sub

Public Sub TestTransposeRectangular1DArrayOf1DArrays()
    Dim var As Variant
    Dim AnArray As Variant
    
    Let AnArray = TransposeRectangular1DArrayOf1DArrays(Array(Array(1, 2, 3), Array(10, 20, 30)))
    
    For Each var In AnArray
        PrintArray var
    Next
    
End Sub

Public Sub TestSimulatingMapThreadTransposeRectangular1DArrayOf1DArrays()
    Dim var As Variant
    Dim AnArray As Variant
    
    Let AnArray = TransposeRectangular1DArrayOf1DArrays(Array(Array(1, 2, 3), Array(10, 20, 30)))
    
    For Each var In AnArray
        Debug.Print Application.Sum(var)
    Next

End Sub

Public Function TestAddListObject() As ListObject
    Dim TheHeaders() As String
    Dim TheData() As Variant
    Dim lo As ListObject
    
    For Each lo In TempComputation.ListObjects
        Call lo.Delete
    Next
    
    Let TheHeaders = Cast(Array("Col1", "Col2", "Col3"), xlParamTypeChar)
    Let TheData = [{1,2,3; 10,20,30; 100,20,300; 1000, 2000, 3000}]
    
    Call ToTemp(TheHeaders)
    Call DumpInSheet(TheData, TempComputation.Range("A2"))
    
    Set TestAddListObject = AddListObject(TempComputation.Range("A1"), "MyListObject")
End Function

Public Sub TestAddColumnsToListObject()
    Dim lo As ListObject
    Dim DataColumns As Variant
    Dim ColumnNames() As String
    Dim col1 As Variant
    Dim col2 As Variant
    
    Set lo = TestAddListObject()
    
    Let col1 = Array(11, 22, 33, 44)
    Let col2 = Array(111, 222, 333, 444)
    Let ColumnNames = Cast(Array("Col4", "Col5"), xlParamTypeChar)
    Let DataColumns = Array(col1, col2)
    
    Call AddColumnsToListObject(AListObject:=lo, _
                                ColumnNames:=ColumnNames, _
                                TheData:=DataColumns)
End Sub
