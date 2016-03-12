Attribute VB_Name = "ArrayFormulas"
Option Explicit
Option Base 1

' This function inserts either an atomic element into a 1D array satisfying RowVectorQ or
' a row into a 2D array filled exclusively with atomic elements
' AnArray is returned unevaluated if either AnArray, TheElt, or ThePos is not as expected.
' To insert in AnArray's first position, set ThePos = LBound(AnArray)
' To insert in AnArray's last position, set ThePos = UBound(AnArray)+1
' You should think of acceptable values for ThePos as insert TheElt at ThePos and
' shifting to the right anything in AnArray at and to the right of ThePos.
Public Function Insert(AnArray As Variant, TheElt As Variant, ThePos As Long) As Variant
    Dim FirstPart As Variant
    Dim LastPart As Variant

    ' Exit returning AnArray unevaluated if neither AnArray nor ThePos make sense
    If Not RowVectorQ(AnArray) And Not MatrixQ(AnArray) Or Not IsNumeric(ThePos) Then
        Let Insert = AnArray
        Exit Function
    End If
    
    ' Exit if AnArray is a row vector TheElt is not an atomic expression
    If RowVectorQ(AnArray) And IsArray(TheElt) Then
        Let Insert = AnArray
        Exit Function
    End If
    
    ' Exit if AnArray is a matrix and TheElt is not a row vector with the same number of columns
    If MatrixQ(AnArray) And (Not RowVectorQ(TheElt) Or GetNumberOfColumns(AnArray) <> GetArrayLength(TheElt)) Then
        Let Insert = AnArray
        Exit Function
    End If

    ' Exit if ThePos has unaceptable values
    If ThePos < LBound(AnArray) Or ThePos > UBound(AnArray) + 1 Then
        Let Insert = AnArray
        Exit Function
    End If
    
    ' Based on ThePos, shift parts of AnArray and insert TheElt
    If ThePos = LBound(AnArray) Then
        If NumberOfDimensions(AnArray) = 1 Then
            Let Insert = Prepend(AnArray, TheElt)
        Else
            Let Insert = Stack2DArrays(TheElt, AnArray)
        End If
    ElseIf ThePos = UBound(AnArray) + 1 Then
        If NumberOfDimensions(AnArray) = 1 Then
            Let Insert = Append(AnArray, TheElt)
        Else
            Let Insert = Stack2DArrays(AnArray, TheElt)
        End If
    ElseIf ThePos > LBound(AnArray) And ThePos <= UBound(AnArray) Then
        Let FirstPart = Take(AnArray, ThePos - 1)
        Let LastPart = Take(AnArray, -(UBound(AnArray) - ThePos + 1))
        
        If NumberOfDimensions(AnArray) = 1 Then
            Let FirstPart = Append(FirstPart, TheElt)
            Let Insert = ConcatenateArrays(FirstPart, LastPart)
        Else
            Let FirstPart = Stack2DArrays(FirstPart, TheElt)
            Let Insert = Stack2DArrays(FirstPart, LastPart)
        End If
    Else
        Let Insert = AnArray
    End If
End Function

' This function turns a 1D array of 1D arrays into a 2D array.
' Each of the elemements (inner arrays) of the outermost array satisfies
' RowVectorQ
'
' This is useful to quickly build a matrix from 1D arrays
' This function assumes that all elements of TheRowsAs1DArrays have the same lbound()
' arg is allowed to have different lbound from that of its elements
'
' The 2D array returned is indexed starting at 1
' If the optional parameter PackAsColumnsQ is set to True, the 1D arrays in TheRowsAs1DArrays become columns.
Public Function Pack2DArray(TheRowsAs1DArrays As Variant, Optional PackAsColumnsQ As Boolean = False) As Variant
    Dim var As Variant
    Dim r As Long
    Dim c As Long
    Dim Results() As Variant
    Dim TheLength As Long
    
    ' Exit if the argument is not the expected type
    If NumberOfDimensions(TheRowsAs1DArrays) <> 1 Or EmptyArrayQ(TheRowsAs1DArrays) Then
        Let Pack2DArray = Null
        Exit Function
    End If
    
    ' Exit if any of the elements in not an atomic array or
    '  if all the array elements do not the same length
    Let TheLength = GetArrayLength(First(TheRowsAs1DArrays))
    For Each var In TheRowsAs1DArrays
        If Not AtomicArrayQ(var) Or GetArrayLength(var) <> TheLength Then
            Let Pack2DArray = Null
            Exit Function
        End If
    Next

    ' Pre-allocate a 2D array filled with Empty
    ReDim Results(1 To GetArrayLength(TheRowsAs1DArrays), 1 To GetArrayLength(First(TheRowsAs1DArrays)))
    
    ' Pack the array
    For r = LBound(TheRowsAs1DArrays) To UBound(TheRowsAs1DArrays)
        For c = LBound(First(TheRowsAs1DArrays)) To UBound(First(TheRowsAs1DArrays))
            Let Results(IIf(LBound(TheRowsAs1DArrays) = 0, 1, 0) + r, IIf(LBound(First(TheRowsAs1DArrays)) = 0, 1, 0) + c) = TheRowsAs1DArrays(r)(c)
        Next c
    Next r
    
    If PackAsColumnsQ Then
        Let Pack2DArray = TransposeMatrix(Results)
        Exit Function
    End If
    
    Let Pack2DArray = Results
End Function

' If n is an integer:
' A negative n is interpreted as counting starting with 1 from the right of the 1D array or the bottom
' of the 2D array respectively.  Returns the first n elements of a 1D array or n rows from a 2D array
' Empty arrays returns an empty array (e.g. Array())
' If AnArray is not an array, function returns Null
' If AnArray has more than 2 dimensions, returns Null
'
' If n is an array of integers:
' Returns the elements or rows index by each element of n
'
' Examples:
' 1. Take(Array(1,2,3,4), 2) returns Array(1,2)
' 2. Take(Array(1,2,3,4), Array(2)) returns 2
'
' A similar thing happens to 2D arrays, but the elements are then the rows of the matrix.
Public Function Take(AnArray As Variant, N As Variant) As Variant
    Dim c As Long
    Dim r As Long
    Dim ResultArray() As Variant
    Dim var As Variant
    Dim RenormalizedIndices() As Long

    If Not DimensionedQ(AnArray) Then
        Let Take = Null
        Exit Function
    End If
    
    If EmptyArrayQ(AnArray) Then
        Let Take = Null
        Exit Function
    End If
    
    ' Exit with argument unchanged if the array has fewer or more than 2 dimensions
    If NumberOfDimensions(AnArray) > 2 Then
        Let Take = Null
        Exit Function
    End If

    ' Exit with argument unchanged if either N is not an integer or an array of integers
    If Not (WholeNumberQ(N) Or WholeNumberArrayQ(N)) Or EmptyArrayQ(N) Then
        Let Take = Null
        
        Exit Function
    End If
    
    ' Proceed with N as an integer index
    If WholeNumberQ(N) Then
        If N = 0 Then
            Let Take = Array()
        ElseIf NumberOfDimensions(AnArray) = 1 And N > 0 Then
            Let Take = GetSubArray(AnArray, LBound(AnArray), CLng(N) - (1 - LBound(AnArray, 1)))
        ElseIf NumberOfDimensions(AnArray) = 1 And N < 0 Then
            Let Take = GetSubArray(AnArray, Application.Max(UBound(AnArray) + N + 1, LBound(AnArray)), UBound(AnArray))
        ElseIf NumberOfDimensions(AnArray) = 2 And N > 0 Then
            Let Take = GetSubMatrix(AnArray, LBound(AnArray, 1), Application.Min(LBound(AnArray, 1) + N - 1, UBound(AnArray, 1)), LBound(AnArray, 2), UBound(AnArray, 2))
        ElseIf NumberOfDimensions(AnArray) = 2 And N < 0 Then
            Let Take = GetSubMatrix(AnArray, Application.Max(UBound(AnArray, 1) + N + 1, LBound(AnArray, 1)), UBound(AnArray, 1), LBound(AnArray, 2), UBound(AnArray, 2))
        End If
        
        Exit Function
    End If
    
    ' Proceed with N as an array of integer indices
    ' Turn all indices in N into their positive equivalents
    Let RenormalizedIndices = ToLongs(NormalizeArrayIndices(AnArray:=AnArray, _
                                                            TheIndices:=ToLongs(N), _
                                                            NormalizeTo1Q:=True))
    
    ' Proceed if AnArray is 1D
    If NumberOfDimensions(AnArray) = 1 Then
        ReDim ResultArray(LBound(N, 1) To UBound(N, 1))
        
        For c = LBound(RenormalizedIndices, 1) To UBound(RenormalizedIndices, 1)
            Let ResultArray(c) = AnArray(RenormalizedIndices(c) - IIf(N(c) <= 0, 0, (1 - LBound(AnArray, 1))))
        Next
        
        Let Take = ResultArray
        
        Exit Function
    End If

    ' Proceed here if AnArray is 2D
    ' Pre-allocate a matrix big enough to hold all the requested elements
    ReDim ResultArray(1 To GetArrayLength(N))
    
    For r = 1 To GetArrayLength(N)
        If N(r) < 0 Then
            Let ResultArray(r) = GetElement(AnArray, GetElement(N, r - (1 - LBound(N, 1))))
        Else
            Let ResultArray(r) = GetElement(AnArray, GetElement(N, r - (1 - LBound(N, 1))) - (1 - LBound(AnArray, 1)))
        End If
    Next
    
    Let Take = Pack2DArray(ResultArray)
End Function

' This function is the exact opposite of Take.  It removes what take would return
Public Function Drop(AnArray As Variant, N As Variant) As Variant
    Dim DualIndex As Variant
    Dim RenormalizedIndices As Variant
    Dim var As Variant
    Dim c As Long

    ' Exit with argument unchanged if the array has fewer or more than 2 dimensions
    If NumberOfDimensions(AnArray) = 0 Or NumberOfDimensions(AnArray) > 2 Then
        Let Drop = AnArray

        Exit Function
    End If

    ' Exit with argument unchanged if either N is not an integer or an array of integers
    If Not (IsNumeric(N) Or IsNumericArrayQ(N)) Or EmptyArrayQ(N) Or EmptyArrayQ(AnArray) Then
        Let Drop = Array()
        
        Exit Function
    End If
    
    If IsNumericArrayQ(N) Then
        For Each var In N
            If CLng(var) <> var Then
                Let Drop = Array()
                
                Exit Function
            End If
        Next
    End If
    
    ' Case of N an integer
    If NumberOfDimensions(N) = 0 Then
        If NumberOfDimensions(AnArray) = 1 Then
            If N > 0 Then
                Let DualIndex = IIf(N > GetArrayLength(AnArray), Array(), -(GetArrayLength(AnArray) - N))
            Else
                Let DualIndex = IIf(Abs(N) > GetArrayLength(AnArray), Array(), GetArrayLength(AnArray) + N)
            End If
        Else
            If N > 0 Then
                Let DualIndex = IIf(N > GetNumberOfRows(AnArray), Array(), -(GetNumberOfRows(AnArray) - N))
            Else
                Let DualIndex = IIf(Abs(N) > GetNumberOfRows(AnArray), Array(), GetNumberOfRows(AnArray) + N)
            End If
        End If
    
        Let Drop = Take(AnArray, DualIndex)
    
        Exit Function
    End If
    
    ' Proceed with case of N a 1D array of integers
    
    ' Turn all indices in N into their positive equivalents
    Let c = 1
    Let RenormalizedIndices = ConstantArray(Empty, GetArrayLength(N))
    For Each var In N
        Let RenormalizedIndices(c) = IIf(var < 0, IIf(NumberOfDimensions(AnArray) = 1, GetNumberOfColumns(AnArray), GetNumberOfRows(AnArray)) + var + 1, var)
        Let c = c + 1
    Next
    
    Let Drop = Take(AnArray, ComplementOfSets(CreateSequentialArray(1, GetArrayLength(AnArray)), RenormalizedIndices))
End Function

' This function does the same thing as Worksheet.CopyFromRecordSet, but it does not fail when a column has an entry more than 255 characters long.
Public Function ConvertRecordSetToMatrix(rst As ADODB.Recordset, Optional ReturnOption As ConvertRecordSetPayloadToMatrixOptionsType = HeadersAndBody) As Variant
    Dim TheResults() As Variant
    Dim CurrentRow As Long
    Dim RowCount As Long
    Dim NColumns As Long
    Dim FirstRow As Long
    Dim r As Long
    Dim c As Long
    Dim h As Long
    
    Let NColumns = rst.Fields.Count
    Select Case ReturnOption
        Case HeadersAndBody
            Let FirstRow = 2
        Case Body
            Let FirstRow = 1
        Case Else
            Let FirstRow = 1
    End Select
    
    ReDim TheResults(1 To NColumns, 1 To 1)
    Let RowCount = 1
    
    If ReturnOption = Headers Or ReturnOption = HeadersAndBody Then
        For h = 0 To NColumns - 1
            Let TheResults(h + 1, 1) = rst.Fields(h).Name
        Next h
    End If
    
    If ReturnOption = Body Then
        Let CurrentRow = 1
    ElseIf ReturnOption = HeadersAndBody Then
        Let CurrentRow = 2
    End If
    
    While Not rst.EOF
        Let RowCount = RowCount + 1
        ReDim Preserve TheResults(1 To NColumns, 1 To RowCount)
    
        For c = 0 To NColumns - 1
            Let TheResults(c + 1, CurrentRow) = rst.Fields(c).Value
        Next c
        
        Call rst.MoveNext
        Let CurrentRow = CurrentRow + 1
    Wend
    
    Let ConvertRecordSetToMatrix = TransposeMatrix(TheResults)
End Function

Public Function Convert1DArrayIntoParentheticalExpression(TheArray As Variant) As String
    Let Convert1DArrayIntoParentheticalExpression = "(" & Join(TheArray, ",") & ")"
End Function

' This function prints a 0D, 1D, or 2D array in the debug window.
Public Sub PrintArray(TheArray As Variant)
    Dim ARow As Variant
    Dim c As Long
    Dim r As Long
    
    If Not IsArray(TheArray) Then
        Debug.Print "Not an array"
    ElseIf NumberOfDimensions(TheArray) = 0 Then
        Debug.Print TheArray
    ElseIf NumberOfDimensions(TheArray) = 1 Then
        If EmptyArrayQ(TheArray) Then
                Debug.Print "Empty 1D array"
        Else
            Let ARow = TheArray(LBound(TheArray))
        
            If UBound(TheArray) - LBound(TheArray) >= 1 Then
                For c = LBound(TheArray) + 1 To UBound(TheArray)
                    Let ARow = ARow & vbTab & TheArray(c)
                Next c
            End If
        
            Debug.Print ARow
        End If
    Else
        For r = LBound(TheArray, 1) To UBound(TheArray, 1)
            Let ARow = TheArray(r, LBound(TheArray, 2))
        
            If UBound(TheArray, 2) - LBound(TheArray, 2) >= 1 Then
                For c = LBound(TheArray, 2) + 1 To UBound(TheArray, 2)
                    Let ARow = ARow & vbTab & TheArray(r, c)
                Next c
            End If
        
            Debug.Print ARow
        Next r
    End If
End Sub

' This function sorts the given 2D matrix by the columns whose positions are given by
' ArrayOfColPos. The sorting orientation in each column are in ArrayOfColsSortOrder
' ArrayOfColsSortOrder is a variant array whose elements are all of enumerated type XLSortOrder
' (e.g. xlAscending, xlDescending)
Public Function Sort2DArray(MyArray As Variant, ArrayOfColPos As Variant, _
                     ArrayOfColsSortOrder As Variant, _
                     WithHeaders As XlYesNoGuess) As Variant
    Dim TheRange As Range
    Dim TmpSheet As Worksheet
    Dim i As Integer
    
    ' Set pointer to temp sheet and clear its used range
    Set TmpSheet = ThisWorkbook.Worksheets("TempComputation")
    TmpSheet.UsedRange.ClearContents
    
    ' Dump array in temp sheet
    'Let TmpSheet.Range("A1").Resize(UBound(MyArray, 1), UBound(MyArray, 2)).Value2 = MyArray
    Call DumpInSheet(MyArray, TempComputation.Range("A1"), True)

    ' Set range pointer to the data we just dumped in the temp sheet
    Set TheRange = TmpSheet.Range("A1").Resize(GetNumberOfRows(MyArray), GetNumberOfColumns(MyArray))

    ' Clear any previous sorting criteria
    TmpSheet.Sort.SortFields.Clear

    ' Add all the sorting criteria
    For i = LBound(ArrayOfColPos) To UBound(ArrayOfColPos)
        ' Add criteria to sort by date
        TheRange.Worksheet.Sort.SortFields.Add _
            Key:=TheRange.Columns(ArrayOfColPos(i)), _
            SortOn:=xlSortOnValues, _
            Order:=ArrayOfColsSortOrder(i), _
            DataOption:=xlSortNormal
    Next i
    
    ' Execute the sort
    With TheRange.Worksheet.Sort
        .SetRange TheRange
        .Header = WithHeaders
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    ' Extract horizontally reversed matrix and set them to return when the function exits
    Let Sort2DArray = TmpSheet.Range("A1").CurrentRegion.Value2
End Function

' This function reverses the horizontal order of an array (1D or 2D)
' It assumes that the first column must remain unchanged.
'***HERE Implement using loops
Public Function ReverseHorizontally(MyArray As Variant) As Variant
    Dim TheRange As Range
    Dim TmpSheet As Worksheet

    ' Set pointer to temp sheet and clear its used range
    Set TmpSheet = ThisWorkbook.Worksheets("TempComputation")
    TmpSheet.UsedRange.ClearContents
    
    ' Dump array in temp sheet
    Let TmpSheet.Range("A1").Resize(UBound(MyArray, 1), UBound(MyArray, 2)).Value2 = MyArray

    ' Insert an empty row at the top where we will place the horizontal column counter
    TmpSheet.Range("1:1").Insert
    
    ' Insert column counter in first row
    Let TmpSheet.Range("A1").Resize(1, TmpSheet.Range("A2").CurrentRegion.Columns.Count).Formula = "=column()"
    
    ' Change the column index of the first column to be larger than the number of columns so it does not change
    ' position
    Let TmpSheet.Range("A1").Value2 = Application.Max(TmpSheet.Range(Range("B1"), Range("B1").End(xlToRight))) + 1
    
    ' Copy and paste column Indices as values
    TmpSheet.Range(Range("A1"), Range("A1").End(xlToRight)).Value2 = TmpSheet.Range(Range("A1"), Range("A1").End(xlToRight)).Value2
    
    ' Set range pointer to the data we just dumped in the temp sheet
    Set TheRange = TmpSheet.Range("A1").CurrentRegion

    ' Clear any previous sorting criteria
    TmpSheet.Sort.SortFields.Clear

    ' Add criteria to sort by date
    TheRange.Worksheet.Sort.SortFields.Add _
        Key:=TheRange.Rows(1), _
        SortOn:=xlSortOnValues, _
        Order:=xlDescending, _
        DataOption:=xlSortNormal
        
    ' Execute the sort
    With TheRange.Worksheet.Sort
        .SetRange TheRange
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlLeftToRight
        .SortMethod = xlPinYin
        .Apply
    End With
    
    ' Delete the header row
    TmpSheet.Range("1:1").Delete
    
    ' Extract horizontally reversed matrix and set them to return when the function exits
    Let ReverseHorizontally = TmpSheet.Range("A1").CurrentRegion.Value2
End Function

' This function reverses the horizontal order of an array (1D or 2D)
' It assumes that the first column must remain unchanged.
'***HERE Implement using loops
Public Function ReverseVertically(MyArray As Variant) As Variant
    Dim TheRange As Range
    Dim TmpSheet As Worksheet
    
    ' Set pointer to temp sheet and clear its used range
    Set TmpSheet = ThisWorkbook.Worksheets("TempComputation")
    TmpSheet.UsedRange.ClearContents
    
    ' Dump array in temp sheet
    Let TmpSheet.Range("A1").Resize(UBound(MyArray, 1), UBound(MyArray, 2)).Value2 = MyArray

    ' Insert an empty row at the top where we will place the horizontal column counter
    TmpSheet.Range("A:A").Insert
    
    ' Insert column counter in first row
    Let TmpSheet.Range("A1").Resize(TmpSheet.Range("A2").CurrentRegion.Rows.Count, 1).Formula = "=row()"
    
    ' Change the row index of the first row to be larger than the number of row so it does not change
    ' position
    Let TmpSheet.Range("A1").Value2 = Application.Max(TmpSheet.Range(TmpSheet.Range("A2"), TmpSheet.Range("A2").End(xlDown))) + 1
    
    ' Copy and paste column Indices as values
    Let TmpSheet.Range("A1").CurrentRegion.Columns(1).Value2 = TmpSheet.Range("A1").CurrentRegion.Columns(1).Value2
    
    ' Set range pointer to the data we just dumped in the temp sheet
    Set TheRange = TmpSheet.Range("A1").CurrentRegion

    ' Clear any previous sorting criteria
    TmpSheet.Sort.SortFields.Clear

    ' Add criteria to sort by date
    TheRange.Worksheet.Sort.SortFields.Add _
        Key:=TheRange.Columns(1), _
        SortOn:=xlSortOnValues, _
        Order:=xlDescending, _
        DataOption:=xlSortNormal
        
    ' Execute the sort
    With TheRange.Worksheet.Sort
        .SetRange TheRange
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    ' Delete the header row
    TmpSheet.Range("A:A").Delete
    
    ' Extract horizontally reversed matrix and set them to return when the function exits
    Let ReverseVertically = TmpSheet.Range("A1").CurrentRegion.Value2
End Function

' This function returns the number of dimensions in an array.  There is no built-in function
' to do this.  Atomic expressions are assigned 0.
' Atomic expressions have dimension 0
' array() has dimension 1
' [{1,2,3; 4, 5, 6}] has dimension 2
' However, [{}] has dimension 0
Public Function NumberOfDimensions(MyArray As Variant) As Long
    Dim temp As Long
    Dim i As Long
    
    On Error GoTo FinalDimension
    
    If EmptyArrayQ(MyArray) Then
        Let NumberOfDimensions = 1
        Exit Function
    End If
    
    If Not IsArray(MyArray) Then
        Let NumberOfDimensions = 0
        Exit Function
    End If
    
    For i = 1 To 60000
        Let temp = LBound(MyArray, i)
    Next i

    Let NumberOfDimensions = i
    
    Exit Function
        
FinalDimension:
    Let NumberOfDimensions = i - 1
    Exit Function
End Function

' Alias for NumberOfDimensions()
Public Function GetNumberOfDimensions(MyArray As Variant) As Long
    Let GetNumberOfDimensions = NumberOfDimensions(MyArray)
End Function

' Alias for Public Function ConvertTo1DArray(a As Variant) As Variant in this module
' This one handles any number of dimensions and nesting.
Public Function Flatten(a As Variant) As Variant
    Dim var As Variant
    Dim var2 As Variant
    Dim ResultsDict As Dictionary
    
    Set ResultsDict = New Dictionary
    
    For Each var In a
        If AtomicQ(var) Then
            Call ResultsDict.Add(Key:=ResultsDict.Count, Item:=var)
        Else
            For Each var2 In Flatten(var)
                Call ResultsDict.Add(Key:=ResultsDict.Count, Item:=var2)
            Next
        End If
    Next
    
    Let Flatten = ResultsDict.Items
End Function

' Convert two-dimensional representation of a one-dimensional array to a one-dimensional array.
' If a is not an array, the function returns Array(a).  If the argument a is an array of dims
' other than Nx1 or a 1xN, the function returns arg unchanged.
Public Function ConvertTo1DArray(a As Variant) As Variant
    Dim nd As Integer
    Dim i As Long
    Dim TheResult() As Variant
    
    If EmptyArrayQ(a) Then
        Let ConvertTo1DArray = Array()
        Exit Function
    ElseIf Not IsArray(a) Then
        Let ConvertTo1DArray = Array(a)
        Exit Function
    End If
    
    Let nd = NumberOfDimensions(a)

    If RowArrayQ(a) Then
        ' Process arg is it has one dimensions
        If nd = 1 Then
            ' If this is a 1-element, 1D array
            If LBound(a) = UBound(a) Then
                If IsArray(a) Then Let ConvertTo1DArray = ConvertTo1DArray(a(LBound(a, 1)))
            ' If this is a multi-element 1D array
            Else
                Let ConvertTo1DArray = a
            End If
            
            Exit Function
        End If

        ' If we get here the array is two dimensional
        ' This is the case when there are more than one row and more than column
        If UBound(a, 1) = LBound(a, 1) And UBound(a, 2) = LBound(a, 2) Then
            If IsArray(a(LBound(a, 1), LBound(a, 2))) Then
                Let ConvertTo1DArray = ConvertTo1DArray(a(LBound(a, 1), LBound(a, 2)))
            Else
                Let ConvertTo1DArray = Array(a(LBound(a, 1), LBound(a, 2)))
            End If
        ElseIf UBound(a, 1) > LBound(a, 1) Then
            Let ConvertTo1DArray = Array()
        ' This is the case when there is just one row
        ElseIf UBound(a, 2) > LBound(a, 2) Then
            ReDim TheResult(LBound(a, 2) To UBound(a, 2))
            For i = LBound(a, 2) To UBound(a, 2)
                Let TheResult(i) = a(UBound(a, 1), i)
            Next i
            
            Let ConvertTo1DArray = TheResult
        ' This is the case when there is just one row
        Else
            ReDim TheResult(LBound(a, 1) To UBound(a, 1))
            For i = LBound(a, 1) To UBound(a, 1)
                Let TheResult(i) = a(i, UBound(a, 2))
            Next i
            
            Let ConvertTo1DArray = TheResult
        End If
    ElseIf ColumnArrayQ(a) Then
        If EmptyArrayQ(a) Then
            Let ConvertTo1DArray = Array()
            Exit Function
        End If

        If nd = 1 Then
            If LBound(a) = UBound(a) Then
                If Not IsArray(a(LBound(a))) Then
                    Let ConvertTo1DArray = a
                Else
                    Let ConvertTo1DArray = ConvertTo1DArray(a(LBound(a)))
                End If
            End If
            
            Exit Function
        End If
        
        If UBound(a, 1) > LBound(a, 1) And UBound(a, 2) > LBound(a, 2) Then
            Let ConvertTo1DArray = Array()
        ElseIf UBound(a, 1) > LBound(a, 1) Then
            ReDim TheResult(LBound(a, 1) To UBound(a, 1))
        
            For i = LBound(a, 1) To UBound(a, 1)
                Let TheResult(i) = a(i, LBound(a, 1))
            Next i
        
            Let ConvertTo1DArray = TheResult
        ElseIf UBound(a, 1) = LBound(a, 1) Then
            If IsArray(a(LBound(a, 1), LBound(a, 2))) Then
                Let ConvertTo1DArray = ConvertTo1DArray(a(LBound(a, 1), LBound(a, 2)))
            Else
                Let ConvertTo1DArray = Array(a(LBound(a, 1), LBound(a, 2)))
            End If
        End If
    Else
        Let ConvertTo1DArray = Array()
    End If
End Function

' Returns a sequential, 1D array of N integers from StartNumber to StartNumber+N-1 inclusively
Public Function CreateSequentialArray(StartNumber As Long, N As Long, Optional TheStep As Integer = 1)
    Dim TheArray As Variant
    Dim i As Long
    
    ReDim TheArray(1 To N)
    
    For i = StartNumber To StartNumber + N - 1
        Let TheArray(i - StartNumber + 1) = StartNumber + (i - StartNumber) * TheStep
    Next i
    
    Let CreateSequentialArray = TheArray
End Function

' This function returns a 1D array unchanged
' Regardless of the array's indexing (starts at 0 or 1), the user must request the first row as row 1.
Public Function GetRow(aMatrix As Variant, RowNumber As Long) As Variant
    Dim nd As Long
    Dim i As Long
    Dim TheResult() As Variant

    If EmptyArrayQ(aMatrix) Then
        Let GetRow = Array()
        Exit Function
    End If
    
    Let nd = NumberOfDimensions(aMatrix)
    
    If nd = 1 Then
        If RowNumber = 1 Then
            Let GetRow = aMatrix
        Else
            Let GetRow = Array()
        End If
        
        Exit Function
    End If
    
    If RowNumber < 1 Or RowNumber > GetNumberOfRows(aMatrix) Then
        Let GetRow = Array()
        Exit Function
    End If
    
    ReDim TheResult(1 To GetNumberOfColumns(aMatrix))
    For i = 1 To GetNumberOfColumns(aMatrix)
        Let TheResult(i) = aMatrix(IIf(LBound(aMatrix, 1) = 0, RowNumber - 1, RowNumber), IIf(LBound(aMatrix, 2) = 0, i - 1, i))
    Next i
    
    Let GetRow = TheResult
End Function

Public Function GetSubRow(aMatrix As Variant, RowNumber As Long, StartColumn As Long, EndColumn As Long) As Variant
    Dim AnArray As Variant
    
    Let AnArray = GetRow(aMatrix, RowNumber)
    
    Let GetSubRow = GetSubArray(AnArray, StartColumn, EndColumn)
End Function

' Works on both 1D and 2D arrays, returning what makes sense (e.g. 1D arrays are interpreted as 1-row 2D arrays)
' Regardless of the array's indexing convention (e.g. start at 0 or 1), the user must refer to the first column as column 1.
' The result has the same number of dimensions as the argument
Public Function GetColumn(aMatrix As Variant, ColumnNumber As Long) As Variant
    Dim i As Long
    Dim TheResults() As Variant
    Dim NDims As Integer

    If EmptyArrayQ(aMatrix) Then
        Let GetColumn = Array()
        Exit Function
    End If
    
    Let NDims = NumberOfDimensions(aMatrix)
    
    If NDims = 1 Then
        If ColumnNumber > GetNumberOfColumns(aMatrix) Or ColumnNumber < 1 Then
            Let GetColumn = Array()
        Else
            Let GetColumn = Array(aMatrix(IIf(LBound(aMatrix) = 0, ColumnNumber - 1, ColumnNumber)))
        End If
        
        Exit Function
    End If
    
    If ColumnNumber < 1 Or ColumnNumber > GetNumberOfColumns(aMatrix) Then
        Let GetColumn = Array()
        Exit Function
    End If

    ReDim TheResults(GetNumberOfRows(aMatrix), 1 To 1)
    For i = IIf(NDims = 1, LBound(aMatrix), LBound(aMatrix, 1)) To IIf(NDims = 1, UBound(aMatrix), UBound(aMatrix, 1))
        Let TheResults(i, 1) = aMatrix(i, ColumnNumber)
    Next i
    
    Let GetColumn = TheResults
End Function

' Returns the requested subset of an column in a two-dimension array
' The result is returned as a one-dimensional array
Public Function GetSubColumn(aMatrix As Variant, ColumnNumber As Long, StartRow As Long, EndRow As Long) As Variant
    Dim AnArray As Variant
    
    Let AnArray = GetColumn(aMatrix, ColumnNumber)
    
    Let GetSubColumn = GetSubArray(AnArray, StartRow, EndRow)
End Function

' Returns the first element of the array if one dimensional or the first row if two dimensional
' Returns Empty if arg not an array
' Returns Empty for an empty array
Public Function First(AnArray As Variant) As Variant
    If Not IsArray(AnArray) Then
        Let First = Empty
    ElseIf NumberOfDimensions(AnArray) = 1 Then
        If EmptyArrayQ(AnArray) Then
            Let First = Empty
        Else
            Let First = AnArray(LBound(AnArray))
        End If
    ElseIf NumberOfDimensions(AnArray) = 2 Then
        Let First = ConvertTo1DArray(GetRow(AnArray, 1))
    Else
        Let First = AnArray
    End If
End Function

' Returns the first element of the array if one dimensional or the last row if two dimensional
' Returns Empty if arg not an array
' Returns Empty for an empty array
Public Function Last(AnArray As Variant) As Variant
    If Not IsArray(AnArray) Then
        Let Last = Empty
    ElseIf NumberOfDimensions(AnArray) = 1 Then
        If EmptyArrayQ(AnArray) Then
            Let Last = Empty
        Else
            Let Last = AnArray(UBound(AnArray))
        End If
    ElseIf NumberOfDimensions(AnArray) = 2 Then
        Let Last = ConvertTo1DArray(GetRow(AnArray, GetNumberOfRows(AnArray)))
    Else
        Let Last = AnArray
    End If
End Function

' Returns a sub-array with all but the last element of the array.
' If the array is empty or has only one element, this function returns array() (e.g. empty array)
Public Function Most(AnArray As Variant) As Variant
    If Not IsArray(AnArray) Then
        Let Most = AnArray
    ElseIf NumberOfDimensions(AnArray) = 1 Then
        If EmptyArrayQ(AnArray) Then
            Let Most = Array()
        ElseIf UBound(AnArray) <= LBound(AnArray) Then
            Let Most = Array()
        ElseIf UBound(AnArray) > LBound(AnArray) Then
            Let Most = GetSubArray(AnArray, LBound(AnArray), UBound(AnArray) - 1)
        End If
    ElseIf NumberOfDimensions(AnArray) = 2 Then
        If EmptyArrayQ(AnArray) Then
            Let Most = Array()
        ElseIf UBound(AnArray) <= LBound(AnArray) Then
            Let Most = Array()
        ElseIf UBound(AnArray) > LBound(AnArray) Then
            Let Most = GetSubMatrix(AnArray, LBound(AnArray, 1), UBound(AnArray, 1) - 1, LBound(AnArray, 2), UBound(AnArray, 2))
        End If
    Else
        Let Most = AnArray
    End If
End Function

' Returns a sub-array with all but the first element of the array.
' If the array is empty or has only one element, this function returns array() (e.g. empty array)
Public Function Rest(AnArray As Variant) As Variant
    If Not IsArray(AnArray) Then
        Let Rest = AnArray
    ElseIf NumberOfDimensions(AnArray) = 1 Then
        If EmptyArrayQ(AnArray) Then
            Let Rest = Array()
        ElseIf UBound(AnArray) <= LBound(AnArray) Then
            Let Rest = Array()
        ElseIf UBound(AnArray) > LBound(AnArray) Then
            Let Rest = GetSubArray(AnArray, LBound(AnArray) + 1, UBound(AnArray))
        End If
    ElseIf NumberOfDimensions(AnArray) = 2 Then
        If EmptyArrayQ(AnArray) Then
            Let Rest = Array()
        ElseIf UBound(AnArray) <= LBound(AnArray) Then
            Let Rest = Array()
        ElseIf UBound(AnArray) > LBound(AnArray) Then
            Let Rest = GetSubMatrix(AnArray, LBound(AnArray, 1) + 1, UBound(AnArray, 1), LBound(AnArray, 2), UBound(AnArray, 2))
        End If
    Else
        Let Rest = AnArray
    End If
End Function

' Appends a new element to the given array handles 1D and 2D arrays. Returns Null
' if the parameters are inconsistent. In the case of AnArray being a 2D array, AnElt
' must be a 1D array with the same number of colummns as AnArray for Append to make sense.
'
' If AnArray is a 1D or 2D array and AnElt is Null, the funtion returns AnArray unchanged.
'
' This works differently from simply using Stack2DArrays(AnArray, AnElt) since it can return
' something that is not a matrix.  Stack2DArrays ALWAYS returns a matrix
Public Function Append(AnArray As Variant, AnElt As Variant) As Variant
    Dim NewArray As Variant
    Dim AnArrayNumberOfDims As Integer
    
    Let AnArrayNumberOfDims = NumberOfDimensions(AnArray)
    
    If Not IsArray(AnArray) Or AnArrayNumberOfDims > 2 Or _
       (AnArrayNumberOfDims = 2 And GetNumberOfColumns(AnArray) <> GetNumberOfColumns(AnElt)) Then
        Let Append = Null
    
        Exit Function
    End If
    
    If IsNull(AnElt) Then
        Let Append = AnArray
        Exit Function
    End If
    
    ' If AnArray is 1D, then put the new element (whatever it may be) as the last element of a
    ' 1D array 1 longer than the original one.
    If NumberOfDimensions(AnArray) = 1 Then
        Let NewArray = AnArray
        ReDim Preserve NewArray(LBound(AnArray) To UBound(AnArray) + 1)
        Let NewArray(UBound(NewArray)) = AnElt
        
        Let Append = NewArray
        
        Exit Function
    End If
    
    ' If AnArray has two dims and the same number of columns as AnElt, then stack AnElt as the bottom
    ' of AnArray
    Let Append = Stack2DArrays(AnArray, AnElt)
End Function

' Stacks two arrays (may be 1 or 2-dimensional) on top of each other, provided they
' have the same number of columns. 1D arrays are allowed and interpreted as 1-row,
' 2D arrays. If either a or b is not an array, have dimensions > 2, or do not have
' the same number of columns, then this function returns an empty array (e.g. Array())
' The resulting arrays are indexed starting with 1
Public Function Stack2DArrays(ByVal a As Variant, ByVal b As Variant) As Variant
    Dim r As Long
    Dim c As Long
    Dim aprime As Variant
    Dim bprime As Variant
    Dim TheResult() As Variant
    
    If EmptyArrayQ(a) Or EmptyArrayQ(b) Or GetNumberOfColumns(a) <> GetNumberOfColumns(b) Then
        Let Stack2DArrays = Array()
        Exit Function
    End If
    
    ' If we have to 1D arrays of the same length, the we stack a on top of b
    If NumberOfDimensions(a) = 1 And NumberOfDimensions(b) = 1 Then
        ReDim TheResult(1 To 2, 1 To GetNumberOfColumns(a))
        
        For c = 1 To GetNumberOfColumns(a)
            Let TheResult(1, c) = a(IIf(LBound(a) = 0, c - 1, c))
            Let TheResult(2, c) = b(IIf(LBound(b) = 0, c - 1, c))
        Next c
    ElseIf NumberOfDimensions(a) = 1 And NumberOfDimensions(b) > 1 Then
        ReDim TheResult(1 To GetNumberOfRows(b) + 1, 1 To GetNumberOfColumns(b))
        
        For c = 1 To GetNumberOfColumns(a)
            Let TheResult(1, c) = a(IIf(LBound(a) = 0, c - 1, c))
        Next c
        
        For r = 1 To GetNumberOfRows(b)
            For c = 1 To GetNumberOfColumns(b)
                Let TheResult(1 + r, c) = b(IIf(LBound(b, 1) = 0, r - 1, r), IIf(LBound(b) = 0, c - 1, c))
            Next c
        Next r
    ElseIf NumberOfDimensions(a) > 1 And NumberOfDimensions(b) = 1 Then
        ReDim TheResult(1 To GetNumberOfRows(a) + 1, 1 To GetNumberOfColumns(a))
        
        For r = 1 To GetNumberOfRows(a) + 1
            For c = 1 To GetNumberOfColumns(a)
                If r < GetNumberOfRows(a) + 1 Then
                    Let TheResult(r, c) = a(IIf(LBound(a, 1) = 0, r - 1, r), IIf(LBound(a, 2) = 0, c - 1, c))
                Else
                    Let TheResult(GetNumberOfRows(a) + 1, c) = b(IIf(LBound(b) = 0, c - 1, c))
                End If
            Next c
        Next r
    Else
        ReDim TheResult(1 To GetNumberOfRows(a) + GetNumberOfRows(b), 1 To GetNumberOfColumns(b))
        For r = 1 To GetNumberOfRows(a)
            For c = 1 To GetNumberOfColumns(a)
                Let TheResult(r, c) = a(IIf(LBound(a, 1) = 0, r - 1, r), IIf(LBound(a, 2) = 0, c - 1, c))
            Next c
        Next r
    
        For r = 1 To GetNumberOfRows(b)
            For c = 1 To GetNumberOfColumns(b)
                Let TheResult(GetNumberOfRows(a) + r, c) = b(IIf(LBound(b, 1) = 0, r - 1, r), IIf(LBound(b, 2) = 0, c - 1, c))
            Next c
        Next r
    End If
    
    Let Stack2DArrays = TheResult
End Function

' Appends a new element to the given array. Handles arrays of dimension 1 and 2
' Returns Null if AnArray is not an array or it has more than two dims,
' or dim(AnArray) = 2 and AnArray and AnElt don't have the same number of columns.
' If AnElt is Null, the function returns AnArray unevaluated.
'
' This works differently from simply using Stack2DArrays(AnArray, AnElt)
' This one can give you something that is not a matrix.  Stack2DArrays ALWAYS returns
' a matrix
Public Function Prepend(AnArray As Variant, AnElt As Variant) As Variant
    Dim NewArray() As Variant
    Dim AnArrayNumberOfDims As Integer
    Dim i As Long
    
    Let AnArrayNumberOfDims = NumberOfDimensions(AnArray)
    
    If Not IsArray(AnArray) Or AnArrayNumberOfDims > 2 Or _
       (AnArrayNumberOfDims = 2 And GetNumberOfColumns(AnArray) <> GetNumberOfColumns(AnElt)) Then
        Let Prepend = Null
    
        Exit Function
    End If
    
    If IsNull(AnElt) Then
        Let Prepend = AnArray
        Exit Function
    End If
    
    ' If AnArray is 1D, then put the new element (whatever it may be) as the last element of a
    ' 1D array 1 longer than the original one.
    If NumberOfDimensions(AnArray) = 1 Then
        ReDim NewArray(LBound(AnArray) To UBound(AnArray) + 1)
        
        For i = LBound(AnArray) + 1 To UBound(AnArray) + 1
            Let NewArray(i) = AnArray(i - 1)
        Next i
        
        Let NewArray(LBound(AnArray)) = AnElt
        
        Let Prepend = NewArray
        
        Exit Function
    End If
    
    ' If AnArray has two dims and the same number of columns as AnElt, then stack AnElt as the bottom
    ' of AnArray
    Let Prepend = Stack2DArrays(AnElt, AnArray)
End Function

' Results a 2D array with the appropriate sub-matrix.  Every attempt is made to return a sensible sub-array when
' one or more optional parameters are missing.
' This allows for requesting all columns between two rows, all rows between two columns, and a fully specified, contiguous
' sub-matrix.
Public Function GetSubMatrix(aMatrix As Variant, Optional TopRowNumber As Variant, Optional BottomRowNumber As Variant, _
                             Optional BottomColumnNumber As Variant, Optional TopColumnNumber As Variant)
    If EmptyArrayQ(aMatrix) Then
        Let GetSubMatrix = Array()
        Exit Function
    End If
                             
    If IsMissing(TopRowNumber) And IsMissing(BottomRowNumber) And IsMissing(BottomColumnNumber) And IsMissing(TopColumnNumber) Then
        Let GetSubMatrix = aMatrix
        Exit Function
    End If
    
    If IsMissing(TopRowNumber) And IsMissing(BottomRowNumber) And Not IsMissing(BottomColumnNumber) And Not IsMissing(TopColumnNumber) Then
        Let GetSubMatrix = GetSubMatrixHelper(aMatrix, 1, GetNumberOfRows(aMatrix), CLng(BottomColumnNumber), CLng(TopColumnNumber))
        Exit Function
    End If
    
    If IsMissing(BottomColumnNumber) And IsMissing(TopColumnNumber) And Not IsMissing(TopRowNumber) And Not IsMissing(BottomRowNumber) Then
        Let GetSubMatrix = GetSubMatrixHelper(aMatrix, CLng(TopRowNumber), CLng(BottomRowNumber), 1, GetNumberOfColumns(aMatrix))
        Exit Function
    End If

    Let GetSubMatrix = GetSubMatrixHelper(aMatrix, CLng(TopRowNumber), CLng(BottomRowNumber), CLng(BottomColumnNumber), CLng(TopColumnNumber))
End Function

' Helper function for GetSubMatrix in this module.  This one ALWAYS requires all parameters.  GetSubMatrix allows for missing parameters, calling
' this one to handle all possible cases of missing parameters.
' This requires TheArray to be a 2D array, not a 1D array of 1D arrays
Private Function GetSubMatrixHelper(TheArray As Variant, TopRowNumber As Long, BottomRowNumber As Long, BottomColumnNumber As Long, TopColumnNumber As Long) As Variant
    Dim r As Long
    Dim c As Long
    Dim SubMatrix() As Variant
    
    If RowArrayQ(TheArray) Then
        ReDim SubMatrix(1, TopColumnNumber - BottomColumnNumber + LBound(TheArray))
        
        For c = BottomColumnNumber To TopColumnNumber
            Let SubMatrix(1, c - BottomColumnNumber + LBound(TheArray)) = TheArray(1, c)
        Next c
        
        Let GetSubMatrixHelper = SubMatrix
        Exit Function
    End If
    
    ReDim SubMatrix(IIf(LBound(TheArray, 1) = 0, 1, 0) To BottomRowNumber - TopRowNumber + IIf(LBound(TheArray, 1) = 0, 1, 0), _
                    IIf(LBound(TheArray, 2) = 0, 1, 0) To TopColumnNumber - BottomColumnNumber + IIf(LBound(TheArray, 2) = 0, 1, 0))
    For r = TopRowNumber To BottomRowNumber
        For c = BottomColumnNumber To TopColumnNumber
            Let SubMatrix(r - TopRowNumber + IIf(LBound(TheArray, 1) = 0, 1, 0), _
                          c - BottomColumnNumber + IIf(LBound(TheArray, 2) = 0, 1, 0)) = TheArray(r, c)
        Next c
    Next r
    
    Let GetSubMatrixHelper = SubMatrix
End Function

' Gets the sub-array between Indices StartIndex and EndIndex from the 1D array AnArray
' If the argument is not an array, then this function returns its argument unevaluated.
' If either StartIndex is below Lbound of the parameter or EndIndex is above UBound then,
' The LBound or UBound respectively is used.
'
' The function returns Null if the number of dimensions is larger than 1. An atomic expression
' is returned as a 1D array with the argument as its only argument. The function returns Null
' when EndIndex<StartIndex then.
'
' This function expects StartIndex and EndIndex to follow the indexing convention of AnArray.
' In other words, it respects 0 and 1 as the starting indices of an Array.  The resulting
' subarray is returned with the same indexing convention as the original array.
Public Function GetSubArray(AnArray As Variant, StartIndex As Long, EndIndex As Long) As Variant
    Dim i As Long
    Dim ReturnedArray() As Variant
    
    If Not IsArray(AnArray) Or EndIndex < StartIndex Then
        Let GetSubArray = Null
        Exit Function
    End If
    
    If NumberOfDimensions(AnArray) = 0 Then
        Let GetSubArray = Array(AnArray)
    ElseIf NumberOfDimensions(AnArray) = 1 Then
        Let StartIndex = Application.Max(StartIndex, LBound(AnArray))
        Let EndIndex = Application.Min(EndIndex, UBound(AnArray))
    
        ReDim ReturnedArray(LBound(AnArray) To EndIndex - StartIndex + LBound(AnArray))
    
        For i = LBound(AnArray) To EndIndex - StartIndex + LBound(AnArray)
            Let ReturnedArray(i) = AnArray(StartIndex + i - LBound(AnArray))
        Next i
        
        Let GetSubArray = ReturnedArray
    Else
        Let GetSubArray = AnArray
    End If
    
    If Not IsArray(GetSubArray) Then
        Let GetSubArray = Array(GetSubArray)
    End If
End Function

' aMatrix is a 1D or 2D array of values (not a range).
' Returns a 1-dimensional array with the unique subset of elements.
' If aMatrix has more than two dimensions, the function returns Null
Public Function UniqueSubset(aMatrix As Variant) As Variant
    Dim r As Long
    Dim c As Long
    Dim UniqueDict As Dictionary
    Dim Dimensionality As Integer
    
    ' Exit, returning an empty array if aMatrix is empty
    If EmptyArrayQ(aMatrix) Or IsEmpty(aMatrix) Then
        Let UniqueSubset = Array()
        Exit Function
    End If
    
    Set UniqueDict = New Dictionary
    Let Dimensionality = NumberOfDimensions(aMatrix)
    
    If Dimensionality = 0 Then
        Let UniqueSubset = aMatrix
        
        Exit Function
    ElseIf Dimensionality = 1 Then
        For r = LBound(aMatrix) To UBound(aMatrix)
            If Not UniqueDict.Exists(aMatrix(r)) Then
                Call UniqueDict.Add(Key:=aMatrix(r), Item:=1)
            End If
        Next r
    ElseIf Dimensionality = 2 Then
        For r = LBound(aMatrix, 1) To UBound(aMatrix, 1)
            For c = LBound(aMatrix, 2) To UBound(aMatrix, 2)
                If Not UniqueDict.Exists(aMatrix(r, c)) Then
                    Call UniqueDict.Add(Key:=aMatrix(r, c), Item:=1)
                End If
            Next c
        Next r
    Else
        Let UniqueSubset = Null
        
        Exit Function
    End If
    
    Let UniqueSubset = UniqueDict.Keys
End Function
' Returns the unique set of values contained in the two 1D arrays. The union of Set1 and Set2 is returned as a 1D array.
' Set1 and Set2 must have no empty rows of columns.
Public Function UnionOfSets(Set1 As Variant, Set2 As Variant) As Variant
    Dim FirstSet As Variant
    Dim SecondSet As Variant
    Dim CombinedSet As Dictionary
    Dim i As Long

    Set CombinedSet = New Dictionary
    
    If (EmptyArrayQ(Set1) Or IsEmpty(Set1)) And (EmptyArrayQ(Set2) Or IsEmpty(Set2)) Then
        Let UnionOfSets = Array()
        Exit Function
    ElseIf (EmptyArrayQ(Set1) Or IsEmpty(Set1)) And Not EmptyArrayQ(Set2) Then
        Let UnionOfSets = UniqueSubset(Set2)
        Exit Function
    ElseIf Not EmptyArrayQ(Set1) And (EmptyArrayQ(Set2) Or IsEmpty(Set2)) Then
        Let UnionOfSets = UniqueSubset(Set1)
        Exit Function
    End If

    ' Stack the two sets on top of each other in the worksheet TempComputation
    Let FirstSet = UniqueSubset(Set1)
    Let SecondSet = UniqueSubset(Set2)
    
    For i = LBound(FirstSet) To UBound(FirstSet)
        If Not CombinedSet.Exists(FirstSet(i)) Then
            Call CombinedSet.Add(Key:=FirstSet(i), Item:=i)
        End If
    Next i
    
    For i = LBound(SecondSet) To UBound(SecondSet)
        If Not CombinedSet.Exists(SecondSet(i)) Then
            Call CombinedSet.Add(Key:=SecondSet(i), Item:=i)
        End If
    Next i
    
    Let UnionOfSets = CombinedSet.Keys
End Function

' This function takes two parameters.  Each could be either a 1D array or a 2D array.
' This function returns the intersection of the two arrays.
Public Function IntersectionOfSets(Set1 As Variant, Set2 As Variant) As Variant
    Dim FirstDict As Dictionary
    Dim First1DSet As Variant
    Dim Second1DSet As Variant
    Dim IntersectionDict As Dictionary
    Dim i As Long
    
    ' Exit returning an empty array if either set is empty
    If EmptyArrayQ(Set1) Or IsEmpty(Set1) Or EmptyArrayQ(Set2) Or IsEmpty(Set2) Then
        Let IntersectionOfSets = Array()
        Exit Function
    End If
    
    ' Instantiate dictionaries
    Set FirstDict = New Dictionary
    Set IntersectionDict = New Dictionary
    
    ' Convert each set to a 1D array
    Let First1DSet = Stack2DArrayAs1DArray(Set1)
    Let Second1DSet = Stack2DArrayAs1DArray(Set2)
    
    ' Load a dictionary with the elements of the first set
    For i = LBound(First1DSet) To UBound(First1DSet)
        Call FirstDict.Add(Key:=First1DSet(i), Item:=1)
    Next i
    
    ' Store the elements of the second set that are in the first
    For i = LBound(Second1DSet) To UBound(Second1DSet)
        If FirstDict.Exists(Key:=Second1DSet(i)) Then
            Call IntersectionDict.Add(Key:=Second1DSet(i), Item:=1)
        End If
    Next i
    
    ' Return the intesection of the two sets
    Let IntersectionOfSets = IntersectionDict.Keys
End Function

' This function returns the complement of set B in A
' Both A and B are required to be 1D arrays
' If the complement is empty, this function returns an empty array (e.g. array())
Public Function ComplementOfSets(a As Variant, b As Variant) As Variant
    Dim BDict As Dictionary
    Dim ComplementDict As Dictionary
    Dim obj As Variant
    
    ' If a is an empty array, exit returning an empty array
    If EmptyArrayQ(a) Or IsEmpty(a) Then
        Let ComplementOfSets = Array()
        Exit Function
    End If
    
    If EmptyArrayQ(b) Or IsEmpty(b) Then
        Let ComplementOfSets = a
        Exit Function
    End If
    
    If NumberOfDimensions(a) < 1 Or NumberOfDimensions(b) < 1 Then
        Let ComplementOfSets = Array()
        
        Exit Function
    End If
    
    ' Instantiate dictionaries
    Set BDict = New Dictionary
    Set ComplementDict = New Dictionary
    
    ' Initialize ADict to get unique subset of ADict
    If GetArrayLength(b) > 0 Then
        For Each obj In b
            If Not BDict.Exists(Key:=obj) Then
                Call BDict.Add(Key:=obj, Item:=obj)
            End If
        Next
    End If
    
    ' Populate ComplementDict
    If GetArrayLength(a) > 0 Then
        For Each obj In a
            If Not BDict.Exists(Key:=obj) And Not ComplementDict.Exists(Key:=obj) Then
                Call ComplementDict.Add(Key:=obj, Item:=obj)
            End If
        Next
    End If
    
    ' Return complement as 1D array
    If ComplementDict.Count = 0 Then
        Let ComplementOfSets = Array()
    Else
        Let ComplementOfSets = ComplementDict.Keys
    End If
End Function

' This function takes a 1D or 2D array (NOT a range) and turns it into a 1D array with the same list of
' elements. It returns a 1D array.  Row 2 is appended to row 1.  Row 3 is then appended to that, etc.
Public Function Stack2DArrayAs1DArray(aMatrix As Variant) As Variant
    Dim TheResults() As Variant
    Dim var As Variant
    Dim j As Long

    If EmptyArrayQ(aMatrix) Or Not IsArray(aMatrix) Or NumberOfDimensions(aMatrix) > 2 Then
        Let Stack2DArrayAs1DArray = Array()
        Exit Function
    ElseIf NumberOfDimensions(aMatrix) = 0 Then
        Let Stack2DArrayAs1DArray = aMatrix
        Exit Function
    End If
    
    If NumberOfDimensions(aMatrix) = 1 Then
        ReDim TheResults(GetArrayLength(aMatrix))
    Else
        ReDim TheResults(GetNumberOfRows(aMatrix) * GetNumberOfColumns(aMatrix))
    End If
    
    Let j = 1
    For Each var In aMatrix
        Let TheResults(j) = var
        Let j = j + 1
    Next
    
    Let Stack2DArrayAs1DArray = TheResults
End Function

' Alias for Stack2DArrayAs1DArray
Public Function StackArrayAs1DArray(aMatrix As Variant) As Variant
    Let StackArrayAs1DArray = Stack2DArrayAs1DArray(aMatrix)
End Function

' This function dumps an array (1D or 2D) into worksheet TempComputation and then returns a reference to
' the underlying range.  Worksheet TempComputation is cleared before dump.  Dimensions are preserved.
' This means that an m x n array is dumped into an m x n range.  This function should not be used if
' leading single quotes (e.g "'") are part of the array's elements.
Public Function ToTemp(AnArray As Variant, Optional PreserveColumnTextFormats As Boolean = False) As Range
    Dim c As Integer
    Dim NumberOfRows As Long
    Dim NumberOfColumns As Integer
    Dim NumDimensions As Integer
    
    ' Exit if AnArray is empty
    If EmptyArrayQ(AnArray) Or IsEmpty(AnArray) Then
        Let ToTemp = Null
        Exit Function
    End If
    
    ' Clear used range
    Call TempComputation.Cells.ClearFormats
    Call TempComputation.UsedRange.ClearContents
        
    Let NumDimensions = NumberOfDimensions(AnArray)
        
    Let NumberOfRows = GetNumberOfRows(AnArray)
    Let NumberOfColumns = GetNumberOfColumns(AnArray)
    
    If PreserveColumnTextFormats Then
        If NumberOfDimensions(AnArray) = 0 Then
            Let TempComputation.Range("A1").NumberFormat = IIf(TypeName(AnArray) = "String", "@", "0")
            Let TempComputation.Range("A1").Value2 = AnArray
        Else
            ' Loop over the target columns, applying the format of the array's first column element to the entire,
            ' corresponding target range column
            For c = 0 To NumberOfColumns - 1
                If NumDimensions = 1 Then
                    Let TempComputation.Range("A1").Offset(0, c).Resize(NumberOfRows, 1).NumberFormat = IIf(TypeName(AnArray(c + 1)) = "String", "@", "0")
                Else
                    Let TempComputation.Range("A1").Offset(0, c).Resize(NumberOfRows, 1).NumberFormat = IIf(TypeName(AnArray(1, c + 1)) = "String", "@", "0")
                End If
            Next c
        End If
    End If
    
    If NumDimensions = 1 Then
        Let TempComputation.Range("A1").Resize(1, NumberOfColumns).Value2 = AnArray
        
        Set ToTemp = TempComputation.Range("A1").Resize(1, NumberOfColumns)
    Else
        Let TempComputation.Range("A1").Resize(NumberOfRows, NumberOfColumns).Value2 = AnArray

        Set ToTemp = TempComputation.Range("A1").Resize(NumberOfRows, NumberOfColumns)
    End If
End Function

' This function dumps an array (1D or 2D) into the worksheet with the range referenced by TopLeftCorner as the cell in the upper-left
' corner. It then returns a reference to the underlying range. Dimensions are preserved.  This means that an m x n array is dumped into
' an m x n range.
Public Function DumpInTempPositionWithoutFirstClearing(AnArray As Variant, TopLeftCorner As Range, Optional PreserveColumnTextFormats As Boolean = False) As Range
    Dim c As Integer
    Dim NumberOfRows As Long
    Dim NumberOfColumns As Integer
    Dim NumDimensions As Integer

    Let NumDimensions = NumberOfDimensions(AnArray)
    Let NumberOfRows = GetNumberOfRows(AnArray)
    Let NumberOfColumns = GetNumberOfColumns(AnArray)
    
    If PreserveColumnTextFormats Then
        If NumDimensions = 0 Then
            Let TempComputation.Range("A1").NumberFormat = IIf(TypeName(AnArray) = "String", "@", "0")
            Let TempComputation.Range("A1").Value2 = AnArray
        Else
            ' Loop over the target columns, applying the format of the array's first column element to
            ' the entire, corresponding target range column
            For c = 0 To NumberOfColumns - 1
                If NumDimensions = 1 Then
                    Let TopLeftCorner.Offset(0, c).Resize(NumberOfRows, 1).NumberFormat = IIf(TypeName(AnArray(c + LBound(AnArray))) = "String", "@", "0")
                Else
                    Let TopLeftCorner.Offset(0, c).Resize(NumberOfRows, 1).NumberFormat = IIf(TypeName(AnArray(1, c + LBound(AnArray, 2))) = "String", "@", "0")
                End If
            Next c
        End If
    End If

    If NumDimensions = 1 Then
        Let TopLeftCorner(1, 1).Resize(1, NumberOfColumns).Value2 = AnArray
        Set DumpInTempPositionWithoutFirstClearing = TopLeftCorner(1, 1).Resize(1, NumberOfColumns)
    Else
        Let TopLeftCorner(1, 1).Resize(NumberOfRows, NumberOfColumns).Value2 = AnArray
        Set DumpInTempPositionWithoutFirstClearing = TopLeftCorner(1, 1).Resize(NumberOfRows, NumberOfColumns)
    End If
End Function

' This is an alias for function DumpInTempPositionWithoutFirstClearing above
' This one handles empty arrays correctly, by doing nothing and returning Null
Public Function DumpInSheet(AnArray As Variant, TopLeftCorner As Range, Optional PreserveColumnTextFormats As Boolean = False) As Range
    If Not DimensionedQ(AnArray) Then
        Set DumpInSheet = Nothing
        Exit Function
    End If

    If EmptyArrayQ(AnArray) Or IsEmpty(AnArray) Then
        Set DumpInSheet = Nothing
        Exit Function
    End If
    
    If GetArrayLength(AnArray) = 0 Then
        Set DumpInSheet = Nothing
        Exit Function
    End If

    If PreserveColumnTextFormats Then
        Set DumpInSheet = DumpInTempPositionWithoutFirstClearing(AnArray, TopLeftCorner, PreserveColumnTextFormats:=PreserveColumnTextFormats)
    Else
        Set DumpInSheet = DumpInTempPositionWithoutFirstClearing(AnArray, TopLeftCorner)
    End If
End Function

' Applies the LAG operator to the power N to TheArray. TheArray has to be a 1D or 2D Array containing timeseries data.
' Timedimension is a parameter indicating whether the timeseries are organised from left to right (horizontal) or from
' top to bottom (vertical)
Public Function LAGN(TheArray As Variant, N As Variant, TimeDimension As String)
    If TimeDimension = "Vertical" Then
        LAGN = LagNRange(GetSubMatrix(TheArray, , UBound(TheArray, 1) - N), N, TimeDimension).Value2
    ElseIf TimeDimension = "Horizontal" Then
        LAGN = LagNRange(GetSubMatrix(TheArray, , , , UBound(TheArray, 2) - N), N, TimeDimension).Value2
    Else
        Exit Function
    End If
End Function

' Calculates the LN of a 1D or 2D Array
Public Function LnOfArray(TheArray As Variant) As Variant
    Dim tmpSht As Worksheet
    Dim RangeOfTheArray As Range

    Set RangeOfTheArray = ToTemp(TheArray)
    
    ' Set reference to the worksheet where TheArray has been dumped in
    Set tmpSht = RangeOfTheArray.Worksheet
    
    RangeOfTheArray.Offset(0, UBound(TheArray, 2) + 1) = Application.Ln(RangeOfTheArray)
    
    LnOfArray = RangeOfTheArray.Offset(0, UBound(TheArray, 2) + 1).Value2
End Function

' This function performs matrix element-wise division on two 0D, 1D, or 2D arrays.  Clearly, the two arrays
' must have the same dimensions.  The result is returned as an array of the same dimensions as those of
' the input.  This function throws an error if there is a division by 0
Public Function ElementwiseAddition(matrix1 As Variant, matrix2 As Variant) As Variant
    Dim TmpSheet As Worksheet
    Dim numRows As Long
    Dim numColumns As Long
    Dim r As Long ' for number of rows
    Dim c As Long ' for number of columns
    Dim rOffset1 As Long
    Dim cOffset1 As Long
    Dim rOffset2 As Long
    Dim cOffset2 As Long
    Dim TheResults() As Double
    Dim var As Variant
    
    ' Check parameter consistency
    If EmptyArrayQ(matrix1) Or EmptyArrayQ(matrix2) Then
        Let ElementwiseAddition = Null
    End If
    
    If Not (IsNumeric(matrix1) Or VectorQ(matrix1) Or MatrixQ(matrix1)) Then
        Let ElementwiseAddition = Null
        Exit Function
    End If
    
    If Not (IsNumeric(matrix2) Or VectorQ(matrix2) Or MatrixQ(matrix2)) Then
        Let ElementwiseAddition = Null
        Exit Function
    End If
    
    If MatrixQ(matrix1) And MatrixQ(matrix2) Then
        If GetNumberOfRows(matrix1) <> GetNumberOfRows(matrix2) Or _
           GetNumberOfColumns(matrix1) <> GetNumberOfColumns(matrix2) Then
           Let ElementwiseAddition = Null
           Exit Function
        End If
    End If
    
    If (RowVectorQ(matrix1) And MatrixQ(matrix2)) Or _
       (MatrixQ(matrix1) And RowVectorQ(matrix2)) Then
        If GetNumberOfColumns(matrix1) <> GetNumberOfColumns(matrix2) Then
           Let ElementwiseAddition = Null
           Exit Function
        End If
    End If
    
    If (ColumnVectorQ(matrix1) And MatrixQ(matrix2)) Or _
       (MatrixQ(matrix1) And ColumnVectorQ(matrix2)) Then
        If GetNumberOfRows(matrix1) <> GetNumberOfRows(matrix2) Then
           Let ElementwiseAddition = Null
           Exit Function
        End If
    End If
    
    ' Perform the calculations
    If IsNumeric(matrix1) And IsNumeric(matrix2) Then
        If CDbl(matrix2) <> 0 Then
            Let ElementwiseAddition = CDbl(matrix1) + CDbl(matrix2)
            Exit Function
        Else
            Let ElementwiseAddition = Null
            Exit Function
        End If
    ElseIf IsNumeric(matrix1) And RowVectorQ(matrix2) Then
        Let numColumns = GetNumberOfColumns(matrix2)
        
        ReDim TheResults(1 To numColumns)

        ' Compute the r and c offsets due to differences in array starts
        Let cOffset2 = IIf(LBound(matrix2, 1) = 0, 1, 0)
        
        For c = 1 To numColumns
            Let TheResults(c) = CDbl(matrix1) + CDbl(matrix2(c + cOffset2))
        Next c
    ElseIf IsNumeric(matrix1) And ColumnVectorQ(matrix2) Then
        Let numRows = GetNumberOfRows(matrix2)
        
        ReDim TheResults(1 To numRows, 1 To 1)

        ' Compute the r and c offsets due to differences in array starts
        Let rOffset2 = IIf(LBound(matrix2, 1) = 0, 1, 0)
        
        For r = 1 To numRows
            Let TheResults(r, 1) = CDbl(matrix1) + CDbl(matrix2(r + rOffset2, 1))
        Next r
    ElseIf IsNumeric(matrix1) And MatrixQ(matrix2) Then
        Let numRows = GetNumberOfRows(matrix2)
        Let numColumns = GetNumberOfColumns(matrix2)
        
        ReDim TheResults(1 To numRows, 1 To numColumns)
    
        ' Compute the r and c offsets due to differences in array starts
        Let rOffset2 = IIf(LBound(matrix2, 1) = 0, 1, 0)
        Let cOffset2 = IIf(LBound(matrix2, 2) = 0, 1, 0)
        
        For r = 1 To numRows
            For c = 1 To numColumns
                Let TheResults(r, c) = CDbl(matrix1) + CDbl(matrix2(r + rOffset2, c + cOffset2))
            Next c
        Next r
    ElseIf RowVectorQ(matrix1) And MatrixQ(matrix2) Then
        ' If the code gets here, we are adding two 2D matrices of the same size
        Let numRows = GetNumberOfRows(matrix2)
        Let numColumns = GetNumberOfColumns(matrix2)
        
        ReDim TheResults(1 To numRows, 1 To numColumns)
    
        ' Compute the r and c offsets due to differences in array starts
        Let rOffset2 = IIf(LBound(matrix2, 1) = 0, 1, 0)
        Let cOffset1 = IIf(LBound(matrix1) = 0, 1, 0)
        Let cOffset2 = IIf(LBound(matrix2, 2) = 0, 1, 0)
        
        For r = 1 To numRows
            For c = 1 To numColumns
                Let TheResults(r, c) = CDbl(matrix1(c + cOffset1)) + CDbl(matrix2(r + rOffset2, c + cOffset2))
            Next c
        Next r
    ElseIf ColumnVectorQ(matrix1) And MatrixQ(matrix2) Then
        Let numRows = GetNumberOfRows(matrix2)
        Let numColumns = GetNumberOfColumns(matrix2)
        
        ReDim TheResults(1 To numRows, 1 To numColumns)
    
        ' Compute the r and c offsets due to differences in array starts
        Let rOffset1 = IIf(LBound(matrix1, 1) = 0, 1, 0)
        Let rOffset2 = IIf(LBound(matrix2, 1) = 0, 1, 0)
        Let cOffset1 = IIf(LBound(matrix1, 2) = 0, 1, 0)
        Let cOffset2 = IIf(LBound(matrix2, 2) = 0, 1, 0)
        
        For r = 1 To numRows
            For c = 1 To numColumns
                Let TheResults(r, c) = CDbl(matrix1(r + rOffset1, 1 + cOffset1)) + CDbl(matrix2(r + rOffset2, c + cOffset2))
            Next c
        Next r
    ElseIf RowVectorQ(matrix1) And IsNumeric(matrix2) Then
        Let numColumns = GetNumberOfColumns(matrix1)
        
        ReDim TheResults(1 To numColumns)

        ' Compute the r and c offsets due to differences in array starts
        Let cOffset1 = IIf(LBound(matrix1) = 0, 1, 0)
        
        For c = 1 To numColumns
            Let TheResults(c) = CDbl(matrix1(c + cOffset1)) + CDbl(matrix2)
        Next c
    ElseIf ColumnVectorQ(matrix1) And IsNumeric(matrix2) Then
        Let numRows = GetNumberOfRows(matrix1)
        
        ReDim TheResults(1 To numRows, 1 To 1)

        ' Compute the r and c offsets due to differences in array starts
        Let rOffset1 = IIf(LBound(matrix1, 1) = 0, 1, 0)
        
        For r = 1 To numRows
            Let TheResults(r, 1) = CDbl(matrix1(r + rOffset1, 1)) + CDbl(matrix2)
        Next r
    ElseIf MatrixQ(matrix1) And IsNumeric(matrix2) Then
        Let numRows = GetNumberOfRows(matrix1)
        Let numColumns = GetNumberOfColumns(matrix1)
        
        ReDim TheResults(1 To numRows, 1 To numColumns)
    
        ' Compute the r and c offsets due to differences in array starts
        Let rOffset1 = IIf(LBound(matrix1, 1) = 0, 1, 0)
        Let cOffset1 = IIf(LBound(matrix1, 2) = 0, 1, 0)
        
        For r = 1 To numRows
            For c = 1 To numColumns
                Let TheResults(r, c) = CDbl(matrix1(r + rOffset1, c + cOffset1)) + CDbl(matrix2)
            Next c
        Next r
    ElseIf MatrixQ(matrix1) And RowVectorQ(matrix2) Then
        ' If the code gets here, we are adding two 2D matrices of the same size
        Let numRows = GetNumberOfRows(matrix1)
        Let numColumns = GetNumberOfColumns(matrix1)
        
        ReDim TheResults(1 To numRows, 1 To numColumns)
    
        ' Compute the r and c offsets due to differences in array starts
        Let rOffset1 = IIf(LBound(matrix1, 1) = 0, 1, 0)
        Let cOffset1 = IIf(LBound(matrix1, 2) = 0, 1, 0)
        Let cOffset2 = IIf(LBound(matrix2) = 0, 1, 0)
        
        For r = 1 To numRows
            For c = 1 To numColumns
                Let TheResults(r, c) = CDbl(matrix1(r + rOffset1, c + cOffset1)) + CDbl(matrix2(c + cOffset2))
            Next c
        Next r
    ElseIf ColumnVectorQ(matrix2) And MatrixQ(matrix1) Then
        Let numRows = GetNumberOfRows(matrix1)
        Let numColumns = GetNumberOfColumns(matrix1)
        
        ReDim TheResults(1 To numRows, 1 To numColumns)
    
        ' Compute the r and c offsets due to differences in array starts
        Let rOffset1 = IIf(LBound(matrix1, 1) = 0, 1, 0)
        Let rOffset2 = IIf(LBound(matrix2, 1) = 0, 1, 0)
        Let cOffset1 = IIf(LBound(matrix1, 2) = 0, 1, 0)
        Let cOffset2 = IIf(LBound(matrix2, 2) = 0, 1, 0)
        
        For r = 1 To numRows
            For c = 1 To numColumns
                Let TheResults(r, c) = CDbl(matrix1(r + rOffset1, c + cOffset1)) + CDbl(matrix2(r + rOffset2, 1 + cOffset2))
            Next c
        Next r
    ElseIf MatrixQ(matrix1) And MatrixQ(matrix2) Then
        ' If the code gets here, we are adding two 2D matrices of the same size
        Let numRows = GetNumberOfRows(matrix1)
        Let numColumns = GetNumberOfColumns(matrix1)
        
        ReDim TheResults(1 To numRows, 1 To numColumns)
    
        ' Compute the r and c offsets due to differences in array starts
        Let rOffset1 = IIf(LBound(matrix1, 1) = 0, 1, 0)
        Let rOffset2 = IIf(LBound(matrix2, 1) = 0, 1, 0)
        Let cOffset1 = IIf(LBound(matrix1, 2) = 0, 1, 0)
        Let cOffset2 = IIf(LBound(matrix2, 2) = 0, 1, 0)
        
        For r = 1 To numRows
            For c = 1 To numColumns
                Let TheResults(r, c) = CDbl(matrix1(r + rOffset1, c + cOffset1)) + CDbl(matrix2(r + rOffset2, c + cOffset2))
            Next c
        Next r
    Else
        Let ElementwiseAddition = Null
    End If
    
    ' Return the result
    Let ElementwiseAddition = TheResults
End Function



' This function performs matrix element-wise division on two 0D, 1D, or 2D arrays.  Clearly, the two arrays
' must have the same dimensions.  The result is returned as an array of the same dimensions as those of
' the input.  This function throws an error if there is a division by 0
Public Function ElementwiseMultiplication(matrix1 As Variant, matrix2 As Variant) As Variant
    Dim TmpSheet As Worksheet
    Dim numRows As Long
    Dim numColumns As Long
    Dim r As Long ' for number of rows
    Dim c As Long ' for number of columns
    Dim rOffset1 As Long
    Dim cOffset1 As Long
    Dim rOffset2 As Long
    Dim cOffset2 As Long
    Dim TheResults() As Double
    Dim var As Variant
    
    ' Check parameter consistency
    If EmptyArrayQ(matrix1) Or EmptyArrayQ(matrix2) Then
        Let ElementwiseMultiplication = Null
    End If
    
    If Not (IsNumeric(matrix1) Or VectorQ(matrix1) Or MatrixQ(matrix1)) Then
        Let ElementwiseMultiplication = Null
        Exit Function
    End If
    
    If Not (IsNumeric(matrix2) Or VectorQ(matrix2) Or MatrixQ(matrix2)) Then
        Let ElementwiseMultiplication = Null
        Exit Function
    End If
    
    If MatrixQ(matrix1) And MatrixQ(matrix2) Then
        If GetNumberOfRows(matrix1) <> GetNumberOfRows(matrix2) Or _
           GetNumberOfColumns(matrix1) <> GetNumberOfColumns(matrix2) Then
           Let ElementwiseMultiplication = Null
           Exit Function
        End If
    End If
    
    If (RowVectorQ(matrix1) And MatrixQ(matrix2)) Or _
       (MatrixQ(matrix1) And RowVectorQ(matrix2)) Then
        If GetNumberOfColumns(matrix1) <> GetNumberOfColumns(matrix2) Then
           Let ElementwiseMultiplication = Null
           Exit Function
        End If
    End If
    
    If (ColumnVectorQ(matrix1) And MatrixQ(matrix2)) Or _
       (MatrixQ(matrix1) And ColumnVectorQ(matrix2)) Then
        If GetNumberOfRows(matrix1) <> GetNumberOfRows(matrix2) Then
           Let ElementwiseMultiplication = Null
           Exit Function
        End If
    End If
    
    ' Perform the calculations
    If IsNumeric(matrix1) And IsNumeric(matrix2) Then
        If CDbl(matrix2) <> 0 Then
            Let ElementwiseMultiplication = CDbl(matrix1) * CDbl(matrix2)
            Exit Function
        Else
            Let ElementwiseMultiplication = Null
            Exit Function
        End If
    ElseIf IsNumeric(matrix1) And RowVectorQ(matrix2) Then
        Let numColumns = GetNumberOfColumns(matrix2)
        
        ReDim TheResults(1 To numColumns)

        ' Compute the r and c offsets due to differences in array starts
        Let cOffset2 = IIf(LBound(matrix2, 1) = 0, 1, 0)
        
        For c = 1 To numColumns
            Let TheResults(c) = CDbl(matrix1) * CDbl(matrix2(c + cOffset2))
        Next c
    ElseIf IsNumeric(matrix1) And ColumnVectorQ(matrix2) Then
        Let numRows = GetNumberOfRows(matrix2)
        
        ReDim TheResults(1 To numRows, 1 To 1)

        ' Compute the r and c offsets due to differences in array starts
        Let rOffset2 = IIf(LBound(matrix2, 1) = 0, 1, 0)
        
        For r = 1 To numRows
            Let TheResults(r, 1) = CDbl(matrix1) * CDbl(matrix2(r + rOffset2, 1))
        Next r
    ElseIf IsNumeric(matrix1) And MatrixQ(matrix2) Then
        Let numRows = GetNumberOfRows(matrix2)
        Let numColumns = GetNumberOfColumns(matrix2)
        
        ReDim TheResults(1 To numRows, 1 To numColumns)
    
        ' Compute the r and c offsets due to differences in array starts
        Let rOffset2 = IIf(LBound(matrix2, 1) = 0, 1, 0)
        Let cOffset2 = IIf(LBound(matrix2, 2) = 0, 1, 0)
        
        For r = 1 To numRows
            For c = 1 To numColumns
                Let TheResults(r, c) = CDbl(matrix1) * CDbl(matrix2(r + rOffset2, c + cOffset2))
            Next c
        Next r
    ElseIf RowVectorQ(matrix1) And MatrixQ(matrix2) Then
        ' If the code gets here, we are adding two 2D matrices of the same size
        Let numRows = GetNumberOfRows(matrix2)
        Let numColumns = GetNumberOfColumns(matrix2)
        
        ReDim TheResults(1 To numRows, 1 To numColumns)
    
        ' Compute the r and c offsets due to differences in array starts
        Let rOffset2 = IIf(LBound(matrix2, 1) = 0, 1, 0)
        Let cOffset1 = IIf(LBound(matrix1) = 0, 1, 0)
        Let cOffset2 = IIf(LBound(matrix2, 2) = 0, 1, 0)
        
        For r = 1 To numRows
            For c = 1 To numColumns
                Let TheResults(r, c) = CDbl(matrix1(c + cOffset1)) * CDbl(matrix2(r + rOffset2, c + cOffset2))
            Next c
        Next r
    ElseIf ColumnVectorQ(matrix1) And MatrixQ(matrix2) Then
        Let numRows = GetNumberOfRows(matrix2)
        Let numColumns = GetNumberOfColumns(matrix2)
        
        ReDim TheResults(1 To numRows, 1 To numColumns)
    
        ' Compute the r and c offsets due to differences in array starts
        Let rOffset1 = IIf(LBound(matrix1, 1) = 0, 1, 0)
        Let rOffset2 = IIf(LBound(matrix2, 1) = 0, 1, 0)
        Let cOffset1 = IIf(LBound(matrix1, 2) = 0, 1, 0)
        Let cOffset2 = IIf(LBound(matrix2, 2) = 0, 1, 0)
        
        For r = 1 To numRows
            For c = 1 To numColumns
                Let TheResults(r, c) = CDbl(matrix1(r + rOffset1, 1 + cOffset1)) * CDbl(matrix2(r + rOffset2, c + cOffset2))
            Next c
        Next r
    ElseIf RowVectorQ(matrix1) And IsNumeric(matrix2) Then
        Let numColumns = GetNumberOfColumns(matrix1)
        
        ReDim TheResults(1 To numColumns)

        ' Compute the r and c offsets due to differences in array starts
        Let cOffset1 = IIf(LBound(matrix1) = 0, 1, 0)
        
        For c = 1 To numColumns
            Let TheResults(c) = CDbl(matrix1(c + cOffset1)) * CDbl(matrix2)
        Next c
    ElseIf ColumnVectorQ(matrix1) And IsNumeric(matrix2) Then
        Let numRows = GetNumberOfRows(matrix1)
        
        ReDim TheResults(1 To numRows, 1 To 1)

        ' Compute the r and c offsets due to differences in array starts
        Let rOffset1 = IIf(LBound(matrix1, 1) = 0, 1, 0)
        
        For r = 1 To numRows
            Let TheResults(r, 1) = CDbl(matrix1(r + rOffset1, 1)) * CDbl(matrix2)
        Next r
    ElseIf MatrixQ(matrix1) And IsNumeric(matrix2) Then
        Let numRows = GetNumberOfRows(matrix1)
        Let numColumns = GetNumberOfColumns(matrix1)
        
        ReDim TheResults(1 To numRows, 1 To numColumns)
    
        ' Compute the r and c offsets due to differences in array starts
        Let rOffset1 = IIf(LBound(matrix1, 1) = 0, 1, 0)
        Let cOffset1 = IIf(LBound(matrix1, 2) = 0, 1, 0)
        
        For r = 1 To numRows
            For c = 1 To numColumns
                Let TheResults(r, c) = CDbl(matrix1(r + rOffset1, c + cOffset1)) * CDbl(matrix2)
            Next c
        Next r
    ElseIf MatrixQ(matrix1) And RowVectorQ(matrix2) Then
        ' If the code gets here, we are adding two 2D matrices of the same size
        Let numRows = GetNumberOfRows(matrix1)
        Let numColumns = GetNumberOfColumns(matrix1)
        
        ReDim TheResults(1 To numRows, 1 To numColumns)
    
        ' Compute the r and c offsets due to differences in array starts
        Let rOffset1 = IIf(LBound(matrix1, 1) = 0, 1, 0)
        Let cOffset1 = IIf(LBound(matrix1, 2) = 0, 1, 0)
        Let cOffset2 = IIf(LBound(matrix2) = 0, 1, 0)
        
        For r = 1 To numRows
            For c = 1 To numColumns
                Let TheResults(r, c) = CDbl(matrix1(r + rOffset1, c + cOffset1)) * CDbl(matrix2(c + cOffset2))
            Next c
        Next r
    ElseIf ColumnVectorQ(matrix2) And MatrixQ(matrix1) Then
        Let numRows = GetNumberOfRows(matrix1)
        Let numColumns = GetNumberOfColumns(matrix1)
        
        ReDim TheResults(1 To numRows, 1 To numColumns)
    
        ' Compute the r and c offsets due to differences in array starts
        Let rOffset1 = IIf(LBound(matrix1, 1) = 0, 1, 0)
        Let rOffset2 = IIf(LBound(matrix2, 1) = 0, 1, 0)
        Let cOffset1 = IIf(LBound(matrix1, 2) = 0, 1, 0)
        Let cOffset2 = IIf(LBound(matrix2, 2) = 0, 1, 0)
        
        For r = 1 To numRows
            For c = 1 To numColumns
                Let TheResults(r, c) = CDbl(matrix1(r + rOffset1, c + cOffset1)) * CDbl(matrix2(r + rOffset2, 1 + cOffset2))
            Next c
        Next r
    ElseIf MatrixQ(matrix1) And MatrixQ(matrix2) Then
        ' If the code gets here, we are adding two 2D matrices of the same size
        Let numRows = GetNumberOfRows(matrix1)
        Let numColumns = GetNumberOfColumns(matrix1)
        
        ReDim TheResults(1 To numRows, 1 To numColumns)
    
        ' Compute the r and c offsets due to differences in array starts
        Let rOffset1 = IIf(LBound(matrix1, 1) = 0, 1, 0)
        Let rOffset2 = IIf(LBound(matrix2, 1) = 0, 1, 0)
        Let cOffset1 = IIf(LBound(matrix1, 2) = 0, 1, 0)
        Let cOffset2 = IIf(LBound(matrix2, 2) = 0, 1, 0)
        
        For r = 1 To numRows
            For c = 1 To numColumns
                Let TheResults(r, c) = CDbl(matrix1(r + rOffset1, c + cOffset1)) * CDbl(matrix2(r + rOffset2, c + cOffset2))
            Next c
        Next r
    Else
        Let ElementwiseMultiplication = Null
    End If
    
    ' Return the result
    Let ElementwiseMultiplication = TheResults
End Function

' This function performs matrix element-wise division on two 0D, 1D, or 2D arrays.  Clearly, the two arrays
' must have the same dimensions.  The result is returned as an array of the same dimensions as those of
' the input.  This function throws an error if there is a division by 0
Public Function ElementWiseDivision(matrix1 As Variant, matrix2 As Variant) As Variant
    Dim TmpSheet As Worksheet
    Dim numRows As Long
    Dim numColumns As Long
    Dim r As Long ' for number of rows
    Dim c As Long ' for number of columns
    Dim rOffset1 As Long
    Dim cOffset1 As Long
    Dim rOffset2 As Long
    Dim cOffset2 As Long
    Dim TheResults() As Double
    Dim var As Variant
    
    ' Check parameter consistency
    If EmptyArrayQ(matrix1) Or EmptyArrayQ(matrix2) Then
        Let ElementWiseDivision = Null
    End If
    
    If Not (IsNumeric(matrix1) Or VectorQ(matrix1) Or MatrixQ(matrix1)) Then
        Let ElementWiseDivision = Null
        Exit Function
    End If
    
    If Not (IsNumeric(matrix2) Or VectorQ(matrix2) Or MatrixQ(matrix2)) Then
        Let ElementWiseDivision = Null
        Exit Function
    End If
    
    If MatrixQ(matrix1) And MatrixQ(matrix2) Then
        If GetNumberOfRows(matrix1) <> GetNumberOfRows(matrix2) Or _
           GetNumberOfColumns(matrix1) <> GetNumberOfColumns(matrix2) Then
           Let ElementWiseDivision = Null
           Exit Function
        End If
    End If
    
    If (RowVectorQ(matrix1) And MatrixQ(matrix2)) Or _
       (MatrixQ(matrix1) And RowVectorQ(matrix2)) Then
        If GetNumberOfColumns(matrix1) <> GetNumberOfColumns(matrix2) Then
           Let ElementWiseDivision = Null
           Exit Function
        End If
    End If
    
    If (ColumnVectorQ(matrix1) And MatrixQ(matrix2)) Or _
       (MatrixQ(matrix1) And ColumnVectorQ(matrix2)) Then
        If GetNumberOfRows(matrix1) <> GetNumberOfRows(matrix2) Then
           Let ElementWiseDivision = Null
           Exit Function
        End If
    End If
    
    ' Perform the calculations
    If IsNumeric(matrix1) And IsNumeric(matrix2) Then
        If CDbl(matrix2) <> 0 Then
            Let ElementWiseDivision = CDbl(matrix1) / CDbl(matrix2)
            Exit Function
        Else
            Let ElementWiseDivision = Null
            Exit Function
        End If
    ElseIf IsNumeric(matrix1) And RowVectorQ(matrix2) Then
        Let numColumns = GetNumberOfColumns(matrix2)
        
        ReDim TheResults(1 To numColumns)

        ' Compute the r and c offsets due to differences in array starts
        Let cOffset2 = IIf(LBound(matrix2, 1) = 0, 1, 0)
        
        For c = 1 To numColumns
            Let TheResults(c) = CDbl(matrix1) / CDbl(matrix2(c + cOffset2))
        Next c
    ElseIf IsNumeric(matrix1) And ColumnVectorQ(matrix2) Then
        Let numRows = GetNumberOfRows(matrix2)
        
        ReDim TheResults(1 To numRows, 1 To 1)

        ' Compute the r and c offsets due to differences in array starts
        Let rOffset2 = IIf(LBound(matrix2, 1) = 0, 1, 0)
        
        For r = 1 To numRows
            Let TheResults(r, 1) = CDbl(matrix1) / CDbl(matrix2(r + rOffset2, 1))
        Next r
    ElseIf IsNumeric(matrix1) And MatrixQ(matrix2) Then
        Let numRows = GetNumberOfRows(matrix2)
        Let numColumns = GetNumberOfColumns(matrix2)
        
        ReDim TheResults(1 To numRows, 1 To numColumns)
    
        ' Compute the r and c offsets due to differences in array starts
        Let rOffset2 = IIf(LBound(matrix2, 1) = 0, 1, 0)
        Let cOffset2 = IIf(LBound(matrix2, 2) = 0, 1, 0)
        
        For r = 1 To numRows
            For c = 1 To numColumns
                Let TheResults(r, c) = CDbl(matrix1) / CDbl(matrix2(r + rOffset2, c + cOffset2))
            Next c
        Next r
    ElseIf RowVectorQ(matrix1) And MatrixQ(matrix2) Then
        ' If the code gets here, we are adding two 2D matrices of the same size
        Let numRows = GetNumberOfRows(matrix2)
        Let numColumns = GetNumberOfColumns(matrix2)
        
        ReDim TheResults(1 To numRows, 1 To numColumns)
    
        ' Compute the r and c offsets due to differences in array starts
        Let rOffset2 = IIf(LBound(matrix2, 1) = 0, 1, 0)
        Let cOffset1 = IIf(LBound(matrix1) = 0, 1, 0)
        Let cOffset2 = IIf(LBound(matrix2, 2) = 0, 1, 0)
        
        For r = 1 To numRows
            For c = 1 To numColumns
                Let TheResults(r, c) = CDbl(matrix1(c + cOffset1)) / CDbl(matrix2(r + rOffset2, c + cOffset2))
            Next c
        Next r
    ElseIf ColumnVectorQ(matrix1) And MatrixQ(matrix2) Then
        Let numRows = GetNumberOfRows(matrix2)
        Let numColumns = GetNumberOfColumns(matrix2)
        
        ReDim TheResults(1 To numRows, 1 To numColumns)
    
        ' Compute the r and c offsets due to differences in array starts
        Let rOffset1 = IIf(LBound(matrix1, 1) = 0, 1, 0)
        Let rOffset2 = IIf(LBound(matrix2, 1) = 0, 1, 0)
        Let cOffset1 = IIf(LBound(matrix1, 2) = 0, 1, 0)
        Let cOffset2 = IIf(LBound(matrix2, 2) = 0, 1, 0)
        
        For r = 1 To numRows
            For c = 1 To numColumns
                Let TheResults(r, c) = CDbl(matrix1(r + rOffset1, 1 + cOffset1)) / CDbl(matrix2(r + rOffset2, c + cOffset2))
            Next c
        Next r
    ElseIf RowVectorQ(matrix1) And IsNumeric(matrix2) Then
        Let numColumns = GetNumberOfColumns(matrix1)
        
        ReDim TheResults(1 To numColumns)

        ' Compute the r and c offsets due to differences in array starts
        Let cOffset1 = IIf(LBound(matrix1) = 0, 1, 0)
        
        For c = 1 To numColumns
            Let TheResults(c) = CDbl(matrix1(c + cOffset1)) / CDbl(matrix2)
        Next c
    ElseIf ColumnVectorQ(matrix1) And IsNumeric(matrix2) Then
        Let numRows = GetNumberOfRows(matrix1)
        
        ReDim TheResults(1 To numRows, 1 To 1)

        ' Compute the r and c offsets due to differences in array starts
        Let rOffset1 = IIf(LBound(matrix1, 1) = 0, 1, 0)
        
        For r = 1 To numRows
            Let TheResults(r, 1) = CDbl(matrix1(r + rOffset1, 1)) / CDbl(matrix2)
        Next r
    ElseIf MatrixQ(matrix1) And IsNumeric(matrix2) Then
        Let numRows = GetNumberOfRows(matrix1)
        Let numColumns = GetNumberOfColumns(matrix1)
        
        ReDim TheResults(1 To numRows, 1 To numColumns)
    
        ' Compute the r and c offsets due to differences in array starts
        Let rOffset1 = IIf(LBound(matrix1, 1) = 0, 1, 0)
        Let cOffset1 = IIf(LBound(matrix1, 2) = 0, 1, 0)
        
        For r = 1 To numRows
            For c = 1 To numColumns
                Let TheResults(r, c) = CDbl(matrix1(r + rOffset1, c + cOffset1)) / CDbl(matrix2)
            Next c
        Next r
    ElseIf MatrixQ(matrix1) And RowVectorQ(matrix2) Then
        ' If the code gets here, we are adding two 2D matrices of the same size
        Let numRows = GetNumberOfRows(matrix1)
        Let numColumns = GetNumberOfColumns(matrix1)
        
        ReDim TheResults(1 To numRows, 1 To numColumns)
    
        ' Compute the r and c offsets due to differences in array starts
        Let rOffset1 = IIf(LBound(matrix1, 1) = 0, 1, 0)
        Let cOffset1 = IIf(LBound(matrix1, 2) = 0, 1, 0)
        Let cOffset2 = IIf(LBound(matrix2) = 0, 1, 0)
        
        For r = 1 To numRows
            For c = 1 To numColumns
                Let TheResults(r, c) = CDbl(matrix1(r + rOffset1, c + cOffset1)) / CDbl(matrix2(c + cOffset2))
            Next c
        Next r
    ElseIf ColumnVectorQ(matrix2) And MatrixQ(matrix1) Then
        Let numRows = GetNumberOfRows(matrix1)
        Let numColumns = GetNumberOfColumns(matrix1)
        
        ReDim TheResults(1 To numRows, 1 To numColumns)
    
        ' Compute the r and c offsets due to differences in array starts
        Let rOffset1 = IIf(LBound(matrix1, 1) = 0, 1, 0)
        Let rOffset2 = IIf(LBound(matrix2, 1) = 0, 1, 0)
        Let cOffset1 = IIf(LBound(matrix1, 2) = 0, 1, 0)
        Let cOffset2 = IIf(LBound(matrix2, 2) = 0, 1, 0)
        
        For r = 1 To numRows
            For c = 1 To numColumns
                Let TheResults(r, c) = CDbl(matrix1(r + rOffset1, c + cOffset1)) / CDbl(matrix2(r + rOffset2, 1 + cOffset2))
            Next c
        Next r
    ElseIf MatrixQ(matrix1) And MatrixQ(matrix2) Then
        ' If the code gets here, we are adding two 2D matrices of the same size
        Let numRows = GetNumberOfRows(matrix1)
        Let numColumns = GetNumberOfColumns(matrix1)
        
        ReDim TheResults(1 To numRows, 1 To numColumns)
    
        ' Compute the r and c offsets due to differences in array starts
        Let rOffset1 = IIf(LBound(matrix1, 1) = 0, 1, 0)
        Let rOffset2 = IIf(LBound(matrix2, 1) = 0, 1, 0)
        Let cOffset1 = IIf(LBound(matrix1, 2) = 0, 1, 0)
        Let cOffset2 = IIf(LBound(matrix2, 2) = 0, 1, 0)
        
        For r = 1 To numRows
            For c = 1 To numColumns
                Let TheResults(r, c) = CDbl(matrix1(r + rOffset1, c + cOffset1)) / CDbl(matrix2(r + rOffset2, c + cOffset2))
            Next c
        Next r
    Else
        Let ElementWiseDivision = Null
    End If
    
    ' Return the result
    Let ElementWiseDivision = TheResults
End Function

'***HERE explain this a bit clearer
' This function dumps an array (1D or 2D) into worksheet TempComputation, shifts the array by N rows or N columns depending
' on the value of TimeDimension, and then returns a reference to the underlying range.  Dimensions are preserved.
' This means that an m x n array is dumped into an m x n range. A 1D array is dumped vertically to allow for bigger arrays since Excel 2010
' has a horizontal maximum under 17K columns.
Public Function LagNRange(AnArray As Variant, N As Variant, TimeDimension As String) As Range
    Dim tmpSht As Worksheet
    
    ' Set reference to worksheet TempComputation
    Set tmpSht = ActiveWorkbook.Worksheets("TempComputation")
    
    ' Clear used range
    tmpSht.UsedRange.ClearContents
    
    If TimeDimension = "Vertical" Then
        If NumberOfDimensions(AnArray) = 1 Then
            Let tmpSht.Range("A1").Offset(N, 0).Resize(UBound(AnArray)).Value2 = Application.Transpose(AnArray)
            Set LagNRange = tmpSht.Range("A1").Resize(UBound(AnArray) + N, 1)
        Else
            Let tmpSht.Range("A1").Offset(N, 0).Resize(UBound(AnArray, 1), UBound(AnArray, 2)).Value2 = AnArray
            Set LagNRange = tmpSht.Range("A1").Resize(UBound(AnArray, 1) + N, UBound(AnArray, 2))
        End If
    ElseIf TimeDimension = "Horizontal" Then
        If NumberOfDimensions(AnArray) = 1 Then
            Let tmpSht.Range("A1").Offset(0, N).Resize(UBound(AnArray)).Value2 = Application.Transpose(AnArray)
            Set LagNRange = tmpSht.Range("A1").Resize(UBound(AnArray) + N, 1)
        Else
            Let tmpSht.Range("A1").Offset(0, N).Resize(UBound(AnArray, 1), UBound(AnArray, 2)).Value2 = AnArray
            Set LagNRange = tmpSht.Range("A1").Resize(UBound(AnArray, 1), UBound(AnArray, 2) + N)
        End If
    Else
        Exit Function:
    End If
End Function

'This function returns an Array of NrOfRows rows and NrOfColumns and all elements are equal to -1
Public Function MatrixOfMinusOnes(NrOfRows As Long, NrOfColumns As Long) As Variant
    Dim tmpSht As Worksheet
    
    ' Set reference to worksheet TempComputation
    Set tmpSht = Worksheets("TempComputation")
    
    ' Clear used range
    tmpSht.UsedRange.ClearContents
    
    Let tmpSht.Range("A1").Resize(NrOfRows, NrOfColumns).Value2 = -1
    Let MatrixOfMinusOnes = tmpSht.Range("A1").Resize(NrOfRows, NrOfColumns)
End Function

'This function returns an Array of NrOfRows rows and NrOfColumns columns and all elements are equal to 1
Public Function Ones(NrOfRows As Long, NrOfColumns As Long) As Variant
    Dim tmpSht As Worksheet
    
    ' Set reference to worksheet TempComputation
    Set tmpSht = ThisWorkbook.Worksheets("TempComputation")
    
    ' Clear used range
    Call tmpSht.UsedRange.ClearContents
    
    Let tmpSht.Range("A1").Resize(NrOfRows, NrOfColumns).Value2 = 1
    Let Ones = tmpSht.Range("A1").Resize(NrOfRows, NrOfColumns)
End Function

'This function returns an Array of NrOfRows rows and NrOfColumns columns and all elements are equal to 1
Public Function Zeroes(NrOfRows As Long, NrOfColumns As Long) As Variant
    Dim tmpSht As Worksheet
    
    ' Set reference to worksheet TempComputation
    Set tmpSht = ThisWorkbook.Worksheets("TempComputation")
    
    ' Clear used range
    Call tmpSht.UsedRange.ClearContents
    
    Let tmpSht.Range("A1").Resize(NrOfRows, NrOfColumns).Value2 = 0
    Let Zeroes = tmpSht.Range("A1").Resize(NrOfRows, NrOfColumns)
End Function

' This function sorts a 1D array in the ascending or descending order, assuming no header.
' This returns a 1D array with the sorted values
Public Function Sort1DArray(MyArray As Variant) As Variant
    Dim TheRange As Range
    Dim TmpSheet As Worksheet
    Dim i As Integer
    Dim LastRowNumber As Long
    
    ' Set pointer to temp sheet and clear its used range
    Set TmpSheet = ThisWorkbook.Worksheets("TempComputation")
    Call TmpSheet.UsedRange.ClearContents
    
    ' Dump array in temp sheet
    Call ToTemp(Application.Transpose(ConvertTo1DArray(MyArray)))
    
    ' Find the last row used
    Let LastRowNumber = TmpSheet.Range("A1").Offset(TmpSheet.Rows.Count - 1, 0).End(xlUp).row

    ' Set range pointer to the data we just dumped in the temp sheet
    Set TheRange = TmpSheet.Range("A1").Resize(LastRowNumber, 1)

    ' Clear any previous sorting criteria
    TmpSheet.Sort.SortFields.Clear

    ' Add all the sorting criteria
    TheRange.Worksheet.Sort.SortFields.Add _
        Key:=TheRange, _
        SortOn:=xlSortOnValues, _
        Order:=xlAscending, _
        DataOption:=xlSortNormal
    
    ' Execute the sort
    With TheRange.Worksheet.Sort
        .SetRange TheRange
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    ' Extract horizontally reversed matrix and set them to return when the function exits
    Let Sort1DArray = ConvertTo1DArray(TheRange.Value2)
End Function

' Returns a 2D array of uniformly distributed random numbers between 0 and 1.
Public Function RandomMatrix(NRows As Long, NColumns As Long) As Variant
    Dim ReturnMatrix() As Double

    If NRows < 0 Or NColumns < 0 Then
        Let RandomMatrix = Empty
        Exit Function
    End If
    
    If NRows = 0 Or NColumns = 0 Then
        Let RandomMatrix = Empty
        Exit Function
    End If
    
    ' Allocate a return
    ReDim ReturnMatrix(NRows, NColumns)
    
    ' Clear TmpSht
    Call ThisWorkbook.Worksheets("TempComputation").UsedRange.ClearContents
    
    ' Create random matrix
    Let Application.Calculation = xlManual
    Let ThisWorkbook.Worksheets("TempComputation").Cells(1, 1).Resize(NRows, NColumns).Formula = "=Rand()"
    Call Application.Calculate
    Let Application.Calculation = xlAutomatic
    
    ' Return random matrix
    Let RandomMatrix = ThisWorkbook.Worksheets("TempComputation").Range("A1").CurrentRegion.Value2
    
    ' Clear TmpSht
    Call ThisWorkbook.Worksheets("TempComputation").UsedRange.ClearContents
End Function

' Alias for GetArrayLength
Public Function Length(AnArray As Variant) As Long
    Let Length = GetArrayLength(AnArray)
End Function

' Alias for GetNumberOfRows
Public Function NumberOfRows(AnArray As Variant) As Long
    Let NumberOfRows = GetNumberOfRows(AnArray)
End Function

' Alias for GetNumberOfColumns
Public Function NumberOfColumns(AnArray As Variant) As Long
    Let NumberOfColumns = GetNumberOfColumns(AnArray)
End Function

' Returns the number of elements in the first dimension
Public Function GetArrayLength(aMatrix As Variant) As Long
    If Not IsArray(aMatrix) Then
        Let GetArrayLength = 0
    Else
        Let GetArrayLength = UBound(aMatrix) - LBound(aMatrix) + 1
    End If
End Function

' Returns the number of rows in a 2D array
Public Function GetNumberOfRows(aMatrix As Variant) As Long
    If Not IsArray(aMatrix) Then
        Let GetNumberOfRows = 0
    ElseIf NumberOfDimensions(aMatrix) = 1 Then
        Let GetNumberOfRows = 1
    Else
        Let GetNumberOfRows = UBound(aMatrix, 1) - LBound(aMatrix, 1) + 1
    End If
End Function

' Returns the number of columns in a 2D array. A 1D array is considered a 2D array with 1 row
Public Function GetNumberOfColumns(aMatrix As Variant) As Long
    If IsArray(aMatrix) Then
        If NumberOfDimensions(aMatrix) = 1 Then
            Let GetNumberOfColumns = UBound(aMatrix, 1) - LBound(aMatrix, 1) + 1
        Else
            Let GetNumberOfColumns = UBound(aMatrix, 2) - LBound(aMatrix, 2) + 1
        End If
    Else
        Let GetNumberOfColumns = 0
    End If
End Function

' This function takes a 1D range and trims its contents after converting it to upper case
Public Function TrimAndConvertArrayToCaps(TheArray As Variant) As Variant
    Dim i As Long
    Dim j As Long
    Dim ResultsArray As Variant
    
    ' Exit with an empty array if TheArray is empty
    If EmptyArrayQ(TheArray) Then
        Let TrimAndConvertArrayToCaps = Array()
        Exit Function
    End If
    
    If NumberOfDimensions(TheArray) = 0 Then
        Let ResultsArray = UCase(Trim(TheArray))
    ElseIf NumberOfDimensions(TheArray) = 1 Then
        Let ResultsArray = TheArray
        
        For i = LBound(TheArray) To UBound(TheArray)
            Let ResultsArray(i) = UCase(Trim(TheArray(i)))
        Next i
    ElseIf NumberOfDimensions(TheArray) = 2 Then
        Let ResultsArray = TheArray
        
        For i = LBound(TheArray, 1) To UBound(TheArray, 1)
            For j = LBound(TheArray, 2) To UBound(TheArray, 2)
                Let ResultsArray(i, j) = UCase(Trim(TheArray(i, j)))
            Next j
        Next i
    Else
        Let ResultsArray = TheArray
    End If
    
    Let TrimAndConvertArrayToCaps = ResultsArray
End Function

' This function returns a constant array with the requested dimensions
' If NCols is not given, the result is a 1D array
Public Function ConstantArray(TheValue As Variant, N As Long, Optional NCols As Variant) As Variant
    Dim TheResult() As Variant
    Dim c As Long
    Dim r As Long
    
    If IsMissing(NCols) Then
        If Not IsNumeric(N) Then
            Let ConstantArray = Array()
            Exit Function
        ElseIf N < 0 Then
            Let ConstantArray = Array()
            Exit Function
        ElseIf N = 1 Then
            Let ConstantArray = Array(TheValue)
            Exit Function
        End If
    
        ReDim TheResult(1 To N)
        
        For c = 1 To N
            Let TheResult(c) = TheValue
        Next c
        
        Let ConstantArray = TheResult
    Else
        ReDim TheResult(1 To N, 1 To NCols)
        
        For r = 1 To N
            For c = 1 To NCols
                Let TheResult(r, c) = TheValue
            Next c
        Next r
        
        Let ConstantArray = TheResult
    End If
End Function

' This function returns the array resulting from concatenating B to the right of A.
' A and B must have the same dimensions (e.g. 1 or 2D)
' If dim(A)<>dim(B) or dim(A)>2 or dim(B)>2 or dim(A)<1 or dim(B)<1 then
' this function returns Array()
Public Function ConcatenateArrays(a As Variant, b As Variant) As Variant
    If NumberOfDimensions(a) = 1 And NumberOfDimensions(b) = 1 Then
        Let ConcatenateArrays = ConvertTo1DArray(GetRow(TransposeMatrix(Stack2DArrays(TransposeMatrix(a), TransposeMatrix(b))), 1))
    Else
        Let ConcatenateArrays = TransposeMatrix(Stack2DArrays(TransposeMatrix(a), TransposeMatrix(b)))
    End If
End Function

' This function sorts the given range by the columns whose positions are given by ' ArrayOfColPos.
' The sorting orientation in each column are in ArrayOfColsSortOrder
' ArrayOfColsSortOrder is a variant array whose elements are all of enumerated type XLSortOrder (e.g. xlAscending, xlDescending)
Public Sub SortRange(MyRange As Range, ArrayOfColPos As Variant, _
                     ArrayOfColsSortOrder As Variant, _
                     WithHeaders As XlYesNoGuess)
    Dim TmpSheet As Worksheet
    Dim i As Integer
    
    ' Clear any previous sorting criteria
    Call MyRange.Worksheet.Sort.SortFields.Clear

    ' Add all the sorting criteria
    For i = LBound(ArrayOfColPos) To UBound(ArrayOfColPos)
        ' Add criteria to sort by date
        MyRange.Worksheet.Sort.SortFields.Add _
            Key:=MyRange.Columns(ArrayOfColPos(i)), _
            SortOn:=xlSortOnValues, _
            Order:=ArrayOfColsSortOrder(i), _
            DataOption:=xlSortNormal
    Next i
    
    ' Execute the sort
    With MyRange.Worksheet.Sort
        .SetRange MyRange
        .Header = WithHeaders
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub

' This function swaps two columns from a range, returning True if the operation is successful or False if it is not
Public Function SwapRangeColumns(TheRange As Range, FirstColumnIndex As Long, SecondColumIndex As Long) As Boolean
    Dim col1 As Range
    
    If FirstColumnIndex < 1 Or FirstColumnIndex > TheRange.Columns.Count Or SecondColumIndex < 1 Or SecondColumIndex > TheRange.Columns.Count Or FirstColumnIndex = SecondColumIndex Then
        Let SwapRangeColumns = False

        Exit Function
    End If
    
    Call TheRange.Worksheet.Activate
    
    Set col1 = TheRange.Worksheet.UsedRange
    Set col1 = col1.Range("a1").Offset(col1.Rows.Count, col1.Columns.Count).Resize(TheRange.Rows.Count, 1)
    
    Call TheRange.Columns(FirstColumnIndex).Copy
    Call col1.Range("A1").Select
    Call TheRange.Worksheet.Paste
    
    Call TheRange.Columns(SecondColumIndex).Copy
    Call TheRange.Cells(1, FirstColumnIndex).Select
    Call TheRange.Worksheet.Paste
    
    Call col1.Copy
    Call TheRange.Columns(SecondColumIndex).Select
    Call TheRange.Worksheet.Paste
    
    Call col1.ClearContents
    Call col1.ClearFormats
    Call col1.ClearComments
    Call col1.ClearHyperlinks
    Call col1.ClearNotes
    Call col1.ClearOutline
    
    Let SwapRangeColumns = True
End Function

' This function swaps two columns from a range, returning True if the operation is successful or False if it is not
Public Function SwapRangeRows(TheRange As Range, FirstRowIndex As Long, SecondRowIndex As Long) As Boolean
    Dim Row1 As Range
    
    If FirstRowIndex < 1 Or FirstRowIndex > TheRange.Columns.Count Or SecondRowIndex < 1 Or SecondRowIndex > TheRange.Columns.Count Or FirstRowIndex = SecondRowIndex Then
        Let SwapRangeRows = False

        Exit Function
    End If
    
    Call TheRange.Worksheet.Activate
    
    Set Row1 = TheRange.Worksheet.UsedRange
    Set Row1 = Row1.Range("a1").Offset(Row1.Rows.Count, Row1.Columns.Count).Resize(1, TheRange.Columns.Count)
    
    Call TheRange.Rows(FirstRowIndex).Copy
    Call Row1.Range("A1").Select
    Call TheRange.Worksheet.Paste
    
    Call TheRange.Rows(SecondRowIndex).Copy
    Call TheRange.Rows(FirstRowIndex).Select
    Call TheRange.Worksheet.Paste
    
    Call Row1.Copy
    Call TheRange.Rows(SecondRowIndex).Select
    Call TheRange.Worksheet.Paste
    
    Call Row1.ClearContents
    Call Row1.ClearFormats
    Call Row1.ClearComments
    Call Row1.ClearHyperlinks
    Call Row1.ClearNotes
    Call Row1.ClearOutline
    
    Let SwapRangeRows = True
End Function

' This function swaps two columns from a range, returning the resulting matrix if the operation is successful or False if it is not
'***HERE
Public Function SwapMatrixColumns(TheMatrix As Variant, FirstColumnIndex As Long, SecondColumIndex As Long) As Variant
    Dim col1 As Variant
    
    If FirstColumnIndex < 1 Or FirstColumnIndex > GetNumberOfColumns(TheMatrix) Or SecondColumIndex < 1 Or SecondColumIndex > GetNumberOfColumns(TheMatrix) Or FirstColumnIndex = SecondColumIndex Then
        Let SwapMatrixColumns = False

        Exit Function
    End If
    
    Call ToTemp(TheMatrix)
    Let col1 = GetColumn(TheMatrix, FirstColumnIndex)
    Call DumpInSheet(GetColumn(TheMatrix, SecondColumIndex), TempComputation.Cells(1, FirstColumnIndex))
    Call DumpInSheet(col1, TempComputation.Cells(1, SecondColumIndex))
    
    Let SwapMatrixColumns = TempComputation.Range("A1").CurrentRegion.Value2
End Function

' This function swaps two columns from a range, returning  the resulting matrix if the operation is successful or False if it is not
'***HERE
Public Function SwapMatrixRows(TheMatrix As Variant, FirstRowIndex As Long, SecondRowIndex As Long) As Variant
    Dim Row1 As Variant

    If FirstRowIndex < 1 Or FirstRowIndex > GetNumberOfRows(TheMatrix) Or SecondRowIndex < 1 Or SecondRowIndex > GetNumberOfRows(TheMatrix) Or FirstRowIndex = SecondRowIndex Then
        Let SwapMatrixRows = False

        Exit Function
    End If

    Call ToTemp(TheMatrix)
    Let Row1 = GetRow(TheMatrix, FirstRowIndex)
    Call DumpInSheet(GetRow(TheMatrix, SecondRowIndex), TempComputation.Cells(FirstRowIndex, 1))
    Call DumpInSheet(Row1, TempComputation.Cells(SecondRowIndex, 1))
    
    Let SwapMatrixRows = TempComputation.Range("A1").CurrentRegion.Value2
End Function

' This function transposes 1D or 2D arrays
' This function uses the built-in transposition function unless the optional parameter
' UseBuiltInQ is passed with a value of False.
Public Function TransposeMatrix(aMatrix As Variant, Optional UseBuiltInQ As Boolean = True) As Variant
    Dim r As Long
    Dim c As Long
    Dim TheResult() As Variant
    
    If Not DimensionedQ(aMatrix) Then
        Let TransposeMatrix = Null
        Exit Function
    End If
    
    If NumberOfDimensions(aMatrix) = 0 Then
        Let TransposeMatrix = aMatrix
        Exit Function
    End If
    
    If EmptyArrayQ(aMatrix) Then
        Let TransposeMatrix = Array()
        Exit Function
    End If
    
    If UseBuiltInQ Then
        Let TransposeMatrix = Application.Transpose(aMatrix)
        Exit Function
    End If
    
    If NumberOfDimensions(aMatrix) = 1 Then
        ReDim TheResult(LBound(aMatrix) To UBound(aMatrix), 1)
    
        For c = LBound(aMatrix) To UBound(aMatrix)
            Let TheResult(c, 1) = aMatrix(c)
        Next c
        
        Let TransposeMatrix = TheResult
        
        Exit Function
    End If
    
    ReDim TheResult(LBound(aMatrix, 2) To UBound(aMatrix, 2), LBound(aMatrix, 1) To UBound(aMatrix, 1))

    For r = LBound(aMatrix, 1) To UBound(aMatrix, 1)
        For c = LBound(aMatrix, 2) To UBound(aMatrix, 2)
            Let TheResult(c, r) = aMatrix(r, c)
        Next c
    Next r
    
    Let TransposeMatrix = TheResult
End Function

' This function transposes a 1D array of 1D arrays into a 1D array of 1D arrays.
' For instance,
' Array(Array(1,2,3), Array(10,20,30)) => Array(Array(1,10), Array(2, 20), Array(3, 30))
Public Function TransposeRectangular1DArrayOf1DArrays(AnArray As Variant) As Variant
    Dim r As Long
    Dim c As Long
    Dim TheResult() As Variant
    Dim TheMatrix As Variant
    Dim var As Variant
    Dim TheLength As Long
    
    ' Exit with null if AnArray is not an array
    If Not IsArray(AnArray) Then
        Let TransposeRectangular1DArrayOf1DArrays = Null
        Exit Function
    End If
    
    ' Exith with Null if AnArray is an empty array
    If EmptyArrayQ(AnArray) Then
        Let TransposeRectangular1DArrayOf1DArrays = Null
        Exit Function
    End If
    
    Let TheLength = GetArrayLength(First(AnArray))
    
    For Each var In AnArray
        If Not AtomicArrayQ(var) Or GetArrayLength(var) <> TheLength Then
            Let TransposeRectangular1DArrayOf1DArrays = Null
            Exit Function
        End If
    Next
    
    Let TheMatrix = Pack2DArray(AnArray)
    ReDim TheResult(1 To GetNumberOfColumns(TheMatrix))
    For c = 1 To GetArrayLength(First(AnArray))
        Let TheResult(c) = ConvertTo1DArray(GetColumn(TheMatrix, c))
    Next
    
    Let TransposeRectangular1DArrayOf1DArrays = TheResult
End Function

' This function returns a 2D array consolidating the special visible cells of an autofiltered range
' The returned array includes the header row
Public Function GetConsolidatedVisibleCells(ARange As Range) As Variant
    Call TempComputation.Range("A1").Resize(ARange.Rows.Count + 1, ARange.Columns.Count + 1).ClearContents
    Call DumpVisibleCells(SourceRange:=ARange, TargetRange:=TempComputation.Range("A1"))
    
    Let GetConsolidatedVisibleCells = TempComputation.Range("A1").CurrentRegion.Value2
End Function

' This function copies the visible cells of a range to the range defined by the upper left corner
' of the target range
Public Sub DumpVisibleCells(SourceRange As Range, TargetRange As Range)
    Call SourceRange.SpecialCells(xlCellTypeVisible).Copy
    Call TargetRange.Range("A1").Resize().PasteSpecial(Paste:=xlPasteAll)
End Sub

' This function returns the autofiltered result of a matrix with a header row. The resulting set is returned
' as a matrix.
' Parameters:
' 1. ColumnsToFilter is an array of integers
' 2. Criteria1List is an array of variants
' 3. OperatorList is an array of xlAutoFilterOperator (e.g. xlFilterValues, xlAnd, xlOr, etc.)
' 4. Criteria2List is an array of variants
'
' If the optional parameters are present, you must make them as long as ColumnsToFilter.
' If a particular column uses Criteria1List exclusively (e.g. filtered on one single condition),
' then enter Empty for the corresponding array elements in OperatorList and Criteria2List.
'
' For example, assume the matrix has 5 columns.  Assume you are filter on columns 1 and 3. Then,
' sample sets of parameters are:
'
' Example 1
' 1. ColumnsToFilter = Array(1,3)
' 2. Criteria1List = Array("<=0.5", "=10")
'
' Example 2
' 1. ColumnsToFilter = Array(1,3)
' 2. Criteria1List = Array("<=0.5", "<=1,5")
' 3. OperatorList = Array(Empty, xlOr)
' 4. Criteria2List = Array(Empty, ">=0.5")
'
' An empty array (identifier via the EmptyArrayQ() predicate, is returned when there is an error.
Public Function AutofilterMatrix(TheMatrix As Variant, ColumnsToFilter As Variant, Criteria1List As Variant, _
                                 Optional OperatorList As Variant, Optional Criteria2List As Variant) As Variant
    Dim i As Integer
    Dim ListObjRef As ListObject
    Dim NCols As Long
    Dim NRows As Long

    If NumberOfDimensions(TheMatrix) <> 2 Or GetNumberOfRows(TheMatrix) < 2 Or _
       NumberOfDimensions(ColumnsToFilter) <> 1 Or NumberOfDimensions(Criteria1List) <> 1 Or _
       GetArrayLength(ColumnsToFilter) <> GetArrayLength(Criteria1List) Then
        Let AutofilterMatrix = Array()
        Exit Function
    End If
    
    If Not IsMissing(OperatorList) Or Not IsMissing(Criteria2List) Then
        ' At least one of the optional parameters is present, both must be present
        If IsMissing(OperatorList) Or IsMissing(Criteria2List) Then
            Let AutofilterMatrix = Array()
            Exit Function
        End If
        
        If NumberOfDimensions(OperatorList) <> 1 Or NumberOfDimensions(Criteria2List) <> 1 Or _
           GetArrayLength(OperatorList) <> GetArrayLength(Criteria1List) Or _
           GetArrayLength(Criteria2List) <> GetArrayLength(Criteria1List) Then
            Let AutofilterMatrix = Array()
            Exit Function
        End If
    End If
    
    ' If the code gets to this point, parameters are consistent.

    ' Get dimensions of source matrix
    Let NCols = GetNumberOfColumns(TheMatrix)
    Let NRows = GetNumberOfRows(TheMatrix)

    ' Clear any existing contents in worksheet TempComputation
    Call TempComputation.UsedRange.ClearContents

    ' Dump matrix in worksheet TempComputation
    Call DumpInSheet(TheMatrix, TempComputation.Range("A1").Offset(NRows + 1))

    ' Turn the range into a table because this makes it easier to do computations
    Set ListObjRef = TempComputation.ListObjects.Add(SourceType:=xlSrcRange, Source:=TempComputation.Range("A1").Offset(NRows + 1).CurrentRegion, XlListObjectHasHeaders:=xlYes)

    ' Add each of the filtering conditions
    For i = LBound(ColumnsToFilter) To UBound(ColumnsToFilter)
        If IsMissing(Criteria2List) Then
            Call ListObjRef.Range.AutoFilter(Field:=ColumnsToFilter(i), Criteria1:=Criteria1List(i))
        Else
            Call ListObjRef.Range.AutoFilter(Field:=ColumnsToFilter(i), Criteria1:=Criteria1List(i), Operator:=OperatorList(i), Criteria2:=Criteria2List(i))
        End If
    Next i

    Call DumpVisibleCells(ListObjRef.Range, TempComputation.Range("A1"))

    Let AutofilterMatrix = TempComputation.Range("A1").CurrentRegion.Value2

    Call TempComputation.UsedRange.EntireColumn.Delete
End Function


' This returns an array of the same dimensions as the original one but with quotes surrounding
' every element.  This is done for dimensions one or two.  Matrices of higher dimensionality
' are returned unchanged.
Public Function AddQuotesToAllArrayElements(TheArray As Variant, Optional EscapeSingleQuote As Boolean = True) As Variant
    Dim NumberOfDims As Integer
    Dim i As Long
    Dim j As Long
    Dim TheResults As Variant
    
    Let NumberOfDims = NumberOfDimensions(TheArray)
    Let TheResults = TheArray
    
    If NumberOfDims = 1 Then
        For i = LBound(TheResults) To UBound(TheResults)
            If IsError(TheResults(i)) Then
                Let TheResults(i) = "NULL"
            ElseIf Trim(TheResults(i)) = "" Or (TheResults(i) = Empty And Not IsNumeric(TheResults(i))) Or TheResults(i) = "NULL" Then
                Let TheResults(i) = "NULL"
            Else
                Let TheResults(i) = """" & IIf(EscapeSingleQuote, esc(CStr(TheResults(i))), TheResults(i)) & """"
            End If
        Next i
    ElseIf NumberOfDims = 2 Then
        For i = LBound(TheResults, 1) To UBound(TheResults, 1)
            For j = LBound(TheResults, 2) To UBound(TheResults, 2)
                If Trim(TheResults(i, j)) = "" Or (TheResults(i, j) = Empty And Not IsNumeric(TheResults(i, j))) Or TheResults(i, j) = "NULL" Then
                    Let TheResults(i, j) = "NULL"
                Else
                    Let TheResults(i, j) = """" & IIf(EscapeSingleQuote, esc(CStr(TheResults(i, j))), TheResults(i, j)) & """"
                End If
            Next j
        Next i
    End If
        
    Let AddQuotesToAllArrayElements = TheResults
End Function

' Returns an identical array, but with single quotes (e.g. ') surrounding every element
Public Function AddSingleQuotesToAllArrayElements(TheArray As Variant, Optional EscapeSingleQuote As Boolean = True) As Variant
    Dim NumberOfDims As Integer
    Dim i As Long
    Dim j As Long
    Dim TheResults As Variant
    
    Let NumberOfDims = NumberOfDimensions(TheArray)
    Let TheResults = TheArray
    
    If NumberOfDims = 1 Then
        For i = LBound(TheResults) To UBound(TheResults)
            If IsError(TheResults(i)) Then
                Let TheResults(i) = "NULL"
            ElseIf Trim(TheResults(i)) = "" Or TheResults(i) = Empty Or TheResults(i) = "NULL" Then
                Let TheResults(i) = "NULL"
            Else
                Let TheResults(i) = "'" & IIf(EscapeSingleQuote, esc(CStr(TheResults(i))), TheResults(i)) & "'"
            End If
        Next i
    ElseIf NumberOfDims = 2 Then
        For i = LBound(TheResults, 1) To UBound(TheResults, 1)
            For j = LBound(TheResults, 2) To UBound(TheResults, 2)
                If IsError(TheResults(i, j)) Then
                    Let TheResults(i, j) = "NULL"
                ElseIf Trim(TheResults(i, j)) = "" Or (TheResults(i, j) = Empty And Not IsNumeric(TheResults(i, j))) Or TheResults(i, j) = "NULL" Then
                    Let TheResults(i, j) = "NULL"
                Else
                    Let TheResults(i, j) = "'" & IIf(EscapeSingleQuote, esc(CStr(TheResults(i, j))), TheResults(i, j)) & "'"
                End If
            Next j
        Next i
    End If
        
    Let AddSingleQuotesToAllArrayElements = TheResults
End Function

Public Function AddSingleBackQuotesToAllArrayElements(TheArray As Variant, Optional EscapeSingleQuote As Boolean = True) As Variant
    Dim NumberOfDims As Integer
    Dim i As Long
    Dim j As Long
    Dim TheResults As Variant
    
    Let NumberOfDims = NumberOfDimensions(TheArray)
    Let TheResults = TheArray
    
    If NumberOfDims = 1 Then
        For i = LBound(TheResults) To UBound(TheResults)
            If IsError(TheResults(i)) Then
                Let TheResults(i) = "NULL"
            ElseIf Trim(TheResults(i)) = "" Or (TheResults(i) = Empty And Not IsNumeric(TheResults(i))) Or TheResults(i) = "NULL" Then
                Let TheResults(i) = "NULL"
            Else
                Let TheResults(i) = "`" & IIf(EscapeSingleQuote, esc(CStr(TheResults(i))), TheResults(i)) & "`"
            End If
        Next i
    ElseIf NumberOfDims = 2 Then
        For i = LBound(TheResults, 1) To UBound(TheResults, 1)
            For j = LBound(TheResults, 2) To UBound(TheResults, 2)
                If IsError(TheResults(i, j)) Then
                    Let TheResults(i, j) = "NULL"
                ElseIf Trim(TheResults(i, j)) = "" Or (TheResults(i, j) = Empty And Not IsNumeric(TheResults(i, j))) Or TheResults(i, j) = "NULL" Then
                    Let TheResults(i, j) = "NULL"
                Else
                    Let TheResults(i, j) = "`" & IIf(EscapeSingleQuote, esc(CStr(TheResults(i, j))), TheResults(i, j)) & "`"
                End If
            Next j
        Next i
    End If
        
    Let AddSingleBackQuotesToAllArrayElements = TheResults
End Function

' This function acts on a 2D array.  It adds quotes around every element of the given submatrix
Public Function AddQuotesToAllSubArrayElements(TheArray As Variant, LeftIndex As Integer, TopIndex As Integer, _
                                               NRows As Integer, NColumns As Integer, Optional EscapeSingleQuote As Boolean = True) As Variant
    Dim i As Long
    Dim j As Long
    
    Dim TheResults As Variant
    
    Let TheResults = TheArray
    
    For i = TopIndex To TopIndex + NRows - 1
        For j = LeftIndex To LeftIndex + NColumns - 1
            If IsError(TheResults(i, j)) Then
                Let TheResults(i, j) = "NULL"
            ElseIf Trim(TheResults(i, j)) = "" Or (TheResults(i, j) = Empty And Not IsNumeric(TheResults(i, j))) Or TheResults(i, j) = "NULL" Then
                Let TheResults(i, j) = "NULL"
            Else
                Let TheResults(i, j) = """" & IIf(EscapeSingleQuote, esc(CStr(TheResults(i, j))), TheResults(i, j)) & """"
            End If
        Next j
    Next i
        
    Let AddQuotesToAllSubArrayElements = TheResults
End Function

Public Function AddSingleQuotesToAllSubArrayElements(TheArray As Variant, LeftIndex As Integer, TopIndex As Integer, _
                                               NRows As Integer, NColumns As Integer, Optional EscapeSingleQuote As Boolean = True) As Variant
    Dim i As Long
    Dim j As Long
    Dim TheResults As Variant
    
    Let TheResults = TheArray
    
    For i = TopIndex To TopIndex + NRows - 1
        For j = LeftIndex To LeftIndex + NColumns - 1
            If IsError(TheResults(i, j)) Then
                Let TheResults(i, j) = "NULL"
            ElseIf Trim(TheResults(i, j)) = "" Or TheResults(i, j) = "NULL" Or TheResults(i, j) = Empty Then
                Let TheResults(i, j) = "NULL"
            Else
                Let TheResults(i, j) = "'" & IIf(EscapeSingleQuote, esc(CStr(TheResults(i, j))), TheResults(i, j)) & "'"
            End If
        Next j
    Next i
        
    Let AddSingleQuotesToAllSubArrayElements = TheResults
End Function

Public Function AddSingleBackQuotesToAllSubArrayElements(TheArray As Variant, LeftIndex As Integer, TopIndex As Integer, _
                                               NRows As Integer, NColumns As Integer, Optional EscapeSingleQuote As Boolean = True) As Variant
    Dim i As Long
    Dim j As Long
    Dim TheResults As Variant
    
    Let TheResults = TheArray
    
    For i = TopIndex To TopIndex + NRows - 1
        For j = LeftIndex To LeftIndex + NColumns - 1
            If IsError(TheResults(i, j)) Then
                Let TheResults(i, j) = "NULL"
            ElseIf Trim(TheResults(i, j)) = "" Or (TheResults(i, j) = Empty And Not IsNumeric(TheResults(i, j))) Or TheResults(i, j) = "NULL" Then
                Let TheResults(i, j) = "NULL"
            Else
                Let TheResults(i, j) = "`" & IIf(EscapeSingleQuote, esc(CStr(TheResults(i, j))), TheResults(i, j)) & "`"
            End If
        Next j
    Next i
        
    Let AddSingleBackQuotesToAllSubArrayElements = TheResults
End Function

' Converts a 2D array into properly quoted array.  This means that numeric columns are not quoted but others are.  The only
' exception are empty columns, which are converted to NULL.  This returns the properly quoted string.
' This function may be used on the data array before calling sub ConnectAndExecuteInsertQuery() in this module
Public Function DoubleQuote2DArray(TheArray As Variant) As Variant
    Dim r As Long
    Dim c As Long
    Dim ParamCopy As Variant
    
    Let ParamCopy = TheArray
    
    For r = 1 To GetNumberOfRows(ParamCopy)
        For c = 1 To GetNumberOfColumns(ParamCopy)
            If IsError(ParamCopy(r, c)) Then
                Let ParamCopy(r, c) = "NULL"
            ElseIf IsEmpty(ParamCopy(r, c)) Or Trim(ParamCopy(r, c)) = "" Then
                Let ParamCopy(r, c) = "NULL"
            ElseIf TypeName(ParamCopy(r, c)) = "String" And Not IsNumeric(ParamCopy(r, c)) Then
                Let ParamCopy(r, c) = """" & ParamCopy(r, c) & """"
            End If
        Next c
    Next r
    
    Let DoubleQuote2DArray = ParamCopy
End Function

' Given a 2D array like [{1,2,3; 4, 5, 6}], this function returns the following STRING
' (1,2,3), (4,5,6)
'
' This is very useful to create insert statements for databases.
Public Function Convert2DArrayIntoListOfParentheticalExpressions(TheArray As Variant) As String
    Dim TheList As String
    Dim i As Integer

    For i = 1 To GetNumberOfRows(TheArray)
        Let TheList = TheList & Convert1DArrayIntoParentheticalExpression(GetRow(TheArray, CLng(i)))
        
        If i < GetNumberOfRows(TheArray) Then
            Let TheList = TheList & ", "
        End If
    Next i
    
    Let Convert2DArrayIntoListOfParentheticalExpressions = TheList
End Function

' This function resizes a 2D array while preserving values.
' This function returns a variant. AnArray must be a 2D array.
' If either NRows and NCols is less than Ubound(AnArray,1) or Ubound(AnArray,2) then the resulting
' matrix truncates the input matrix accordingly
Public Function Redim2DArray(AnArray As Variant, NRows As Long, NCols As Long) As Variant
    Dim NewArray() As Variant
    Dim r As Long
    Dim c As Long
    
    If NumberOfDimensions(AnArray) <> 2 Then
        Let Redim2DArray = AnArray
        Exit Function
    End If
    
    If Not IsArray(AnArray) Then
        Let Redim2DArray = AnArray
        Exit Function
    End If
    
    If NRows < LBound(AnArray, 1) Or NCols < LBound(AnArray, 2) Then
        Let Redim2DArray = AnArray
        Exit Function
    End If
            
    ReDim NewArray(LBound(AnArray, 1) To IIf(LBound(AnArray, 1) = 0, NRows - 1, NRows), _
                   LBound(AnArray, 2) To IIf(LBound(AnArray, 2) = 0, NCols - 1, NCols))
                   
    For r = LBound(AnArray, 1) To Application.Min(NRows + IIf(LBound(AnArray, 1) = 0, NRows - 1, NRows), UBound(AnArray, 1))
        For c = LBound(AnArray, 2) To Application.Min(NCols + IIf(LBound(AnArray, 2) = 0, NCols - 1, NRows), UBound(AnArray, 2))
            Let NewArray(r, c) = AnArray(r, c)
        Next c
    Next r
    
    Let Redim2DArray = NewArray
End Function

' The purpose of this function is to extend a2Darray1 with data from a2Darray2 using
' equality on the given key columns.  The function returns the "left joined" 2D array.
' This means that all rows in array1 are included. The data from array2 is included
' only if its key in also in array 1.
'
' If a key if found more than once in a2Darray, the first ocurrance is used. The resulting
' 2D array uses data from the columns in a2DArray1 specified in ColsPosArrayFrom2DArray1
' and the columns in a2DArray2 specified in ColsPosArrayFrom2DArray2
'
' When the parameters are inconsistent, the function returns Null
Public Function LeftJoin2DArraysOnKeyEquality(a2DArray1 As Variant, _
                                              a2DArray1KeyColPos As Integer, _
                                              ColsPosArrayFrom2DArray1 As Variant, _
                                              a2DArray2 As Variant, _
                                              a2DArray2KeyColPos As Integer, _
                                              ColsPosArrayFrom2DArray2 As Variant, _
                                              Optional ArraysHaveHeadersQ As Boolean = True, _
                                              Optional IncludeHeadersQ As Boolean = True) As Variant
    Dim ResultsDict As Dictionary
    Dim Array2TrackingDict As Dictionary
    Dim r As Long
    Dim var As Variant
    Dim TheKey As Variant
    Dim TheItems As Variant
    Dim AppendedItems As Variant
    Dim NumColsArray1 As Integer
    Dim NumColsArray2 As Integer
    Dim JoinedHeadersRow As Variant
    Dim TheResults As Variant

    ' Parameter consistency checks
    
    ' Exit with Null if ArraysHaveHeadersQ is False and IncludeHeadersQ = True
    If Not ArraysHaveHeadersQ And IncludeHeadersQ Then
        Let LeftJoin2DArraysOnKeyEquality = Null
        Exit Function
    End If
    
    ' Exit with Null if either a2Array1, a2Array2, ColsPosArrayFrom2DArray1
    ' ColsPosArrayFrom2DArray2 is not an array
    If Not (IsArray(a2DArray1) And IsArray(a2DArray2) And _
            IsArray(ColsPosArrayFrom2DArray1) And IsArray(ColsPosArrayFrom2DArray2)) Then
        Let LeftJoin2DArraysOnKeyEquality = Null
        Exit Function
    End If
    
    ' Exit with Null if a2DArray1 and a2DArray2 don't have the number of dimensions
    If NumberOfDimensions(a2DArray1) <> NumberOfDimensions(a2DArray2) Or _
       NumberOfDimensions(a2DArray1) <> 2 Then
        Let LeftJoin2DArraysOnKeyEquality = Null
        Exit Function
    End If
    
    ' Exit if either a2DArray1KeyColPos or a2DArray2KeyColPos is not an integer
    If (TypeName(a2DArray1KeyColPos) <> TypeName(1) And TypeName(a2DArray1KeyColPos) <> TypeName(99999999)) Or _
       (TypeName(a2DArray2KeyColPos) <> TypeName(1) And TypeName(a2DArray2KeyColPos) <> TypeName(99999999)) Then
        Let LeftJoin2DArraysOnKeyEquality = Null
        Exit Function
    End If
    
    ' Exit with Null if the index of key1 is non-positive or larger than the number
    ' of columns in array1 or the index of key2 is non-positive or larger than the number
    ' of columns in array2
    If a2DArray1KeyColPos < 1 Or a2DArray1KeyColPos > GetNumberOfColumns(a2DArray1) Or _
       a2DArray2KeyColPos < 1 Or a2DArray2KeyColPos > GetNumberOfColumns(a2DArray2) Then
        Let LeftJoin2DArraysOnKeyEquality = Null
        Exit Function
    End If
    
    ' Exit with Null if with ColsPosArrayFrom2DArray1 and ColsPosArrayFrom2DArray2
    ' are not numeric, positive, integer 1D arrays
    If Not (IsPositiveIntegerArrayQ(ColsPosArrayFrom2DArray1) And _
            IsPositiveIntegerArrayQ(ColsPosArrayFrom2DArray2)) Then
        Let LeftJoin2DArraysOnKeyEquality = Null
        Exit Function
    End If

    ' Exit with Null if any of the indices in ColsPosArrayFrom2DArray1 is less than 1
    ' or larger than the number of columns in array1
    For Each var In ColsPosArrayFrom2DArray1
        If var < 1 Or var > GetNumberOfColumns(a2DArray1) Then
            Let LeftJoin2DArraysOnKeyEquality = Null
            Exit Function
        End If
    Next
    
    ' Exit with Null if any of the indices in ColsPosArrayFrom2DArray2 is less than 1
    ' or larger than the number of columns in array2
    For Each var In ColsPosArrayFrom2DArray2
        If var < 1 Or var > GetNumberOfColumns(a2DArray2) Then
            Let LeftJoin2DArraysOnKeyEquality = Null
            Exit Function
        End If
    Next
    
    ' Exit with Null if either a2DArray1 or a2DArray2 has only 1 row when ArraysHaveHeadersQ
    ' is True
    If ArraysHaveHeadersQ And _
       (GetNumberOfRows(a2DArray1) < 2 Or GetNumberOfRows(a2DArray1) < 2) Then
        Let LeftJoin2DArraysOnKeyEquality = Null
        Exit Function
    End If

    ' Determine the number of columns from array 1
    Let NumColsArray1 = GetArrayLength(ColsPosArrayFrom2DArray1)
    
    ' Determine the number of columns from array 2 to append to array 1
    Let NumColsArray2 = GetArrayLength(ColsPosArrayFrom2DArray2)

    ' Load all information from a2DArray1 into a dictionary
    Set ResultsDict = New Dictionary
    For r = IIf(ArraysHaveHeadersQ, LBound(a2DArray1, 1) + 1, LBound(a2DArray1, 1)) To UBound(a2DArray1, 1)
        ' Get the for the current row
        Let TheKey = a2DArray1(r, a2DArray1KeyColPos)

        If Not ResultsDict.Exists(Key:=TheKey) Then
            ' Extract the columns needed from this security's row
            Let TheItems = Take(GetRow(a2DArray1, r), _
                                Prepend(ColsPosArrayFrom2DArray1, a2DArray1KeyColPos))
            
            ' Pad TheItems with enough slots for the items appended from a2DArray2
            Let TheItems = ConcatenateArrays(TheItems, ConstantArray(Empty, CLng(NumColsArray2)))

            ' Add the array of values to this security's entry
            Call ResultsDict.Add(Key:=TheKey, Item:=TheItems)
        End If
    Next r
    
    ' Scan a2DArray2 appending to the array of elements of each element in a2DArray1 the
    ' elements in a2DArray2
    Set Array2TrackingDict = New Dictionary
    For r = IIf(ArraysHaveHeadersQ, LBound(a2DArray2, 1) + 1, LBound(a2DArray2, 1)) To UBound(a2DArray2, 1)
        ' Get the for the current row
        Let TheKey = a2DArray2(r, a2DArray1KeyColPos)
        
        ' Append to the items in the results dicts for the current key if the current key
        ' is in the results dictionary already, and they has not already been appended
        If ResultsDict.Exists(Key:=TheKey) And Not Array2TrackingDict.Exists(Key:=TheKey) Then
            ' Mark this row in a2DArray2 as having been processed
            Call Array2TrackingDict.Add(Key:=TheKey, Item:=Empty)
            
            ' Take the portion of the items corresponding to array 1
            Let TheItems = Take(ResultsDict.Item(Key:=TheKey), 1 + NumColsArray1)
            
            ' Get the required columns from this row to append to those already in the results
            ' dictionary
            Let AppendedItems = Take(GetRow(a2DArray2, r), ColsPosArrayFrom2DArray2)
            
            Let ResultsDict.Item(Key:=TheKey) = ConcatenateArrays(TheItems, AppendedItems)
        End If
    Next r
    
    ' Repack the results as a 2D array
    Let TheResults = Pack2DArray(ResultsDict.Items)
    
    If Not IncludeHeadersQ Then
        Let LeftJoin2DArraysOnKeyEquality = TheResults
    
        Exit Function
    End If
    
    ' Prepend the headers row if the user chose to
    Let JoinedHeadersRow = ConcatenateArrays(Take(GetRow(a2DArray1, 1), _
                                                  Prepend(ColsPosArrayFrom2DArray1, a2DArray1KeyColPos)), _
                                             Take(GetRow(a2DArray2, 1), _
                                                  ColsPosArrayFrom2DArray2))
    
    ' Prepend headers to return matrix
    Let LeftJoin2DArraysOnKeyEquality = Prepend(TheResults, JoinedHeadersRow)
End Function

' The purpose of this function is to extend a2Darray1 with data from a2Darray2 using
' equality on the given key columns.  The function returns the "left joined" 2D array.
' This means that all rows in array1 are included. The data from array2 is included
' only if its key in also in array 1.
'
' If a key if found more than once in a2Darray, the first ocurrance is used.  The resulting
' 2D array uses data from the columns in a2DArray1 specified in ColsPosArrayFrom2DArray1
' and the columns in a2DArray2 specified in ColsPosArrayFrom2DArray2
'
' When the parameters are inconsistent, the function returns Null
Public Function InnerJoin2DArraysOnKeyEquality(a2DArray1 As Variant, _
                                               a2DArray1KeyColPos As Integer, _
                                               ColsPosArrayFrom2DArray1 As Variant, _
                                               a2DArray2 As Variant, _
                                               a2DArray2KeyColPos As Integer, _
                                               ColsPosArrayFrom2DArray2 As Variant, _
                                               Optional ArraysHaveHeadersQ As Boolean = True, _
                                               Optional IncludeHeadersQ As Boolean = True) As Variant
    Dim Array2Dict As Dictionary
    Dim ResultsDict As Dictionary
    Dim r As Long
    Dim var As Variant
    Dim TheKey As Variant
    Dim TheItems As Variant
    Dim AppendedItems As Variant
    Dim NumColsArray1 As Integer
    Dim NumColsArray2 As Integer
    Dim JoinedHeadersRow As Variant
    Dim TheResults As Variant

    ' Parameter consistency checks

    ' Exit with Null if ArraysHaveHeadersQ is False and IncludeHeadersQ = True
    If Not ArraysHaveHeadersQ And IncludeHeadersQ Then
        Let InnerJoin2DArraysOnKeyEquality = Null
        Exit Function
    End If

    ' Exit with Null if either a2Array1, a2Array2, ColsPosArrayFrom2DArray1
    ' ColsPosArrayFrom2DArray2 is not an array
    If Not (IsArray(a2DArray1) And IsArray(a2DArray2) And _
            IsArray(ColsPosArrayFrom2DArray1) And IsArray(ColsPosArrayFrom2DArray2)) Then
        Let InnerJoin2DArraysOnKeyEquality = Null
        Exit Function
    End If

    ' Exit with Null if a2DArray1 and a2DArray2 don't have the number of dimensions
    If NumberOfDimensions(a2DArray1) <> NumberOfDimensions(a2DArray2) Or _
       NumberOfDimensions(a2DArray1) <> 2 Then
        Let InnerJoin2DArraysOnKeyEquality = Null
        Exit Function
    End If

    ' Exit if either a2DArray1KeyColPos or a2DArray2KeyColPos is not an integer
    If (TypeName(a2DArray1KeyColPos) <> TypeName(1) And TypeName(a2DArray1KeyColPos) <> TypeName(99999999)) Or _
       (TypeName(a2DArray2KeyColPos) <> TypeName(1) And TypeName(a2DArray2KeyColPos) <> TypeName(99999999)) Then
        Let InnerJoin2DArraysOnKeyEquality = Null
        Exit Function
    End If

    ' Exit with Null if the index of key1 is non-positive or larger than the number
    ' of columns in array1 or the index of key2 is non-positive or larger than the number
    ' of columns in array2
    If a2DArray1KeyColPos < 1 Or a2DArray1KeyColPos > GetNumberOfColumns(a2DArray1) Or _
       a2DArray2KeyColPos < 1 Or a2DArray2KeyColPos > GetNumberOfColumns(a2DArray2) Then
        Let InnerJoin2DArraysOnKeyEquality = Null
        Exit Function
    End If

    ' Exit with Null if with ColsPosArrayFrom2DArray1 and ColsPosArrayFrom2DArray2
    ' are not numeric, positive, integer 1D arrays
    If Not (IsPositiveIntegerArrayQ(ColsPosArrayFrom2DArray1) And _
            IsPositiveIntegerArrayQ(ColsPosArrayFrom2DArray2)) Then
        Let InnerJoin2DArraysOnKeyEquality = Null
        Exit Function
    End If

    ' Exit with Null if any of the indices in ColsPosArrayFrom2DArray1 is less than 1
    ' or larger than the number of columns in array1
    For Each var In ColsPosArrayFrom2DArray1
        If var < 1 Or var > GetNumberOfColumns(a2DArray1) Then
            Let InnerJoin2DArraysOnKeyEquality = Null
            Exit Function
        End If
    Next

    ' Exit with Null if any of the indices in ColsPosArrayFrom2DArray2 is less than 1
    ' or larger than the number of columns in array2
    For Each var In ColsPosArrayFrom2DArray2
        If var < 1 Or var > GetNumberOfColumns(a2DArray2) Then
            Let InnerJoin2DArraysOnKeyEquality = Null
            Exit Function
        End If
    Next

    ' Exit with Null if either a2DArray1 or a2DArray2 has only 1 row when ArraysHaveHeadersQ
    ' is True
    If ArraysHaveHeadersQ And _
       (GetNumberOfRows(a2DArray1) < 2 Or GetNumberOfRows(a2DArray1) < 2) Then
        Let InnerJoin2DArraysOnKeyEquality = Null
        Exit Function
    End If

    ' Determine the number of columns in arrays 1 and 2
    Let NumColsArray1 = GetArrayLength(ColsPosArrayFrom2DArray1)
    Let NumColsArray2 = GetArrayLength(ColsPosArrayFrom2DArray2)

    ' Index the contents of array2
    Set Array2Dict = New Dictionary
    For r = IIf(ArraysHaveHeadersQ, LBound(a2DArray2, 1) + 1, LBound(a2DArray2, 1)) To UBound(a2DArray2, 1)
        ' Get the for the current row
        Let TheKey = a2DArray2(r, a2DArray2KeyColPos)

        If Not Array2Dict.Exists(Key:=TheKey) Then
            ' Add the array of values to this security's entry
            Call Array2Dict.Add(Key:=TheKey, _
                                Item:=Take(GetRow(a2DArray2, r), ColsPosArrayFrom2DArray2))
        End If
    Next r

    ' Scan a2DArray2 appending to the array of elements of each element in a2DArray1 the
    ' elements in a2DArray2
    Set ResultsDict = New Dictionary
    For r = IIf(ArraysHaveHeadersQ, LBound(a2DArray1, 1) + 1, LBound(a2DArray1, 1)) To UBound(a2DArray1, 1)
        ' Get the for the current row
        Let TheKey = a2DArray1(r, a2DArray1KeyColPos)

        ' Create the join for this row in array1 if it is found in array2 based on the key.
        If Not ResultsDict.Exists(Key:=TheKey) And Array2Dict.Exists(Key:=TheKey) Then
            ' Extract the columns required from this row in array 1
            Let TheItems = Take(GetRow(a2DArray1, r), _
                                Prepend(ColsPosArrayFrom2DArray1, a2DArray1KeyColPos))

            ' Get the corresponding items from array 2
            Let AppendedItems = Array2Dict.Item(Key:=TheKey)

            ' Index the joined items for this row in array 1
            Call ResultsDict.Add(Key:=TheKey, Item:=ConcatenateArrays(TheItems, AppendedItems))
        End If
    Next r

    ' Repack the results as a 2D array
    Let TheResults = Pack2DArray(ResultsDict.Items)

    If Not IncludeHeadersQ Then
        Set InnerJoin2DArraysOnKeyEquality = TheResults

        Exit Function
    End If

    ' Prepend the headers row if the user chose to
    Let JoinedHeadersRow = ConcatenateArrays(Take(GetRow(a2DArray1, 1), _
                                                  Prepend(ColsPosArrayFrom2DArray1, a2DArray1KeyColPos)), _
                                             Take(GetRow(a2DArray2, 1), _
                                                  ColsPosArrayFrom2DArray2))

    ' Prepend headers to return matrix
    Let InnerJoin2DArraysOnKeyEquality = Prepend(TheResults, JoinedHeadersRow)
End Function

' Calculates the dot product of two vectors. Returns Null if the parameters are incompatible.
Public Function DotProduct(v1 As Variant, v2 As Variant) As Variant
    Dim v1prime As Variant
    Dim v2prime As Variant
    Dim i As Long
    Dim TheResult As Double

    If Not (VectorQ(v1) And VectorQ(v2)) Then
        Let DotProduct = Null
        Exit Function
    End If
    
    If Length(v1) <> Length(v2) Then
        Let DotProduct = Null
        Exit Function
    End If
    
    Let v1prime = ConvertTo1DArray(v1)
    Let v2prime = ConvertTo1DArray(v2)
    
    For i = 1 To Length(v1prime)
        Let TheResult = TheResult + v1prime(i) * v2prime(i)
    Next i
    
    Let DotProduct = TheResult
End Function

' Performs matrix multiplication. Both parameters must satisfy VectorQ() or MatrixQ() The function returns Null if the parameters are incompatible.
Public Function MatrixMultiply(m1 As Variant, m2 As Variant) As Variant
    Dim r As Long
    Dim c As Long
    Dim TheResult() As Double
    
    If Not (MatrixQ(m1) And MatrixQ(m2)) Then
        Let MatrixMultiply = Null
        Exit Function
    End If
    
    If NumberOfColumns(m1) <> NumberOfRows(m2) Then
        Let MatrixMultiply = Null
        Exit Function
    End If
    
    ReDim TheResult(1 To NumberOfRows(m1), 1 To NumberOfColumns(m2))
    For r = 1 To NumberOfRows(m1)
        For c = 1 To NumberOfColumns(m2)
            Let TheResult(r, c) = DotProduct(GetRow(m1, r), GetColumn(m2, c))
        Next c
    Next r
    
    Let MatrixMultiply = TheResult
End Function

' Fills in the banks in a 1D array.  Repeats the value in the first cell until it finds a different value.
' It then repeats that one until a new one is found.  And so forth.
Public Function FillArrayBlanks(ByVal AnArray As Variant) As Variant
    Dim CurrentValue As Variant
    Dim c As Long

    If Not DimensionedQ(AnArray) Then
        Let FillArrayBlanks = Null
        Exit Function
    End If
    
    If EmptyArrayQ(AnArray) Then
        Let FillArrayBlanks = Array()
        Exit Function
    End If

    If IsNull(First(AnArray)) Then
        Let CurrentValue = Empty
    Else
        Let CurrentValue = First(AnArray)
    End If
    
    For c = LBound(AnArray, 1) To UBound(AnArray, 1)
        If IsEmpty(AnArray(c)) Or IsNull(AnArray(c)) Then
            Let AnArray(c) = CurrentValue
        ElseIf AnArray(c) = Empty Then
            Let AnArray(c) = CurrentValue
        ElseIf CurrentValue <> AnArray(c) Then
            Let CurrentValue = AnArray(c)
        End If
    Next c
            
    Let FillArrayBlanks = AnArray
End Function

' DESCRIPTION
' This function return the given array with sequential repeatitions blanked out. It is
' useful to turn arrays into columns for Pivot Table-like arrangements.  For example,
' Array(Empty, Empty, 1, 1, 1, 2, 2, 3) turns into
' Array(Empty, Empty, 1, Empty, Empty, 2, Empty, 3)
' If AnArray fails Predicates.AtomicArrayQ, the fuction returns Null.  The same thing
' happens whenever the parameter makes no sense.
'
' PARAMETERS
' 1. AnArray       - An array satisfying Predicates.AtomicArrayQ
'
' RETURNED VALUE
' AnAtomicArrayOrTable after blanking out sequential repetitions of its elements.
Public Function BlankOutArraySequentialRepetitions(ByVal AnArray As Variant)
    Dim CurrentValue As Variant
    Dim c As Long

    If Not DimensionedQ(AnArray) Then
        Let BlankOutArraySequentialRepetitions = Null
        Exit Function
    End If
    
    If EmptyArrayQ(AnArray) Then
        Let BlankOutArraySequentialRepetitions = Array()
        Exit Function
    End If
    
    If Not AtomicArrayQ(AnArray) Then
        Let BlankOutArraySequentialRepetitions = Null
        Exit Function
    End If
    
    Let CurrentValue = First(AnArray)
    For c = LBound(AnArray, 1) + 1 To UBound(AnArray, 1)
        If IsEmpty(AnArray(c)) Or IsNull(AnArray(c)) Then
            Let AnArray(c) = Empty
        ElseIf CurrentValue = AnArray(c) Then
            Let AnArray(c) = Empty
        Else
            Let CurrentValue = AnArray(c)
        End If
    Next c
            
    Let BlankOutArraySequentialRepetitions = AnArray
End Function

' DESCRIPTION
' Returns the nth element in an array. Negative indices start at the end of the array with index = -1.
' This function returns Null if the index falls outside of the array or if the index is 0 or not an integer.
' If the given array is 2D, the function
'public function GetElement(AnArray,

' DESCRIPTION
' This function inserts new elements in a repeated fashion in between the elements of the
' given array. It can take several forms:
' 1. Riffle(Array(e1, e2, ...), elt)
' 2. Riffle(Array(e1, e2, ...), Array(elt1, elt2, ...))
' 3. Riffle(Array(e1, e2, ...), elt, n)
' 4. Riffle(Array(e1, e2, ...), elt, Array(imin, imax, n))
'    This case requires assumes that 1 instead of 0 is the first array position.
'    It also requires imin<=imax
'
' If there are fewer elements in Array(elt1, elt2, ...) than gaps between  in Riffle(Array(e1,e2,…),Array(x1,x2,…)),
' the Array(elt1, elt2, ...) are used cyclically. Riffle(Array(e),x) gives Array(e). The specification  is of the
' type used in Take. Negative indices count from the end of the list.
'
' In Riffle[list, xlist], if list and xlist are of the same length, then their elements are directly interleaved,
' so that the last element of the result is the last element of xlist.
'
' When A1DArray is an empty array, the empty array is returned unchanged.  When the parameters are inconsistent,
' the function returns Null.
Public Function Riffle(A1DArray As Variant, Arg2 As Variant, Optional StepInterval As Variant) As Variant
    Dim ResultsDict As Dictionary
    Dim var As Dictionary
    Dim r As Long
    Dim s As Long
    Dim c As Long

    If Not DimensionedQ(A1DArray) Then
        Let Riffle = Null
        Exit Function
    End If

    If EmptyArrayQ(A1DArray) Then
        Let Riffle = A1DArray
        Exit Function
    End If
    
    If Not AtomicArrayQ(A1DArray) Then
        Let Riffle = Null
        Exit Function
    End If
    
    If Not IsMissing(StepInterval) Then
        If Not (PositiveWholeNumberQ(StepInterval) Or IsArray(StepInterval)) Then
            Let Riffle = Null
            Exit Function
        End If
        
        If IsArray(StepInterval) Then
            If Not AtomicQ(Arg2) Then
                Let Riffle = Null
                Exit Function
            End If
            
            If Not PositiveIntegerArrayQ(StepInterval) And GetArrayLength(StepInterval) <> 3 Then
                Let Riffle = Null
                Exit Function
            End If
            
            If First(StepInterval) > First(Rest(StepInterval)) Then
                Let Riffle = Null
                Exit Function
            End If
        End If
    End If
    
    If Not (AtomicQ(Arg2) Or AtomicArrayQ(Arg2)) Then
        Let Riffle = Null
        Exit Function
    End If
    
    If IsArray(Arg2) Then
        If Not DimensionedQ(Arg2) Then
            Let Riffle = Null
            Exit Function
        End If

        If EmptyArrayQ(Arg2) Then
            Let Riffle = Null
            Exit Function
        End If
        
        If Not AtomicArrayQ(Arg2) Then
            Let Riffle = Null
            Exit Function
        End If
        
        If Not IsMissing(StepInterval) Then
            Let Riffle = Null
            Exit Function
        End If
    End If
    
    ' Case Riffle(Array(e1, e2, ...), elt) and Riffle(Array(e1, e2, ...), elt, n)
    If AtomicArrayQ(A1DArray) And AtomicQ(Arg2) Then
        If IsMissing(StepInterval) Then
            Let StepInterval = 1
        End If
        
        Let s = 1
        Set ResultsDict = New Dictionary
        For r = LBound(A1DArray, 1) To UBound(A1DArray, 1) - 1
            Call ResultsDict.Add(Key:=r, Item:=A1DArray(r))
            
            If s Mod StepInterval = 0 Then
                Call ResultsDict.Add(Key:=r & "-separator", Item:=Arg2)
                Let s = 1
            Else
                Let s = s + 1
            End If
        Next
        Call ResultsDict.Add(Key:=UBound(A1DArray, 1), Item:=A1DArray(UBound(A1DArray, 1)))
        
        Let Riffle = ResultsDict.Items
        Exit Function
    End If
    
    ' Case Riffle(Array(e1, e2, ...), Array(elt1, elt2, ...))
    If AtomicArrayQ(A1DArray) And AtomicArrayQ(Arg2) Then
        Let s = LBound(Arg2, 1)
        Let c = 1
        
        Set ResultsDict = New Dictionary
        For r = LBound(A1DArray, 1) To UBound(A1DArray, 1) - 1
            Call ResultsDict.Add(Key:=c, Item:=A1DArray(r))
            Call ResultsDict.Add(Key:=c + 1, Item:=Arg2(s))
                
            If s = UBound(Arg2, 1) Then
                Let s = LBound(Arg2, 1)
            Else
                Let s = s + 1
            End If
            
            Let c = c + 2
        Next
        Call ResultsDict.Add(Key:=c, Item:=A1DArray(UBound(A1DArray, 1)))
        If GetArrayLength(A1DArray) = GetArrayLength(Arg2) Then
            Call ResultsDict.Add(Key:=c + 1, _
                                 Item:=Arg2(UBound(Arg2, 1)))
        End If
        
        Let Riffle = ResultsDict.Items
        Exit Function
    End If
    
    ' Case Riffle(Array(e1, e2, ...), elt, Array(imin, imax, n))
    Let s = 1
    '***HERE1for r =
End Function

' Returns the nth element of an array, using the array's indexing conventions.  It returns
' Null if the arguments are inconsistent:
' 1. AnArray is not an array
' 2. AnArray is empty
' 3. AnArray is not dimensioned
' 4. AnArray has number of dimesions other than 1 or 2
' 5. TheColumnIndex is given but is not a whole number
Public Function GetElement(AnArray As Variant, TheIndex As Long) As Variant
    Dim NormalizedIndex As Long
       
    If Not DimensionedQ(AnArray) Then
        Let GetElement = Null
        Exit Function
    End If
        
    If EmptyArrayQ(AnArray) Then
        Let GetElement = Null
        Exit Function
    End If
    
    If NumberOfDimensions(AnArray) = 1 Then
        Let GetElement = AnArray(NormalizeArrayIndex(AnArray, TheIndex))
    ElseIf NumberOfDimensions(AnArray) = 2 Then
        Let GetElement = GetRow(AnArray, NormalizeArrayIndex(AnArray, TheIndex) + 1 - LBound(AnArray, 1))
    Else
        Let GetElement = Null
        Exit Function
    End If
End Function

' This function caps non-negative indices at LBound and UBound if it falls outside of
' the interval LBound to UBound exclusively.  Negative indices are translated into
' non-negative ones, with -1 mapping to UBound.
'
' Normally, the function normalize TheIndex, assuming it refers to the first dimension.
' However, when the optional Boolean flag NormalizeWithRespectToColumnsQ is passed with
' value of True, TheIndex is normalized with respect to the second dimension.  Clearly,
' this makes sense only when AnArray has at least two dimensions.  The function
' returns Null if this is not the case.
'
' If the optional boolean paramater NormalizeTo1Q is set to True, then index normalization
' happens assuming 1 is the first positive index in the array.
'
' ***HERE Test the case when NormalizeWithRespectToColumnsQ = True
Public Function NormalizeArrayIndex(AnArray As Variant, _
                                    TheIndex As Long, _
                                    Optional NormalizeWithRespectToColumnsQ As Boolean = False, _
                                    Optional NormalizeTo1Q As Boolean = False) As Variant
    Dim IndexOffset As Long
                                    
    If Not DimensionedQ(AnArray) Then
        Let NormalizeArrayIndex = Null
        Exit Function
    End If
    
    If EmptyArrayQ(AnArray) Then
        Let NormalizeArrayIndex = Null
        Exit Function
    End If
    
    If NormalizeWithRespectToColumnsQ And NumberOfDimensions(AnArray) < 2 Then
         Let NormalizeArrayIndex = Null
        Exit Function
    End If

    ' Case when NormalizeWithRespectToColumnsQ is passed explicitly with a value of True
    If NormalizeWithRespectToColumnsQ Then
        If NonNegativeWholeNumberQ(TheIndex) Then
            If TheIndex < LBound(AnArray, 2) Then
                Let NormalizeArrayIndex = LBound(AnArray, 2)
            ElseIf TheIndex > UBound(AnArray, 1) Then
                Let NormalizeArrayIndex = UBound(AnArray, 2)
            Else
                Let NormalizeArrayIndex = TheIndex
            End If
        Else
            If TheIndex < -NumberOfColumns(AnArray) Then
                Let NormalizeArrayIndex = LBound(AnArray, 2)
            Else
                Let NormalizeArrayIndex = UBound(AnArray, 2) + TheIndex + 1
            End If
        End If
    End If
    
    ' Default case
    If NonNegativeWholeNumberQ(TheIndex) Then
        If NormalizeTo1Q Then
            If TheIndex <= 1 Then
                Let NormalizeArrayIndex = 1
            ElseIf TheIndex >= GetArrayLength(AnArray) Then
                Let NormalizeArrayIndex = GetArrayLength(AnArray)
            Else
                Let NormalizeArrayIndex = TheIndex - 1 + LBound(AnArray, 1)
            End If
            
            Exit Function
        End If
    
        If TheIndex <= LBound(AnArray, 1) Then
            Let NormalizeArrayIndex = LBound(AnArray, 1)
        ElseIf TheIndex >= UBound(AnArray, 1) Then
            Let NormalizeArrayIndex = UBound(AnArray, 1)
        Else
            Let NormalizeArrayIndex = TheIndex
        End If
        
        Exit Function
    Else
        If TheIndex < -GetArrayLength(AnArray) Then
            Let NormalizeArrayIndex = LBound(AnArray, 1)
        Else
            Let NormalizeArrayIndex = UBound(AnArray, 1) + TheIndex + 1
        End If
    End If
    
    Exit Function
End Function

' Same thing as ArrayFormulas.NormalizeArrayIndex, but for an array of indices
Public Function NormalizeArrayIndices(AnArray As Variant, _
                                      TheIndices() As Long, _
                                      Optional NormalizeTo1Q As Boolean = False) As Variant
    Dim r As Long
    Dim ResultsArray() As Variant
    
    If Not DimensionedQ(AnArray) Then
        Let NormalizeArrayIndices = Null
        Exit Function
    End If
    
    If EmptyArrayQ(AnArray) Then
        Let NormalizeArrayIndices = Null
        Exit Function
    End If

    ReDim ResultsArray(LBound(TheIndices) To UBound(TheIndices))
    For r = LBound(TheIndices) To UBound(TheIndices)
        Let ResultsArray(r) = NormalizeArrayIndex(AnArray:=AnArray, _
                                                  TheIndex:=TheIndices(r), _
                                                  NormalizeTo1Q:=NormalizeTo1Q)
    Next
    
    Let NormalizeArrayIndices = ResultsArray
End Function

' DESCRIPTION
' This function translates an array index given in a convention to the array's.
' For instance, AnArray has LBound=0 and UBound=2 and we want to convert TheIndex 1
' with the convention that LBound=1.  Then, this function maps TheIndex to 0.
'
' Negative indices map -1 to the last element of the array.
'
' This function provides an easy to refer to array positions using 1 to reference
' the nth element in the array regardless of how the indexing convention of the
' array.  For example, the second element in the array is at index
' TranslateIndex(AnArray, 2, 1)
Public Function TranslateIndices(AnArray As Variant, TheIndexOrIndexArray As Variant, FromLBound As Long) As Variant
    Dim TheResult As Variant
    Dim ResultArray() As Variant
    Dim r As Long

    ' Exit with Null if AnArray is either not an array or has not been dimensioned.
    If Not DimensionedQ(AnArray) Then
        Let TranslateIndex = Null
        Exit Function
    End If

    ' Exit if AnArray is the empty 1D array
    If EmptyArrayQ(AnArray) Then
        Let TranslateIndex = Null
        Exit Function
    End If
    
    ' Handle the case of TheIndexOrIndexArray being an array of indices
    If IsArray(TheIndexOrIndexArray) Then
        If Not DimensionedQ(TheIndexOrIndexArray) Then
            Let TranslateIndex = Null
            Exit Function
        End If
    
        ' Exit if AnArray is the empty 1D array
        If EmptyArrayQ(TheIndexOrIndexArray) Then
            Let TranslateIndex = Null
            Exit Function
        End If
        
        ' Exit if any of the elements in TheIndexOrIndexArray is not a whole number
        If Not WholeNumberArrayQ(TheIndexOrIndexArray) Then
            Let TranslateIndex = Null
            Exit Function
        End If
        
        ' Recurse of each index in array TheIndexOrIndexArray
        ReDim ResultArray(LBound(TheIndexOrIndexArray, 1) To UBound(TheIndexOrIndexArray, 1))
        For r = LBound(TheIndexOrIndexArray, 1) To UBound(TheIndexOrIndexArray, 1)
            Let ResultArray(r) = TranslateIndices(AnArray, TheIndexOrIndexArray(r), FromLBound)
        Next
        
        ' Return the array of translated indices
        Let TranslateIndices = ResultArray
        
        Exit Function
    End If
    
    ' Handles case of TheIndex non-negative
    If NonNegativeWholeNumberQ(TheIndexOrIndexArray) Then
        If TheIndex <= FromLBound Then
            Let TheResult = FromLBound
        ElseIf TheIndex >= FromLBound + GetArrayLength(AnArray) - 1 Then
            Let TheResult = FromLBound + GetArrayLength(AnArray) - 1
        Else
            Let TheResult = TheIndexOrIndexArray
        End If
        
        Let TranslateIndex = TheResult + LBound(AnArray, 1) - FromLBound
        Exit Function
    End If

    ' Handles case of TheIndex negative
    If TheIndex <= -GetArrayLength(AnArray) Then
        Let TranslateIndex = LBound(AnArray, 1)
    Else
        Let TranslateIndex = TheIndex + 1 + UBound(AnArray, 1)
    End If
End Function
