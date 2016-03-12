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
' Each of the elemements (inner arrays) of the outermost array satisfies RowVectorQ
'
' This is useful to quickly build a matrix from 1D arrays
' This function assumes that all elements of args have the same lbound()
' arg is allowed to have different lbound from that of its elements
'
' The 2D array returned is indexed starting at 1
Public Function Pack2DArray(args As Variant) As Variant
    Dim RowSize As Long
    Dim var As Variant
    Dim r As Long
    Dim c As Long
    Dim Results As Variant
    Dim IndexOffset As Long
    Dim StartIndex As Long
    Dim EndIndex As Long
    
    ' Exit if the argument is not the expected type
    If NumberOfDimensions(args) <> 1 Or EmptyArrayQ(args) Then
        Let Pack2DArray = Empty
        Exit Function
    End If
    
    Let RowSize = GetArrayLength(First(args))
        
    For Each var In args
        If Not RowVectorQ(var) Or GetArrayLength(var) <> RowSize Then
            Let Pack2DArray = Empty
            Exit Function
        End If
    Next

    ' Pre-allocate a 2D array filled with Empty
    Let Results = ConstantArray(Empty, GetArrayLength(args), RowSize)

    ' Get the starting/ending indices and the index offset
    Let StartIndex = LBound(First(args))
    Let EndIndex = UBound(First(args))
    Let IndexOffset = LBound(First(args)) - LBound(args)
    
    ' Pack the array
    For r = LBound(args) To UBound(args)
        For c = StartIndex To EndIndex
            Let Results(IndexOffset + r, IndexOffset + c) = args(r)(c)
        Next c
    Next r
    
    Let Pack2DArray = Results
End Function

' If n is an integer:
' A negative n is interpreted as counting starting with 1 from the right of the 1D array or the bottom of the 2D array respectively
' Returns the first n elements of a 1D array or n rows from a 2D array
' Empty arrays returns an empty array (e.g. Array())
' If AnArray is not an array, function returns Empty
' If AnArray has more than 2 dimensions, parameter AnArray is returned unevaluated
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
    Dim TheResult As Variant
    Dim ResultArray() As Variant
    Dim var As Variant
    Dim RenormalizedIndices As Variant
    
    ' Exit with argument unchanged if the array has fewer or more than 2 dimensions
    If Not IsArray(AnArray) Or NumberOfDimensions(AnArray) = 0 Or NumberOfDimensions(AnArray) > 2 Then
        Let Take = AnArray

        Exit Function
    End If

    ' Exit with argument unchanged if either N is not an integer or an array of integers
    If Not (IsNumeric(N) Or IsNumericArrayQ(N)) Or EmptyArrayQ(N) Then
        Let Take = Array()
        
        Exit Function
    End If
    
    If IsNumericArrayQ(N) Then
        For Each var In N
            If CLng(var) <> var Then
                Let Take = Array()
                
                Exit Function
            End If
        Next
    End If
    
    ' Proceed with N as an array of integer indices
    If IsArray(N) Then
        ' Exit if any of the row indices is invalid
        For c = LBound(N) To UBound(N)
            If Abs(N(c)) < LBound(AnArray) Or Abs(N(c)) > UBound(AnArray) Then
                Let Take = Array()
                
                Exit Function
            End If
        Next c
        
        ' Turn all indices in N into their positive equivalents
        Let RenormalizedIndices = ConstantArray(Empty, GetArrayLength(N))
        Let c = 1
        For Each var In N
            Let RenormalizedIndices(c) = IIf(var < 0, IIf(NumberOfDimensions(AnArray) = 1, GetNumberOfColumns(AnArray), GetNumberOfRows(AnArray)) + var + 1, var)
            Let c = c + 1
        Next
        
        ' Proceed if Anrray is 1D
        If NumberOfDimensions(AnArray) = 1 Then
            Let TheResult = ConstantArray(Empty, GetArrayLength(N))
            
            Let c = 1
            For Each var In RenormalizedIndices
                Let TheResult(c) = AnArray(var)
                
                Let c = c + 1
            Next
            
            Let Take = TheResult
            
            Exit Function
        End If
    
        ' Proceed here if AnArray is 2D
        ' Pre-allocate a matrix big enough to hold all the requested elements
        ReDim ResultArray(GetArrayLength(N), GetNumberOfColumns(AnArray))
        
        Let r = 1
        For Each var In RenormalizedIndices
            For c = 1 To GetNumberOfColumns(AnArray)
                Let ResultArray(r, c) = AnArray(var, IIf(LBound(AnArray, 2) = 0, c - 1, c))
            Next c
            
            Let r = r + 1
        Next
        
        Let Take = ResultArray
    
    ' Proceed with N as an integer index
    Else
        If N = 0 Then
            Let Take = Array()
        ElseIf NumberOfDimensions(AnArray) = 1 And N > 0 Then
            Let Take = GetSubArray(AnArray, LBound(AnArray), Application.Min(LBound(AnArray) + N - 1, UBound(AnArray)))
        ElseIf NumberOfDimensions(AnArray) = 1 And N < 0 Then
            Let Take = GetSubArray(AnArray, Application.Max(UBound(AnArray) + N + 1, LBound(AnArray)), UBound(AnArray))
        ElseIf NumberOfDimensions(AnArray) = 2 And N > 0 Then
            Let Take = GetSubMatrix(AnArray, LBound(AnArray, 1), Application.Min(LBound(AnArray, 1) + N - 1, UBound(AnArray, 1)), LBound(AnArray, 2), UBound(AnArray, 2))
        ElseIf NumberOfDimensions(AnArray) = 2 And N < 0 Then
            Let Take = GetSubMatrix(AnArray, Application.Max(UBound(AnArray, 1) + N + 1, LBound(AnArray, 1)), UBound(AnArray, 1), LBound(AnArray, 2), UBound(AnArray, 2))
        End If
    End If
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

'Public Function Drop(AnArray As Variant, N As Variant)
'    Dim LeftPart As Variant
'    Dim RightPart As Variant
'    Dim var As Variant
'    Dim TempArray As Variant
'
'    ' Exit conditions
'    If Not IsArray(AnArray) Then
'        Let Drop = AnArray
'    ElseIf EmptyArrayQ(AnArray) Then
'        Let Drop = AnArray
'    ElseIf NumberOfDimensions(AnArray) > 2 Or NumberOfDimensions(AnArray) = 0 Then
'        Let Drop = AnArray
'    ElseIf NumberOfDimensions(N) = 0 And Not IsNumeric(N) Then
'        Let Drop = AnArray
'    ElseIf NumberOfDimensions(N) = 1 And Not IsNumericArrayQ(N) Then
'        Let Drop = AnArray
'    ElseIf NumberOfDimensions(N) > 1 Then
'        Let Drop = AnArray
'
'    ' One-dimensional arrays
'    ElseIf NumberOfDimensions(AnArray) = 1 Then
'        If NumberOfDimensions(N) = 0 Then
'            If N > LBound(AnArray) And N < UBound(AnArray) Then
'                Let LeftPart = Take(AnArray, N - 1)
'                Let RightPart = Take(AnArray, -GetArrayLength(AnArray) + N)
'                Let Drop = ConcatenateArrays(LeftPart, RightPart)
'            ElseIf N = LBound(AnArray) Then
'                Let Drop = Take(AnArray, -GetArrayLength(AnArray) + 1)
'            ElseIf N = UBound(AnArray) Then
'                Let Drop = Take(AnArray, GetArrayLength(AnArray) - 1)
'            ElseIf Abs(N) < GetArrayLength(AnArray) And Abs(N) > 1 Then
'                Let LeftPart = Take(AnArray, GetArrayLength(AnArray) + N)
'                Let RightPart = Take(AnArray, N + 1)
'                Let Drop = ConcatenateArrays(LeftPart, RightPart)
'            ElseIf Abs(N) = GetArrayLength(AnArray) Then
'                Let Drop = Take(AnArray, 1 - GetArrayLength(AnArray))
'            ElseIf Abs(N) = 1 Then
'                Let Drop = Take(AnArray, GetArrayLength(AnArray) - 1)
'            Else
'                Let Drop = AnArray
'            End If
'        ElseIf IsNumericArrayQ(N) Then
'            Let TempArray = AnArray
'            For Each var In N
'                If var > LBound(TempArray) And N < UBound(TempArray) Then
'                    Let LeftPart = Take(TempArray, var - 1)
'                    Let RightPart = Take(TempArray, -GetArrayLength(TempArray) + var)
'                    Let TempArray = ConcatenateArrays(LeftPart, RightPart)
'                ElseIf var = LBound(TempArray) Then
'                    Let TempArray = Take(TempArray, -GetArrayLength(TempArray) + 1)
'                ElseIf var = UBound(TempArray) Then
'                    Let TempArray = Take(TempArray, GetArrayLength(TempArray) - 1)
'                End If
'            Next
'
'            Let Drop = TempArray
'        Else
'            Let Drop = AnArray
'        End If
'
'    ' Two-dimensional arrays
'    Else
'        If NumberOfDimensions(N) = 0 Then
'            If N > LBound(AnArray) And N < UBound(AnArray) Then
'                Let LeftPart = Take(AnArray, N - 1)
'                Let RightPart = Take(AnArray, -GetArrayLength(AnArray) + N)
'                Let Drop = Stack2DArrays(LeftPart, RightPart)
'            ElseIf N = LBound(AnArray) Then
'                Let Drop = Take(AnArray, -GetArrayLength(AnArray) + 1)
'            ElseIf N = UBound(AnArray) Then
'                Let Drop = Take(AnArray, GetArrayLength(AnArray) - 1)
'            ElseIf Abs(N) < GetArrayLength(AnArray) And Abs(N) > 1 Then
'                Let LeftPart = Take(AnArray, GetArrayLength(AnArray) + N)
'                Let RightPart = Take(AnArray, N + 1)
'                Let Drop = Stack2DArrays(LeftPart, RightPart)
'            ElseIf Abs(N) = GetArrayLength(AnArray) Then
'                Let Drop = Take(AnArray, 1 - GetArrayLength(AnArray))
'            ElseIf Abs(N) = 1 Then
'                Let Drop = Take(AnArray, GetArrayLength(AnArray) - 1)
'            Else
'                Let Drop = AnArray
'            End If
'        ElseIf IsNumericArrayQ(N) Then
'            Let TempArray = AnArray
'            For Each var In N
'                If var > LBound(TempArray) And N < UBound(TempArray) Then
'                    Let LeftPart = Take(TempArray, var - 1)
'                    Let RightPart = Take(TempArray, -GetArrayLength(TempArray) + var)
'                    Let TempArray = Stack2DArrays(LeftPart, RightPart)
'                ElseIf var = LBound(TempArray) Then
'                    Let TempArray = Take(TempArray, -GetArrayLength(TempArray) + 1)
'                ElseIf var = UBound(TempArray) Then
'                    Let TempArray = Take(TempArray, GetArrayLength(TempArray) - 1)
'                End If
'            Next
'
'            Let Drop = TempArray
'        Else
'            Let Drop = AnArray
'        End If
'    End If
'End Function

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

' This function sorts the given 2D matrix by the columns whose positions are given by ' ArrayOfColPos.
' The sorting orientation in each column are in ArrayOfColsSortOrder
' ArrayOfColsSortOrder is a variant array whose elements are all of enumerated type XLSortOrder (e.g. xlAscending, xlDescending)
' ***HERE
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

' This function performs matrix addition on two 2D arrays of the same dimensions
'***HERE
Public Function MAdd(matrix1 As Variant, matrix2 As Variant) As Variant
    Dim TmpSheet As Worksheet
    Dim targetMatrix1Address As String
    Dim targetMatrix2Address As String
    Dim resultMatrixAddress As String
    
    ' Set pointer to temp sheet and clear its used range
    Set TmpSheet = ThisWorkbook.Worksheets("TempComputation")
    TmpSheet.UsedRange.ClearContents

    ' Dump matrices into temp sheet
    Let targetMatrix1Address = TmpSheet.Range("A1").Resize(UBound(matrix1, 1), UBound(matrix1, 2)).Address
    Let TmpSheet.Range(targetMatrix1Address).Value2 = matrix1
    
    Let targetMatrix2Address = TmpSheet.Range(targetMatrix1Address).Offset(0, UBound(matrix1, 2)).Address
    Let TmpSheet.Range(targetMatrix2Address).Value2 = matrix2
    
    ' Inset array formula to add the two matrices
    Let resultMatrixAddress = TmpSheet.Range(targetMatrix1Address).Offset(0, 2 * UBound(matrix1, 2)).Address
    Let TmpSheet.Range(resultMatrixAddress).FormulaArray = "= " & targetMatrix1Address & "+" & targetMatrix2Address
    
    ' Import result of adding the two matrices
    Let MAdd = TmpSheet.Range(resultMatrixAddress).Value2
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
                If IsArray(a) Then Let ConvertTo1DArray = ConvertTo1DArray(a(LBound(a)))
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
                Let TheResult(i) = a(i, 1)
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
Public Function GetRow(AMatrix As Variant, RowNumber As Long) As Variant
    Dim nd As Long
    Dim i As Long
    Dim TheResult() As Variant

    If EmptyArrayQ(AMatrix) Then
        Let GetRow = Array()
        Exit Function
    End If
    
    Let nd = NumberOfDimensions(AMatrix)
    
    If nd = 1 Then
        If RowNumber = 1 Then
            Let GetRow = AMatrix
        Else
            Let GetRow = Array()
        End If
        
        Exit Function
    End If
    
    If RowNumber < 1 Or RowNumber > GetNumberOfRows(AMatrix) Then
        Let GetRow = Array()
        Exit Function
    End If
    
    ReDim TheResult(1 To GetNumberOfColumns(AMatrix))
    For i = 1 To GetNumberOfColumns(AMatrix)
        Let TheResult(i) = AMatrix(IIf(LBound(AMatrix, 1) = 0, RowNumber - 1, RowNumber), IIf(LBound(AMatrix, 2) = 0, i - 1, i))
    Next i
    
    Let GetRow = TheResult
End Function

Public Function GetSubRow(AMatrix As Variant, RowNumber As Long, StartColumn As Long, EndColumn As Long) As Variant
    Dim AnArray As Variant
    
    Let AnArray = GetRow(AMatrix, RowNumber)
    
    Let GetSubRow = GetSubArray(AnArray, StartColumn, EndColumn)
End Function

' Works on both 1D and 2D arrays, returning what makes sense (e.g. 1D arrays are interpreted as 1-row 2D arrays)
' Regardless of the array's indexing convention (e.g. start at 0 or 1), the user must refer to the first column as column 1.
' The result has the same number of dimensions as the argument
Public Function GetColumn(AMatrix As Variant, ColumnNumber As Long) As Variant
    Dim i As Long
    Dim TheResults() As Variant
    Dim NDims As Integer

    If EmptyArrayQ(AMatrix) Then
        Let GetColumn = Array()
        Exit Function
    End If
    
    Let NDims = NumberOfDimensions(AMatrix)
    
    If NDims = 1 Then
        If ColumnNumber > GetNumberOfColumns(AMatrix) Or ColumnNumber < 1 Then
            Let GetColumn = Array()
        Else
            Let GetColumn = Array(AMatrix(IIf(LBound(AMatrix) = 0, ColumnNumber - 1, ColumnNumber)))
        End If
        
        Exit Function
    End If
    
    If ColumnNumber < 1 Or ColumnNumber > GetNumberOfColumns(AMatrix) Then
        Let GetColumn = Array()
        Exit Function
    End If

    ReDim TheResults(GetNumberOfRows(AMatrix), 1 To 1)
    For i = IIf(NDims = 1, LBound(AMatrix), LBound(AMatrix, 1)) To IIf(NDims = 1, UBound(AMatrix), UBound(AMatrix, 1))
        Let TheResults(i, 1) = AMatrix(i, ColumnNumber)
    Next i
    
    Let GetColumn = TheResults
End Function

' Returns the requested subset of an column in a two-dimension array
' The result is returned as a one-dimensional array
Public Function GetSubColumn(AMatrix As Variant, ColumnNumber As Long, StartRow As Long, EndRow As Long) As Variant
    Dim AnArray As Variant
    
    Let AnArray = GetColumn(AMatrix, ColumnNumber)
    
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

' Appends a new element to the given array
' Handles arrays of dimension 1 and 2
' Returns an AnArray unevaluated if AnArray is not an array or it has more than two dims,
' or dim(AnArray) = 2 and AnArray and AnElt don't have the same number of columns
' This works differently from simply using Stack2DArrays(AnArray, AnElt)
' This one can give you something that is not a matrix.  Stack2DArrays ALWAYS returns
' a matrix
Public Function Append(AnArray As Variant, AnElt As Variant) As Variant
    Dim NewArray As Variant
    Dim AnArrayNumberOfDims As Integer
    
    Let AnArrayNumberOfDims = NumberOfDimensions(AnArray)
    
    If Not IsArray(AnArray) Or AnArrayNumberOfDims > 2 Or _
       (AnArrayNumberOfDims = 2 And GetNumberOfColumns(AnArray) <> GetNumberOfColumns(AnElt)) Then
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

' Appends a new element to the given array
' Handles arrays of dimension 1 and 2
' Returns an AnArray unevaluated if AnArray is not an array or it has more than two dims,
' or dim(AnArray) = 2 and AnArray and AnElt don't have the same number of columns
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
Public Function GetSubMatrix(AMatrix As Variant, Optional TopRowNumber As Variant, Optional BottomRowNumber As Variant, _
                             Optional BottomColumnNumber As Variant, Optional TopColumnNumber As Variant)
    If EmptyArrayQ(AMatrix) Then
        Let GetSubMatrix = Array()
        Exit Function
    End If
                             
    If IsMissing(TopRowNumber) And IsMissing(BottomRowNumber) And IsMissing(BottomColumnNumber) And IsMissing(TopColumnNumber) Then
        Let GetSubMatrix = AMatrix
        Exit Function
    End If
    
    If IsMissing(TopRowNumber) And IsMissing(BottomRowNumber) And Not IsMissing(BottomColumnNumber) And Not IsMissing(TopColumnNumber) Then
        Let GetSubMatrix = GetSubMatrixHelper(AMatrix, 1, GetNumberOfRows(AMatrix), CLng(BottomColumnNumber), CLng(TopColumnNumber))
        Exit Function
    End If
    
    If IsMissing(BottomColumnNumber) And IsMissing(TopColumnNumber) And Not IsMissing(TopRowNumber) And Not IsMissing(BottomRowNumber) Then
        Let GetSubMatrix = GetSubMatrixHelper(AMatrix, CLng(TopRowNumber), CLng(BottomRowNumber), 1, GetNumberOfColumns(AMatrix))
        Exit Function
    End If

    Let GetSubMatrix = GetSubMatrixHelper(AMatrix, CLng(TopRowNumber), CLng(BottomRowNumber), CLng(BottomColumnNumber), CLng(TopColumnNumber))
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
    
    ReDim SubMatrix(BottomRowNumber - TopRowNumber + LBound(TheArray, 1), TopColumnNumber - BottomColumnNumber + LBound(TheArray, 2))
    For r = TopRowNumber To BottomRowNumber
        For c = BottomColumnNumber To TopColumnNumber
            Let SubMatrix(r - TopRowNumber + LBound(TheArray, 1), c - BottomColumnNumber + LBound(TheArray, 2)) = TheArray(r, c)
        Next c
    Next r
    
    Let GetSubMatrixHelper = SubMatrix
End Function

' Gets the sub-array between Indices StartIndex and EndIndex from the 1D array AnArray
' If the argument is not an array, then this function returns its argument unevaluated.
' If either StartIndex is below Lbound of the parameter or EndIndex is above UBound then,
' The LBound or UBound respectively is used.
' If the number of dimensions is larger than 1, then the array is returned unevaluated.
' An atomic expression is returned as a 1D array with the argument as its only argument.
' If EndIndex<StartIndex then the argument is returned unevaluated.
Public Function GetSubArray(AnArray As Variant, StartIndex As Long, EndIndex As Long) As Variant
    Dim i As Long
    Dim ReturnedArray() As Variant
    
    If Not IsArray(AnArray) Or EndIndex < StartIndex Then
        Let GetSubArray = AnArray
        Exit Function
    End If
    
    If NumberOfDimensions(AnArray) = 0 Then
        Let GetSubArray = Array(AnArray)
    ElseIf NumberOfDimensions(AnArray) = 1 Then
        ReDim ReturnedArray(1 To IIf(EndIndex > UBound(AnArray), UBound(AnArray), EndIndex) - IIf(StartIndex < LBound(AnArray), LBound(AnArray), StartIndex) + 1)
    
        For i = StartIndex To EndIndex
            Let ReturnedArray(i - StartIndex + 1) = AnArray(i)
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
' If aMatrix has more than two dimensions, the function returns Nothing
Public Function UniqueSubset(AMatrix As Variant) As Variant
    Dim r As Long
    Dim c As Long
    Dim UniqueDict As Dictionary
    Dim Dimensionality As Integer
    
    ' Exit, returning an empty array if aMatrix is empty
    If EmptyArrayQ(AMatrix) Or IsEmpty(AMatrix) Then
        Let UniqueSubset = Array()
        Exit Function
    End If
    
    Set UniqueDict = New Dictionary
    Let Dimensionality = NumberOfDimensions(AMatrix)
    
    If Dimensionality = 0 Then
        Let UniqueSubset = AMatrix
        
        Exit Function
    ElseIf Dimensionality = 1 Then
        For r = LBound(AMatrix) To UBound(AMatrix)
            If Not UniqueDict.Exists(AMatrix(r)) Then
                Call UniqueDict.Add(Key:=AMatrix(r), Item:=1)
            End If
        Next r
    ElseIf Dimensionality = 2 Then
        For r = LBound(AMatrix, 1) To UBound(AMatrix, 1)
            For c = LBound(AMatrix, 2) To UBound(AMatrix, 2)
                If Not UniqueDict.Exists(AMatrix(r, c)) Then
                    Call UniqueDict.Add(Key:=AMatrix(r, c), Item:=1)
                End If
            Next c
        Next r
    Else
        Set UniqueSubset = Nothing
        
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
Public Function Stack2DArrayAs1DArray(AMatrix As Variant) As Variant
    Dim TheResults() As Variant
    Dim var As Variant
    Dim j As Long

    If EmptyArrayQ(AMatrix) Or Not IsArray(AMatrix) Or NumberOfDimensions(AMatrix) > 2 Then
        Let Stack2DArrayAs1DArray = Array()
        Exit Function
    ElseIf NumberOfDimensions(AMatrix) = 0 Then
        Let Stack2DArrayAs1DArray = AMatrix
        Exit Function
    End If
    
    If NumberOfDimensions(AMatrix) = 1 Then
        ReDim TheResults(GetArrayLength(AMatrix))
    Else
        ReDim TheResults(GetNumberOfRows(AMatrix) * GetNumberOfColumns(AMatrix))
    End If
    
    Let j = 1
    For Each var In AMatrix
        Let TheResults(j) = var
        Let j = j + 1
    Next
    
    Let Stack2DArrayAs1DArray = TheResults
End Function

' Alias for Stack2DArrayAs1DArray
Public Function StackArrayAs1DArray(AMatrix As Variant) As Variant
    Let StackArrayAs1DArray = Stack2DArrayAs1DArray(AMatrix)
End Function

' This function dumps an array (1D or 2D) into worksheet TempComputation and then returns a reference to
' the underlying range.  Dimensions are preserved.  This means that an m x n array is dumped into an m x n
' range.  This function should not be used if leading single quotes (e.g "'") are part of the array's elements.
Public Function ToTemp(AnArray As Variant, Optional PreserveColumnTextFormats As Boolean = False) As Range
    Dim c As Integer
    Dim NumberOfRows As Long
    Dim NumberOfColumns As Integer
    Dim NumDimensions As Integer
    
    ' Exit if AnArray is empty
    If EmptyArrayQ(AnArray) Or IsEmpty(AnArray) Then
        Set ToTemp = Nothing
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
' This one handles empty arrays correctly, by doing nothing and returning a reference to Nothing
Public Function DumpInSheet(AnArray As Variant, TopLeftCorner As Range, Optional PreserveColumnTextFormats As Boolean = False) As Range
    If EmptyArrayQ(AnArray) Or IsEmpty(AnArray) Or TopLeftCorner Is Nothing Then
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

' This function performs matrix element-wise addition on two 0D, 1D, 2D arrays.  Clearly, the two arrays
' must have the same dimensions.  The result is returned as an array of the same dimensions as those of
' the input.
Public Function ElementWiseAddition(matrix1 As Variant, matrix2 As Variant) As Variant
    Dim TmpSheet As Worksheet
    Dim r As Long ' for number of rows
    Dim c As Long ' for number of columns
    
    ' Set pointer to temp sheet and clear its used range
    Set TmpSheet = ThisWorkbook.Worksheets("TempComputation")
    Call TmpSheet.Cells.ClearFormats
    Call TmpSheet.UsedRange.ClearContents

    ' Store number of rows and columns base on number of dimensions
    If NumberOfDimensions(matrix1) = 0 And NumberOfDimensions(matrix2) = 0 Then
        ' Return the result
        Let ElementWiseAddition = matrix1 + matrix2
    ElseIf NumberOfDimensions(matrix1) = 0 And NumberOfDimensions(matrix2) = 1 Then
        ' Get matrix dimensions
        Let r = 1
        Let c = UBound(matrix2) - LBound(matrix2) + 1
        
        ' Dump matrices into temp sheet
        Let TmpSheet.Range("A1").Resize(r, c).Value2 = matrix2
        
        ' Drop formula to add elementwise
        Let TmpSheet.Range("A1").Offset(1, 0).Resize(r, c).FormulaR1C1 = "=R[" & -r & "]C[0] + " & matrix1
    
        ' Return the result
        Let ElementWiseAddition = ConvertTo1DArray(TmpSheet.Range("A1").Offset(r, 0).Resize(r, c).Value2)
    ElseIf NumberOfDimensions(matrix1) = 1 And NumberOfDimensions(matrix2) = 0 Then
        ' Get matrix dimensions
        Let r = 1
        Let c = UBound(matrix1) - LBound(matrix1) + 1
        
        ' Dump matrices into temp sheet
        Let TmpSheet.Range("A1").Resize(r, c).Value2 = matrix1
        
        ' Drop formula to add elementwise
        Let TmpSheet.Range("A1").Offset(1, 0).Resize(r, c).FormulaR1C1 = "=R[" & -r & "]C[0] + " & matrix2
    
        ' Return the result
        Let ElementWiseAddition = ConvertTo1DArray(TmpSheet.Range("A1").Offset(r, 0).Resize(r, c).Value2)
    ElseIf NumberOfDimensions(matrix1) = 1 And NumberOfDimensions(matrix2) = 1 Then
        ' Get matrix dimensions
        Let r = 1
        Let c = UBound(matrix1) - LBound(matrix1) + 1
        
        ' Dump matrices into temp sheet
        Let TmpSheet.Range("A1").Resize(r, c).Value2 = matrix1
        Let TmpSheet.Range("A1").Offset(r, 0).Resize(r, c).Value2 = matrix2
        
        ' Drop formula to add elementwise
        Let TmpSheet.Range("A1").Offset(2 * r, 0).Resize(r, c).FormulaR1C1 = "=R[" & -2 * r & "]C[0] + R[" & -r & "]C[0]"
    
        ' Return the result
        Let ElementWiseAddition = ConvertTo1DArray(TmpSheet.Range("A1").Offset(2 * r, 0).Resize(r, c).Value2)
    ElseIf NumberOfDimensions(matrix1) = 0 And NumberOfDimensions(matrix2) = 2 Then
        ' Get matrix dimensions
        Let r = UBound(matrix2, 1) - LBound(matrix2, 1) + 1
        Let c = UBound(matrix2, 2) - LBound(matrix2, 2) + 1
        
        ' Dump matrices into temp sheet
        Let TmpSheet.Range("A1").Resize(r, c).Value2 = matrix2
        
        ' Drop formula to add elementwise
        Let TmpSheet.Range("A1").Offset(r, 0).Resize(r, c).FormulaR1C1 = "=R[" & -r & "]C[0] + " & matrix1
    
        ' Return the result
        Let ElementWiseAddition = TmpSheet.Range("A1").Offset(r, 0).Resize(r, c).Value2
    ElseIf NumberOfDimensions(matrix1) = 2 And NumberOfDimensions(matrix2) = 0 Then
        ' Get matrix dimensions
        Let r = UBound(matrix1, 1) - LBound(matrix1, 1) + 1
        Let c = UBound(matrix1, 2) - LBound(matrix1, 2) + 1
        
        ' Dump matrices into temp sheet
        Let TmpSheet.Range("A1").Resize(r, c).Value2 = matrix1
        
        ' Drop formula to add elementwise
        Let TmpSheet.Range("A1").Offset(r, 0).Resize(r, c).FormulaR1C1 = "=R[" & -r & "]C[0] + " & matrix2
    
        ' Return the result
        Let ElementWiseAddition = TmpSheet.Range("A1").Offset(r, 0).Resize(r, c).Value2
    Else
        ' Get matrix dimensions
        Let r = UBound(matrix1, 1) - LBound(matrix1, 1) + 1
        Let c = UBound(matrix1, 2) - LBound(matrix1, 2) + 1
        
        ' Dump matrices into temp sheet
        Let TmpSheet.Range("A1").Resize(r, c).Value2 = matrix1
        Let TmpSheet.Range("A1").Offset(r, 0).Resize(r, c).Value2 = matrix2
        
        ' Drop formula to add elementwise
        Let TmpSheet.Range("A1").Offset(2 * r, 0).Resize(r, c).FormulaR1C1 = "=R[" & -2 * r & "]C[0] + R[" & -r & "]C[0]"
    
        ' Return the result
        Let ElementWiseAddition = TmpSheet.Range("A1").Offset(2 * r, 0).Resize(r, c).Value2
    End If
End Function

' This function performs matrix element-wise multiplication on two 0D, 1D, or 2D arrays.  Clearly, the two arrays
' must have the same dimensions.  The result is returned as an array of the same dimensions as those of
' the input.
Public Function ElementWiseMultiplication(matrix1 As Variant, matrix2 As Variant) As Variant
    Dim TmpSheet As Worksheet
    Dim r As Long ' for number of rows
    Dim c As Long ' for number of columns
    
    ' Set pointer to temp sheet and clear its used range
    Set TmpSheet = ThisWorkbook.Worksheets("TempComputation")
    Call TmpSheet.Cells.ClearFormats
    Call TmpSheet.UsedRange.ClearContents

    ' Store number of rows and columns base on number of dimensions
    If NumberOfDimensions(matrix1) = 0 And NumberOfDimensions(matrix2) = 0 Then
        ' Return the result
        Let ElementWiseMultiplication = matrix1 * matrix2
    ElseIf NumberOfDimensions(matrix1) = 0 And NumberOfDimensions(matrix2) = 1 Then
        ' Get matrix dimensions
        Let r = 1
        Let c = UBound(matrix2) - LBound(matrix2) + 1
        
        ' Dump matrices into temp sheet
        Let TmpSheet.Range("A1").Resize(r, c).Value2 = matrix2
        
        ' Drop formula to add elementwise
        Let TmpSheet.Range("A1").Offset(1, 0).Resize(r, c).FormulaR1C1 = "=R[" & -r & "]C[0] * " & matrix1
    
        ' Return the result
        Let ElementWiseMultiplication = TmpSheet.Range("A1").Offset(r, 0).Resize(r, c).Value2
    ElseIf NumberOfDimensions(matrix1) = 1 And NumberOfDimensions(matrix2) = 0 Then
        ' Get matrix dimensions
        Let r = 1
        Let c = UBound(matrix1) - LBound(matrix1) + 1
        
        ' Dump matrices into temp sheet
        Let TmpSheet.Range("A1").Resize(r, c).Value2 = matrix1
        
        ' Drop formula to add elementwise
        Let TmpSheet.Range("A1").Offset(1, 0).Resize(r, c).FormulaR1C1 = "=R[" & -r & "]C[0] * " & matrix2
    
        ' Return the result
        Let ElementWiseMultiplication = TmpSheet.Range("A1").Offset(r, 0).Resize(r, c).Value2
    ElseIf NumberOfDimensions(matrix1) = 1 And NumberOfDimensions(matrix2) = 1 Then
        ' Get matrix dimensions
        Let r = 1
        Let c = UBound(matrix1) - LBound(matrix1) + 1
        
        ' Dump matrices into temp sheet
        Let TmpSheet.Range("A1").Resize(r, c).Value2 = matrix1
        Let TmpSheet.Range("A1").Offset(r, 0).Resize(r, c).Value2 = matrix2
        
        ' Drop formula to add elementwise
        Let TmpSheet.Range("A1").Offset(2 * r, 0).Resize(r, c).FormulaR1C1 = "=R[" & -2 * r & "]C[0] * R[" & -r & "]C[0]"
    
        ' Return the result
        Let ElementWiseMultiplication = TmpSheet.Range("A1").Offset(2 * r, 0).Resize(r, c).Value2
    ElseIf NumberOfDimensions(matrix1) = 0 And NumberOfDimensions(matrix2) = 2 Then
        ' Get matrix dimensions
        Let r = UBound(matrix2, 1) - LBound(matrix2, 1) + 1
        Let c = UBound(matrix2, 2) - LBound(matrix2, 2) + 1
        
        ' Dump matrices into temp sheet
        Let TmpSheet.Range("A1").Resize(r, c).Value2 = matrix2
        
        ' Drop formula to add elementwise
        Let TmpSheet.Range("A1").Offset(r, 0).Resize(r, c).FormulaR1C1 = "=R[" & -r & "]C[0] * " & matrix1
    
        ' Return the result
        Let ElementWiseMultiplication = TmpSheet.Range("A1").Offset(r, 0).Resize(r, c).Value2
    ElseIf NumberOfDimensions(matrix1) = 2 And NumberOfDimensions(matrix2) = 0 Then
        ' Get matrix dimensions
        Let r = UBound(matrix1, 1) - LBound(matrix1, 1) + 1
        Let c = UBound(matrix1, 2) - LBound(matrix1, 2) + 1
        
        ' Dump matrices into temp sheet
        Let TmpSheet.Range("A1").Resize(r, c).Value2 = matrix1
        
        ' Drop formula to add elementwise
        Let TmpSheet.Range("A1").Offset(r, 0).Resize(r, c).FormulaR1C1 = "=R[" & -r & "]C[0] * " & matrix2
    
        ' Return the result
        Let ElementWiseMultiplication = TmpSheet.Range("A1").Offset(r, 0).Resize(r, c).Value2
    Else
        ' Get matrix dimensions
        Let r = UBound(matrix1, 1) - LBound(matrix1, 1) + 1
        Let c = UBound(matrix1, 2) - LBound(matrix1, 2) + 1
        
        ' Dump matrices into temp sheet
        Let TmpSheet.Range("A1").Resize(r, c).Value2 = matrix1
        Let TmpSheet.Range("A1").Offset(r, 0).Resize(r, c).Value2 = matrix2
        
        ' Drop formula to add elementwise
        Let TmpSheet.Range("A1").Offset(2 * r, 0).Resize(r, c).FormulaR1C1 = "=R[" & -2 * r & "]C[0] * R[" & -r & "]C[0]"
    
        ' Return the result
        Let ElementWiseMultiplication = TmpSheet.Range("A1").Offset(2 * r, 0).Resize(r, c).Value2
    End If
End Function

' This function performs matrix element-wise division on two 0D, 1D, or 2D arrays.  Clearly, the two arrays
' must have the same dimensions.  The result is returned as an array of the same dimensions as those of
' the input.
Public Function ElementWiseDivision(matrix1 As Variant, matrix2 As Variant) As Variant
    Dim TmpSheet As Worksheet
    Dim r As Long ' for number of rows
    Dim c As Long ' for number of columns
    
    ' Set pointer to temp sheet and clear its used range
    Set TmpSheet = ThisWorkbook.Worksheets("TempComputation")
    Call TmpSheet.Cells.ClearFormats
    Call TmpSheet.UsedRange.ClearContents

    ' Store number of rows and columns base on number of dimensions
    If NumberOfDimensions(matrix1) = 0 And NumberOfDimensions(matrix2) = 0 Then
        ' Return the result
        Let ElementWiseDivision = matrix1 / matrix2
    ElseIf NumberOfDimensions(matrix1) = 0 And NumberOfDimensions(matrix2) = 1 Then
        ' Get matrix dimensions
        Let r = 1
        Let c = UBound(matrix2) - LBound(matrix2) + 1
        
        ' Dump matrices into temp sheet
        Let TmpSheet.Range("A1").Resize(r, c).Value2 = matrix2
        
        ' Drop formula to add elementwise
        Let TmpSheet.Range("A1").Offset(1, 0).Resize(r, c).FormulaR1C1 = "=R[" & -r & "]C[0] / " & matrix1
    
        ' Return the result
        Let ElementWiseDivision = TmpSheet.Range("A1").Offset(r, 0).Resize(r, c).Value2
    ElseIf NumberOfDimensions(matrix1) = 1 And NumberOfDimensions(matrix2) = 0 Then
        ' Get matrix dimensions
        Let r = 1
        Let c = UBound(matrix1) - LBound(matrix1) + 1
        
        ' Dump matrices into temp sheet
        Let TmpSheet.Range("A1").Resize(r, c).Value2 = matrix1
        
        ' Drop formula to add elementwise
        Let TmpSheet.Range("A1").Offset(1, 0).Resize(r, c).FormulaR1C1 = "=R[" & -r & "]C[0] / " & matrix2
    
        ' Return the result
        Let ElementWiseDivision = TmpSheet.Range("A1").Offset(r, 0).Resize(r, c).Value2
    ElseIf NumberOfDimensions(matrix1) = 1 And NumberOfDimensions(matrix2) = 1 Then
        ' Get matrix dimensions
        Let r = 1
        Let c = UBound(matrix1) - LBound(matrix1) + 1
        
        ' Dump matrices into temp sheet
        Let TmpSheet.Range("A1").Resize(r, c).Value2 = matrix1
        Let TmpSheet.Range("A1").Offset(r, 0).Resize(r, c).Value2 = matrix2
        
        ' Drop formula to add elementwise
        Let TmpSheet.Range("A1").Offset(2 * r, 0).Resize(r, c).FormulaR1C1 = "=R[" & -2 * r & "]C[0] / R[" & -r & "]C[0]"
    
        ' Return the result
        Let ElementWiseDivision = TmpSheet.Range("A1").Offset(2 * r, 0).Resize(r, c).Value2
    ElseIf NumberOfDimensions(matrix1) = 0 And NumberOfDimensions(matrix2) = 2 Then
        ' Get matrix dimensions
        Let r = UBound(matrix2, 1) - LBound(matrix2, 1) + 1
        Let c = UBound(matrix2, 2) - LBound(matrix2, 2) + 1
        
        ' Dump matrices into temp sheet
        Let TmpSheet.Range("A1").Resize(r, c).Value2 = matrix2
        
        ' Drop formula to add elementwise
        Let TmpSheet.Range("A1").Offset(r, 0).Resize(r, c).FormulaR1C1 = "=R[" & -r & "]C[0] / " & matrix1
    
        ' Return the result
        Let ElementWiseDivision = TmpSheet.Range("A1").Offset(r, 0).Resize(r, c).Value2
    ElseIf NumberOfDimensions(matrix1) = 2 And NumberOfDimensions(matrix2) = 0 Then
        ' Get matrix dimensions
        Let r = UBound(matrix1, 1) - LBound(matrix1, 1) + 1
        Let c = UBound(matrix1, 2) - LBound(matrix1, 2) + 1
        
        ' Dump matrices into temp sheet
        Let TmpSheet.Range("A1").Resize(r, c).Value2 = matrix1
        
        ' Drop formula to add elementwise
        Let TmpSheet.Range("A1").Offset(r, 0).Resize(r, c).FormulaR1C1 = "=R[" & -r & "]C[0] / " & matrix2
    
        ' Return the result
        Let ElementWiseDivision = TmpSheet.Range("A1").Offset(r, 0).Resize(r, c).Value2
    Else
        ' Get matrix dimensions
        Let r = UBound(matrix1, 1) - LBound(matrix1, 1) + 1
        Let c = UBound(matrix1, 2) - LBound(matrix1, 2) + 1
        
        ' Dump matrices into temp sheet
        Let TmpSheet.Range("A1").Resize(r, c).Value2 = matrix1
        Let TmpSheet.Range("A1").Offset(r, 0).Resize(r, c).Value2 = matrix2
        
        ' Drop formula to add elementwise
        Let TmpSheet.Range("A1").Offset(2 * r, 0).Resize(r, c).FormulaR1C1 = "=R[" & -2 * r & "]C[0] / R[" & -r & "]C[0]"
    
        ' Return the result
        Let ElementWiseDivision = TmpSheet.Range("A1").Offset(2 * r, 0).Resize(r, c).Value2
    End If
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

' Returns the number of elements in the first dimension
Public Function GetArrayLength(AMatrix As Variant) As Long
    If IsArray(AMatrix) Then
        Let GetArrayLength = UBound(AMatrix) - LBound(AMatrix) + 1
    Else
        Let GetArrayLength = 0
    End If
End Function

' Returns the number of rows in a 2D array
Public Function GetNumberOfRows(AMatrix As Variant) As Long
    If IsArray(AMatrix) Then
        If NumberOfDimensions(AMatrix) = 1 Then
            Let GetNumberOfRows = 1
        Else
            Let GetNumberOfRows = UBound(AMatrix, 1) - LBound(AMatrix, 1) + 1
        End If
    Else
        Let GetNumberOfRows = 0
    End If
End Function

' Returns the number of columns in a 2D array. A 1D array is considered a 2D array with 1 row
Public Function GetNumberOfColumns(AMatrix As Variant) As Long
    If IsArray(AMatrix) Then
        If NumberOfDimensions(AMatrix) = 1 Then
            Let GetNumberOfColumns = UBound(AMatrix, 1) - LBound(AMatrix, 1) + 1
        Else
            Let GetNumberOfColumns = UBound(AMatrix, 2) - LBound(AMatrix, 2) + 1
        End If
    Else
        Let GetNumberOfColumns = 0
    End If
End Function

' Returns a 1D array with the same length as the input but with all its entries converted to strings.
' The array of strings is returned indexed from 1 to N.
Public Function Cast1DArrayToStrings(TheArray As Variant) As String()
    Dim i As Long
    Dim StringArray() As String
    
    ' Exit with empty array if TheArray is empty
    If EmptyArrayQ(TheArray) Then
        Let Cast1DArrayToStrings = Array()
        Exit Function
    End If
    
    If IsArray(TheArray) Then
        ReDim StringArray(1 To GetArrayLength(TheArray))
    
        For i = LBound(TheArray) To UBound(TheArray)
            Let StringArray(i + (1 - LBound(TheArray))) = CStr(TheArray(i))
        Next i
    
        Let Cast1DArrayToStrings = StringArray
    Else
        Let Cast1DArrayToStrings = Array(CStr(TheArray))
    End If
End Function

' Returns a 1D array with the same length as the input but with all its entries converted to worksheet object references.
' The array of worksheet references is returned indexed from 1 to N.
Public Function Cast1DArrayToWorksheets(VariantWorksheetsArray As Variant) As Worksheet()
    Dim MyArray() As Worksheet
    Dim i As Long
    
    ' Exit with empty array if TheArray is empty
    If EmptyArrayQ(VariantWorksheetsArray) Then
        Let Cast1DArrayToWorksheets = Array()
        Exit Function
    End If

    If IsArray(VariantWorksheetsArray) Then
        ReDim MyArray(1 To GetArrayLength(VariantWorksheetsArray))
        
        For i = 1 To GetArrayLength(VariantWorksheetsArray)
            Set MyArray(i) = VariantWorksheetsArray(IIf(LBound(VariantWorksheetsArray) = 0, i - 1, i))
        Next i
        
        Let Cast1DArrayToWorksheets = MyArray
    Else
        Let Cast1DArrayToWorksheets = Array(VariantWorksheetsArray)
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
    Dim Col1 As Range
    
    If FirstColumnIndex < 1 Or FirstColumnIndex > TheRange.Columns.Count Or SecondColumIndex < 1 Or SecondColumIndex > TheRange.Columns.Count Or FirstColumnIndex = SecondColumIndex Then
        Let SwapRangeColumns = False

        Exit Function
    End If
    
    Call TheRange.Worksheet.Activate
    
    Set Col1 = TheRange.Worksheet.UsedRange
    Set Col1 = Col1.Range("a1").Offset(Col1.Rows.Count, Col1.Columns.Count).Resize(TheRange.Rows.Count, 1)
    
    Call TheRange.Columns(FirstColumnIndex).Copy
    Call Col1.Range("A1").Select
    Call TheRange.Worksheet.Paste
    
    Call TheRange.Columns(SecondColumIndex).Copy
    Call TheRange.Cells(1, FirstColumnIndex).Select
    Call TheRange.Worksheet.Paste
    
    Call Col1.Copy
    Call TheRange.Columns(SecondColumIndex).Select
    Call TheRange.Worksheet.Paste
    
    Call Col1.ClearContents
    Call Col1.ClearFormats
    Call Col1.ClearComments
    Call Col1.ClearHyperlinks
    Call Col1.ClearNotes
    Call Col1.ClearOutline
    
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
    Dim Col1 As Variant
    
    If FirstColumnIndex < 1 Or FirstColumnIndex > GetNumberOfColumns(TheMatrix) Or SecondColumIndex < 1 Or SecondColumIndex > GetNumberOfColumns(TheMatrix) Or FirstColumnIndex = SecondColumIndex Then
        Let SwapMatrixColumns = False

        Exit Function
    End If
    
    Call ToTemp(TheMatrix)
    Let Col1 = GetColumn(TheMatrix, FirstColumnIndex)
    Call DumpInSheet(GetColumn(TheMatrix, SecondColumIndex), TempComputation.Cells(1, FirstColumnIndex))
    Call DumpInSheet(Col1, TempComputation.Cells(1, SecondColumIndex))
    
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

'***TEST
' This function transposes 1D or 2D arrays
Public Function TransposeMatrix(AMatrix As Variant) As Variant
    Dim r As Long
    Dim c As Long
    Dim TheResult() As Variant
    
    If NumberOfDimensions(AMatrix) = 0 Then
        Let TransposeMatrix = AMatrix
        Exit Function
    End If
    
    If EmptyArrayQ(AMatrix) Then
        Let TransposeMatrix = Array()
        Exit Function
    End If
    
    If NumberOfDimensions(AMatrix) = 1 Then
        ReDim TheResult(LBound(AMatrix) To UBound(AMatrix), 1)
    
        For c = LBound(AMatrix) To UBound(AMatrix)
            Let TheResult(c, 1) = AMatrix(c)
        Next c
        
        Let TransposeMatrix = TheResult
        
        Exit Function
    End If
    
    ReDim TheResult(LBound(AMatrix, 2) To UBound(AMatrix, 2), LBound(AMatrix, 1) To UBound(AMatrix, 1))

    For r = LBound(AMatrix, 1) To UBound(AMatrix, 1)
        For c = LBound(AMatrix, 2) To UBound(AMatrix, 2)
            Let TheResult(c, r) = AMatrix(r, c)
        Next c
    Next r
    
    Let TransposeMatrix = TheResult
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
Public Function AddSingleQuotesToAllArrayElements(TheArray As Variant, Optional EscapeSingleQuote As Boolean = True) As Variant '***HERE
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
