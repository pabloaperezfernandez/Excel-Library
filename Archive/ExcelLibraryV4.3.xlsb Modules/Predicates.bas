Attribute VB_Name = "Predicates"
Option Explicit
Option Base 1

' This function returns TRUE only if arg is a 1D array of positive integers
Public Function IsPositiveIntegerArrayQ(Arg As Variant) As Boolean
    Dim var As Variant
    
    Let IsPositiveIntegerArrayQ = True
    
    If Not IsArray(Arg) Then
        Let IsPositiveIntegerArrayQ = False
    ElseIf NumberOfDimensions(Arg) <> 1 Then
        Let IsPositiveIntegerArrayQ = False
    Else
        For Each var In Arg
            ' Exit with False if current item is not a number or string
            ' that can be interpreted as a number
            If Not IsNumeric(var) Then
                Let IsPositiveIntegerArrayQ = False
                Exit For
            End If
            
            ' If the truncation of var is not equalt to var, it is not
            ' and integer.  Exit with False
            If CLng(var) <> var Then
                Let IsPositiveIntegerArrayQ = False
                Exit For
            End If
            
            ' If var is an integer below 1, exit with False
            If CLng(var) < 1 Then
                Let IsPositiveIntegerArrayQ = False
                Exit For
            End If
        Next
    End If
End Function

' This function returns TRUE only if arg is a 1D array of numerics
Public Function IsNumericArrayQ(Arg As Variant) As Boolean
    Dim var As Variant
    
    Let IsNumericArrayQ = True
    
    If Not IsArray(Arg) Then
        Let IsNumericArrayQ = False
    ElseIf NumberOfDimensions(Arg) <> 1 Then
        Let IsNumericArrayQ = False
    Else
        For Each var In Arg
            If Not IsNumeric(var) Then
                Let IsNumericArrayQ = False
                Exit For
            End If
        Next
    End If
End Function

' This function returns TRUE if the given directory exists.  Otherwise, it returns FALSE
Public Function DirectoryExistsQ(TheDirPath As String)
    Let DirectoryExistsQ = False
    If Not Dir(TheDirPath, vbDirectory) = vbNullString Then DirectoryExistsQ = True
End Function

' Boolean function that returns TRUE if the given table exists in the given worksheet.  Otherwise, it returns FALSE
Public Function ListObjectExistsQ(WorkSheetReference As Worksheet, ListObjectName As String) As Boolean
    Dim AName As String
    
    Let ListObjectExistsQ = True
    
    On Error GoTo ErrorHandler
    
    Let AName = WorkSheetReference.ListObjects(ListObjectName).Name
    
    Exit Function
    
ErrorHandler:
    Let ListObjectExistsQ = False
End Function

' Returns True is the argument is 2D matrix with atomic entries.  No attempt is made to check if entries
' are numeric.
Public Function MatrixQ(Arg As Variant) As Boolean
    Dim var As Variant
    
    If NumberOfDimensions(Arg) <> 2 Then
        Let MatrixQ = False
        Exit Function
    End If

    For Each var In Arg
        If IsArray(var) Then
            Let MatrixQ = False
            Exit Function
        End If
    Next
    
    Let MatrixQ = True
End Function

' Returns True is the argument is a row or column vector as defined by ColumVectorQ and RowVectorQ
Function VectorQ(Arg As Variant) As Boolean
    Let VectorQ = RowVectorQ(Arg) Or ColumnVectorQ(Arg)
End Function

' Returns True is the argument is a column vector (e.g. n x 1 2D array with atomic entries exclusively)
Public Function ColumnVectorQ(Arg As Variant) As Boolean
    Dim nd As Long
    Dim i As Long
    
    If EmptyArrayQ(Arg) Then
        Let ColumnVectorQ = True
        Exit Function
    Else
        Let ColumnVectorQ = False
        Let nd = NumberOfDimensions(Arg)
    End If
    
    If nd <> 2 Then Exit Function
    
    If LBound(Arg, 2) <> UBound(Arg, 2) Then Exit Function
    
    For i = LBound(Arg, 1) To UBound(Arg, 1)
        If IsArray(Arg(i, 1)) Then Exit Function
    Next i

    Let ColumnVectorQ = True
End Function

' Returns True if the argument is a row vector (e.g. 1D non of whose entries is an array)
Public Function RowVectorQ(Arg As Variant) As Boolean
    Dim nd As Long
    Dim i As Long
    
    Let RowVectorQ = False
    
    If NumberOfDimensions(Arg) <> 1 Then Exit Function
    
    For i = LBound(Arg) To UBound(Arg)
        If IsArray(Arg(i)) Then Exit Function
    Next i
    
    Let RowVectorQ = True
End Function

' Returns True if the argument is a 1D array row vector all of whose entries are either numeric, strings,
' dates, or Boolean. Empty qualifies as numeric (e.g. IsNumeric(Empty) = True)
'***HERE TEST
Public Function AtomicArrayQ(Arg As Variant) As Boolean
    Dim nd As Long
    Dim i As Long
    
    Let AtomicArrayQ = False
    
    If NumberOfDimensions(Arg) <> 1 Then Exit Function
    
    For i = LBound(Arg) To UBound(Arg)
        If IsArray(Arg(i)) Then
            Exit Function
        End If
        
        If Not (TypeName(Arg(i)) = TypeName(1) Or _
                TypeName(Arg(i)) = TypeName(1.1) Or _
                TypeName(Arg(i)) = TypeName("a") Or _
                TypeName(Arg(i)) = TypeName(True) Or _
                TypeName(Arg(i)) = TypeName(#1/10/2015#)) Then
            Exit Function
        End If
    Next i
    
    Let AtomicArrayQ = True
End Function


' Returns true if the array could be interpreseted as a 1D array of atomic elements.
' This means that this function returns True for each of the following examples:
' 1. Array(1,2,3)
' 2. Array(Array(1,2,3))
' 3. Array(Array(Array(1,2,3)))),
' 4. [{1,2,3}], [{Array(1,2,3)}], etc. will evaluate to True.
Public Function RowArrayQ(a As Variant) As Boolean
    Dim nd As Integer
    Dim i As Long
    
    Let RowArrayQ = False
    Let nd = NumberOfDimensions(a)
    
    ' If a has more than two or fewer than one dimension, then exit with False.
    If nd > 2 Or nd < 1 Then
        Exit Function
    End If
    
    ' Process arg is it has one dimensions
    If nd = 1 Then
        ' If this is a 1-element, 1D array
        If LBound(a) = UBound(a) Then
            If IsArray(a) Then Let RowArrayQ = RowArrayQ(a(LBound(a)))
            Exit Function
        ' If this is a multi-element 1D array
        Else
            For i = LBound(a) To UBound(a)
                If IsArray(a(i)) Then Exit Function
            Next i
        End If
        
        Let RowArrayQ = True
        Exit Function
    End If
    
    ' If we get here the array is two dimensional
    ' This is the 2D, single-element case
    If UBound(a, 1) = LBound(a, 1) And UBound(a, 2) = LBound(a, 2) Then
        Let RowArrayQ = RowArrayQ(a(LBound(a, 1), LBound(a, 2)))
        Exit Function
    ' This is the case when we have a matrix that cannot be interpreted as a row
    ElseIf UBound(a, 1) > LBound(a, 1) And UBound(a, 2) > LBound(a, 2) Then
        Exit Function
    ' This is the case when there is just one row
    ElseIf (UBound(a, 1) = LBound(a, 1)) And (UBound(a, 2) > LBound(a, 2)) Then
        For i = LBound(a, 2) To UBound(a, 2)
            If IsArray(a(LBound(a, 1), i)) Then
                Exit Function
            End If
        Next i
    ' This is the case when there is just one column of length > 1
    Else
        Exit Function
    End If
    
    Let RowArrayQ = True
End Function

' Returns true if the array is a row array (must be a 2D array)
' This returns TRUE only for a 2D array with one column
Public Function ColumnArrayQ(a As Variant) As Boolean
    Dim nd As Integer
    Dim i As Long
    
    If EmptyArrayQ(a) Then
        Let ColumnArrayQ = True
        Exit Function
    End If
    
    Let ColumnArrayQ = False
    Let nd = NumberOfDimensions(a)
    
    If nd < 1 Or nd > 2 Then
        Exit Function
    End If
    
    If nd = 1 Then
        If LBound(a) = UBound(a) Then
            If Not IsArray(a(LBound(a))) Then
                Let ColumnArrayQ = True
            Else
                Let ColumnArrayQ = ColumnArrayQ(a(LBound(a)))
            End If
        End If
        
        Exit Function
    End If
    
    If (UBound(a, 1) > LBound(a, 1)) And (UBound(a, 2) > LBound(a, 2)) Then
        Exit Function
    ElseIf (UBound(a, 1) > LBound(a, 1)) And (UBound(a, 2) = LBound(a, 2)) Then
        For i = LBound(a, 1) To UBound(a, 1)
            If IsArray(a(i, 1)) Then Exit Function
        Next i
    
        Let ColumnArrayQ = True
    ElseIf (UBound(a, 1) = LBound(a, 1)) And (UBound(a, 2) = LBound(a, 2)) Then
        If IsArray(a(LBound(a, 1), LBound(a, 2))) Then
            Let ColumnArrayQ = ColumnArrayQ(a(LBound(a, 1), LBound(a, 2)))
        Else
            Let ColumnArrayQ = True
        End If
    End If
End Function

' This function determines if the given value is in the given array (1D or 2D)
Public Function MemberQ(TheArray As Variant, TheValue As Variant) As Boolean
    Dim i As Long
    Dim TheResultFlag As Boolean
    
    ' Assume result is False and change TheValue is in any one column of TheArray
    Let TheResultFlag = False
    
    If NumberOfDimensions(TheArray) <= 1 Then
        Let MemberQ = IsValueIn1DArray(TheArray, TheValue)
        
        Exit Function
    End If
    
    ' Go through all the columns, checking if the result is true for one of the columns
    For i = 1 To UBound(TheArray, 2)
        Let TheResultFlag = TheResultFlag Or IsValueIn1DArray(ConvertTo1DArray(GetColumn(TheArray, i)), TheValue)
    Next i
    
    ' Set the value to return
    Let MemberQ = TheResultFlag
End Function

' This function determines if the given value is missig from the given array (1D or 2D).
' It is the complete opposite of MemberQ
Public Function FreeQ(TheArray As Variant, TheValue As Variant) As Boolean
    Let FreeQ = Not MemberQ(TheArray, TheValue)
End Function

' This is private, helper function used by MemberQ above. This function returns true only
' if TheArray has dimension 0 and is equal to TheValue or if TheArray has dimension 1 and
' TheValue is in TheArray.
Private Function IsValueIn1DArray(TheArray As Variant, TheValue As Variant) As Boolean
    If EmptyArrayQ(TheArray) Then
        Let IsValueIn1DArray = False
    
        Exit Function
    End If
    
    If NumberOfDimensions(TheArray) > 1 Then
        Let IsValueIn1DArray = False
    ElseIf NumberOfDimensions(TheArray) = 0 Then
        Let IsValueIn1DArray = (TheValue = TheArray)
    ElseIf IsNumeric(Application.Match(TheValue, TheArray, 0)) Then
        Let IsValueIn1DArray = True
    Else
        Let IsValueIn1DArray = False
    End If
End Function

' This function returns true if the array passed is 1D and empty (e.g. array())
' An atomic parameter evaluates to False
' This function correctly handles empty multi-dimensional arrays
Public Function EmptyArrayQ(AnArray As Variant) As Boolean
    ' Return True if the argument is not an array
    If Not IsArray(AnArray) Then
        Let EmptyArrayQ = False
        Exit Function
    End If

    ' Return True if we have either representatio of an empty array.
    ' One representation starts with index 0 and the other with 1.
    If LBound(AnArray) = 0 And UBound(AnArray) = -1 Then
        Let EmptyArrayQ = True
    ElseIf LBound(AnArray) = 1 And UBound(AnArray) = 0 Then
        Let EmptyArrayQ = True
    Else
        Let EmptyArrayQ = False
    End If
End Function

' This predicate returns true if the given workbook has a worksheet with name WorksheetName.
' Otherwise, it returns false
Public Function WorksheetExistsQ(AWorkBook As Workbook, WorksheetName As String) As Boolean
    Let WorksheetExistsQ = False
    
    On Error Resume Next
    
    Let WorksheetExistsQ = AWorkBook.Worksheets(WorksheetName).Name <> ""
    Exit Function
    
    On Error GoTo 0
End Function

' This predicate returns true if the given workbook has a worksheet with name WorksheetName.
' Otherwise, it returns false
Public Function SheetExistsQ(AWorkBook As Workbook, SheetName As String) As Boolean
    ' returns TRUE if the sheet exists in the active workbook
    Let SheetExistsQ = False
    
    On Error GoTo NoSuchSheet
    If Len(AWorkBook.Sheets(SheetName).Name) > 0 Then
        Let SheetExistsQ = True
        Exit Function
    End If

NoSuchSheet:

End Function

' Returns true if the argument's type is string
' Otherwise, it returns false
Public Function StringQ(AnArg As Variant) As Boolean
    Let StringQ = TypeName(AnArg) = TypeName("a")
End Function

' Returns true if the argument is a 1D array with all arguments
' satisfying StringQ() in this module
Public Function String1DArrayQ(AnArray As Variant) As Boolean
    Dim var As Variant

    Let String1DArrayQ = True

    If Not IsArray(AnArray) Then
        Let String1DArrayQ = False
        Exit Function
    End If
    
    For Each var In AnArray
        If Not StringQ(var) Then
            Let String1DArrayQ = False
            Exit Function
        End If
    Next
End Function
