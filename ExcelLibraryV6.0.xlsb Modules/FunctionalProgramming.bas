Attribute VB_Name = "FunctionalProgramming"
Option Explicit
Option Base 1

' DESCRIPTION
' Boolean function returning True if its argument is an array that has been dimensioned.
' Returns False otherwise.  In other words, it returns False if its arg is neither an
' an array nor dimensioned.
'
' PARAMETERS
' 1. arg - any value or object reference
'
' RETURNED VALUE
' True when arg is a dimensioned array. False otherwise.
Public Function Through(AFunctionNameArray As Variant, CallingWorkbook As Workbook, AnElement As Variant) As Variant
    Dim ResultArray() As Variant
    Dim i As Long

    ' Exit with Null if AFunctionNameArray is undimensioned
    If Not StringArrayQ(AFunctionNameArray) Then
        Let Through = Null
        Exit Function
    End If
    
    ' Exit the empty array if AFunctionNameArray satisfies EmptyArrayQ
    If EmptyArrayQ(AFunctionNameArray) Then
        Let Through = EmptyArray()
        Exit Function
    End If
    
    ' Exit with Null if AnElement fails AtomicQ
    If Not AtomicQ(AnElement) Then
        Let Through = Null
        Exit Function
    End If
    
    ReDim ResultArray(LBound(AFunctionNameArray) To UBound(AFunctionNameArray))
    For i = LBound(AFunctionNameArray) To UBound(AFunctionNameArray)
        Let ResultArray(i) = Run("'" & CallingWorkbook.Name & "'!" & AFunctionNameArray(i), AnElement)
    Next i
    
    Let Through = ResultArray
End Function

' This function returns an array of the same length as A1DArray with the result of apply
' the function with name AFunctionName to each element of A1Array
'
' The function being mapped over the array must be an array of variants.  If it is not,
' one has to pass the optional TheDataType
Public Function ArrayMap(AFunctionName As String, CallingWorkbook As Workbook, A1DArray As Variant, _
                         Optional TheDataType As DataType) As Variant
    Dim TheResults() As Variant
    Dim IntegerArray() As Integer
    Dim LongArray() As Integer
    Dim DoubleArray() As Double
    Dim StringArray() As String
    Dim BooleanArray() As Boolean
    Dim WorksheetArray() As Worksheet
    Dim WorkbookArray() As Workbook
    Dim c As Long

    ' Parameter consistency checks
    If Not IsArray(A1DArray) Then
        Let ArrayMap = EmptyArray()
        
        Exit Function
    End If
    
    If NumberOfDimensions(A1DArray) <> 1 Then
        Let ArrayMap = EmptyArray()
        
        Exit Function
    End If
    
    
    If Not DimensionedQ(A1DArray) Then
        Let ArrayMap = EmptyArray()
        
        Exit Function
    End If

    If EmptyArrayQ(A1DArray) Then
        Let ArrayMap = EmptyArray()
        
        Exit Function
    End If
    
    ' Pre-allocate results array
    ReDim TheResults(LBound(A1DArray, 1) To UBound(A1DArray, 1))
    
    ' Exit with Null if TheDataType is not one of the supported types
    If Not IsMissing(TheDataType) Then
        If FreeQ(Array(StringType, IntegerType, LongType, DoubleType, _
                       BooleanType, WorksheetType, WorkbookType), _
                 TheDataType) Then
            Let ArrayMap = Null
            Exit Function
        End If
        
        For c = LBound(A1DArray) To UBound(A1DArray)
            Let TheResults(c) = Run("'" & CallingWorkbook.Name & "'!" & AFunctionName, A1DArray(c))
        Next c
    
        Let ArrayMap = TheResults
        
        Exit Function
    End If

    Select Case TheDataType
        Case IntegerType
            ReDim IntegerArray(LBound(A1DArray) To UBound(A1DArray))
            
            For c = LBound(A1DArray) To UBound(A1DArray)
                Let IntegerArray(c) = CInt(A1DArray(c))
            Next
            
            For c = LBound(A1DArray) To UBound(A1DArray)
                Let TheResults(c) = Run("'" & CallingWorkbook.Name & "'!" & AFunctionName, IntegerArray(c))
            Next c
        
            Let ArrayMap = TheResults
        Case LongType
            ReDim LongArray(LBound(A1DArray) To UBound(A1DArray))
            
            For c = LBound(A1DArray) To UBound(A1DArray)
                Let LongArray(c) = CLng(A1DArray(c))
            Next

            For c = LBound(A1DArray) To UBound(A1DArray)
                Let TheResults(c) = Run("'" & CallingWorkbook.Name & "'!" & AFunctionName, LongArray(c))
            Next c
        
            Let ArrayMap = TheResults
        Case DoubleType
            ReDim DoubleArray(LBound(A1DArray) To UBound(A1DArray))
            
            For c = LBound(A1DArray) To UBound(A1DArray)
                Let DoubleArray(c) = CDbl(A1DArray(c))
            Next

            For c = LBound(A1DArray) To UBound(A1DArray)
                Let TheResults(c) = Run("'" & CallingWorkbook.Name & "'!" & AFunctionName, DoubleArray(c))
            Next c
        
            Let ArrayMap = TheResults
        Case StringType
            ReDim StringArray(LBound(A1DArray) To UBound(A1DArray))
            
            For c = LBound(A1DArray) To UBound(A1DArray)
                Let StringArray(c) = CStr(A1DArray(c))
            Next

            For c = LBound(A1DArray) To UBound(A1DArray)
                Let TheResults(c) = Run("'" & CallingWorkbook.Name & "'!" & AFunctionName, StringArray(c))
            Next c
        
            Let ArrayMap = TheResults
        Case BooleanType
            ReDim BooleanArray(LBound(A1DArray) To UBound(A1DArray))
            
            For c = LBound(A1DArray) To UBound(A1DArray)
                Let BooleanArray(c) = CBool(A1DArray(c))
            Next

            For c = LBound(A1DArray) To UBound(A1DArray)
                Let TheResults(c) = Run("'" & CallingWorkbook.Name & "'!" & AFunctionName, BooleanArray(c))
            Next c
        
            Let ArrayMap = TheResults
        Case WorksheetType
            ReDim WorksheetArray(LBound(A1DArray) To UBound(A1DArray))
            
            For c = LBound(A1DArray) To UBound(A1DArray)
                Set WorksheetArray(c) = A1DArray(c)
            Next

            For c = LBound(A1DArray) To UBound(A1DArray)
                Let TheResults(c) = Run("'" & CallingWorkbook.Name & "'!" & AFunctionName, WorksheetArray(c))
            Next c
        
            Let ArrayMap = TheResults
        Case WorkbookType
            ReDim WorkbookArray(LBound(A1DArray) To UBound(A1DArray))
            
            For c = LBound(A1DArray) To UBound(A1DArray)
                Set WorkbookArray(c) = A1DArray(c)
            Next

            For c = LBound(A1DArray) To UBound(A1DArray)
                Let TheResults(c) = Run("'" & CallingWorkbook.Name & "'!" & AFunctionName, WorkbookArray(c))
            Next c
        
            Let ArrayMap = TheResults
    End Select
End Function

' Returns the result of performing a Mathematica-like MapThread.  It returns an array with the same
' length as any of the array elements of parameter ArrayOfEqualLength1DArrays after the sequential
' application of the function with name AFunctionName to the arrays resulting from packing ith element
' of each of the elements in ArrayOfEqualLength1DArrays.
'
' If the parameters are compatible with expectations, the function returns Nothing
'
' Example: ArrayMapThread("StringJoin", array(array(1,2,3), array(10,20,30))) returns
'          ("110", "220", "330")
Public Function ArrayMapThread(AFunctionName As String, _
                               CallingWorkbook As Workbook, _
                               ArrayOfEqualLength1DArrays As Variant) As Variant
    Dim var As Variant
    Dim r As Long
    Dim c As Long
    Dim ParamsArray() As Variant
    Dim ResultsArray() As Variant

    ' Input consistency checks
    
    ' Exit with Nothing if ArrayOfEqualLength1DArrays
    If Not IsArray(ArrayOfEqualLength1DArrays) Then
        Set ArrayMapThread = Null
        Exit Function
    End If
    
    ' Exit with Nothing if any one of the elts in ArrayOfEqualLength1DArrays is not an array
    For Each var In ArrayOfEqualLength1DArrays
        If Not IsArray(var) Then
            Set ArrayMapThread = Null
            Exit Function
        End If
    Next
    
    ' Exit with Nothing if elts of ArrayOfEqualLength1DArrays are not of equal length
    For Each var In ArrayOfEqualLength1DArrays
        If GetArrayLength(var) <> GetArrayLength(First(ArrayOfEqualLength1DArrays)) Then
            Set ArrayMapThread = Null
            Exit Function
        End If
    Next
    
    ' Exit if ArrayOfEqualLength1DArrays is an empty array
    If EmptyArrayQ(ArrayOfEqualLength1DArrays) Then
        Set ArrayMapThread = Null
        Exit Function
    End If
    
    ' Exit with Nothing if any of the elements of ArrayOfEqualLength1DArrays is an empty array
    If EmptyArrayQ(First(ArrayOfEqualLength1DArrays)) Then
        Set ArrayMapThread = Null
        Exit Function
    End If
    
    ' If the code gets here, inputs are consistent
    ReDim ParamsArray(LBound(ArrayOfEqualLength1DArrays) To UBound(ArrayOfEqualLength1DArrays))
    ReDim ResultsArray(LBound(First(ArrayOfEqualLength1DArrays)) To UBound(First(ArrayOfEqualLength1DArrays)))

    For r = LBound(First(ArrayOfEqualLength1DArrays)) To UBound(First(ArrayOfEqualLength1DArrays))
        For c = LBound(ArrayOfEqualLength1DArrays) To UBound(ArrayOfEqualLength1DArrays)
            Let ParamsArray(c) = ArrayOfEqualLength1DArrays(c)(r)
        Next c
        
        Let ResultsArray(r) = Run("'" & CallingWorkbook.Name & "'!" & AFunctionName, ParamsArray)
    Next r
    
    Let ArrayMapThread = ResultsArray
End Function

' This function returns the sub-array of A1DArray defined by those elements for which the funciton yields
' AFunctionName is the string name of a boolean function.  This function must be able to act of each element of
' A1DArray.
Public Function ArraySelect(A1DArray As Variant, TheWorkbook, AFunctionName As String) As Variant
    Dim TheResults As Dictionary
    Dim i As Long
        
    If NumberOfDimensions(A1DArray) <> 1 Then
        Let ArraySelect = EmptyArray()
        
        Exit Function
    End If
    
    Set TheResults = New Dictionary
    
    For i = LBound(A1DArray) To UBound(A1DArray)
        If Run("'" & TheWorkbook.Name & "'!" & AFunctionName, A1DArray(i)) Then
            Call TheResults.Add(Key:=i, Item:=A1DArray(i))
        End If
    Next i
    
    If TheResults.Count = 0 Then
        Let ArraySelect = EmptyArray()
    ElseIf TheResults.Count = 1 Then
        Let ArraySelect = Array(TheResults.Item(Key:=LBound(A1DArray)))
    Else
        Let ArraySelect = TheResults.Items
    End If
End Function

