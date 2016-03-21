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
    Dim Result1DArray() As Variant
    Dim i As Long

    ' Exit with Null if AFunctionNameArray is undimensioned or not 1D
    If NumberOfDimensions(AFunctionNameArray) <> 1 Then
        Let Through = Null
        Exit Function
    End If
    
    ' Exit the empty array if AFunctionNameArray satisfies EmptyArrayQ
    If EmptyArrayQ(AFunctionNameArray) Then
        Let Through = EmptyArray()
        Exit Function
    End If
    
    ' Exit with Null if any of the elements of AFunctionNameArray fails StringQ
    For i = LBound(AFunctionNameArray, 1) To UBound(AFunctionNameArray, 1)
        If Not StringQ(AFunctionNameArray(i)) Then
            Let Through = Null
            Exit Function
        End If
    Next
    
    ' Pre-allocate array to hold results
    ReDim Result1DArray(LBound(AFunctionNameArray) To UBound(AFunctionNameArray))
    
    ' Compute values from applying each function to AnElement
    For i = LBound(AFunctionNameArray) To UBound(AFunctionNameArray)
        Let Result1DArray(i) = Run("'" & CallingWorkbook.Name & "'!" & AFunctionNameArray(i), AnElement)
    Next i
    
    ' Return results array
    Let Through = Result1DArray
End Function

' This function returns an array of the same length as A1DArray with the result of apply
' the function with name AFunctionName to each element of A1Array
'
' The function being mapped over the array must be an array of variants.  If it is not,
' one has to pass the optional TheDataType
Public Function ArrayMap(AFunctionName As String, CallingWorkbook As Workbook, A1DArray As Variant) As Variant
    Dim Result1DArray() As Variant
    Dim c As Long
    
    ' Exit with Null if A1DArray is not a dimensioned, 1D array
    If NumberOfDimensions(A1DArray) <> 1 Then
        Let ArrayMap = Null
        
        Exit Function
    End If

    ' Exit with the empty array if A1DArray is empty
    If EmptyArrayQ(A1DArray) Then
        Let ArrayMap = EmptyArray()
        
        Exit Function
    End If
    
    ' Pre-allocate results array
    ReDim Result1DArray(LBound(A1DArray, 1) To UBound(A1DArray, 1))
    
    ' Compute the values from mapping the function over the array
    For c = LBound(A1DArray) To UBound(A1DArray)
        Let Result1DArray(c) = Run("'" & CallingWorkbook.Name & "'!" & AFunctionName, A1DArray(c))
    Next c

    ' Return the array holding the mapped results
    Let ArrayMap = Result1DArray
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
    Dim Result1DArray() As Variant

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
    ReDim Result1DArray(LBound(First(ArrayOfEqualLength1DArrays)) To UBound(First(ArrayOfEqualLength1DArrays)))

    For r = LBound(First(ArrayOfEqualLength1DArrays)) To UBound(First(ArrayOfEqualLength1DArrays))
        For c = LBound(ArrayOfEqualLength1DArrays) To UBound(ArrayOfEqualLength1DArrays)
            Let ParamsArray(c) = ArrayOfEqualLength1DArrays(c)(r)
        Next c
        
        Let Result1DArray(r) = Run("'" & CallingWorkbook.Name & "'!" & AFunctionName, ParamsArray)
    Next r
    
    Let ArrayMapThread = Result1DArray
End Function

' This function returns the sub-array of A1DArray defined by those elements for which the funciton yields
' AFunctionName is the string name of a boolean function.  This function must be able to act of each element of
' A1DArray.
Public Function ArraySelect(A1DArray As Variant, TheWorkbook, AFunctionName As String) As Variant
    Dim Result1DArray As Dictionary
    Dim i As Long
        
    If NumberOfDimensions(A1DArray) <> 1 Then
        Let ArraySelect = EmptyArray()
        
        Exit Function
    End If
    
    Set Result1DArray = New Dictionary
    
    For i = LBound(A1DArray) To UBound(A1DArray)
        If Run("'" & TheWorkbook.Name & "'!" & AFunctionName, A1DArray(i)) Then
            Call Result1DArray.Add(Key:=i, Item:=A1DArray(i))
        End If
    Next i
    
    If Result1DArray.Count = 0 Then
        Let ArraySelect = EmptyArray()
    ElseIf Result1DArray.Count = 1 Then
        Let ArraySelect = Array(Result1DArray.Item(Key:=LBound(A1DArray)))
    Else
        Let ArraySelect = Result1DArray.Items
    End If
End Function

