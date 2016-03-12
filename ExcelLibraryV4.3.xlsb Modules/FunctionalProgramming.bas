Attribute VB_Name = "FunctionalProgramming"
Option Explicit
Option Base 1

' This function returns an array of the same length as A1DArray with the result of apply
' the function with name AFunctionName to each element of A1Array
Public Function ArrayMap(AFunctionName As String, A1DArray As Variant) As Variant
    Dim TheResults() As Variant
    Dim i As Long
    
    If EmptyArrayQ(A1DArray) Then
        Let ArrayMap = Array()
        
        Exit Function
    End If
    
    ReDim TheResults(LBound(A1DArray, 1) To UBound(A1DArray, 1))
    
    For i = LBound(A1DArray) To UBound(A1DArray)
        Let TheResults(i) = Run(AFunctionName, GetRow(A1DArray, i))
    Next i

    Let ArrayMap = TheResults
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
Public Function ArrayMapThread(AFunctionName As String, ArrayOfEqualLength1DArrays As Variant) As Variant
    Dim var As Variant
    Dim r As Long

    ' Input consistency checks
    
    ' Exit with Nothing if ArrayOfEqualLength1DArrays
    If Not IsArray(ArrayOfEqualLength1DArrays) Then
        Set ArrayMapThread = Nothing
        Exit Function
    End If
    
    ' Exit with Nothing if any one of the elts in ArrayOfEqualLength1DArrays is not an array
    For Each var In ArrayOfEqualLength1DArrays
        If Not IsArray(var) Then
            Set ArrayMapThread = Nothing
            Exit Function
        End If
    Next
    
    ' Exit with Nothing if elts of ArrayOfEqualLength1DArrays are not of equal length
    For Each var In ArrayOfEqualLength1DArrays
        If GetArrayLength(var) <> GetArrayLength(First(ArrayOfEqualLength1DArrays)) Then
            Set ArrayMapThread = Nothing
            Exit Function
        End If
    Next
    
    ' Exit if ArrayOfEqualLength1DArrays is an empty array
    If EmptyArrayQ(ArrayOfEqualLength1DArrays) Then
        Set ArrayMapThread = Nothing
        Exit Function
    End If
    
    ' Exit with Nothing if any of the elements of ArrayOfEqualLength1DArrays is an empty array
    If EmptyArrayQ(First(ArrayOfEqualLength1DArrays)) Then
        Set ArrayMapThread = Nothing
        Exit Function
    End If
    
    ' If the code gets here, inputs are consistent
    Let ArrayMapThread = ArrayMap(AFunctionName, TransposeMatrix(Pack2DArray(ArrayOfEqualLength1DArrays)))
End Function

' This function returns the sub-array of A1DArray defined by those elements for which the funciton yields
' AFunctionName is the string name of a boolean function.  This function must be able to act of each element of
' A1DArray.
Public Function ArraySelect(A1DArray As Variant, AFunctionName As String) As Variant
    Dim TheResults As Dictionary
    Dim i As Long
        
    If NumberOfDimensions(A1DArray) <> 1 Then
        Let ArraySelect = Array()
        
        Exit Function
    End If
    
    Set TheResults = New Dictionary
    
    For i = LBound(A1DArray) To UBound(A1DArray)
        If Run(AFunctionName, A1DArray(i)) Then
            Call TheResults.Add(Key:=i, Item:=A1DArray(i))
        End If
    Next i
    
    If TheResults.Count = 0 Then
        Let ArraySelect = Array()
    ElseIf TheResults.Count = 1 Then
        Let ArraySelect = Array(TheResults.Item(Key:=LBound(A1DArray)))
    Else
        Let ArraySelect = TheResults.Items
    End If
End Function

