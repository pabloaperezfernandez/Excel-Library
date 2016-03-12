Attribute VB_Name = "FunctionalProgramming"
Option Explicit
Option Base 1

' Alias for function MapFunctionOverArray()
Public Function ArrayMap(AFunctionName As String, A1DArray As Variant) As Variant
    Let ArrayMap = MapFunctionOverArray(AFunctionName, A1DArray)
End Function

' This function returns an array of the same length as A1DArray with the result of apply
' the function with name AFunctionName to each element of A1Array
Public Function MapFunctionOverArray(AFunctionName As String, A1DArray As Variant) As Variant
    Dim TheResults() As Variant
    Dim i As Long
    
    If EmptyArrayQ(A1DArray) Then
        Let MapFunctionOverArray = Array()
        
        Exit Function
    End If
    
    ReDim TheResults(LBound(A1DArray) To UBound(A1DArray))
    
    For i = LBound(A1DArray) To UBound(A1DArray)
        Let TheResults(i) = Run(AFunctionName, A1DArray(i))
    Next i
    
    Let MapFunctionOverArray = TheResults
End Function
' Alias for function SelectFromArrayWithFunction()
Public Function ArraySelect(A1DArray As Variant, AFunctionName As String) As Variant
    Let ArraySelect = SelectFromArrayWithFunction(A1DArray, AFunctionName)
End Function

' This function returns the sub-array of A1DArray defined by those elements for which the funciton yields
' AFunctionName is the string name of a boolean function.  This function must be able to act of each element of
' A1DArray.
Public Function SelectFromArrayWithFunction(A1DArray As Variant, AFunctionName As String) As Variant
    Dim TheResults As Dictionary
    Dim i As Long
        
    If NumberOfDimensions(A1DArray) <> 1 Then
        Let SelectFromArrayWithFunction = Array()
        
        Exit Function
    End If
    
    Set TheResults = New Dictionary
    
    For i = LBound(A1DArray) To UBound(A1DArray)
        If Run(AFunctionName, A1DArray(i)) Then
            Call TheResults.Add(Key:=i, Item:=A1DArray(i))
        End If
    Next i
    
    If TheResults.Count = 0 Then
        Let SelectFromArrayWithFunction = Array()
    ElseIf TheResults.Count = 1 Then
        Let SelectFromArrayWithFunction = Array(TheResults.Item(Key:=LBound(A1DArray)))
    Else
        Let SelectFromArrayWithFunction = TheResults.Items
    End If
End Function

