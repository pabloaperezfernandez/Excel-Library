Attribute VB_Name = "FunctionalProgramming"
Option Explicit
Option Base 1

' This function returns an array of the same length as A1DArray with the result of apply
' the function with name AFunctionName to each element of A1Array
'
' The function being mapped over the array must be an array of variants.  If it is not,
' one has to pass the optional TheDataType
'***HERE
Public Function ArrayMap(AFunctionName As String, TheWorkbook As Workbook, A1DArray As Variant, _
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


    If Not IsArray(A1DArray) Then
        Let ArrayMap = Array()
        
        Exit Function
    End If
    
    If Not DimensionedQ(A1DArray) Then
        Let ArrayMap = Array()
        
        Exit Function
    End If

    If EmptyArrayQ(A1DArray) Then
        Let ArrayMap = Array()
        
        Exit Function
    End If
    
    
    ' Exit with Null if TheDataType is not one of the supported types
    If FreeQ(Array(StringType, IntegerType, LongType, DoubleType, _
                   BooleanType, WorksheetType, WorkbookType), _
             TheDataType) Then
        Let Cast = Null
        Exit Function
    End If

    ReDim TheResults(LBound(A1DArray, 1) To UBound(A1DArray, 1))
    Select Case TheDataType
        Case IntegerType
            ReDim IntegerArray(LBound(arg) To UBound(arg))
            
            For c = LBound(arg) To UBound(arg)
                Let IntegerArray(c) = CInt(arg(c))
            Next
            
            Let Cast = IntegerArray
        Case LongType
            ReDim LongArray(LBound(arg) To UBound(arg))
            
            For c = LBound(arg) To UBound(arg)
                Let IntegerArray(c) = CLng(arg(c))
            Next
            
            Let Cast = LongArray
        Case DoubleType
            ReDim DoubleArray(LBound(arg) To UBound(arg))
            
            For c = LBound(arg) To UBound(arg)
                Let DoubleArray(c) = CDbl(arg(c))
            Next
            
            Let Cast = DoubleArray
        Case StringType
            ReDim StringArray(LBound(arg) To UBound(arg))
            
            For c = LBound(arg) To UBound(arg)
                Let StringArray(c) = CStr(arg(c))
            Next
            
            Let Cast = StringArray
        Case BooleanType
            ReDim BooleanArray(LBound(arg) To UBound(arg))
            
            For c = LBound(arg) To UBound(arg)
                Let BooleanArray(c) = CBool(arg(c))
            Next
            
            Let Cast = BooleanArray
        Case WorksheetType
            ReDim WorksheetArray(LBound(arg) To UBound(arg))
            
            For c = LBound(arg) To UBound(arg)
                Let WorksheetArray(c) = arg(c)
            Next
            
            Let Cast = WorksheetArray
        Case WorkbookType
            ReDim WorkbookArray(LBound(arg) To UBound(arg))
            
            For c = LBound(arg) To UBound(arg)
                Let WorkbookArray(c) = arg(c)
            Next
            
            Let Cast = WorkbookArray
    End Select
    
    For c = LBound(A1DArray) To UBound(A1DArray)
        Let TheResults(i) = Run("'" & TheWorkbook.Name & "'!" & AFunctionName, GetRow(A1DArray, c))
    Next c

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
Public Function ArrayMapThread(AFunctionName As String, TheWorkbook As Workbook, ArrayOfEqualLength1DArrays As Variant) As Variant
    Dim var As Variant
    Dim r As Long

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
    Let ArrayMapThread = ArrayMap(AFunctionName, TheWorkbook, Pack2DArray(ArrayOfEqualLength1DArrays, True))
End Function

' This function returns the sub-array of A1DArray defined by those elements for which the funciton yields
' AFunctionName is the string name of a boolean function.  This function must be able to act of each element of
' A1DArray.
Public Function ArraySelect(A1DArray As Variant, TheWorkbook, AFunctionName As String) As Variant
    Dim TheResults As Dictionary
    Dim i As Long
        
    If NumberOfDimensions(A1DArray) <> 1 Then
        Let ArraySelect = Array()
        
        Exit Function
    End If
    
    Set TheResults = New Dictionary
    
    For i = LBound(A1DArray) To UBound(A1DArray)
        If Run("'" & TheWorkbook.Name & "'!" & AFunctionName, A1DArray(i)) Then
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

