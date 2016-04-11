Attribute VB_Name = "FunctionalProgramming"
Option Explicit
Option Base 1

' DESCRIPTION
' Applies a sequence of functions to an element, returning an array holding the result of applying
' each of the functions to the element. This function returns Null if the function array fails
' Predicates.StringArrayQ.  This function returns an empty array if the array of function names is empty.
' This function returns Null if AnElement is not atomic.
'
' PARAMETERS
' 1. AFunctionNameArray - An array of function names
' 2. CallingWorkbook - A reference to the workbook where the functions in AFunctionNameArray are located
' 3. AnElement - Any atomic value to which each of the functions will be applied
' 4. ParameterCheckQ (optional) - If this explicitly set to False, no parameter consistency checks are
'    perform.
'
' RETURNED VALUE
' An array with the results of applying a sequence of the functions to an element.
Public Function Through(AFunctionNameArray As Variant, _
                        CallingWorkbook As Workbook, _
                        AnElement As Variant, _
                        Optional ParameterCheckQ As Boolean = True) As Variant
    Dim ReturnArray() As Variant
    Dim c As Long
    
    ' Set default return value
    Let Through = Null
    
    ' Check parameters for consistency only if ParameterCheckQ is True
    If ParameterCheckQ Then
        ' Exit with Null if AFunctionNameArray is undimensioned or not 1D
        If Not StringArrayQ(AFunctionNameArray) Then Exit Function
        
        ' Exit the empty array if AFunctionNameArray satisfies EmptyArrayQ
        If EmptyArrayQ(AFunctionNameArray) Then
            Let Through = EmptyArray()
            
            Exit Function
        End If
        
        ' Exit with Null if AnElement is not atomic
        If Not (AtomicQ(AnElement) Or AtomicArrayQ(AnElement)) Then Exit Function
    End If
    
    ' Pre-allocate array to hold results
    ReDim ReturnArray(1 To Length(AFunctionNameArray))
    
    ' Compute values from applying each function to AnElement
    For c = 1 To Length(AFunctionNameArray)
        Let ReturnArray(c) = Run("'" & CallingWorkbook.Name & "'!" & AFunctionNameArray(c + LBound(AFunctionNameArray) - 1), _
                                 AnElement)
    Next
    
    ' Return results array
    Let Through = ReturnArray
End Function

' DESCRIPTION
' Applies a function to an array of atomic elements, returning an array with the results.
' This function returns Null if the array of elements is not atomic. This function returns an
' empty array if the array of elements is empty.  This function interprets a 2D array as a
' 1D arrays of its rows.
'
' PARAMETERS
' 1. AFunctionName - The name of the function to apply to each of the elements in the array
' 2. CallingWorkbook - A reference to the workbook where the function is located
' 3. A1DArray - An array of atomic elements
'
' RETURNED VALUE
' An array with the results of applying a sequence of the functions to an element.
Public Function ArrayMap(AFunctionName As String, _
                         CallingWorkbook As Workbook, _
                         A1DArray As Variant) As Variant
    Dim ReturnArray As Variant
    Dim c As Long
    
    ' Set default return value
    Let ArrayMap = Null
    
    ' Exit with Null if A1DArray is not dimensioned
    If Not DimensionedQ(A1DArray) Then Exit Function
    
    ' Check parameters for consistency
    ' Exit with the empty array if A1DArray is empty
    If EmptyArrayQ(A1DArray) Then
        Let ArrayMap = EmptyArray()
        
        Exit Function
    End If

    ' Pre-allocate results array
    ReDim ReturnArray(1 To Length(A1DArray))
    
    ' Compute the values from mapping the function over the array
    For c = 1 To Length(A1DArray)
        Let ReturnArray(c) = Run("'" & CallingWorkbook.Name & "'!" & AFunctionName, _
                                 Part(A1DArray, c))
    Next c

    ' Return the array holding the mapped results
    Let ArrayMap = ReturnArray
End Function

' DESCRIPTION
' Returns an array with the elements in the given array returning True when applying the
' given function. This function returns Null if the array of elements is not atomic. It
' returns an empty array if the array of elements is empty.
'
' PARAMETERS
' 1. AnArray - An array of atomic elements
' 2. CallingWorkbook - A reference to the workbook where the function is located
' 3. AFunctionName - The name of the function to apply to each of the elements in the array
' 4. ParameterCheckQ (optional) - If this explicitly set to False, no parameter consistency checks are
'    perform.
'
' RETURNED VALUE
' An array with the results of applying a sequence of the functions to an element.
Public Function ArraySelect(AnArray As Variant, _
                            CallingWorkbook, _
                            AFunctionName As String) As Variant
    Dim ReturnArray As Variant
    Dim i As Long
    Dim c As Long
    
    ' Set default return value
    Let ArraySelect = Null
    
    ' Check parameters for consistency
    If Not DimensionedQ(AnArray) Then Exit Function
    
    If NumberOfDimensions(AnArray) < 1 Or NumberOfDimensions(AnArray) > 2 Then Exit Function
    
    ' Exit with the empty array if AnArray is empty
    If EmptyArrayQ(AnArray) Then
        Let ArrayMap = EmptyArray
        
        Exit Function
    End If
    
    ' Pre-allocated array to return at most all of the elements in the array
    ReDim ReturnArray(1 To Length(AnArray))
    
    ' Cycle through the array, adding to the return array those elements yielding True
    For i = 1 To Length(AnArray)
        If Run("'" & CallingWorkbook.Name & "'!" & AFunctionName, Part(AnArray, i)) Then
            Let ReturnArray(i) = var
            Let c = c + 1
        End If
    Next
    
    If c = 0 Then
        Let ArraySelect = EmptyArray()
    Else
        ' Throw away the unused slots in ReturnArray
        ReDim Preserve ReturnArray(1 To c)
        Let ArraySelect = ReturnArray
    End If
End Function

' DESCRIPTION
' Returns the result of performing a Mathematica-like MapThread.  It returns an array with the same
' length as any of the array elements of parameter ArrayOfEqualLength1DArrays after the sequential
' application of the function with name AFunctionName to the arrays resulting from packing the ith
' element of each of the arrays in ArrayOfEqualLength1DArrays.
'
' If the parameters are compatible with expectations, the function returns Null
'
' Example: ArrayMapThread("StringJoin", ThisWorkbook, array(1,2,3), array(10,20,30)) returns
'          ("110", "220", "330")
'
' PARAMETERS
' 1. AFunctionName - Name of the function to apply
' 2. CallingWorkbook - A reference to the workbook where the function is located
' 3. ArrayOfEqualLength1DArrays - A sequence of equal-length, atomic arrays
'
' RETURNED VALUE
' An array with the results of applying the given function to the threading of the parameter arrays
Public Function ArrayMapThread(AFunctionName As String, _
                               CallingWorkbook As Workbook, _
                               ParamArray ArrayOfEqualLength1DArrays() As Variant) As Variant
    Dim var As Variant
    Dim N As Long
    Dim r As Long
    Dim c As Long
    Dim ArrayNumber As Long
    Dim ElementNumber As Long
    Dim ParamsArray As Variant
    Dim CallArray As Variant
    Dim ReturnArray As Variant
    
    ' Set default return value
    Let ArrayMapThread = Null
    
    ' Make a copy of array of parameters so we may apply other functions to it
    Let ParamsArray = CopyParamArray(ArrayOfEqualLength1DArrays)
    
    ' Exit with Null if ParamsArray is not dimensioned
    If Not DimensionedQ(ParamsArray) Then Exit Function
    
    ' Exit with Null if ParamsArray is empty
    If EmptyArrayQ(ParamsArray) Then
        Let ArrayMapThread = EmptyArray
        Exit Function
    End If
    
    ' Get the length of the first parameter array
    Let N = Length(First(ParamsArray))
    
    ' Exit with Null if the arrays don't all have the same length or are atomic
    For Each var In ParamsArray
        If Length(var) <> N Then Exit Function
    Next
    
    ' Pre-allocate array to hold results and each function call's array
    ReDim ReturnArray(1 To N)
    ReDim CallArray(1 To Length(ParamsArray))

    ' Loop over the array elements to compute results to return
    For r = 1 To N
        ' Assemble array of value for this function call
        For c = 1 To Length(ParamsArray)
            Let CallArray(c) = Part(Part(ParamsArray, c), r)
        Next
        
        Let ReturnArray(r) = Run("'" & CallingWorkbook.Name & "'!" & AFunctionName, CallArray)
    Next
    
    ' Return the array used
    Let ArrayMapThread = ReturnArray
End Function

' DESCRIPTION
' Returns the result of performing a Mathematica-like MapThread.  It returns an array with the same
' length as any of the array elements of parameter ArrayOfEqualLength1DArrays after the sequential
' application of the function with name AFunctionName to the arrays resulting from packing the ith
' element of each of the arrays in ArrayOfEqualLength1DArrays.
'
' If the parameters are compatible with expectations, the function returns Null
'
' Example: ArrayMapThread("StringJoin", ThisWorkbook, array(1,2,3), array(10,20,30)) returns
'          ("110", "220", "330")
'
' PARAMETERS
' 1. AFunctionName - Name of the function to apply
' 2. CallingWorkbook - A reference to the workbook where the function is located
' 3. ArrayOfEqualLength1DArrays - A sequence of equal-length, atomic arrays
'
' RETURNED VALUE
' An array with the results of applying the given function to the threading of the parameter arrays
Public Function Nest(AFunctionName As String, _
                     CallingWorkbook As Workbook, _
                     N As Long) As Variant
'***HERE
End Function

' DESCRIPTION
' Returns the result of performing a Mathematica-like MapThread.  It returns an array with the same
' length as any of the array elements of parameter ArrayOfEqualLength1DArrays after the sequential
' application of the function with name AFunctionName to the arrays resulting from packing the ith
' element of each of the arrays in ArrayOfEqualLength1DArrays.
'
' If the parameters are compatible with expectations, the function returns Null
'
' Example: ArrayMapThread("StringJoin", ThisWorkbook, array(1,2,3), array(10,20,30)) returns
'          ("110", "220", "330")
'
' PARAMETERS
' 1. AFunctionName - Name of the function to apply
' 2. CallingWorkbook - A reference to the workbook where the function is located
' 3. ArrayOfEqualLength1DArrays - A sequence of equal-length, atomic arrays
'
' RETURNED VALUE
' An array with the results of applying the given function to the threading of the parameter arrays
Public Function NestList(AFunctionName As String, _
                         CallingWorkbook As Workbook, _
                         N As Long) As Variant
'***HERE
End Function

