Attribute VB_Name = "FunctionalProgramming"
Option Explicit
Option Base 1

' DESCRIPTION
' This function is the equivalent of Mathematica's Scan.  It applies the function or
' sub with name AFunctionName to each element of A1DArray without storing the returned
' result.  Usage examples include:
'
' Call ForEach("MySub", Array(1, 2, 3, 4))
'
' Applies a function to an array of atomic elements, returning an array with the results.
' This function returns Null if the array of elements is not atomic. This function returns an
' empty array if the array of elements is empty.  This function interprets a 2D array as a
' 1D arrays of its rows.
'
' PARAMETERS
' 1. AFunctionName - The name of the function to apply to each of the elements in the array
' 2. A1DArray - An array of atomic elements
'
' RETURNED VALUE
' An array with the results of applying a sequence of the functions to an element.
Public Function Scan(aFunctionName As String, A1DArray As Variant) As Variant
    Dim var As Variant
    
    ' Set default return value
    Let Scan = Null
    
    ' Exit with Null if A1DArray is not dimensioned
    If Not DimensionedQ(A1DArray) Then Exit Function
    
    ' Check parameters for consistency
    ' Exit with the empty array if A1DArray is empty
    If EmptyArrayQ(A1DArray) Then
        Let Scan = True
        
        Exit Function
    End If
    
    ' Compute the values from mapping the function over the array
    For Each var In A1DArray
        ' Exit with Null if AFunctionName returs Null on current array element
        If IsNull(Run(aFunctionName, var)) Then Exit Function
    Next

    ' Return the array holding the mapped results
    Let Scan = True
End Function

' DESCRIPTION
' Applies a sequence of functions to an element, returning an array holding the result of applying
' each of the functions to the element. This function returns Null if the function array fails
' Predicates.StringArrayQ.  This function returns an empty array if the array of function names is empty.
' This function returns Null if AnElement is not atomic.
'
' PARAMETERS
' 1. AFunctionNameArray - An array of function names
' 2. AnElement - Any element to which each of the functions can be applied
' 3. ParameterCheckQ (optional) - If this explicitly set to False, no parameter consistency checks are
'    perform.
'
' RETURNED VALUE
' An array with the results of applying a sequence of the functions to an element.
Public Function Through(AFunctionNameArray As Variant, _
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
    End If
    
    ' Pre-allocate array to hold results
    ReDim ReturnArray(1 To Length(AFunctionNameArray))
    
    ' Compute values from applying each function to AnElement
    For c = 1 To Length(AFunctionNameArray)
        Let ReturnArray(c) = Run(AFunctionNameArray(c + LBound(AFunctionNameArray) - 1), _
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
' 2. A1DArray - An array of atomic elements
'
' RETURNED VALUE
' An array with the results of applying a sequence of the functions to an element.
Public Function ArrayMap(aFunctionName As String, A1DArray As Variant) As Variant
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
        Let ReturnArray(c) = Run(aFunctionName, Part(A1DArray, c))
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
' 2. AFunctionName - The name of the function to apply to each of the elements in the array
' 3. ParameterCheckQ (optional) - If this explicitly set to False, no parameter consistency checks are
'    perform.
'
' RETURNED VALUE
' An array with the results of applying a sequence of the functions to an element.
Public Function ArraySelect(AnArray As Variant, aFunctionName As String) As Variant
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
        Let ArraySelect = EmptyArray
        
        Exit Function
    End If
    
    ' Pre-allocated array to return at most all of the elements in the array
    ReDim ReturnArray(1 To Length(AnArray))
    
    ' Cycle through the array, adding to the return array those elements yielding True
    Let c = 0
    For i = 1 To Length(AnArray)
        If Run(aFunctionName, Part(AnArray, i)) Then
            Let c = c + 1
            Let ReturnArray(c) = Part(AnArray, i)
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
' Example: ArrayMapThread("add", array(1,2,3), array(10,20,30)) -> (11, 22, 33)
'
' PARAMETERS
' 1. AFunctionName - Name of the function to apply
' 2. ArrayOfEqualLength1DArrays - A sequence of equal-length, atomic arrays
'
' RETURNED VALUE
' An array with the results of applying the given function to the threading of the parameter arrays
Public Function MapThread(aFunctionName As String, _
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
    Dim SplicingDelegateName As String
    
    ' Set default return value
    Let MapThread = Null
    
    ' Make a copy of array of parameters so we may apply other functions to it
    Let ParamsArray = CopyParamArray(ArrayOfEqualLength1DArrays)
    
    ' Exit with Null if ParamsArray is not dimensioned
    If Not DimensionedQ(ParamsArray) Then Exit Function
    
    ' Exit with Null if ParamsArray is empty
    If EmptyArrayQ(ParamsArray) Then
        Let MapThread = EmptyArray
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
    
    ' Create splicing delegate name
    Let SplicingDelegateName = ParameterSplicingDelegate(aFunctionName, Length(ParamsArray))

    ' Loop over the array elements to compute results to return
    For r = 1 To N
        ' Assemble array of value for this function call
        For c = 1 To Length(ParamsArray)
            Let CallArray(c) = Part(Part(ParamsArray, c), r)
        Next
        
        Let ReturnArray(r) = Run(SplicingDelegateName, CallArray)
    Next
    
    ' Return the array used
    Let MapThread = ReturnArray
End Function

' DESCRIPTION
' This function computes the sum of the elements of the given array.  If AnArray is a 2D array,
' this function returns the sum of the columns.  This is equivalent to
'
' Total(AnArray, 1)
'
' To add the rows, use Total(AnArray, 2)
'
' PARAMETERS
' 1. AnArray - a 1D or 2D numeric array
' 2. DimensionalIndex (optional) - Set by default to 1, indicates whether to perform the operation
'    along dimension 1 or 2
'
' RETURNED VALUE
' The result of the operation the whole array for 1D arrays, the columns for 2D arrays, or the rows
' if requested for a 2D array
Public Function Total(AnArray As Variant, _
                      Optional DimensionalIndex As Integer = 1, _
                      Optional ParameterCheckQ As Boolean = True) As Variant
    Dim var As Variant
    Dim ResultArray As Variant
    Dim i As Integer
    Dim j As Integer
    
    Let Total = Null

    If ParameterCheckQ Then
       If Not DimensionedQ(AnArray) Then Exit Function
       
        If NumberOfDimensions(AnArray) < 1 Or NumberOfDimensions(AnArray) > 2 Then Exit Function
        
        For Each var In AnArray
            If Not NumberQ(var) Then Exit Function
        Next
    End If
    
    If EmptyArrayQ(AnArray) Then
        Let Total = 0
        Exit Function
    End If
    
    If DimensionalIndex = 2 Then
        If NumberOfDimensions(AnArray) = 1 Then
            Let ResultArray = AnArray
        Else
            ReDim ResultArray(1 To Length(AnArray))
        
            For i = 1 To Length(AnArray)
                Let ResultArray(i) = Total(Part(AnArray, i))
            Next i
        End If
    Else
        If NumberOfDimensions(AnArray) = 1 Then
            Let ResultArray = 0
            For Each var In AnArray
                Let ResultArray = ResultArray + var
            Next
        Else
            ReDim ResultArray(1 To NumberOfColumns(AnArray))
            
            For j = 1 To NumberOfColumns(AnArray)
                Let ResultArray(j) = Total(Part(AnArray, Span(1, -1), j))
            Next
        End If
    End If
    
    Let Total = ResultArray
End Function

' DESCRIPTION
' This function computes the product of the elements of the given array.  If AnArray is a
' 2D array, this function returns the product of the columns.  This is equivalent to
'
' Times(AnArray, 1)
'
' To add the rows, use Total(AnArray, 2)
'
' PARAMETERS
' 1. AnArray - a 1D or 2D numeric array
' 2. DimensionalIndex (optional) - Set by default to 1, indicates whether to perform the operation
'    along dimension 1 or 2
'
' RETURNED VALUE
' The result of the operation the whole array for 1D arrays, the columns for 2D arrays, or the rows
' if requested for a 2D array
Public Function Times(AnArray As Variant, _
                      Optional DimensionalIndex As Integer = 1, _
                      Optional ParameterCheckQ As Boolean = True) As Variant
    Dim var As Variant
    Dim ResultArray As Variant
    Dim i As Integer
    Dim j As Integer
    
    Let Times = Null

    If ParameterCheckQ Then
       If Not DimensionedQ(AnArray) Then Exit Function
       
        If NumberOfDimensions(AnArray) < 1 Or NumberOfDimensions(AnArray) > 2 Then Exit Function
        
        For Each var In AnArray
            If Not NumberQ(var) Then Exit Function
        Next
    End If
    
    If EmptyArrayQ(AnArray) Then
        Let Times = 1
        Exit Function
    End If
    
    If DimensionalIndex = 2 Then
        If NumberOfDimensions(AnArray) = 1 Then
            Let ResultArray = AnArray
        Else
            ReDim ResultArray(1 To Length(AnArray))
        
            For i = 1 To Length(AnArray)
                Let ResultArray(i) = Times(Part(AnArray, i))
            Next i
        End If
    Else
        If NumberOfDimensions(AnArray) = 1 Then
            Let ResultArray = 1
            For Each var In AnArray
                Let ResultArray = ResultArray * var
            Next
        Else
            ReDim ResultArray(1 To NumberOfColumns(AnArray))
            
            For j = 1 To NumberOfColumns(AnArray)
                Let ResultArray(j) = Times(Part(AnArray, Span(1, -1), j))
            Next
        End If
    End If
    
    Let Times = ResultArray
End Function

' DESCRIPTION
' This function computes the array that results from the successive addition of its elements. It
' adds repeatedly along columns, returning a 1D array of 1D arrays.  When, applied to a 2D matrix,
' it adds the rows repeatedly as if we had a sequence of rows, returning a 1D array of 1D
' arrays. At the moment is does not work along the second dimension (adding along columns)
'
' PARAMETERS
' 1. AnArray - a 1D or 2D numeric array
'
' RETURNED VALUE
' The array of successive sums of the elements of the array or the columns of the 2D array
Public Function Accumulate(AnArray As Variant, _
                           Optional ParameterCheckQ As Boolean = True) As Variant
    Dim var As Variant
    Dim ResultArray As Variant
    Dim i As Integer
    
    Let Accumulate = Null

    If ParameterCheckQ Then
       If Not DimensionedQ(AnArray) Then Exit Function
       
        If NumberOfDimensions(AnArray) < 1 Or NumberOfDimensions(AnArray) > 2 Then Exit Function
        
        For Each var In AnArray
            If Not NumberQ(var) Then Exit Function
        Next
    End If
    
    If EmptyArrayQ(AnArray) Then
        Let Accumulate = 0
        
        Exit Function
    End If
    
    If NumberOfDimensions(AnArray) = 1 Then
        ReDim ResultArray(1 To Length(AnArray))
    
        Let ResultArray(1) = First(AnArray)
        For i = 2 To Length(AnArray)
            Let ResultArray(i) = ResultArray(i - 1) + Part(AnArray, i)
        Next
    Else
        ' Pre-allocate result array
        ReDim ResultArray(1 To Length(AnArray))
    
        ' Set the initial value to the first row
        Let ResultArray(1) = Part(AnArray, 1)

        For i = 2 To Length(AnArray)
            Let ResultArray(i) = Add(ResultArray(i - 1), Part(AnArray, i))
        Next
    End If
    
    Let Accumulate = ResultArray
End Function

' DESCRIPTION
' Return the result of applying the given function N times iteratively to the given argument.
' Returns the argument when N = 0.  Returns Null when N<0.
'
' Example: ArrayMapThread("StringJoin", array(1,2,3), array(10,20,30)) returns
'          ("110", "220", "330")
'
' PARAMETERS
' 1. AFunctionName - Name of the function to apply
' 2. Arg - Initial value for initial function call
' 3. N - Number of times to apply functional nesting
'
' RETURNED VALUE
' Returns the Nth iteration of a function to an argument
Public Function Nest(aFunctionName As String, _
                     arg As Variant, _
                     N As Long) As Variant
    If N < 0 Then
        Let Nest = Null
    ElseIf N = 0 Then
        Let Nest = arg
    Else
        Let Nest = Nest(aFunctionName, Run(aFunctionName, arg), N - 1)
    End If
End Function

' DESCRIPTION
' Return an array recording the result of applying the given function iterativesly N
' times to the given argument. Returns the argument when N = 0.  Returns Null when N<0.
'
' Example: ArrayMapThread("StringJoin", array(1,2,3), array(10,20,30)) returns
'          ("110", "220", "330")
'
' PARAMETERS
' 1. AFunctionName - Name of the function to apply
' 2. Arg - Initial value for initial function call
' 3. N - Number of times to apply functional nesting
'
' RETURNED VALUE
' Returns array with the iterative application of a function to an argument
Public Function NestList(aFunctionName As String, _
                         arg As Variant, _
                         N As Long) As Variant
    Dim ResultArray As Variant
    Dim i As Long
    Dim CurrentValue As Variant
    
    If N < 0 Then
        Let ResultArray = Null
    ElseIf N = 0 Then
        Let ResultArray = Array(arg)
    Else
        ReDim ResultArray(1 To N + 1)
        
        Let ResultArray(1) = arg
        Let CurrentValue = arg
        For i = 2 To N + 1
            Let ResultArray(i) = Run(aFunctionName, CurrentValue)
            Let CurrentValue = ResultArray(i)
        Next
    End If
    
    Let NestList = ResultArray
End Function

' DESCRIPTION
' Return the result of applying the given function N times iteratively to the given
' argument. Returns the argument when N = 0.  Returns Null when N<0.
'
' The function with name aFunctionName must accept two arguments.
'
' Example: Fold("Times", 10, [{1,2,3}]) returns 60
'
' PARAMETERS
' 1. AFunctionName - Name of the function to apply
' 2. Arg - Initial value for initial function call
' 3. N - Number of times to apply functional nesting
'
' RETURNED VALUE
' Returns the Nth iteration of a function to an argument
Public Function Fold(aFunctionName As String, _
                     FirstArg As Variant, _
                     AnArrayForSecondArgs As Variant) As Variant
    Dim i As Long
    Dim CurrentValue As Variant
                     
    ' Set default return value for errors
    Let Fold = Null
    
    ' Exit with Null if AnArrayForSecordArgs is not dimensioned
    If Not DimensionedQ(AnArrayForSecondArgs) Then Exit Function
    
    ' Return an empty list if AnArrayForSecondArgs is not array or undimensioned
    If Not DimensionedQ(AnArrayForSecondArgs) Then Exit Function
    
    Let CurrentValue = FirstArg
    For i = 1 To Length(AnArrayForSecondArgs)
        Let CurrentValue = Run(aFunctionName, CurrentValue, Part(AnArrayForSecondArgs, i))
    Next
    
    Let Fold = CurrentValue
End Function

' DESCRIPTION
' Return an array with each step in the computation resulting from applying the given
' function N times iteratively to the given argument. Returns a 1D array with FirstArg
' as its sole element when N = 0. Returns Null when N<0.
'
' The function with name aFunctionName must accept two arguments.
'
' Example: FoldList("Times", 10, [{1,2,3}]) returns {1, 1, 2, 6}
'
' PARAMETERS
' 1. AFunctionName - Name of the function to apply
' 2. Arg - Initial value for initial function call
' 3. N - Number of times to apply functional nesting
'
' RETURNED VALUE
' Returns all steps in the N iterations of the function on the given list
Public Function FoldList(aFunctionName As String, _
                     FirstArg As Variant, _
                     AnArrayForSecondArgs As Variant) As Variant
    Dim i As Long
    Dim ResultArray() As Variant
                     
    ' Set default return value for errors
    Let FoldList = Null
    
    ' Exit with Null if AnArrayForSecordArgs is not dimensioned
    If Not DimensionedQ(AnArrayForSecondArgs) Then Exit Function
    
    ' Return a 1D, 1-elt array with FirstArg if AnArrayForSecondArgs empty
    If EmptyArrayQ(AnArrayForSecondArgs) Then
        Let FoldList = Array(FirstArg)
        Exit Function
    End If
    
    ReDim ResultArray(1 To 1 + Length(AnArrayForSecondArgs))
    Let ResultArray(1) = FirstArg
    For i = 1 To Length(AnArrayForSecondArgs)
        Let ResultArray(i + 1) = Run(aFunctionName, ResultArray(i), Part(AnArrayForSecondArgs, i))
    Next
    
    Let FoldList = ResultArray
End Function

' DESCRIPTION
' Creates an anonymous function and returns its name so it may be used
' by other functional programming constructs.
'
' Example: ArrayMap(Lambda("x", "", "2*x"), [{1,2,3,4,5}]) ->
'          Array(2,4,6,8,10)
'
'          This also returns a function name for the anonymous function.
'
' PARAMETERS
' 1. ParameterNameArray - A string or list of strings, with each string the
'    valid name of a variable
' 2. FunctionBody - Either a string representing a valid VBA function body,
'    or an array of strings representing a valid VBA function body
' 3. ReturnExpression - A valid VBA expression whose evaluation is returned by
'    the anonymous function
'
' RETURNED VALUE
' Returns the name of the anonymous function
Public Function Lambda(ParameterNameArray As Variant, FunctionBody As Variant, ReturnExpression As String) As String
    Dim AnonCounter As Integer
    Dim FunctionName As String
    
    ' Get the current lambda counter
    Let AnonCounter = CInt(Right(ThisWorkbook.Names("LambdaFunctionCounter").Value, _
                                 Len(ThisWorkbook.Names("LambdaFunctionCounter").Value) - 1))
    
    
    ' Generate new, unique lambda name
    Let FunctionName = "Lambda" & AnonCounter
    
    ' Increase lambda counter
    Let ThisWorkbook.Names("LambdaFunctionCounter").Value = AnonCounter + 1
    
    ' Generate new lambda function
    Call InsertFunction(ThisWorkbook, _
                        "LambdaFunctionsTemp", _
                        FunctionName, _
                        IIf(StringQ(ParameterNameArray), Array(ParameterNameArray), ParameterNameArray), _
                        Append(IIf(StringQ(FunctionBody), Array(FunctionBody), FunctionBody), _
                               vbCrLf & "Let " & FunctionName & "=" & ReturnExpression))
    
    ' Return fully-qualified function name
    Let Lambda = "'" & ThisWorkbook.Name & "'!" & FunctionName
End Function

' DESCRIPTION
' Returns the result from the application of a function with the values of the
' given array as the values of its parameters. This function does not checking to
' ensure that the length of the parameter array and the number of required funcitonal
' parameters are equal.
'
' Example: Apply(Lambda([{"x","y"}], "", "2*x*y"), [{1,2}]) ->
'          Lambda([{"x"<-1,"y"<-2}], "", "2*x*y") -> 2*1*2 = 4
'
' PARAMETERS
' 1. FunctionName - The name of the function
' 2. ParameterArray - The array holding the values to evaluate the function
'
' RETURNED VALUE
' The result from evaluating the function with parameter values specified by the
' given array
Public Function Apply(ByVal FunctionName As String, _
                      ByVal ParameterArray As Variant) As Variant
    Dim ParentExpr As String
    Dim ParamNames As Variant
    Dim FunctionBody() As String
    Dim VarName As Variant
    Dim i As Long
    Dim LambdaName As String
    Dim LambdaCounter As Integer

    ' Set default return value in case of error
    Let Apply = Null
    
    ' ErrorCheck: Exit with Null if ParameterArray not an array
    If Not IsArray(ParameterArray) Then Exit Function
    
    ' ErrorCheck: Exit with Null if ParameterArray not initialized and non-empty
    If Not DimensionedQ(ParameterArray) Then Exit Function
    
    ' ErrorCheck: Exit with Null if ParameterArray is empty
    If EmptyArrayQ(ParameterArray) Then Exit Function
    
    ' Create delegating function
    Let LambdaName = ParameterSplicingDelegate(FunctionName, Length(ParameterArray))
    
    ' Apply the delegation funciton to the parameter array
    Let Apply = Run(LambdaName, ParameterArray)
End Function

' DESCRIPTION
' Creates and returns the name of a function to allow splice the contents of a 1D array
' as the arguments of a function. This function provides the same functionality as the
' apply function in EmacsLisp.
'
' There is no way in to take a<-array(1,2) and then apply("add", a) to get add(1,2)
'
' This function returns a function name that can be passed to Apply to achieve this
' behavior. One does not need to call ParameterSplicingDelegate directly. This helper
' function is called by Apply. However, it is also called by MapThread to achieve similar
' behavior.
'
' PARAMETERS
' 1. FunctionName - The name of the function
' 2. N - A positive integer indicating the number of arguments the delegate must accept
'
' RETURNED VALUE
' Returns the name of the delegate function for splicing an array into the parameter
' slots of the given function
Public Function ParameterSplicingDelegate(FunctionName As String, N As Integer) As Variant
    Dim ParamNames As Variant
    Dim ParenString As String
    Dim FunctionBody() As String
    Dim VarName As Variant
    Dim i As Long
    
    ' Set default return value in case of error
    Let ParameterSplicingDelegate = Null
    
    ' ErrorCheck: Exit with Null if N not a positive integer
    If Not PositiveWholeNumberQ(N) Then Exit Function
    
    ' Create parameter list for anonymous function
    Let ParamNames = GenerateStringSequence("Param", 1, N)
    
    ' Create the function body
    
    ' Construct a let statement for bind each var to its intended value
    ReDim FunctionBody(1 To 2 * N)
    For i = 1 To N
        Let FunctionBody(i) = _
            "Dim " & Part(ParamNames, i) & " as Variant"
    Next
    
    For i = N + 1 To 2 * N
        Let FunctionBody(i) = _
            "Let " & Part(ParamNames, i - N) & " = Part(ArrayToSplice," & i - N & ")"
    Next
        
    ' Create the anonymous function
    Let ParenString = ToParentheticalString(ParamNames)
    Let ParenString = Right(ParenString, Len(ParenString) - 1)
    Let ParameterSplicingDelegate = Lambda("ArrayToSplice", _
                                           FunctionBody, _
                                           "run(" & Chr(34) & FunctionName & Chr(34) & _
                                          "," & ParenString _
                                          )
End Function

' DESCRIPTION
' Clears the contents of module LambdaFunctionsTemp and resets workbook name
' LambdaFunctionCounter to 0.
'
' PARAMETERS
' None
'
' RETURNED VALUE
' None
Public Sub ClearLambdaFunctionsData()
    Dim CodeModule As VBIDE.CodeModule

    ' Simply continue without raising error if the module is not found in
    ' the given module
    On Error Resume Next
    
    ' Set reference to approrpriate code module
    Set CodeModule = ThisWorkbook.VBProject.VBComponents("LambdaFunctionsTemp").CodeModule

    ' Delete the funnction from the code module
    Call CodeModule.DeleteLines(1, CodeModule.CountOfLines)
    
    ' Reset lambda counter
    Let ThisWorkbook.Names("LambdaFunctionCounter").Value = 0
End Sub
