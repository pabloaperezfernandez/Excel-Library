Attribute VB_Name = "ClassConstructors"
Option Base 1
Option Explicit

' DESCRIPTION
' This function returns an instance of class Span. It makes sense only when used in the
' context of a given array and relative to one of the array's dimensions.
'
' For details on how to call this function, see Arrays.NumericalSequence.
'
' PARAMETERS
' 1. StartNumber - First number in the span
' 2. EndNumber - Last number in the span
' 3. TheStep (optional) - To create a sequence using a sequential step different from 1
'
' RETURNED VALUE
' An object representing the desired span of indices
Public Function Span(StartNumber As Long, _
                     EndNumber As Long, _
                     Optional TheStep As Long = 1) As Span
    Dim obj As New Span
    
    Let obj.TheStart = StartNumber
    Let obj.TheEnd = EndNumber
    Let obj.TheStep = TheStep
    
    ' Return object
    Set Span = obj
End Function

' DESCRIPTION
' Creates an anonymous function and returns it as an instance of class Lambda
'
' Example: ArrayMap(Lambda("x", "", "2*x"), [{1,2,3,4,5}]) ->
'          Array(2,4,6,8,10)
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
' Returns an instance of class Lambda (an anonymous function)
Public Function Lambda(ParameterNameArray As Variant, _
                       FunctionBody As Variant, _
                       ReturnExpression As String) As Lambda
    Dim LambdaCounter As Integer
    Dim FunctionName As String
    Dim obj As New Lambda
    
    ' Get the current lambda counter
    Let LambdaCounter = CInt(Right(ThisWorkbook.Names("LambdaFunctionCounter").Value, _
                                 Len(ThisWorkbook.Names("LambdaFunctionCounter").Value) - 1))
    
    
    ' Generate new, unique lambda name
    Let FunctionName = "Lambda" & LambdaCounter
    
    ' Increase lambda counter
    Let ThisWorkbook.Names("LambdaFunctionCounter").Value = LambdaCounter + 1
    
    ' Generate new lambda function
    Call InsertFunction(ThisWorkbook, _
                        "LambdaFunctionsTemp", _
                        FunctionName, _
                        IIf(StringQ(ParameterNameArray), Array(ParameterNameArray), ParameterNameArray), _
                        Append(IIf(StringQ(FunctionBody), Array(FunctionBody), FunctionBody), _
                               vbCrLf & "Let " & FunctionName & "=" & ReturnExpression))
    
    ' Set the function name for Lambda instance
    Let obj.FunctionName = MakeRoutineName(ThisWorkbook, "LambdaFunctionsTemp", FunctionName)

    ' Return the newly created Lambda class instance
    Set Lambda = obj
End Function

