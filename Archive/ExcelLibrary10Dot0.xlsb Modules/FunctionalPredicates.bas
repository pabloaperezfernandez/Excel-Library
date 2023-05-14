Attribute VB_Name = "FunctionalPredicates"
' PURPOSE OF THIS MODULE
'
' The purpose of this module is to a collection of fundamental functional
' predicates. Examples include AllTrueQ(AnArray, Optional ALambda) that
' accept an optional Lambda. All of the functions in this module return
' Booleans.
'
' By convention, predicates in this module return True when acting on an
' empty array or any object or value that is not an array. From a logical
' point of view, AllTrueQ returns True for empty arrays and anything that
' is not an array because the statement "for all x in S," is vacously true.

Option Explicit
Option Base 1

' DESCRIPTION
' Boolean function returning True if all of the elements in AnArray satisfy the
' given predicate. If ALambda is not provided, then the function returns True
' when all of the elements of AnArray are True.  The predicate returns Null if
' ALambda is not a string. Returns True for an empty array.
'
' PARAMETERS
' 1. AnArray - A dimensioned 1D or 2D array
' 2. ALambda - An instance of class Lambda or string name of a function
'
' RETURNED VALUE
' True or False depending on whether arg is dimensioned and all its elements
' satisfy ALambda or are True when ALambda is missing.
Public Function AllTrueQ(AnArray As Variant, _
                         Optional aLambda As Variant) As Variant
    Dim var As Variant
    Dim ProcName As String
    
    ' Set the default return value of True
    Let AllTrueQ = True
    
    ' ErrorCheck: Exit with False if AnArrat is not dimensioned
    If Not DimensionedQ(AnArray) Then
        Let AllTrueQ = False
        Exit Function
    End If
    
    ' ErrorCheck: Exit with Null if ALambdaOrFunctionName is neither a
    ' Lambda or a string
    If Not IsMissing(aLambda) Then
        If Not (LambdaQ(aLambda) Or StringQ(aLambda)) Then Exit Function
    End If
    
    ' Exit the True of AnArray is an empty array
    If EmptyArrayQ(AnArray) Then Exit Function
    
    ' Get the name of the predicate if provided
    If Not IsMissing(aLambda) Then
        If StringQ(aLambda) Then
            Let ProcName = aLambda
        Else
            Let ProcName = aLambda.FunctionName
        End If
    End If
    
    On Error GoTo ErrorHandler
    
    ' Exit with False if any of the entries is not Boolean or is Boolean and False
    ' Splits into two cases: predicate given and not given
    ' Cannot use BooleanArrayQ here because it would cause a circular definitional reference
    If IsMissing(aLambda) Then
        For Each var In AnArray
            If FalseQ(var) Then
                Let AllTrueQ = False
                Exit Function
            End If
        Next
    Else
        For Each var In AnArray
            If FalseQ(Run(ProcName, var)) Then
                Let AllTrueQ = False
                Exit Function
            End If
        Next
    End If
    
    Exit Function

ErrorHandler:
    Let AllTrueQ = False
End Function

' DESCRIPTION
' Boolean function returning True if at least one of the elements in AnArray satisfy the
' predicate ALambda. If ALambda is missing, the function returns True if any of the
' elements in AnArray is True.  The predicate returns Null if ALambda is not a string.
' The function returns False for an empty array.
'
' PARAMETERS
' 1. arg - any value or object reference
' 2. ALambda - An instance of class Lambda or string name of a function
'
' RETURNED VALUE
' True or False depending on whether arg is dimensioned and at least one of its elements
' satisfies the predicate with name ALambda
Public Function AnyTrueQ(AnArray As Variant, _
                         Optional aLambda As Variant) As Variant
    Dim var As Variant
    Dim ProcName As String
    
    ' Set the default return value of True
    Let AnyTrueQ = False

    ' ErrorCheck: Exit with False if AnArray is not dimensioned array
    If Not DimensionedQ(AnArray) Then Exit Function
    
    ' ErrorCheck: Exit with False if ALambdaOrFunctionName is neither a Lambda or a string
    If Not IsMissing(aLambda) Then
        If Not (LambdaQ(aLambda) Or StringQ(aLambda)) Then Exit Function
    End If
    
    ' Exit the False if AnArray is an empty array
    If EmptyArrayQ(AnArray) Then Exit Function
    
    ' Get the name of the predicate if provided
    If Not IsMissing(aLambda) Then
        If StringQ(aLambda) Then
            Let ProcName = aLambda
        Else
            Let ProcName = aLambda.FunctionName
        End If
    End If
    
    On Error GoTo ErrorHandler
    
    ' Exit with Null if ALambda not missing and ALambda not a string
    ' Exit with Null if ALambda missing and AnArray is not an array of Booleans
    ' Cannot use BooleanArrayQ here because it would cause a circular definitional reference
    If IsMissing(aLambda) Then
        For Each var In AnArray
            If TrueQ(var) Then
                Let AnyTrueQ = True
                Exit Function
            End If
        Next
    Else
        For Each var In AnArray
            If TrueQ(Run(ProcName, var)) Then
                Let AnyTrueQ = True
                Exit Function
            End If
        Next
    End If
    
    Exit Function

ErrorHandler:
    Let AnyTrueQ = False
End Function

' DESCRIPTION
' Boolean function returning True if all of the elements in AnArray fail the
' predicate ALambda.  If ALambda is missing, the function returns True when all
' of AnArray are False.
'
' PARAMETERS
' 1. AnArray - A dimensioned 1D or 2D array
' 2. ALambda - An instance of class Lambda or string name of a function
'
' RETURNED VALUE
' True or False depending on whether arg is dimensioned and all its elements fail
' the predicate with name ALambda or are False when ALambda is missing.
Public Function NoneTrueQ(AnArray As Variant, _
                          Optional aLambda As Variant) As Boolean
    Dim var As Variant
    Dim ProcName As String
                              
    Let NoneTrueQ = True
    
    ' Exit with False if AnArray is not dimensioned
    If Not DimensionedQ(AnArray) Then
        Let NoneTrueQ = False
        Exit Function
    End If
    
    ' ErrorCheck: Exit with False if ALambdaOrFunctionName is neither a Lambda or a string
    If Not IsMissing(aLambda) Then
        If Not (LambdaQ(aLambda) Or StringQ(aLambda)) Then Exit Function
    End If
    
    ' Exit with True if AnArray is empty. This case is necessary because NoneTrueQ is not logically
    ' the negation of AllTrueQ.  For NoneTrueQ to be true, all elements of AnArray must be False.
    ' However, all elements of AnArray are False if AnArray is an empty set.
    If EmptyArrayQ(AnArray) Then Exit Function

    ' Get the name of the predicate if provided
    If Not IsMissing(aLambda) Then
        If StringQ(aLambda) Then
            Let ProcName = aLambda
        Else
            Let ProcName = aLambda.FunctionName
        End If
    End If
    
    On Error GoTo ErrorHandler

    ' Exit with Null if ALambda not missing and ALambda not a string
    ' Exit with Null if ALambda missing and AnArray is not an array of Booleans
    ' Cannot use BooleanArrayQ here because it would cause a circular definitional reference
    If IsMissing(aLambda) Then
        For Each var In AnArray
            If TrueQ(var) Then
                Let NoneTrueQ = False
                Exit Function
            End If
        Next
    Else
        For Each var In AnArray
            If TrueQ(Application.Run(ProcName, var)) Then
                Let NoneTrueQ = False
                Exit Function
            End If
        Next
    End If
    
    Exit Function

ErrorHandler:
    Let NoneTrueQ = False
End Function

' DESCRIPTION
' Boolean function returning True if all of the elements in AnArray fail the
' predicate ALambda.  If ALambda is not provided, then the function returns
' True when all of the elements of AnArray are False.  The predicate returns
' Null if ALambda is not a string. Returns True for an empty array.
'
' PARAMETERS
' 1. AnArray - A dimensioned 1D array
' 2. ALambda - An instance of class Lambda or string name of a function
'
' RETURNED VALUE
' True or False depending on whether arg is dimensioned and all its elements
' fails the predicate with name ALambda or are True when ALambda is missing.
Public Function AllFalseQ(AnArray As Variant, _
                          Optional aLambda As Variant) As Variant
    Dim ProcName As String
    Dim var As Variant
    
    ' Set the default return value of True
    Let AllFalseQ = True
    
    ' ErrorCheck: Exit with False if AnArrat is not dimensioned
    If Not DimensionedQ(AnArray) Then
        Let AllFalseQ = False
        Exit Function
    End If
    
    ' ErrorCheck: Exit with False if ALambdaOrFunctionName is neither a Lambda or a string
    If Not IsMissing(aLambda) Then
        If Not (LambdaQ(aLambda) Or StringQ(aLambda)) Then Exit Function
    End If
    
    ' Exit the True of AnArray is an empty array
    If EmptyArrayQ(AnArray) Then Exit Function
    
    ' Get the name of the predicate if provided
    If Not IsMissing(aLambda) Then
        If StringQ(aLambda) Then
            Let ProcName = aLambda
        Else
            Let ProcName = aLambda.FunctionName
        End If
    End If
    
    On Error GoTo ErrorHandler
    
    ' Cannot use BooleanArrayQ here because it would cause a circular definitional reference
    If IsMissing(aLambda) Then
        For Each var In AnArray
            If TrueQ(var) Then
                Let AllFalseQ = False
                Exit Function
            End If
        Next
    Else
        For Each var In AnArray
            If TrueQ(Run(ProcName, var)) Then
                Let AllFalseQ = False
                Exit Function
            End If
        Next
    End If

    Exit Function

ErrorHandler:
    Let AllFalseQ = False
End Function

' DESCRIPTION
' Boolean function returning True if at least one of the elements in AnArray fails the
' predicate whose name is ALambda. If ALambda is missing, the function returns
' True if any of the elements in AnArray is True.  The predicate returns Null if
' ALambda is not a string.  The function returns False for an empty array.
'
' PARAMETERS
' 1. arg - any value or object reference
' 2. ALambda - An instance of class Lambda or string name of a function
'
' RETURNED VALUE
' True or False depending on whether arg is dimensioned and at least one of its elements
' fails the predicate with name ALambda
Public Function AnyFalseQ(AnArray As Variant, _
                         Optional aLambda As Variant) As Variant
    Dim ProcName As String
    Dim var As Variant
    
    ' Set the default return value of True
    Let AnyFalseQ = False

    ' Exit with False if AnArray is not dimensioned array
    If Not DimensionedQ(AnArray) Then
        Let AnyFalseQ = False
        Exit Function
    End If
    
    ' ErrorCheck: Exit with False if ALambdaOrFunctionName is neither a Lambda or a string
    If Not IsMissing(aLambda) Then
        If Not (LambdaQ(aLambda) Or StringQ(aLambda)) Then Exit Function
    End If
    
    ' Exit the False if AnArray is an empty array
    If EmptyArrayQ(AnArray) Then
        Exit Function
    End If
    
    ' Get the name of the predicate if provided
    If Not IsMissing(aLambda) Then
        If StringQ(aLambda) Then
            Let ProcName = aLambda
        Else
            Let ProcName = aLambda.FunctionName
        End If
    End If
    
    On Error GoTo ErrorHandler
    
    ' Cannot use BooleanArrayQ here because it would cause a circular definitional reference
    If IsMissing(aLambda) Then
        For Each var In AnArray
            If FalseQ(var) Then
                Let AnyFalseQ = True
                Exit Function
            End If
        Next
    Else
        For Each var In AnArray
            If FalseQ(Run(ProcName, var)) Then
                Let AnyFalseQ = True
                Exit Function
            End If
        Next
    End If
    
    Exit Function

ErrorHandler:
    Let AnyFalseQ = False
End Function

' DESCRIPTION
' Boolean function returning True if none of the elements in AnArray fail the
' predicate ALambda.  If ALambda is missing, the function returns True when
' all of AnArray are False.
'
' PARAMETERS
' 1. AnArray - A dimensioned 1D or 2D array
' 2. ALambda (optional) - A string representing the predicates name
'
' RETURNED VALUE
' True or False depending on whether arg is dimensioned and all its elements fail
' the predicate ALambda or are False when ALambda is missing.
Public Function NoneFalseQ(AnArray As Variant, Optional aLambda As Variant) As Boolean
    Dim ProcName As String
    Dim var As Variant
    
    Let NoneFalseQ = True
                          
    ' Exit with True if AnArray is not dimensioned
    If Not DimensionedQ(AnArray) Then Exit Function
    
    ' ErrorCheck: Exit with False if ALambda is not missing and isneither
    ' a Lambda or a string
    If Not IsMissing(aLambda) Then
        If Not (LambdaQ(aLambda) Or StringQ(aLambda)) Then Exit Function
    End If
    
    ' Exit with True if AnArray is empty. This case is necessary because NoneTrueQ is
    ' not logically the negation of AllTrueQ.  For NoneTrueQ to be true, all elements
    ' of AnArray must be False. However, all elements of AnArray are False if AnArray
    ' is an empty set.
    If EmptyArrayQ(AnArray) Then Exit Function

    ' Get the name of the predicate if provided
    If Not IsMissing(aLambda) Then
        If StringQ(aLambda) Then
            Let ProcName = aLambda
        Else
            Let ProcName = aLambda.FunctionName
        End If
    End If
    
    On Error GoTo ErrorHandler
    
    ' Cannot use BooleanArrayQ here because it would cause a circular definitional reference
    If IsMissing(aLambda) Then
        For Each var In AnArray
            If FalseQ(var) Then
                Let NoneFalseQ = False
                Exit Function
            End If
        Next
    Else
        For Each var In AnArray
            If FalseQ(Run(ProcName, var)) Then
                Let NoneFalseQ = False
                Exit Function
            End If
        Next
    End If

    Exit Function

ErrorHandler:
    Let NoneFalseQ = False
End Function
