Attribute VB_Name = "FunctionalPredicates"
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
' True or False depending on whether arg is dimensioned and all its elements satisfy
' ALambda or are True when ALambda is missing.
Public Function AllTrueQ(AnArray As Variant, _
                         Optional ALambda As Variant) As Variant
    Dim var As Variant
    Dim ProcName As String
    
    ' Set the default return value of True
    Let AllTrueQ = True
    
    ' ErrorCheck: Exit with False if AnArrat is not dimensioned
    If Not DimensionedQ(AnArray) Then
        Let AllTrueQ = False
        Exit Function
    End If
    
    ' ErrorCheck: Exit with Null if ALambdaOrFunctionName is neither a Lambda or a string
    If Not IsMissing(ALambda) Then
        If Not (LambdaQ(ALambda) Or StringQ(ALambda)) Then Exit Function
    End If
    
    ' Exit the True of AnArray is an empty array
    If EmptyArrayQ(AnArray) Then Exit Function
    
    ' Get the name of the predicate if provided
    If Not IsMissing(ALambda) Then
        If StringQ(ALambda) Then
            Let ProcName = ALambda
        Else
            Let ProcName = ALambda.FunctionName
        End If
    End If
    
    On Error GoTo ErrorHandler
    
    ' Exit with False if any of the entries is not Boolean or is Boolean and False
    ' Splits into two cases: predicate given and not given
    ' Cannot use BooleanArrayQ here because it would cause a circular definitional reference
    If IsMissing(ALambda) Then
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
                         Optional ALambda As Variant) As Variant
    Dim var As Variant
    Dim ProcName As String
    
    ' Set the default return value of True
    Let AnyTrueQ = False

    ' ErrorCheck: Exit with False if AnArray is not dimensioned array
    If Not DimensionedQ(AnArray) Then Exit Function
    
    ' ErrorCheck: Exit with False if ALambdaOrFunctionName is neither a Lambda or a string
    If Not IsMissing(ALambda) Then
        If Not (LambdaQ(ALambda) Or StringQ(ALambda)) Then Exit Function
    End If
    
    ' Exit the False if AnArray is an empty array
    If EmptyArrayQ(AnArray) Then Exit Function
    
    ' Get the name of the predicate if provided
    If Not IsMissing(ALambda) Then
        If StringQ(ALambda) Then
            Let ProcName = ALambda
        Else
            Let ProcName = ALambda.FunctionName
        End If
    End If
    
    On Error GoTo ErrorHandler
    
    ' Exit with Null if ALambda not missing and ALambda not a string
    ' Exit with Null if ALambda missing and AnArray is not an array of Booleans
    ' Cannot use BooleanArrayQ here because it would cause a circular definitional reference
    If IsMissing(ALambda) Then
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
' 2. ALambda (optional) - A string representing the predicates name
' 3. WorkbookReference (optional) - A workbook reference to the workbook holding
'    the predicate
'
' RETURNED VALUE
' True or False depending on whether arg is dimensioned and all its elements fail
' the predicate with name ALambda or are False when ALambda is missing.
Public Function NoneTrueQ(AnArray As Variant, _
                          Optional ALambda As Variant) As Boolean
    Dim var As Variant
    Dim ProcName As String
                              
    Let NoneTrueQ = True
    
    ' Exit with False if AnArray is not dimensioned
    If Not DimensionedQ(AnArray) Then
        Let NoneTrueQ = False
        Exit Function
    End If
    
    ' ErrorCheck: Exit with False if ALambdaOrFunctionName is neither a Lambda or a string
    If Not IsMissing(ALambda) Then
        If Not (LambdaQ(ALambda) Or StringQ(ALambda)) Then Exit Function
    End If
    
    ' Exit with True if AnArray is empty. This case is necessary because NoneTrueQ is not logically
    ' the negation of AllTrueQ.  For NoneTrueQ to be true, all elements of AnArray must be False.
    ' However, all elements of AnArray are False if AnArray is an empty set.
    If EmptyArrayQ(AnArray) Then Exit Function

    ' Get the name of the predicate if provided
    If Not IsMissing(ALambda) Then
        If StringQ(ALambda) Then
            Let ProcName = ALambda
        Else
            Let ProcName = ALambda.FunctionName
        End If
    End If
    
    On Error GoTo ErrorHandler

    ' Exit with Null if ALambda not missing and ALambda not a string
    ' Exit with Null if ALambda missing and AnArray is not an array of Booleans
    ' Cannot use BooleanArrayQ here because it would cause a circular definitional reference
    If IsMissing(ALambda) Then
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
' Boolean function returning True if all of the elements in AnArray fail the predicate whose
' name is PredicateName.  If PredicateName is not provided, then the function returns True when
' all of the elements of AnArray are False.  The predicate returns Null if PredicateName is not a
' string. Returns True for an empty array.
'
' PARAMETERS
' 1. AnArray - A dimensioned 1D array
' 2. PredicateName (optional) - A string representing the predicates name
'
' RETURNED VALUE
' True or False depending on whether arg is dimensioned and all its elements fails the
' predicate with name PredicateName or are True when PredicateName is missing.
Public Function AllFalseQ(AnArray As Variant, _
                         Optional PredicateName As Variant) As Variant
    Dim var As Variant
    
    ' Set the default return value of True
    Let AllFalseQ = True
    
    ' Exit with False if AnArrat is not dimensioned
    If Not DimensionedQ(AnArray) Then
        Let AllFalseQ = False
        Exit Function
    End If
    
    ' Exit the True of AnArray is an empty array
    If EmptyArrayQ(AnArray) Then Exit Function
    
    ' Exit with Null if PredicateName not missing and PredicateName not a string
    ' Exit with Null if PredicateName missing and AnArray is not an array of Booleans
    ' Cannot use BooleanArrayQ here because it would cause a circular definitional reference
    If IsMissing(PredicateName) Then
        For Each var In AnArray
            If Not BooleanQ(var) Then
                Let AllFalseQ = False
                Exit Function
            End If
        Next
    Else
        If Not StringQ(PredicateName) Then
            Let AllFalseQ = Null
            Exit Function
        End If
    End If
    
    If Not IsMissing(PredicateName) Then
        For Each var In AnArray
            If Run(PredicateName, var) Then
                Let AllFalseQ = False
                Exit Function
            End If
        Next
    Else
        For Each var In AnArray
            If var Then
                Let AllFalseQ = False
                Exit Function
            End If
        Next
    End If
End Function

' DESCRIPTION
' Boolean function returning True if at least one of the elements in AnArray fails the
' predicate whose name is PredicateName. If PredicateName is missing, the function returns
' True if any of the elements in AnArray is True.  The predicate returns Null if PredicateName
' is not a string.  The function returns False for an empty array.
'
' PARAMETERS
' 1. arg - any value or object reference
' 2. PredicateName (optional) - A string representing the predicates name
'
' RETURNED VALUE
' True or False depending on whether arg is dimensioned and at least one of its elements fails
' the predicate with name PredicateName
Public Function AnyFalseQ(AnArray As Variant, _
                         Optional PredicateName As Variant) As Variant
    Dim var As Variant
    
    ' Set the default return value of True
    Let AnyFalseQ = False

    ' Exit with False if AnArray is not dimensioned array
    If Not DimensionedQ(AnArray) Then
        Let AnyFalseQ = False
        Exit Function
    End If
    
    ' Exit the False if AnArray is an empty array
    If EmptyArrayQ(AnArray) Then
        Exit Function
    End If
    
    ' Exit with Null if PredicateName not missing and PredicateName not a string
    ' Exit with Null if PredicateName missing and AnArray is not an array of Booleans
    ' Cannot use BooleanArrayQ here because it would cause a circular definitional reference
    If IsMissing(PredicateName) Then
        For Each var In AnArray
            If Not BooleanQ(var) Then
                Let AnyFalseQ = False
                Exit Function
            End If
        Next
    Else
        If Not StringQ(PredicateName) Then
            Let AnyFalseQ = Null
            Exit Function
        End If
    End If
    
    If Not IsMissing(PredicateName) Then
        For Each var In AnArray
            If Not Run(PredicateName, var) Then
                Let AnyFalseQ = True
                Exit Function
            End If
        Next
    Else
        For Each var In AnArray
        If Not var Then
            Let AnyFalseQ = True
            Exit Function
        End If
        Next
    End If
End Function

' DESCRIPTION
' Boolean function returning True if none of the elements in AnArray fail the predicate whose
' name is PredicateName.  If PredicateName is missing, the function returns True when all
' of AnArray are False.
'
' PARAMETERS
' 1. AnArray - A dimensioned 1D or 2D array
' 2. PredicateName (optional) - A string representing the predicates name
' 3. WorkbookReference (optional) - A workbook reference to the workbook holding the predicate
'
' RETURNED VALUE
' True or False depending on whether arg is dimensioned and all its elements fail the
' predicate with name PredicateName or are False when PredicateName is missing.
Public Function NoneFalseQ(AnArray As Variant, _
                          Optional PredicateName As Variant) As Boolean
    Dim var As Variant
    
    Let NoneFalseQ = True
                          
    ' Exit with Null if AnArray is not dimensioned
    If Not DimensionedQ(AnArray) Then
        Let NoneFalseQ = False
        Exit Function
    End If
    
    ' Exit with True if AnArray is empty. This case is necessary because NoneTrueQ is not logically
    ' the negation of AllTrueQ.  For NoneTrueQ to be true, all elements of AnArray must be False.
    ' However, all elements of AnArray are False if AnArray is an empty set.
    If EmptyArrayQ(AnArray) Then Exit Function

    ' Exit with Null if PredicateName not missing and PredicateName not a string
    ' Exit with Null if PredicateName missing and AnArray is not an array of Booleans
    ' Cannot use BooleanArrayQ here because it would cause a circular definitional reference
    If IsMissing(PredicateName) Then
        For Each var In AnArray
            If Not BooleanQ(var) Then
                Let NoneFalseQ = False
                Exit Function
            End If
        Next
    Else
        If Not StringQ(PredicateName) Then
            Let NoneFalseQ = False
            Exit Function
        End If
    End If
    
    For Each var In AnArray
        If Not IsMissing(PredicateName) Then
            If Application.Run(PredicateName, var) Then
                Let NoneFalseQ = False
                Exit Function
            End If
        Else
            If var Then
                Let NoneFalseQ = False
                Exit Function
            End If
        End If
    Next
End Function

' DESCRIPTION
' Boolean function returning True if TheValue is in the given 1D array.
'
' PARAMETERS
' 1. TheArray - A 1D array
' 2. TheValue - Any Excel value or reference
'
' RETURNED VALUE
' Returns True or False depending on whether or not the given value is in the given array
Public Function MemberQ(TheArray As Variant, TheValue As Variant) As Boolean
    Dim i As Long
    Dim var As Variant
    
    ' Assume result is False and change TheValue is in any one column of TheArray
    Let MemberQ = False
    
    ' Exit if TheArray is not a 1D array
    If NumberOfDimensions(TheArray) <> 1 Then Exit Function
    
    For Each var In TheArray
        If IsError(var) Then
            Let MemberQ = False
            Exit Function
        End If
        
        If IsEmpty(var) And IsEmpty(TheValue) Then
            Let MemberQ = True
            Exit Function
        End If
        
        If IsNull(var) And IsNull(TheValue) Then
            Let MemberQ = True
            Exit Function
        End If
        
        If IsObject(var) Then
            Let MemberQ = TheValue Is var
            Exit Function
        End If

        If VarType(var) = VarType(TheValue) And var = TheValue Then
            Let MemberQ = True
            Exit Function
        End If
    Next
End Function

' DESCRIPTION
' Boolean function returning True if TheValue is not in the given 1D array. TheValue must
' satisfy NumberOrStringQ. Every element in TheArray must also satisfy NumberOfStringQ
'
' PARAMETERS
' 1. TheArray - A 1D array satisfying PrintableArrayQ
' 2. TheValue - Any value satisfying PrintableQ
'
' RETURNED VALUE
' Returns True or False depending on whether or not the given value is in the given array
Public Function FreeQ(TheArray As Variant, TheValue As Variant) As Boolean
    Let FreeQ = IsArray(TheArray) And Not MemberQ(TheArray, TheValue)
End Function
