Attribute VB_Name = "ClassConstructors"
Option Base 1
Option Explicit

' DESCRIPTION
' This function returns an instance of class Span .Value holds the result of running
' Arrays.CreateSequentialArray.  If the result of running Arrays.CreateSequentialArray
' is Null, the .Value is set to Null.
'
' For details on how to call this function, see Arrays.CreateSequentialArray.
'
' PARAMETERS
' 1. StartNumber - First number in the array
' 2. N - Number of elements in the sequence or the ending number, depending on the calling modality
' 3. TheStep (optional) - To create a sequence using a sequential step different from 1
' 4. ToEndNumberQ (optional) - When passed explicitly as True, it activates calling modality 2
'
' RETURNED VALUE
' An object requested numerical sequence
Public Function Span(StartNumber As Variant, _
                     N As Variant, _
                     Optional TheStep As Variant, _
                     Optional ToEndNumberQ As Boolean = False) As Span
    Dim obj As New Span
    Dim TheStepCopy As Variant
    
    Set obj = New Span
    
    Let TheStepCopy = IIf(IsMissing(TheStep), 1, TheStep)
    
    If ToEndNumberQ Then
        Let obj.Value = CreateSequentialArray(StartNumber, N, TheStepCopy, ToEndNumberQ)
    Else
        Let obj.Value = CreateSequentialArray(StartNumber, N, TheStepCopy)
    End If
    
    ' Return object
    Set Span = obj
End Function
