Attribute VB_Name = "ClassConstructors"
Option Base 1
Option Explicit

' DESCRIPTION
' This function returns an instance of class Span. It makes sense only when used in the context
' of a given array and relative to one of the array's dimensions.
'
' For details on how to call this function, see Arrays.CreateSequentialArray.
'
' PARAMETERS
' 1. StartNumber - First number in the span
' 2. EndNumber - Last number in the span
' 3. TheStep (optional) - To create a sequence using a sequential step different from 1
'
' RETURNED VALUE
' An object represing the desired span of indices
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
