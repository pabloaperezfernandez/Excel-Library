Attribute VB_Name = "UI"
' The purpose of this module is to provide facilities for the
' creation of UIs. The functions herein automate processes we
' have found to repeat quite often.

Option Explicit
Option Base 1

' DESCRIPTION
' Deletes the shapes whose references are contained in the given array
'
' PARAMETERS
' 1. ShapeArray - An array of shapes
'
' RETURNED VALUE
' Null if an error is encountered.  True if deletion suceeds.
Public Function DeleteShapes(ShapeArray As Variant) As Variant
    Dim VarShape As Variant
    
    ' Set default value in case of error
    Let DeleteShapes = Null
    
    ' Exit with Null if ShapeArray not as expected
    If Not DimensionedNonEmptyArrayQ(ShapeArray) Then Exit Function
    
    ' Exit with Null if any item is not a shape
    For Each VarShape In ShapeArray
        If Not TypeName(VarShape) = "Shape" Then Exit Function
    Next
    
    ' Delete all the shapes in the array
    For Each VarShape In ShapeArray
        Call VarShape.Delete
    Next
    
    Let DeleteShapes = True
End Function

' DESCRIPTION
'
'
' PARAMETERS
' 1.
'
' RETURNED VALUE
'
' This uses Shape.Type which takes its value from enumeration MsoShapeType.
' Form controls has value msoFormControl
' It also uses Shape.FormControlType which takes its value from enumeration
' xlFormControl
Public Function LayoutButtonsOnGrid(Wsht As Worksheet, _
                                    UpperLeftCell As Range, _
                                    NumPerRow As Integer, _
                                    HorizontalSpacing As Long, _
                                    VerticalSpacing As Long, _
                                    ButtonWidth As Long, _
                                    ButtonHeight As Long, _
                                    ButtonCaptionsArray As Variant, _
                                    RoutineNamesOrLambdasArray As Variant) As Variant
    Dim AShape As Shape
    Dim r As Integer
    Dim c As Integer
    Dim TheLeft As Long
    Dim TheTop As Long
    Dim N As Integer
    
    Let N = Length(ButtonCaptionsArray)
    For r = 1 To Application.Quotient(N, NumPerRow)
        For c = 1 To NumPerRow
            Let TheLeft = UpperLeftCell.Left + (c - 1) * (ButtonWidth + HorizontalSpacing)
            Let TheTop = UpperLeftCell.Top + (r - 1) * (ButtonHeight + VerticalSpacing)
        
            Set AShape = Wsht.Shapes.AddFormControl(xlButtonControl, _
                                                    TheLeft, _
                                                    TheTop, _
                                                    ButtonWidth, _
                                                    ButtonHeight)
                                                    
            If StringQ(Part(RoutineNamesOrLambdasArray, (r - 1) * NumPerRow + c)) Then
                Let AShape.OnAction = Part(RoutineNamesOrLambdasArray, (r - 1) * NumPerRow + c)
            Else
                Let AShape.OnAction = Part(RoutineNamesOrLambdasArray, (r - 1) * NumPerRow + c).FunctionName
            End If
            
            Let AShape.AlternativeText = Part(ButtonCaptionsArray, (r - 1) * NumPerRow + c)
        Next
    Next
    
    If N Mod NumPerRow <> 0 Then
        Let TheTop = Application.Quotient(N, NumPerRow) * (ButtonHeight + VerticalSpacing)
        Let TheTop = UpperLeftCell.Top + TheTop
        
        For c = 1 To N Mod NumPerRow
            Let TheLeft = UpperLeftCell.Left + (c - 1) * (ButtonWidth + HorizontalSpacing)

            Set AShape = Wsht.Shapes.AddFormControl(xlButtonControl, _
                                                    TheLeft, _
                                                    TheTop, _
                                                    ButtonWidth, _
                                                    ButtonHeight)

            If StringQ(Part(RoutineNamesOrLambdasArray, Application.Quotient(N, NumPerRow) * NumPerRow + c)) Then
                Let AShape.OnAction = Part(RoutineNamesOrLambdasArray, _
                                           Application.Quotient(N, NumPerRow) * NumPerRow + c)
            Else
                Let AShape.OnAction = Part(RoutineNamesOrLambdasArray, _
                                           Application.Quotient(N, NumPerRow) * NumPerRow + c).FunctionName
            End If

            Let AShape.AlternativeText = Part(ButtonCaptionsArray, _
                                              Application.Quotient(N, NumPerRow) * NumPerRow + c)
        Next
    End If
End Function

' DESCRIPTION
'
'
' PARAMETERS
' 1.
'
' RETURNED VALUE
'
Public Function CreateValidatedDropDownCell(CellLabel As String, TheListObject As ListObject, _
                                            ListColumnHeader As String) As Variant
                                  
End Function

