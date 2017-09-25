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
' Creates and arranges a list of buttons (form controls) in a grid pattern
' with the given spacing and assigns them to macros.
'
' PARAMETERS
' 1. Wsht - Target worksheet
' 2. UpperLeftCell
'
' RETURNED VALUE
' Returns True on the sucessful create of the buttons grid. Returns Null when
' an error is encountered.
'
' This uses Shape.Type which takes its value from enumeration MsoShapeType.
' Form controls has value msoFormControl
' It also uses Shape.FormControlType which takes its value from enumeration
' xlFormControl
Public Function CreateButtonsGrid(Wsht As Worksheet, _
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
    Dim n As Integer
    
    ' Set default return value in case of error
    Let CreateButtonsGrid = Null
    
    ' Compute number of buttons to arrange
    Let n = Length(ButtonCaptionsArray)
    
    ' Insert as many full rows of buttons as needed
    For r = 1 To Application.Quotient(n, NumPerRow)
        ' Insert the buttons for this row
        For c = 1 To NumPerRow
            ' Compute the left and top positions of this button
            Let TheLeft = UpperLeftCell.Left + (c - 1) * (ButtonWidth + HorizontalSpacing)
            Let TheTop = UpperLeftCell.Top + (r - 1) * (ButtonHeight + VerticalSpacing)
        
            ' Create the button
            Set AShape = Wsht.Shapes.AddFormControl(xlButtonControl, _
                                                    TheLeft, _
                                                    TheTop, _
                                                    ButtonWidth, _
                                                    ButtonHeight)
                                                                
            ' Assign the macro
            If StringQ(Part(RoutineNamesOrLambdasArray, (r - 1) * NumPerRow + c)) Then
                Let AShape.OnAction = Part(RoutineNamesOrLambdasArray, _
                                           (r - 1) * NumPerRow + c)
            Else
                Let AShape.OnAction = Part(RoutineNamesOrLambdasArray, _
                                           (r - 1) * NumPerRow + c).FunctionName
            End If
            
            ' Set the button's text
            Let Wsht.Buttons(AShape.Name).Caption = Part(ButtonCaptionsArray, _
                                                         (r - 1) * NumPerRow + c)
        Next
    Next
    
    ' Arrange another partially-filled row if needed
    If n Mod NumPerRow <> 0 Then
        ' Compute the top position for this row's buttons
        Let TheTop = Application.Quotient(n, NumPerRow) * (ButtonHeight + VerticalSpacing)
        Let TheTop = UpperLeftCell.Top + TheTop
        
        ' Insert the buttons for this row
        For c = 1 To n Mod NumPerRow
            ' Compute the left position for this button
            Let TheLeft = UpperLeftCell.Left + (c - 1) * (ButtonWidth + HorizontalSpacing)

            ' Create the button
            Set AShape = Wsht.Shapes.AddFormControl(xlButtonControl, _
                                                    TheLeft, _
                                                    TheTop, _
                                                    ButtonWidth, _
                                                    ButtonHeight)
            
            ' Assign the macro
            If StringQ(Part(RoutineNamesOrLambdasArray, _
                            Application.Quotient(n, NumPerRow) * NumPerRow + c)) Then
                Let AShape.OnAction = Part(RoutineNamesOrLambdasArray, _
                                           Application.Quotient(n, NumPerRow) * NumPerRow + c)
            Else
                Let AShape.OnAction = Part(RoutineNamesOrLambdasArray, _
                                           Application.Quotient(n, NumPerRow) * NumPerRow + c).FunctionName
            End If

            ' Set the button's text
            Let Wsht.Buttons(AShape.Name).Caption = _
                Part(ButtonCaptionsArray, Application.Quotient(n, NumPerRow) * NumPerRow + c)
        Next
    End If
    
    ' Return True to indicate success
    Let CreateButtonsGrid = True
End Function

' DESCRIPTION
' Re-arranges a set of buttons to be equally distributed on a range.
' The buttons are well distributed horizontally in the given range.
' If the sum of the buttons' widths is less than the width of the
' target range, the excess range width is distributed evenly between
' the buttons and the edges of the target range. If the sums of the
' buttons' widths exceeds the width of the target range, we simply
' ensure the centers are equally spaced horizontally in the target
' range.

' The distribution
' is done so the centers of the buttons are equally spaced in the range.
' No assumption is made about the sizes of the buttons.
'
' PARAMETERS
' 1. ButtonArray - Array of buttons
' 2. TargetRange - Range over which to arrange the buttons
'
' RETURNED VALUE
' Returns True on the sucessful create of the buttons grid. Returns Null when
' an error is encountered.
Public Function DistributeButtonsHorizontally(ButtonArray As Variant, _
                                              TargetRange As Range) As Variant
    Dim HorizontalSpacing As Double
    Dim RangeVerticalCenter As Long
    Dim AShape As Shape
    Dim NewLefts() As Long
    Dim NewTops() As Long
    Dim c As Integer
    Dim ButtonsTotalWidth As Long
    Dim SpaceCentersEquallyQ As Boolean
    
    ' Set default return value in case of error
    Let DistributeButtonsHorizontally = Null
    
    ' Error Check: Exit with Null if ButtonArray is not as expected
    If Not DimensionedNonEmptyArrayQ(ButtonArray) Then Exit Function
    
    ' Error Check: Exit with Null if any elt in ButtonArray is not a shape
    If Not AllTrueQ(ButtonArray, _
                    MakeRoutineName(ThisWorkbook, "Predicates", "FormControlButtonQ")) Then
        Exit Function
    End If
    
    ' Error Check: Exit with Null if TargetRange is Nothing
    If TargetRange Is Nothing Then Exit Function
    
    ' Determine if we should equally space buttons' centers or equally distribute
    ' the target range's excess width
    Let ButtonsTotalWidth = Total(Map(Lambda("b", "", "b.Width"), ButtonArray))
    Let SpaceCentersEquallyQ = ButtonsTotalWidth > TargetRange.Width

    ' Error Check: Exit with Null if the left of top of any button would end up
    ' outside of the worksheet
    ReDim NewLefts(1 To Length(ButtonArray))
    ReDim NewTops(1 To Length(ButtonArray))
    
    ' Decide if the widths of the buttons exceeds the width of the target range
    If SpaceCentersEquallyQ Then
        Let HorizontalSpacing = TargetRange.Width / (Length(ButtonArray) + 1)
    Else
        Let HorizontalSpacing = (TargetRange.Width - ButtonsTotalWidth) / (Length(ButtonArray) + 1)
    End If

    Let RangeVerticalCenter = TargetRange.Top + CLng(TargetRange.Height / 2)
    For c = 1 To Length(ButtonArray)
        Set AShape = Part(ButtonArray, c)
    
        If SpaceCentersEquallyQ Then
            Let NewLefts(c) = TargetRange.Left + HorizontalSpacing * c - CLng(AShape.Width / 2)
        Else

            If c = 1 Then
                Let NewLefts(1) = TargetRange.Left + HorizontalSpacing
            Else
                Let NewLefts(c) = NewLefts(c - 1) + Part(ButtonArray, c - 1).Width + HorizontalSpacing
            End If
        End If
        Let NewTops(c) = RangeVerticalCenter - AShape.Height / 2
        
        If NewLefts(c) < 1 Or NewTops(c) < 1 Then Exit Function
    Next
    
    ' Store the coordinates, sizes of the buttons, and coordinates of centers
    ' Compute new left and top coordinates and reposition buttons
    For c = 1 To Length(ButtonArray)
        Set AShape = Part(ButtonArray, c)
        
        Let AShape.Left = NewLefts(c)
        Let AShape.Top = NewTops(c)
    Next

    ' Return True to indicate success
    Let DistributeButtonsHorizontally = True
End Function

' DESCRIPTION
' Resizes each of the form control buttons in the given array to
' the size.
'
' PARAMETERS
' 1. ButtonArray - Array of buttons
' 2. TheWidth - The desired width
' 3. TheHeight - The desired height
'
' RETURNED VALUE
' Returns True on the sucessful resize. Returns Null when an error
' is encountered.
Public Function EqualizeFormControlButtonSizes(ButtonArray As Variant, _
                                               TheWidth As Long, _
                                               TheHeight As Long) As Variant
    ' Set default return value in case of error
    Let EqualizeFormControlButtonSizes = Null
    
    ' Error Check: Exit with Null if ButtonArray is not as expected
    If Not DimensionedNonEmptyArrayQ(ButtonArray) Then Exit Function
    
    ' Error Check: Exit with Null if any elt in ButtonArray is not a shape
    If Not AllTrueQ(ButtonArray, _
                    MakeRoutineName(ThisWorkbook, "Predicates", "FormControlButtonQ")) Then
        Exit Function
    End If

    ' Error Check: Exit with Null if either TheWidth or TheHeight < 1
    If TheWidth < 1 Or TheHeight < 1 Then Exit Function
    
    ' Resize all the buttons
    Call Scan(Lambda("x", _
                     Array("Call x.ScaleWidth(" & TheWidth & "/x.Width, msoFalse)", _
                           "Call x.ScaleHeight(" & TheHeight & "/x.Height, msoFalse)"), _
                     "True"), _
              ButtonArray)

    ' Return True to indicate success
    Let EqualizeFormControlButtonSizes = True
End Function

' DESCRIPTION
' Creates a horizontal label/form control pair useful when creating forms.
' Returns Null in case of an error and True when successful
'
' PARAMETERS
' 1. TheLabel - Array of buttons
' 2. ControlType - The desired width
' 3. LabelCell - The desired height
' 4. TheWidth
' 5. TheHeight
'
' RETURNED VALUE
' Returns True on the sucessful creation and Null when an error encountered.
Public Function CreateLabelAndFormControlPair(TheLabel As String, _
                                              ControlType As XlFormControl, _
                                              LabelCell As Range, _
                                              TheWidth As Long, _
                                              TheHeight As Long) As Variant
    ' Set default return value in case of error
    Let CreateLabelAndFormControlPair = Null
    
    ' Return True to indicate success
    Let CreateLabelAndFormControlPair = True '***HERE
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

