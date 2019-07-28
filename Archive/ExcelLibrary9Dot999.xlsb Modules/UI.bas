Attribute VB_Name = "UI"
' The purpose of this module is to provide facilities for the
' creation of UIs. The functions herein automate processes we
' have found to repeat quite often.

Option Explicit
Option Base 1

' DESCRIPTION
' Deletes the shapes whose references are contained in the given array or collection.
'
' PARAMETERS
' 1. ShapeArray - An array of shapes
'
' RETURNED VALUE
' N/A
Public Sub DeleteShapes(ShapeArray As Variant)
    Dim VarShape As Variant
    
    ' Exit with Null if any item is not a shape
    For Each VarShape In ShapeArray
        If Not TypeName(VarShape) = "Shape" Then Exit Sub
    Next
    
    ' Delete all the shapes in the array
    For Each VarShape In ShapeArray
        Call VarShape.Delete
    Next
End Sub

' DESCRIPTION
' Deletes all shapes in the given worksheet
'
' PARAMETERS
' 1. Wsht - A reference to a worksheet object
'
' RETURNED VALUE
' N/A
Public Sub DeleteAllShapesInWorkSheet(wsht As Worksheet)
    Dim aShape As Shape

    For Each aShape In wsht.Shapes: Call aShape.Delete: Next
End Sub

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
Public Function CreateButtonsGrid(wsht As Worksheet, _
                                  UpperLeftCell As Range, _
                                  NumPerRow As Integer, _
                                  HorizontalSpacing As Long, _
                                  VerticalSpacing As Long, _
                                  ButtonWidth As Long, _
                                  ButtonHeight As Long, _
                                  ButtonCaptionsArray As Variant, _
                                  RoutineNamesOrLambdasArray As Variant) As Variant
    Dim aShape As Shape
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
            Set aShape = wsht.Shapes.AddFormControl(xlButtonControl, _
                                                    TheLeft, _
                                                    TheTop, _
                                                    ButtonWidth, _
                                                    ButtonHeight)
                                                                
            ' Assign the macro
            If StringQ(Part(RoutineNamesOrLambdasArray, (r - 1) * NumPerRow + c)) Then
                Let aShape.OnAction = Part(RoutineNamesOrLambdasArray, _
                                           (r - 1) * NumPerRow + c)
            Else
                Let aShape.OnAction = Part(RoutineNamesOrLambdasArray, _
                                           (r - 1) * NumPerRow + c).FunctionName
            End If
            
            ' Set the button's text
            Let wsht.Buttons(aShape.Name).Caption = Part(ButtonCaptionsArray, _
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
            Set aShape = wsht.Shapes.AddFormControl(xlButtonControl, _
                                                    TheLeft, _
                                                    TheTop, _
                                                    ButtonWidth, _
                                                    ButtonHeight)
            
            ' Assign the macro
            If StringQ(Part(RoutineNamesOrLambdasArray, _
                            Application.Quotient(n, NumPerRow) * NumPerRow + c)) Then
                Let aShape.OnAction = Part(RoutineNamesOrLambdasArray, _
                                           Application.Quotient(n, NumPerRow) * NumPerRow + c)
            Else
                Let aShape.OnAction = Part(RoutineNamesOrLambdasArray, _
                                           Application.Quotient(n, NumPerRow) * NumPerRow + c).FunctionName
            End If

            ' Set the button's text
            Let wsht.Buttons(aShape.Name).Caption = _
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
    Dim aShape As Shape
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
        Set aShape = Part(ButtonArray, c)
    
        If SpaceCentersEquallyQ Then
            Let NewLefts(c) = TargetRange.Left + HorizontalSpacing * c - CLng(aShape.Width / 2)
        Else

            If c = 1 Then
                Let NewLefts(1) = TargetRange.Left + HorizontalSpacing
            Else
                Let NewLefts(c) = NewLefts(c - 1) + Part(ButtonArray, c - 1).Width + HorizontalSpacing
            End If
        End If
        Let NewTops(c) = RangeVerticalCenter - aShape.Height / 2
        
        If NewLefts(c) < 1 Or NewTops(c) < 1 Then Exit Function
    Next
    
    ' Store the coordinates, sizes of the buttons, and coordinates of centers
    ' Compute new left and top coordinates and reposition buttons
    For c = 1 To Length(ButtonArray)
        Set aShape = Part(ButtonArray, c)
        
        Let aShape.Left = NewLefts(c)
        Let aShape.Top = NewTops(c)
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
' Returns the list chosen using a file chooser dialogue
'
' PARAMETERS
' 1. TextForDialogueWindow - Text to use as the dialogue's title
'
' RETURNED VALUE
' The chosen file.  An empty string is no file chosen (e.g. CANCEL chosen)
Public Function PickFileUsingFileChooser(TextForDialogueWindow As String) As String
        With Application.FileDialog(msoFileDialogFolderPicker)
            Let .Title = TextForDialogueWindow
            Let .InitialFileName = ThisWorkbook.Path
            Let .AllowMultiSelect = False
            If .Show = 0 Then Exit Function
            Let PickFileUsingFileChooser = .SelectedItems(1)
        End With
End Function

' ***HERE
' DESCRIPTION
' Creates a horizontal label/form control pair useful when creating forms.
' Returns Null in case of an error and True when successful.
'
' This function does not type checking to verify that the elements of the shape array are
' of type Shape.
'
' PARAMETERS
' 1. TheShapes - Array of shapes (form controls)
' 2. GroupBoxTitle - A string representing the title of the group box
'
' RETURNED VALUE
' Returns True on the sucessful creation and Null when an error encountered.
Public Function SurroundControlsWithGroupBox(TheShapes() As Shape, _
                                             GroupBoxTitle As String) As Variant
    Dim TopMostPos As Long
    Dim LeftMostPos As Long
    Dim BottomMostPos As Long
    Dim RightMostPos As Long
    Dim aShape As Variant
    Dim GroupShape As Variant
    Dim ShapeWsht As Worksheet
    
    Let SurroundControlsWithGroupBox = Null
    
    ' Exit with null if all the shapes are not in the same worksheet
    Set ShapeWsht = First(TheShapes).TopLeftCell.Worksheet
    For Each aShape In TheShapes
        If Not aShape.TopLeftCell.Worksheet Is ShapeWsht Then Exit Function
    Next
    
    ' Initialize the top, left, bottom, and right positions
    Let TopMostPos = First(TheShapes).Top
    Let LeftMostPos = First(TheShapes).Left
    Let BottomMostPos = First(TheShapes).Top + First(TheShapes).Height
    Let RightMostPos = First(TheShapes).Left + First(TheShapes).Width
    
    ' Loop through all the shapes to determine dimensions of group box
    For Each aShape In TheShapes
        If aShape.Top < TopMostPos Then Let TopMostPos = aShape.Top
        If aShape.Left < LeftMostPos Then Let LeftMostPos = aShape.Left
        If RightMostPos < aShape.Width + aShape.Left Then
            Let RightMostPos = aShape.Left + aShape.Width
        End If
        If BottomMostPos < aShape.Top + aShape.Height Then
            Let BottomMostPos = aShape.Top + aShape.Height
        End If
    Next
    
    ' Create a group box around the group of shapes
    Set GroupShape = ShapeWsht.GroupBoxes.Add(LeftMostPos - 20, _
                                              TopMostPos - 20, _
                                              RightMostPos - LeftMostPos + 40, _
                                              BottomMostPos - TopMostPos + 40)
    Let GroupShape.Characters.Text = GroupBoxTitle
    
    ' Exit with success code
    Set SurroundControlsWithGroupBox = GroupShape
End Function

' DESCRIPTION
' Creates a validation dropdown control at the given cell.  The allowed values
' can come from a 1xN or Nx1 array, a 1xN or Nx1 range of cells, or a listcolumn.
'
' PARAMETERS
' 1. ACellRange - A reference to a single cell to hold the dropdown control
' 2. CellLabel - String to insert to the left of ACellRange
' 3. ListColumnValue1DArrayOr1DRange
'    - Initialized 1D range of cells (1 column or 1 row)
'    - Initialized 1D array
'    - Initialized listcolumn
'
' RETURNED VALUE
' Null on error and True on success
'
Public Function CreateValidatedDropDownCell(ACellRange As Range, _
                                            CellLabel As String, _
                                            ListColumn1DArrayOr1DRange As Variant) As Variant
    Dim ListColumnName As String
    Dim RangeAddress As String
    
    ' Set default return value in case of error
    Let CreateValidatedDropDownCell = Null
    
    ' Exit with error code if Value1DArrayOr1DRange is neither a 1D array nor a range reference
    If Not (RowArrayQ(ListColumn1DArrayOr1DRange) Or _
            RangeQ(ListColumn1DArrayOr1DRange) Or _
            ListColumnQ(ListColumn1DArrayOr1DRange) Or _
            ListRowQ(ListColumn1DArrayOr1DRange)) Then
        Exit Function
    End If
    
    ' If ListColumn1DArrayOr1DRange is a range reference, exit if not 1D
    If RangeQ(ListColumn1DArrayOr1DRange) Then
        If ListColumn1DArrayOr1DRange.Rows.Count > 1 And ListColumn1DArrayOr1DRange.Columns.Count > 1 Then
            Exit Function
        End If
    End If
    
    ' Exit if the ACellRange points to a cell in column 1
    If ACellRange.Column = 1 Then Exit Function
    
    ' Drop the label
    Let ACellRange(1, 1).Offset(0, -1).Value2 = CellLabel
    
    ' Handle the case of a listcolumn
    If ListColumnQ(ListColumn1DArrayOr1DRange) Then
        Let ListColumnName = ListColumn1DArrayOr1DRange.Parent.Name & "[" & ListColumn1DArrayOr1DRange.Name & "]"
    
        With ACellRange.Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, _
                 Formula1:="=Indirect(""" & ListColumnName & """)"
            .IgnoreBlank = True
            .InCellDropdown = True
            .InputTitle = ""
            .ErrorTitle = ""
            .InputMessage = ""
            .ErrorMessage = ""
            .ShowInput = True
            .ShowError = True
        End With
    ElseIf RowArrayQ(ListColumn1DArrayOr1DRange) Then
        With ACellRange.Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, _
                 Formula1:=Join(ListColumn1DArrayOr1DRange, ",")
            .IgnoreBlank = True
            .InCellDropdown = True
            .InputTitle = ""
            .ErrorTitle = ""
            .InputMessage = ""
            .ErrorMessage = ""
            .ShowInput = True
            .ShowError = True
        End With
    Else
        Let RangeAddress = "'" & ListColumn1DArrayOr1DRange.Worksheet.Name & "'!" & ListColumn1DArrayOr1DRange.Address
    
        With ACellRange.Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, _
                 Formula1:="=" & RangeAddress
            .IgnoreBlank = True
            .InCellDropdown = True
            .InputTitle = ""
            .ErrorTitle = ""
            .InputMessage = ""
            .ErrorMessage = ""
            .ShowInput = True
            .ShowError = True
        End With
    End If
    
    ' Exit with success code
    Let CreateValidatedDropDownCell = True
End Function
