Attribute VB_Name = "Forms"
Option Explicit
Option Base 1

' This sub centers a shape in both dimensions relative to a given range (provided the dimensions work out)
Public Sub CenterShapeInRange(AShape As Shape, aRange As Range)
    Dim TheShapeLeftPos As Integer
    Dim TheShapeLeftPosOffset As Integer
    Dim TheShapeTopPos As Integer
    Dim TheShapeWidth As Integer
    Dim TheShapeHeight As Integer
    Dim TheShape As Shape
    Dim var As Variant
    
    Let AShape.Left = aRange.Left + Application.Max((aRange.Width - AShape.Width) / 2, 0)
    Let AShape.Top = aRange.Top + Application.Max((aRange.Height - AShape.Height) / 2, 0)
End Sub

Public Sub FormManipulationExample()
Attribute FormManipulationExample.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim theHeight As Double
    Dim theWidth As Double
    Dim i As Integer
    Dim obj() As OLEObject
    Dim N As Integer
    
    Dim var As Variant
    
    For Each var In TempComputation.Shapes
        Call var.Delete
    Next
    
    Let N = 10
    
    Let theWidth = TempComputation.Range("A1").Width
    Let theHeight = TempComputation.Range("A1").Height

    ReDim obj(1 To N)
    For i = 1 To N
        Set obj(i) = TempComputation.OLEObjects.Add(ClassType:="Forms.TextBox.1", Link:=False, DisplayAsIcon:=False, _
                                                    Left:=50, Top:=i * 50, Width:=theWidth, Height:=theHeight)
        
        Let obj(i).Name = "Pablo" & i
    Next i
End Sub

