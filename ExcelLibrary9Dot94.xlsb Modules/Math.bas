Attribute VB_Name = "Math"
Option Explicit
Option Base 1

Public Function Floor(ANumber As Double) As Long
    Let Floor = Application.Floor_Precise(ANumber)
End Function

Public Function Ceiling(ANumber As Double) As Long
    Let Ceiling = Application.Ceiling_Precise(ANumber)
End Function
