VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Span"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Base 1
Option Explicit

Private IntervalValue As Variant

Property Get Value() As Variant
    Let Value = IntervalValue
End Property

Property Let Value(TheValue As Variant)
    Let IntervalValue = TheValue
End Property
