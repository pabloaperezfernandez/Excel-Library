VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Lambda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' This class is used in conjunction with a number of array functions and routines
' to specify ranges of indices.  It works identically to Mathematica's Span.  It
' can be used as a parmeter for Arrays.Part() and other arrays slicing formulas.

Option Base 1
Option Explicit

' Fully-qualified name including workbook name (e.g. "'MyBook.xlsb'!MyFunctionName")
Public FunctionName As String
