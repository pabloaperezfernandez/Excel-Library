VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Lambda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' This class is used to define the Lambda data type. Instances of this class are used as
' arguments to many functional programming facilities. See for instance the FunctionalProgramming
' module

Option Base 1
Option Explicit

' Fully-qualified name including workbook name (e.g. "'MyBook.xlsb'!MyModule.MyFunctionName")
' Does not have to be fully-qualified, but it can be.
Public FunctionName As String
