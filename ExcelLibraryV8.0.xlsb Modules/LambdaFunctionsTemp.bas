Attribute VB_Name = "LambdaFunctionsTemp"
Public Function Lambda7(ArrayToSplice) As Variant
Dim x1 As Variant
Dim x2 As Variant
Let x1 = Part(ArrayToSplice, 1)
Let x2 = Part(ArrayToSplice, 2)

Let Lambda7 = Run("'ExcelLibraryV8.0.xlsb'!Lambda6", x1, x2)
End Function
Public Function Lambda6(x, y) As Variant


Let Lambda6 = x + y
End Function

Public Function Lambda4(x, y) As Variant


Let Lambda4 = x + y
End Function
Public Function Lambda3(ArrayToSplice) As Variant
Dim x1 As Variant
Dim x2 As Variant
Let x1 = Part(ArrayToSplice, 1)
Let x2 = Part(ArrayToSplice, 2)

Let Lambda3 = Run("'ExcelLibraryV8.0.xlsb'!Lambda2", x1, x2)
End Function
Public Function Lambda2(x, y) As Variant


Let Lambda2 = x + y
End Function
Public Function Lambda1(ArrayToSplice) As Variant
Dim x1 As Variant
Dim x2 As Variant
Let x1 = Part(ArrayToSplice, 1)
Let x2 = Part(ArrayToSplice, 2)

Let Lambda1 = Add(x1, x2)
End Function
Public Function Lambda0(ArrayToSplice) As Variant
Dim x1 As Variant
Dim x2 As Variant
Let x1 = Part(ArrayToSplice, 1)
Let x2 = Part(ArrayToSplice, 2)

Let Lambda0 = Add(x1, x2)
End Function
