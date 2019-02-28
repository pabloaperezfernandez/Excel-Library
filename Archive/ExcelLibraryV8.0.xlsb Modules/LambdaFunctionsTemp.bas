Attribute VB_Name = "LambdaFunctionsTemp"
Public Function Lambda8(ArrayToSplice) As Variant
Dim x1 As Variant
Dim x2 As Variant
Dim x3 As Variant
Let x1 = Part(ArrayToSplice, 1)
Let x2 = Part(ArrayToSplice, 2)
Let x3 = Part(ArrayToSplice, 3)

Let Lambda8 = Run("'ExcelLibraryV8.0.xlsb'!Lambda7", x1, x2, x3)
End Function
Public Function Lambda7(x, y, z) As Variant


Let Lambda7 = x + y + z
End Function
Public Function Lambda6(x) As Variant


Let Lambda6 = 2 * x
End Function
Public Function Lambda5(x) As Variant


Let Lambda5 = 2 * x
End Function
Public Function Lambda4(ArrayToSplice) As Variant
Dim x1 As Variant
Dim x2 As Variant
Dim x3 As Variant
Let x1 = Part(ArrayToSplice, 1)
Let x2 = Part(ArrayToSplice, 2)
Let x3 = Part(ArrayToSplice, 3)

Let Lambda4 = Run("'ExcelLibraryV8.0.xlsb'!Lambda3", x1, x2, x3)
End Function
Public Function Lambda3(x, y, z) As Variant


Let Lambda3 = x + y + z
End Function
