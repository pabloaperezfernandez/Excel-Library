Attribute VB_Name = "LambdaFunctionsTemp"
Public Function Lambda14(ArrayToSplice) As Variant
Dim x1 As Variant
Dim x2 As Variant
Let x1 = Part(ArrayToSplice, 1)
Let x2 = Part(ArrayToSplice, 2)

Let Lambda14 = Run("'ExcelLibraryV9.0.xlsb'!Lambda13", x1, x2)
End Function
Public Function Lambda13(x, y) As Variant


Let Lambda13 = 2 * x * y
End Function
Public Function Lambda12(ArrayToSplice) As Variant
Dim x1 As Variant
Dim x2 As Variant
Let x1 = Part(ArrayToSplice, 1)
Let x2 = Part(ArrayToSplice, 2)

Let Lambda12 = Run("'ExcelLibraryV9.0.xlsb'!Lambda11", x1, x2)
End Function
Public Function Lambda11(x, y) As Variant


Let Lambda11 = 2 * x * y
End Function
