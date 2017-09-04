Attribute VB_Name = "LambdaFunctionsTemp"
Public Function Lambda2(ArrayToSplice) As Variant
Dim x1 As Variant
Dim x2 As Variant
Let x1 = Part(ArrayToSplice, 1)
Let x2 = Part(ArrayToSplice, 2)

Let Lambda2 = Run("Add", x1, x2)
End Function
Public Function Lambda1(ArrayToSplice) As Variant
Dim x1 As Variant
Dim x2 As Variant
Dim x3 As Variant
Dim x4 As Variant
Let x1 = Part(ArrayToSplice, 1)
Let x2 = Part(ArrayToSplice, 2)
Let x3 = Part(ArrayToSplice, 3)
Let x4 = Part(ArrayToSplice, 4)

Let Lambda1 = Run("'ExcelLibraryV9.0.xlsb'!LambdaFunctionsTemp.Lambda0", x1, x2, x3, x4)
End Function
Public Function Lambda0(x1, x2, x3, x4) As Variant


Let Lambda0 = x1 + x2 + x3 + x4
End Function
