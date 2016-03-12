Attribute VB_Name = "TypeConversions"
Option Base 1
Option Explicit

' Returns a 1D array with the same length as the input but with all its entries converted to strings.
' The array of strings is returned indexed from 1 to N.
Public Function Cast1DArrayToStrings(TheArray As Variant) As String()
    Dim i As Long
    Dim StringArray() As String
    
    ' Exit with Array() if arg is not an array
    If Not IsArray(TheArray) Then
        Let Cast1DArrayToStrings = Array()
        Exit Function
    End If
    
    ' Exit with empty array if TheArray is empty
    If EmptyArrayQ(TheArray) Then
        Let Cast1DArrayToStrings = Array()
        Exit Function
    End If
    
    If IsArray(TheArray) Then
        ReDim StringArray(1 To GetArrayLength(TheArray))
    
        For i = LBound(TheArray) To UBound(TheArray)
            Let StringArray(i + (1 - LBound(TheArray))) = CStr(TheArray(i))
        Next i
    
        Let Cast1DArrayToStrings = StringArray
    Else
        Let Cast1DArrayToStrings = Array(CStr(TheArray))
    End If
End Function

' Returns a 1D array with the same length as the input but with all its entries converted to strings.
' The array of strings is returned indexed from 1 to N.
Public Function Cast1DArrayToIntegers(TheArray As Variant) As Integer()
    Dim i As Long
    Dim ResultsArray() As Integer

    ' Exit with Array() if arg is not an array
    If Not IsArray(TheArray) Then
        Let Cast1DArrayToIntegers = Array()
        Exit Function
    End If

    ' Exit with empty array if TheArray is empty
    If EmptyArrayQ(TheArray) Then
        Let Cast1DArrayToIntegers = Array()
        Exit Function
    End If
    
    If IsArray(TheArray) Then
        ReDim ResultsArray(1 To GetArrayLength(TheArray))
    
        For i = LBound(TheArray) To UBound(TheArray)
            Let ResultsArray(i + (1 - LBound(TheArray))) = CStr(TheArray(i))
        Next i
    
        Let Cast1DArrayToIntegers = ResultsArray
    Else
        Let Cast1DArrayToIntegers = Array(CInt(TheArray))
    End If
End Function

' Returns a 1D array with the same length as the input but with all its entries converted to strings.
' The array of strings is returned indexed from 1 to N.
Public Function Cast1DArrayToLongs(TheArray As Variant) As Long()
    Dim i As Long
    Dim ResultsArray() As Long
    
    ' Exit with Array() if arg is not an array
    If Not IsArray(TheArray) Then
        Let Cast1DArrayToLongs = Array()
        Exit Function
    End If
    
    ' Exit with empty array if TheArray is empty
    If EmptyArrayQ(TheArray) Then
        Let Cast1DArrayToLongs = Array()
        Exit Function
    End If
    
    If IsArray(TheArray) Then
        ReDim ResultsArray(1 To GetArrayLength(TheArray))
    
        For i = LBound(TheArray) To UBound(TheArray)
            Let ResultsArray(i + (1 - LBound(TheArray))) = CStr(TheArray(i))
        Next i
    
        Let Cast1DArrayToLongs = ResultsArray
    Else
        Let Cast1DArrayToLongs = Array(CLng(TheArray))
    End If
End Function

' Returns a 1D array with the same length as the input but with all its entries converted to strings.
' The array of strings is returned indexed from 1 to N.
Public Function Cast1DArrayToDoubles(TheArray As Variant) As Double()
    Dim i As Long
    Dim ResultsArray() As Double

    ' Exit with Array() if arg is not an array
    If Not IsArray(TheArray) Then
        Let Cast1DArrayToDoubles = Array()
        Exit Function
    End If

    ' Exit with empty array if TheArray is empty
    If EmptyArrayQ(TheArray) Then
        Let Cast1DArrayToLongs = Array()
        Exit Function
    End If
    
    If IsArray(TheArray) Then
        ReDim ResultsArray(1 To GetArrayLength(TheArray))
    
        For i = LBound(TheArray) To UBound(TheArray)
            Let ResultsArray(i + (1 - LBound(TheArray))) = CStr(TheArray(i))
        Next i
    
        Let Cast1DArrayToLongs = ResultsArray
    Else
        Let Cast1DArrayToLongs = Array(CDbl(TheArray))
    End If
End Function

' Returns a 1D array with the same length as the input but with all its entries converted to worksheet object references.
' The array of worksheet references is returned indexed from 1 to N.
Public Function Cast1DArrayToWorksheets(VariantWorksheetsArray As Variant) As Worksheet()
    Dim MyArray() As Worksheet
    Dim i As Long

    ' Exit with Array() if arg is not an array
    If Not IsArray(TheArray) Then
        Let Cast1DArrayToWorksheets = Array()
        Exit Function
    End If

    ' Exit with empty array if TheArray is empty
    If EmptyArrayQ(VariantWorksheetsArray) Then
        Let Cast1DArrayToWorksheets = Array()
        Exit Function
    End If

    If IsArray(VariantWorksheetsArray) Then
        ReDim MyArray(1 To GetArrayLength(VariantWorksheetsArray))
        
        For i = 1 To GetArrayLength(VariantWorksheetsArray)
            Set MyArray(i) = VariantWorksheetsArray(IIf(LBound(VariantWorksheetsArray) = 0, i - 1, i))
        Next i
        
        Let Cast1DArrayToWorksheets = MyArray
    Else
        Let Cast1DArrayToWorksheets = Array(VariantWorksheetsArray)
    End If
End Function

' This function casts atomic arguments and 1D arrays into the desired type.
' At the moment, it handles, integers, longs, doubles, strings, and Booleans
Public Function Cast(Arg As Variant, TheDataType As XlParameterDataType) As Variant
    Dim IntegerArray() As Integer
    Dim LongArray() As Integer
    Dim DoubleArray() As Double
    Dim StringArray() As String
    Dim BooleanArray() As Boolean
    Dim c As Long
    
    ' Exit with Null if TheDataType is not one of the supported types
    If FreeQ(Array(xlParamTypeInteger, xlParamTypeLongVarBinary, _
                   xlParamTypeDouble, xlParamTypeChar, _
                   xlParamTypeBinary), _
               TheDataType) Then
        Let Cast = Null
        Exit Function
    End If

    If IsArray(Arg) Then
        Select Case TheDataType
            Case xlParamTypeInteger
                ReDim IntegerArray(LBound(Arg) To UBound(Arg))
                
                For c = LBound(Arg) To UBound(Arg)
                    Let IntegerArray(c) = CInt(Arg(c))
                Next
                
                Let Cast = IntegerArray
            Case xlParamTypeLongVarBinary
                ReDim LongArray(LBound(Arg) To UBound(Arg))
                
                For c = LBound(Arg) To UBound(Arg)
                    Let IntegerArray(c) = CLng(Arg(c))
                Next
                
                Let Cast = LongArray
            Case xlParamTypeDouble
                ReDim DoubleArray(LBound(Arg) To UBound(Arg))
                
                For c = LBound(Arg) To UBound(Arg)
                    Let DoubleArray(c) = CDbl(Arg(c))
                Next
                
                Let Cast = DoubleArray
            Case xlParamTypeChar
                ReDim StringArray(LBound(Arg) To UBound(Arg))
                
                For c = LBound(Arg) To UBound(Arg)
                    Let StringArray(c) = CStr(Arg(c))
                Next
                
                Let Cast = StringArray
            Case xlParamTypeBinary
                ReDim BooleanArray(LBound(Arg) To UBound(Arg))
                
                For c = LBound(Arg) To UBound(Arg)
                    Let BooleanArray(c) = CBool(Arg(c))
                Next
                
                Let Cast = BooleanArray
        End Select
        
        Exit Function
    End If
    
    Select Case TheDataType
        Case xlParamTypeInteger
            Let Cast = CInt(Arg)
        Case xlParamTypeLongVarBinary
            Let Cast = CLng(Arg)
        Case xlParamTypeDouble
            Let Cast = CDbl(Arg)
        Case xlParamTypeChar
            Let Cast = CStr(Arg)
        Case xlParamTypeBinary
            Let Cast = CBool(Arg)
    End Select
End Function
