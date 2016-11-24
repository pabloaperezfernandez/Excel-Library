Attribute VB_Name = "TypeConversions"
Option Base 1
Option Explicit

' Casts to strings a 1D array satisfying Predicates.AtomicArrayQ() or 2D array satisfying
' Predicates.AtomicTableQ() to strings.  If there is a problem with the parameter, the function
' returns an uninitialized array (e.g. an array returning False from Predicates.DimensionedQ())
'
' This function does not check if the elements of AnArray can be casted to the desired type.
Public Function ToStrings(AnArray As Variant) As String()
    Dim ResultArray() As String
    Dim RowOffset As Long
    Dim columnOffset As Long
    Dim r As Long
    Dim c As Long
    
    If EmptyArrayQ(AnArray) Then
        Let ToStrings = ResultArray
        Exit Function
    End If

    ' DimensionQ returns false if AnArray is not an array or it is not initialized
    If Not DimensionedQ(AnArray) Then
        Let ToStrings = ResultArray
        Exit Function
    End If
    
    ' Handle the case of a true 1D array
    If AtomicArrayQ(AnArray) Then
        ReDim ResultArray(1 To Length(AnArray))
        
        Let columnOffset = LBound(AnArray)
        For r = 1 To Length(AnArray)
            Let ResultArray(r) = CStr(AnArray(r - 1 + columnOffset))
        Next r
    End If

    ' Handle the case of a true 2D array
    If AtomicTableQ(AnArray) Then
        ReDim ResultArray(1 To NumberOfRows(AnArray), 1 To NumberOfColumns(AnArray))

        Let RowOffset = LBound(AnArray, 1)
        Let columnOffset = LBound(AnArray, 2)
        For r = 1 To NumberOfRows(AnArray)
            For c = 1 To NumberOfColumns(AnArray)
                Let ResultArray(r, c) = CStr(AnArray(r - 1 + RowOffset, c - 1 + columnOffset))
            Next c
        Next r
    End If
    
    Let ToStrings = ResultArray
End Function

' Casts to integers a 1D array satisfying Predicates.AtomicArrayQ() or 2D array satisfying
' Predicates.AtomicTableQ() a integers. If there is a problem with the parameter, the function
' returns an uninitialized array (e.g. an array returning False from Predicates.DimensionedQ())
'
' This function does not check if the elements of AnArray can be casted to the desired type.
Public Function ToIntegers(AnArray As Variant) As Integer()
    Dim ResultArray() As Integer
    Dim RowOffset As Long
    Dim columnOffset As Long
    Dim r As Long
    Dim c As Long
    
    If EmptyArrayQ(AnArray) Then
        Let ToIntegers = ResultArray
        Exit Function
    End If

    ' DimensionQ returns false if AnArray is not an array or it is not initialized
    If Not DimensionedQ(AnArray) Then
        Let ToIntegers = ResultArray
        Exit Function
    End If
    
    ' Handle the case of a true 1D array
    If AtomicArrayQ(AnArray) Then
        ReDim ResultArray(1 To Length(AnArray))
        
        Let columnOffset = LBound(AnArray)
        For r = 1 To Length(AnArray)
            Let ResultArray(r) = CInt(AnArray(r - 1 + columnOffset))
        Next r
    End If

    ' Handle the case of a true 2D array
    If AtomicTableQ(AnArray) Then
        ReDim ResultArray(1 To NumberOfRows(AnArray), 1 To NumberOfColumns(AnArray))

        Let RowOffset = LBound(AnArray, 1)
        Let columnOffset = LBound(AnArray, 2)
        For r = 1 To NumberOfRows(AnArray)
            For c = 1 To NumberOfColumns(AnArray)
                Let ResultArray(r, c) = CInt(AnArray(r - 1 + RowOffset, c - 1 + columnOffset))
            Next c
        Next r
    End If
    
    Let ToIntegers = ResultArray
End Function

' Casts to longs a 1D array satisfying Predicates.AtomicArrayQ() or 2D array satisfying
' Predicates.AtomicTableQ() a integers. If there is a problem with the parameter, the function
' returns an uninitialized array (e.g. an array returning False from Predicates.DimensionedQ())
'
' This function does not check if the elements of AnArray can be casted to the desired type.
Public Function ToLongs(AnArray As Variant) As Long()
    Dim ResultArray() As Long
    Dim RowOffset As Long
    Dim columnOffset As Long
    Dim r As Long
    Dim c As Long
    
    If EmptyArrayQ(AnArray) Then
        Let ToLongs = ResultArray
        Exit Function
    End If

    ' DimensionQ returns false if AnArray is not an array or it is not initialized
    If Not DimensionedQ(AnArray) Then
        Let ToLongs = ResultArray
        Exit Function
    End If
    
    ' Handle the case of a true 1D array
    If AtomicArrayQ(AnArray) Then
        ReDim ResultArray(1 To Length(AnArray))
        
        Let columnOffset = LBound(AnArray)
        For r = 1 To Length(AnArray)
            Let ResultArray(r) = CLng(AnArray(r - 1 + columnOffset))
        Next r
    End If

    ' Handle the case of a true 2D array
    If AtomicTableQ(AnArray) Then
        ReDim ResultArray(1 To NumberOfRows(AnArray), 1 To NumberOfColumns(AnArray))

        Let RowOffset = LBound(AnArray, 1)
        Let columnOffset = LBound(AnArray, 2)
        For r = 1 To NumberOfRows(AnArray)
            For c = 1 To NumberOfColumns(AnArray)
                Let ResultArray(r, c) = CLng(AnArray(r - 1 + RowOffset, c - 1 + columnOffset))
            Next c
        Next r
    End If
    
    Let ToLongs = ResultArray
End Function

' Casts to doubles a 1D array satisfying Predicates.AtomicArrayQ() or 2D array satisfying
' Predicates.AtomicTableQ() a integers. If there is a problem with the parameter, the function
' returns an uninitialized array (e.g. an array returning False from Predicates.DimensionedQ())
'
' This function does not check if the elements of AnArray can be casted to the desired type.
Public Function ToDoubles(AnArray As Variant) As Double()
    Dim ResultArray() As Double
    Dim RowOffset As Long
    Dim columnOffset As Long
    Dim r As Long
    Dim c As Long
    
    If EmptyArrayQ(AnArray) Then
        Let ToDoubles = ResultArray
        Exit Function
    End If

    ' DimensionQ returns false if AnArray is not an array or it is not initialized
    If Not DimensionedQ(AnArray) Then
        Let ToDoubles = ResultArray
        Exit Function
    End If
    
    ' Handle the case of a true 1D array
    If AtomicArrayQ(AnArray) Then
        ReDim ResultArray(1 To Length(AnArray))
        
        Let columnOffset = LBound(AnArray)
        For r = 1 To Length(AnArray)
            Let ResultArray(r) = CDbl(AnArray(r - 1 + columnOffset))
        Next r
    End If

    ' Handle the case of a true 2D array
    If AtomicTableQ(AnArray) Then
        ReDim ResultArray(1 To NumberOfRows(AnArray), 1 To NumberOfColumns(AnArray))

        Let RowOffset = LBound(AnArray, 1)
        Let columnOffset = LBound(AnArray, 2)
        For r = 1 To NumberOfRows(AnArray)
            For c = 1 To NumberOfColumns(AnArray)
                Let ResultArray(r, c) = CDbl(AnArray(r - 1 + RowOffset, c - 1 + columnOffset))
            Next c
        Next r
    End If
    
    Let ToDoubles = ResultArray
End Function

' Casts to booleans a 1D array satisfying Predicates.AtomicArrayQ() or 2D array satisfying
' Predicates.AtomicTableQ() a integers. If there is a problem with the parameter, the function
' returns an uninitialized array (e.g. an array returning False from Predicates.DimensionedQ())
'
' This function does not check if the elements of AnArray can be casted to the desired type.
Public Function ToBooleans(AnArray As Variant) As Boolean()
    Dim ResultArray() As Boolean
    Dim RowOffset As Long
    Dim columnOffset As Long
    Dim r As Long
    Dim c As Long
    
    If EmptyArrayQ(AnArray) Then
        Let ToBooleans = ResultArray
        Exit Function
    End If

    ' DimensionQ returns false if AnArray is not an array or it is not initialized
    If Not DimensionedQ(AnArray) Then
        Let ToBooleans = ResultArray
        Exit Function
    End If
    
    ' Handle the case of a true 1D array
    If AtomicArrayQ(AnArray) Then
        ReDim ResultArray(1 To Length(AnArray))
        
        Let columnOffset = LBound(AnArray)
        For r = 1 To Length(AnArray)
            Let ResultArray(r) = CBool(AnArray(r - 1 + columnOffset))
        Next r
    End If

    ' Handle the case of a true 2D array
    If AtomicTableQ(AnArray) Then
        ReDim ResultArray(1 To NumberOfRows(AnArray), 1 To NumberOfColumns(AnArray))

        Let RowOffset = LBound(AnArray, 1)
        Let columnOffset = LBound(AnArray, 2)
        For r = 1 To NumberOfRows(AnArray)
            For c = 1 To NumberOfColumns(AnArray)
                Let ResultArray(r, c) = CBool(AnArray(r - 1 + RowOffset, c - 1 + columnOffset))
            Next c
        Next r
    End If
    
    Let ToBooleans = ResultArray
End Function

' Casts to worksheets a 1D array satisfying Predicates.AtomicArrayQ() or 2D array satisfying
' Predicates.AtomicTableQ() a integers. If there is a problem with the parameter, the function
' returns an uninitialized array (e.g. an array returning False from Predicates.DimensionedQ())
'
' This function does not check if the elements of AnArray can be casted to the desired type.
Public Function ToWorksheets(AnArray As Variant) As Worksheet()
    Dim ResultArray() As Worksheet
    Dim RowOffset As Long
    Dim columnOffset As Long
    Dim r As Long
    Dim c As Long

    ' DimensionQ returns false if AnArray is not an array or it is not initialized
    If Not DimensionedQ(AnArray) Then
        Let ToWorksheets = ResultArray
        Exit Function
    End If

    If EmptyArrayQ(AnArray) Then
        Let ToWorksheets = ResultArray
        Exit Function
    End If
    
    ' Handle the case of a true 1D array
    If AtomicArrayQ(AnArray) Then
        ReDim ResultArray(1 To Length(AnArray))
        
        Let columnOffset = LBound(AnArray)
        For r = 1 To Length(AnArray)
            Set ResultArray(r) = AnArray(r - 1 + columnOffset)
        Next r
    End If

    ' Handle the case of a true 2D array
    If AtomicTableQ(AnArray) Then
        ReDim ResultArray(1 To NumberOfRows(AnArray), 1 To NumberOfColumns(AnArray))

        Let RowOffset = LBound(AnArray, 1)
        Let columnOffset = LBound(AnArray, 2)
        For r = 1 To NumberOfRows(AnArray)
            For c = 1 To NumberOfColumns(AnArray)
                Set ResultArray(r, c) = AnArray(r - 1 + RowOffset, c - 1 + columnOffset)
            Next c
        Next r
    End If
    
    Let ToWorksheets = ResultArray
End Function

' Casts to workbooks 1D array satisfying Predicates.AtomicArrayQ() or 2D array satisfying
' Predicates.AtomicTableQ() a integers. If there is a problem with the parameter, the function
' returns an uninitialized array (e.g. an array returning False from Predicates.DimensionedQ())
'
' This function does not check if the elements of AnArray can be casted to the desired type.
Public Function ToWorkbooks(AnArray As Variant) As Workbook()
    Dim ResultArray() As Workbook
    Dim RowOffset As Long
    Dim columnOffset As Long
    Dim r As Long
    Dim c As Long
    
    If EmptyArrayQ(AnArray) Then
        Let ToWorkbooks = ResultArray
        Exit Function
    End If

    ' DimensionQ returns false if AnArray is not an array or it is not initialized
    If Not DimensionedQ(AnArray) Then
        Let ToWorkbooks = ResultArray
        Exit Function
    End If
    
    ' Handle the case of a true 1D array
    If AtomicArrayQ(AnArray) Then
        ReDim ResultArray(1 To Length(AnArray))
        
        Let columnOffset = LBound(AnArray)
        For r = 1 To Length(AnArray)
            Let ResultArray(r) = AnArray(r - 1 + columnOffset)
        Next r
    End If

    ' Handle the case of a true 2D array
    If AtomicTableQ(AnArray) Then
        ReDim ResultArray(1 To NumberOfRows(AnArray), 1 To NumberOfColumns(AnArray))

        Let RowOffset = LBound(AnArray, 1)
        Let columnOffset = LBound(AnArray, 2)
        For r = 1 To NumberOfRows(AnArray)
            For c = 1 To NumberOfColumns(AnArray)
                Let ResultArray(r, c) = AnArray(r - 1 + RowOffset, c - 1 + columnOffset)
            Next c
        Next r
    End If
    
    Let ToWorkbooks = ResultArray
End Function

' Returns a 1D array with the same length as the input but with all its entries converted to strings.
' The array of strings is returned indexed from 1 to N.
Public Function Cast1DArrayToStrings(TheArray As Variant) As String()
    Dim i As Long
    Dim StringArray() As String
    
    ' Exit with EmptyArray() if arg is not an array
    If Not IsArray(TheArray) Then
        Let Cast1DArrayToStrings = StringArray
        Exit Function
    End If
    
    ' Exit with empty array if TheArray is empty
    If EmptyArrayQ(TheArray) Then
        Let Cast1DArrayToStrings = StringArray
        Exit Function
    End If
    
    If IsArray(TheArray) Then
        ReDim StringArray(1 To Length(TheArray))
    
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

    ' Exit with EmptyArray() if arg is not an array
    If Not IsArray(TheArray) Then
        Let Cast1DArrayToIntegers = EmptyArray()
        Exit Function
    End If

    ' Exit with empty array if TheArray is empty
    If EmptyArrayQ(TheArray) Then
        Let Cast1DArrayToIntegers = EmptyArray()
        Exit Function
    End If
    
    If IsArray(TheArray) Then
        ReDim ResultsArray(1 To Length(TheArray))
    
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
    
    ' Exit with EmptyArray() if arg is not an array
    If Not IsArray(TheArray) Then
        Let Cast1DArrayToLongs = EmptyArray()
        Exit Function
    End If
    
    ' Exit with empty array if TheArray is empty
    If EmptyArrayQ(TheArray) Then
        Let Cast1DArrayToLongs = EmptyArray()
        Exit Function
    End If
    
    If IsArray(TheArray) Then
        ReDim ResultsArray(1 To Length(TheArray))
    
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

    ' Exit with EmptyArray() if arg is not an array
    If Not IsArray(TheArray) Then
        Let Cast1DArrayToDoubles = EmptyArray()
        Exit Function
    End If

    ' Exit with empty array if TheArray is empty
    If EmptyArrayQ(TheArray) Then
        Let Cast1DArrayToLongs = EmptyArray()
        Exit Function
    End If
    
    If IsArray(TheArray) Then
        ReDim ResultsArray(1 To Length(TheArray))
    
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

    ' Exit with EmptyArray() if arg is not an array
    If Not IsArray(TheArray) Then
        Let Cast1DArrayToWorksheets = EmptyArray()
        Exit Function
    End If

    ' Exit with empty array if TheArray is empty
    If EmptyArrayQ(VariantWorksheetsArray) Then
        Let Cast1DArrayToWorksheets = EmptyArray()
        Exit Function
    End If

    If IsArray(VariantWorksheetsArray) Then
        ReDim MyArray(1 To Length(VariantWorksheetsArray))
        
        For i = 1 To Length(VariantWorksheetsArray)
            Set MyArray(i) = VariantWorksheetsArray(IIf(LBound(VariantWorksheetsArray) = 0, i - 1, i))
        Next i
        
        Let Cast1DArrayToWorksheets = MyArray
    Else
        Let Cast1DArrayToWorksheets = Array(VariantWorksheetsArray)
    End If
End Function

' This function casts atomic arguments and 1D arrays into the desired type.
' At the moment, it handles, integers, longs, doubles, strings, and Booleans
' This function can be used convert a variant type to the desired type, but
' it does not work if the conversion is made on a function call.  Type checking
' in VBA is done at the definition time and not at run time.
'
' This function is deprecated in favor of ToStrings(), etc.  It is kept here for
' backward compatibility reasons.
Public Function Cast(arg As Variant, TheDataType As XlParameterDataType) As Variant
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

    If IsArray(arg) Then
        Select Case TheDataType
            Case xlParamTypeInteger
                ReDim IntegerArray(LBound(arg) To UBound(arg))
                
                For c = LBound(arg) To UBound(arg)
                    Let IntegerArray(c) = CInt(arg(c))
                Next
                
                Let Cast = IntegerArray
            Case xlParamTypeLongVarBinary
                ReDim LongArray(LBound(arg) To UBound(arg))
                
                For c = LBound(arg) To UBound(arg)
                    Let IntegerArray(c) = CLng(arg(c))
                Next
                
                Let Cast = LongArray
            Case xlParamTypeDouble
                ReDim DoubleArray(LBound(arg) To UBound(arg))
                
                For c = LBound(arg) To UBound(arg)
                    Let DoubleArray(c) = CDbl(arg(c))
                Next
                
                Let Cast = DoubleArray
            Case xlParamTypeChar
                ReDim StringArray(LBound(arg) To UBound(arg))
                
                For c = LBound(arg) To UBound(arg)
                    Let StringArray(c) = CStr(arg(c))
                Next
                
                Let Cast = StringArray
            Case xlParamTypeBinary
                ReDim BooleanArray(LBound(arg) To UBound(arg))
                
                For c = LBound(arg) To UBound(arg)
                    Let BooleanArray(c) = CBool(arg(c))
                Next
                
                Let Cast = BooleanArray
        End Select
        
        Exit Function
    End If
    
    Select Case TheDataType
        Case xlParamTypeInteger
            Let Cast = CInt(arg)
        Case xlParamTypeLongVarBinary
            Let Cast = CLng(arg)
        Case xlParamTypeDouble
            Let Cast = CDbl(arg)
        Case xlParamTypeChar
            Let Cast = CStr(arg)
        Case xlParamTypeBinary
            Let Cast = CBool(arg)
    End Select
End Function


