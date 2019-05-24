Attribute VB_Name = "TypeConversions"
Option Base 1
Option Explicit

' Cast an object reference to a listobject. This function is used to convert variants
' referencing listobjects to a reference with type listobject
Public Function ToListObject(AnObjectRerence As Variant) As ListObject
    Set ToListObject = AnObjectRerence
End Function

' Casts to ListObjects a 1D array satisfying Predicates.AtomicArrayQ() or 2D array satisfying
' Predicates.AtomicTableQ().  Each of the elements must be an object referencing a listobject
' If there is a problem with the parameter, the function returns an uninitialized array
' (e.g. an array returning False from Predicates.DimensionedQ())
'
' This function does not check if the elements of AnArray can be casted to the desired type.
Public Function ToListObjects(AnArray As Variant) As ListObject()
    Dim ResultArray() As ListObject
    Dim RowOffset As Long
    Dim columnOffset As Long
    Dim r As Long
    Dim c As Long
    
    ' DimensionQ returns false if AnArray is not an array or it is not initialized
    If Not DimensionedQ(AnArray) Then
        Let ToListObjects = ResultArray
        Exit Function
    End If

    ' Exit with an empty array with the argument is an empty array
    If EmptyArrayQ(AnArray) Then
        Let ToListObjects = ResultArray
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
    
    Let ToListObjects = ResultArray
End Function

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
    
    ' DimensionQ returns false if AnArray is not an array or it is not initialized
    If Not DimensionedQ(AnArray) Then
        Let ToStrings = ResultArray
        Exit Function
    End If

    ' Exit with an empty array with the argument is an empty array
    If EmptyArrayQ(AnArray) Then
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

' DESCRIPTION
' Converts a shapes collection to an array. Returns Null if the collection
' is empty
'
' PARAMETERS
' 1. AShapesCollection - A shapes collection
'
' RETURNED VALUE
' An array holding the elements of the shapes collection
Public Function CastShapesToArray(AShapesCollection As Shapes) As Variant
    Dim ResultArray() As Variant
    Dim c As Long
    
    Let CastShapesToArray = Null

    If AShapesCollection.Count = 0 Then Exit Function
    
    ReDim ResultArray(1 To AShapesCollection.Count)
    For c = 1 To AShapesCollection.Count
        Set ResultArray(c) = AShapesCollection(c)
    Next
    
    Let CastShapesToArray = ResultArray
End Function

