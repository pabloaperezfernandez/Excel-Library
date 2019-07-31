Attribute VB_Name = "Predicates"
Option Explicit
Option Base 1

Dim SampleLambda As New Lambda

' DESCRIPTION
' This function returns an array of strings with the names the predicates used to
' identify atomic values.
'
' PARAMETERS
'
' RETURNED VALUE
' array of strings with the names the predicates used to identify atomic values.
Public Function GetAtomicTypePredicateNames() As Variant
    Let GetAtomicTypePredicateNames = Array("BooleanQ", _
                                            "DateQ", _
                                            "DictionaryQ", _
                                            "ListObjectQ", _
                                            "StringQ", _
                                            "WorkbookQ", _
                                            "WorksheetQ", _
                                            "EmptyQ", _
                                            "ErrorQ", _
                                            "NullQ", _
                                            "NumberQ", _
                                            "PositiveQ", _
                                            "NegativeQ", _
                                            "NonPositiveQ", _
                                            "NonNegativeQ", _
                                            "NonWholeNumberQ", _
                                            "WholeNumberQ", _
                                            "PositiveWholeNumberQ", _
                                            "NegativeWholeNumberQ", _
                                            "NonPositiveWholeNumberQ", _
                                            "NonNegativeWholeNumberQ", _
                                            "NonZeroWholeNumberQ", _
                                            "TrueQ", _
                                            "FalseQ")
End Function

' DESCRIPTION
' This function returns an array of strings with the names the predicates used to
' identify printable values (e.g. a date is printable but a workbook is not)
'
' PARAMETERS
'
' RETURNED VALUE
' array of strings with the names the predicates used to identify pritable values.
Public Function GetPrintableTypePredicateNames() As Variant
    Let GetPrintableTypePredicateNames = Array("BooleanQ", _
                                               "DateQ", _
                                               "NumberQ", _
                                               "StringQ", _
                                               "EmptyQ", _
                                               "NullQ", _
                                               "TrueQ", _
                                               "FalseQ")
End Function

' DESCRIPTION
' Boolean function returning True if its argument is printable (e.g. numeric, string, date, Boolean,
' Empty or Null).  It returns False even if its argument is a printable array or table.  This is used
' to detect printable atomic elements.
'
' PARAMETERS
' 1. arg - Any Excel value or reference
'
' RETURNED VALUE
' Returns True or False depending on whether or not its argument can be considered a printable atom.
Public Function PrintableQ(arg As Variant) As Boolean
    Let PrintableQ = AnyTrueQ(Through(GetPrintableTypePredicateNames(), arg))
End Function

' DESCRIPTION
' This function returns an array of intergers with the VB values for the various
' vartypes representing integers.  These are the values returned by VarType() and
' come from enumeration
'
' PARAMETERS
'
' RETURNED VALUE
' array of intergers with the VB values for the various vartypes representing integers
Public Function GetNumericVarTypes() As Variant
    #If Win64 Then
        Let GetNumericVarTypes = Array(vbByte, vbCurrency, vbDecimal, vbDouble, vbInteger, _
                                       vbLong, vbLongLong, vbSingle)
    #Else
        Let GetNumericVarTypes = Array(vbByte, vbCurrency, vbDecimal, vbDouble, vbInteger, _
                                       vbLong, vbSingle)
    #End If
End Function

' DESCRIPTION
' Returns True when its argument is one of the types returned by function
' GetAtomicTypePredicateNames(). Returns False otherwise
'
' PARAMETERS
' 1. arg - any value or object reference
'
' RETURNED VALUE
' True when arg has one of the types detailed above. Returns False otherwise
Public Function AtomicQ(arg As Variant) As Boolean
Attribute AtomicQ.VB_Description = "This is the documentation"
    Dim var As Variant
    
    Let AtomicQ = True
    
    For Each var In GetAtomicTypePredicateNames()
        If Run(var, arg) Then Exit Function
    Next
    
    Let AtomicQ = False
End Function

' DESCRIPTION
' Boolean function returning True if its argument is True. Returns False otherwise.
'
' PARAMETERS
' 1. arg - any value or object reference
'
' RETURNED VALUE
' True or False depending on whether or not its argument is equal to True
Public Function TrueQ(vValue As Variant) As Boolean
    If IsObject(vValue) Then
        Let TrueQ = False
    Else
        Let TrueQ = (vValue = True)
    End If
End Function

' DESCRIPTION
' Boolean function returning True if its argument is False. Returns False otherwise.
'
' PARAMETERS
' 1. arg - any value or object reference
'
' RETURNED VALUE
' True or False depending on whether or not its argument is equal to False
Public Function FalseQ(vValue As Variant) As Boolean
    If IsObject(vValue) Then
        Let FalseQ = False
    Else
        Let FalseQ = (vValue = False)
    End If
End Function

' DESCRIPTION
' Boolean function returning True if its argument satisfies IsEmpty(). Returns False otherwise.
'
' PARAMETERS
' 1. arg - any value or object reference
'
' RETURNED VALUE
' True or False depending on whether or not its argument satisfies IsEmpty()
Public Function EmptyQ(vValue As Variant) As Boolean
    Let EmptyQ = IsEmpty(vValue)
End Function

' DESCRIPTION
' Boolean function returning True if its argument satisfies IsNull(). Returns False otherwise.
'
' PARAMETERS
' 1. arg - any value or object reference
'
' RETURNED VALUE
' True or False depending on whether or not its argument satisfies IsNull()
Public Function NullQ(vValue As Variant) As Boolean
    Let NullQ = IsNull(vValue)
End Function

' DESCRIPTION
' Boolean function returning True if its argument satisfies IsError(). Returns False otherwise.
'
' PARAMETERS
' 1. arg - any value or object reference
'
' RETURNED VALUE
' True or False depending on whether or not its argument satisfies IsError()
Public Function ErrorQ(vValue As Variant) As Boolean
    Let ErrorQ = IsNull(vValue)
End Function

' DESCRIPTION
' Boolean function returning True if the given parameter is 0.
'
' PARAMETERS
' 1. arg - any Excel expression
'
' RETURNED VALUE
' Returns True or False depending on whether or not its argument is zero
Public Function ZeroQ(arg As Variant) As Boolean
    Let ZeroQ = False

    If NumberQ(arg) Then Let ZeroQ = arg = 0
End Function

' DESCRIPTION
' Boolean function returning True if its argument is a nonzero number.
'
' PARAMETERS
' 1. arg - any value or object reference
'
' RETURNED VALUE
' True or False depending on whether or not its argument is a nonzero number.
Public Function NonZeroQ(arg As Variant) As Boolean
    Let NonZeroQ = False
    
    If NumberQ(arg) Then Let NonZeroQ = arg <> 0
End Function

' DESCRIPTION
' Boolean function returning True if the given parameter is a number equal to 1.
'
' PARAMETERS
' 1. arg - any Excel expression
'
' RETURNED VALUE
' Returns True or False depending on whether or not its argument is a number equal
' to 1
Public Function OneQ(arg As Variant) As Boolean
    Let OneQ = False

    If NumberQ(arg) Then Let OneQ = arg = 1
End Function

' DESCRIPTION
' Boolean function returning True if its argument is a number not equal to 1.
'
' PARAMETERS
' 1. arg - any value or object reference
'
' RETURNED VALUE
' True or False depending on whether or not its argument is a number not equal to 1.
Public Function NonOneQ(arg As Variant) As Boolean
    Let NonOneQ = False
    
    If NumberQ(arg) Then Let NonOneQ = arg <> 1
End Function

' DESCRIPTION
' Boolean function returning True if its argument is a number.  Returns False otherwise
'
' PARAMETERS
' 1. arg - any value or object reference
'
' RETURNED VALUE
' True when arg is byte, integer, long, single, double, or currency. LongLong
' is available on 64-bit systems exclusively, and Decimal is not found in
' my systems.  Returns False otherwise.
Public Function NumberQ(arg As Variant) As Boolean
    Dim var As Variant
    
    Let NumberQ = True
    
    For Each var In GetNumericVarTypes()
        If var = VarType(arg) Then Exit Function
    Next
    
    Let NumberQ = False
End Function

' DESCRIPTION
' Boolean function returning True if its argument is positive. Returns
' False otherwise.
'
' PARAMETERS
' 1. arg - any value or object reference
'
' RETURNED VALUE
' True or False depending on whether or not its argument is positive.
Public Function PositiveQ(arg As Variant) As Boolean
    Let PositiveQ = False

    If NumberQ(arg) Then Let PositiveQ = arg > 0
End Function

' DESCRIPTION
' Boolean function returning True if its argument is negative. Returns
' False otherwise.
'
' PARAMETERS
' 1. arg - any value or object reference
'
' RETURNED VALUE
' True or False depending on whether or not its argument is negative.
Public Function NegativeQ(arg As Variant) As Boolean
    Let NegativeQ = False

    If NumberQ(arg) Then Let NegativeQ = arg < 0
End Function

' DESCRIPTION
' Boolean function returning True if its argument is non-positive. Returns
' False otherwise.
'
' PARAMETERS
' 1. arg - any value or object reference
'
' RETURNED VALUE
' True or False depending on whether or not its argument is non-positive.
Public Function NonPositiveQ(arg As Variant) As Boolean
    Let NonPositiveQ = False

    If NumberQ(arg) Then Let NonPositiveQ = arg <= 0
End Function

' DESCRIPTION
' Boolean function returning True if its argument is non-negative. Returns
' False otherwise.
'
' PARAMETERS
' 1. arg - any value or object reference
'
' RETURNED VALUE
' True or False depending on whether or not its argument is non-negative.
Public Function NonNegativeQ(arg As Variant) As Boolean
    Let NonNegativeQ = False

    If NumberQ(arg) Then Let NonNegativeQ = arg >= 0
End Function

' DESCRIPTION
' Boolean function returning True if its argument is a any non-whole number. Returns
' False otherwise.
'
' PARAMETERS
' 1. arg - any value or object reference
'
' RETURNED VALUE
' True or False depending on whether or not its argument is a non-whole number.
Public Function NonWholeNumberQ(arg As Variant) As Boolean
    Let NonWholeNumberQ = False

    If Not NumberQ(arg) Then Exit Function
    
    If CLng(arg) = arg Then Exit Function
    
    Let NonWholeNumberQ = True
End Function

' DESCRIPTION
' Boolean function returning True if its argument is a any whole number type. Returns False otherwise.
'
' PARAMETERS
' 1. arg - any value or object reference
'
' RETURNED VALUE
' True or False depending on whether or not its argument is a whole number.
Public Function WholeNumberQ(arg As Variant) As Boolean
    Dim TheVarType As VbVarType

    Let TheVarType = VarType(arg)

    If TheVarType = vbByte Or TheVarType = vbInteger Or TheVarType = vbLong Then
        Let WholeNumberQ = True
    ' The reason this next iff is necessary is that [{1,2}] passes 1 and 2 with type
    ' double.  Don't know why.  Makes no sense.  In any case, all we care about is that
    ' this has no nonzero decimal part.
    ElseIf TheVarType = vbDecimal Or TheVarType = vbDouble Or TheVarType = vbDouble Then
        Let WholeNumberQ = CLng(arg) = arg
    Else
        Let WholeNumberQ = False
    End If
End Function

' DESCRIPTION
' Boolean function returning True if its argument is a any whole number type and is positive.
' Returns False otherwise.
'
' PARAMETERS
' 1. arg - any value or object reference
'
' RETURNED VALUE
' True or False depending on whether or not its argument is a positive whole number.
Public Function PositiveWholeNumberQ(arg As Variant) As Boolean
    Let PositiveWholeNumberQ = False
    
    If WholeNumberQ(arg) Then Let PositiveWholeNumberQ = arg > 0
End Function

' DESCRIPTION
' Boolean function returning True if its argument is a negative whole number. Returns False
' otherwise.
'
' PARAMETERS
' 1. arg - any value or object reference
'
' RETURNED VALUE
' True or False depending on whether or not its argument is a negative whole number.
Public Function NegativeWholeNumberQ(arg As Variant) As Boolean
    Let NegativeWholeNumberQ = False
    
    If WholeNumberQ(arg) Then Let NegativeWholeNumberQ = arg < 0
End Function

' DESCRIPTION
' Boolean function returning True if its argument is a non-positive whole number. Returns
' False otherwise.
'
' PARAMETERS
' 1. arg - any value or object reference
'
' RETURNED VALUE
' True or False depending on whether or not its argument is a non-positive whole number.
Public Function NonPositiveWholeNumberQ(arg As Variant) As Boolean
    Let NonPositiveWholeNumberQ = False
    
    If WholeNumberQ(arg) Then Let NonPositiveWholeNumberQ = arg <= 0
End Function

' DESCRIPTION
' Boolean function returning True if its argument is a non-negative whole number. Returns
' False otherwise.
'
' PARAMETERS
' 1. arg - any value or object reference
'
' RETURNED VALUE
' True or False depending on whether or not its argument is a non-negative whole number.
Public Function NonNegativeWholeNumberQ(arg As Variant) As Boolean
    Let NonNegativeWholeNumberQ = False
    
    If WholeNumberQ(arg) Then Let NonNegativeWholeNumberQ = arg >= 0
End Function

' DESCRIPTION
' Boolean function returning True if its argument is a nonzero whole number.
'
' PARAMETERS
' 1. arg - any value or object reference
'
' RETURNED VALUE
' True or False depending on whether or not its argument is a nonzero, whole number.
Public Function NonZeroWholeNumberQ(arg As Variant) As Boolean
    Let NonZeroWholeNumberQ = False
    
    If WholeNumberQ(arg) Then Let NonZeroWholeNumberQ = arg <> 0
End Function

' DESCRIPTION
' Boolean function returns True if Returns True is its argument is 2D matrix with numeric entries.
'
' PARAMETERS
' 1. arg - Any Excel value or reference
'
' RETURNED VALUE
' Returns True or False depending on whether or not its argument can be considered a numerical matrix.
Public Function BooleanQ(arg As Variant) As Boolean
    Let BooleanQ = VarType(arg) = vbBoolean
End Function

' DESCRIPTION
' Boolean function returning True if its argument is a string. Returns False otherwise.
'
' PARAMETERS
' 1. arg - any value or object reference
'
' RETURNED VALUE
' True or False depending on whether or not its argument is a string
Public Function StringQ(AnArg As Variant) As Boolean
    Let StringQ = VarType(AnArg) = vbString
End Function

' DESCRIPTION
' Boolean function returning True if its argument is range reference.
'
' PARAMETERS
' 1. arg - an initialized workbook reference
'
' RETURNED VALUE
' True or False depending on whether arg its argument is an initialized range reference.
Public Function RangeQ(arg As Variant) As Boolean
    Let RangeQ = TypeName(arg) = TypeName(ThisWorkbook.Worksheets(1).Range("A1"))
End Function

' DESCRIPTION
' Boolean function returning True if its argument is an initialized workbook reference.
'
' PARAMETERS
' 1. arg - an initialized workbook reference
'
' RETURNED VALUE
' True or False depending on whether arg its argument is an initialized workbook reference.
Public Function WorkbookQ(arg As Variant) As Boolean
    Let WorkbookQ = TypeName(arg) = TypeName(ThisWorkbook)
End Function

' DESCRIPTION
' Boolean function returning True if its argument is an initialized worksheet reference.
'
' PARAMETERS
' 1. arg - an initialized workbook reference
'
' RETURNED VALUE
' True or False depending on whether arg its argument is an initialized worksheet reference.
Public Function WorksheetQ(arg As Variant) As Boolean
    Let WorksheetQ = TypeName(arg) = TypeName(ThisWorkbook.Worksheets(1))
End Function

' DESCRIPTION
' Boolean function returning True if its argument is a dictionary. Returns False otherwise.
'
' PARAMETERS
' 1. arg - any value or object reference
'
' RETURNED VALUE
' True or False depending on whether or not its argument is a dictionary
Public Function DictionaryQ(vValue As Variant) As Boolean
    Dim obj As New Dictionary
    
    Let DictionaryQ = TypeName(vValue) = TypeName(obj)
End Function

' DESCRIPTION
' Boolean function returning True if its argument satisfies IsDate(). Returns False otherwise.
'
' PARAMETERS
' 1. arg - any value or object reference
'
' RETURNED VALUE
' True or False depending on whether or not its argument satisfies IsDate()
Public Function DateQ(vValue As Variant) As Boolean
    Let DateQ = IsDate(vValue)
End Function

' DESCRIPTION
' Boolean function returning True if its argument is an initialized ListObject reference.
'
' PARAMETERS
' 1. arg - an initialized workbook reference
'
' RETURNED VALUE
' True or False depending on whether arg its argument is an initialized ListObject reference.
Public Function ListObjectQ(arg As Variant) As Boolean
    Let ListObjectQ = TypeName(arg) = "ListObject"
End Function

' DESCRIPTION
' Boolean function returning True if its parameter satisfies the requirements to be considered
' an index for Arrays.Part()
'
' PARAMETERS
' 1. arg - Expressions with of the following forms:
'
'        1. n - to get element n.  If given a 2D array, n refers to the row number
'        2. Span - An instance of class Span, which can be conveniently generated using
'                  ClassConstructors.Span()
'
' RETURNED VALUE
' Returns True or False depending on whether or not the given parameter has one of the acceptable forms
Public Function PartIndexQ(arg As Variant) As Boolean
    Let PartIndexQ = NonZeroWholeNumberQ(arg) Or NonzeroWholeNumberArrayQ(arg) Or SpanQ(arg)
End Function

' DESCRIPTION
' Boolean function returning True if its parameter satisfies the requirements to be considered
' an index for Take.
'
' PARAMETERS
' 1. TheIndex - Any Excel expression with of the following forms:
'
'        1. n - with n nonzero
'        2. [{n_1, n_2}] - with n_i nonzero
'        3. [{n_1, n_2, TheStep}] - with n_1, n2_2, TheStep<>0
'
' RETURNED VALUE
' Returns True or False depending on whether or not the given parameter has one of the acceptable forms
Public Function TakeIndexQ(TheIndex As Variant) As Boolean
     Let TakeIndexQ = True

    If NonZeroWholeNumberQ(TheIndex) Then
        Exit Function
    ElseIf NonzeroWholeNumberArrayQ(TheIndex) And Length(TheIndex) >= 1 And Length(TheIndex) <= 3 Then
        Exit Function
    Else
        Let TakeIndexQ = False
    End If
End Function

' DESCRIPTION
' Boolean function returning True if its argument is an instance of class Span.
'
' PARAMETERS
' 1. arg - any Excel value or reference
'
' RETURNED VALUE
' Returns True or False depending on whether or not its argument is an instance of class Span
Public Function SpanQ(arg As Variant) As Boolean
    Let SpanQ = TypeName(arg) = TypeName(ClassConstructors.Span(1, 1))
End Function

' DESCRIPTION
' Boolean function returning True if its argument is an instance of class Lambda.
'
' PARAMETERS
' 1. arg - any Excel value or reference
'
' RETURNED VALUE
' Returns True or False depending on whether or not its argument is an instance
' of class Span
Public Function LambdaQ(arg As Variant) As Boolean
    Let LambdaQ = TypeName(arg) = TypeName(Predicates.SampleLambda)
End Function

' DESCRIPTION
' Boolean function returning True if its argument is a control form button.
' Returns False otherwise.
'
' PARAMETERS
' 1. arg - any value or object reference
'
' RETURNED VALUE
' True or False depending on whether or not its argument is a button.
Public Function FormControlButtonQ(arg As Variant) As Boolean
    Let FormControlButtonQ = False
    
    If TypeName(arg) = "Shape" Then
        If arg.Type = msoFormControl Then
            If arg.FormControlType = xlButtonControl Then
                Let FormControlButtonQ = True
            End If
        End If
    End If
End Function

' DESCRIPTION
' Boolean function returning True if its argument is a string or a number.
' Returns False otherwise.
'
' PARAMETERS
' 1. arg - any value or object reference
'
' RETURNED VALUE
' True or False depending on whether or not its argument is a string or
' a number.
Public Function NumberOrStringQ(arg As Variant) As Boolean
    Let NumberOrStringQ = NumberQ(arg) Or StringQ(arg)
End Function

' DESCRIPTION
' Boolean function returning True if its argument is a string or a whole number. Returns False otherwise.
'
' PARAMETERS
' 1. arg - any value or object reference
'
' RETURNED VALUE
' True or False depending on whether or not its argument string or a whole number.
Public Function WholeNumberOrStringQ(arg As Variant) As Boolean
    Let WholeNumberOrStringQ = WholeNumberQ(arg) Or StringQ(arg)
End Function

' DESCRIPTION
' Boolean function returns True if the given file exists.  Returns False
' otherwise.
'
' PARAMETERS
' 1. ThePath - A valid path
'
' RETURNED VALUE
' True or False depending on whether the given file exists
Function FileExistsQ(ThePath As String) As Boolean
    Dim TheAttribute As Variant
                          
    On Error GoTo ErrorHandler

    Let TheAttribute = GetAttr(ThePath)
    
    Let FileExistsQ = (TheAttribute And vbNormal) Or _
                      (TheAttribute And vbReadOnly) Or _
                      (TheAttribute And vbHidden) Or _
                      (TheAttribute And vbSystem) Or _
                      (TheAttribute And vbArchive) Or _
                      (TheAttribute And vbAlias)
                      
    Exit Function
ErrorHandler:
    Let FileExistsQ = False
End Function

' DESCRIPTION
' Boolean function returns True if the given directory exists.  Returns False
' otherwise.
'
' PARAMETERS
' 1. ThePath - A valid path
'
' RETURNED VALUE
' True or False depending on whether the given directory exists
Function DirectoryExistsQ(ThePath As String) As Boolean
    On Error GoTo ErrorHandler

    Let DirectoryExistsQ = GetAttr(ThePath) And vbDirectory
                      
    Exit Function
ErrorHandler:
    Let DirectoryExistsQ = False
End Function

' DESCRIPTION
' Boolean function returns True if there obj is a reference initialized to a listcolumn.
' It returns False otherwise.
'
' PARAMETERS
' 1. WorkSheetReference - A worksheet reference
' 2. ListObjectName - A string representing the name of the listobject in question
'
' RETURNED VALUE
' True or False depending on whether obj is initialized with a reference to a listcolumn
Public Function ListColumnQ(obj As Variant) As Boolean
    Let ListColumnQ = False
    
    If Not IsObject(obj) Then Exit Function
    If Not TypeName(obj) = "ListColumn" Then Exit Function
    
    Let ListColumnQ = True
End Function


' DESCRIPTION
' Boolean function returns True if there is a listobject with the given name in the given
' worksheet.  Returns False otherwise.
'
' PARAMETERS
' 1. WorkSheetReference - A worksheet reference
' 2. ListObjectName - A string representing the name of the listobject in question
'
' RETURNED VALUE
' True or False depending on whether the a listobject with the given name exists on the
' given worksheet
Public Function ListObjectExistsQ(WorkSheetReference As Worksheet, ListObjectName As String) As Boolean
    Dim AName As String
    
    Let ListObjectExistsQ = True
    
    On Error GoTo ErrorHandler
    
    Let AName = WorkSheetReference.ListObjects(ListObjectName).Name
    
    Exit Function
    
ErrorHandler:
    Let ListObjectExistsQ = False
End Function

' DESCRIPTION
' Boolean function returns True if there obj is a reference initialized to a listcolumn.
' It returns False otherwise.
'
' PARAMETERS
' 1. WorkSheetReference - A worksheet reference
' 2. ListObjectName - A string representing the name of the listobject in question
'
' RETURNED VALUE
' True or False depending on whether obj is initialized with a reference to a listcolumn
Public Function ListRowQ(obj As Variant) As Boolean
    Let ListRowQ = False
    
    If Not IsObject(obj) Then Exit Function
    If Not TypeName(obj) = "ListRow" Then Exit Function
    
    Let ListRowQ = True
End Function

' DESCRIPTION
' Boolean function returning True if the given workbook has a worksheet with the given name.
'
' PARAMETERS
' 1. aWorkbook - A workbook reference
' 2. WorksheetName - A worksheet reference
'
' RETURNED VALUE
' Returns True or False depending on whether or not the given workbook has a worksheet with
' the given name
Public Function WorksheetExistsQ(AWorkbook As Workbook, WorksheetName As String) As Boolean
    Let WorksheetExistsQ = False
    
    On Error Resume Next
    
    Let WorksheetExistsQ = AWorkbook.Worksheets(WorksheetName).Name <> ""
    Exit Function
    
    On Error GoTo 0
End Function

' DESCRIPTION
' Boolean function returning True if the given workbook has a sheet with
' the given name.
'
' PARAMETERS
' 1. aWorkbook - A workbook reference
' 2. WorksheetName - A worksheet reference
'
' RETURNED VALUE
' Returns True or False depending on whether or not the given workbook has
' a sheet with the given name
Public Function SheetExistsQ(AWorkbook As Workbook, SheetName As String) As Boolean
    Let SheetExistsQ = False
    
    On Error GoTo NoSuchSheet
    If Len(AWorkbook.Sheets(SheetName).Name) > 0 Then
        Let SheetExistsQ = True
        Exit Function
    End If

NoSuchSheet:
End Function

' DESCRIPTION
' Boolean function returning True if TheValue is in the given 1D array.
' This function cannot compare arrays at the moment. For instance,
' MemberQ(Array(Array(1,2)), Array(1,2)) -> True. Objects are equal
' if they are references to the same object.
'
' PARAMETERS
' 1. TheArray - A 1D array
' 2. TheValue - Any Excel value or reference
'
' RETURNED VALUE
' Returns True or False depending on whether or not the given value is
' in the given array
Public Function MemberQ(TheArray As Variant, TheValue As Variant) As Boolean
    Dim i As Long
    
    ' Assume result is False and change TheValue is in any one column
    ' of TheArray
    Let MemberQ = False
    
    ' Exit if TheArray is not a 1D array
    If NumberOfDimensions(TheArray) <> 1 Then Exit Function
    
    For i = 1 To Length(TheArray)
        If EqualQ(Part(TheArray, i), TheValue) Then
            Let MemberQ = True
            Exit Function
        End If
    Next
End Function

' DESCRIPTION
' Boolean function returning True if TheValue is not in the given 1D
' array. TheValue must satisfy NumberOrStringQ. Every element in
' TheArray must also satisfy NumberOfStringQ
'
' PARAMETERS
' 1. TheArray - A 1D array satisfying PrintableArrayQ
' 2. TheValue - Any value satisfying PrintableQ
'
' RETURNED VALUE
' Returns True or False depending on whether or not the given value is
' in the given array
Public Function FreeQ(TheArray As Variant, TheValue As Variant) As Boolean
    Let FreeQ = IsArray(TheArray) And Not MemberQ(TheArray, TheValue)
End Function

' DESCRIPTION
' Boolean function returning True if its parameter has type integer
'
' PARAMETERS
' 1. var - any expression
'
' RETURNED VALUE
' True if the arguments has type integer
Public Function IntegerTypeQ(var As Variant) As Boolean
    Let IntegerTypeQ = VarType(var) = vbInteger
End Function

' DESCRIPTION
' Boolean function returning True if its parameter has type long
'
' PARAMETERS
' 1. var - any expression
'
' RETURNED VALUE
' True if the arguments has type long
Public Function LongTypeQ(var As Variant) As Boolean
    Let LongTypeQ = VarType(var) = vbLong
End Function

' This could be rewritten using VarType and comparing to vbSingle
' DESCRIPTION
' Boolean function returning True if its parameter has type single
'
' PARAMETERS
' 1. var - any expression
'
' RETURNED VALUE
' True if the arguments has type single
Public Function SingleTypeQ(var As Variant) As Boolean
    Let SingleTypeQ = VarType(var) = vbSingle
End Function

' DESCRIPTION
' Boolean function returning True if its parameter has type double
'
' PARAMETERS
' 1. var - any expression
'
' RETURNED VALUE
' True if the arguments has type double
Public Function DoubleTypeQ(var As Variant) As Boolean
    Let DoubleTypeQ = VarType(var) = vbDouble
End Function

' DESCRIPTION
' Boolean function returning True if its two parameters are equal according
' to the following rules:
'
' 1. If IsNull(x) and IsNull(y) then EqualQ(x,y) -> True
' 2. If IsObject(x) and IsObject(y) and (x Is y) then EqualQ(x,y)->True
' 3. If neither of x,y Null or Object then x=y -> EqualQ(x,y)
' 4. If x,y arrays then x,y is they have the same dimensions and
'    the corresponding elements all satisfy EqualQ
'
' Arguments to this function can be any valid expression. However, it returns
' false is any element at any level has more than two dimensions. This could
' be implemented, but I can only think about how to write it by writing code
' dynamically.
'
' PARAMETERS
' 1. x - any expression
' 2. y - any expression
' 3. Optional NumericallyExactQ = True - Number comparisons must be exact
' 4. Optional NumericalThreshold = False - Threshold under which two numbers are considered equal
'
' RETURNED VALUE
' True if the arguments are the same and False otherwise
Public Function EqualQ(x As Variant, y As Variant, _
                       Optional NumericallyExactQ As Boolean = False, _
                       Optional NumericalThreshold As Double = 0.000001) As Boolean
    Dim i As Long
    Dim j As Long
    
    ' Assume result is False and change TheValue is in any one column
    ' of TheArray
    Let EqualQ = False
    
    ' Exit if the two objects don't have the same dimensions. That means
    ' that either one is an array and the other is not or that they are
    ' arrays of different dimensions.
    If Dimensions(x) <> Dimensions(y) Then Exit Function
    
    ' Exit with False if either x or y has dimension larger than 2
    If Dimensions(x) > 2 Or Dimensions(y) > 2 Then Exit Function
    
    ' If we two 1D arrays, we recurse on each element and exit on the
    ' first False
    If IsArray(x) And IsArray(y) Then
        If Dimensions(x) = 1 And Dimensions(y) = 1 Then
            If Length(x) = Length(y) Then
                For i = 1 To Length(x)
                    If Not EqualQ(Part(x, i), Part(y, i)) Then Exit Function
                Next
                
                Let EqualQ = True
                Exit Function
            Else
                Exit Function
            End If
        ElseIf Dimensions(x) = 2 And Dimensions(y) = 2 Then
            If NumberOfRows(x) = NumberOfRows(y) And _
               NumberOfColumns(x) = NumberOfColumns(y) Then
                For i = 1 To NumberOfRows(x)
                    For j = 1 To NumberOfColumns(x)
                        If Not EqualQ(Part(x, i, j), Part(y, i, j)) Then
                            Exit Function
                        End If
                    Next
                Next
                
                Let EqualQ = True
                Exit Function
            End If
        Else
            Exit Function
        End If
        
        Let EqualQ = True
    ElseIf NumberQ(x) And NumberQ(y) Then
        If NumericallyExactQ Then
            Let EqualQ = (x = y) And _
                         ((IntegerTypeQ(x) And IntegerTypeQ(y)) Or _
                          (IntegerTypeQ(x) And LongTypeQ(y)) Or _
                          (LongTypeQ(x) And IntegerTypeQ(y)) Or _
                          (LongTypeQ(x) And LongTypeQ(y)) Or _
                          (SingleTypeQ(x) And SingleTypeQ(y)) Or _
                          (SingleTypeQ(x) And DoubleTypeQ(y)) Or _
                          (DoubleTypeQ(x) And SingleTypeQ(y)) Or _
                          (DoubleTypeQ(x) And DoubleTypeQ(y)))
        Else
            Let EqualQ = (Abs(x - y) < NumericalThreshold)
        End If
    ElseIf IsNull(x) And IsNull(y) Then
        Let EqualQ = True
    ElseIf IsObject(x) And IsObject(y) Then
        Let EqualQ = (x Is y)
    ElseIf IsError(x) And IsError(y) Then
        Let EqualQ = (CLng(x) = CLng(y))
    ElseIf TypeName(x) = TypeName(y) Then
        Let EqualQ = (x = y)
    End If
End Function

' DESCRIPTION
' Returns True if its argument is a directory node and False otherwise. A Directory
' node is defined as 1-element array whose only element is a string or a 2-element
' array whose first element is a string and and whose second element is an array of
' directory tree nodes.  In other words,
'
' - Array("Dir")
' - Array("Dir", Array())
' - Array("Dir1", Array(...)), where ... is of the types above
'
' PARAMETERS
' 1. arg - any Excel expression
'
' RETURNED VALUE
' True if the argument is a directory tree.
Public Function DirectoryTreeQ(arg As Variant) As Boolean
    Dim var As Variant
    
    ' Return False as the default value
    Let DirectoryTreeQ = False
    
    ' Exit if the argument is not a 2-element array
    If Not DimensionedQ(arg) Then Exit Function
    If Length(arg) < 1 Then Exit Function
    
    ' Exit with error code if first arg is not a directory name
    If Not StringQ(First(arg)) Then Exit Function
    
    ' Make sure second element of argument is an array of tree nodes
    If Length(arg) = 1 Then
        Let DirectoryTreeQ = True
    ElseIf EmptyArrayQ(Last(arg)) Then
        Let DirectoryTreeQ = True
    Else
        For Each var In Last(arg): Let DirectoryTreeQ = DirectoryTreeQ(var): Next
    End If
End Function

' DESCRIPTION
' Boolean function returns True if its argument is a reference to a shape
'
' PARAMETERS
' 1. arg - Any Excel value or reference
'
' RETURNED VALUE
' Returns True or False depending on whether or not its argument is a reference to a shape
Public Function ShapeQ(arg As Variant) As Boolean
    Let ShapeQ = (VarType(arg) = vbObject) And (TypeName(arg) = "Shape")
End Function

' DESCRIPTION
' Boolean function returns True if the given name has been defined in the given workshet
'
' PARAMETERS
' 1. TheName - Name of the range
' 2. Wsht - Reference to a worksheet
'
' RETURNED VALUE
' Returns True or False depending on whether or not the given name exists in the given worksheet
Public Function WorkSheetNameExistsQ(TheName As String, wsht As Worksheet) As Boolean
    Dim AString As String

    On Error GoTo ErrorHandler
    
    Let WorkSheetNameExistsQ = True

    Let AString = wsht.Names(TheName).Name
    
    Exit Function
    
ErrorHandler:
    Let WorkSheetNameExistsQ = False
End Function

' DESCRIPTION
' Boolean function returns True if the given name has been defined in the given workbook
'
' PARAMETERS
' 1. TheName - Name of the range
' 2. Wbk - Reference to a worksheet
'
' RETURNED VALUE
' Returns True or False depending on whether or not the given name exists in the given workbook
Public Function WorkbookNameExistsQ(TheName As String, Wbk As Workbook) As Boolean
    Dim AString As String

    On Error GoTo ErrorHandler
    
    Let WorkbookNameExistsQ = True

    Let AString = Wbk.Names(TheName).Name
    
    Exit Function
    
ErrorHandler:
    Let WorkbookNameExistsQ = False
End Function

' DESCRIPTION
' Boolean function returns True if the given worksheet name refers to a range
'
' PARAMETERS
' 1. TheName - Name of the range
' 2. Wsht - Reference to a worksheet
'
' RETURNED VALUE
' True if the given name is for a worksheet name that refers to a range
Public Function WorkSheetNameRefersToRangeQ(TheName As String, wsht As Worksheet) As Boolean
    Dim ARange As Range

    On Error GoTo ErrorHandler
    
    Let WorkSheetNameRefersToRangeQ = True

    Set ARange = wsht.Names(TheName).RefersToRange
    
    Exit Function
    
ErrorHandler:
    Let WorkSheetNameRefersToRangeQ = False
End Function

' DESCRIPTION
' Boolean function returns True if the given workbook name refers to a range
'
' PARAMETERS
' 1. TheName - Name of the range
' 2. Wbk - Reference to a workbook
'
' RETURNED VALUE
' True if the given name is for a workbook name that refers to a range
Public Function WorkbookNameRefersToRangeQ(TheName As String, Wbk As Workbook) As Boolean
    Dim ARange As Range

    On Error GoTo ErrorHandler
    
    Let WorkbookNameRefersToRangeQ = True

    Set ARange = Wbk.Names(TheName).RefersToRange
    
    Exit Function
    
ErrorHandler:
    Let WorkbookNameRefersToRangeQ = False
End Function

' DESCRIPTION
' Boolean function returns True if the first cell in the range has validation set.
' Returns False if the argument is not initiazed
'
' PARAMETERS
' 1. ARange (Range) - A reference to a range
'
' RETURNED VALUE
' True if the first cell in the given range has been validated
Public Function CellValidatedQ(ARange As Range) As Boolean
    Dim v As Variant

    Let CellValidatedQ = True
    
    If ARange Is Nothing Then
        Let CellValidatedQ = False
        Exit Function
    End If
    
    On Error GoTo ErrorChecker
    
    Let v = ARange(1, 1).Validation.Type
    
    Exit Function
    
ErrorChecker:
    Let CellValidatedQ = False
End Function
