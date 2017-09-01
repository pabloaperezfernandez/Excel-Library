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
                                            "NonZeroWholeNumberQ")
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
                                               "NullQ")
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
    Let StringQ = TypeName(AnArg) = TypeName("a")
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
' Returns True or False depending on whether or not its argument is an instance of class Span
Public Function LambdaQ(arg As Variant) As Boolean
    Let LambdaQ = TypeName(arg) = TypeName(Predicates.SampleLambda)
End Function

' DESCRIPTION
' Boolean function returning True if its argument is a string or a number. Returns False otherwise.
'
' PARAMETERS
' 1. arg - any value or object reference
'
' RETURNED VALUE
' True or False depending on whether or not its argument string or a number.
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
' Boolean function returns True if the given directory path or file exists.  Returns False
' otherwise.  You may pass a directory path or a full filename.
'
' PARAMETERS
' 1. ThePath - A valid path
'
' RETURNED VALUE
' True or False depending on whether the given path exists
Function FileExistsQ(ThePath As String) As Boolean
    If Not dir(ThePath, vbDirectory) = vbNullString Then
        FileExistsQ = True
    Else
        Let FileExistsQ = False
    End If
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
' Boolean function returning True if the given workbook has a sheet with the given name.
'
' PARAMETERS
' 1. aWorkbook - A workbook reference
' 2. WorksheetName - A worksheet reference
'
' RETURNED VALUE
' Returns True or False depending on whether or not the given workbook has a sheet with
' the given name
Public Function SheetExistsQ(AWorkbook As Workbook, SheetName As String) As Boolean
    Let SheetExistsQ = False
    
    On Error GoTo NoSuchSheet
    If Len(AWorkbook.Sheets(SheetName).Name) > 0 Then
        Let SheetExistsQ = True
        Exit Function
    End If

NoSuchSheet:
End Function
