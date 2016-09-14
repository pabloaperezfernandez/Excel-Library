Attribute VB_Name = "Predicates"
Option Explicit
Option Base 1

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
                                            "NumberQ", _
                                            "StringQ", _
                                            "WorkbookQ", _
                                            "WorksheetQ", _
                                            "EmptyQ", _
                                            "ErrorQ", _
                                            "NullQ")
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
Public Function AtomicQ(Arg As Variant) As Boolean
    Dim var As Variant
    
    Let AtomicQ = True
    
    For Each var In GetAtomicTypePredicateNames()
        If Run(var, Arg) Then Exit Function
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
Public Function ZeroQ(Arg As Variant) As Boolean
    Let ZeroQ = False

    If NumberQ(Arg) Then Let ZeroQ = Arg = 0
End Function

' DESCRIPTION
' Boolean function returning True if its argument is a nonzero number.
'
' PARAMETERS
' 1. arg - any value or object reference
'
' RETURNED VALUE
' True or False depending on whether or not its argument is a nonzero number.
Public Function NonZeroQ(Arg As Variant) As Boolean
    Let NonZeroQ = False
    
    If NumberQ(Arg) Then Let NonZeroQ = Arg <> 0
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
Public Function OneQ(Arg As Variant) As Boolean
    Let OneQ = False

    If NumberQ(Arg) Then Let OneQ = Arg = 1
End Function

' DESCRIPTION
' Boolean function returning True if its argument is a number not equal to 1.
'
' PARAMETERS
' 1. arg - any value or object reference
'
' RETURNED VALUE
' True or False depending on whether or not its argument is a number not equal to 1.
Public Function NonOneQ(Arg As Variant) As Boolean
    Let NonOneQ = False
    
    If NumberQ(Arg) Then Let NonOneQ = Arg <> 1
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
Public Function NumberQ(Arg As Variant) As Boolean
    Dim var As Variant
    
    Let NumberQ = True
    
    For Each var In GetNumericVarTypes()
        If var = VarType(Arg) Then Exit Function
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
Public Function PositiveQ(Arg As Variant) As Boolean
    Let PositiveQ = False

    If NumberQ(Arg) Then Let PositiveQ = Arg > 0
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
Public Function NegativeQ(Arg As Variant) As Boolean
    Let NegativeQ = False

    If NumberQ(Arg) Then Let NegativeQ = Arg < 0
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
Public Function NonPositiveQ(Arg As Variant) As Boolean
    Let NonPositiveQ = False

    If NumberQ(Arg) Then Let NonPositiveQ = Arg <= 0
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
Public Function NonNegativeQ(Arg As Variant) As Boolean
    Let NonNegativeQ = False

    If NumberQ(Arg) Then Let NonNegativeQ = Arg >= 0
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
Public Function NonWholeNumberQ(Arg As Variant) As Boolean
    Let NonWholeNumberQ = False

    If Not NumberQ(Arg) Then Exit Function
    
    If CLng(Arg) = Arg Then Exit Function
    
    Let NonWholeNumber = True
End Function

' DESCRIPTION
' Boolean function returning True if its argument is a any whole number type. Returns False otherwise.
'
' PARAMETERS
' 1. arg - any value or object reference
'
' RETURNED VALUE
' True or False depending on whether or not its argument is a whole number.
Public Function WholeNumberQ(Arg As Variant) As Boolean
    Dim TheVarType As VbVarType

    Let TheVarType = VarType(Arg)

    If TheVarType = vbByte Or TheVarType = vbInteger Or TheVarType = vbLong Then
        Let WholeNumberQ = True
    ' The reason this next iff is necessary is that [{1,2}] passes 1 and 2 with type
    ' double.  Don't know why.  Makes no sense.  In any case, all we care about is that
    ' this have no nonzero decimal part.
    ElseIf TheVarType = vbDecimal Or TheVarType = vbDouble Or TheVarType = vbDouble Then
        Let WholeNumberQ = CLng(Arg) = Arg
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
Public Function PositiveWholeNumberQ(Arg As Variant) As Boolean
    Let PositiveWholeNumberQ = False
    
    If WholeNumberQ(Arg) Then Let PositiveWholeNumberQ = Arg > 0
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
Public Function NegativeWholeNumberQ(Arg As Variant) As Boolean
    Let NegativeWholeNumberQ = False
    
    If WholeNumberQ(Arg) Then Let NegativeWholeNumberQ = Arg < 0
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
Public Function NonPositiveWholeNumberQ(Arg As Variant) As Boolean
    Let NonPositiveWholeNumberQ = False
    
    If WholeNumberQ(Arg) Then Let NonPositiveWholeNumberQ = Arg <= 0
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
Public Function NonNegativeWholeNumberQ(Arg As Variant) As Boolean
    Let NonNegativeWholeNumberQ = False
    
    If WholeNumberQ(Arg) Then Let NonNegativeWholeNumberQ = Arg >= 0
End Function

' DESCRIPTION
' Boolean function returning True if its argument is a nonzero whole number.
'
' PARAMETERS
' 1. arg - any value or object reference
'
' RETURNED VALUE
' True or False depending on whether or not its argument is a nonzero, whole number.
Public Function NonZeroWholeNumberQ(Arg As Variant) As Boolean
    Let NonZeroWholeNumberQ = False
    
    If WholeNumberQ(Arg) Then Let NonZeroWholeNumberQ = Arg <> 0
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
Public Function WorkbookQ(Arg As Variant) As Boolean
    Let WorkbookQ = TypeName(Arg) = TypeName(ThisWorkbook)
End Function

' DESCRIPTION
' Boolean function returning True if its argument is an initialized worksheet reference.
'
' PARAMETERS
' 1. arg - an initialized workbook reference
'
' RETURNED VALUE
' True or False depending on whether arg its argument is an initialized worksheet reference.
Public Function WorksheetQ(Arg As Variant) As Boolean
    Let WorksheetQ = TypeName(Arg) = TypeName(ThisWorkbook.Worksheets(1))
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
Public Function ListObjectQ(Arg As Variant) As Boolean
    Let ListObjectQ = TypeName(Arg) = "ListObject"
End Function

' DESCRIPTION
' Boolean function returning True if its parameter any of the following forms:
'
' PARAMETERS
' 1. arg - Expressions with of the following forms:
'
'        1. n - to get element n.  If given a 2D array, n refers to the row number
'        2. AnArray() - A non-empty array of valid indices
'        3. Span - An instance of class Span, which can be conveniently generated using
'                  ClassConstructors.Span()
'
' RETURNED VALUE
' Returns True or False depending on whether or not the given parameter has one of the acceptable forms
Public Function PartIndexQ(Arg As Variant) As Boolean
    Let PartIndexQ = NonZeroWholeNumberQ(Arg) Or _
                     (NonEmptyArrayQ(Arg) And NonzeroWholeNumberArrayQ(Arg)) Or _
                     SpanQ(Arg)
End Function

' DESCRIPTION
' Boolean function returning True if its parameter any of the following forms:
'
' PARAMETERS
' 1. TheIndex - Any Excel expression with of the following forms:
'
'        1. n - with n nonzero
'        2. [{n_1, n_2}] - with n_i nonzero
'
' RETURNED VALUE
' Returns True or False depending on whether or not the given parameter has one of the acceptable forms
Public Function TakeIndexQ(TheIndex As Variant) As Boolean
    Let TakeIndexQ = NonZeroWholeNumberQ(TheIndex) Or (NonzeroWholeNumberArrayQ(TheIndex) And Length(TheIndex) = 2)
End Function

' DESCRIPTION
' Boolean function returning True if its argument is an instance of class Span.
'
' PARAMETERS
' 1. arg - any Excel value or reference
'
' RETURNED VALUE
' Returns True or False depending on whether or not its argument is an instance of class Span
Public Function SpanQ(Arg As Variant) As Boolean
    Let SpanQ = TypeName(Arg) = TypeName(ClassConstructors.Span(1, 1))
End Function

' DESCRIPTION
' Boolean function returning True if its argument is a string or a number. Returns False otherwise.
'
' PARAMETERS
' 1. arg - any value or object reference
'
' RETURNED VALUE
' True or False depending on whether or not its argument string or a number.
Public Function NumberOrStringQ(Arg As Variant) As Boolean
    Let NumberOrStringQ = NumberQ(Arg) Or StringQ(Arg)
End Function

' DESCRIPTION
' Boolean function returning True if its argument is a string or a whole number. Returns False otherwise.
'
' PARAMETERS
' 1. arg - any value or object reference
'
' RETURNED VALUE
' True or False depending on whether or not its argument string or a whole number.
Public Function WholeNumberOrStringQ(Arg As Variant) As Boolean
    Let WholeNumberOrStringQ = WholeNumberQ(Arg) Or StringQ(Arg)
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
Public Function WorksheetExistsQ(aWorkbook As Workbook, WorksheetName As String) As Boolean
    Let WorksheetExistsQ = False
    
    On Error Resume Next
    
    Let WorksheetExistsQ = aWorkbook.Worksheets(WorksheetName).Name <> ""
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
Public Function SheetExistsQ(aWorkbook As Workbook, SheetName As String) As Boolean
    Let SheetExistsQ = False
    
    On Error GoTo NoSuchSheet
    If Len(aWorkbook.Sheets(SheetName).Name) > 0 Then
        Let SheetExistsQ = True
        Exit Function
    End If

NoSuchSheet:
End Function
