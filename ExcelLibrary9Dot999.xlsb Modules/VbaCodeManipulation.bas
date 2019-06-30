Attribute VB_Name = "VbaCodeManipulation"
' PURPOSE OF THIS MODULE
'
' The purpose of this module is to provide faciities to manipulate VBA code
' programmatically. This modules introduces the functionality to enable
' functional programming in VBA.
'
' Remember to add a reference to Microsoft Visual Basic for Applications Extensibility

Option Base 1
Option Explicit

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" _
(ByVal lpClassName As String, ByVal lpWindowName As String) As Long
 
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" _
(ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, _
ByVal lpsz2 As String) As Long
 
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" _
(ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
 
Private Declare Function GetWindowTextLength Lib "user32" Alias _
"GetWindowTextLengthA" (ByVal hWnd As Long) As Long
 
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
 
Dim Ret As Long, ChildRet As Long
Dim OpenRet As Long
Dim strBuff As String
Dim ButCap As String
Dim MyPassword As String
 
Const WM_SETTEXT = &HC
Const BM_CLICK = &HF5

Public Sub UnlockVbaAndAddReference(xlapp As Application, wbk As Workbook, MyPassword As String)
    Call DetectExcel
    
    '~~> Launch the VBA Project Password window
    '~~> I am assuming that it is protected. If not then
    '~~> put a check here.
    Call xlapp.VBE.CommandBars(1).FindControl(ID:=2578, recursive:=True, Visible:=False).Execute
 
    '~~> Get the handle of the "VBAProject Password" Window
    Let Ret = FindWindow(vbNullString, "VBAProject Password")
 
    If Ret <> 0 Then
        'MsgBox "VBAProject Password Window Found"
 
        '~~> Get the handle of the TextBox Window where we need to type the password
        Let ChildRet = FindWindowEx(Ret, ByVal 0&, "Edit", vbNullString)
 
        If ChildRet <> 0 Then
            'MsgBox "TextBox's Window Found"
            '~~> This is where we send the password to the Text Window
            Call SendMess(MyPassword, ChildRet)
 
            DoEvents
 
            '~~> Get the handle of the Button's "Window"
            Let ChildRet = FindWindowEx(Ret, ByVal 0&, "Button", vbNullString)
 
            '~~> Check if we found it or not
            If ChildRet <> 0 Then
                'MsgBox "Button's Window Found"
 
                '~~> Get the caption of the child window
                Let strBuff = String(GetWindowTextLength(ChildRet) + 1, Chr$(0))
                Call GetWindowText(ChildRet, strBuff, Len(strBuff))
                Let ButCap = strBuff
 
                '~~> Loop through all child windows
                Do While ChildRet <> 0
                    '~~> Check if the caption has the word "OK"
                    If InStr(1, ButCap, "OK") Then
                        '~~> If this is the button we are looking for then exit
                        Let OpenRet = ChildRet
                        Exit Do
                    End If
 
                    '~~> Get the handle of the next child window
                    Let ChildRet = FindWindowEx(Ret, ChildRet, "Button", vbNullString)
                    '~~> Get the caption of the child window
                    Let strBuff = String(GetWindowTextLength(ChildRet) + 1, Chr$(0))
                    Call GetWindowText(ChildRet, strBuff, Len(strBuff))
                    Let ButCap = strBuff
                Loop
 
                '~~> Check if we found it or not
                If OpenRet <> 0 Then
                    '~~> Click the OK Button
                    Call SendMessage(ChildRet, BM_CLICK, 0, vbNullString)
                    
                    ' To get rid of the VBA Project properties dialogue
                    DoEvents
                    Call SendKeys("{ESC}")
                    DoEvents
                Else
                    Call MsgBox("The Handle of OK Button was not found for file " & wbk.Name)
                End If
            Else
                 Call MsgBox("Button's Window Not Found for file " & wbk.Name)
            End If
        Else
            Call MsgBox("The Edit Box was not found for file " & wbk.Name)
        End If
    Else
        Call MsgBox("VBAProject Password Window was not Found" & "for file " & wbk.Name)
    End If
End Sub

' Helper for Sub VabCodeManipulation.UnlockVbaAndAddReference
Private Sub SendMess(Message As String, hWnd As Long)
    Call SendMessage(hWnd, WM_SETTEXT, False, ByVal Message)
End Sub

' Helper for Sub VabCodeManipulation.UnlockVbaAndAddReference
Private Sub DetectExcel()
    ' The procedure detects that it is running Excel and registers it.
    Const WM_USER = 1024
    Dim hWnd As Long
    
    ' If Excel is running this API call returns the controller.
    Let hWnd = FindWindow("XLMAIN", 0)
    
    If hWnd = 0 Then
        ' 0 mean that Excel is not running.
        Exit Sub
    Else
    'Excel is running, so the SendMessage API function is used
     'To enter it into the Running Object table.
        Call SendMessage(hWnd, WM_USER + 18, 0, 0)
    End If
End Sub

' Exports all VBA project components containing code to a folder in the same directory
' as this spreadsheet.
Public Sub ExportAllComponents(TheWorkbook As Workbook)
    Dim VBComp As VBIDE.VBComponent
    Dim destDir As String, fName As String, ext As String
    
    'Create the directory where code will be created.
    'Alternatively, you could change this so that the user is prompted
    If TheWorkbook.Path = "" Then
        Call MsgBox("You must first save this workbook somewhere so that it has a path.", , "Error")
        Exit Sub
    End If
    
    Let destDir = TheWorkbook.Path & "\" & TheWorkbook.Name & " Modules"
    
    ' Delet the directory (if it exists already) where the exported components are stored
    If DirectoryExistsQ(destDir) Then
        If Not EmptyArrayQ(GetFileNames(destDir & "\*.*")) Then
            Call Kill(destDir & "\*.*")
        End If
    Else
        Call MkDir(destDir)
    End If
    
    'Export all non-blank components to the directory
    For Each VBComp In TheWorkbook.VBProject.VBComponents
        If VBComp.CodeModule.CountOfLines > 0 Then
            'Determine the standard extention of the exported file.
            'These can be anything, but for re-importing, should be the following:
            Select Case VBComp.Type
                Case vbext_ct_ClassModule: ext = ".cls"
                Case vbext_ct_Document: ext = ".cls"
                Case vbext_ct_StdModule: ext = ".bas"
                Case vbext_ct_MSForm: ext = ".frm"
                Case Else: ext = vbNullString
            End Select
            
            If ext <> vbNullString Then
                Let fName = destDir & "\" & VBComp.Name & ext
                'Overwrite the existing file
                If Dir(fName, vbNormal) <> vbNullString Then Kill (fName)
                Call VBComp.Export(fName)
            End If
        End If
    Next VBComp
End Sub

' Calls ExportAllComponents(ThisWorkbook)
' It is to use only when coding the library. Does not matter who calls
' this sub. It always exports the library.
Public Sub BackUp()
    Call ExportAllComponents(ThisWorkbook)
End Sub

' The purpose of this function is to dump in the immediate console a
' function to draw this application's UI.  This function should be run
' right after "drawing" the UI by hand (which is easier).
'
' This function handles all form controls.  It does not handle the
' data for dropdowns, etc.  It handles the size, location, and macros
' associated with the controls.
Public Sub WriteUiCode(Wsht As Worksheet)
    Dim AShape As Shape
    Dim aShapeName As Variant
    Dim anListObjectName As Variant
    Dim AnAddress As String
    Dim ANumberFormat As String
    Dim AnAlignment As XlHAlign
    Dim AWidth As Double
    Dim ShapesDict As Dictionary
    Dim PropertiesDict As Dictionary
    Dim lo As ListObject
    Dim r As Long
    Dim LosDict As Dictionary            ' To store characteristics of listobjects in Wsht
    Dim NumberFormatsArray() As String
    Dim Alignments() As XlHAlign
    Dim TheFormulas() As String
    Dim TheColumnWidths() As Double
    
    Debug.Print Wsht.Name & " has " & Wsht.Shapes.Count & " shapes."
    
    ' Store the properties of every shape in the worksheet
    Set ShapesDict = New Dictionary
    For Each AShape In Wsht.Shapes
        Set PropertiesDict = New Dictionary
        
        Call PropertiesDict.Add("Top", AShape.Top)
        Call PropertiesDict.Add("Left", AShape.Left)
        Call PropertiesDict.Add("Width", AShape.Width)
        Call PropertiesDict.Add("Height", AShape.Height)
        Call PropertiesDict.Add("AlternativeText", AShape.AlternativeText)
        Call PropertiesDict.Add("OnAction", AShape.OnAction)
        If AShape.Type = 4 Then
            Call PropertiesDict.Add("XlFormControl", Null)
        Else
            Call PropertiesDict.Add("XlFormControl", AShape.FormControlType)
        End If
        
        Call ShapesDict.Add(Key:=AShape.Name, Item:=PropertiesDict)
    Next
    
    ' Store the properties of every listobject in the worksheet
    Set LosDict = New Dictionary
    For Each lo In Wsht.ListObjects
        ReDim NumberFormatsArray(1 To lo.ListColumns.Count)
        ReDim Alignments(1 To lo.ListColumns.Count)
        ReDim TheFormulas(1 To lo.ListColumns.Count)
        ReDim TheColumnWidths(1 To lo.ListColumns.Count)
        
        ' Store listcolumns' formats.  The format of the first cell in each listcolumn is
        ' taken as the format of the entire listcolumn
        For r = 1 To lo.ListColumns.Count
            Let NumberFormatsArray(r) = lo.ListColumns(r).DataBodyRange(1, 1).NumberFormat
            Let Alignments(r) = lo.ListColumns(r).DataBodyRange(1, 1).HorizontalAlignment
            Let TheFormulas(r) = lo.ListColumns(r).DataBodyRange(1, 1).Formula
            Let TheColumnWidths(r) = lo.ListColumns(r).Range.EntireColumn.Width
        Next

        Set PropertiesDict = New Dictionary
        ' Store reference to listobject's upper-left cell
        Call PropertiesDict.Add("UpperLeftCell", lo.Range(1, 1))
        ' Store listobject's headers
        Call PropertiesDict.Add("Headers", lo.HeaderRowRange.Value2)
        Call PropertiesDict.Add("Formats", NumberFormatsArray)
        Call PropertiesDict.Add("HorizontalAlignments", Alignments)
        Call PropertiesDict.Add("Formulas", TheFormulas)
        Call PropertiesDict.Add("ColumnWidths", TheColumnWidths)
        
        Call LosDict.Add(Key:=lo.Name, Item:=PropertiesDict)
    Next
    
    ' Write the code to recreate the UI in the given worksheet
    Debug.Print "Public Sub ReCreateUi()"
    Debug.Print "   Dim aShape As Shape"
    Debug.Print "   Dim wsht As Worksheet"
    Debug.Print "   Dim lo as ListObject"
    Debug.Print
    Debug.Print "   Set wsht = ThisWorkbook.Worksheets(""" & Wsht.Name & """)"
    Debug.Print "   For Each aShape In wsht.Shapes: Call aShape.Delete: Next"
    Debug.Print
    For Each aShapeName In ShapesDict.Keys
        If Not IsNull(ShapesDict.Item(aShapeName).Item("XlFormControl")) Then
            Debug.Print vbTab & "Set aShape = wsht.Shapes.AddFormControl(" & _
                        ShapesDict.Item(aShapeName).Item("XlFormControl") & _
                        ", " & ShapesDict.Item(aShapeName).Item("Left") & ", " & _
                        ShapesDict.Item(aShapeName).Item("Top") & ", " & _
                        ShapesDict.Item(aShapeName).Item("Width") & ", " & _
                        ShapesDict.Item(aShapeName).Item("Height") & ")"
            Debug.Print vbTab & "Let aShape.Name = """ & aShapeName & """"
            Debug.Print vbTab & "Let aShape.TextFrame.Characters.Text = """ & _
                        ShapesDict.Item(aShapeName).Item("AlternativeText") & """"
            Debug.Print vbTab & "Let aShape.OnAction = """ & _
                        ShapesDict.Item(aShapeName).Item("OnAction") & """"
            Debug.Print
        End If
    Next
    
    For Each anListObjectName In LosDict.Keys
        Let AnAddress = LosDict.Item(anListObjectName).Item("UpperLeftCell").Address
        Debug.Print vbTab & "Call DumpInSheet(" & _
                    Convert1DArrayOfStringToCode(LosDict.Item(anListObjectName).Item("Headers")) & _
                    ", wsht.Range(""" & AnAddress & """))"
        Debug.Print vbTab & "Set lo = AddListObject(wsht.Range(""" & AnAddress & """).CurrentRegion, """ & anListObjectName & """)"
        
        For r = 1 To LosDict.Item(anListObjectName).Item("Formats")
            Let ANumberFormat = Part(LosDict.Item(anListObjectName).Item("Formats"), r)
            Let AnAlignment = Part(LosDict.Item(anListObjectName).Item("HorizontalAlignments"), r)
            Let AWidth = Part(LosDict.Item(anListObjectName).Item("ColumnWidths"), r)
            Debug.Print vbTab & "Let lo.ListColumns(" & r & ").DataBodyRange.NumberFormat = """ & ANumberFormat & """"
            Debug.Print vbTab & "Let lo.ListColumns(" & r & ").Range.EntireColumn.HorizontalAlignment = " & AnAlignment
            Debug.Print vbTab & "Let lo.ListColumns(" & r & ").Range.EntireColumn.Width = " & AWidth
            '***HERE add formulas
        Next
    Next
    
    Debug.Print
    Debug.Print "End Sub"
End Sub

' DESCRIPTION
' Predicate indicating if the given module exists
'
' PARAMETERS
' 1. AWorkBook - A reference to a workbook object
' 2. ModuleName - The name of a module
'
' RETURNED VALUE
' True or False according to the exitence of the module
Public Function ModuleExistsQ(AWorkbook As Workbook, _
                              ModuleName As String) As Boolean
    Dim CodeModule As VBIDE.CodeModule
    Dim AVar As Variant
    
    ' call ErrorHandler in case of error
    On Error GoTo ErrorHandler
    
    ' Set default return value
    Let ModuleExistsQ = True
    
    ' Set reference to target code module
    Set CodeModule = AWorkbook.VBProject.VBComponents(ModuleName).CodeModule
    
    ' Exit before ErrorHandler since no error occured.
    Exit Function

' Handle errors. Returns False if module missing
ErrorHandler:
    Let ModuleExistsQ = False
End Function

' DESCRIPTION
' Predicate indicating if the given function has been defined
'
' PARAMETERS
' 1. AWorkBook - A reference to a workbook object
' 2. ModuleName - The name of a module
' 3. FunctionName - The name of a function
'
' RETURNED VALUE
' True or False according to the exitence of the function
Public Function FunctionExistsQ(AWorkbook As Workbook, _
                                ModuleName As String, _
                                FunctionName As String) As Boolean
    Dim CodeModule As VBIDE.CodeModule
    Dim AVar As Variant
    
    ' call ErrorHandler in case of error
    On Error GoTo ErrorHandler
    
    ' Set default return value
    Let FunctionExistsQ = True
    
    ' Set reference to target code module
    Set CodeModule = AWorkbook.VBProject.VBComponents(ModuleName).CodeModule

    Let AVar = CodeModule.ProcStartLine(FunctionName, vbext_pk_Proc)
    
    ' Exit before ErrorHandler since no error occured.
    Exit Function

' Handle errors. Returns False if either the module or function missing
ErrorHandler:
    Let FunctionExistsQ = False
End Function

' DESCRIPTION
' Inserts the a function into a target module. This sub does NOT checks for proper
' syntax.
'
' PARAMETERS
' 1. AWorkBook - A reference to a workbook object
' 2. ModuleName - The name of a module
' 3. FunctionName - The name of a function
' 4. ParameterNameArray - A 1D array of strings with the names of the function's parameters
' 5. FunctionBody - An array of strings with the body of the function. Each element must be
'                   one line of the function's body.
'
' RETURNED VALUE
' Inserts the function in the target module provided it does not already exists.
Public Sub InsertFunction(AWorkbook As Workbook, _
                          ModuleName As String, _
                          FunctionName As String, _
                          ParameterNameArray As Variant, _
                          FunctionBody As Variant)
    Dim CodeModule As VBIDE.CodeModule
    Dim CodeString As String
    Dim TmpStr As Variant
    
    ' Exit if the target module does not exists
    If Not ModuleExistsQ(AWorkbook, ModuleName) Then Exit Sub
    
    ' Exit if the function already exists in the given module and workbook
    If FunctionExistsQ(AWorkbook, ModuleName, FunctionName) Then Exit Sub

    ' Set reference to appropriate code module
    Set CodeModule = AWorkbook.VBProject.VBComponents(ModuleName).CodeModule

    ' Create string to hold function body
    Let CodeString = "Public Function " & FunctionName
    Let CodeString = CodeString & _
                     ToParentheticalString(ParameterNameArray) & _
                     " As Variant" & vbCrLf

    For Each TmpStr In FunctionBody
        Let CodeString = CodeString & TmpStr & vbCrLf
    Next

    Let CodeString = CodeString & "End Function"

    ' Add function to top of module
    Call CodeModule.AddFromString(CodeString)
End Sub

' DESCRIPTION
' Deletes the function with name FunctionName from ModuleName in workbook AWorkBook
'
' PARAMETERS
' 1. AWorkBook - A reference of type Workbook
' 2. ModuleName - A string name for the module where the function is contained
' 3. FunctionName - The string name of the function to delete
'
' RETURNED VALUE
' None
Public Sub DeleteFunction(AWorkbook As Workbook, _
                          ModuleName As String, _
                          FunctionName As String)
    Dim CodeModule As VBIDE.CodeModule

    ' Simply continue without raising error if the function is not found in
    ' the given module
    On Error Resume Next
    
    ' Set reference to approrpriate code module
    Set CodeModule = AWorkbook.VBProject.VBComponents(ModuleName).CodeModule

    ' Delete the funnction from the code module
    Call CodeModule.DeleteLines(CodeModule.ProcStartLine(FunctionName, vbext_pk_Proc), _
                                CodeModule.ProcCountLines(FunctionName, vbext_pk_Proc))
End Sub

' DESCRIPTION
' Clears the contents of module LambdaFunctionsTemp and resets workbook name
' LambdaFunctionCounter to 0.
'
' PARAMETERS
' None
'
' RETURNED VALUE
' None
Public Sub ClearLambdaFunctionsData()
    Dim CodeModule As VBIDE.CodeModule

    ' Simply continue without raising error if the module is not found in
    ' the given module
    On Error Resume Next
    
    ' Set reference to approrpriate code module
    Set CodeModule = ThisWorkbook.VBProject.VBComponents("LambdaFunctionsTemp").CodeModule

    ' Delete the funnction from the code module
    Call CodeModule.DeleteLines(1, CodeModule.CountOfLines)
    
    ' Reset lambda counter
    Let ThisWorkbook.Names("LambdaFunctionCounter").Value = 0
End Sub

' DESCRIPTION
' Creates and returns the name of a function to allow splicing the contents of a 1D array
' as the arguments of a function. This function provides the same functionality as the
' apply function in EmacsLisp.
'
' There is no way in to take a<-array(1,2) and then apply("add", a) to get add(1,2)
'
' This function returns a function name that can be passed to Apply to achieve this
' behavior. One does not need to call ParameterSplicingDelegate directly. This helper
' function is called by Apply. However, it is also called by MapThread to achieve similar
' behavior.
'
' PARAMETERS
' 1. FunctionName - The name of the function
' 2. N - A positive integer indicating the number of arguments the delegate must accept
'
' RETURNED VALUE
' Returns the name of the delegate function for splicing an array into the parameter
' slots of the given function
Public Function ParameterSplicingDelegate(FunctionName As String, n As Integer) As Variant
    Dim ParamNames As Variant
    Dim ParenString As String
    Dim FunctionBody() As String
    Dim VarName As Variant
    Dim i As Long
    
    ' Set default return value in case of error
    Let ParameterSplicingDelegate = Null
    
    ' ErrorCheck: Exit with Null if N not a positive integer
    If Not PositiveWholeNumberQ(n) Then Exit Function
    
    ' Create parameter list for anonymous function
    Let ParamNames = GenerateStringSequence("Param", 1, n)

    ' Create the function body
    
    ' Construct a let statement for bind each var to its intended value
    ReDim FunctionBody(1 To 2 * n)
    For i = 1 To n
        Let FunctionBody(i) = _
            "Dim " & Part(ParamNames, i) & " as Variant"
    Next
    
    For i = n + 1 To 2 * n
        Let FunctionBody(i) = _
            "Let " & Part(ParamNames, i - n) & " = Part(ArrayToSplice," & i - n & ")"
    Next
        
    ' Create the anonymous function
    Let ParenString = ToParentheticalString(ParamNames)
    Let ParenString = Right(ParenString, Len(ParenString) - 1)
    Let ParameterSplicingDelegate = Lambda("ArrayToSplice", _
                                           FunctionBody, _
                                           "run(" & Chr(34) & FunctionName & Chr(34) & _
                                           "," & ParenString _
                                          ).FunctionName
End Function

' Returns a VBA code representation of the 2D array of values in a range of cells.
' This function is meant to be used to create data test cases in a worksheet and then
' have the application automatically write the VBA code representing the string.
Public Function ConvertRangeToVbaArray(ARange As Range, Optional As2DArrayQ As Boolean = True) As String
    Dim r As Long
    Dim c As Long
    Dim RowStringArray() As String
    Dim RowsArray() As String
    
    ReDim RowsArray(1 To ARange.Rows.Count)
    ReDim RowStringArray(1 To ARange.Columns.Count)
    
    For r = 1 To ARange.Rows.Count
        For c = 1 To ARange.Columns.Count
            If EmptyQ(ARange(r, c)) Then
                Let RowStringArray(c) = "Empty"
            ElseIf ARange(r, c).NumberFormat = "@" Or StringQ(ARange(r, c).Value2) Then
                Let RowStringArray(c) = """" & ARange(r, c).Value2 & """"
            Else
                Let RowStringArray(c) = ARange(r, c).Value2
            End If
        Next c
        
        Let RowsArray(r) = "Array(" & Join(RowStringArray, ",") & ")"
        
        If Not As2DArrayQ Then
            Let ConvertRangeToVbaArray = RowsArray(r)
            Exit Function
        End If
    Next r
    
    Let ConvertRangeToVbaArray = "Array(" & Join(RowsArray, "," & vbCrLf) & ")"
End Function

