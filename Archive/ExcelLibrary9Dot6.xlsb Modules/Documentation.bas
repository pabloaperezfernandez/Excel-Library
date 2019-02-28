Attribute VB_Name = "Documentation"
' PURPOSE OF THIS MODULE
'
' The purpose of this module is to provide facilities to retrieve
' documentation for modules and routines. In addition, this module
' provides functions to get the names of all components in a workbook.
'
' The module also provides an automatic documentation generator. The
' output is in Org format.

Option Explicit
Option Base 1

' DESCRIPTION
' Returns a routine's documentation. It is available as comments before the
' routine's declaration line
'
' PARAMETERS
' 1. AWorkbook - A reference of type Workbook
' 2. ModuleName - Name of a module in AWorkbook
' 3. RoutineName - The string name of the sub/function
'
' RETURNED VALUE
' Returns the requested routine documentation
Public Function GetRoutineDocumentation(AWorkbook As Workbook, _
                                        ModuleName As String, _
                                        RoutineName As String) As Variant
    Dim FirstLine As Long
    Dim DeclarationLine As Long
    Dim TheRoutineName As String
    Dim CodeModule As VBIDE.CodeModule
    
    ' Set default return value in case of error
    Let GetRoutineDocumentation = Null
    
    ' Set location for error handler
    On Error GoTo ErrorHandler
    
    ' Set reference to appropriate code module
    Set CodeModule = AWorkbook.VBProject.VBComponents(ModuleName).CodeModule
    
    ' Create routine name
    Let TheRoutineName = MakeRoutineName(AWorkbook, ModuleName, RoutineName)
    
    If FunctionExistsQ(AWorkbook, ModuleName, RoutineName) Then
        Let FirstLine = CodeModule.ProcStartLine(RoutineName, vbext_pk_Proc)
        Let DeclarationLine = CodeModule.ProcBodyLine(RoutineName, vbext_pk_Proc)

        ' Split by vbCfLf whatever is stored in the lines above the routine's declaration
        ' Get rid of the "' " at the beginning of each line.
        Let GetRoutineDocumentation = _
            Map(Lambda("ByVal s As String", _
                       "If len(s)>2 Then" & vbCrLf & _
                       "    Let s = Right(s, len(s)-2)" & vbCrLf & _
                       "Else" & vbCrLf & _
                       "    Let s = vbNullString" & vbCrLf & _
                       "End If", _
                       "s").FunctionName, _
                Split(CodeModule.Lines(FirstLine, DeclarationLine - FirstLine), vbCrLf))
        Let GetRoutineDocumentation = Join(GetRoutineDocumentation, vbCrLf)
    End If

    ' Through away everything before the first non-newline character
    While (Len(GetRoutineDocumentation) > 0 And _
          (Left(GetRoutineDocumentation, 1) = vbCrLf Or Left(GetRoutineDocumentation, 1) = vbCr Or _
           Left(GetRoutineDocumentation, 1) = vbLf))
        Let GetRoutineDocumentation = Right(GetRoutineDocumentation, Len(GetRoutineDocumentation) - 1)
    Wend
    
    Exit Function
    
ErrorHandler:
    Let GetRoutineDocumentation = Null
End Function

' DESCRIPTION
' Returns an array with the names of all of the routines defined in the given
' module and workbook.
'
' PARAMETERS
' 1. AWorkbook - A reference of type Workbook
' 2. ModuleName - Name of a module in AWorkbook
'
' RETURNED VALUE
' Returns list of routine defined in the given workbook and module
Public Function GetRoutineNames(AWorkbook As Workbook, _
                                ModuleName As String) As Variant
    Dim CodeModule As VBIDE.CodeModule
    Dim i As Long
    Dim RoutineName As String
    Dim ADict As Dictionary
    
    ' Initialize a dictionary to stored the routine names
    Set ADict = New Dictionary
    
    ' Set location of error handler
    On Error GoTo ErrorHandler
    
    ' Find the code module for the project.
    Set CodeModule = AWorkbook.VBProject.VBComponents(ModuleName).CodeModule

    ' Scan through the code module, looking for procedures.
    Let i = 1
    Do While i < CodeModule.CountOfLines
        ' Get the routine name corresponding to the current line
        Let RoutineName = CodeModule.ProcOfLine(i, vbext_pk_Proc)
        
        ' If the current line contains a declaration, store the routine's name
        ' if not already in the dictionary
        If RoutineName <> vbNullString Then
            If Not ADict.Exists(Key:=RoutineName) Then
                Call ADict.Add(Key:=RoutineName, Item:=Empty)
            End If
            
            ' Move to the line below the current procedure so we may find the
            ' next procedure
            Let i = i + CodeModule.ProcCountLines(RoutineName, vbext_pk_Proc)
        Else
            ' Move to the next line in this code module
            Let i = i + 1
        End If
    Loop
    
    ' Return the list of routine names stored as keys in the dictionary
    Let GetRoutineNames = ADict.Keys
    
    Exit Function
    
ErrorHandler:
    Let GetRoutineNames = Null
End Function

' DESCRIPTION
' Returns an array with all the modules in the current workbook
'
' PARAMETERS
' 1. AWorkbook - A reference of type Workbook
'
' RETURNED VALUE
' An array with the names of all modules in the workbook
Public Function GetModuleNames(AWorkbook As Workbook) As Variant
    Dim i As Long
    Dim ADict As Dictionary
    
    ' Initialize a dictionary to stored the routine names
    Set ADict = New Dictionary
    
    ' Set location of error handler
    On Error GoTo ErrorHandler
    
    ' Loop through all the components in the workbook, storing their names
    Let i = 1
    Do While i <= AWorkbook.VBProject.VBComponents.Count
        If AWorkbook.VBProject.VBComponents(i).Name <> "License" And _
           AWorkbook.VBProject.VBComponents(i).Name <> "LibraryTesting" And _
           Not ADict.Exists(Key:=AWorkbook.VBProject.VBComponents(i).Name) Then
            Call ADict.Add(Key:=AWorkbook.VBProject.VBComponents(i).Name, _
                           Item:=Empty)
        End If
        
        ' Move to the next component
        Let i = i + 1
    Loop
    
    ' Return the module names, which are stored as keys in the dictionary
    Let GetModuleNames = ADict.Keys

    Exit Function
    
ErrorHandler:
    Let GetModuleNames = Null
End Function

' DESCRIPTION
' Returns the declaration of the procedure with the given procedure
'
' PARAMETERS
' 1. AWorkbook - A reference of type Workbook
' 2. ModuleName - Name of a module in AWorkbook
' 3. RoutineName - The string name of the sub/function
'
' RETURNED VALUE
' Returns the requested procedure declaration
Public Function GetRoutineDeclaration(AWorkbook As Workbook, _
                                      ModuleName As String, _
                                      RoutineName As String) As Variant
    Dim ALine As String
    Dim Declaration As String
    Dim DeclarationLine As Long
    Dim TheRoutineName As String
    Dim TheCodeModule As VBIDE.CodeModule
    
    ' Set default return value in case of error
    Let GetRoutineDeclaration = Null
    
    ' ErrorCheck: Exit with Null if the requested routine does not exist
    If Not FunctionExistsQ(AWorkbook, ModuleName, RoutineName) Then Exit Function

    ' Set location of error handler
    On Error GoTo ErrorHandler
    
    ' Set reference to code module holding the routine
    Set TheCodeModule = AWorkbook.VBProject.VBComponents(ModuleName).CodeModule
    
    ' Create rotuine name
    Let TheRoutineName = MakeRoutineName(AWorkbook, ModuleName, RoutineName)
    Let DeclarationLine = TheCodeModule.ProcBodyLine(RoutineName, vbext_pk_Proc)
    
    ' Starting from DeclarationLine, read one line at a timme. Concatenate these
    ' lines until you find a line with no " _" as the last character. Declarations
    ' are by definition one liners that maay be split across multiple lines by using
    ' the " -" at the end of each physical code line.
    Let ALine = TheCodeModule.Lines(DeclarationLine, 1)
    Do While Right(ALine, 1) = "_"
        Let ALine = Left(ALine, Len(ALine) - 1) & " "
        Let Declaration = Declaration & ALine
        Let DeclarationLine = DeclarationLine + 1
        Let ALine = TheCodeModule.Lines(DeclarationLine, 1)
    Loop
    
    Let GetRoutineDeclaration = RemoveDuplicatedSpaces(Declaration & ALine)
    
    Exit Function
    
ErrorHandler:
    Let GetRoutineDeclaration = Null
End Function

' DESCRIPTION
' Returns the documentation at the top of a code module
'
' PARAMETERS
' 1. AWorkbook - A reference of type Workbook
' 2. ModuleName - Name of a module in AWorkbook
'
' RETURNED VALUE
' Returns the requested procedure declaration
Public Function GetModuleDocumentation(AWorkbook As Workbook, ModuleName As String) As Variant
    Dim ModuleDocumentation As String
    Dim ALine As String
    Dim LineCounter As Long
    Dim TheCodeModule As VBIDE.CodeModule
    
    ' Set default return value in case of error
    Let GetModuleDocumentation = Null
    
    ' Set location of error handler
    On Error GoTo ErrorHandler
    
    ' Set reference to code module holding the routine
    Set TheCodeModule = AWorkbook.VBProject.VBComponents(ModuleName).CodeModule
        
    ' Starting from the first line of the module, proceed to the first comment line
    ' (e.g. lines with "'" as the first character.
    Let LineCounter = 1
    Let ALine = TheCodeModule.Lines(LineCounter, 1)
    Do While Left(ALine, 1) <> "'"
        Let ALine = ALine + 1
    Loop
    
    ' Collect all lines that start with "' "
    If Len(ALine) > 2 Then
        Let ModuleDocumentation = Right(ALine, Len(ALine) - 2)
    Else
        Let ModuleDocumentation = vbNullString
    End If
    
    Let LineCounter = LineCounter + 1
    
    Let ALine = TheCodeModule.Lines(LineCounter, 1)
    Do While Left(ALine, 1) = "'"
        If Len(ALine) = 1 Then
            Let ModuleDocumentation = ModuleDocumentation + vbCrLf
        Else
            Let ModuleDocumentation = ModuleDocumentation & Right(ALine, Len(ALine) - 2) & vbCrLf
        End If
        Let LineCounter = LineCounter + 1
        Let ALine = TheCodeModule.Lines(LineCounter, 1)
    Loop
    
    Let GetModuleDocumentation = ModuleDocumentation
    
    Exit Function
    
ErrorHandler:
    Let GetModuleDocumentation = Null
End Function

' DESCRIPTION
' Returns name, full-path, and description for each non-built-in
' references used by the given workbook.
'
' PARAMETERS
' 1. Wbk - A reference of type Workbook
'
' RETURNED VALUE
' A dictionary of dictionaries. This dictionary's keys are the
' reference names. Its items are dictionaries containing full paths
' and descriptions. Each inner dictionary has two keys: "FullPath"
' and "Description"
Public Function GetReferences(Wbk As Workbook) As Dictionary
    Dim AReference As Reference
    Dim ADict As Dictionary
    Dim RowDict As Dictionary

    ' Instantiate dictionary to hold reference data
    Set ADict = New Dictionary
    
    ' Error Check: Exit with Null if project has no references
    ' This is virtually impossible since all workbooks contain
    ' some default references by default
    If Wbk.VBProject.References.Count = 0 Then Exit Function

    ' Pre-allocate array to hold references
    ReDim ReferenceNames(1 To Wbk.VBProject.References.Count, 1 To 3)
    
    ' For each reference get its name, full path, and description
    For Each AReference In Wbk.VBProject.References
        If Not (AReference.BuiltIn Or ADict.Exists(Key:=AReference.Name)) Then
            ' Instantiate dictionary to store this reference's data
            Set RowDict = New Dictionary
            
            ' Store this reference's data
            Call RowDict.Add(Key:="FullPath", Item:=AReference.FullPath)
            Call RowDict.Add(Key:="Description", Item:=AReference.Description)
            
            ' Add this reference's dictionary to outer dictionary
            Call ADict.Add(Key:=AReference.Name, Item:=RowDict)
        End If
    Next
    
    ' Return array with
    Set GetReferences = ADict
End Function


' DESCRIPTION
' Generates Org-formatted documentation for a workbook
'
' PARAMETERS
' 1. Wbk - A reference of type Workbook
' 2. FullPathFileName - Filename for Org file (full path)
'
' RETURNED VALUE
' Produces the requested documentation file
Public Sub GenerateOrgDocumentation(Wbk As Workbook, _
                                    FullPathFileName As String, _
                                    AuthorName As String, _
                                    AuthorEmail As String)
    Dim ModuleNames As Variant
    Dim AModuleName As Variant
    Dim RoutineNames As Variant
    Dim ARoutineName As Variant
    Dim RefsDict As Dictionary
    Dim RefDict As Dictionary
    Dim i As Integer
    
    On Error GoTo ErrorHandler
    
    Set RefsDict = GetReferences(ThisWorkbook)
    
    Open FullPathFileName For Output As #1
    
    Print #1, "#+TITLE: Functional Documentation for " & Wbk.Name
    Print #1, "#+AUTHOR: " & AuthorName
    Print #1, "#+DATE: " & Format(Now, "YYYYMMDD")
    Print #1, "#+EMAIL: " & AuthorEmail
    Print #1, "#+INFOJS_OPT: view:info" & vbNewLine

    Print #1, "* References"
    Print #1, "This workbook has " & RefsDict.Count & " non-built-in references."
    For i = 0 To RefsDict.Count - 1
        Set RefDict = RefsDict.Items(i)

        Print #1, "** Description: " & RefDict.Item(Key:="Description")
        Print #1, "   - Name: " & RefsDict.Keys(i)
        Print #1, "   - Full Path: " & RefDict.Item(Key:="FullPath")
    Next
    
    Let Application.StatusBar = "References Done"
    
    For Each AModuleName In GetModuleNames(Wbk)
        Let Application.StatusBar = "Working on module " & AModuleName
        
        Print #1, "* " & AModuleName
        Print #1, GetModuleDocumentation(Wbk, CStr(AModuleName))
        
        For Each ARoutineName In GetRoutineNames(Wbk, CStr(AModuleName))
            Let Application.StatusBar = "Working on routine " & ARoutineName
        
            Print #1, "** " & ARoutineName
            Print #1, GetRoutineDeclaration(Wbk, CStr(AModuleName), _
                                            CStr(ARoutineName))
            Print #1, "#+BEGIN_EXAMPLE"
            Print #1, GetRoutineDocumentation(Wbk, CStr(AModuleName), _
                                              CStr(ARoutineName))
            Print #1, "#+END_EXAMPLE"
        Next
        
        Let Application.StatusBar = "Done with module " & AModuleName
    Next
    
    Close #1
    
    Let Application.StatusBar = "Ready"
    
    Exit Sub
    
ErrorHandler:
    Let Application.StatusBar = "Ready"
    Call MsgBox("There was an error creating the documentation file.")
End Sub
