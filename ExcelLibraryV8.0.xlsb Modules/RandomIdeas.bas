Attribute VB_Name = "RandomIdeas"
' This module constains random test code for various things

Public Sub TestAddNewModule()
  Dim proj As VBIDE.VBProject
  Dim comp As VBIDE.VBComponent
  Dim codeMode As VBIDE.CodeModule
  Dim lineNum As Integer

  Set proj = ActiveWorkbook.VBProject
  Set comp = proj.VBComponents.Add(vbext_ct_StdModule)
  comp.Name = "MyNewModule"

  Set codeMode = comp.CodeModule

  With codeMode
    lineNum = .CountOfLines + 1
    .InsertLines lineNum, vbCrLf
    lineNum = lineNum + 1
    .InsertLines lineNum, vbCrLf
    lineNum = lineNum + 1
    .InsertLines lineNum, "Public Sub ANewSub()"
    lineNum = lineNum + 1
    .InsertLines lineNum, "  MsgBox " & """" & "I added a module!" & """"
    lineNum = lineNum + 1
    .InsertLines lineNum, "End Sub"
  End With

    Debug.Print "ANewSub starts in line " & codeMode.ProcStartLine("ANewSub", vbext_pk_Proc)
End Sub

'Make sure to insert this in module ANewModule before running the code
'
'
'
'
'Public Sub ANewSub()
'  MsgBox "Sub1"
'End Sub
'
'Public Sub ANewSub2()
'  MsgBox "Sub2"
'End Sub
'
'Public Sub ANewSub3()
'  MsgBox "Sub3"
'End Sub
'
' a Line
' Another Line
'
'Public Sub ANewSub4()
'  MsgBox "Sub4"
'  MsgBox "Sub4 again"
'End Sub
'
'Public Sub ANewSub5()
'  MsgBox "Sub5"
'  MsgBox "Sub5 lines 2"
'End Sub
Public Sub TestAddNewModule2()
    Dim proj As VBIDE.VBProject
    Dim comp As VBIDE.VBComponent
    Dim codeMode As VBIDE.CodeModule
    Dim lineNum As Integer
    
    Set proj = ActiveWorkbook.VBProject
    Set comp = proj.VBComponents("MyNewModule")
    Set codeMode = comp.CodeModule

    Debug.Print "ANewSub starts in line " & codeMode.ProcStartLine("ANewSub", vbext_pk_Proc)
    Debug.Print "ANewSub2 starts in line " & codeMode.ProcStartLine("ANewSub2", vbext_pk_Proc)
    Debug.Print "ANewSub3 starts in line " & codeMode.ProcStartLine("ANewSub3", vbext_pk_Proc)
    Debug.Print "ANewSub4 starts in line " & codeMode.ProcStartLine("ANewSub4", vbext_pk_Proc)
    
    Debug.Print "Lines 13 and 14 are: " & codeMode.Lines(13, 2)
    Debug.Print "The body of ANewSub4 begins on " & codeMode.ProcBodyLine("ANewSub4", vbext_pk_Proc)
    Debug.Print "ANewSub4 has # lines = " & codeMode.ProcCountLines("ANewSub4", vbext_pk_Proc)
    Debug.Print "ANewSub3 has # lines = " & codeMode.ProcCountLines("ANewSub3", vbext_pk_Proc)
    
    Debug.Print "---------------------"
    Debug.Print codeMode.Lines(codeMode.ProcStartLine("ANewSub4", vbext_pk_Proc), _
                               codeMode.ProcCountLines("ANewSub4", vbext_pk_Proc))
    Debug.Print "---------------------"
    
    Call codeMode.DeleteLines(codeMode.ProcStartLine("ANewSub4", vbext_pk_Proc), _
                              codeMode.ProcCountLines("ANewSub4", vbext_pk_Proc))
                              
    Call codeMode.AddFromString("Public Function Func1(x) as Integer" & vbCrLf & _
                                "   debug.print ""Done""" & vbCrLf & _
                                " END FUNCTion")
    On Error GoTo ErrorHandler
    Debug.Print IsError(codeMode.ProcStartLine("ANewSub4", vbext_pk_Proc))
    
    Debug.Print "The code had ended."
    Exit Sub

ErrorHandler:
    Debug.Print "ANewSub4 does not exit. The error is " & Err.Number
    Debug.Print "The error name is """ & Err.Description & """"
End Sub

' Testing my code to insert and delete functions
Public Sub ATest()
    Call InsertFunction(ThisWorkbook, _
                        "MyNewModule", _
                        "ANewFunc", _
                        Array("x", "y"), _
                        Array("debug.print x", "debug.print y") _
                        )

    Call InsertFunction(ThisWorkbook, _
                        "MyNewModule", _
                        "ANewFunc2", _
                        ToStrings([{"x","y"}]), _
                        ToStrings([{"debug.print x", "debug.print y"}]) _
                        )
                        
    Call UseTheFuncs
    
    Call DeleteFunction(ThisWorkbook, "MyNewModule", "ANewFunc")
    Call DeleteFunction(ThisWorkbook, "MyNewModule", "ANewFunc2")
End Sub

Private Sub UseTheFuncs()
    Debug.Print "Calling ANewFunc"
    Call ANewFunc(1, 2)
    Debug.Print
    Debug.Print "Calling ANewFunc2"
    Call ANewFunc(10, 20)
End Sub
