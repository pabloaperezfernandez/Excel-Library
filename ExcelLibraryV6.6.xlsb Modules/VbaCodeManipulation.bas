Attribute VB_Name = "VbaCodeManipulation"
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

' Remember to add a reference to Microsoft Visual Basic for Applications Extensibility
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
    If FileExistsQ(destDir) Then
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
                If dir(fName, vbNormal) <> vbNullString Then Kill (fName)
                Call VBComp.Export(fName)
            End If
        End If
    Next VBComp
End Sub


' The purpose of this function is to dump in the immediate console a
' function to draw this application's UI.  This function should be run
' right after "drawing" the UI by hand (which is easier).
'
' This function handles all form controls.  It does not handle the
' data for dropdowns, etc.  It handles the size, location, and macros
' associated with the controls.
Public Sub WriteUiCode(wsht As Worksheet)
    Dim aShape As Shape
    Dim aShapeName As Variant
    Dim ShapesDict As Dictionary
    Dim PropertiesDict As Dictionary
    
    Debug.Print wsht.Name & " has " & wsht.Shapes.Count & " shapes."
    
    Set ShapesDict = New Dictionary
    For Each aShape In wsht.Shapes
        Set PropertiesDict = New Dictionary
        
        Call PropertiesDict.Add("Top", aShape.Top)
        Call PropertiesDict.Add("Left", aShape.Left)
        Call PropertiesDict.Add("Width", aShape.Width)
        Call PropertiesDict.Add("Height", aShape.height)
        Call PropertiesDict.Add("AlternativeText", aShape.AlternativeText)
        Call PropertiesDict.Add("OnAction", aShape.OnAction)
        Call PropertiesDict.Add("XlFormControl", aShape.FormControlType)
        
        Call ShapesDict.Add(Key:=aShape.Name, Item:=PropertiesDict)
    Next

    Debug.Print "Public Sub ReCreateUi()"
    Debug.Print "   Dim aShape As Shape"
    Debug.Print
    Debug.Print "   For Each aShape In wsht.Shapes: Call aShape.Delete: Next"
    Debug.Print
    For Each aShapeName In ShapesDict.Keys
        Debug.Print "   Set aShape = wsht.Shapes.AddFormControl(" & _
                    ShapesDict.Item(aShapeName).Item("XlFormControl") & _
                    ", " & ShapesDict.Item(aShapeName).Item("Left") & ", " & _
                    ShapesDict.Item(aShapeName).Item("Top") & ", " & _
                    ShapesDict.Item(aShapeName).Item("Width") & ", " & _
                    ShapesDict.Item(aShapeName).Item("Height") & ")"
        Debug.Print "   Let aShape.Name = """ & aShapeName & """"
        Debug.Print "   Let aShape.TextFrame.Characters.Text = """ & _
                    ShapesDict.Item(aShapeName).Item("AlternativeText") & """"
        Debug.Print "   Let aShape.OnAction = """ & ShapesDict.Item(aShapeName).Item("OnAction") & """"
        Debug.Print
    Next
    Debug.Print "End Sub"
    
End Sub

