Attribute VB_Name = "VbaCodeManipulation"
Option Base 1
Option Explicit

'Remember to add a reference to Microsoft Visual Basic for Applications Extensibility
'Exports all VBA project components containing code to a folder in the same directory as this spreadsheet.
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
                'Alternatively, you can prompt the user before killing the file.
                If Dir(fName, vbNormal) <> vbNullString Then Kill (fName)
                Call VBComp.Export(fName)
            End If
        End If
    Next VBComp
End Sub
