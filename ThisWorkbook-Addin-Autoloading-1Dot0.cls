VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' PURPOSE OF THIS WORKSHEET MODULE
'
' The purpose of this worksheet is hold routines that run when
' the workbook opens and closes.

Option Explicit
Option Base 1

' DESCRIPTION
' Returns the XLAM addin library name to load on open and unload before close.
'
' PARAMETERS
' None
'
' RETURNED VALUE
' The requested string
Public Function ExcelLibraryName() As String
    Let ExcelLibraryName = "ExcelLibrary9Dot4"
End Function

' DESCRIPTION
' Returns the file name for the XLAM addin to load on open and unload before close.
'
' PARAMETERS
' None
'
' RETURNED VALUE
' The requested string
Public Function ExcelLibraryFilename() As String
    Let ExcelLibraryFilename = "ExcelLibraryV9.4.xlam"
End Function

' DESCRIPTION
' Pulls all files in the current directory and subdirectories and writes them
' to Sheet1.
'
' PARAMETERS
' None
'
' RETURNED VALUE
' The list of filenames in Sheet1.
Public Sub ListAllFiles()
    Dim NextRow As Long
    Dim file As Variant

    Call Sheet1.UsedRange.ClearContents

    Let Sheet1.Range("A1:G1").Value = Array("Name", "Size", "Type", "Created", "Accessed", "Modified", "Path")
    
    Let NextRow = 2
    With CreateObject("Scripting.FileSystemObject")
        For Each file In Filter(Split(CreateObject("WScript.Shell").Exec("CMD /C DIR ""C:\Users\pablo.perez.GLOBAL\GlobalCapital Malta\GC Board - Documents\Board Packs Through Jan 2018\*.*"" /S /B /A:-D").StdOut.ReadAll, vbCrLf), ".")
            With .GetFile(CStr(file))
                Let Sheet1.Range("A" & NextRow).Value = .Name
                Let Sheet1.Range("B" & NextRow).Value = Format((.Size / 1024), "000") & " KB"
                Let Sheet1.Range("C" & NextRow).Value = .Type
                Let Sheet1.Range("D" & NextRow).Value = .DateCreated
                Let Sheet1.Range("E" & NextRow).Value = .DateLastAccessed
                Let Sheet1.Range("F" & NextRow).Value = .DateLastModified
                Let Sheet1.Range("G" & NextRow).Value = .Path
            End With
            
            Let NextRow = NextRow + 1
        Next
    End With
End Sub

' DESCRIPTION
' Removes the reference to the ExcelLibrary XLAM addin prior to closing.
'
' PARAMETERS
' None
'
' RETURNED VALUE
' None
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Call RemoveXlamAddinReference(ExcelLibraryFilename)
    Call ThisWorkbook.Save
End Sub

' DESCRIPTION
' Adds reference to the ExcelLibrary XLAM addin.
'
' PARAMETERS
' None
'
' RETURNED VALUE
' None
Private Sub Workbook_Open()
    Call AddXlamAddinReference(ThisWorkbook.Path & "\..\Common\" & ExcelLibraryFilename, _
                               ExcelLibraryName)
End Sub

' DESCRIPTION
' Checks if an XLAM addin is referenced by ThisWorkbook
'
' PARAMETERS
' 1. VBAProjectReferenceName - Addin's VBAProject name
'
' RETURNED VALUE
' True if the reference loaded. False if the reference not loaded.
Public Function XlamAddinLoadedQ(VBAProjectReferenceName As String) As Boolean
    Dim var As Variant
    
    Let XlamAddinLoadedQ = True
    
    On Error GoTo ErrorHandler
    Let var = ThisWorkbook.VBProject.References(VBAProjectReferenceName).Name
    
    Exit Function
ErrorHandler:
    Let XlamAddinLoadedQ = False
End Function

' DESCRIPTION
' Removes the XLAM addin with the given name.
'
' PARAMETERS
' 1. VBAProjectReferenceName - Addin's VBAProject name
'
' RETURNED VALUE
' None
Public Sub RemoveXlamAddinReference(VBAProjectReferenceName As String)
    If XlamAddinLoadedQ(VBAProjectReferenceName) Then
        With ThisWorkbook.VBProject
            Call .References.Remove(.References(VBAProjectReferenceName))
        End With
    End If
End Sub

' DESCRIPTION
' Returns a 2D array with the data contained in a dictionary of dictionaries, where
' each inner dictionary has the same number of fields. All of the inner dictionaries
' must contain the same number of items.
'
' At the moment, this function does not perform detailed error checking.
'
' PARAMETERS
' 1. FullPathToXalmAddin - Full path to the XALM addin's file
' 2. VBAProjectReferenceName - Addin's VBAProject name
'
' RETURNED VALUE
' None
Public Sub AddXlamAddinReference(FullPathToXalmAddin As String, VBAProjectReferenceName As String)
    If Not XlamAddinLoadedQ(FullPathToXalmAddin) Then
        Call ThisWorkbook.VBProject.References.AddFromFile(FullPathToXalmAddin)
    End If
End Sub
