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
    Let ExcelLibraryName = "ExcelLibrary9Dot5"
End Function

' DESCRIPTION
' Returns the file name for the XLAM addin to load on open and unload before close.
'
' PARAMETERS
' None
'
' RETURNED VALUE
' The requested string
Public Function ExcelLibraryFileName() As String
    Let ExcelLibraryFileName = ExcelLibraryName() & ".xlam"
End Function

' DESCRIPTION
' Removes the reference to the ExcelLibrary XLAM addin prior to closing.
'
' PARAMETERS
' None
'
' RETURNED VALUE
' None
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Call RemoveXlamAddinReference(ExcelLibraryName)
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
    ' Load Excel Library
    Call AddXlamAddinReference(ThisWorkbook.Path & "\..\Common\" & ExcelLibraryFileName, _
                               ExcelLibraryName)
                               
    ' Clear the body of the list object holding the list of filenames.
    If Not FileListWsht.ListObjects(1).DataBodyRange Is Nothing Then
        Call FileListWsht.ListObjects(1).DataBodyRange.Delete
    End If
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
Public Sub AddXlamAddinReference(FullPathToXalmAddin As String, ExcelAddInReferenceName As String)
    If Not XlamAddinLoadedQ(ExcelAddInReferenceName) Then
        Call ThisWorkbook.VBProject.References.AddFromFile(FullPathToXalmAddin)
    End If
End Sub
