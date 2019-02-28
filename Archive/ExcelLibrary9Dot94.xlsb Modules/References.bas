Attribute VB_Name = "References"
' PURPOSE OF THIS MODULE
'
' The purpose of this module is to provide faciities to handle XLAM addin
' references programmatically.

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
    Call AddXlamAddinReference(ThisWorkbook.Path & "\..\Common\" & ExcelLibraryFileName, _
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
' 2. ExcelAddInReferenceName - Addin's VBAProject name
'
' RETURNED VALUE
' None
Public Sub AddXlamAddinReference(FullPathToXalmAddin As String, ExcelAddInReferenceName As String)
    If Not XlamAddinLoadedQ(ExcelAddInReferenceName) Then
        Call ThisWorkbook.VBProject.References.AddFromFile(FullPathToXalmAddin)
    End If
End Sub


