Attribute VB_Name = "References"
' PURPOSE OF THIS MODULE
'
' The purpose of this module is to provide faciities to handle XLAM addin
' references programmatically.

Option Explicit
Option Base 1

' DESCRIPTION
' Checks if an XLAM addin is referenced by ThisWorkbook
'
' PARAMETERS
' 1. VBAProjectReferenceName - Addin's VBAProject name
' 2. Wbk - A reference to the workbook we are testing for the presence addin
'
' RETURNED VALUE
' True if the reference loaded. False if the reference not loaded.
Public Function XlamAddinLoadedQ(VBAProjectReferenceName As String, Wbk As Workbook) As Boolean
    Dim var As Variant
    
    On Error GoTo ErrorHandler
    Let var = Wbk.VBProject.References(VBAProjectReferenceName).Name
    
    Let XlamAddinLoadedQ = True
    
    Exit Function
ErrorHandler:
    Let XlamAddinLoadedQ = False
End Function

' DESCRIPTION
' Removes the XLAM addin with the given name.
'
' PARAMETERS
' 1. VBAProjectReferenceName - Addin's VBAProject name
' 2. Wbk - A reference to the workbook from which to remove the addin
'
' RETURNED VALUE
' None
Public Sub RemoveXlamAddinReference(VBAProjectReferenceName As String, Wbk As Workbook)
    If XlamAddinLoadedQ(VBAProjectReferenceName) Then
        Call Wbk.VBProject.References.Remove(Wbk.VBProject.References(VBAProjectReferenceName))
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
' 3. Wbk - A reference to the workbook to which the addin must be added
'
' RETURNED VALUE
' None
Public Sub AddXlamAddinReference(FullPathToXalmAddin As String, VBAProjectReferenceName As String, Wbk As Workbook)
    If Not XlamAddinLoadedQ(FullPathToXalmAddin, Wbk) Then
        Call Wbk.VBProject.References.AddFromFile(FullPathToXalmAddin)
    End If
End Sub

