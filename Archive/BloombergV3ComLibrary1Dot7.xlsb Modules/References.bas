Attribute VB_Name = "References"
' PURPOSE OF THIS MODULE
'
' The purpose of this module is to hold the code used to autoload references
' at runtimme. These are called by code in module VBAProject's ThisWorkbook

Option Explicit
Option Base 1

Public Const EXCEL_LIBRARY_NAME As String = "ExcelLibrary9Dot997"
Public Const EXCEL_LIBRARY_FILENAME As String = "ExcelLibrary9Dot997.xlam"

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
    If Not XlamAddinLoadedQ(VBAProjectReferenceName) Then
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
    If Not XlamAddinLoadedQ(VBAProjectReferenceName) Then
        Call ThisWorkbook.VBProject.References.AddFromFile(FullPathToXalmAddin)
    End If
End Sub

