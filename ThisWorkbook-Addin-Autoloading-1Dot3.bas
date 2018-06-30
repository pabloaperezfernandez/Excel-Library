' PURPOSE OF THIS WORKSHEET MODULE
'
' The purpose of this worksheet is hold routines that run when
' the workbook opens and closes.

Option Explicit
Option Base 1

' DESCRIPTION
' Removes the reference to the ExcelLibrary XLAM addin prior to closing.
'
' PARAMETERS
' None
'
' RETURNED VALUE
' None
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Call RemoveXlamAddinReference(EXCEL_LIBRARY_NAME)
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
    Call AddXlamAddinReference(ThisWorkbook.Path & "\..\..\Common\" & EXCEL_LIBRARY_FILENAME, _
                               EXCEL_LIBRARY_NAME)
End Sub