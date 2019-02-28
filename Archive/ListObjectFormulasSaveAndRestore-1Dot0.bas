Attribute VB_Name = "ListObjectFormulasSaveAndRestore"
Option Explicit
Option Base 1

' DESCRIPTION
' This sub applies to the listobject holding the given selected cell.
' This function restores the formulas stored in the comments of the listobject's headers row.
' Columns with no comments are ignored.
'
' PARAMETERS
' None
'
' RETURNED VALUE
' No value is returned
Public Sub RetoreTableFormulas()
    Dim c As Integer
    Dim lo As ListObject
    
    ' Set a reference to the listobject holding the selected cell
    Set lo = Selection.ListObject
    
    ' Extract formula for each column header and insert it in the corresponding columns
    For c = 1 To lo.HeaderRowRange.Count
        If Not (lo.HeaderRowRange(1, c).Comment Is Nothing) Then
            Let lo.ListColumns(c).DataBodyRange.Formula = lo.HeaderRowRange(1, c).Comment.Text
        End If
    Next
End Sub

' DESCRIPTION
' This sub applies to the listobject holding the given selected cell.
' Stores in the comments of the headers row the formulas in the first listrow.
' Columns without formulas are ignored.
'
' Note that this function wipes all pre-exisiting comments in the headers row.
'
' PARAMETERS
' None
'
' RETURNED VALUE
' No value is returned
Public Sub SaveTableFormulas()
    Dim c As Integer
    Dim lo As ListObject

    ' Set a reference to the listobject holding the selected cell
    Set lo = Selection.ListObject

    ' Store formulas
    For c = 1 To lo.HeaderRowRange.Count
        ' Clear any previously stored comments
        Call lo.HeaderRowRange(1, c).ClearComments
        
        ' Store the formulas of every column but Date and ID_BB_GLOBAL
        If lo.DataBodyRange(1, c).Formula <> vbNullString Then
            If Left(Trim(lo.DataBodyRange(1, c).Formula), 1) = "=" Then
                Call lo.HeaderRowRange(1, c).AddComment(lo.DataBodyRange(1, c).Formula)
            End If
        End If
    Next
    
    Let lo.DataBodyRange.Value2 = lo.DataBodyRange.Value2
End Sub


