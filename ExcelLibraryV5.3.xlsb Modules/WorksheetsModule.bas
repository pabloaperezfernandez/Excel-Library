Attribute VB_Name = "WorksheetsModule"
Option Explicit
Option Base 1

' This function returns the row number of the first cell from the bottom to the top that is not empty (e.g. isempty() is false)
' and not equal to the given ComparisonScalars (if a scalar) or any of the scalars in 1D array ComparisonScalars.  If the
' number of dimensions of ComparisonScalars is > 1, the function returns -1
Public Function LastNonBlankRowIndexInColumnNotEqualTo(aColumnRange As Range, ComparisonScalars As Variant) As Long
    Dim LastRowIndex As Long
    Dim i As Long
    Dim j As Integer
    
    ' Exit if ComparisonScalars has dimensions > 1
    If NumberOfDimensions(ComparisonScalars) > 1 Then
        Let LastNonBlankRowIndexInColumnNotEqualTo = -1
        
        Exit Function
    End If
    
    ' Get the last non-emty cell in the worksheet's column holding the given column range
    Let LastRowIndex = aColumnRange.Worksheet.Cells(aColumnRange.Worksheet.Rows.Count, aColumnRange.Column).End(xlUp).row

    For i = LastRowIndex To 1 Step -1
        If Not IsArray(ComparisonScalars) Then
            If Not Application.IsNA(aColumnRange.Worksheet.Cells(i, aColumnRange.Column).Value2) Then
                If aColumnRange.Worksheet.Cells(i, aColumnRange.Column).Value2 <> ComparisonScalars Then
                    Exit For
                End If
            End If
        Else
            If FreeQ(ComparisonScalars, aColumnRange.Worksheet.Cells(i, aColumnRange.Column).Value2) Then
                Exit For
            End If
        End If
    Next i

    
    Let LastNonBlankRowIndexInColumnNotEqualTo = i
End Function

' This function returns a an 2D array with consolidated data contained in DB-like tables found in an
' the worksheets referenced by the given worksheet reference array.  All of the referenced worksheets
' must contain identical DB-like tables, including the headers. All worksheets must contain headers in the
' first row, unless an optional starting row is provided.  If StartingRow is provided, the function
' includes the first StartingRow-1 rows as headers in the consolidated array.  If StartingRow is not given,
' StartingRow is set to 1.
'
' The data starting in row StartingRow from each workbook is consolidated in the consolidated array.
' The worksheets referenced by array WorksheetsArray() may be in different workbooks.
'
' If the optional parameter TargetWorksheet is supplied, then the function returns Null
' and copies the consolidated data into the target worksheet starting in cell A1.
Public Function ConsolidateWorksheets(WorksheetsArray() As Worksheet, Optional StartingRow As Variant, Optional TargetWorksheet As Variant) As Variant
    Dim Headers As Variant
    Dim NumberOfColumns As Integer
    Dim NumberOfRows As Long
    Dim FirstDataRow As Long
    Dim LastDataRow As Long
    Dim ConsolidationWorksheet As Worksheet
    Dim SourceRange As Range
    Dim TargetRange As Range
    Dim N As Long
    Dim RowCursor As Long
    
    ' Set the first row (where data starts, not headers)
    If IsMissing(StartingRow) Then
        Let FirstDataRow = 1
    ElseIf StartingRow >= 1 Then
        Let FirstDataRow = StartingRow
    Else
        Let FirstDataRow = 1
    End If
    
    ' Instantiate the consolidated worksheet.  It will be deleted at the end once data
    ' is copied to the target worksheet or returned as a 2D array.
    Set ConsolidationWorksheet = ThisWorkbook.Worksheets.Add
    
    ' Determine the number of columns
    Let NumberOfColumns = WorksheetsArray(1).UsedRange.Columns.Count
    
    ' Get the headers. May be multiple rows, starting from the top
    If FirstDataRow > 1 Then
        Let Headers = WorksheetsArray(1).Range("A1").Resize(FirstDataRow, NumberOfColumns).Value2
    
        ' Copy the headers row to ConsolidatedWorksheet
        Call DumpInTempPositionWithoutFirstClearing(Headers, ConsolidationWorksheet.Range("A1"))
    End If
    
    Let RowCursor = FirstDataRow
    For N = LBound(WorksheetsArray) To UBound(WorksheetsArray)
        ' Compute the last row to consolidate in this worksheet
        Let LastDataRow = WorksheetsArray(N).Cells(WorksheetsArray(1).Rows.Count, 1).End(xlUp).row
        
        ' Compute the number of rows to consolidate
        Let NumberOfRows = LastDataRow - FirstDataRow + 1
                        
        ' Set the source range
        Set SourceRange = WorksheetsArray(N).Cells(FirstDataRow, 1).Resize(NumberOfRows, NumberOfColumns)
        
        ' Set the target range
        Set TargetRange = ConsolidationWorksheet.Cells(RowCursor, 1).Resize(NumberOfRows, NumberOfColumns)
        
        ' Do the actual consolidation
        Call SourceRange.Copy
        Call TargetRange.PasteSpecial(Paste:=xlPasteAll)
        
        ' Update RowCursor
        Let RowCursor = RowCursor + NumberOfRows
    Next N
    
    If Not IsMissing(TargetWorksheet) Then
        Call ConsolidationWorksheet.Cells(1, 1).Resize(RowCursor, NumberOfColumns).Copy
        Call TargetWorksheet.Range("A1").PasteSpecial(Paste:=xlPasteAll)
        
        ' Delete the consolidation worksheet
        Call ConsolidationWorksheet.Delete
        
        Exit Function
    End If
    
    ' Since a target worksheet was not provided to use as consolidator, we simply return a 2D matrix with the data
    Let ConsolidateWorksheets = ConsolidationWorksheet.Cells(1, 1).Resize(RowCursor, NumberOfColumns).Value2
    
    ' Delete the consolidation worksheet
    Call ConsolidationWorksheet.Delete
End Function

Public Function ConsolidateWorksheetsHorizontally(WorksheetsArray() As Worksheet, Optional StartingColumn As Variant, Optional TargetWorksheet As Variant) As Variant
    Dim var As Variant
    Dim NRows As Variant
    Dim NCols As Variant
    Dim SourceRange As Range
    Dim TargetRange As Range
    Dim LastColUsed As Long
    Dim ConsolidationWorksheet As Worksheet
    
    Set ConsolidationWorksheet = ThisWorkbook.Worksheets.Add
    
    Let LastColUsed = 0
    For Each var In WorksheetsArray
        If LastColUsed = 0 Then
            ' This is the first workbook. Copy names and tickers
            Let LastColUsed = LastColUsed + 1
            Set SourceRange = var.Worksheets(1).Range("A1").CurrentRegion
            Set TargetRange = ConsolidationWorksheet.Range("A1").Resize(SourceRange.Rows.Count, SourceRange.Columns.Count)
            Call SourceRange.Copy
            Call TargetRange.PasteSpecial(Paste:=xlPasteAll)
            Let LastColUsed = LastColUsed + SourceRange.Columns.Count
        Else
            Set SourceRange = var.Worksheets(1).Range("A1").CurrentRegion
            Set SourceRange = SourceRange.Offset(0, StartingColumn - 1).Resize(SourceRange.Rows.Count, SourceRange.Columns.Count - 2)
            Set TargetRange = ConsolidationWorksheet.Range("A1").Offset(0, LastColUsed - 1).Resize(SourceRange.Rows.Count, SourceRange.Columns.Count)
            Call SourceRange.Copy
            Call TargetRange.PasteSpecial(Paste:=xlPasteAll)
            Let LastColUsed = LastColUsed + SourceRange.Columns.Count
        End If
    Next var
    
    ' Autofit all columns and then center all columns to the right of StartingColumn
    Set TargetRange = ConsolidationWorksheet.Range("A1").CurrentRegion
    Call TargetRange.EntireColumn.AutoFit
    Set TargetRange = TargetRange.Offset(0, 2).Resize(TargetRange.Rows.Count, TargetRange.Columns.Count - 2)
    Let TargetRange.EntireColumn.HorizontalAlignment = xlCenter
    
    If IsMissing(TargetWorksheet) Then
        Let ConsolidateWorksheetsHorizontally = ConsolidationWorksheet.Range("A1").CurrentRegion.Value2
        Call ConsolidationWorksheet.Delete
    
        Exit Function
    End If
    
    Call ConsolidationWorksheet.Cells(1, 1).Resize(RowCursor, NumberOfColumns).Copy
    Call TargetWorksheet.Range("A1").PasteSpecial(Paste:=xlPasteAll)
    
    ' Delete the consolidation worksheet
    Call ConsolidationWorksheet.Delete
End Function

' This function deletes all worksheets in the worksheet's parent with the exception of the given
' worksheet.
Public Sub RemoveAllOtherWorksheets(TheWorksheet As Worksheet)
    Dim WorkbookRef As Workbook
    Dim WorksheetNames() As String
    Dim i As Integer

    ' Set a reference to the workbook
    Set WorkbookRef = TheWorksheet.Parent

    ' Delete any other worksheets beside this one
    ReDim WorksheetNames(WorkbookRef.Worksheets.Count)
    For i = 1 To WorkbookRef.Worksheets.Count
        Let WorksheetNames(i) = WorkbookRef.Worksheets(i).Name
    Next i
    
    ' Delete all worksheets other than the template worksheet just copied
    For i = 1 To WorkbookRef.Worksheets.Count
        If WorksheetNames(i) <> TheWorksheet.Name Then
            Call WorkbookRef.Worksheets(WorksheetNames(i)).Delete
        End If
    Next i
End Sub
