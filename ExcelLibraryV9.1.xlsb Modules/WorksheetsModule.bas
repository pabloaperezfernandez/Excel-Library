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
    Let LastRowIndex = aColumnRange.Worksheet.Cells(aColumnRange.Worksheet.Rows.Count, aColumnRange.Column).End(xlUp).Row

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

' This function consolidates identically formatted, DB-like tables. The result is returned either
' as a 2D array or consolidated in the optionally given TargetWorksheet. The function requires the
' following:
'
' 1. All of the referenced worksheets in parameter WorksheetsArray() must contain identical DB-like tables
' 2. All of the referenced worksheeets must contained the same header row
' 3. If provided FirstDataRow, FirstDataRow must be > 2. The header row must in in row position
'    FirstDataRow-1
' 4. If FirstDataRow is not given, it is assumed FirstDataRow=2. If FirstDataRow<=1, then
'    FirstDataRow is assumed to equal 2.
'
' If the optional parameter TargetWorksheet is supplied, then the function returns Null
' and copies the consolidated data into the target worksheet starting in cell A1.
Public Function ConsolidateWorksheets(WorksheetsArray() As Worksheet, _
                                      Optional TheFirstDataRow As Variant, _
                                      Optional TargetWorksheet As Variant) As Variant
    Dim Headers As Variant
    Dim NumColumns As Integer
    Dim NumRows As Long
    Dim FirstDataRow As Long
    Dim LastDataRow As Long
    Dim ConsolidationWorksheet As Worksheet
    Dim SourceRange As Range
    Dim TargetRange As Range
    Dim N As Long
    Dim RowCursor As Long
    
    ' Default return value. Only changed below iff TargetWorksheet missing
    Let ConsolidateWorksheets = Null
    
    ' Set the first row (where data starts, not headers)
    If IsMissing(FirstDataRow) Then
        Let FirstDataRow = 2
    ElseIf FirstDataRow > 1 Then
        Let FirstDataRow = TheFirstDataRow
    Else
        Let FirstDataRow = 2
    End If
    
    ' Instantiate the consolidated worksheet.  It will be deleted at the end once data
    ' is copied to the target worksheet or returned as a 2D array.
    If IsMissing(TargetWorksheet) Then
        Set ConsolidationWorksheet = TempComputation
        
        ' Clear any pre=existing formats and content
        Call ConsolidationWorksheet.UsedRange.ClearFormats
        Call ConsolidationWorksheet.UsedRange.ClearContents
    Else
        Set ConsolidationWorksheet = TargetWorksheet
    End If
    
    ' Determine the number of columns in first worksheet to consolidate
    Let NumColumns = WorksheetsArray(1).UsedRange.Columns.Count
    
    ' Get the headers. May be multiple rows, starting from the top
    Let Headers = WorksheetsArray(1).Range("A1").Resize(FirstDataRow - 1, NumColumns).Value2

    ' Copy the headers row to ConsolidatedWorksheet
    Call DumpInSheet(Headers, ConsolidationWorksheet.Range("A1"))
    
    ' Set the current row where data should be dropped in the consolidation worksheet
    Let RowCursor = FirstDataRow
    For N = LBound(WorksheetsArray) To UBound(WorksheetsArray)
        ' Compute the last row to consolidate in this worksheet
        Let LastDataRow = WorksheetsArray(N).Cells(WorksheetsArray(1).Rows.Count, 1).End(xlUp).Row
        
        ' Compute the number of rows to consolidate
        Let NumRows = LastDataRow - FirstDataRow + 1
                        
        ' Set the source range
        Set SourceRange = WorksheetsArray(N).Cells(FirstDataRow, 1).Resize(NumRows, NumColumns)
        
        ' Set the target range
        Set TargetRange = ConsolidationWorksheet.Cells(RowCursor, 1).Resize(NumRows, NumColumns)
        
        ' Do the actual consolidation
        Call SourceRange.Copy
        Call TargetRange.PasteSpecial(Paste:=xlPasteAll)
        
        ' Update RowCursor
        Let RowCursor = RowCursor + NumRows
    Next N
    
    ' Exit since data in TargetWorksheet by this point in code if TargetWorksheet give
    If Not IsMissing(TargetWorksheet) Then Exit Function
    
    ' If TargetWorksheet not provided, return a 2D matrix with the data
    Let ConsolidateWorksheets = ConsolidationWorksheet.Cells(1, 1).Resize(RowCursor, NumColumns).Value2
    
    ' Delete the consolidation worksheet
    Call ConsolidationWorksheet.UsedRange.ClearFormats
    Call ConsolidationWorksheet.UsedRange.ClearContents
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
