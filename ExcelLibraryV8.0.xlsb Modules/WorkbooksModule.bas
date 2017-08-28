Attribute VB_Name = "WorkbooksModule"
Option Explicit
Option Base 1

' This function returns a consolidated DB-like table from an array of workbook references,
' all of which contain identical DB-like tables in their sole worksheets. All worksheets must
' contain headers in the first row, unless the optional starting row is provide.  If StartingRow
' is provided, the function includes the first StartingRow-1 rows as headers in the consolidated array.
' The data starting in row StartingRow from each workbook is consolidated in the consolidated array.
Public Function ConsolidateWorkbooks(WorkbooksArray() As Workbook, Optional WorkSheetNamesArray As Variant, Optional StartingRow As Variant) As Variant
    Dim WorksheetsArray() As Worksheet
    Dim N As Integer
    
    ReDim WorksheetsArray(LBound(WorkbooksArray) To UBound(WorkbooksArray))
    For N = LBound(WorkbooksArray) To UBound(WorkbooksArray)
        If IsMissing(WorkSheetNamesArray) Then
            Set WorksheetsArray(N) = WorkbooksArray(N).Worksheets(1)
        Else
            Set WorksheetsArray(N) = WorkbooksArray(N).Worksheets(WorkSheetNamesArray(N))
        End If
    Next N

    ' Return a reference to the consolidation worksheet
    If IsMissing(StartingRow) Then
        Let ConsolidateWorkbooks = ConsolidateWorksheets(WorksheetsArray)
    ElseIf IsNumeric(StartingRow) Then
        If StartingRow >= 1 Then
            Let ConsolidateWorkbooks = ConsolidateWorksheets(WorksheetsArray, StartingRow)
        Else
            Let ConsolidateWorkbooks = ConsolidateWorksheets(WorksheetsArray)
        End If
    Else
        Let ConsolidateWorkbooks = ConsolidateWorksheets(WorksheetsArray)
    End If
End Function

' This sub closes all workbooks other than the one whose reference has been passed
Public Sub CloseAllOtherWorkbooks(CallingWorkbook As Workbook)
    Dim wbk As Workbook
    
    For Each wbk In Application.Workbooks
        If Not wbk Is CallingWorkbook Then Call wbk.Close(SaveChanges:=False)
    Next
End Sub
