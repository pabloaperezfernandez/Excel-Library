Attribute VB_Name = "PivotTablesModule"
Option Explicit
Option Base 1


' DESCRIPTION
' This function generates a pivot table in a new worksheet of the given workbook using
' the requested column, row, and data fields.
'
' INPUTS
' 1. TargetWorkbook - Reference to the workbook where pivot tables should be created
' 2. TargetPivotTableName - The name to give to the pivot table. This is also used as
'    the name of the worksheet holding the pivot table
' 3. SourceListObject - The listobject holding the data for the pivot table's cache
' 4. RowPivotFieldName - The name of the listobject's column to use for the pivot
'    table's row field
' 5. ColumnPivotFieldName - The name of the listobject's column to use for the pivot
'    table's column field
' 6. PagePivotFieldName - The name of the listobject's column to use for the pivot
'    table's page field
' 7. PageFieldValue - Value to use in the page field to filter the pivot summary
' 8. DataFieldName - The name of the listobject's column to summarize using addition
' 9. ConsolidationFunction - One of the value for enumerated data type XlConsolidationFunction.
'                            For example, xlSum use to summarize the chosen DataFieldName
' 10. ConsolidationFieldLabel - Label such as "Sum of " to let the user know how
'     data field is summarized
' 11. DataFieldFormatString - Formatting string to use with .NumberFormat on the summarized
'     data field
'
' OUTPUT
' A pivot table in a new worksheet with the chosen name in TargetWorkbook. This pivot
' summarizes the chosen data field with respect to the chosen cuts.
Public Sub GeneratePivotTable(TargetWorksheet As Worksheet, _
                              TargetPivotTableName As String, _
                              SourceListObject As ListObject, _
                              RowPivotFieldName As String, _
                              ColumnPivotFieldName As String, _
                              PagePivotFieldName As String, _
                              PageFieldValue As String, _
                              DataFieldName As String, _
                              ConsolidationFunction As XlConsolidationFunction, _
                              ConsolidationFieldLabel As String, _
                              DataFieldFormatString As String)
    Dim pc As PivotCache
    Dim pt As Variant ' Cannot get it work with pt As PivotTable
    Dim SrcWbk As Workbook
    Dim SrcWsht As Worksheet
    Dim TgtWbk As Workbook

    ' Set reference to source workbook, worksheet, and range
    Set SrcWsht = SourceListObject.Range.Worksheet
    Set SrcWbk = SrcWsht.Parent
    Set TgtWbk = TargetWorksheet.Parent
    
    ' Create sectoral exposures pivot table
    Set pc = TgtWbk.PivotCaches.Create(SourceType:=xlDatabase, _
                                       sourceData:=StringJoin(Array("'[", SrcWbk.Name, "]", SrcWsht.Name, "'!", SourceListObject.Name)), _
                                       Version:=xlPivotTableVersion14)

    ' Create pivot table from above pivot cache
    Set pt = pc.CreatePivotTable(TableDestination:=TargetWorksheet.Range("A1"), _
                                 TableName:=TargetPivotTableName, _
                                 DefaultVersion:=xlPivotTableVersion14)

    With pt
        ' Add page field
        If Not (PagePivotFieldName <> Empty Or IsEmpty(PagePivotFieldName) Or PagePivotFieldName = "") Then
            With .PivotFields(PagePivotFieldName)
                .Orientation = xlPageField
                .Position = 1
            End With
        End If

        ' Add column field
        With .PivotFields(ColumnPivotFieldName)
            .Orientation = xlColumnField
            .Position = 1
        End With

        ' Add row field
        With .PivotFields(RowPivotFieldName)
            .Orientation = xlRowField
            .Position = 1
        End With

        ' Add data field to summarize using addition
        Let .AddDataField(.PivotFields(DataFieldName), _
                          ConsolidationFieldLabel & DataFieldName, _
                          ConsolidationFunction).NumberFormat = DataFieldFormatString

        ' Set the page field to the chosen date
        If PagePivotFieldName <> vbNullString And Not IsEmpty(PagePivotFieldName) Then
            Let .PivotFields(PagePivotFieldName).CurrentPage = PageFieldValue
        End If

        ' Display only column grand totals
        Let .ColumnGrand = True
        Let .RowGrand = False

        ' Hide display captions
        Let .DisplayFieldCaptions = False
    End With
End Sub
