Attribute VB_Name = "Main_UI"
' PURPOSE OF THIS MODULE
'
' The purpose of this module is to provide the functionality that runs the UI
' in worksheet UI
' and filenames.
Option Explicit
Option Base 1

' DESCRIPTION
' Main application drivers. Executes when the user clicks on the button "CONSOLIDATE"
' in the UI worksheet
'
' PARAMETERS
' 1. Worksheets("UI").Range("SharePointDirectoryPath") - The path to the target SharePoint DIR
' 2. Worksheets("UI").ListObjects("TeamInputWorkbooksLo") - List of team-specific inputs sheets
'
' RETURNED VALUE
' Returns a filename with the full path
Public Sub MainDriver()
    ' Turn off screen updating and user confirmations
    Let Application.ScreenUpdating = False
    Let Application.DisplayAlerts = False
    Let Application.Calculation = xlCalculationManual
    
    ' Save archived copy of this workbook in user's downloads directory
    Call ThisWorkbook.SaveCopyAs(filename:=MakeWorkbookArchivalFileName(ThisWorkbook.Name))
    
    ' Unprotect workbook and worksheets
    Call UnprotectWorkbookAndWorksheets

    ' Loop over input files, pulling data into this workbook
    Call ImportInputFiles
    
    ' Archive targetBisAllocationLo table and two other input tables
    Call ArchiveInputs
    
    ' Update unique dates for comparison of changes in inputs
    Call UpdateListOfTimeDatedCrossSections

    ' Protect workbook worksheets before saving the changes
    Call ProtectWorkbookAndWorksheets
    
    ' Save changes
    Call ThisWorkbook.Save
    
    ' Make sure the UI worksheet is visible when the screen refreshes
    Call ThisWorkbook.Worksheets("UI").Activate
    
    ' Turn on auto screen refresh and user confirmations
    Let Application.Calculation = xlCalculationAutomatic
    Let Application.ScreenUpdating = True
    Let Application.DisplayAlerts = True
    
    ' Display msg letting the user know the consolidation is done
    Call MsgBox("Consolidation Done")
End Sub

' DESCRIPTION
' Unprotect this workook and certain worksheets to prevent accidental editing.
' It uses to global variable Main_UI.CONST_password.
'
' PARAMETERS
' 1. Main_UI.CONST_password (string): Global variable in this module
'
' RETURNED VALUE
' None
Public Sub UnprotectWorkbookAndWorksheets()
    Dim varStr As Variant
    Dim workSheetNames() As String
    Dim thePassword As String
    
    With ThisWorkbook.Worksheets("UI")
        Let workSheetNames = TypeConversions.ToStrings(Flatten(.ListObjects("ProtectedWshtsLuLo").DataBodyRange.Value2))
        Let thePassword = .Range("Password")
    End With

    ' Unprotect workbook worksheets
    Call ThisWorkbook.Unprotect(password:=thePassword)
    For Each varStr In workSheetNames
        Call UnprotectWorksheet(wsht:=ThisWorkbook.Worksheets(varStr), thePassword:=thePassword)
    Next
    
    ' Make sure the UI worksheet is visible when the screen refreshes
    Call ThisWorkbook.Worksheets("UI").Activate
End Sub

' DESCRIPTION
' Protect this workook and certain worksheets to prevent accidental editing.
' It uses to global variable Main_UI.CONST_password.
'
' PARAMETERS
' 1. Main_UI.CONST_password (string): Global variable in this module
'
' RETURNED VALUE
' None
Public Sub ProtectWorkbookAndWorksheets()
    Dim varStr As Variant
    Dim workSheetNames() As String
    Dim thePassword As String
    
    With ThisWorkbook.Worksheets("UI")
        Let workSheetNames = TypeConversions.ToStrings(Flatten(.ListObjects("ProtectedWshtsLuLo").DataBodyRange.Value2))
        Let thePassword = .Range("Password")
    End With
    
    ' Protect workbook worksheets
    Call ThisWorkbook.Protect(password:=thePassword, Structure:=True)
    For Each varStr In workSheetNames
        Call ProtectWorksheet(wsht:=ThisWorkbook.Worksheets(varStr), thePassword:=thePassword)
    Next
    
    ' Make sure the UI worksheet is visible when the screen refreshes
    Call ThisWorkbook.Worksheets("UI").Activate
End Sub

' DESCRIPTION
' Re-populate the L7 input workbooks when we make changes to the lookup values or mapping tables.
' This is particularly useful when new projects are added. The function creates a timed-dated
' snapshot in the appropriate worksheet and listobjects
'
' PARAMETERS
' N/A
'
' RETURNED VALUE
' None
Public Sub RegenerateL7OrgInputSheets()
    Dim filteredRangeValues As Variant
    Dim tgtValsWsht As Worksheet
    Dim luLoNames As Variant
    Dim inputWbkNames As Variant
    Dim bisAliasAllocationsWbkNames As Variant
    Dim twoPizzaTeamNames As Variant
    Dim wbksLo As ListObject
    Dim wbk As Workbook
    Dim mngrToOrgLo As ListObject
    Dim srcLo As ListObject
    Dim tgtLo As ListObject
    Dim strVariant As Variant
    Dim strWbkNameVariant As Variant
    Dim strColNameVariant As Variant
    Dim ArchiveGoalsToLevelLo As ListObject
    Dim ArchiveManagerToOrgLo As ListObject
    
    Call UnprotectWorkbookAndWorksheets

    ' Set a few references
    With ThisWorkbook.Worksheets("UI")
        Let luLoNames = Flatten(.ListObjects("LookupTableNamesLuLo").DataBodyRange.Value2)
        Let inputWbkNames = Flatten(.ListObjects("TeamInputWorkbooksLo").DataBodyRange.Value2)
        Let bisAliasAllocationsWbkNames = Flatten(.ListObjects("TeamInputWorkbooksWithBisAliasAllocationsLuLo").DataBodyRange.Value2)
    End With
    
    With ThisWorkbook
        Set ArchiveGoalsToLevelLo = .Worksheets("ArchiveGoalsToLevelLo").ListObjects("ArchiveGoalsToLevelLo")
        Set ArchiveManagerToOrgLo = .Worksheets("ArchiveManagerToOrgLo").ListObjects("ArchiveManagerToOrgLo")
    End With
    
    ' Archive GoalsToLevelLo, ManagerToOrgLo, and LookupValues
    With ThisWorkbook
        Call ArchiveTable(.Worksheets("GoalsToLevelLo").ListObjects("GoalsToLevelLo"), _
                          ArchiveGoalsToLevelLo)
        Call ArchiveTable(.Worksheets("ManagerToOrgLo").ListObjects("ManagerToOrgLo"), _
                          ArchiveManagerToOrgLo)
    End With
    
    ' Archive lookup values
    For Each strVariant In luLoNames
        ' Set references to sources and target listobjects
        With ThisWorkbook
            Set srcLo = .Worksheets("LookupValues").ListObjects(strVariant)
            Set tgtLo = .Worksheets("ArchiveLookupValues").ListObjects("Archive" & strVariant)
        End With
        ' Populate the target with the values from the source
        Call ArchiveTable(srcLo, tgtLo)
    Next
    
    ' Needed to determine two-pizza teams rolling up to an L7 org input file
    Set mngrToOrgLo = ThisWorkbook.Worksheets("ManagerToOrgLo").ListObjects("ManagerToOrgLo")
    
    ' Loop over input files, pulling data into this workbook
    For Each strWbkNameVariant In inputWbkNames
        ' Open the source file and set a reference
        Set wbk = Application.Workbooks.Open(filename:=GetInputWorkbookName(CStr(strWbkNameVariant)), _
                                             UpdateLinks:=False, ReadOnly:=False)
        Call wbk.LockServerFile
        
        ' Re-populate the lookup tables in the input wbk with the values in this workbook
        For Each strVariant In luLoNames
            ' Set references to sources and target listobjects
            Set srcLo = ThisWorkbook.Worksheets("LookupValues").ListObjects(strVariant)
            Set tgtLo = wbk.Worksheets("LookupValues").ListObjects(strVariant)
            ' Clear the target listobject
            Call ClearListObjectDataBodyRange(tgtLo)
            Let tgtLo.ShowTotals = False
            ' Populate the target with the values from the source
            Call AppendToListObject(TheData:=srcLo.DataBodyRange.Value2, lo:=tgtLo)
        Next
        
        ' Re-populate listobjects the input workbook using the data in this workbook
        For Each strVariant In Array("ManagerToOrgLo", "GoalsToLevelLo")
            Set srcLo = ThisWorkbook.Worksheets(strVariant).ListObjects(strVariant)
            Set tgtLo = wbk.Worksheets(strVariant).ListObjects(strVariant)
            Call ClearListObjectDataBodyRange(tgtLo)
            Let tgtLo.ShowTotals = False
            Call AppendToListObject(TheData:=srcLo.DataBodyRange.Value2, lo:=tgtLo)
        Next

        ' Turn on autofilters on listobject and show all data
        Let twoPizzaTeamNames = GetTwoPizzaTeamsForL7OrgInputFile(CStr(strWbkNameVariant))
        
        ' Get BisAllocationsLo subset for the L7 Org
        Let filteredRangeValues = GetBisAllocationsLoForL7Org(CStr(strWbkNameVariant))
        Set tgtLo = wbk.Worksheets("BisAllocationsLo").ListObjects("bisAllocationsLo")
        Call ClearListObjectDataBodyRange(tgtLo)
        Let tgtLo.ShowTotals = False
        Call AppendToListObject(filteredRangeValues, tgtLo)
        
        ' If this is one the input workbooks that includes a BIS Alias allocations worksheet:
        ' 1. Copy the bis alias allocations worksheet contents from this workbook to the target
        '    workbook and worksheet
        ' 2. Filter out any rows that do not correspond to that L7 org
        ' 3. Clear values from the databody range of the target listobject
        ' 4. Restore the formulas of the target listobject
        If MemberQ(bisAliasAllocationsWbkNames, strWbkNameVariant) Then
            ' ***HERE
        End If

        ' Close the source workbook
        Call wbk.Close(SaveChanges:=False)
    Next
    
    Call ProtectWorkbookAndWorksheets

    ' Make sure the UI worksheet is visible when the screen refreshes
    Call ThisWorkbook.Worksheets("UI").Activate
End Sub

' DESCRIPTION
' HELPER sub for Public Sub MainDriver(). Its function is to store a date-time
' snapshot of the three input tables that consolidate inputs from the two pizza
' teams so we can later determine what changed between updates
'
' PARAMETERS
' None
'
' RETURNED VALUE
' Archival tables in worksheets ArchiveBisAllocationsLo and ArchiveInputTables are
' appended with the state of the corresponding input tables at the start of the consolidation
Private Sub ArchiveTable(sourceListObject As ListObject, targetListObject As ListObject)
    Dim DumpArray As Variant
    Dim theSerialDate As Variant
    Dim TotalsQ As Boolean
    
    ' Store state of autofilters
    Let TotalsQ = sourceListObject.ShowAutoFilter

    If Not (sourceListObject.DataBodyRange Is Nothing) Then
        ' Turn off autofilters if on
        If TotalsQ Then Let sourceListObject.ShowAutoFilter = False
        
        Let DumpArray = sourceListObject.DataBodyRange.Value2
        Let theSerialDate = ConstantArray(Now(), Length(DumpArray))
        If NumberOfColumns(DumpArray) = 1 Then
            Let DumpArray = Array(Transpose(DumpArray, UseBuiltInQ:=True, ParameterCheckQ:=False))
        Else
            Let DumpArray = Transpose(DumpArray, UseBuiltInQ:=True, ParameterCheckQ:=False)
        End If
        Let DumpArray = Prepend(DumpArray, theSerialDate)
        Let DumpArray = Transpose(DumpArray, UseBuiltInQ:=True, ParameterCheckQ:=False)
        Call AppendToListObject(DumpArray, targetListObject)
        
        ' Turn autofilters on if they were when ArchiveTable called.
        If TotalsQ Then Let sourceListObject.ShowAutoFilter = True
    End If
End Sub

' DESCRIPTION
' HELPER sub for Public Sub MainDriver(). Its function is to cycle through the two-pizza
' team input files, importing them into the target tables to consolidate into this workbook.
'
' PARAMETERS
' 1. Worksheets("UI").Range("SharePointDirectoryPath") - The path to the target SharePoint DIR
' 2. Worksheets("UI").ListObjects("TeamInputWorkbooksLo") - List of team-specific inputs sheets
'
' RETURNED VALUE
' Input tables in worksheet Overview with all data from the two-pizza team files consolidated
Private Sub ImportInputFiles()
    Dim i As Integer
    Dim wbk As Workbook
    Dim inputWbkNames() As Variant
    Dim sourceBisAllocationLo As ListObject
    Dim sourceBisCapacityLo As ListObject
    Dim sourceTeamHcLo As ListObject
    Dim targetBisAllocationLo As ListObject
    Dim targetBisCapacityLo As ListObject
    Dim targetTeamHcLo As ListObject
    Dim targetTeamBisAllocationErrorCheckLo As ListObject
    Dim targetTeamBisToHcGapLo As ListObject
    Dim inputFileRowCountLo As ListObject
    Dim varStr As Variant
    Dim varLo As Variant
    Dim lo As ListObject
    Dim RowCounts As Variant
    Dim TgTRange As Range
    
    ' Set reference to this workbook
    With ThisWorkbook.Worksheets("UI").ListObjects("TeamInputWorkbooksLo")
        Let inputWbkNames = Flatten(.DataBodyRange.Value2)
    End With
    
    Let RowCounts = ConstantArray(Empty, Length(inputWbkNames) + 2, 2)
    Let RowCounts(1, 1) = "InputFileRowCountLo"
    Let RowCounts(1, 2) = Empty
    Let RowCounts(2, 1) = "ERROR CHECK: Row Count in BisAllocationsLo Equals Sum of Rows in Input Wbks"
    Let RowCounts(2, 2) = Empty
    
    Set targetBisAllocationLo = ThisWorkbook.Worksheets("BisAllocationsLo").ListObjects("BisAllocationsLo")
    With ThisWorkbook.Worksheets("Overview")
        Set targetBisCapacityLo = .ListObjects("TeamBISCapacityLo")
        Set targetTeamHcLo = .ListObjects("TeamHCLo")
        Set targetTeamBisAllocationErrorCheckLo = .ListObjects("TeamBisAllocationErrorCheckLo")
        Set targetTeamBisToHcGapLo = .ListObjects("TeamBisToHcGapLo")
        Set inputFileRowCountLo = ThisWorkbook.Worksheets("Overview").ListObjects("inputFileRowCountLo")
    End With
    
    Call inputFileRowCountLo.Range.CurrentRegion.ClearFormats
    Call inputFileRowCountLo.Range.CurrentRegion.ClearContents

    ' Save formulas in table
    For Each varLo In Array(targetBisAllocationLo, targetTeamBisAllocationErrorCheckLo, targetTeamBisToHcGapLo)
        Set lo = varLo: Call SaveListObjectFormulas(lo)
    Next

    ' Clear tables where we will consolidate data from input files
    For Each varLo In Array(targetBisAllocationLo, targetBisCapacityLo, targetTeamHcLo)
        Set lo = varLo: Call ClearListObjectDataBodyRange(lo)
    Next

    ' Import input files.
    For i = LBound(inputWbkNames) To UBound(inputWbkNames)
        ' Open the source file and set a reference
        Set wbk = Application.Workbooks.Open(filename:=GetInputWorkbookName(CStr(inputWbkNames(i))), _
                                             UpdateLinks:=False, ReadOnly:=True)
        
        ' Set references to the tables we are consolidating from the source workbook
        With wbk
            Set sourceBisAllocationLo = .Worksheets("BisAllocationsLo").ListObjects("BisAllocationsLo")
            Set sourceBisCapacityLo = .Worksheets("Overview").ListObjects("TeamBISCapacityLo")
            Set sourceTeamHcLo = .Worksheets("Overview").ListObjects("TeamHCLo")
        End With
        
        ' Copy table data from the source tables to the target locations in this workbook
        Call AppendToListObject(sourceBisAllocationLo.DataBodyRange.Value2, targetBisAllocationLo)
        Call AppendToListObject(sourceBisCapacityLo.DataBodyRange.Value2, targetBisCapacityLo)
        Call AppendToListObject(sourceTeamHcLo.DataBodyRange.Value2, targetTeamHcLo)
        
        Let RowCounts(i + 2, 1) = inputWbkNames(i)
        Let RowCounts(i + 2, 2) = sourceBisAllocationLo.ListRows.Count
        
        ' Close the source workbook
        Call wbk.Close(SaveChanges:=False)
    Next i
    
    ' Sort target tables in Worksheets("Overview") so all tables in the worksheet are in the same order
    ' and all formulas work
    Call SortListObject(targetBisCapacityLo, "Team", xlAscending)
    Call SortListObject(targetTeamHcLo, "Team", xlAscending)
    
    ' Clear listcolumns before restoring formulas in listobjects TeamBisAllocationErrorCheckLo and
    ' TeamBisToHcGapLo in Worksheets("Overview"). The reason we need to do this is because
    ' Deleting the databodyranges of tables TeamHCLo and TeamBISCapacityLo destroys formulas.
    For Each varStr In Take(Flatten(targetTeamHcLo.HeaderRowRange.Value2), Array(4, -1))
        Call targetTeamBisAllocationErrorCheckLo.ListColumns(varStr).DataBodyRange.ClearContents
        Call targetTeamBisToHcGapLo.ListColumns(varStr).DataBodyRange.ClearContents
    Next

    ' Restore formulas to update computations
    For Each varLo In Array(targetBisAllocationLo, targetTeamBisAllocationErrorCheckLo, targetTeamBisToHcGapLo)
        Set lo = varLo: Call RestoreListObjectFormulas(lo)
    Next
    
    ' Populate Worksheets("Overview").ListObjects("inputFileRowCountLo")
    Set TgTRange = targetBisCapacityLo.HeaderRowRange.Range("A1").Offset(0, targetBisCapacityLo.ListColumns.Count + 1)
    Call DumpInSheet(Array("Filename", "Row Count"), TgTRange)
    Call DumpInSheet(Rest(Rest(RowCounts)), TgTRange.Offset(1, 0).Range("A1"))
    Set inputFileRowCountLo = AddListObject(ARangeInCurrentRegion:=TgTRange, ListObjectName:="InputFileRowCountLo")
    Call DumpInSheet(Take(RowCounts, 2), inputFileRowCountLo.HeaderRowRange.Offset(-2, 0))
    Let inputFileRowCountLo.HeaderRowRange.Offset(-2, 0).Interior.ColorIndex = 1
    Let inputFileRowCountLo.HeaderRowRange.Offset(-2, 0).Font.Color = 16777215
    Let inputFileRowCountLo.HeaderRowRange.Offset(-1, 0).Interior.ColorIndex = 12
    Let inputFileRowCountLo.HeaderRowRange.Offset(-1, 0).Font.Color = 16777215
    Let inputFileRowCountLo.ShowTotals = True
    Set TgTRange = inputFileRowCountLo.Range(1, 1).Offset(inputFileRowCountLo.ListRows.Count + 2, 0)
    Let TgTRange.Value2 = "Number of Rows in BisAllocationsLo"
    Let TgTRange.Offset(0, 1).Formula = "=COUNTA(BisAllocationsLo[Project])"
    Call TurnOffAutoFilters(inputFileRowCountLo)
End Sub

' DESCRIPTION
' HELPER sub for Public Sub MainDriver(). Its function is to determine the unique set of date/times
' in the crossectional archive in worksheet ArchiveBisAllocationsLo
'
' PARAMETERS
' None directly. The sole input is column ArchiveBisAllocationsLo.ListColumns("Serial Snapshot Date")
'
' RETURNED VALUE
' The table UniqueArchiveDatesLuLo is updated with the unique set of archival dates/times
Private Sub UpdateListOfTimeDatedCrossSections()
    Dim uniqueArchiveDatesLuLo As ListObject
    Dim archiveTargetBisAllocationLo As ListObject
    Dim DumpArray As Variant
    
    With ThisWorkbook
        Set uniqueArchiveDatesLuLo = .Worksheets("ChangesToPlan").ListObjects("UniqueArchiveDatesLuLo")
        Set archiveTargetBisAllocationLo = .Worksheets("ArchiveBisAllocationsLo").ListObjects("ArchiveBisAllocationsLo")
    End With

    Call ClearListObjectDataBodyRange(ThisWorkbook.Worksheets("ChangesToPlan").ListObjects("UniqueArchiveDatesLuLo"))
    If Not (archiveTargetBisAllocationLo.ListColumns("Serial Snapshot Date").DataBodyRange Is Nothing) Then
        Let DumpArray = UniqueSubset(Flatten(archiveTargetBisAllocationLo.ListColumns("Serial Snapshot Date").DataBodyRange.Value2))
        Call ClearListObjectDataBodyRange(uniqueArchiveDatesLuLo)
        Call AppendToListObject(Transpose(DumpArray), uniqueArchiveDatesLuLo)
    End If
End Sub

' DESCRIPTION
' HELPER sub for Public Sub MainDriver(). Its function is to determine the unique set of date/times
' in the crossectional archive in worksheet ArchiveBisAllocationsLo
'
' PARAMETERS
' None directly. The sole input is column ArchiveBisAllocationsLo.ListColumns("Serial Snapshot Date")
'
' RETURNED VALUE
' The table UniqueArchiveDatesLuLo is updated with the unique set of archival dates/times
Private Sub ArchiveInputs()
    Dim archiveTargetBisAllocationLo As ListObject
    Dim archiveTargetBisCapacityLo As ListObject
    Dim archiveTargetTeamHcLo As ListObject
    Dim targetBisAllocationLo As ListObject
    Dim targetBisCapacityLo As ListObject
    Dim targetTeamHcLo As ListObject

    ' Set references to worksheets in this workbook where we will copy or modify inputs
    With ThisWorkbook
        Set targetBisAllocationLo = .Worksheets("BisAllocationsLo").ListObjects("BisAllocationsLo")
        Set targetBisCapacityLo = .Worksheets("Overview").ListObjects("TeamBISCapacityLo")
        Set targetTeamHcLo = .Worksheets("Overview").ListObjects("TeamHCLo")
        Set archiveTargetBisAllocationLo = .Worksheets("ArchiveBisAllocationsLo").ListObjects("ArchiveBisAllocationsLo")
        Set archiveTargetBisCapacityLo = .Worksheets("ArchiveInputTables").ListObjects("ArchiveTeamBISCapacityLo")
        Set archiveTargetTeamHcLo = .Worksheets("ArchiveInputTables").ListObjects("ArchiveTeamHCLo")
    End With

    Call ArchiveTable(sourceListObject:=targetBisAllocationLo, targetListObject:=archiveTargetBisAllocationLo)
    Call ArchiveTable(sourceListObject:=targetBisCapacityLo, targetListObject:=archiveTargetBisCapacityLo)
    Call ArchiveTable(sourceListObject:=targetTeamHcLo, targetListObject:=archiveTargetTeamHcLo)
End Sub

' DESCRIPTION
' HELPER sub. It returns the list of two-pizza teams rolling up to an L7 org. The L7 org
' is specified by the name of its L7 org input workbook
'
' PARAMETERS
' None directly. The sole input is column ArchiveBisAllocationsLo.ListColumns("Serial Snapshot Date")
'
' RETURNED VALUE
' The table UniqueArchiveDatesLuLo is updated with the unique set of archival dates/times
Private Function GetTwoPizzaTeamsForL7OrgInputFile(wbkName As String) As Variant
    Dim mngrToOrgLo As ListObject
    Dim filteredRangeValues As Variant
    Dim fieldPos As Integer
    
    ' Set a reference to listobject Worksheets("ManagerToOrgLo").ListObjects("ManagerToOrgLo")
    Set mngrToOrgLo = ThisWorkbook.Worksheets("ManagerToOrgLo").ListObjects("ManagerToOrgLo")

    ' Turn on autofilters on listobject and show all data
    Call TurnOnAutoFilters(mngrToOrgLo)
    
    ' Get two-pizza team names corresponding to L7 org
    Let fieldPos = Application.Match("Input Workbook Name", mngrToOrgLo.HeaderRowRange, 0)
    Call mngrToOrgLo.Range.AutoFilter(Field:=fieldPos, _
                                      Criteria1:=Array(wbkName), _
                                      Operator:=xlFilterValues)
    Let filteredRangeValues = GetVisibleSpecialValues(mngrToOrgLo)
    If mngrToOrgLo.ShowTotals = True Then filteredRangeValues = Most(filteredRangeValues)
    Let filteredRangeValues = UniqueSubset(Rest(GetColumnByHeader(filteredRangeValues, "Team")))
    
    ' Show all data and turn off autofilters
    Call TurnOffAutoFilters(mngrToOrgLo)
    
    Let GetTwoPizzaTeamsForL7OrgInputFile = filteredRangeValues
End Function

' DESCRIPTION
' HELPER sub. It takes the basefilename of the L7 org input workbook
'
' PARAMETERS
' None directly. The sole input is listobject WorkSheets("BisAllocationsLo").ListObjects("BisAllocationsLo")
'
' RETURNED VALUE
' The subset of values from BisAllocationsLo corresponding to the L7 org. Header row excluded.
Public Function GetBisAllocationsLoForL7Org(wbkName As String) As Variant
    Dim mngrToOrgLo As ListObject
    Dim bisAllocationsLo As ListObject
    Dim twoPizzaTeams As Variant
    Dim filteredRangeValues As Variant
    
    ' Set references to listobjects
    With ThisWorkbook
        Set mngrToOrgLo = .Worksheets("ManagerToOrgLo").ListObjects("ManagerToOrgLo")
        Set bisAllocationsLo = .Worksheets("BisAllocationsLo").ListObjects("BisAllocationsLo")
    End With
    
    ' Get the list of two-pizza teams rolling up to the given L7 org filename
    Let twoPizzaTeams = GetTwoPizzaTeamsForL7OrgInputFile(wbkName)

    ' Turn on autofilters on listobject and show all data
    Call TurnOnAutoFilters(bisAllocationsLo)
    
    ' Get two-pizza team names corresponding to L7 org
    Call bisAllocationsLo.Range.AutoFilter(Field:=Application.Match("Team", bisAllocationsLo.HeaderRowRange, 0), _
                                           Criteria1:=twoPizzaTeams, _
                                           Operator:=xlFilterValues)
    Let filteredRangeValues = Rest(GetVisibleSpecialValues(bisAllocationsLo))
    If bisAllocationsLo.ShowTotals = True Then filteredRangeValues = Most(filteredRangeValues)
    
    ' Show all data and turn off autofilters
    Call TurnOffAutoFilters(bisAllocationsLo)
    
    Let GetBisAllocationsLoForL7Org = filteredRangeValues
End Function
