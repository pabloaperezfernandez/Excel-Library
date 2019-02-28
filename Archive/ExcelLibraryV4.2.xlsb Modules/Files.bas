Attribute VB_Name = "Files"
Option Explicit
Option Base 1


'*********************************************************************************************************
'* This section sets the base directories for the production environment
'*********************************************************************************************************

' This sets the root production directory
Public Function ProductionRootDir() As String
    Let ProductionRootDir = ThisWorkbook.Path & "\..\"
End Function

' This sets the production mappings directory
Public Function ProductionMappingsDir() As String
    Let ProductionMappingsDir = ProductionRootDir() & "Common Software Directory\Mappings\"
End Function

'*********************************************************************************************************
'* This section sets the production directories for the trade execution transformer
'*********************************************************************************************************

' This sets the base root directory of account admin's production environment
Public Function TransformerProductionDir() As String
    Let TransformerProductionDir = ProductionRootDir() & "Transformer Production Directory\"
End Function

' This sets the production input directory for account administration files
Public Function TransformerProductionInputDir() As String
    Let TransformerProductionInputDir = TransformerProductionDir() & "Transformer Input Directory\"
End Function

' This sets the production output directory for account administration files
Public Function TransformerProductionOutputDir() As String
    Let TransformerProductionOutputDir = TransformerProductionDir() & "Transformer Output Directory\"
End Function

' This sets the production intermediate directory for account administration file
Public Function TransformerProductionInterdiateDir() As String
    Let TransformerProductionInterdiateDir = TransformerProductionDir() & "Transformer Intermediate Directory\"
End Function

'*********************************************************************************************************
'* This section sets the production directories for the WIP
'*********************************************************************************************************

' This sets the base root directory of account admin's production environment
Public Function WipProductionDir() As String
    Let WipProductionDir = ProductionRootDir() & "EquityWip Production Directory\"
End Function

' This sets the production input directory for account administration files
Public Function WipProductionInputDir() As String
    Let WipProductionInputDir = WipProductionDir() & "EquityWip Input Directory\"
End Function

' This sets the production output directory for account administration files
Public Function WipProductionOutputDir() As String
    Let WipProductionOutputDir = WipProductionDir() & "EquityWip Output Directory\"
End Function

' This sets the production intermediate directory for account administration file
Public Function WipProductionIntermediateDir() As String
    Let WipProductionIntermediateDir = WipProductionDir() & "EquityWip Intermediate Directory\"
End Function

'*********************************************************************************************************
'* This section sets the production directories for the ETF WIP
'*********************************************************************************************************

' This sets the base root directory of account admin's production environment
Public Function EtfWipProductionDir() As String
    Let EtfWipProductionDir = ProductionRootDir() & "ET Wip Production Directory\"
End Function

' This sets the production input directory for account administration files
Public Function EtfWipProductionInputDir() As String
    Let EtfWipProductionInputDir = EtfWipProductionDir() & "ET Wip Input Directory\"
End Function

' This sets the production output directory for account administration files
Public Function EtfWipProductionOutputDir() As String
    Let EtfWipProductionOutputDir = EtfWipProductionDir() & "ET Wip Output Directory\"
End Function

' This sets the production intermediate directory for account administration file
Public Function EtfWipProductionIntermediateDir() As String
    Let EtfWipProductionIntermediateDir = EtfWipProductionDir() & "ET Wip Intermediate Directory\"
End Function

'*********************************************************************************************************
'* This section sets the development
'*********************************************************************************************************

' This sets the dev directory on this computer
Public Function DevRootDir() As String
    Let DevRootDir = "C:\Users\PabloA\Documents\QuantModels\Development\Excel\"
    'Let DevRootDir = "U:\WIP_TEST_RUN\"
    'Let DevRootDir = "O:\Development\PabloDevTesting\"
End Function

'This sets the data sets root directory
Public Function DataSetRootDir() As String
    Let DataSetRootDir = "E:\DataSets\"
End Function

Public Function CountryRegionRelationMapFileName() As String
    Let CountryRegionRelationMapFileName = ThisWorkbook.Path() & "\Mappings\Country Region Relation.xlsb"
End Function

Public Function GicsCodeToMsciSubIndustryCodeMapFileName() As String
    Let GicsCodeToMsciSubIndustryCodeMapFileName = ThisWorkbook.Path() & "\Mappings\GicsCodeToMsciSubIndustryCodeMap.xlsb"
End Function

Public Function IsinToAssetClassMapFileName() As String
    Let IsinToAssetClassMapFileName = ThisWorkbook.Path() & "\Mappings\ISIN-to-PiamAssetClassMap.xlsb"
End Function

' This function returns a 1D array with the list of all file names in the given directory.
' The filenames are pathless.
Public Function GetFileNames(ThePath As String, Optional ByVal ThePattern As Variant) As Variant
    Dim FileNamesDict As Dictionary
    Dim i As Long
    Dim AFileName As String
    Dim TheResult() As Variant

    Set FileNamesDict = New Dictionary

    ' Iterate until there are no more files to get
    Let AFileName = Dir(ThePath & IIf(IsMissing(ThePattern), "", ThePattern))
    While Len(AFileName) > 0
        Call FileNamesDict.Add(Key:=AFileName, Item:=AFileName)
        Let AFileName = Dir
    Wend
    
    If FileNamesDict.Count = 0 Then
        Let GetFileNames = Array()
        Exit Function
    End If

    ReDim TheResult(1 To FileNamesDict.Count)
    For i = 1 To FileNamesDict.Count
        Let TheResult(i) = FileNamesDict.Items(i - 1)
    Next i
    
    Let GetFileNames = TheResult
End Function

' This function returns a 1D array with the list of all Non Backup Filenames in the given directory. The filenames are pathless.
' The assumption is that the filenames starting with ~$ are system generated backup files.
Public Function GetNonBackupFileNames(ThePath As String) As Variant
    Dim MyFso As New FileSystemObject
    Dim MyFolder As Folder
    Dim MyFile As File
    Dim FileNames() As String
    Dim i As Integer
    
    Set MyFso = New Scripting.FileSystemObject
    Set MyFolder = MyFso.GetFolder(ThePath)
        
    ' Store all filenames located in ThePath that do not start with ~$
    Let i = 1
    For Each MyFile In MyFolder.Files
        If Not (Left(MyFile.Name, 2) = "~$") Then
            ReDim Preserve FileNames(i)
            Let FileNames(i) = MyFile.Name
            Let i = i + 1
        End If
    Next

    Let GetNonBackupFileNames = FileNames
End Function

' This function extracts the directory from a filename with its full path
Public Function ExtractDirectoryFromFullPathFileName(TheFullPathFileName As String) As String
    Dim TheFileName As String
    Dim TheDirectory As String
    Dim FileNameWithNoExtension As String
    
    ' Split the directory and the filename
    Let TheFileName = StringSplit(TheFullPathFileName, "\", UBound(Split(TheFullPathFileName, "\")) + 1)
    Let TheDirectory = Left(TheFullPathFileName, Len(TheFullPathFileName) - Len(TheFileName))
    ' Strip the extension from the filename
    Let FileNameWithNoExtension = StringSplit(CStr(TheFileName), ".", 1)
    
    ' Set the value to return
    Let ExtractDirectoryFromFullPathFileName = TheDirectory
End Function

' This function extracts the filename and extension from a filename with its full path
Public Function ExtractFilenameAndExtensionFromFullPathFileName(TheFullPathFileName As String) As String
    Dim TheFileName As String
    Dim TheDirectory As String
    Dim FileNameWithNoExtension As String
    
    ' Split the directory and the filename
    Let TheFileName = StringSplit(TheFullPathFileName, "\", UBound(Split(TheFullPathFileName, "\")) + 1)
    Let TheDirectory = Left(TheFullPathFileName, Len(TheFullPathFileName) - Len(TheFileName))
    ' Strip the extension from the filename
    Let FileNameWithNoExtension = StringSplit(CStr(TheFileName), ".", 1)
    
    ' Set the value to return
    Let ExtractFilenameAndExtensionFromFullPathFileName = TheFileName
End Function

' This function extracts the filename with no extension from a filename with its full path
Public Function ExtractFileNameWithNoExtensionFromFullPathFileName(TheFullPathFileName As String) As String
    Dim TheFileName As String
    Dim TheDirectory As String
    Dim FileNameWithNoExtension As String
    
    ' Split the directory and the filename
    Let TheFileName = StringSplit(TheFullPathFileName, "\", UBound(Split(TheFullPathFileName, "\")) + 1)
    Let TheDirectory = Left(TheFullPathFileName, Len(TheFullPathFileName) - Len(TheFileName))
    ' Strip the extension from the filename
    Let FileNameWithNoExtension = StringSplit(CStr(TheFileName), ".", 1)
    
    ' Set the value to return
    Let ExtractFileNameWithNoExtensionFromFullPathFileName = FileNameWithNoExtension
End Function

' This function extracts the extension from a filename with its full path
Public Function ExtractExtensionFromFullPathFileName(TheFullPathFileName As String) As String
    Dim TheFileName As String
    Dim TheDirectory As String
    Dim FileExtension As String
    
    ' Split the directory and the filename
    Let TheFileName = StringSplit(TheFullPathFileName, "\", UBound(Split(TheFullPathFileName, "\")) + 1)
    Let TheDirectory = Left(TheFullPathFileName, Len(TheFullPathFileName) - Len(TheFileName))
    ' Strip the extension from the filename
    Let FileExtension = StringSplit(CStr(TheFileName), ".", 2)
    
    ' Set the value to return
    Let ExtractExtensionFromFullPathFileName = FileExtension
End Function

' Gets executed when clicking MasterControl´s STEP 14 button
'
' The purpose of this routine is to:
' 1. Inject the consolidated screen just before executing the
'    steps for generating the master file
Public Sub ExportListObjectAsTsvFile(AListObject As ListObject, _
                                     TheFullPathFileName As String, _
                                     Optional IncludeHeaderRowQ As Boolean = False)
    Dim r As Long
    Dim c As Long
    Dim ARow As String
    
    ' Open the file
    Open TheFullPathFileName For Output As #1

    If IncludeHeaderRowQ Then
        Let ARow = ""
        For c = 1 To AListObject.ListColumns.Count - 1
            Let ARow = ARow & AListObject.HeaderRowRange(1, c).Value2 & vbTab
        Next c
        Let ARow = ARow & AListObject.HeaderRowRange(1, AListObject.ListColumns.Count).Value2
        Print #1, ARow
    End If
        
    ' Write the listrows
    For r = 1 To AListObject.ListRows.Count
        Let ARow = ""
        For c = 1 To AListObject.ListColumns.Count - 1
            Let ARow = ARow & AListObject.DataBodyRange(r, c).Value2 & vbTab
        Next c
        Let ARow = ARow & AListObject.DataBodyRange(r, AListObject.ListColumns.Count).Value2
        Print #1, ARow
    Next r
    
    ' Close the file
    Close #1
End Sub

