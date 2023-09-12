Attribute VB_Name = "Main_Files"
' PURPOSE OF THIS MODULE
'
' The purpose of this module is to provide faciities to manipulate files
' and filenames.
Option Explicit
Option Base 1


' DESCRIPTION
' Returns a filename with full path for a file in the targer SharePoint
' directory specified in named range Worksheets("UI").Range("SharePointDirectoryPath")
'
' PARAMETERS
' 1. workbookName (string) - a name of the file
'
' RETURNED VALUE
' Returns a filename with the full path
Public Function GetInputWorkbookName(workbookName As String) As String
    Dim SharePointUrl As Variant
    Dim SharePointDirectory As Variant
    Dim StringsArray As Variant
    
    With ThisWorkbook
        Let SharePointUrl = FileNameSplit(.Worksheets("UI").Range("SharePointSiteUrl").Value2, True)
        Let SharePointDirectory = FileNameSplit(.Worksheets("UI").Range("SharePoinSiteDirectory").Value2, True)
    End With

    Let StringsArray = ConcatenateArrays(SharePointUrl, SharePointDirectory)
    Let StringsArray = ConcatenateArrays(StringsArray, Array(workbookName))
    Let GetInputWorkbookName = FileNameJoin(StringsArray, UnixPathSeparatorQ:=True)
End Function

' DESCRIPTION
' Returns a filename with full path for a file in the targer SharePoint
' directory specified in named range Worksheets("UI").Range("SharePointDirectoryPath")
'
' PARAMETERS
' 1. workbookName (string) - a name of the file
'
' RETURNED VALUE
' Returns a filename with the full path
Public Function MakeWorkbookArchivalFileName(wbkName As String) As String
    Dim theDateTime As Date
    Dim dateSuffix As String
    Dim timeSuffix As String
    Dim theFileName As String
    Dim FilePath As String

    ' Pull the current system date/time
    Let theDateTime = Now()
    ' Convert the current date/time to a serial year string
    Let dateSuffix = Format(theDateTime, "YYYYMMDD")
    ' Convert the current date/time to a serial time string
    Let timeSuffix = Format(theDateTime, "HHMMDD")
    
    ' Create the filename and path
    Let FilePath = FileNameJoin(Array(Environ("HOMEDRIVE"), Environ("HOMEPATH"), "Downloads"))
    Let theFileName = Join(Array(FileBaseName(wbkName), dateSuffix, timeSuffix, "." & FileExtension(wbkName)), "-")
    
    ' Return fully qualified filename (i.e., includes path)
    Let MakeWorkbookArchivalFileName = FileNameJoin(Array(FilePath, theFileName))
End Function
