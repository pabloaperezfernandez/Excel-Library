Attribute VB_Name = "Files"
Option Explicit
Option Base 1

' This function returns a 1D array with the list of all file names in the given directory.
' The filenames are pathless.
Public Function GetFileNames(ThePath As String, Optional ByVal ThePattern As Variant) As Variant
    Dim FileNamesDict As Dictionary
    Dim i As Long
    Dim AFileName As String
    Dim TheResult() As Variant

    Set FileNamesDict = New Dictionary

    ' Iterate until there are no more files to get
    Let AFileName = dir(ThePath & IIf(IsMissing(ThePattern), "", ThePattern))
    While Len(AFileName) > 0
        Call FileNamesDict.Add(Key:=AFileName, Item:=AFileName)
        Let AFileName = dir
    Wend
    
    If FileNamesDict.Count = 0 Then
        Let GetFileNames = EmptyArray()
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

' DESCRIPTION
' This function returns a properly formatted, full-path file or directory name.
'
' INPUT
' 1. ArrayOfDirNames - a 1D array of the components of the directory or file's path
'
' OUTPUT
' A properly formatted, full-path file or directory name
Public Function FileNameJoin(ArrayOfFileNames As Variant) As Variant
    If Not IsArray(ArrayOfFileNames) Or IsNull(ArrayOfFileNames) Then
        Let FileNameJoin = Null
        Exit Function
    End If
    
    If Not DimensionedQ(ArrayOfFileNames) Then
        Let FileNameJoin = Null
        Exit Function
    End If
    
    If EmptyArrayQ(ArrayOfFileNames) Then
        Let FileNameJoin = Empty
        Exit Function
    End If
        
    Let FileNameJoin = Join(ArrayOfFileNames, Application.PathSeparator)
End Function

' DESCRIPTION
' This function splits the given filename into its components.
'
' INPUT
' 1. AFileName - a filepath
'
' OUTPUT
' A 1D array with the parts of the array
Public Function FileNameSplit(AFileName As String) As Variant
    If IsNull(AFileName) Or IsEmpty(AFileName) Or AFileName = Empty Then
        Let FileNameSplit = Null
        Exit Function
    End If
    
    Let FileNameSplit = Split(AFileName, Application.PathSeparator)
End Function

' DESCRIPTION
' This function splits the given filename into its components.
'
' INPUT
' 1. AFileName - a filepath
'
' OUTPUT
' A 1D array with the parts of the array
Public Function FileExtension(AFileName As String) As Variant
    If IsNull(AFileName) Or IsEmpty(AFileName) Or AFileName = Empty Then
        Let FileExtension = Null
        Exit Function
    End If
    
    Let FileExtension = Split(Last(FileNameSplit(AFileName)), ".")
    
    If GetArrayLength(FileExtension) = 1 Then
        Let FileExtension = Empty
    Else
        Let FileExtension = Last(FileExtension)
    End If
End Function

' DESCRIPTION
' This function splits the given filename into its components.
'
' INPUT
' 1. AFileName - a filepath
'
' OUTPUT
' A 1D array with the parts of the array
Public Function FileBaseName(AFileName As String) As Variant
    If IsNull(AFileName) Or IsEmpty(AFileName) Or AFileName = Empty Then
        Let FileBaseName = Null
        Exit Function
    End If
    
    Let FileBaseName = Split(Last(FileNameSplit(AFileName)), ".")
    
    If GetArrayLength(FileBaseName) = 1 Then
        Let FileBaseName = First(FileBaseName)
    Else
        Let FileBaseName = Join(Most(FileBaseName), ".")
    End If
End Function
