Attribute VB_Name = "Files"
Option Explicit
Option Base 1

' DESCRIPTION
' This function returns a 1D array with the list of all directories
' names in the given directory or all directories with the given pattern.
' If the optional parameter AddPathQ is set to True, the path is prepended
' to all directory names.
'
' The function returns Null if the directory does not exists to differentiate
' from the case where the directory has no subdirectories.
'
' ThePath may include wildchars
'
' PARAMETERS
' 1. ThePath - A filename path, possible including wildchars
' 2. AddPathQ - Set to False by default. When False, the pathless directory
'    names are returned. When set to True, the full-path directory names
'    are returned.
'
' RETURNED VALUE
' An array with the names of the directories inside the given path if the
' directory exists. The function returns Null if the directory does not
' exists.
Public Function GetDirectories(ByVal ThePath As String, _
                               Optional AddPathQ As Boolean = False) As Variant
    Dim FileNamesDict As Dictionary
    Dim aFileName As String
    
    ' Trim ThePath
    Let ThePath = Trim(ThePath)

    ' Exit with Null if the directory does not exists
    Let GetDirectories = Null
    
    ' Remove the last section of the directory path if it has wildchars
    If InStr(FileBaseName(ThePath), "*") + InStr(FileBaseName(ThePath), "?") > 0 Then
        If Not DirectoryExistsQ(FileNameJoin(Most(FileNameSplit(DirectoryName(ThePath))))) Then Exit Function
    Else
        If Not DirectoryExistsQ(ThePath) Then Exit Function
    End If

    Set FileNamesDict = New Dictionary

    ' Iterate until there are no more files to get
    ' Ensure the directory name ends with "\". Otherwise, you can end up with one subdiretory
    ' in an empty directory
    If InStr(FileBaseName(ThePath), "*") + InStr(FileBaseName(ThePath), "?") > 0 Then
        If Right(ThePath, 1) = "\" Then
            Let aFileName = Dir(Left(ThePath, Len(ThePath) - 1), vbDirectory + vbHidden)
        Else
            Let aFileName = Dir(ThePath, vbDirectory + vbHidden)
        End If
    Else
        Let aFileName = Dir(ThePath, vbDirectory + vbHidden) & IIf(Right(ThePath, 1) <> "\", "\", "")
    End If
                        
    While Len(aFileName) > 0
        If AddPathQ And FreeQ(Array(".", ".."), aFileName) Then
            Call FileNamesDict.Add(Key:=aFileName, _
                                   Item:=FileNameJoin(Array(ThePath, aFileName, "\")))
        ElseIf FreeQ(Array(".", ".."), aFileName) Then
            Call FileNamesDict.Add(Key:=aFileName, Item:=aFileName)
        End If
        
        Let aFileName = Dir
    Wend
    
    If FileNamesDict.Count = 0 Then
        Let GetDirectories = EmptyArray()
        Exit Function
    End If
    
    Let GetDirectories = Flatten(FileNamesDict.Items)
End Function

' DESCRIPTION
' This function returns a 1D array with the list of all filenames
' matching the given path. If the optional parameter AddPathQ is set
' to True, the path is prepended to all directory names.
'
' ThePath may include wildchars
'
' PARAMETERS
' 1. ThePath - A filename path, possible including wildchars
' 2. AddPathQ - Set to False by default. When False, the pathless directory
'    names are returned. When set to True, the full-path directory names
'    are returned.
'
' RETURNED VALUE
' An array with the names of the files matching the given pattern
Public Function GetFileNames(ThePath As String, _
                             Optional AddPathQ As Boolean = False) As Variant
    Dim FileNamesDict As Dictionary
    Dim i As Long
    Dim aFileName As String
    Dim TheResult() As Variant

    On Error Resume Next

    Set FileNamesDict = New Dictionary

    ' Iterate until there are no more files to get
    Let aFileName = Dir(ThePath, vbHidden + vbNormal + vbReadOnly + vbSystem)
    While Len(aFileName) > 0
        If AddPathQ Then
            If InStr(FileBaseName(ThePath), "*") + InStr(FileBaseName(ThePath), "?") > 0 Then
                Call FileNamesDict.Add(Key:=aFileName, _
                                       Item:=FileNameJoin(Array(FileNameJoin(Most(FileNameSplit(ThePath))), aFileName)))
            Else
                Call FileNamesDict.Add(Key:=aFileName, _
                                       Item:=FileNameJoin(Array(ThePath, aFileName)))
            End If
        Else
            Call FileNamesDict.Add(Key:=aFileName, Item:=aFileName)
        End If
        
        Let aFileName = Dir
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

' DESCRIPTION
' This function returns a properly formatted, full-path file or directory name.
'
' INPUT
' 1. ArrayOfDirNames - a 1D array of the components of the directory or file's path
'
' OUTPUT
' A properly formatted, full-path file or directory name
Public Function FileNameJoin(ArrayOfFileNames As Variant, Optional ParamConsistencyChecksQ = False) As String
    Dim c As Long
    Dim ResultArray As Variant

    ' Default return value
    Let FileNameJoin = vbNullString
    
    ' Parameter check. If parameters inconsistent, exit with empty string
    If ParamConsistencyChecksQ Then
        If Not DimensionedQ(ArrayOfFileNames) Then Exit Function
        If Not StringArrayQ(ArrayOfFileNames) Then Exit Function
    End If

    ' Exit with empty string if there is nothing to join
    If EmptyArrayQ(ArrayOfFileNames) Then Exit Function
    
    Let ResultArray = EmptyArray()
    For c = LBound(ArrayOfFileNames) To UBound(ArrayOfFileNames)
        Let ResultArray = ConcatenateArrays(ResultArray, _
                                            ComplementOfSets(Split(ArrayOfFileNames(c), Application.PathSeparator), _
                                                             Array(vbNullString)))
    Next
    
    Let FileNameJoin = Join(ResultArray, Application.PathSeparator)
    
    If Left(Trim(First(ArrayOfFileNames)), 2) = "\\" Then Let FileNameJoin = "\\" & FileNameJoin
End Function

' DESCRIPTION
' This function splits the given filename into its components.
'
' INPUT
' 1. AFileName - a filepath
'
' OUTPUT
' A 1D array with the parts of the array
Public Function FileNameSplit(aFileName As String) As Variant
    Let FileNameSplit = vbNullString

    If aFileName = vbNullString Then Exit Function
    
    Let FileNameSplit = Split(aFileName, Application.PathSeparator)
    
    If Left(aFileName, 2) = "\\" Then
        Let FileNameSplit = Prepend(FileNameSplit, "\\")
    End If
    
    Let FileNameSplit = ComplementOfSets(FileNameSplit, Array(vbNullString))
End Function

' DESCRIPTION
' This function returns the directory corresponding to the file name.
' If the filename has no explicit directory, this function returns
' vbNullString
'
' INPUT
' 1. AFileName - a filepath
'
' OUTPUT
' The directory in the filename or vbNullString if no directory is explicit
Public Function DirectoryName(aFileName As String) As String
    Dim TheParts As Variant
    
    Let DirectoryName = vbNullString

    If aFileName = vbNullString Then Exit Function
    
    Let TheParts = Split(aFileName, "\")
    
    If Length(TheParts) = 1 Then
        Let DirectoryName = vbNullString
    Else
        Let DirectoryName = FileNameJoin(Most(TheParts))
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
Public Function FileExtension(aFileName As String) As String
    Dim TheParts As Variant

    Let FileExtension = vbNullString

    If aFileName = vbNullString Then Exit Function
    
    Let TheParts = Split(Last(FileNameSplit(aFileName)), ".")
    
    If Length(TheParts) = 1 Then
        Let FileExtension = Empty
    Else
        Let FileExtension = Last(TheParts)
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
Public Function FileBaseName(aFileName As String) As Variant
    If IsNull(aFileName) Or IsEmpty(aFileName) Or aFileName = Empty Then
        Let FileBaseName = Null
        Exit Function
    End If
    
    Let FileBaseName = Split(Last(FileNameSplit(aFileName)), ".")
    
    If Length(FileBaseName) = 1 Then
        Let FileBaseName = First(FileBaseName)
    Else
        Let FileBaseName = Join(Most(FileBaseName), ".")
    End If
End Function

' DESCRIPTION
' This sub deletes all files in the given directory
'
' INPUT
' 1. ThePath - a filepath
'
' OUTPUT
' The chosen directory is cleared
Public Sub ClearDirectory(ThePath As String)
    Dim TheFileNames As Variant
    
    Let TheFileNames = GetFileNames(FileNameJoin(Array(ThePath, "*.*")))
    If Not EmptyArrayQ(TheFileNames) Then
        Kill FileNameJoin(Array(ThePath, "*.*"))
    End If
End Sub

' DESCRIPTION
' This function returns the full path a workbook.
'
' PARAMETERS
' 1. Wbk - A reference to a workbook
'
' RETURNED VALUE
' Full path of the given workbook
Public Function WorkbookFullPath(Wbk As Workbook) As String
    Let WorkbookFullPath = FileNameJoin(Array(Wbk.Path, ThisWorkbook.Name))
End Function

