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
            Let aFileName = dir(Left(ThePath, Len(ThePath) - 1), vbDirectory + vbHidden)
        Else
            Let aFileName = dir(ThePath, vbDirectory + vbHidden)
        End If
    Else
        Let aFileName = dir(ThePath, vbDirectory + vbHidden) & IIf(Right(ThePath, 1) <> "\", "\", "")
    End If
                        
    While Len(aFileName) > 0
        If AddPathQ And FreeQ(Array(".", ".."), aFileName) Then
            Call FileNamesDict.Add(Key:=aFileName, _
                                   Item:=FileNameJoin(Array(ThePath, aFileName, "\")))
        ElseIf FreeQ(Array(".", ".."), aFileName) Then
            Call FileNamesDict.Add(Key:=aFileName, Item:=aFileName)
        End If
        
        Let aFileName = dir
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
    Let aFileName = dir(ThePath, vbHidden + vbNormal + vbReadOnly + vbSystem)
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
        
        Let aFileName = dir
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
Public Function FileNameJoin(ArrayOfFileNames As Variant, Optional ParamConsistencyChecksQ As Boolean = False) As String
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
        If StrComp(ArrayOfFileNames(c), vbNullString) Then
            Let ResultArray = ConcatenateArrays(ResultArray, _
                                                FilterValueFromArray(Split(ArrayOfFileNames(c), Application.PathSeparator), _
                                                                     vbNullString))
        End If
    Next
    
    ' Exit if there was an error
    If NullQ(ResultArray) Then Exit Function
    
    ' Join the parts and separate them with the file separator string
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
Public Function WorkbookFullPath(wbk As Workbook) As String
    Let WorkbookFullPath = FileNameJoin(Array(wbk.Path, ThisWorkbook.Name))
End Function

' DESCRIPTION
' Returns the list of all filenames contained in a given directory
'
' PARAMETERS
' 1. sPath - The directory path from where to get filenames.
'
' RETURNED VALUE
' The list of filenames
Public Function GetFileNamesInTree(ByVal sPath As String) As Variant
    ' attribute mask
    Const iAttr     As Long = vbNormal + vbReadOnly + vbSystem + vbDirectory
    Dim jAttr       As Long        ' file attributes
    Dim Col         As Collection  ' queued directories
    Dim iFile       As Long        ' file counter
    Dim sFile       As String      ' file name
    Dim sName       As String      ' full file name
    
    Let Application.ScreenUpdating = False
    
    Let iFile = 1

    With TempComputation.UsedRange: Call .ClearFormats: Call .ClearContents: End With

    If Right(sPath, 1) <> "\" Then sPath = sPath & "\"
    
    ' Create a collection to hold all files under a directory
    Set Col = New Collection
    
    ' Add the give path to the collection as the first path
    Call Col.Add(sPath)

    ' While there are paths to process, continue
    Do While Col.Count
        ' Start with the path at the root of the collection
        Let sPath = Col(1)
        
        On Error Resume Next
        Let sFile = dir(sPath, iAttr)

        Do While Len(sFile)
            Let sName = sPath & sFile

            On Error Resume Next
            Let jAttr = GetAttr(sName)

            If Err.Number Then
                ' You can't get attributes for files with Unicode characters in
                ' the name, or some particular files (e.g., "C:\System Volume Information")
                Debug.Print sName
                Err.Clear
            Else
                On Error GoTo 0
                If jAttr And vbDirectory Then
                    If Right(sName, 1) <> "." Then Call Col.Add(sName & "\")
                Else
                    Let TempComputation.Range("A1").Offset(iFile - 1).Value2 = sName
                    Let iFile = iFile + 1
                End If
            End If
            
            ' Get the next file in the current directory
            Let sFile = dir()
        Loop
        
        Call Col.Remove(1)
    Loop

    Call TempComputation.Range("A1").CurrentRegion.Sort(Key1:=TempComputation.Range("A1"), Header:=xlYes)
    
    ' Return the list of filenames
    Let GetFileNamesInTree = Flatten(TempComputation.Range("A1").CurrentRegion.Value2)
    Call TempComputation.UsedRange.ClearFormats
    Call TempComputation.UsedRange.ClearContents
    
    Application.ScreenUpdating = True
End Function

' DESCRIPTION
' Returns the sequence of directories increasing to the given one. For example,
' given C:\Dir1\Dir2\Dir3 this function returns
'
' Array("C:", "C:\Dir1", "C:\Dir1\Dir2", "C:\Dir1\Dir2\Dir3")
'
' PARAMETERS
' 1. ADirPath - A directory path
'
' RETURNED VALUE
' The sequence of paths increasing to the given one
Public Function DirectorySequence(ADirPath As String) As Variant
    Dim CurrentDir As String
    Dim CurrentDirPart As Variant
    Dim ADict As Dictionary
    
    Set ADict = New Dictionary
    
    Let CurrentDir = vbNullString
    For Each CurrentDirPart In FileNameSplit(ADirPath)
        Let CurrentDir = FileNameJoin(Array(CurrentDir, CurrentDirPart))
        Call ADict.Add(Key:=CurrentDir, Item:=Null)
    Next
    
    Let DirectorySequence = Flatten(ADict.Keys)
End Function

' DESCRIPTION
' Creates the directory structure specified by the user.
'
' 1. Root
' 2. DirectoryTree - An element satisfying DirectoryTreeQ. For example:
'    - Array("Dir")
'    - Array("Dir", Array())
'    - Array("Dir1", Array(...)), where ... is of the types above
'
' PARAMETERS
' Returns Null is case of error and True if successful
'
' RETURNED VALUE
' The requested directory structure is created.
Public Function CreateDirectoryTree(Root As String, DirectoryTree As Variant) As Variant
    Dim var As Variant
    Dim FullPath As String

    Let CreateDirectoryTree = Null
    
    If Not DirectoryTreeQ(DirectoryTree) Then Exit Function
    
    Let FullPath = FileNameJoin(Array(Root, First(DirectoryTree)))
    If DirectoryExistsQ(FullPath) Then Exit Function
    MkDir FullPath
    
    If Length(DirectoryTree) = 2 Then
        For Each var In Last(DirectoryTree)
            Let CreateDirectoryTree = CreateDirectoryTree(FullPath, var)
            If NullQ(CreateDirectoryTree) Then Exit Function
        Next
    End If
    
    Let CreateDirectoryTree = True
End Function

' DESCRIPTION
' Creates the basic directory tree used by an Excel application. The following
' directories are created:
' 1. Code
' 2. Inputs
' 3. Output
' 4. Archive
' 5. Documentation
'
' PARAMETERS
' Returns Null is case of error and True if successful
'
' RETURNED VALUE
' The requested directory structure is created.
Public Function CreateAppDirectoryTree(RootAppDir As String) As Variant
    Let CreateAppDirectoryTree = CreateDirectoryTree(RootAppDir, _
                                                     Array(".", Array(Array("Code"), _
                                                               Array("Inputs"), _
                                                               Array("Outputs"), _
                                                               Array("Documentation"))))
End Function
