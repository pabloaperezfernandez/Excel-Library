Attribute VB_Name = "Files"
Option Explicit
Option Base 1

' This function returns a 1D array with the list of all file names in the given directory.
' The filenames are pathless.
Public Function GetFileNames(ThePath As String) As Variant
    Dim FileNamesDict As Dictionary
    Dim i As Long
    Dim AFileName As String
    Dim TheResult() As Variant

    Set FileNamesDict = New Dictionary

    ' Iterate until there are no more files to get
    Let AFileName = dir(ThePath)
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
Public Function FileNameSplit(AFileName As String) As Variant
    Let FileNameSplit = vbNullString

    If AFileName = vbNullString Then Exit Function
    
    Let FileNameSplit = Split(AFileName, Application.PathSeparator)
    
    If Left(AFileName, 2) = "\\" Then
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
Public Function DirectoryName(AFileName As String) As String
    Let DirectoryName = vbNullString

    If AFileName = vbNullString Then Exit Function
    
    Let DirectoryName = Split(AFileName, "\")
    
    If Length(DirectoryName) = 1 Then
        Let DirectoryName = vbNullString
    Else
        Let DirectoryName = FileNameJoin(Most(DirectoryName))
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
Public Function FileExtension(AFileName As String) As String
    Dim TheParts As Variant

    Let FileExtension = vbNullString

    If AFileName = vbNullString Then Exit Function
    
    Let TheParts = Split(Last(FileNameSplit(AFileName)), ".")
    
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
Public Function FileBaseName(AFileName As String) As Variant
    If IsNull(AFileName) Or IsEmpty(AFileName) Or AFileName = Empty Then
        Let FileBaseName = Null
        Exit Function
    End If
    
    Let FileBaseName = Split(Last(FileNameSplit(AFileName)), ".")
    
    If Length(FileBaseName) = 1 Then
        Let FileBaseName = First(FileBaseName)
    Else
        Let FileBaseName = Join(Most(FileBaseName), ".")
    End If
End Function
