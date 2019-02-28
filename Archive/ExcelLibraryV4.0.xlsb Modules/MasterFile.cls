VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MasterControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Option Base 1

Public Sub CreateDirectoryStructure()
    Dim aListObj As ListObject
    Dim currentDirName As String
    Dim currentLevel As Integer
    Dim priorLevel As Integer
    Dim i As Integer
    Dim j As Integer
    
    Set aListObj = DirectoryHierarchy.ListObjects("DirectoryHierarchy")
    
    Call ChDir(MasterControl.Range("TreeRoot").Value2)
    Let priorLevel = 1
    For i = 1 To aListObj.ListRows.Count
        Let currentLevel = aListObj.ListColumns("Level").DataBodyRange.Cells(i, 1).Value2
        Let currentDirName = aListObj.ListColumns("DIR Name").DataBodyRange.Cells(i, 1).Value2
    
        If currentLevel < priorLevel Then
            For j = currentLevel To priorLevel
                Call ChDir("..")
            Next j
        ElseIf currentLevel = priorLevel And currentLevel <> 1 Then
            Call ChDir("..")
        End If
        
        Let priorLevel = currentLevel
        Call MkDir(currentDirName)
        Call ChDir(currentDirName)
        
        Debug.Print Application.WorksheetFunction.Rept("  ", currentLevel) & currentDirName & ", DIR #" & i & ", level = " & currentLevel
    Next i
End Sub

Public Sub CopyFilesToDirectories()
    Dim dirListObject As ListObject
    Dim progListObject As ListObject
    Dim originalName As String
    Dim newName As String
    Dim pkTargetDirectory As Integer
    Dim dirIndex As Integer
    Dim progIndex As Integer
    Dim rootDir As String
    Dim currentDirName As String
    Dim aProgramList As Variant
    Dim currentLevel As Integer
    Dim priorLevel As Integer
    Dim j As Integer
    
    Let rootDir = MasterControl.Range("ProgramsSourceDir").Value2
    Set dirListObject = DirectoryHierarchy.ListObjects("DirectoryHierarchy")
    Set progListObject = FileLocations.ListObjects("ProgramLocations")
    Let aProgramList = UniqueSubset(ConvertTo1DArray(progListObject.ListColumns("Original Application Name").DataBodyRange.Value2))

    Let progIndex = 1
    
    Call ChDir(MasterControl.Range("TreeRoot").Value2)
    Let priorLevel = 1
    Let progIndex = 1
    
    Let dirIndex = 1
    While (dirIndex < dirListObject.ListRows.Count)
        Let currentLevel = dirListObject.ListColumns("Level").DataBodyRange.Cells(dirIndex, 1).Value2
        Let currentDirName = dirListObject.ListColumns("DIR Name").DataBodyRange.Cells(dirIndex, 1).Value2
    
        If currentLevel < priorLevel Then
            For j = currentLevel To priorLevel
                Call ChDir("..")
            Next j
        ElseIf currentLevel = priorLevel And currentLevel <> 1 Then
            Call ChDir("..")
        End If
        
        Let originalName = progListObject.ListColumns("Original Application Name").DataBodyRange.Cells(progIndex, 1).Value2
        Let newName = progListObject.ListColumns("New Application Name").DataBodyRange.Cells(progIndex, 1).Value2
        Let pkTargetDirectory = progListObject.ListColumns("PK Target Directory").DataBodyRange.Cells(progIndex, 1).Value2
        
        Call ChDir(currentDirName)
        
        If dirIndex = pkTargetDirectory Then
            Call FileCopy(rootDir & originalName, CurDir() & "\" & newName)
            Let progIndex = progIndex + 1
            
            Debug.Print "Wrote " & originalName & " to " & CurDir() & "\" & newName
        Else
            Debug.Print "We just processed directory with index " & dirIndex
            Let dirIndex = dirIndex + 1
        End If
        
        Let priorLevel = currentLevel
    Wend
End Sub

Public Sub AddLibraryReferences()
    Dim dirListObject As ListObject
    Dim progListObject As ListObject
    Dim originalName As String
    Dim newName As String
    Dim pkTargetDirectory As Integer
    Dim dirIndex As Integer
    Dim progIndex As Integer
    Dim rootDir As String
    Dim environmentRoot As String
    Dim currentDirName As String
    Dim aProgramList As Variant
    Dim currentLevel As Integer
    Dim priorLevel As Integer
    Dim j As Integer
    Dim vbReference As Variant
    Dim vbReferenceFilename As String
    Dim wbk As Workbook
    
    Let Application.EnableEvents = False
    Let Application.ScreenUpdating = False
    
    On Error GoTo ErrorTrap
    
    Let rootDir = MasterControl.Range("ProgramsSourceDir").Value2
    Let environmentRoot = MasterControl.Range("TreeRoot").Value2
    Set dirListObject = DirectoryHierarchy.ListObjects("DirectoryHierarchy")
    Set progListObject = FileLocations.ListObjects("ProgramLocations")
    Let aProgramList = UniqueSubset(ConvertTo1DArray(progListObject.ListColumns("Original Application Name").DataBodyRange.Value2))

    Let progIndex = 1
    
    Call ChDir(MasterControl.Range("TreeRoot").Value2)
    Let priorLevel = 1
    Let progIndex = 1
    
    Let dirIndex = 1
    While (dirIndex < dirListObject.ListRows.Count And progIndex < progListObject.ListRows.Count)
        Let currentLevel = dirListObject.ListColumns("Level").DataBodyRange.Cells(dirIndex, 1).Value2
        Let currentDirName = dirListObject.ListColumns("DIR Name").DataBodyRange.Cells(dirIndex, 1).Value2
    
        If currentLevel < priorLevel Then
            For j = currentLevel To priorLevel
                Call ChDir("..")
            Next j
        ElseIf currentLevel = priorLevel And currentLevel <> 1 Then
            Call ChDir("..")
        End If
        
        Let originalName = progListObject.ListColumns("Original Application Name").DataBodyRange.Cells(progIndex, 1).Value2
        Let newName = progListObject.ListColumns("New Application Name").DataBodyRange.Cells(progIndex, 1).Value2
        Let pkTargetDirectory = progListObject.ListColumns("PK Target Directory").DataBodyRange.Cells(progIndex, 1).Value2
        Let vbReferenceFilename = progListObject.ListColumns("Library").DataBodyRange.Cells(progIndex, 1).Value2
        
        Call ChDir(currentDirName)
        
        If dirIndex = pkTargetDirectory Then
            If pkTargetDirectory <> 2 Then
                Debug.Print "Just opened " & CurDir() & currentDirName & "\" & newName & " to add library references."
                Set wbk = Application.Workbooks.Open(Filename:=newName)
                
                Let vbReferenceFilename = environmentRoot & "Production IM\Common Software Directory\" & vbReferenceFilename
                Call wbk.VBProject.References.AddFromFile(vbReferenceFilename)

                Call wbk.Close(SaveChanges:=True)
            
                Debug.Print "Added library reference to " & newName
            End If
            
            Let progIndex = progIndex + 1
        Else
            Let dirIndex = dirIndex + 1
        End If
        
        Let priorLevel = currentLevel
    Wend

    Let Application.ScreenUpdating = True
    Let Application.EnableEvents = True
    
    Exit Sub
    
ErrorTrap:
    MsgBox Err.Description
End Sub
