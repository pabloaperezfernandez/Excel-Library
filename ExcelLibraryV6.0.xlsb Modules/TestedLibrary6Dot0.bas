Attribute VB_Name = "TestedLibrary6Dot0"
Option Explicit
Option Base 1

Public Sub TestFileNameJoin()
    Dim AnArray As Variant
    
    Let AnArray = Array("c:", "dir1", "dir2")
    Debug.Print "The result is " & FileNameJoin(AnArray)
    
    Let AnArray = Array("c:", "dir1", "file.txt")
    Debug.Print "The result is " & FileNameJoin(AnArray)
    
    Let AnArray = Array("c:", "dir1", "file.txt")
    Debug.Print "The result is " & FileNameJoin(AnArray)
    
    Let AnArray = Array()
    Debug.Print "The result is " & FileNameJoin(AnArray)
End Sub

Public Sub TestFileNameSplit()
    Dim AnArray As Variant
    
    Let AnArray = Array("c:", "dir1", "dir2")
    Debug.Print "The result is"
    PrintArray FileNameSplit(FileNameJoin(AnArray))
    Debug.Print
    
    Let AnArray = Array("c:", "dir1", "file.txt")
    Debug.Print "The result is"
    PrintArray FileNameSplit(FileNameJoin(AnArray))
    Debug.Print
    
    Let AnArray = Empty
    Debug.Print "The result is"
    PrintArray FileNameSplit(Empty)
    Debug.Print "IsNull(FileNameSplit(Empty)) = " & IsNull(FileNameSplit(Empty))
    Debug.Print
End Sub

Public Sub TestFileBaseName()
    Dim var As Variant

    For Each var In Array("c:\dir1\dir2\base.ext1.ext2", _
                          "c:\dir\base.txt", _
                          "c:\dir\base", _
                          "base.txt", _
                          "base")
        Debug.Print "For " & var & " the base name is -" & FileBaseName(CStr(var)) & "-"
    Next
End Sub

Public Sub TestFileExtension()
    Dim var As Variant

    For Each var In Array("c:\dir1\dir2\base.ext1.ext2", _
                          "c:\dir\base.txt", _
                          "c:\dir\base", _
                          "base.txt", _
                          "base")
        Debug.Print "For " & var & " the extension is -" & FileExtension(CStr(var)) & "-"
    Next
End Sub
