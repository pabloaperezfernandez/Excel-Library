Attribute VB_Name = "XML"
Option Explicit
Option Base 1

' This function returns the first element with the given tag.
' Opening and closing tags are included if the optional argument IncludeTags is set to True.
' By default, IncludeTags is set to false. This function assumes that the XML is well-form
' Also, it assumes that there are no extra spaces in closing tags before ">"
' We don't allow the <xml ...> tag usually placed at the beginning of an XML document
Public Function GetXmlElement(TheString As String, TheTag As String, _
                              Optional IncludeTags As Boolean = False) As String
    Dim StartPos As Long
    Dim EndingPos As Long
    Dim NextOpeningPos As Long
    Dim NextClosingPos As Long
    Dim OpeningTagCount As Long
    Dim ClosingTagCount As Long
    Dim FixedTag As String
    Dim i As Integer
    Dim CurrentPos As Long
    
    ' Make sure FixedTag contains no "<" or ">."  We want the name of the tag
    ' Make sure TheTag has been trimmed
    Let FixedTag = Replace(Replace(Trim(TheTag), "<", ""), ">", "")
    
    Let OpeningTagCount = GetArrayLength(Split(TheString, "<" & FixedTag))
    Let ClosingTagCount = GetArrayLength(Split(TheString, "</" & FixedTag & ">"))
    
    If OpeningTagCount <= 1 Or (OpeningTagCount > 1 And OpeningTagCount <> ClosingTagCount) Then
        Let GetXmlElement = ""
        Exit Function
    End If
    
    ' Record the starting position of the opening tag
    Let OpeningTagCount = 0
    Let ClosingTagCount = 0
    Let StartPos = InStr(Trim(TheString), "<" & FixedTag)
    Let CurrentPos = StartPos
    While CurrentPos < Len(Trim(TheString)) And ((OpeningTagCount = 0 And ClosingTagCount = 0) Or OpeningTagCount <> ClosingTagCount)
        Let NextOpeningPos = InStr(CurrentPos, Trim(TheString), "<" & FixedTag)
        Let NextClosingPos = InStr(CurrentPos, Trim(TheString), "</" & FixedTag & ">")
        
        If NextOpeningPos + NextClosingPos = 0 Then
            Let CurrentPos = Len(Trim(TheString)) + 1
        End If
        
        If (OpeningTagCount <> ClosingTagCount And NextClosingPos = 0) Or (OpeningTagCount + ClosingTagCount = 0 And NextClosingPos = 0) Then
            Let GetXmlElement = ""
            Exit Function
        ElseIf NextOpeningPos <> 0 And NextOpeningPos < NextClosingPos Then
            Let OpeningTagCount = OpeningTagCount + 1
            Let CurrentPos = NextOpeningPos + 1
        ElseIf NextOpeningPos = 0 And NextClosingPos <> 0 Then
            Let ClosingTagCount = ClosingTagCount + 1
            Let CurrentPos = NextClosingPos + 1
        ElseIf NextClosingPos <> 0 And NextClosingPos < NextOpeningPos Then
            Let ClosingTagCount = ClosingTagCount + 1
            Let CurrentPos = NextClosingPos + 1
        ElseIf NextClosingPos = 0 And NextOpeningPos <> 0 Then
            Let OpeningTagCount = OpeningTagCount + 1
            Let CurrentPos = NextOpeningPos + 1
        End If
    Wend
    
    If OpeningTagCount = ClosingTagCount And OpeningTagCount + ClosingTagCount <> 0 And IncludeTags = True Then
        Let GetXmlElement = Mid(Trim(TheString), StartPos, CurrentPos - StartPos + Len("</" & FixedTag & ">") - 1)
    ElseIf OpeningTagCount = ClosingTagCount And OpeningTagCount + ClosingTagCount <> 0 And IncludeTags = False Then
        Let EndingPos = InStr(StartPos, Trim(TheString), ">")
    
        Let GetXmlElement = Mid(Trim(TheString), EndingPos + 1, CurrentPos - EndingPos - 2)
    Else
        Let GetXmlElement = ""
    End If
End Function

Public Function GetFirstXmlBracketExpression(TheString As String) As String
    Dim StartPos As Long
    Dim EndPos As Long
    
    Let StartPos = InStr(TheString, "<")
    Let EndPos = InStr(StartPos, TheString, ">")
    
    If StartPos = 0 Or EndPos = 0 Then
        Let GetFirstXmlBracketExpression = ""
        Exit Function
    End If
    
    Let GetFirstXmlBracketExpression = "<" & Trim(Mid(TheString, StartPos + 1, EndPos - StartPos - 1)) & ">"
End Function
