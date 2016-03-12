Attribute VB_Name = "IconSetFormatting"
Sub PasteIconSetRule()
    For Each Item In Selection
        Call SetIconSetRule(Item.Address, Item.Offset(0, 1).Address)
    Next Item
End Sub

Sub PasteIconSetRule2()
    For Each Item In Selection
        Call SetIconSetRule(Item.Address, Item.Offset(0, 0).Address)
    Next Item
End Sub

' This module inserts an icon set rule since copying of icon set rules does not work all that well.
Private Sub SetIconSetRule(TheAddress As String, TheOtherCell As String)
    Range(TheAddress).Select
    Selection.FormatConditions.AddIconSetCondition
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    
    With Selection.FormatConditions(1)
        .ReverseOrder = True
        .ShowIconOnly = False
        .IconSet = ActiveWorkbook.IconSets(xl3TrafficLights1)
    End With
    
    With Selection.FormatConditions(1).IconCriteria(2)
        .Type = xlConditionValueFormula
        .Value = "=" & TheOtherCell & "+2"
        .Operator = 5
    End With
    
    With Selection.FormatConditions(1).IconCriteria(3)
        .Type = xlConditionValueFormula
        .Value = "=" & TheOtherCell & "+4"
        .Operator = 5
    End With
End Sub
