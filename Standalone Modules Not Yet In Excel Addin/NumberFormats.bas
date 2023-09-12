Attribute VB_Name = "NumberFormats"
' PURPOSE OF MODULE
' The purpose of this module is to centralized

Option Explicit
Option Base 1

Public Function EuroNumberFormat(Optional NumberOfDecimalPlaces As Integer = 0) As String
    If NumberOfDecimalPlaces = 0 Then
        Let EuroNumberFormat = "_-[$—-x-euro2] * #,##0_-;-[$—-x-euro2] * #,##0_-;" & _
                               "_-[$—-x-euro2] * ""-""??_-;_-@_-"
    Else
        Let EuroNumberFormat = "_-[$—-x-euro2] * #,##0." & String$("0", NumberOfDecimalPlaces) & "_-;" & _
                               "-[$—-x-euro2] * #,##0." & String$("0", NumberOfDecimalPlaces) & "_-;" & _
                               "_-[$—-x-euro2] * ""-""??_-;_-@_-"
    End If
End Function

Public Function DecimalNumberFormat(Optional NumberOfDecimalPlaces As Integer = 0) As String
    If NumberOfDecimalPlaces = 0 Then
        Let DecimalNumberFormat = "#,##0"
    Else
        Let DecimalNumberFormat = "#,##0." & String$("0", NumberOfDecimalPlaces)
    End If
End Function


