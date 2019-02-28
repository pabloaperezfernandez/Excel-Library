Attribute VB_Name = "DateAndTime"
Option Base 1
Option Explicit

' Converts a serial date to date data type
Public Function ConvertToDateFromSerial(aSerialDate As Long) As Date
    Let ConvertToDateFromSerial = CDate(GetMonthFromSerialDate(aSerialDate) & "/" & _
                                        GetDayFromSerialDate(aSerialDate) & "/" & _
                                        GetYearFromSerialDate(aSerialDate))
End Function

' This function converts an Excel date into a serial date (e.g. YYYYMMDD)
' A Blank is Interpreted in Excel as 0 so we check to this
' Date after extracting using Year/Month/Day function
Public Function ConvertDateToSerial(aDate As Date) As Long
    Dim TheYear As String
    Dim TheMonth As String
    Dim TheDay As String
    
    If aDate = CDate(0) Then
        Let ConvertDateToSerial = 0
        Exit Function
    End If
    
    Let TheYear = CStr(Year(aDate))
    Let TheMonth = CStr(Month(aDate))
    Let TheDay = CStr(Day(aDate))
    
    If Len(TheMonth) < 2 Then
        Let TheMonth = "0" & TheMonth
    End If

    If Len(TheDay) < 2 Then
        Let TheDay = "0" & TheDay
    End If

    Let ConvertDateToSerial = CLng(TheYear & TheMonth & TheDay)
End Function

' Extract parts of a serial date
Public Function GetYearFromSerialDate(aDate As Long) As Long
    Let GetYearFromSerialDate = CLng(Left(aDate, 4))
End Function

Public Function GetMonthFromSerialDate(aDate As Long) As Long
    Let GetMonthFromSerialDate = CLng(Mid(aDate, 5, 2))
End Function

Public Function GetDayFromSerialDate(aDate As Long) As Long
    Let GetDayFromSerialDate = CLng(Right(aDate, 2))
End Function

' This function converts an Excel time (which is of type date) into a serial time (e.g. HHMMSS)
' The output may have fewer than six digits.  This is because the output may have leading zeroes.
' However, this is not a problem since smaller numbers happen earlier in the day.  So, this format
' preserves ascending chronological order. The returned serial represents to be in 24-hour format.
Public Function ConvertTimeToSerial(aTime As Date) As Long
    Dim TheHours As String
    Dim TheMinutes As String
    Dim TheSeconds As String
    
    Let TheHours = CStr(Hour(aTime))
    Let TheMinutes = CStr(Minute(aTime))
    Let TheSeconds = CStr(Second(aTime))
    
    If Len(TheHours) < 2 Then
        Let TheHours = "0" & TheHours
    End If

    If Len(TheMinutes) < 2 Then
        Let TheMinutes = "0" & TheMinutes
    End If

    If Len(TheSeconds) < 2 Then
        Let TheSeconds = "0" & TheSeconds
    End If
    
    Let ConvertTimeToSerial = CLng(TheHours & TheMinutes & TheSeconds)
End Function

' Extract parts of a serial time
Public Function GetHourFromSerialTime(aTime As Long) As Long
    Dim aTimeStr As String
    
    Let aTimeStr = CStr(aTime)
    
    If Len(aTimeStr) < 6 Then
        Let aTimeStr = String(6 - Len(aTimeStr), "0") & aTime
    End If
        
    Let GetHourFromSerialTime = CLng(Left(aTimeStr, 2))
End Function

Public Function GetMinuteFromSerialTime(aTime As Long) As Long
    Dim aTimeStr As String
    
    Let aTimeStr = CStr(aTime)
    
    If Len(aTimeStr) < 6 Then
        Let aTimeStr = String(6 - Len(aTimeStr), "0") & aTime
    End If
        
    Let GetMinuteFromSerialTime = CLng(Mid(aTimeStr, 3, 2))
End Function

Public Function GetSecondFromSerialTime(aTime As Long) As Long
    Dim aTimeStr As String
    
    Let aTimeStr = CStr(aTime)
    
    If Len(aTimeStr) < 6 Then
        Let aTimeStr = String(6 - Len(aTimeStr), "0") & aTime
    End If
        
    Let GetSecondFromSerialTime = CLng(Right(aTimeStr, 2))
End Function

Public Function ConvertDateTimeToMySQLFormat(aDate As Date) As String
    Let ConvertDateTimeToMySQLFormat = Format(aDate, "yyyy-mm-dd hh:mm:ss")
End Function
