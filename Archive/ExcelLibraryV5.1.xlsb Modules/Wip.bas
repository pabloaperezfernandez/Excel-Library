Attribute VB_Name = "Wip"
Option Explicit
Option Base 1

' This function returns the number of the latest run for the given asset class and strategy.
' The function returns 0 if there is no data in the run table
Public Function GetLatestWipRun(TheAssetClass As String, TheStrategy As String) As Integer
    Dim TheQuery As String
    Dim TheResults As Variant
    Dim TheLatestRunNumber As Integer

    ' Get the latest run's number
    Let TheQuery = "SELECT MAX(`runnumber`) FROM `etwip2dot0`.`tradingprocessledger` WHERE `assetclasscode` = '" & TheAssetClass & "' AND `strategycode` = '" & TheStrategy & "';"
    Let TheResults = ConnectAndSelect(TheQuery, "etwip2dot0", DbServerAddress, DbUserName, DbPassword)
    
    If GetNumberOfRows(TheResults) < 2 Then
        Let GetLatestWipRun = 0
        Exit Function
    End If

    Let GetLatestWipRun = TheResults(2, 1)
End Function

' This function returns TRUE if there is an ongoing run for the given asset class and strategy.
' Otherwise if returns FALSE
Public Function LatestWipRunOpenedQ(TheAssetClass As String, TheStrategy As String) As Boolean
    Dim TheQuery As String
    Dim LatestRunNumber As Integer
    Dim TheResults As Variant
    Dim TheDate As Long
    Dim TheLatestRunNumber As Integer

    ' Get the latest run's number
    Let TheQuery = "SELECT MAX(`runnumber`) FROM `etwip2dot0`.`tradingprocessledger` WHERE `assetclasscode` = '" & TheAssetClass & "' AND `strategycode` = '" & TheStrategy & "';"
    Let TheResults = ConnectAndSelect(TheQuery, "etwip2dot0", DbServerAddress, DbUserName, DbPassword)
    
    If GetNumberOfRows(TheResults) < 2 Then
        Let LatestWipRunOpenedQ = False
        Exit Function
    End If
    
    ' Exit if there is no authorized run
    If IsEmpty(TheResults(2, 1)) Or IsNull(TheResults(2, 1)) Then Exit Function
    
    ' Get the latest run's start and end dates
    Let TheQuery = "SELECT `authorizationdate`, `completiondate` FROM `etwip2dot0`.`tradingprocessledger` WHERE `assetclasscode` = '" & TheAssetClass & "' AND `strategycode` = '" & TheStrategy & "' AND `runnumber` = " & TheResults(2, 1) & ";"
    Let TheResults = ConnectAndSelect(TheQuery, "etwip2dot0", DbServerAddress, DbUserName, DbPassword)
    If IsEmpty(TheResults(2, 2)) Or IsNull(TheResults(2, 2)) Then
        Let LatestWipRunOpenedQ = True
    Else
        Let LatestWipRunOpenedQ = False
    End If
End Function

