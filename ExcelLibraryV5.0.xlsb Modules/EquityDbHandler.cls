VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EquityDbHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Purpose   :    Query DB and return mnemonics for factors pre-analyzed using mimicking portfolios
'Inputs    :    studyType
'Outputs   :    Returns a 1D string array of mnemonics
'Author    :    Jesus Carrero
'Date      :    25 March 2014
'Notes     :
Public Function GetPallFactorsStudied(studyType As QuantileType) As Variant
    If studyType = Quintile Then
        Let GetPallFactorsStudied = ConnectAndSelect("SELECT distinct(substring(type, 1, length(type)-18)) AS MNEMONIC " _
                    & " FROM modeling.`mimicking-factor-quintile-returns` WHERE type like '%PALL_%';", _
                    "modeling", DbServerAddress, DbUserName, DbPassword)
    ElseIf studyType = Decile Then
        Let GetPallFactorsStudied = ConnectAndSelect("SELECT distinct(substring(type, 1, length(type)-18)) AS MNEMONIC " _
                    & " FROM modeling.`mimicking-factor-decile-returns` WHERE type like '%PALL_%';", _
                    "modeling", DbServerAddress, DbUserName, DbPassword)
    End If
End Function

'Purpose   :    Query returns with the requested rebalancing frequency for pre-computed Mimicking returns.
'               Compute cummulative returns, drawdown, and compute the data needed to plot return histogram
'Inputs    :    factor          factor mnemonic [e.g PB, PS, ATO, ...]
'               region          [REG1, REG2, ...]
'               weight          [EQW, VLW, and MKW]
'               startDate       in serial format
'               endDate         in serial format
'               rebalPeriod     [28,...] in days
'               studyType
'               revert          [Boolean, revert the sort]
'Outputs   :    Return 1D array with three 2D arrays (three returned values):
'               1. 2D array with headers row and the following data organized in columns
'                  a. date in serial format
'                  b. the factor's mnemonic
'                  c. weighted return (per user's param above) for the low quantile
'                  d. weighted return (per user's param above) for the high quantile
'                  e. spread of high - low weighted return (per user's param above) (of c and d above)
'               2. 2D array without headers row and the following data organized in colunns
'                  a. date in serial format
'                  b. cummulative return of high quantile
'                  c. cummulative return of low quantile
'                  d. cummulative return of spread
'                  e. loss (represented as a negative % of original investment) or 0 if capital above initial investment
'                  f. column of auxiliary compututations (don't use) ('***HERE TODO remove this columnn before return matrix)
'Author    :    Jesus Carrero
'Date      :    25 March 2014
'
'Notes
Public Function GetMimickingReturnsAndStats(factor As String, Region As String, Weight As String, _
                             StartDate As Long, EndDate As Long, rebalPeriod As Long, _
                             studyType As QuantileType, revert As Boolean) As Variant

    ' Clear Range
    Call TempComputation.UsedRange.ClearContents

    ' Uses Store procedure defined in the Database.
    Dim factorRetunrs As Variant
    Let factorRetunrs = GetTimeSeriesOfMimickingPortfolioReturns( _
                        factor, Region, Weight, str(StartDate), str(EndDate), studyType, revert)

    If (GetNumberOfRows(factorRetunrs) > 1) Then

        ' Write Data to Spread Sheet
        Call TempComputation.UsedRange.ClearContents ' Clear Data
        Call DumpInTempPositionWithoutFirstClearing(factorRetunrs, TempComputation.Range("E10"))
        Call TempComputation.UsedRange.EntireColumn.AutoFit
          
        Dim LastRow As Long
        LastRow = TempComputation.Range("E11").End(xlDown).row
    
        ' Compute Returns
        TempComputation.Range("T9").Value2 = "Date"
        TempComputation.Range("U9:W9").Value2 = TempComputation.Range("G10:I10").Value2
        TempComputation.Range("U10:Y10").Value = 1
        TempComputation.Range("X10").Value = 0
        TempComputation.Range("U11:W" & Trim(str(LastRow))).FormulaR1C1 = "=(1+RC[-14]/100)*R[-1]C"
    
        ' CopyDates
        TempComputation.Range("T11:T" & Trim(str(LastRow))).Value = _
        TempComputation.Range("E11:E" & Trim(str(LastRow))).Value
    
        ' Compute drawdown
        TempComputation.Range("Y11:Y" & Trim(str(LastRow))).FormulaR1C1 = "=MIN((1+RC[-16]/100)*R[-1]C, 1)"
        TempComputation.Range("X11:X" & Trim(str(LastRow))).FormulaR1C1 = "=RC[1] - 1"
 
        ' Generate Histogram
        Dim histogram As Variant
        Let histogram = GenerateHistogram(TempComputation.Range("$I$11:I" & Trim(str(LastRow))), TempComputation, 10)
 
        Let GetMimickingReturnsAndStats = Array(TempComputation.Range("E10").CurrentRegion.Value2, _
                            TempComputation.Range("T9").CurrentRegion.Value2, _
                            histogram)
    Else
        Let GetMimickingReturnsAndStats = Array(factorRetunrs, factorRetunrs, factorRetunrs)
    End If
    
End Function

'Purpose   :    Query from DB the avg (longitudinally) return spread between two dates
'Inputs    :    factor          factor mnemonic [e.g PB, PS, ATO, ...]
'               region          [REG1, REG2, ...]
'               weight          [EQW, VLW, and MKW]
'               startDate       in serial format
'               endDate         in serial format
'Outputs   :    2D array with headers row and the following data organized in columns:
'                  a. date in serial format
'                  b. the factor's mnemonic
'                  c. longitudinal average return spread for the period between startDate and endDate
'Author    :    Jesus Carrero
'Date      :    25 March 2014
'Notes     :
Public Function GetAvgReturnFromMimickingQuintile(factor As String, Region As String, Weight As String, _
                     StartDate As String, EndDate As String) As Variant
                     
    Let GetPallFactorsStudied = ConnectAndSelect("CALL modeling.mimickingAvgReturnSectorRegionFromQuintileTable " _
            & "'PALL_" & factor & "','" & Region & "_" & Weight & "'," & StartDate & "," & EndDate & ")", _
                    "modeling", DbServerAddress, DbUserName, DbPassword)
End Function

'Purpose   :    Query the factor weights and mnemonics for the given model name (e.g modelDescriptor) at the given serial date
'Inputs    :    modelDescriptor
'               modelDate
'Outputs   :    2D array in Variant format with headers for the following columns:
'               a. date
'               b. factor mnemonic
'               c. weight
'               d. source table name (table holding factor's data)
'Author    :    Jesus Carrero
'Date      :    25 March 2014
'Notes     :
Public Function GetLatestModel(modelDescriptor As String, modelDate As Long) As Variant

    Let GetLatestModel = ConnectAndSelect("CALL modeling.`model-get-most-recent`('" _
                    & modelDescriptor & "'," & str(modelDate) & ")", _
                    "modeling", DbServerAddress, DbUserName, DbPassword)
End Function

'Purpose   :    Query avg returns from database.
'Inputs    :    factor          factor mnemonic [e.g PB, PS, ATO, ...]
'               region          [REG1, REG2, ...]
'               weight          [EQW, VLW, MKW]
'               startDate
'               endDate
'               rebalPeriod     [28,...]
'               studyType
'               revert          [Boolean, revert the sort]
'Outputs   :    Table in Variant Format
'Author    :    Jesus Carrero
'Date      :    25 March 2014
'Notes
Public Function GetTimeSeriesOfMimickingPortfolioReturns _
        (factor As String, Region As String, Weight As String, _
        StartDate As String, EndDate As String, studyType As QuantileType, revert As Boolean) As Variant
        
    Dim tmpTable As Variant
    If studyType = Quintile Then
        Let tmpTable = ConnectAndSelect("CALL `modeling`.`mimicking-ts-spread-quintile-returns`(" _
                    & "'" & factor & "','" & Region & "','" & Weight & "'," & StartDate & "," & EndDate & ")", _
                    "modeling", DbServerAddress, DbUserName, DbPassword)
    ElseIf studyType = Decile Then
        Let tmpTable = ConnectAndSelect("CALL `modeling`.`mimicking-ts-spread-decile-returns`(" _
                    & "'" & factor & "','" & Region & "','" & Weight & "'," & StartDate & "," & EndDate & ")", _
                    "modeling", DbServerAddress, DbUserName, DbPassword)
    End If
    
    If revert Then
        'Dump in TempComputation, flip LOW, HIGH and change return sign
        Call TempComputation.UsedRange.ClearContents ' Clear Data
        Call DumpInTempPositionWithoutFirstClearing(tmpTable, TempComputation.Range("A1"))
        
        Dim LastRow As Long
        LastRow = TempComputation.Range("A1").End(xlDown).row
        ' CopyLows to Temp Position.
        TempComputation.Range("H2:H" & Trim(str(LastRow))).Value = _
        TempComputation.Range("C2:C" & Trim(str(LastRow))).Value
    
        ' Copy Highs To Low
        TempComputation.Range("C2:C" & Trim(str(LastRow))).Value = _
        TempComputation.Range("D2:D" & Trim(str(LastRow))).Value
    
        ' Copy Lows To High Position
        TempComputation.Range("D2:D" & Trim(str(LastRow))).Value = _
        TempComputation.Range("H2:H" & Trim(str(LastRow))).Value
    
        ' Multiply return column by -1 and copy back to D postion
        TempComputation.Range("H2:H" & Trim(str(LastRow))).FormulaR1C1 = "= -1*C[-3]"
        TempComputation.Range("E2:E" & Trim(str(LastRow))).Value = _
        TempComputation.Range("H2:H" & Trim(str(LastRow))).Value
        
        ' Return resutls
        Let GetTimeSeriesOfMimickingPortfolioReturns = TempComputation.Range("A1").CurrentRegion.Value2
    Else
        Let GetTimeSeriesOfMimickingPortfolioReturns = tmpTable
    End If
    
End Function

'Purpose   :    Compute Histogram of data
'Inputs    :    dataRange              [data to bin]
'               targetWsht             [spreadsheet to make computation]
'               numBins                [number of bins]
'Outputs   :    Table in Variant Format
'Author    :    Jesus Carrero
'Date      :    25 March 2014
'Notes
Private Function GenerateHistogram(dataRange As Range, targetWsht As Worksheet, numBins As Long) As Variant
    ' Compute Histogram Parameters
    Dim MinAbsReturn As Double, MaxAbsReturn As Double, Increment As Double
    MinAbsReturn = Application.Min(dataRange.Value2)
    MaxAbsReturn = Application.WorksheetFunction.Max(dataRange.Value2)
    Increment = (MaxAbsReturn - MinAbsReturn) / numBins
    
    ' Generate Bins Boundaries
    Let targetWsht.Range("$K$11").Value2 = Round(MinAbsReturn, 2)
    Let targetWsht.Range("$K$12").Resize(numBins - 1, 1).FormulaR1C1 = "= ROUND(R[-1]C[0] + " & Increment & ",2)"
    
    'Count frequency
    Let targetWsht.Range("$L$11:L" & Trim(str(11 + numBins - 1))).FormulaArray = _
        "= frequency(" & dataRange.Address & ", " & targetWsht.Range("$K$11").Resize(numBins, 1).Address & ")"
  
    Let GenerateHistogram = targetWsht.Range("$K$11").CurrentRegion.Value2
End Function
