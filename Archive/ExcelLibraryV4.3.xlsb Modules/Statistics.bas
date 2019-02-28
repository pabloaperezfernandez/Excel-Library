Attribute VB_Name = "Statistics"
Option Base 1
Option Explicit

' Notes:
' 1. Some of these functions require the workbook to have a worksheet called "TempComputations."
'    This is necessary because the on-demand creation and deletion of a temp worksheet is time
'    consuming.

' Computes and returns the correlation matrix of the return vector set packed into a return
' matrix.
Public Function VarianceCovarianceMatrix(MatrixOfColumnReturnVectors)
    Dim i As Long, j As Long, k As Long
    Dim nc As Integer, nr As Integer
    Dim RowVector As Variant, ColumnVector As Variant, CorrMat As Variant
    Dim Results As Variant
    
    ' Determine the number of rows and columns
    Let nr = UBound(MatrixOfColumnReturnVectors, 1)
    Let nc = UBound(MatrixOfColumnReturnVectors, 2)
    
    ' Pre-allocate
    ReDim Results(1 To nc, 1 To nc)
    
    ' Extract the Nr row vectors
    For i = 1 To nc
        For j = 1 To nc
            ' Take the correlation coefficient of column i and column j
            Let Results(i, j) = Application.Covar(Application.Index(MatrixOfColumnReturnVectors, 0, i), _
                                                  Application.Index(MatrixOfColumnReturnVectors, 0, j))
        Next j
    Next i

    Let VarianceCovarianceMatrix = Results
End Function

' Computes and returns the correlation matrix of the return vector set packed into a return
' matrix.
Public Function CorrelationMatrix(MatrixOfColumnReturnVectors)
    Dim i As Long, j As Long, k As Long
    Dim nc As Integer, nr As Integer
    Dim RowVector As Variant, ColumnVector As Variant, CorrMat As Variant
    Dim Results As Variant
    
    ' Determine the number of rows and columns
    Let nr = UBound(MatrixOfColumnReturnVectors, 1)
    Let nc = UBound(MatrixOfColumnReturnVectors, 2)
    
    ' Pre-allocate
    ReDim Results(1 To nc, 1 To nc)
    
    ' Extract the Nr row vectors
    For i = 1 To nc
        For j = 1 To nc
            ' Take the correlation coefficient of column i and column j
            Let Results(i, j) = Application.Correl(Application.Index(MatrixOfColumnReturnVectors, 0, i), _
                                                   Application.Index(MatrixOfColumnReturnVectors, 0, j))
        Next j
    Next i

    Let CorrelationMatrix = Results
End Function

' Returns the size in elements of each quantile given then number of elements and the
' number of quantiles.  It distributes the excess evenly split over the first
' group of quantiles.  This is not the standard way of doing this, but it works fine for
' our purposes.
Public Function ComputeQuantileSizes(NumberOfElements As Long, NumberOfQuantiles As Integer)
    Dim BaseQuantileSize As Long
    Dim QuantileSizes() As Long
    Dim i As Integer
    
    ' Pre-allocate array to return quantile sizes
    ReDim QuantileSizes(1 To NumberOfQuantiles)
    
    ' Compute base quantile size (in number of elements). baseQuantileSize should theoretically be an integer, but
    ' we want to ensure no errors that may arise from the numeric, machine representation.
    Let BaseQuantileSize = CInt((NumberOfElements - NumberOfElements Mod NumberOfQuantiles) / NumberOfQuantiles)
    
    For i = 1 To NumberOfQuantiles
        If i <= NumberOfElements - NumberOfQuantiles * BaseQuantileSize Then
            Let QuantileSizes(i) = BaseQuantileSize + 1
        Else
            Let QuantileSizes(i) = BaseQuantileSize
        End If
    Next i
    
    ' Set return value
    Let ComputeQuantileSizes = QuantileSizes
End Function

' This function returns a vector with the ranks of each number in the array DataToQuantile. DataToQuantile is
' expected to be either a 1D array or a 1-column matrix.  DataToQuantile is expected to be sorted in ascending
' order. This function needs to sort DataToQuantile and then return them to the original order after
' computing the quantile ranks. The easiest way to do this is to record the original row numbers and use that to
' return them to their original positions after computing quantile ranks.
Public Function ComputeQuantiles(DataToQuantile As Variant, NumberOfQuantiles As Integer)
    Dim sht As Worksheet
    Dim TheRange As Range
    Dim OriginalPositions As Variant
    Dim QuantileSizes As Variant
    Dim QuantileRanks() As Integer
    Dim QuantileIndex As Integer
    Dim i As Long
    Dim j As Long
    
    ' Pre-allocate array to hold quantile ranks
    ReDim QuantileRanks(1 To UBound(DataToQuantile))
    
    ' Pre-allocate array to hold original positions
    ReDim OriginalPositions(1 To UBound(DataToQuantile))
    
    ' Set reference to temp computation worksheet
    Set sht = Worksheets("TempComputation")
    
    ' Clear any previous contents
    Call sht.UsedRange.ClearContents
    
    ' Dump data into temp computation worksheet
    Let sht.Range("A1").Resize(UBound(DataToQuantile), 1).Value2 = Application.Transpose(DataToQuantile)
    
    ' Compute original positions for original data so may sort the data back into the original order
    Let sht.Range("B1").Resize(UBound(DataToQuantile), 1).FormulaR1C1 = "=Row(R[0]C[-1])"
    Let sht.Range("B1").Resize(UBound(DataToQuantile), 1).Value2 = sht.Range("B1").Resize(UBound(DataToQuantile), 1).Value2
    Let OriginalPositions = ConvertTo1DArray(sht.Range("B1").Resize(UBound(DataToQuantile), 1).Value2)
    
    ' Dump data to be sorted into temp computation worksheet and sort in ascending order
    sht.UsedRange.ClearContents
    Let sht.Range("A1").Resize(UBound(DataToQuantile), 1).Value2 = Application.Transpose(DataToQuantile)
    Let sht.Range("B1").Resize(UBound(DataToQuantile), 1).Value2 = Application.Transpose(OriginalPositions)
    
    ' Sort DataToQuantile in ascending order
    ' Set range pointer to the data we just dumped in worksheet TempComputation
    Set TheRange = sht.Range("A1").CurrentRegion

    ' Clear any previous sorting criteria
    sht.Sort.SortFields.Clear

    ' Add criteria to sort by date
    sht.Sort.SortFields.Add _
        Key:=TheRange.Columns(1), _
        SortOn:=xlSortOnValues, _
        Order:=xlAscending, _
        DataOption:=xlSortNormal
        
    ' Execute the sort
    With sht.Sort
        .SetRange TheRange
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    ' Compute the quantile sizes.
    Let QuantileSizes = ComputeQuantileSizes(UBound(DataToQuantile), NumberOfQuantiles)
    
    ' Generate array with a quantile rank for each element in dataToQuantile
    Let QuantileIndex = 1
    ' Initialize quantile counter (how many elts have been ranked in the quintile QuantileIndex.
    Let j = 0
    ' Cycle through every item in the data to be quantile, computing a quantile rank for each.
    For i = 1 To UBound(DataToQuantile)
        ' Check if we have assigned the current QuantileIndex to the correct number of dataset elements
        If j < QuantileSizes(QuantileIndex) Then
            ' Increase the counter to record that we did one quantile ranking assignment
            Let j = j + 1
        Else
            ' We need to move to the next quantile
            Let QuantileIndex = QuantileIndex + 1
            ' Record one quantile ranking assigned since this will happen right after exiting this If statement.
            Let j = 1
        End If
        
        ' Assign the current QuantileIndex to the current dataset element.
        Let QuantileRanks(i) = QuantileIndex
    Next i
    
    ' Dump the quantile ranks into the third column of worksheet TempComputation
    Let sht.Range("C1").Resize(UBound(DataToQuantile), 1).Value2 = Application.Transpose(QuantileRanks)
    
    ' Sort data according to original position
    Set TheRange = sht.Range("A1").CurrentRegion

    ' Clear any previous sorting criteria
    sht.Sort.SortFields.Clear

    ' Sort data back into the original order
    sht.Sort.SortFields.Add _
        Key:=TheRange.Columns(2), _
        SortOn:=xlSortOnValues, _
        Order:=xlAscending, _
        DataOption:=xlSortNormal
        
    With sht.Sort
        .SetRange TheRange
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    ' Return the array of quantile ranks in original order of DataToQuantile
    Let ComputeQuantiles = ConvertTo1DArray(TheRange.Columns(3).Value2)
End Function

' Given a factor's vector, its corresponding return vector, and the number of quantiles, this function
' returns the average return for each of the factor's quantile ranks.  In other words, say the factor is P/E and
' the other factor is one-week, forward returns.  Say we are interesting in the average, one-week, forward
' performance of each quintile.  The number of quintiles would be 5, the first factor's vector is the list of P/Es
' (one per security) and the second factor's vector is the list of one-week forward returns for each security.
' This function would return a five-element array with the average, one-week forward return of each security.
Public Function ComputeAveragePerformanceByQuantile(FactorVector As Variant, ReturnVector As Variant, _
                                                    NumberOfQuantiles As Integer)
    Dim TempSheet As Worksheet
    Dim FactorRange As Range
    Dim ReturnRange As Range
    Dim QuantileRanks As Variant
    Dim QuantileRange As Range
    Dim Results As Variant
    Dim i As Integer
    
    ' Pre-allocate array for average value of ReturnVector for each quantile index
    ReDim Results(1 To NumberOfQuantiles)
    
    ' Set reference to temp sheet and clear it
    Set TempSheet = Worksheets("TempComputation")
    TempSheet.UsedRange.ClearContents
    
    ' Quantile QuantileVector
    Let QuantileRanks = ComputeQuantiles(FactorVector, NumberOfQuantiles)
    
    ' Dump FactorVector in TempSheet
    Set FactorRange = TempSheet.Range("A1").Resize(UBound(FactorVector), 1)
    Let FactorRange.Value2 = Application.Transpose(FactorVector)
    
    ' Dump QuantileRanks in TempSheet
    Set QuantileRange = TempSheet.Range("B1").Resize(UBound(QuantileRanks), 1)
    Let QuantileRange.Value2 = Application.Transpose(QuantileRanks)
    
    ' Dump ReturnVector in Tempsheeet
    Set ReturnRange = TempSheet.Range("C1").Resize(UBound(ReturnVector), 1)
    Let ReturnRange.Value2 = Application.Transpose(ReturnVector)

    ' Compute the average value of CorrespondingVector for each quantile index.
    For i = 1 To NumberOfQuantiles
        Let Results(i) = Application.AverageIf(QuantileRange, "=" & i, ReturnRange)
    Next i
    
    ' Return results
    Let ComputeAveragePerformanceByQuantile = Results
End Function
                                           
' Given a factor's vector, its corresponding return vector, a weight vector, and the number of quantiles, this
' function returns the weighted, average return for each of the factor's quantile ranks.  In other words, say
' the factor is P/E and the other factor is one-week, forward returns.  Say we are interesting in the average,
' one-week, forward performance of each quintile.  The number of quintiles would be 5, the first factor's vector
' is the list of P/Es (one per security) and the second factor's vector is the list of one-week forward returns
' for each security. This function would return a five-element array with the weighted, average, one-week forward
' return of each security.
Public Function ComputeWeightedAveragePerformanceByQuantile(FactorVector As Variant, ReturnVector As Variant, _
                                                            WeightVector As Variant, NumberOfQuantiles As Integer)
    Dim TempSheet As Worksheet
    Dim FactorRange As Range
    Dim ReturnRange As Range
    Dim WeightRange As Range
    Dim QuantileRanks As Variant
    Dim QuantileRange As Range
    Dim WeightedRange As Range
    Dim Results As Variant
    Dim i As Integer
    
    ' Pre-allocate array for average value of ReturnVector for each quantile index
    ReDim Results(1 To NumberOfQuantiles)
    
    ' Set reference to temp sheet and clear it
    Set TempSheet = Worksheets("TempComputation")
    TempSheet.UsedRange.ClearContents
    
    ' Quantile QuantileVector
    Let QuantileRanks = ComputeQuantiles(FactorVector, NumberOfQuantiles)
    
    ' Dump FactorVector in TempSheet
    Set FactorRange = TempSheet.Range("A1").Resize(UBound(FactorVector), 1)
    Let FactorRange.Value2 = Application.Transpose(FactorVector)
    
    ' Dump QuantileRanks in TempSheet
    Set QuantileRange = TempSheet.Range("B1").Resize(UBound(QuantileRanks), 1)
    Let QuantileRange.Value2 = Application.Transpose(QuantileRanks)
    
    ' Dump ReturnVector in Tempsheeet
    Set ReturnRange = TempSheet.Range("C1").Resize(UBound(ReturnVector), 1)
    Let ReturnRange.Value2 = Application.Transpose(ReturnVector)
    
    ' Dump WeightVector in Tempsheeet
    Set WeightRange = TempSheet.Range("D1").Resize(UBound(WeightVector), 1)
    Let WeightRange.Value2 = Application.Transpose(WeightVector)
    
    ' Weight the factor values and dump them into the spreadsheet
    Set WeightedRange = TempSheet.Range("E1").Resize(UBound(WeightVector), 1)
    Let WeightedRange.FormulaR1C1 = "=RC[-2]*RC[-1]"

    ' Compute the average value of CorrespondingVector for each quantile index.
    For i = 1 To NumberOfQuantiles
        Let Results(i) = Application.SumIf(QuantileRange, "=" & i, WeightedRange)
        Let Results(i) = Results(i) / Application.WorksheetFunction.SumIf(QuantileRange, "=" & i, WeightRange)
    Next i
    
    ' Return results
    Let ComputeWeightedAveragePerformanceByQuantile = Results
End Function

' This function returns a matrix representing a time series of column vectors, each of which is the average
' return of the quantiles on a given week.  This function takes a matrix of factor values (e.g. the columns are
' time-slices of the some factor, which means that each week's worth of the factor's vector is represented by a
' different column.  The same holds by the ReturnMatrix.  This function returns a matrix with the number of rows
' equal to the number of quantiles and the each column holds the average return of each quantile on a given date.
Function ComputeQuantileAveragePerformanceTimeSeries(FactorMatrix As Variant, ReturnMatrix As Variant, _
                                                     NumberOfQuantiles As Integer)
    Dim i As Integer
    Dim j As Integer
    Dim Results As Variant
    Dim ResultsOneWeek As Variant
    
    ' Pre-allocate the Results array
    ReDim Results(1 To NumberOfQuantiles, 1 To UBound(FactorMatrix, 2))
    
    For i = 1 To UBound(FactorMatrix, 2)
        Let ResultsOneWeek = ComputeAveragePerformanceByQuantile(ConvertTo1DArray(Application.Index(FactorMatrix, 0, i)), _
                                                                 ConvertTo1DArray(Application.Index(ReturnMatrix, 0, i)), _
                                                                 NumberOfQuantiles)
    
        For j = 1 To NumberOfQuantiles
            Let Results(j, i) = ResultsOneWeek(j)
        Next j
    Next i
    
    Let ComputeQuantileAveragePerformanceTimeSeries = Results
End Function

' This function returns a matrix representing a time series of column vectors, each of which is the average
' return of the quantiles on a given week.  This function takes a matrix of factor values (e.g. the columns are
' time-slices of the some factor, which means that each week's worth of the factor's vector is represented by a
' different column.  The same holds by the ReturnMatrix.  This function returns a matrix with the number of rows
' equal to the number of quantiles and the each column holds the average return of each quantile on a given date.
Function ComputeQuantileAveragePerformanceTimeSeries2(FactorMatrix As Variant, ReturnMatrix As Variant, _
                                                      NumberOfQuantiles As Integer)
    Dim i As Integer
    Dim Results() As Variant
    
    ' Pre-allocate the Results array
    ReDim Results(1 To UBound(FactorMatrix, 2))
    
    ' Loop through the weeks.
    For i = 1 To UBound(FactorMatrix, 2)
        ' Compute the average performance of each quantile on this week.
        Let Results(i) = ComputeAveragePerformanceByQuantile(ConvertTo1DArray(Application.Index(FactorMatrix, 0, i)), _
                                                             ConvertTo1DArray(Application.Index(ReturnMatrix, 0, i)), _
                                                             NumberOfQuantiles)
                                                             
        ' Turn this week's average quantile performance into a 1D array
        Let Results(i) = ConvertTo1DArray(Results(i))
    Next i
    
    ' Set the results matrix as the return value of the function
    Let ComputeQuantileAveragePerformanceTimeSeries2 = Application.Transpose(Results)
End Function

' This function returns the distribution for the given return times series.  This
' divides the range of values evenly into NumberOfBins bins.  The function returns an
' Nx2 matrix.  The left column are the right-end points of the bin intervals.  The
' right column are the counts.
Function ComputeDistribution(TimeSeries As Variant, NumberOfBins As Integer)
    Dim TmpSheet As Worksheet
    Dim MaxVal As Double
    Dim MinVal As Double
    Dim BinSize As Double
    Dim BinLeftEndPoints As Variant
    Dim i As Integer
    
    ' Set reference to worksheet TempComputation
    Set TmpSheet = Worksheets("TempComputation")
    
    ' Pre-allocate array to hold the left endpoints of each bin
    ReDim BinLeftEndPoints(1 To NumberOfBins)
    
    ' Determine the max and min values in the time series
    Let MaxVal = Application.Max(TimeSeries)
    Let MinVal = Application.Min(TimeSeries)
    
    ' Determine bin size
    Let BinSize = (MaxVal - MinVal) / (NumberOfBins - 1)
    
    ' Create sequence of bins
    For i = 0 To NumberOfBins - 1
        Let BinLeftEndPoints(i + 1) = MinVal + BinSize * i
    Next i
    
    ' Dump the dataset into TempComputation worksheet
    Let TmpSheet.Range("A1").Resize(UBound(TimeSeries, 1), 1).Value2 = Application.Transpose(TimeSeries)
    
    ' Dump bins', left-end points into TempComputation worksheet
    Let TmpSheet.Range("B1").Resize(NumberOfBins, 1).Value2 = Application.Transpose(BinLeftEndPoints)
    
    ' Dump formulas to compute frequencies. Each bin contains the count that is less than or equal to the
    ' right-hand end point of the interval.
    Let TmpSheet.Range("C2").Resize(NumberOfBins - 1, 1).FormulaR1C1 = _
        "=countif(R1C1:R" & UBound(TimeSeries, 1) & "C1, " & Chr(34) & "<=" & Chr(34) & "&R[0]C[-1])-sum(R1C3:R[-1]C[0])"
        
    ' Set value to return
    Let ComputeDistribution = TmpSheet.Range("B2").Resize(NumberOfBins - 1, 2).Value2
End Function

' This function returns the distribution for the given return times series using the
' Frequency() worksheet funciton.  This divides the range of values evenly into NumberOfBins
' bins.  The function returns an Nx2 matrix.  The left column are the right-end points of the
' bin intervals.  The right column are the counts.
Function ComputeDistribution2(TimeSeries As Variant, NumberOfBins As Integer)
    Dim TmpSheet As Worksheet
    Dim MaxVal As Double
    Dim MinVal As Double
    Dim BinSize As Double
    Dim BinRightEndPoints As Variant
    Dim i As Integer
    
    ' Set reference to worksheet TempComputation
    Set TmpSheet = Worksheets("TempComputation")
    
    ' Pre-allocate array to hold the left endpoints of each bin
    ReDim BinRightEndPoints(1 To NumberOfBins)
    
    ' Determine the max and min values in the time series
    Let MaxVal = Application.Max(TimeSeries)
    Let MinVal = Application.Min(TimeSeries)
    
    ' Determine bin size
    Let BinSize = (MaxVal - MinVal) / NumberOfBins
    
    ' Create the sequence of right, end-points for the bins
    For i = 1 To NumberOfBins
        Let BinRightEndPoints(i) = MinVal + BinSize * i
    Next i
    
    ' Dump the dataset into TempComputation worksheet
    Let TmpSheet.Range("A1").Resize(UBound(TimeSeries, 1), 1).Value2 = Application.Transpose(TimeSeries)
    
    ' Dump bins', right-end points into TempComputation worksheet
    Let TmpSheet.Range("B1").Resize(NumberOfBins, 1).Value2 = Application.Transpose(BinRightEndPoints)
    
    ' Dump formulas to compute frequencies. Each bin contains the count that is less than or equal to the
    ' right-hand end point of the interval.
    Let TmpSheet.Range("C1").Resize(NumberOfBins, 1).Value2 = _
        Application.Frequency(TmpSheet.Range("A1").Resize(UBound(TimeSeries, 1), 1).Value2, _
                              TmpSheet.Range("B1").Resize(NumberOfBins + 1, 1).Value2)
        
    ' Set value to return. We ignore the last element returned by Frequency() because it
    ' has zero elements.  Since we are using right endpoints for the bins, and the last
    ' right, endpoint is the max of the zero, the NumberOfBins+1 element returns zero.
    Let ComputeDistribution2 = TmpSheet.Range("B1").Resize(NumberOfBins, 2).Value2
End Function

' This function returns the distribution for the given return times series using the bins
' defined by the right-end points passed as the second argument of this function.  The left
' end-point of the first bin is defined as the smallest element in the dataset. The function
' returns an Nx2 matrix.  The left column are the right-end points of the bin intervals.  The
' right column are the counts. BinRightEndPointsArray is expected to be a 1D array.
Function ComputeDistributionWithGivenBins(TimeSeries As Variant, BinRightEndPointsArray As Variant)
    Dim TmpSheet As Worksheet
    Dim MaxVal As Double
    Dim MinVal As Double
    Dim i As Integer
    
    ' Set reference to worksheet TempComputation
    Set TmpSheet = Worksheets("TempComputation")
        
    ' Dump the dataset into TempComputation worksheet
    Let TmpSheet.Range("A1").Resize(UBound(TimeSeries), 1).Value2 = Application.Transpose(TimeSeries)
    
    ' Dump bins', left-end points into TempComputation worksheet
    Let TmpSheet.Range("B2").Resize(UBound(BinRightEndPointsArray), 1).Value2 = Application.Transpose(BinRightEndPointsArray)
    
    ' Dump formulas to compute frequencies. Each bin contains the count that is less than or equal to the
    ' right-hand end point of the interval.
    Let TmpSheet.Range("C1").Resize(UBound(BinRightEndPointsArray), 1).Value2 = _
        Application.WorksheetFunction.Frequency(TmpSheet.Range("A1").Resize(UBound(TimeSeries), 1).Value2, _
                                                TmpSheet.Range("B1").Resize(UBound(BinRightEndPointsArray) - 1, 1).Value2)
        
    ' Set value to return
    Let ComputeDistributionWithGivenBins = TmpSheet.Range("B1").Resize(UBound(BinRightEndPointsArray), 2).Value2
End Function


