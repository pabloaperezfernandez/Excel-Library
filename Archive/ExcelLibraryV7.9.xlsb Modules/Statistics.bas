Attribute VB_Name = "Statistics"
Option Base 1
Option Explicit

' Notes:
' 1. Some of these functions require the workbook to have a worksheet called "TempComputations."
'    This is necessary because the on-demand creation and deletion of a temp worksheet is time
'    consuming.

' All of the array parameter must be 1D arrays of the same length
Public Function ComputeSCurvedZScores(FactorValues() As Double, _
                                      TheWeights() As Double, _
                                      TheCoordinates() As String, _
                                      FactorLowerBound As Double, _
                                      FactorUpperBound As Double, _
                                      Optional SCurveQ As Boolean = True) As Double()
    Dim TempMatrix As Variant
    Dim CoordinatesMeanDict As Dictionary ' to hold weighted mean of each bucket
    Dim CoordinatesStdDevDict As Dictionary ' to hold weighted stdev of each bucket
    Dim CoordinatesWeightDict As Dictionary ' to hold sum of weights in each bucket
    Dim NumSecurities As Long ' Number of securities in data set
    Dim TheCoordinate As String ' To use as a temp string in this function
    Dim BucketWeight As Double ' To hold the weight of current bucket
    Dim BucketMean As Double ' To hold weighted mean factor value of current bucket
    Dim BucketStandardDev As Double ' To hold weighted mean std dev of current bucket
    Dim NormalizedScores() As Double ' To hold array of 2*Normdist(f_i,0,1,True)-1
    Dim i As Long
    Dim j As Long
    Dim InclusionFlag() As Integer
    Dim zScore As Double
    Const NumberOfDecimalPlacesToleranceLevel As Integer = 10

    ' Instantiate required dictionaries
    Set CoordinatesMeanDict = New Dictionary
    Set CoordinatesStdDevDict = New Dictionary
    Set CoordinatesWeightDict = New Dictionary

    ' Compute number of securities for this factor
    Let NumSecurities = Length(FactorValues)
    
    ' Create copy of the array of security coordinates.  We set this copy of a security's combined
    ' coordinate to 0 (e.g. something that is not a valid combined coordinate) so that security is not
    ' used in the computation of weighted standard deviation.
    ReDim InclusionFlag(1 To NumSecurities)
    For i = 1 To NumSecurities
        If FactorValues(i) <= FactorLowerBound Or FactorValues(i) > FactorUpperBound Then
            Let InclusionFlag(i) = 0
        Else
            Let InclusionFlag(i) = 1
        End If
    Next i

    ' Package into a rectagular matrix:
    ' 1. inclusion flag (e.g. security's factor value in or outside of bdries),
    ' 2. a string joining with "-" each security's country ISO code and MSCI sub-industry numeric code
    ' 3. Each security's market cap-relative weight (relative to total universe)
    ' 4. The column of factor values
    ' We do this to we sort by coordinate to get all security's in a bucket together.  This simplifies and
    ' speeds up computations from potentially O(n^2) to O(n log n)
    Let TempMatrix = TransposeMatrix(Pack2DArray(Array(InclusionFlag, TheCoordinates, TheWeights, FactorValues)))
    Let TempMatrix = Sort2DArray(TempMatrix, Array(1, 2), Array(xlDescending, xlDescending), xlNo)
    
    ' Compute and store the weights of each bucket and their weighted, mean, factor values.
    Let TheCoordinate = TempMatrix(1, 2)
    Let BucketWeight = 0
    Let BucketMean = 0
    For i = 1 To NumSecurities
        If TheCoordinate = TempMatrix(i, 2) Then
            ' Add the factor contribution for the current bucket if the coordinate is the same for
            ' this security than the last.
            Let BucketWeight = BucketWeight + TempMatrix(i, 1) * TempMatrix(i, 3)
            Let BucketMean = BucketMean + TempMatrix(i, 1) * TempMatrix(i, 3) * TempMatrix(i, 4)
        ElseIf TempMatrix(i, 1) = 0 Then
            ' If the code gets here, we have reached the first of the excluded companies,
            ' all of which are at the bottom of the matrix.  Since they don't contribute to
            ' mean and std dev of any bucket, we exit the for.
            Call CoordinatesWeightDict.Add(Key:=TheCoordinate, Item:=BucketWeight)
            Call CoordinatesMeanDict.Add(Key:=TheCoordinate, _
                                         Item:=BucketMean / CoordinatesWeightDict.Item(Key:=TheCoordinate))
            Exit For
        Else
            ' If the code gets here, this security's coordinate is different from the prior's.  Store
            ' the sum of the weights of this bucket's included companies and compute its
            ' mean factor value.
            Call CoordinatesWeightDict.Add(Key:=TheCoordinate, Item:=BucketWeight)
            Call CoordinatesMeanDict.Add(Key:=TheCoordinate, _
                                         Item:=BucketMean / CoordinatesWeightDict.Item(Key:=TheCoordinate))
            
            Let TheCoordinate = TempMatrix(i, 2)
            Let BucketWeight = TempMatrix(i, 1) * TempMatrix(i, 3)
            Let BucketMean = TempMatrix(i, 1) * TempMatrix(i, 3) * TempMatrix(i, 4)
        End If
    Next i
    
    ' Check if the coordinate of the last security was added to the dictionary
    If Not CoordinatesWeightDict.Exists(Key:=TheCoordinate) Then
        Call CoordinatesWeightDict.Add(Key:=TheCoordinate, Item:=BucketWeight)
        Call CoordinatesMeanDict.Add(Key:=TheCoordinate, _
                                     Item:=BucketMean / CoordinatesWeightDict.Item(Key:=TheCoordinate))
    End If
    
    ' Loop through all the securities, compute the weighted, squared deviations of each bucket.  When
    ' the bucket coordinate changes, then take the square root of the running sum of the contributions
    ' to compute the standard deviation.
    Let TheCoordinate = TempMatrix(1, 2)
    Let BucketWeight = CoordinatesWeightDict.Item(Key:=TheCoordinate)
    Let BucketMean = CoordinatesMeanDict.Item(Key:=TheCoordinate)
    Let BucketStandardDev = 0
    For i = 1 To NumSecurities
        If TheCoordinate = TempMatrix(i, 2) Then
            Let BucketStandardDev = BucketStandardDev + Round(TempMatrix(i, 1) * (TempMatrix(i, 3) / BucketWeight) * (TempMatrix(i, 4) - BucketMean) ^ 2, NumberOfDecimalPlacesToleranceLevel)
        ElseIf TempMatrix(i, 1) = 0 Then
            Call CoordinatesStdDevDict.Add(Key:=TheCoordinate, Item:=Sqr(BucketStandardDev))
            Exit For
        Else
            Call CoordinatesStdDevDict.Add(Key:=TheCoordinate, Item:=Sqr(BucketStandardDev))
            
            Let TheCoordinate = TempMatrix(i, 2)
            Let BucketWeight = CoordinatesWeightDict.Item(Key:=TheCoordinate)
            Let BucketMean = CoordinatesMeanDict.Item(Key:=TheCoordinate)
            Let BucketStandardDev = Round(TempMatrix(i, 1) * (TempMatrix(i, 3) / BucketWeight) * (TempMatrix(i, 4) - BucketMean) ^ 2, NumberOfDecimalPlacesToleranceLevel)
        End If
    Next i
    
    ' Check if the coordinate of the last security was added to the dictionary
    If Not CoordinatesStdDevDict.Exists(Key:=TheCoordinate) Then
        Call CoordinatesStdDevDict.Add(Key:=TheCoordinate, Item:=Sqr(BucketStandardDev))
    End If
    
    ' Compute S-curved, z-score of each security
    ReDim NormalizedScores(1 To NumSecurities)
    For i = 1 To NumSecurities
        ' Compute the z-score of this security's factor value
        If CoordinatesStdDevDict.Item(TheCoordinates(i)) = 0 Then
            Let NormalizedScores(i) = 0
        Else
            Let zScore = (FactorValues(i) - CoordinatesMeanDict.Item(TheCoordinates(i))) _
                         / CoordinatesStdDevDict.Item(TheCoordinates(i))
            Let NormalizedScores(i) = IIf(SCurveQ, 2 * Application.NormDist(zScore, 0, 1, True) - 1, zScore)
        End If
    Next i

    ' Return array of normalized scores
    Let ComputeSCurvedZScores = NormalizedScores
End Function


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
Public Function ComputeQuantileSizes(NumberOfElements As Long, NumberOfQuantiles As Integer) As Long()
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
Public Function ComputeQuantiles(DataToQuantile As Variant, NumberOfQuantiles As Integer) As Variant()
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
    Let OriginalPositions = Flatten(sht.Range("B1").Resize(UBound(DataToQuantile), 1).Value2)
    
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
    Let ComputeQuantiles = Flatten(TheRange.Columns(3).Value2)
End Function

' Given a factor's vector, its corresponding return vector, and the number of quantiles, this function
' returns the average return for each of the factor's quantile ranks.  In other words, say the factor is P/E and
' the other factor is one-week, forward returns.  Say we are interesting in the average, one-week, forward
' performance of each quintile.  The number of quintiles would be 5, the first factor's vector is the list of P/Es
' (one per security) and the second factor's vector is the list of one-week forward returns for each security.
' This function would return a five-element array with the average, one-week forward return of each security.
Public Function ComputeAveragePerformanceByQuantile(FactorVector As Variant, ReturnVector As Variant, _
                                                    NumberOfQuantiles As Integer) As Variant
    Dim TempSheet As Worksheet
    Dim FactorRange As Range
    Dim ReturnRange As Range
    Dim QuantileRanks As Variant
    Dim QuantileRange As Range
    Dim Results() As Variant
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
    ReDim Results(1 To Length(FactorVector))
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
                                                     NumberOfQuantiles As Integer) As Variant
    Dim i As Integer
    Dim j As Integer
    Dim Results As Variant
    Dim ResultsOneWeek As Variant
    
    ' Pre-allocate the Results array
    ReDim Results(1 To NumberOfQuantiles, 1 To NumberOfColumns(FactorMatrix))
    
    For i = 1 To UBound(FactorMatrix, 2)
        Let ResultsOneWeek = ComputeAveragePerformanceByQuantile(Part(FactorMatrix, Span(1, -1), i), _
                                                                 Part(ReturnMatrix, Span(1, -1), i), _
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
                                                      NumberOfQuantiles As Integer) As Variant
    Dim i As Integer
    Dim Results() As Variant
    
    ' Pre-allocate the Results array
    ReDim Results(1 To UBound(FactorMatrix, 2))
    
    ' Loop through the weeks.
    For i = 1 To UBound(FactorMatrix, 2)
        ' Compute the average performance of each quantile on this week.
        Let Results(i) = ComputeAveragePerformanceByQuantile(Flatten(Application.Index(FactorMatrix, 0, i)), _
                                                             Flatten(Application.Index(ReturnMatrix, 0, i)), _
                                                             NumberOfQuantiles)
                                                             
        ' Turn this week's average quantile performance into a 1D array
        Let Results(i) = Flatten(Results(i))
    Next i
    
    ' Set the results matrix as the return value of the function
    Let ComputeQuantileAveragePerformanceTimeSeries2 = TransposeMatrix(Results)
End Function

' This function returns the distribution for the given return times series using the
' Frequency() worksheet funciton.  This divides the range of values evenly into NumberOfBins
' bins.  The function returns an Nx2 matrix.  The left column are the right-end points of the
' bin intervals.  The right column are the counts.
'
' The algorithm does the following:
' 1. This data is divided into N (equal to NumberOfBins) bins of equal size
' 2. The number of results returned by the function is one more than the number of bins
'
' If the bins are (b_1, ..., b_n) and the data is (d_1, ..., d_m), the counts are
' (c_1, ..., c_m, c_{m+1}) and are given by
' c_i = size of {d in (d_1, ..., d_m) s.t. b_{i-1}<d<=b_i}, where b_0=\infty and
' b_{n+1}=\infty
Function ComputeDistribution(TimeSeries As Variant, NumberOfBins As Integer) As Variant
    Dim MaxVal As Double
    Dim MinVal As Double
    Dim BinSize As Double
    Dim BinRightEndPoints() As Double
    Dim i As Integer
    
    ' Determine the max and min values of time series
    Let MaxVal = Application.Max(TimeSeries)
    Let MinVal = Application.Min(TimeSeries)
    
    ' Determine bin size
    Let BinSize = (MaxVal - MinVal) / NumberOfBins
    
    ' Create the sequence of right, end-points for the bins
    Let BinRightEndPoints = Add(Multiply(NumericalSequence(1, CLng(NumberOfBins)), _
                                                                          BinSize), _
                                                MinVal)
        
    ' Set value to return. We ignore the last element returned by Frequency() because it
    ' has zero elements.  Since we are using right endpoints for the bins, and the last
    ' right, endpoint is the max of the zero, the NumberOfBins+1 element returns zero.
    Let ComputeDistribution = Flatten(Application.Frequency(TimeSeries, BinRightEndPoints))
End Function

' Same as ComputeDistribution() in this module but using the given bins
Function ComputeDistributionWithGivenBins(TimeSeries As Variant, BinRightEndPointsArray As Variant)
    Let ComputeDistributionWithGivenBins = Flatten(Application.Frequency(TimeSeries, BinRightEndPointsArray))
End Function


