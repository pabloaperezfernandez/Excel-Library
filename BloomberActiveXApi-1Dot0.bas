Attribute VB_Name = "BloombergApi"
' PURPOSE OF MODULE
' This module holds code to fetch data from Bloomberg and manipulate it into usable
' formats.  This module requires Bloomberg's deprecated ActiveX DLL. The library's name
' is Bloomberg Data Type Library, and the filename is blpdatax.dll.
'
' The best way to use this library is to import directly and to ensure the application
' directly references the Bloomberg Data Type Library.

Option Explicit
Option Base 1

' DESCRIPTION
' Download FX rates for all calendar days between the given dates inclusively for all of the
' cross rates requested.
'
' PARAMETERS
' 1. TickersArray - An array of Bloomberg cross rates tickers (i.e. EURUSD)
' 2. SerialStartDate - A YYYYMMDD integer representing the start date
' 3. SerialEndDate - A YYYYMMDD integer representing the end date
' 4. lo - ListObject reference to drop data. The first column of this listobject
'    must have hold the dates, the second the "from" ICO currency codes, the third
'    the "to" ISO currency codes, and the last the FX rates between the "from" and
'    the "to" currency codes.
'
' RETURNED VALUE
' Null if an error is encountered. Otherwise, it returns a 2D array with a header row
' comprised of "Date" followed by the tickers for the request cross rates. All other rows
' include the date followed by the cross rates requested on that date.
Public Function DownloadAndConvertFxTimeSeriesToDbTable(TickerList As Variant, _
                                                        StartDate As Long, _
                                                        EndDate As Long, _
                                                        lo As ListObject) As Variant
    Dim FxRates As Variant
    Dim c As Integer
    Dim FromCode As String
    Dim ToCode As String
    Dim TheDates As Variant
    Dim DumpArray As Variant
    Dim lo As ListObject
    
    ' Truncate the listobject
    If Not lo.DataBodyRange Is Nothing Then Call lo.DataBodyRange.Delete
    
    ' Fetch the table of FX rates for all of the cross rates
    Let FxRates = BlpFxRates(TickerList, StartDate, EndDate)
    
    ' Splice the dates column, throwing away the column's header
    ' This converts the column --a 2D array-- to a 1D array
    Let TheDates = Part(FxRates, Span(2, -1), 1)
    
    ' Create a data set to append to the FX listobject the data in each column of
    ' of the data set obtained from Bloomberg. Do this as follows:
    ' 1. Use each column's header to obtain the to and from currency codes
    ' 2. Extract the FX rates from rows 2 to the last row of the column
    ' 3. Create a matrix moving horizontally in time for this currency pairs
    '    FX rates. This requires four rows: dates, from codes, to codes, and FX rates
    ' 4. Transpose this 2D array to get a matrix moving vertically in time with the
    '    following columns: dates, from codes, to codes, and FX rates.
    ' 5. Append the resulting data set to the FX listobject
    For c = 2 To NumberOfColumns(FxRates)
        Let FromCode = StringSplit(Part(TickerList, c - 1), " ", 1)
        Let ToCode = Right(FromCode, 3)
        Let FromCode = Left(FromCode, 3)
        
        Let DumpArray = Pack2DArray(Array(TheDates, _
                                          ConstantArray(FromCode, Length(FxRates) - 1), _
                                          ConstantArray(ToCode, Length(FxRates) - 1), _
                                          Part(FxRates, Span(2, -1), c)))
        Let DumpArray = Transpose(DumpArray)
        
        Call AppendToListObject(DumpArray, lo)
    Next
End Function

' DESCRIPTION
' Download FX rates for all calendar days between the given dates inclusively for all of the
' cross rates requested.
'
' PARAMETERS
' 1. TickersArray - An array of Bloomberg cross rates tickers (i.e. EURUSD)
' 2. SerialStartDate - A YYYYMMDD integer representing the start date
' 3. SerialEndDate - A YYYYMMDD integer representing the end date
'
' RETURNED VALUE
' Null if an error is encountered. Otherwise, it returns a 2D array with a header row comprised
' of "Date" followed by the tickers for the request cross rates. All other rows include the date
' followed by the cross rates requested on that date.
Public Function BlpFxRates(TickersArray As Variant, _
                           SerialStartDate As Long, _
                           SerialEndDate As Long) As Variant
    Dim Blp As New BlpData
    Dim A2DArray() As Variant
    Dim TheData As Variant
    Dim t As Long
    Dim s As Long
    
    
    ' Set option to pull all calendar days
    Let Blp.NonTradingDayValue = PreviousDays
    Let Blp.Periodicity = bbDaily
    
    ' Set option to carry previous day's value when none available
    Let Blp.NonTradingDayValue = PreviousDays
    
    ' Pull the data
    Let TheData = Blp.BLPGetHistoricalData2(TickersArray, "PX_LAST", CStr(SerialStartDate), _
                                            Null, CStr(SerialEndDate))
    
    ' Pre-allocate array to hold values for each factor
    ' The first coordinate is for the header plus the dates
    ' The second coordinate is for the date and data items
    ReDim A2DArray(1 To Length(TheData) + 1, 1 To Length(TickersArray) + 1)
                   
    ' Add headers to the array
    Let A2DArray = SetPart(A2DArray, Prepend(TickersArray, "Date"), 1)
    
    ' Extract and store the cross rates from each date into a single
    ' row of the 2D array to return
    For t = 1 To UBound(TheData, 1) - LBound(TheData, 1) + 1
        ' Extract the date
        Let A2DArray(t + 1, 1) = TheData(t - 1, 0, 0)
    
        For s = 1 To UBound(TheData, 2) - LBound(TheData, 2) + 1
            Let A2DArray(t + 1, s + 1) = TheData(t - 1, s - 1, 1)
        Next
    Next
        
    Let BlpFxRates = A2DArray
End Function

' DESCRIPTION
' Download the requested historical fields for all calendar days, inclusively between the
' given dates. Unfortunately, the time series returned by Bloomberg skips some days.  The
' data is returned as a dictionary indexed by security tickers and mapping to a 2D arrays
'
'
' PARAMETERS
' 1. TickersArray - An array of Bloomberg cross rates tickers (i.e. EURUSD)
' 2. SerialStartDate - Start date in YYYYMMDD format
' 3. SerialEndDate - End date in YYYYMMDD format
'
' RETURNED VALUE
' Null if an error is encountered. A dictionary indexed by currency cross pair tickers with
' the items being 2D arrays (e.g. first column is dates and second column is fx rates.
Public Function BlpDownload(TickersArray As Variant, FieldsArray As Variant, SerialStartDate As Long, _
                            SerialEndDate As Long) As Dictionary
    Dim Blp As New BlpData
    Dim ADict As Dictionary
    Dim A2DArray() As Variant
    Dim TheData As Variant
    Dim t As Long
    Dim s As Long
    Dim i As Long
    
    ' Initialize dictionary
    Set ADict = New Dictionary
    
    ' Set option to pull all calendar days
    Let Blp.Periodicity = bbDaily
    
    ' Set option to carry previous day's value when none available
    Let Blp.NonTradingDayValue = PreviousDays
    
    ' Pull the data
    Let TheData = Blp.BLPGetHistoricalData2(TickersArray, FieldsArray, CStr(SerialStartDate), _
                                            Null, CStr(SerialEndDate))
    
    ' Pre-allocate array to hold values for each factor
    ReDim A2DArray(1 To UBound(TheData, 1) - LBound(TheData, 1) + 2, _
                   1 To UBound(TheData, 3) - LBound(TheData, 3) + 1)
    
    For s = 1 To UBound(TheData, 2) - LBound(TheData, 2) + 1
        For t = 2 To UBound(TheData, 1) - LBound(TheData, 1) + 1
            For i = 1 To UBound(TheData, 3) - LBound(TheData, 3) + 1
                Let A2DArray(t, i) = TheData(t - 2, s - 1, i - 1)
            Next
        Next
        
        ' Add the headers to this security's 2D data array
        Let A2DArray = SetPart(A2DArray, Prepend(FieldsArray, "Date"), 1)
        
        ' Store this security's 2D data array
        Let ADict.Item(Key:=Part(TickersArray, s)) = A2DArray
    Next
    
    ' Return the dictionary holding the Bloomberg data
    Set BlpDownload = ADict
End Function

