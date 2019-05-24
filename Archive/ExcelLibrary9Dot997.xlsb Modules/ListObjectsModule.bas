Attribute VB_Name = "ListObjectsModule"
Option Explicit
Option Base 1

' Turn a range of cells into a list object
' When provided, the list object's name is set to the user's choice
' Unless explicit set to False,UserCurrentRegionQ defaults to True
' and allows the user to specify just one cell in the range to
' convert to a list object. The range used in this case is the cell's
' current region. If there is any problem with the parameters, the
' function returns Nothing
Public Function AddListObject(ARangeInCurrentRegion As Range, _
                              Optional ListObjectName As String = Empty, _
                              Optional UseCurrentRegionQ As Boolean = True) As ListObject
    Dim lo As ListObject
    Dim wsht As Worksheet
    
    Set wsht = ARangeInCurrentRegion.Parent
    
    ' Exit if ARange has not been initialized
    If ARangeInCurrentRegion Is Nothing Then
        Set AddListObject = Nothing
        Exit Function
    End If
    
    Set lo = wsht.ListObjects.Add(SourceType:=xlSrcRange, _
                                  Source:=IIf(UseCurrentRegionQ, ARangeInCurrentRegion.CurrentRegion, ARangeInCurrentRegion), _
                                  XlListObjectHasHeaders:=xlYes)
    
    If ListObjectName <> Empty Then Let lo.Name = ListObjectName
    
    Set AddListObject = lo
End Function

' The purpose of this sub is to add column with the given names to the given list object.
' Moreover, data is passed, it is dumped in each of the columns.  Each column in the data
' is dumped in one of the new columns. The function exists with Nothing if their is a
' problem with the arguments. Otherwise, it returns a reference to the listobject
'
' This has been written this to return a reference to the modified ListObject. However,
' it is impossible to avoid the sideeffect of altering the object referenced by the
' AListObject parameter.
'
' Example:
' AListObject has data [{1,2,3; 10, 20, 30; 100, 200, 300}]
' ColumnNames is Array("Col1", "Col2")
' TheData is Array(Col1, Col2), with
'                Col1 = Array(1000, 2000, 3000) and
'                Col2 = Array(10000, 20000, 30000)
Public Function AddColumnsToListObject(aListObject As ListObject, _
                                       ColumnNames() As String, _
                                       Optional TheData As Variant) As ListObject
    Dim var As Variant
    Dim i As Long
    Dim lc As ListColumn
                                       
    ' Exit with null if there is a problem with the arguments
    If aListObject Is Nothing Then
        Set AddColumnsToListObject = Nothing
        Exit Function
    End If
    
    ' Exit if ColumnNames is not an array, ColumnNames is an empty array, or any of the
    ' elements in the array is not a string
    If EmptyArrayQ(ColumnNames) Or Not StringArrayQ(ColumnNames) Then
        Set AddColumnsToListObject = Nothing
        Exit Function
    End If
    
    ' If param TheData is present, then it must be a 1D array of the same length as ColumnNames
    If Not IsMissing(TheData) Then
        If EmptyArrayQ(TheData) Or Length(TheData) <> Length(ColumnNames) Then
            Set AddColumnsToListObject = Nothing
            Exit Function
        End If
        
        For Each var In TheData
            If NumberOfDimensions(var) <> 1 Then
                Set AddColumnsToListObject = Nothing
                Exit Function
            End If
            
            If Length(var) <> aListObject.ListRows.Count Then
                Set AddColumnsToListObject = Nothing
                Exit Function
            End If
        Next
    End If
    
    ' If the code gets here, all parameters are consistent.
    
    If IsMissing(TheData) Then
        For Each var In ColumnNames
            Let aListObject.ListColumns.Add.Name = var
        Next
        
        Set AddColumnsToListObject = aListObject
        Exit Function
    End If
    
    ' If the code gets here, optional parameter TheData has been provided
    For i = 1 To Length(ColumnNames)
        Set lc = aListObject.ListColumns.Add
        Let lc.Name = ColumnNames(i)
        Call DumpInSheet(TransposeMatrix(TheData(i)), lc.DataBodyRange(1, 1))
    Next
    
    Set AddColumnsToListObject = aListObject
End Function

' The purpose of this function is to extend a2Darray1 with data from a2Darray2 using
' equality on the given key columns.  The function returns the "left joined" 2D array.
' This means that all rows in array1 are included. The data from array2 is included
' only if its key in also in array 1.
'
' If a key if found more than once in a2Darray, the first ocurrance is used. The resulting
' 2D array uses data from the columns in a2DArray1 specified in ColsPosArrayFrom2DArray1
' and the columns in a2DArray2 specified in ColsPosArrayFrom2DArray2
'
' When the parameters are inconsistent, the function returns Null
Public Function LeftJoinListObjectsOnKeyEquality(ListObject1 As ListObject, _
                                                 Array1Key As String, _
                                                 Array1HeadersArray() As String, _
                                                 ListObject2 As ListObject, _
                                                 Array2Key As String, _
                                                 Array2HeadersArray() As String, _
                                                 Optional ReturnAsDictOfArraysQ As Boolean = False) As Variant
    Dim Array1HeadersPos() As Integer
    Dim Array2HeadersPos() As Integer
    Dim NumRequestedCols1 As Integer
    Dim NumRequestedCols2 As Integer
    Dim StringArray() As String
    Dim ResultsDict As Dictionary
    Dim ListObject2TrackingDict As Dictionary
    Dim r As Long
    Dim var As Variant
    Dim TheKey As Variant
    Dim TheItems() As Variant
    Dim AppendedItems() As Variant
    Dim JoinedHeadersRow As Variant
    Dim TheResults As Variant

    ' Parameter consistency checks
    
    ' Exit with Null if either a2Array1, a2Array2, ColsPosArrayFrom2DArray1
    ' ColsPosArrayFrom2DArray2 is not an array
    If ListObject1.ListRows.Count < 2 Then
        If ReturnAsDictOfArraysQ Then
            Set LeftJoinListObjectsOnKeyEquality = Nothing
        Else
            Let LeftJoinListObjectsOnKeyEquality = Null
        End If
        
        Exit Function
    End If
    
    ' Exit with Null if any other headers in Array1HeadersArray() is not a header in ListObject1
    Let StringArray = Cast(Flatten(ListObject1.HeaderRowRange.Value2), xlParamTypeChar)
    For Each var In Array1HeadersArray
        If FreeQ(StringArray, CStr(var)) Then
            If ReturnAsDictOfArraysQ Then
                Set LeftJoinListObjectsOnKeyEquality = Nothing
            Else
                Let LeftJoinListObjectsOnKeyEquality = Null
            End If
            
            Exit Function
        End If
    Next
    
    ' Exit with Null if any other headers in Array2HeadersArray() is not a header in ListObject2
    Let StringArray = Cast(Flatten(ListObject2.HeaderRowRange.Value2), xlParamTypeChar)
    For Each var In Array2HeadersArray
        If FreeQ(StringArray, CStr(var)) Then
            If ReturnAsDictOfArraysQ Then
                Set LeftJoinListObjectsOnKeyEquality = Nothing
            Else
                Let LeftJoinListObjectsOnKeyEquality = Null
            End If
            
            Exit Function
        End If
    Next
    
    ' Get the column indices of the headers for ListObject1 and ListObject2
    ReDim Array1HeadersPos(1 To Length(Array1HeadersArray))
    Let r = 1
    For Each var In Array1HeadersArray
        Let Array1HeadersPos(r) = ListObject1.ListColumns(var).Index
        Let r = r + 1
    Next
    
    ReDim Array2HeadersPos(1 To Length(Array2HeadersArray))
    Let r = 1
    For Each var In Array2HeadersArray
        Let Array2HeadersPos(r) = ListObject2.ListColumns(var).Index
        Let r = r + 1
    Next
    
    ' Count the number of requested columns in each list object
    Let NumRequestedCols1 = Length(Array1HeadersArray)
    Let NumRequestedCols2 = Length(Array2HeadersArray)

    ' Load all information from ListObject1 into a dictionary
    ReDim TheItems(1 To Length(Array1HeadersArray) + Length(Array2HeadersArray))
    
    Set ResultsDict = New Dictionary
    For r = 1 To ListObject1.ListRows.Count
        ' Get the for the current row
        Let TheKey = ListObject1.ListColumns(Array1Key).DataBodyRange(r, 1).Value2

        If Not ResultsDict.Exists(Key:=TheKey) Then
            ' Extract the columns needed from this security's row
            If ReturnAsDictOfArraysQ Then
                ' The key is not added when the result is requested as a dictionary
                Let TheItems = Take(Flatten(ListObject1.DataBodyRange.Rows(r).Value2), _
                                    Array1HeadersPos)
            Else
                ' The key is added when the result is requested as a dictionary
                Let TheItems = Take(Flatten(ListObject1.DataBodyRange.Rows(r).Value2), _
                                    Prepend(Array1HeadersPos, ListObject1.ListColumns(Array1Key).Index))
            End If
            
            ' Pad TheItems with enough slots for the items appended from a2DArray2
            Let TheItems = ConcatenateArrays(TheItems, ConstantArray(Empty, CLng(NumRequestedCols2)))

            ' Add the array of values to this security's entry
            Call ResultsDict.Add(Key:=TheKey, Item:=TheItems)
        End If
    Next r

    ' Scan a2DArray2 appending to the array of elements of each element in a2DArray1 the
    ' elements in a2DArray2
    Set ListObject2TrackingDict = New Dictionary
    For r = 1 To ListObject2.ListRows.Count
        ' Get the for the current row
        Let TheKey = ListObject2.ListColumns(Array2Key).DataBodyRange(r, 1).Value2
        
        ' Append to the items in the results dicts for the current key if the current key
        ' is in the results dictionary already, and they has not already been appended
        If ResultsDict.Exists(Key:=TheKey) And Not ListObject2TrackingDict.Exists(Key:=TheKey) Then
            ' Mark this row in ListObject2 as having been processed
            Call ListObject2TrackingDict.Add(Key:=TheKey, Item:=Empty)
            
            ' Take the portion of the items corresponding to array 1
            If ReturnAsDictOfArraysQ Then
                Let TheItems = Take(ResultsDict.Item(Key:=TheKey), NumRequestedCols1)
            Else
                Let TheItems = Take(ResultsDict.Item(Key:=TheKey), 1 + NumRequestedCols1)
            End If
            
            ' Get the required columns from this row to append to those already in the results
            ' dictionary
            Let AppendedItems = Take(Flatten(ListObject2.DataBodyRange.Rows(r).Value2), _
                                     Array2HeadersPos)
            
            Let ResultsDict.Item(Key:=TheKey) = ConcatenateArrays(TheItems, AppendedItems)
        End If
    Next r

    ' Prepend headers to return matrix
    If ReturnAsDictOfArraysQ Then
        ' The user requested that the result be returned as a dictionary or arrays
        Set LeftJoinListObjectsOnKeyEquality = ResultsDict
        Exit Function
    End If
    
    ' Repack the results as a 2D array
    Let TheResults = Pack2DArray(ResultsDict.Items)
    
    ' Prepend the headers row if the user chose to
    Let JoinedHeadersRow = Prepend(ConcatenateArrays(Array1HeadersArray, Array2HeadersArray), _
                                   Array1Key)
    
    ' The user requested the result as a 2D array
    Let LeftJoinListObjectsOnKeyEquality = Prepend(TheResults, JoinedHeadersRow)
End Function

' The purpose of this function is to extend ListObject1 with data from ListObject2 using
' equality on the given key columns.  The function returns the "left joined" 2D array.
' This means that all rows in array1 are included. The data from array2 is included
' only if its key in also in array 1.
'
' If a key if found more than once in a2Darray, the first ocurrance is used.  The resulting
' 2D array uses data from the columns in ListObject1 specified in ColsPosArrayFrom2DArray1
' and the columns in ListObject2 specified in ColsPosArrayFrom2DArray2
'
' When the parameters are inconsistent, the function returns Null
Public Function InnerJoinListObjectsOnKeyEquality(ListObject1 As ListObject, _
                                                  Array1Key As String, _
                                                  Array1HeadersArray() As String, _
                                                  ListObject2 As ListObject, _
                                                  Array2Key As String, _
                                                  Array2HeadersArray() As String, _
                                                  Optional ReturnAsDictOfArraysQ As Boolean = False) As Variant
    Dim Array1HeadersPos() As Integer
    Dim Array2HeadersPos() As Integer
    Dim NumRequestedCols1 As Integer
    Dim NumRequestedCols2 As Integer
    Dim StringArray() As String
    Dim ResultsDict As Dictionary
    Dim ListObject2Dict As Dictionary
    Dim r As Long
    Dim var As Variant
    Dim TheKey As Variant
    Dim TheItems() As Variant
    Dim AppendedItems() As Variant
    Dim JoinedHeadersRow As Variant
    Dim TheResults As Variant

    ' Parameter consistency checks

    ' Exit with Null if either a2Array1, a2Array2, ColsPosArrayFrom2DArray1
    ' ColsPosArrayFrom2DArray2 is not an array
    If ListObject1.ListRows.Count < 2 Then
        If ReturnAsDictOfArraysQ Then
            Set InnerJoinListObjectsOnKeyEquality = Nothing
        Else
            Let InnerJoinListObjectsOnKeyEquality = Null
        End If
        
        Exit Function
    End If
    
    ' Exit with Null if any other headers in Array1HeadersArray() is not a header in ListObject1
    Let StringArray = Cast(Flatten(ListObject1.HeaderRowRange.Value2), xlParamTypeChar)
    For Each var In Array1HeadersArray
        If FreeQ(StringArray, CStr(var)) Then
            If ReturnAsDictOfArraysQ Then
                Set InnerJoinListObjectsOnKeyEquality = Nothing
            Else
                Let InnerJoinListObjectsOnKeyEquality = Null
            End If
            
            Exit Function
        End If
    Next
    
    ' Exit with Null if any other headers in Array2HeadersArray() is not a header in ListObject2
    Let StringArray = Cast(Flatten(ListObject2.HeaderRowRange.Value2), xlParamTypeChar)
    For Each var In Array2HeadersArray
        If FreeQ(StringArray, CStr(var)) Then
            If ReturnAsDictOfArraysQ Then
                Set InnerJoinListObjectsOnKeyEquality = Nothing
            Else
                Let InnerJoinListObjectsOnKeyEquality = Null
            End If
            
            Exit Function
        End If
    Next
    
    ' Get the column indices of the headers for ListObject1 and ListObject2
    ReDim Array1HeadersPos(1 To Length(Array1HeadersArray))
    Let r = 1
    For Each var In Array1HeadersArray
        Let Array1HeadersPos(r) = ListObject1.ListColumns(var).Index
        Let r = r + 1
    Next
    
    ReDim Array2HeadersPos(1 To Length(Array2HeadersArray))
    Let r = 1
    For Each var In Array2HeadersArray
        Let Array2HeadersPos(r) = ListObject2.ListColumns(var).Index
        Let r = r + 1
    Next
    
    ' Count the number of requested columns in each list object
    Let NumRequestedCols1 = Length(Array1HeadersArray)
    Let NumRequestedCols2 = Length(Array2HeadersArray)

    ' Index the contents of array2
    Set ListObject2Dict = New Dictionary
    For r = 1 To ListObject2.ListRows.Count
        ' Get the for the current row
        Let TheKey = ListObject2.ListColumns(Array2Key).DataBodyRange(r, 1).Value2
        
        If Not ListObject2Dict.Exists(Key:=TheKey) Then
            ' Add the array of values to this security's entry
            Call ListObject2Dict.Add(Key:=TheKey, _
                                     Item:=Take(Flatten(ListObject2.ListRows(r).Range.Value2), _
                                                Array2HeadersPos))
        End If
    Next r

    ' Scan ListObject2 appending to the array of elements of each element in ListObject1 the
    ' elements in ListObject2
    Set ResultsDict = New Dictionary
    For r = 1 To ListObject1.ListRows.Count
        ' Get the for the current row
        Let TheKey = ListObject1.ListColumns(Array1Key).DataBodyRange(r, 1).Value2

        ' Create the join for this row in array1 if it is found in array2 based on the key.
        If Not ResultsDict.Exists(Key:=TheKey) And ListObject2Dict.Exists(Key:=TheKey) Then
            ' Extract the columns required from this row in array 1
            If ReturnAsDictOfArraysQ Then
                Let TheItems = Take(Flatten(ListObject1.ListRows(r).Range.Value2), _
                                    Array1HeadersPos)
            Else
                Let TheItems = Take(Flatten(ListObject1.ListRows(r).Range.Value2), _
                                    Prepend(Array1HeadersPos, ListObject1.ListColumns(Array1Key).Index))
            End If

            ' Get the corresponding items from array 2
            Let AppendedItems = ListObject2Dict.Item(Key:=TheKey)

            ' Index the joined items for this row in array 1
            Call ResultsDict.Add(Key:=TheKey, Item:=ConcatenateArrays(TheItems, AppendedItems))
        End If
    Next r

    If ReturnAsDictOfArraysQ Then
        Set InnerJoinListObjectsOnKeyEquality = ResultsDict
        Exit Function
    End If
    
    ' Repack the results as a 2D array
    Let TheResults = Pack2DArray(ResultsDict.Items)

    ' Prepend the headers row if the user chose to
    Let JoinedHeadersRow = Prepend(ConcatenateArrays(Array1HeadersArray, Array2HeadersArray), _
                                   Array1Key)

    ' Prepend headers to return matrix
    Let InnerJoinListObjectsOnKeyEquality = Prepend(TheResults, JoinedHeadersRow)
End Function

' The purpose of this function is to return a dictionary indexing the items in paramater
' TheItems using the elements of TheKeys.  Both parameters must be 1D arrays.  2D arrays
' are disallowed unless they are 1D arrays of arrays.  TheKeys and TheItems must be
' non-empty and have the same length.  No element of TheKeys can be an array because
' arrays are not allowed as keys in dictionaries.
'
' If the parameters are inconsistent, the function returns Nothing
Public Function CreateListObjectDictionary(aListObject As ListObject, _
                                           IndexColumnName As String, _
                                           ItemColumnNames As Variant, _
                                           Optional RowsAsDictionariesQ As Variant = True) As Dictionary
    Dim ADict As Dictionary
    Dim RowDict As Dictionary
    Dim r As Long
    Dim c As Long
    Dim var As Variant
    Dim TheKey As Variant
    Dim TempVariantArray() As Variant
    
    ' check inputs consistency
    
    ' Exit with Null if the list object has not list rows
    If aListObject.ListRows.Count = 0 Then
        Set CreateListObjectDictionary = Nothing
        Exit Function
    End If
    
    ' Exit with Null if IndexColumnName does not correspond to a
    ' column in AListObject
    If Not MemberQ(Flatten(aListObject.HeaderRowRange.Value2), IndexColumnName) Then
        Set CreateListObjectDictionary = Nothing
        Exit Function
    End If
    
    ' Exit with Null if ItemColumnNames is not an array
    If Not IsArray(ItemColumnNames) Then
        Set CreateListObjectDictionary = Nothing
        Exit Function
    End If
    
    ' Exit with Null with Null if ItemColumnNames
    If EmptyArrayQ(ItemColumnNames) Then
        Set CreateListObjectDictionary = Nothing
        Exit Function
    End If
    
    ' Exit with Null if either ItemColumnNames are not arrays or empty
    If Not StringArrayQ(ItemColumnNames) Then
        Set CreateListObjectDictionary = Nothing
        Exit Function
    End If
    
    ' Exit with Null if any one of the names in ItemColumnNames does not correspond
    ' to a column name in AListObject
    For Each var In ItemColumnNames
        If Not MemberQ(Flatten(aListObject.HeaderRowRange.Value2), CStr(var)) Then
        Set CreateListObjectDictionary = Nothing
            Exit Function
        End If
    Next
    
    ' If the code gets to this point, all inputs are consistent
    Set ADict = New Dictionary
    
    For r = 1 To aListObject.ListRows.Count
        Let TheKey = aListObject.ListColumns(IndexColumnName).DataBodyRange(r, 1).Value2
        
        If Not ADict.Exists(Key:=TheKey) Then
            If RowsAsDictionariesQ Then
                Set RowDict = New Dictionary
                
                For Each var In ItemColumnNames
                    Call RowDict.Add(Key:=CStr(var), _
                                     Item:=aListObject.ListColumns(var).DataBodyRange(r, 1).Value2)
                Next
                
                Call ADict.Add(Key:=TheKey, Item:=RowDict)
            Else
                ReDim TempVariantArray(LBound(ItemColumnNames) To UBound(ItemColumnNames))
                
                For c = LBound(ItemColumnNames) To UBound(ItemColumnNames)
                    Let var = ItemColumnNames(c)
                    Let TempVariantArray(c) = aListObject.ListColumns(var).DataBodyRange(r, 1).Value2
                Next c
                
                Call ADict.Add(Key:=TheKey, Item:=TempVariantArray)
            End If
        End If
    Next r
    
    Set CreateListObjectDictionary = ADict
End Function

' The purpose of this routine is to export the given listobject as a CSV file.
' The optional flag (set to False by default) determines if the header row is exported.
Public Sub ExportListObjectAsTsvFile(aListObject As ListObject, _
                                     TheFullPathFileName As String, _
                                     Optional IncludeHeaderRowQ As Boolean = False)
    Dim r As Long
    Dim c As Long
    Dim ARow As String
    
    ' Open the file
    Open TheFullPathFileName For Output As #1

    If IncludeHeaderRowQ Then
        Let ARow = ""
        For c = 1 To aListObject.ListColumns.Count - 1
            Let ARow = ARow & aListObject.HeaderRowRange(1, c).Value2 & vbTab
        Next c
        Let ARow = ARow & aListObject.HeaderRowRange(1, aListObject.ListColumns.Count).Value2
        Print #1, ARow
    End If
        
    ' Write the listrows
    For r = 1 To aListObject.ListRows.Count
        Let ARow = ""
        For c = 1 To aListObject.ListColumns.Count - 1
            Let ARow = ARow & aListObject.DataBodyRange(r, c).Value2 & vbTab
        Next c
        Let ARow = ARow & aListObject.DataBodyRange(r, aListObject.ListColumns.Count).Value2
        Print #1, ARow
    Next r
    
    ' Close the file
    Close #1
End Sub

' This function deletes the databodyrange of a list object
' .DataBodyRange will point to Nothing after calling this
' subroutine
Public Sub ClearListObjectDataBodyRange(lo As ListObject)
    If lo.ListRows.Count > 0 Then Call lo.DataBodyRange.Delete
End Sub

' DESCRIPTION
' This function appends the given 1D or 2D array to a listobject.
'
' PARAMETERS
' None
'
' RETURNED VALUE
' Nothing but it appends the given data to the listobject passed.
Public Sub AppendToListObject(TheData As Variant, lo As ListObject)
    Call DumpInSheet(TheData, lo.Range(1, 1).Offset(1 + lo.ListRows.Count, 0))
End Sub

' DESCRIPTION
' This function applies the rule to use the current value if numeric or the
' latest numeric value available to the desired column in a given listobject.
' The table is required to have a column called Date that is numerically
' sortable.  The table must also have an identifier column that is unique.
' We are using FIGIs here.
'
' PARAMETERS
' 1. lo - A reference to a listobject
' 2. ColumnName - Name of the column to act on
' 3. SortByDateQ - Sort in ascending order if True. False by default
'
' RETURNED VALUE
' No value returned
Public Sub ApplyLastAvailableDataRuleToListObject(lo As ListObject, ColumnName As String, _
                                                  Optional SortByDateQ As Boolean = False)
    Dim DataDict As Dictionary
    Dim DumpDict As Dictionary
    Dim i As Long
    Dim AFigi As String
    Dim TheData As Variant
    
    ' Exit if the listobject or its DataBodyRange are empty
    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub
    
    ' Create instances of dictionaries
    Set DataDict = New Dictionary
    Set DumpDict = New Dictionary
    
    ' Ensure listobject is sorted according to date
    If SortByDateQ Then Call SortListObject(lo, "Date")
    
    ' Apply the last available data rule by scanning the column in time ascending order
    For i = 1 To lo.ListRows.Count
        ' Pull FIGI, date, and price for current listrow
        With lo
            Let AFigi = .ListColumns("ID_BB_GLOBAL").DataBodyRange(i, 1).Value2
            Let TheData = .ListColumns(ColumnName).DataBodyRange(i, 1).Value2
        End With
        
        ' Store current price if numeric
        If NumberQ(TheData) Then Let DataDict.Item(Key:=AFigi) = TheData
        
        ' Store the price to use for the current listrow
        Call DumpDict.Add(Key:=i, Item:=DataDict.Item(Key:=AFigi))
    Next
    
    ' Dump the prices into the listobject
    Call DumpInSheet(Transpose(DumpDict.Items), lo.ListColumns(ColumnName).DataBodyRange(1, 1))
End Sub

' DESCRIPTION
' Sorts the given listobject in the desired direction along the chosen listcolumn.
'
' PARAMETERS
' None
'
' RETURNED VALUE
' No value returned
Public Sub SortListObject(lo As ListObject, FieldName As String, Optional TheOrder As XlSortOrder = xlAscending)
    ' Exit if the listobject or its DataBodyRange are empty
    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub

    ' Clear all prior sorting criteria
    Call lo.Sort.SortFields.Clear
    
    ' Add sorting criteria
    Call lo.Sort.SortFields.Add(Key:=lo.ListColumns(FieldName).Range, _
                                SortOn:=xlSortOnValues, _
                                Order:=TheOrder, _
                                DataOption:=xlSortNormal)
    
    ' Perform the actual sort
    With lo.Sort
        Let .Header = xlYes
        Let .MatchCase = False
        Let .Orientation = xlTopToBottom
        Let .SortMethod = xlPinYin
        Call .Apply
    End With
End Sub

' DESCRIPTION
' This function restores the formulas stored in the comments of the listobject's headers row.
' Columns with no comments are ignored.
'
' PARAMETERS
' None
'
' RETURNED VALUE
' No value is returned
Public Sub RestoreListObjectFormulas(lo As ListObject)
    Dim c As Integer
    
    ' Extract formula for each column header and insert it in the corresponding columns
    For c = 1 To lo.HeaderRowRange.Count
        If Not (lo.HeaderRowRange(1, c).Comment Is Nothing) Then
            Let lo.ListColumns(c).DataBodyRange.Formula = lo.HeaderRowRange(1, c).Comment.Text
        End If
    Next
End Sub

' DESCRIPTION
' Stores in the comments of the headers row the formulas in the first listrow.
' Columns without formulas are ignored.
'
' Note that this function wipes all pre-exisiting comments in the headers row.
'
' PARAMETERS
' None
'
' RETURNED VALUE
' No value is returned
Public Sub SaveListObjectFormulas(lo As ListObject)
    Dim c As Integer

    ' Store formulas
    For c = 1 To lo.HeaderRowRange.Count
        ' Store the formulas of every column that has a formulas in its first cell
        If lo.DataBodyRange(1, c).HasFormula Then
            ' Clear any previously stored comments
            Call lo.HeaderRowRange(1, c).ClearComments
        
            If Left(Trim(lo.DataBodyRange(1, c).Formula), 1) = "=" Then
                Call lo.HeaderRowRange(1, c).AddComment(lo.DataBodyRange(1, c).Formula)
            End If
        End If
    Next
    
    ' Copy and paste the databodyrange as values
    Let lo.DataBodyRange.Value2 = lo.DataBodyRange.Value2
End Sub

' DESCRIPTION
' This sub applies to the listobject holding the given selected cell.
' This function restores the formulas stored in the comments of the listobject's headers row.
' Columns with no comments are ignored.
'
' PARAMETERS
' None
'
' RETURNED VALUE
' No value is returned
Public Sub RestoreTableFormulas()
    Call RestoreListObjectFormulas(Selection.ListObject)
End Sub

' DESCRIPTION
' This sub applies to the listobject holding the given selected cell.
' Stores in the comments of the headers row the formulas in the first listrow.
' Columns without formulas are ignored.
'
' Note that this function wipes all pre-exisiting comments in the headers row.
'
' PARAMETERS
' None
'
' RETURNED VALUE
' No value is returned
Public Sub SaveTableFormulas()
    Call SaveListObjectFormulas(Selection.ListObject)
End Sub

' DESCRIPTION
' This function returns a 1D array with the formulas stored in the header row
' of a listobject. The returned value corresponding to listcolumns with no formula
' stored in the header is VbNulLString.  If the parameter lo is not initialized,
' the function returns Null as an error code.
'
' PARAMETERS
' 1. lo - A reference to the a listobject
'
' RETURNED VALUE
' Null in case of error and a 1D array of strings with the formulas stored in the
' listobject's header row.
Public Function GetListObjectHeaderFormulas(lo As ListObject) As Variant
    Dim c As Integer
    Dim Results As Variant
    
    ' Set default error return code
    Let GetListObjectHeaderFormulas = Null
    
    ' Exit with an error code if lo has not been initialized
    If lo Is Nothing Then Exit Function
    
    ' Pre-allocate array to hold header formulas
    ReDim Results(1 To lo.ListColumns.Count)
    
    ' Extract formula for each column header and insert it in the corresponding columns
    For c = 1 To lo.HeaderRowRange.Count
        If lo.HeaderRowRange(1, c).Comment Is Nothing Then
            Let Results(c) = vbNullString
        Else
            Let Results(c) = lo.HeaderRowRange(1, c).Comment.Text
        End If
    Next
    
    Let GetListObjectHeaderFormulas = Results
End Function

' DESCRIPTION
' This function returns a 1D array with the formulas stored in the header row
' of a listobject. The returned value corresponding to listcolumns with no formula
' stored in the header is VbNulLString.  If the parameter lo is not initialized,
' the function returns Null as an error code.
'
' PARAMETERS
' 1. lo - A reference to the a listobject
'
' RETURNED VALUE
' Null in case of error and a 1D array of strings with the formulas stored in the
' listobject's header row.
Public Function SetListObjectHeaderFormulas(lo As ListObject, String1DArray As Variant) As Variant
    Dim c As Integer
    
    ' Set default error return code
    Let SetListObjectHeaderFormulas = Null
    
    ' Exit with an error code if lo has not been initialized
    If lo Is Nothing Then Exit Function
    
    ' Exit with error code if String1DArray is not a string of array of the same width as lo
    If Not StringArrayQ(String1DArray) Then Exit Function
    If Length(String1DArray) <> lo.ListColumns.Count Then Exit Function
    
    ' Extract formula for each column header and insert it in the corresponding columns
    For c = 1 To lo.HeaderRowRange.Count
        ' Clear any previously stored comments
        Call lo.HeaderRowRange(1, c).ClearComments
        Call lo.HeaderRowRange(1, c).AddComment(Part(String1DArray, c))
    Next
    
    ' Return True to indicate success
    Let SetListObjectHeaderFormulas = True
End Function

' DESCRIPTION
' Trims the values in the given listcolumn.  The number format of the
' column us assumed to be that of the first cell in its DataBodyRange.
' The entire column is set to this number format.
'
' PARAMETERS
' 1. lc - A reference to a listcolumn
'
' RETURNED VALUE
' N/A
Public Sub TrimListColumn(lc As ListColumn)
    Dim DumpArray() As Variant
    Dim r As Long
    Dim ListColumnNumberFormat As String
    
    ' Store the number format of the list column
    Let ListColumnNumberFormat = lc.DataBodyRange(1, 1).NumberFormat
    
    ' Pre-allocate array to read listcolumn data and trim values
    ReDim DumpArray(1 To lc.Parent.ListRows.Count, 1 To 1)
    
    ' Loop through cells in listcolumn, storing trimmed values
    For r = 1 To lc.Parent.ListRows.Count
        Let DumpArray(r, 1) = Trim(lc.DataBodyRange(r, 1).Value2)
    Next
    
    ' Pre-format listcolumn as text to avoid dropping leading 0s if
    ' the listcolumn is not formatted as text
    Let lc.DataBodyRange.NumberFormat = "@"
    
    ' Update listcolumn with trimmed values
    Call DumpInSheet(DumpArray, lc.DataBodyRange(1, 1))
    
    ' Return the list column to its original number format
    Let lc.DataBodyRange.NumberFormat = ListColumnNumberFormat
End Sub

' DESCRIPTION
' Returns a 2D array with the visible data in the given listobject.
'
' PARAMETERS
' 1. lo - A reference to a listobject
'
' RETURNED VALUE
' A 2D array with the visible value
Public Function GetVisibleSpecialValues(lo As ListObject) As Variant
    Dim TheAddresses As Variant
    Dim VarString As Variant
    Dim ResultArray As Variant
    
    Let TheAddresses = Split(lo.Range.SpecialCells(xlCellTypeVisible).Address, ",")
    
    Let ResultArray = lo.HeaderRowRange.Value2
    For Each VarString In TheAddresses
        Let ResultArray = StackArrays(ResultArray, lo.Range.Worksheet.Range(VarString).Value2)
    Next
    
    Let GetVisibleSpecialValues = ResultArray
End Function

' Executes queries and dumps the resulting record set the chosen listobject
'
' PARAMETERS
' 1. lo - A reference to a listobject
' 2. DataGetterFunctionName - The name of a function to perform a querry and return
'    the results.
' 3. WorksheetPassword (Optional, vbNullString default) - Unprotects and the protects given
'    worksheet with given password if not a null string.
'
' RETURNED VALUE
' Dumped querry results in chosen listobject
Private Sub DumpDataInListObject(lo As ListObject, _
                                 DataGetterFunctionName As String, _
                                 Optional WorksheetPassword As String = vbNullString)
    If WorksheetPassword <> vbNullString Then Call UnprotectWorksheet(lo.Parent)
    
    Let lo.ShowAutoFilter = False
    Call ClearListObjectDataBodyRange(lo)
    Call DumpInSheet(Rest(Application.Run(DataGetterFunctionName)), lo.Range(1, 1).Offset(1, 0))
    Let lo.ShowAutoFilter = True
    Call lo.AutoFilter.ShowAllData
    
    If WorksheetPassword <> vbNullString Then Call ProtectWorksheet(lo.Parent)
End Sub

