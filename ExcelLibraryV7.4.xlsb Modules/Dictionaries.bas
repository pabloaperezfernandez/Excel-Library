Attribute VB_Name = "Dictionaries"
Option Explicit
Option Base 1

' The purpose of this function is to return a dictionary indexing the items in paramater
' TheItems using the elements of TheKeys.  Both parameters must be 1D arrays.  2D arrays
' are disallowed unless they are 1D arrays of arrays.  TheKeys and TheItems must be
' non-empty and have the same length.  No element of TheKeys can be an array because
' arrays are not allowed as keys in dictionaries.
'
' If the parameters are inconsistent, the function returns Nothing
Public Function CreateDictionary(TheKeys As Variant, TheItems As Variant) As Dictionary
    Dim aDict As Dictionary
    Dim KeysIndex As Long
    Dim ItemsIndex As Long
    
    ' Check inputs consistency
    
    ' Exit if either argument is not an array
    If Not IsArray(TheKeys) Or Not IsArray(TheItems) Then
        Set CreateDictionary = Nothing
        Exit Function
    End If
    
    ' Exit if either argument is an Empty1D array
    If EmptyArrayQ(TheKeys) Or EmptyArrayQ(TheItems) Then
        Set CreateDictionary = Nothing
        Exit Function
    End If
    
    ' Exit if either argument has dimension other than 1
    If NumberOfDimensions(TheKeys) <> 1 Or NumberOfDimensions(TheItems) <> 1 Then
        Set CreateDictionary = Nothing
        Exit Function
    End If
    
    ' Exit if the arguments don't have the same length
    If Length(TheKeys) <> Length(TheItems) Then
        Set CreateDictionary = Nothing
        Exit Function
    End If
    
    ' Exit if any element in TheKeys is an array since arrays cannot be used as keys
    For KeysIndex = LBound(TheKeys) To UBound(TheKeys)
        If IsArray(TheKeys(KeysIndex)) Then
        Set CreateDictionary = Nothing
            Exit Function
        End If
    Next KeysIndex
    
    ' If the code gets here, the arguments are consistent with expectations
    Set aDict = New Dictionary
    
    Let ItemsIndex = LBound(TheItems)
    For KeysIndex = LBound(TheKeys) To UBound(TheKeys)
        If Not aDict.Exists(Key:=TheKeys(KeysIndex)) Then
            Call aDict.Add(Key:=TheKeys(KeysIndex), Item:=TheItems(ItemsIndex))
        End If
        
        Let ItemsIndex = ItemsIndex + 1
    Next KeysIndex
    
    Set CreateDictionary = aDict
End Function

' DESCRIPTION
' This function translates the elements of a 1D or 2D arrays using a dictionary.
' We impose no restriction on the dictionary's elements.  Hence, an arg satisfying
' Predicates.AtomicArrayQ or Predicates.AtomicTableQ may fail to after passing through this function.
'
' PARAMETERS
' 1. AnAtomicArrayOrTable       - An array satisfying Predicates.AtomicArrayQ or Predicates.AtomicTableQ
' 2. aDict                      - A dictionary
'
' RETURNED VALUE
' AnAtomicArrayOrTable after applying aDict to each of its elements.  If an element in AnAtomicArrayOrTable
' is not found in the dictionary, that element is not translated.
Public Function TranslateUsingDictionary(AnAtomicArrayOrTable As Variant, _
                                         aDict As Dictionary, _
                                         Optional ParameterCheckQ As Boolean = False) As Variant
    Dim ReturnArray() As Variant
    Dim r As Long
    Dim c As Long

    If aDict.Count = 0 Then
        Let TranslateUsingDictionary = AnAtomicArrayOrTable
        Exit Function
    End If
    
    If Not IsArray(AnAtomicArrayOrTable) Then
        Let TranslateUsingDictionary = Null
        Exit Function
    End If
    
    If Not DimensionedQ(AnAtomicArrayOrTable) Then
        Let TranslateUsingDictionary = AnAtomicArrayOrTable
        Exit Function
    End If
    
    If EmptyArrayQ(AnAtomicArrayOrTable) Then
        Let TranslateUsingDictionary = AnAtomicArrayOrTable
        Exit Function
    End If
    
    If ParameterCheckQ Then
        If Not (AtomicArrayQ(AnAtomicArrayOrTable) Or AtomicTableQ(AnAtomicArrayOrTable)) Then
            Let TranslateUsingDictionary = Null
            Exit Function
        End If
    End If
    
    If NumberOfDimensions(AnAtomicArrayOrTable) = 1 Then
        ReDim ReturnArray(LBound(AnAtomicArrayOrTable, 1) To UBound(AnAtomicArrayOrTable, 1))
        
        For r = LBound(AnAtomicArrayOrTable, 1) To UBound(AnAtomicArrayOrTable, 1)
            If aDict.Exists(Key:=AnAtomicArrayOrTable(r)) Then
                Let ReturnArray(r) = aDict.Item(Key:=AnAtomicArrayOrTable(r))
            Else
                Let ReturnArray(r) = AnAtomicArrayOrTable(r)
            End If
        Next
    Else
        ReDim ReturnArray(LBound(AnAtomicArrayOrTable, 1) To UBound(AnAtomicArrayOrTable, 1), _
                          LBound(AnAtomicArrayOrTable, 2) To UBound(AnAtomicArrayOrTable, 2))
                          
        For r = LBound(AnAtomicArrayOrTable, 1) To UBound(AnAtomicArrayOrTable, 1)
            For c = LBound(AnAtomicArrayOrTable, 2) To UBound(AnAtomicArrayOrTable, 2)
                If aDict.Exists(Key:=AnAtomicArrayOrTable(r, c)) Then
                    Let ReturnArray(r, c) = aDict.Item(Key:=AnAtomicArrayOrTable(r, c))
                Else
                    Let ReturnArray(r, c) = AnAtomicArrayOrTable(r, c)
                End If
            Next
        Next
    End If
        
    Let TranslateUsingDictionary = ReturnArray
End Function

'***HERE
' DESCRIPTION
' The purpose of this function is to get data on the given argument using a sequence of dictionaries.
' The idea is to get the information from the first dictionary or recurse on the rest of the dictionaries
' if the info is missing in the first.
'
' The dictionaries are structured using items that are either arrays, or dictionaries themselves.
' The function finds the key in the first dictionary.  If not found, it returns Null. If found, it looks
' at the corresponding item.  If the item's index k corresponding to this dictionary is a number --with 1
' representing the first position--, we return the kth element in the the key's item.  If the item's index
' is a string ColumnName, we have another dictionary and return the item corresponding to index ColumnName.
' If either method results in Null (e.g. case of item being an array) or ColumnName does not exists in the
' item's dictionary, we then recurse on the remaining dictionaries.
'
' PARAMETERS
' 1. Needles            - The term(s) being seached for in the dictonaries
' 2. HayStacks          - A dictionary or array of dictionaries, all of whose items are either arrays or dictionaries.
'                         If the columns indices are arrays, they must all have the same legth.  If the items are
'                         dictionaries they must all the same keys. If the value corresponding to a needle is missing,
'                         the corresponding item must be Null or equal to one of the Optional MissingValueFlags.
'
'                         One essential feature is that of NextKey.  If key has Null value in its item, then we need to
'                         the key to use for the next dictionary.  This is expected in either the 0th position when using
'                         arrays or in the item with key = "NextKey" when using dictionaries.
' 3. ColumnIndices      - These are either integers larger than or equal to 1 or strings. ColumnIndicesArray must have
'                         the same length as HayStackArray (e.g. the same number of indices as dictionaries
'
' RETURNED VALUE
' Returns either the value(s) corresponding to the needle(s) or Null(s) when not found in the sequence of dictionaries
Public Function GetValueFromDictionaries(Needles As Variant, _
                                         HayStacks As Variant, _
                                         ColumnIndices As Variant, _
                                         Optional MissingValueFlags As Variant) As Variant
    Dim ResultsArray() As Variant
    Dim r As Long
                                         
    ' Check consistency of parameter TheNeedleOrNeedles
    If Not (StringQ(Needles) Or StringArrayQ(Needles)) Then
        Let GetValueFromDictionaries = Null
        Exit Function
    End If
    
    ' Check consistency of parameter HayStackArray
    If Not (DictionaryQ(HayStacks) Or Dictionary1DArrayQ(HayStacks)) Then
        Let GetValueFromDictionaries = Null
        Exit Function
    End If
    
    ' Exit if ColumnIndicesArray is neither an atom (string or integer) nor an array of these
    If Not (IntergerLongOrStringQ(ColumnIndices) Or IntergerLongOrString1DArrayQ(ColumnIndices)) Then
        Let GetValueFromDictionaries = Null
        Exit Function
    End If
    
    ' Exit with null if HayStacks, and ColumnIndices don't have the same length
    If Length(HayStacks) <> Length(ColumnIndices) Then
        Let GetValueFromDictionaries = Null
        Exit Function
    End If
    
    ' If Needles is an atom and HayStacks is just a dictionary, then process it
    If DictionaryQ(HayStacks) And Not IsArray(Needles) Then
        If HayStacks.Exists(Key:=Needles) Then
            Let GetValueFromDictionaries = HayStacks.Exists(Key:=Needles)
        Else
            Let GetValueFromDictionaries = Null
        End If
        
        Exit Function
    End If
    
    ' If needles is an atom and HayStacks and array, then process it and recurse only if necessary
    If DictionaryQ(HayStacks) And IsArray(Needles) Then
        ReDim ResultsArray(LBound(Needles) To UBound(Needles))
        
        For r = LBound(Needles) To UBound(Needles)
            If HayStacks.Exists(Key:=Needles(r)) Then
                Let ResultsArray(r) = HayStacks.Item(Key:=Needles(r))
            Else
                Let ResultsArray(r) = Null
            End If
        Next
        
        Let GetValueFromDictionaries = ResultsArray
    End If
    
    '***HERE
End Function
