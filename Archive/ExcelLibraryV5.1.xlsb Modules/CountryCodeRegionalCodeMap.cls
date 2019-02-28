VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CountryCodeRegionalCodeMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Base 1
Option Explicit

' Private instance variable to hold the database
Private MapDb As Variant
' Private instance variable to hold the lookup dictionary
Private RowPositionDict As Dictionary
Private EnglishNameDict As Dictionary

' Custom initializer. This sets MapDb as a 2D array
Public Sub InitializeWithRange(ARange As Range)
    ' Set reference to the underlying range holding the data
    Let MapDb = ARange.Offset(1, 0).Resize(ARange.Rows.Count - 1, ARange.Columns.Count).Value2
    
    ' Instantiate and initialize row indexing dictionary for fast searches
    Call InitializeRowIndexDictionary
    Call InitializeEnglishNameDict
End Sub


' Initializes the object with the data in the file with the given filename.
Public Sub InitializeWithFile(AFileName As String)
    Dim WorkbookObject As Workbook
    Dim TheRange As Range
    
    ' Open workbook containing map
    Set WorkbookObject = Application.Workbooks.Open(AFileName)
    
    ' Extract range holding the data
    Set TheRange = WorkbookObject.Worksheets(1).Range("A1").CurrentRegion
    
    Call InitializeWithRange(TheRange)
    
    ' Close workbook holding the map
    Call WorkbookObject.Close(SaveChanges:=False)
    
    Call InitializeRowIndexDictionary
End Sub

' Initializes the object with the data in the file with the given filename.
Public Sub InitializeWithMySql(DbServerAddress As String, DbName As String, DbUserName As String, DbPassword As String, TableName As String)
    Dim wsht As Worksheet
    Dim TheResults As Variant
    
    ' Open workbook containing map
    Set wsht = Application.Workbooks.Add.Worksheets(1)
    Call RemoveAllOtherWorksheets(wsht)
    
    ' Get table from DB
    Let TheResults = ConnectAndSelect("SELECT * FROM `" & DbName & "`.`" & TableName & "`;", DbName, DbServerAddress, DbUserName, DbPassword)
    
    ' Dump data in target worksheet
    Call DumpInTempPositionWithoutFirstClearing(TheResults, wsht.Range("A1"))
    
    ' Initialize with the range
    Call InitializeWithRange(wsht.Range("A1").CurrentRegion)
    
    Call InitializeRowIndexDictionary
End Sub

' This function initializes a dictionary indexing the row positions SEDOLs in the private file
Public Sub InitializeRowIndexDictionary()
    Dim i As Integer
    
    ' Instantiate dictionary
    Set RowPositionDict = New Dictionary
    
    ' Add the SEDOL row Indices to the dictionary
    For i = 1 To Me.GetDbLength
        ' Insert this sedol's row position in the dictionary
        Me.GetRowIndicesDb.Add Me.GetIthCountryCode(i), Me.GetIthRegionalCode(i)
    Next i
End Sub

Public Sub InitializeEnglishNameDict()
    Dim i As Integer
    
    ' Instantiate dictionary
    Set EnglishNameDict = New Dictionary
    
    ' Add the SEDOL row Indices to the dictionary
    For i = 1 To Me.GetDbLength
        ' Insert this sedol's row position in the dictionary
        Call Me.GetShortEnglishNameDb.Add(Me.GetIthCountryCode(i), Me.GetIthShortEnglishName(i))
    Next i
End Sub

' Getters
' Returns the array holding the data
Property Get GetMapDb() As Variant
    Let GetMapDb = MapDb
End Property

Property Get GetShortEnglishNameDb() As Dictionary
    Set GetShortEnglishNameDb = EnglishNameDict
End Property

Property Get GetIthCountryCode(i As Integer) As String
    Let GetIthCountryCode = Me.GetMapDb(i, 2)
End Property

Property Get GetIthRegionalCode(i As Integer) As String
    Let GetIthRegionalCode = Me.GetMapDb(i, 4)
End Property

Property Get GetIthShortEnglishName(i As Integer) As String
    Let GetIthShortEnglishName = Me.GetMapDb(i, 5)
End Property

Property Get GetRowIndicesDb() As Dictionary
    Set GetRowIndicesDb = RowPositionDict
End Property

Property Get GetDbLength() As Integer
    Let GetDbLength = UBound(GetMapDb, 1)
End Property

Property Get GetRegionalCode(CountryIsoCode As String) As String
    Let GetRegionalCode = Me.GetRowIndicesDb.Item(CountryIsoCode)
End Property

Property Get GetShortEnglishName(CountryIsoCode As String) As String
    Let GetShortEnglishName = Me.GetShortEnglishNameDb.Item(CountryIsoCode)
End Property





