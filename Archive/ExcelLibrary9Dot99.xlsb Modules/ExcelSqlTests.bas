Attribute VB_Name = "ExcelSqlTests"
' PURPOSE OF MODULE
'
' The purpose of this module is to test functionality from the ExcelSql module
' Some of them, like GetDividendCashTest3() below require that you have workbook
' PmsDbTables-1Dot0.xlsb open.  This can be found in the TestingDataFiles folder
' of the folder containing this file.

Option Explicit
Option Base 1

' This one has been tested against the results of the DB
Public Function GetDividendCashTest1(PmsTableWbk As Workbook) As Dictionary
    Dim lo As ListObject
    Dim r As Long
    Dim NumOper As Long
    Dim ADict As Dictionary
    Dim AcctNum As String
    Dim PmsCode As String
    Dim Amount0 As Double
    Dim Amount1 As Double
    Dim MntCaisse As Double
    Dim CaisseRef As String
    Dim Curncy As String
    Dim NomStat As String
    Dim SettAff As String
    Dim TypeTrn As Integer
    Dim Qte As Double
    Dim DateTrn As Date
    Dim DateSett As Date
    Dim ColArray As Variant

    Set lo = PmsTableWbk.Worksheets("TransactionsLo").ListObjects("TransactionsLo")
    
    Set ADict = New Dictionary
    Let ColArray = Array("NumOper", "PmsCode", "Amount0", "Amount1", "MntCaisse", _
                         "CaisseRef", "Curncy", "NomStat", "SettAff", "TypeTrn", _
                         "Qte", "DateTrn", "DateSett" _
                        )
    Call ADict.Add(Key:="ColPos", _
                   Item:=CreateDictionary(ColArray, NumericalSequence(1, Length(ColArray))) _
                  )

    For r = 1 To lo.ListRows.Count
        Let NumOper = lo.ListColumns("Num_Oper").DataBodyRange(r, 1).Value2
        Let AcctNum = lo.ListColumns("Cpt_no").DataBodyRange(r, 1).Value2
        Let PmsCode = lo.ListColumns("Code_Titre").DataBodyRange(r, 1).Value2
        Let Amount1 = lo.ListColumns("Amount1").DataBodyRange(r, 1).Value2
        Let Amount0 = lo.ListColumns("Amount0").DataBodyRange(r, 1).Value2
        Let MntCaisse = lo.ListColumns("Mnt_Caisse").DataBodyRange(r, 1).Value2
        Let CaisseRef = lo.ListColumns("Caisse_Ref").DataBodyRange(r, 1).Value2
        Let Curncy = lo.ListColumns("Curr").DataBodyRange(r, 1).Value2
        Let NomStat = lo.ListColumns("NomStat").DataBodyRange(r, 1).Value2
        Let SettAff = lo.ListColumns("SettAff").DataBodyRange(r, 1).Value2
        Let TypeTrn = lo.ListColumns("Type_Trn").DataBodyRange(r, 1).Value2
        Let Qte = lo.ListColumns("Units").DataBodyRange(r, 1).Value2
        Let DateTrn = lo.ListColumns("Date_Trn").DataBodyRange(r, 1).Value2
        Let DateSett = lo.ListColumns("Date_Sett").DataBodyRange(r, 1).Value2

        If CBool(SettAff) And _
            MemberQ(Array(6, 9, 16, 17), TypeTrn) And _
            NomStat = "O" And _
            Mid(CaisseRef, 1, 1) = "0" And _
            Mid(CaisseRef, 2, 1) <> "0" And _
            Amount1 > 3 And _
            SettAff Then

            Let ADict.Item(Key:=NumOper) = Array(NumOper, PmsCode, Amount0, Amount1, _
                                                 MntCaisse, CaisseRef, Curncy, NomStat, _
                                                 SettAff, TypeTrn, Qte, DateTrn, DateSett)
        End If
    Next

    Set GetDividendCashTransactions = ADict
End Function

' This one returns the same thing as the DB but works a lot faster than the function about
Public Function GetDividendCashTest2() As Variant
    Dim lo As ListObject
    Dim data As Variant
    Dim wsht As Worksheet
    Dim TheQuery As String

    Set lo = Application.Workbooks("PmsDbTables-1Dot0.xlsb").Worksheets("TransactionsLo").ListObjects("TransactionsLo")
    
    Call ListObjectsModule.AddColumnsToListObject(lo, ToStrings(Array("PickQ")))
    Let lo.ListColumns("PickQ").DataBodyRange.Formula = "=And(Left([@[Caisse_Ref]],1)=""0"", Mid([@[Caisse_Ref]],2,1)<>""0"")"
    Let lo.ListColumns("PickQ").DataBodyRange.Value2 = lo.ListColumns("PickQ").DataBodyRange.Value2
    
    Call lo.Range.AutoFilter
    Call lo.Range.AutoFilter(Field:=Application.Match("Type_Trn", lo.HeaderRowRange, 0), _
                             Criteria1:=Array("6", "9", "16", "17"), _
                             Operator:=xlFilterValues)
    Call lo.Range.AutoFilter(Field:=Application.Match("SettAff", lo.HeaderRowRange, 0), _
                             Criteria1:="TRUE")
    Call lo.Range.AutoFilter(Field:=Application.Match("Amount1", lo.HeaderRowRange, 0), _
                             Criteria1:=">3")
    Call lo.Range.AutoFilter(Field:=Application.Match("PickQ", lo.HeaderRowRange, 0), _
                             Criteria1:="TRUE")
    Call lo.Range.AutoFilter(Field:=Application.Match("NomStat", lo.HeaderRowRange, 0), _
                             Criteria1:="O")
                             
    Let GetDividendCashTets = lo.Range.SpecialCells(xlCellTypeVisible).Copy
    
    Set wsht = ThisWorkbook.Worksheets.Add
    
    Call wsht.Range("A1").PasteSpecial(xlPasteValuesAndNumberFormats)
    Call wsht.Cells(1, lo.ListColumns("PickQ").Index).EntireColumn.Delete
    
    Call lo.Range.AutoFilter
    Call lo.ListColumns("PickQ").Range.EntireColumn.Delete
    
    Let TheQuery = _
        "SELECT [Cpt_No], SUM(Amount1) FROM [" & wsht.Name & "$] GROUP by [Cpt_No];"
    Let data = SelectUsingSql(TheQuery, ThisWorkbook.Path & "\" & ThisWorkbook.Name)
    
    Call DumpInSheet(data, wsht.Range("AS1"))
End Function

Public Sub GetDividendCashTest3()
    Dim wbk As Workbook
    Dim data As Variant
    Dim TheQuery As String

    Set wbk = Application.Workbooks("PmsDbTables-1Dot0.xlsb")

    Let TheQuery = _
        "SELECT COUNT(*) FROM [TransactionsLo$];"
    Let data = SelectUsingSql(TheQuery, wbk.Path & "\" & wbk.Name)
    Debug.Print TheQuery
    PrintArray data
    Debug.Print
    
    Let TheQuery = _
        "SELECT COUNT(*) FROM [TransactionsLo$]" & vbCrLf & _
        "WHERE Left([Caisse_Ref],1)='0';"
    Let data = SelectUsingSql(TheQuery, wbk.Path & "\" & wbk.Name)
    Debug.Print TheQuery
    PrintArray data
    Debug.Print

    Let TheQuery = _
        "SELECT COUNT(*) FROM [TransactionsLo$]" & vbCrLf & _
        "WHERE MID([Caisse_Ref],2,1)<>'0';"
    Let data = SelectUsingSql(TheQuery, wbk.Path & "\" & wbk.Name)
    Debug.Print TheQuery
    PrintArray data
    Debug.Print
    
    Let TheQuery = _
        "SELECT COUNT(*) FROM [TransactionsLo$]" & vbCrLf & _
        "WHERE [SettAff]"
    Let data = SelectUsingSql(TheQuery, wbk.Path & "\" & wbk.Name)
    Debug.Print TheQuery
    PrintArray data
    Debug.Print
    
    Let TheQuery = _
        "SELECT COUNT(*) FROM [TransactionsLo$]" & vbCrLf & _
        "WHERE [NomStat] = 'O';"
    Let data = SelectUsingSql(TheQuery, wbk.Path & "\" & wbk.Name)
    Debug.Print TheQuery
    PrintArray data
    Debug.Print

    Let TheQuery = _
        "SELECT COUNT(*) FROM [TransactionsLo$]" & vbCrLf & _
        "WHERE [Amount1]>3;"
    Let data = SelectUsingSql(TheQuery, wbk.Path & "\" & wbk.Name)
    Debug.Print TheQuery
    PrintArray data
    Debug.Print
    
    Let TheQuery = _
        "SELECT COUNT(*) FROM [TransactionsLo$]" & vbCrLf & _
        "WHERE Year([Date_Sett])=2018;"
    Let data = SelectUsingSql(TheQuery, wbk.Path & "\" & wbk.Name)
    Debug.Print TheQuery
    PrintArray data
    Debug.Print
    
    Let TheQuery = _
        "SELECT COUNT(*) FROM [TransactionsLo$]" & vbCrLf & _
        "WHERE [Date_Sett]>=#1/1/2018# and [Date_Sett]<=#12/31/2018#;"
    Let data = SelectUsingSql(TheQuery, wbk.Path & "\" & wbk.Name)
    Debug.Print TheQuery
    PrintArray data
    Debug.Print

    Let TheQuery = _
        "SELECT COUNT(*) FROM [TransactionsLo$]" & vbCrLf & _
        "WHERE Left([Caisse_Ref],1)='0' AND MID([Caisse_Ref],2,1)<>'0' AND " & vbCrLf & _
        "      [SettAff] AND [NomStat] = 'O' AND " & vbCrLf & _
        "      Year([Date_Sett])=2018 AND [Amount1]>3;"
    Let data = SelectUsingSql(TheQuery, wbk.Path & "\" & wbk.Name)
    Debug.Print TheQuery
    PrintArray data
    Debug.Print
End Sub
