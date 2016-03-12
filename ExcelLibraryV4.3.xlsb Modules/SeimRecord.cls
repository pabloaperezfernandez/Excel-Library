VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SeimRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Option Base 1

' Private instance variable to hold the trade database
Private RangeDb As Range
Public UnderlyingWorkbook As Workbook

' Custom initializer. The first row is ignored because it is expected to be the
' header row
Public Sub InitializeWithRange(ARange As Range)
    Set RangeDb = ARange.Resize(1, Me.GetNumberOfColumns)
End Sub

' Getters
' This returns the range holding the asset's value. This could be useful when trying to get the address
' of the range.
Property Get GetRange() As Range
    Set GetRange = RangeDb
End Property

Property Get GetHeaders() As Variant
    Let GetHeaders = Array("DateTime", "User", "Type", "Action", "SourceType", "SourceId", "Appliance", "Xml", "SourceFilename", "ComplianceComments", "PrimaryKey")
End Property

Property Get GetNumberOfColumns() As Integer
    Let GetNumberOfColumns = GetArrayLength(GetHeaders())
End Property

Property Get GetDateTime() As Range
    Set GetDateTime = Me.GetRange.Cells(1)
End Property

Property Get GetUser() As Range
    Set GetUser = Me.GetRange.Cells(2)
End Property

Property Get GetType() As Range
    Set GetType = Me.GetRange.Cells(3)
End Property

Property Get GetAction() As Range
    Set GetAction = Me.GetRange.Cells(4)
End Property

Property Get GetSourceType() As Range
    Set GetSourceType = Me.GetRange.Cells(5)
End Property

Property Get GetSourceId() As Range
    Set GetSourceId = Me.GetRange.Cells(6)
End Property

Property Get GetAppliance() As Range
    Set GetAppliance = Me.GetRange.Cells(7)
End Property

Property Get GetXml() As Range
    Set GetXml = Me.GetRange.Cells(8)
End Property

Property Get GetSourceFileName() As Range
    Set GetSourceFileName = Me.GetRange.Cells(9)
End Property

Property Get GetComplianceComments() As Range
    Set GetComplianceComments = Me.GetRange.Cells(10)
End Property

Property Get GetPrimaryKey() As Range
    Set GetPrimaryKey = Me.GetRange.Cells(11)
End Property

Property Let SetDateTime(Arg As Date)
    Let Me.GetDateTime.Value2 = Arg
End Property

Property Let SetUser(Arg As String)
    Let Me.GetUser.Value2 = Arg
End Property

Property Let SetType(Arg As String)
    Let Me.GetType.Value2 = Arg
End Property

Property Let SetAction(Arg As String)
    Let Me.GetAction.Value2 = Arg
End Property

Property Let SetSourceType(Arg As String)
    Let Me.GetSourceType.Value2 = Arg
End Property

Property Let SetSourceId(Arg As String)
    Let Me.GetSourceId.Value2 = Arg
End Property

Property Let SetAppliance(Arg As String)
    Let Me.GetAppliance.Value2 = Arg
End Property

Property Let LetXml(Arg As String)
    Let Me.GetXml.Value2 = Arg
End Property

Property Let LetSourceFileName(Arg As String)
    Let Me.GetSourceFileName.Value2 = Arg
End Property

Property Let LetComplianceComments(Arg As String)
    Let Me.GetComplianceComments.Value2 = Arg
End Property
