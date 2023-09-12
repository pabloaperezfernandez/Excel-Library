Attribute VB_Name = "Module1"
Option Explicit

Sub Macro2()
Attribute Macro2.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro2 Macro
'

'
    Windows("CDP Roadmap Planner-Catalog_Glenn.xlsx").Activate
    ActiveWorkbook.LockServerFile
End Sub
