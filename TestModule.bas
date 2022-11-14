Attribute VB_Name = "TestModule"
Option Explicit

Private Sub delRowTest()
Attribute delRowTest.VB_ProcData.VB_Invoke_Func = " \n14"
'
' delRowTest Macro
'

'
    Rows("8:8").Select
    Selection.Delete Shift:=xlUp
    Rows("7:52").Select
    Selection.Delete Shift:=xlUp
End Sub
