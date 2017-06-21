Attribute VB_Name = "Quickies"
Option Explicit

Sub DateStamp()
Attribute DateStamp.VB_ProcData.VB_Invoke_Func = " \n14"
    On Error Resume Next
    Selection = Date
End Sub

Sub TimeStamp()
    On Error Resume Next
    Selection = Time
End Sub

Sub DateTimeStamp()
    On Error Resume Next
    Selection = Now
End Sub
Sub PasteAsValues()
    Dim rng As Range
    On Error Resume Next
    Set rng = Selection
    rng = rng.Value
    
    'Duane's code
'    Selection.Copy
'    Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
'        False, Transpose:=False
End Sub
