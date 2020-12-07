Option Explicit

Sub Changetheordinal()
'
' Macro1 Macro
'

'
    Columns("A:A").Select
    Selection.Cut
    Columns("C:C").Select
    ActiveSheet.Paste
    Columns("A:A").Select
    Selection.Delete Shift:=xlToLeft
End Sub
