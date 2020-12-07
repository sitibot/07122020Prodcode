Option Explicit


Sub vlookup(filepath As String)
'
' vlookup Macro
'

'
Dim lastrow As Long
Range("I1").Select
    ActiveCell.FormulaR1C1 = "Amount"
    

lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    ActiveCell.FormulaR1C1 = "Amount"
    Range("I2").Select
    Application.CutCopyMode = False
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = _
        "=VLOOKUP(RC[-7]," & filepath & "]Sheet1'!C1:C3,3,FALSE)"
    Selection.AutoFill Destination:=Range("I2:I" & lastrow)
    Range("I2:I" & lastrow).Select
End Sub

