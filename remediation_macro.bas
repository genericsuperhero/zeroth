Attribute VB_Name = "Module1"
Sub remediation()
'
' remediation Macro
'
' Keyboard Shortcut: Ctrl+o
'
    Rows("1:7").Select
    Selection.Delete Shift:=xlUp
    Columns("B:B").Select
    Selection.EntireColumn.Hidden = True
    Columns("E:E").Select
    Selection.EntireColumn.Hidden = True
    Columns("F:F").Select
    Selection.EntireColumn.Hidden = True
    Columns("H:H").Select
    Selection.EntireColumn.Hidden = True
    Columns("J:R").Select
    Selection.EntireColumn.Hidden = True
    Columns("V:W").Select
    Selection.EntireColumn.Hidden = True
     Columns("B:B").Select
    Application.WindowState = xlMinimized
    Application.WindowState = xlNormal
    Selection.Delete Shift:=xlToLeft
    Columns("D:D").Select
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Columns("E:E").Select
    Selection.Delete Shift:=xlToLeft
    Columns("F:N").Select
    Selection.Delete Shift:=xlToLeft
    Columns("I:J").Select
    Selection.Delete Shift:=xlToLeft
    Range("K1").Select
    ActiveCell.FormulaR1C1 = "Notes"
    Range("L1").Select
    ActiveCell.FormulaR1C1 = "Rescource"
    Range("M1").Select
    ActiveCell.FormulaR1C1 = "Target Date"
    Rows("1:1").Select
    Selection.Font.Bold = True
    Range("K2").Select
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "Vuln Type"
    Range("A1").Select
    Selection.Font.Bold = True
    Range("A2").Select
End Sub
