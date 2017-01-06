Attribute VB_Name = "Module2"
Sub Per_cent_col()

'   Adds extra column to Report2. Calculates % absence, assuming 150 sessions

    Sheets("Report2").Select
    Range("P1").Select
    ActiveCell.FormulaR1C1 = "%"
    Range("P2").Select
    ActiveCell.FormulaR1C1 = "=IFERROR((150-RC[-1])/150,"""")"
    Range("P2").Select
    Selection.AutoFill Destination:=Range("P2:P10000")
    Range("P2:P10000").Select
    Columns("P:P").Select
    Selection.Style = "Percent"
    
End Sub

Sub Absence_per_cent()

'   Adds extra column to Report2. Calculates % absence (absent/(present+absent))

    Sheets("Report2").Select
    Range("M1").Select
    ActiveCell.FormulaR1C1 = "% absent"
    Range("M2").Select
    ActiveCell.FormulaR1C1 = "=iferror(RC[-1]/(RC[-2]+RC[-1]),0)"
    Range("M2").Select
    Selection.AutoFill Destination:=Range("M2:M10000")
    Range("M2:M10000").Select
    Columns("M:M").Select
    Selection.NumberFormat = "0.00%"
    
End Sub

Sub Beh_weeks()

'   Adds extra column to report of behavioural incidents showing the week in which they occured
'   Requires adapting before use

 Sheets("report3").Select
    Range("Q1").Select
    ActiveCell.FormulaR1C1 = "Weeks"
    Range("T4").Select
    ActiveCell.FormulaR1C1 = "=MIN(C[-6])"
    Range("Q2").Select
    ActiveCell.FormulaR1C1 = "=INT((RC[-3]-R4C20)/7)+1"
Dim r As Long
    r = Cells(Rows.Count, "A").End(xlUp).Row
    Range("q2").Select
    Selection.AutoFill Destination:=Range("q2:q" & r), Type:=xlFillDefault

    Columns("Q:Q").Select
    ActiveWorkbook.Worksheets("Report3").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Report3").Sort.SortFields.Add Key:=Range("Q2:Q80") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Report3").Sort
        .SetRange Range("A1:Q80")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

End Sub



