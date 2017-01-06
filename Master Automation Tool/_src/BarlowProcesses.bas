Attribute VB_Name = "BarlowProcesses"
Option Explicit


Public Sub CopyStaffWithNoIncidentsToDashboard()

    With Worksheets("Expected Staff")
        .Activate
        .Range("$A$1:$C$80").AutoFilter Field:=3, Criteria1:="0"
        .Range("A2:a80").Select
        Selection.Copy
    End With


    Sheets("Dashboard").Select
    Range("A75").Select
    ActiveSheet.Paste
    Selection.ClearFormats

End Sub
