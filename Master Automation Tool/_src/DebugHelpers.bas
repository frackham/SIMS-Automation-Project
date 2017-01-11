Attribute VB_Name = "DebugHelpers"

Option Explicit
'None of the functions here should be run by code - only manually. They are to support development.

Private Sub UnhideAllSheets()
    Dim ws As Worksheet
    
    For Each ws In ThisWorkbook.Worksheets
        ws.Visible = xlSheetVisible
    Next ws
End Sub

Private Sub ShowChartObjectX()
    Dim n As Integer
    Dim originalColour As Variant
    n = 1
    
    With ThisWorkbook.Worksheets("Dashboard")
        .Activate
        .ChartObjects(n).Activate
        originalColour = .ChartObjects(n).Border.Color
        .ChartObjects(n).Border.Color = RGB(255, 0, 0)
        Application.Wait (10)
        .ChartObjects(n).Border.Color = originalColour
    End With
End Sub

Private Sub GetChartObjectValueOfSelectedChart()
    Dim co As ChartObject
    For Each co In ThisWorkbook.Worksheets("Dashboard").ChartObjects
        If co.Name = ActiveChart.Name Then
            MsgBox (co.Name)
        End If
    Next co
End Sub

Private Sub NameSelectedChart()
    Dim sName As String
    On Error GoTo ERRORNAMING
    sName = InputBox("Input new name of chart")
    ActiveChart.Parent.Name = sName
    Exit Sub
    
ERRORNAMING:
    MsgBox ("Name not accepted, please try again." & vbCrLf & Err.Description & vbCrLf & Err.Number)
    On Error GoTo 0
End Sub

Private Sub CheckChartObjectsNames()
    Dim n As Integer, co As ChartObject
    For Each co In ThisWorkbook.Worksheets("Dashboard").ChartObjects
        n = n + 1
        MsgBox (n & "  |  Name:" & co.Name & "   |  ParentName:" & co.Parent.Name)
    Next co
End Sub

