Attribute VB_Name = "Logging"

Option Explicit


    'ID  DateStamp   Intended Outcome    Log Value

    'TODO: Create message log function.
    
    
Sub TestSystemLog()
    Call WriteLineToSystemLog("Test: System Log Test", "---", "---")
End Sub
    
Sub TestMessageLog()
    Call WriteLineToSystemLog("Test: Message Log Test", "---", "TODO: Write to correct log.")
End Sub
        
    

Public Sub WriteLineToMessageLog()
    'Our Name    SIMS Report Name    Success?    ErrorVal    Last Success    Last Time Taken Average Time Taken  Last Fail
End Sub



Public Sub WriteLineToSystemLog(LogValue As String, IntendedOutcome As String, ActualOutcome As String)
    Dim ws As Worksheet, y As Integer
    Set ws = ThisWorkbook.Worksheets("SystemLog")
    y = GetLastLineOnSheet(ws)
    
    'Write values.
    ws.Range("A1").Offset(y, 0).Value = y
    ws.Range("A1").Offset(y, 1).Value = CStr(TimeStampNow)
    ws.Range("A1").Offset(y, 1).NumberFormat = "yyyy-mm-dd:hhmm"
    ws.Range("A1").Offset(y, 2).Value = LogValue
    ws.Range("A1").Offset(y, 3).Value = IntendedOutcome
    ws.Range("A1").Offset(y, 4).Value = ActualOutcome
    
End Sub


Public Function GetLastLineOnSheet(ws As Worksheet) As Integer
    Dim y As Integer
    
    y = 0
    While ws.Range("A1").Offset(y, 0).Value & "" <> ""
        
        y = y + 1
    Wend
    GetLastLineOnSheet = y
End Function
