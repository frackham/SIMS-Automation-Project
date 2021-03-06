Attribute VB_Name = "Logging"

Option Explicit


    'ID  DateStamp   Intended Outcome    Log Value
    'TODO: Create message log function.
    'TODO: Add logging levels.
    

Sub DemoSystemLog()
    Call WriteLineToSystemLog("Test: System Log Test", "---", "---")
End Sub
    
Sub DemoMessageLog()
    Call WriteLineToSystemLog("Test: Message Log Test", "---", "TODO: Write to correct log.")
End Sub
        
    

Public Sub WriteLineToMessageLog(LogValue As String, IntendedOutcome As String, ActualOutcome As String)
    'ID   Datestamp   Our Name    SIMS Report Name    Success?    ErrorVal    Last Success    Last Time Taken Average Time Taken  Last Fail
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
    ws.Calculate 'TODO: Trying to make sure that system log always writes.
    
End Sub


Public Function GetLastLineOnSheet(ws As Worksheet) As Integer
    Dim y As Integer
    
    y = 0
    While ws.Range("A1").Offset(y, 0).Value & "" <> ""
        
        y = y + 1
    Wend
    GetLastLineOnSheet = y
End Function

Public Function GetLastLog_System() As String
    Dim ws As Worksheet
    Dim y As Integer, sArr() As String, i As Integer
    Set ws = ThisWorkbook.Worksheets("SystemLog")
    y = GetLastLineOnSheet(ws)
    ReDim sArr(5)
    For i = 0 To 5
        sArr(i) = ws.Range("A1").Offset(y, i)
    Next i
    
    GetLastLog_System = Join(sArr, ",")
    Set ws = Nothing
End Function

Public Function GetLastLog_Message() As String
'TODO: Write.
End Function
