VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

'Code for events (kept separate, so that they can be run manually).
Private Sub ImportModulesIfMoreRecent()
    Dim bImported As Boolean
    'Sheets("Dashboard").Select
    'Sheets("Dashboard").Worksheet_Activate
    
    'Check if any of the _src files have updated since the last time
    Call WriteLineToSystemLog("ImportModulesIfMoreRecent", "-", "Start")
    bImported = False
    If CheckWhetherToImportSRCModules = True Then
        'xbreaker = 0 / 0
        On Error GoTo error
        ImportModules
        bImported = True
    End If
error:
    On Error GoTo 0
    If bImported = True Then
        Call WriteLineToSystemLog("ImportModulesIfMoreRecent", "IMPORTED.", "")
    Else
        Call WriteLineToSystemLog("ImportModulesIfMoreRecent", "DID NOT IMPORT.", "")
    End If
    Call WriteLineToSystemLog("ImportModulesIfMoreRecent", "-", "End")
End Sub
Private Sub ExportAndUpdateTimestamp()
    Call WriteLineToSystemLog("ExportAndUpdateTimestamp", "-", "Start")
    ExportModules
    TimeStampThisLastModified_Update
    Call WriteLineToSystemLog("ExportAndUpdateTimestamp", "-", "End")
End Sub

'Events
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Call WriteLineToSystemLog("Event: BeforeClose.", "-", "Start")
    If ActiveWorkbook.Saved = True Then
        ExportAndUpdateTimestamp
    End If
    Call WriteLineToSystemLog("*Event: BeforeClose.", "-", "End")
End Sub
Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    On Error GoTo failsafesave
    Call WriteLineToSystemLog("Event: BeforeSave.", "-", "Start")
    ExportAndUpdateTimestamp
    Call WriteLineToSystemLog("Event: BeforeSave.", "-", "End")
Exit Sub
failsafesave:
'Should still save the workbook in the event of a failure.
Application.DisplayAlerts = False
Application.EnableEvents = False
ThisWorkbook.Save
End Sub
Private Sub Workbook_Open()
    Call WriteLineToSystemLog("*Event: Open.", "-", "Start")
    ImportModulesIfMoreRecent
    Call WriteLineToSystemLog("Event: Open.", "-", "End")
End Sub

