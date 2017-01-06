VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CommandLineHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Class CommandLineHandler.
Dim TaskID As Long
Dim hProc As Long
Dim lExitCode As Long
Dim ACCESS_TYPE As Integer, STILL_ACTIVE As Integer
Dim program As String
Dim thisSystem As FraSys
Dim TESTMODE As Boolean

'NEXT STEP: Move all the command line logic here.
'CHECK: That still works.

'TEST: Minimal test would be a built-in SIMS report and check the parameters run (or error code?).



'States: Initialising, initialised, running (passes to command line), run (should have access to an error code? How?).


Function Construct_CMD_CallString(path As String, ReportName As String, save_destination As String, parameter_id As String, parameter_values As String) As String
    Dim parameters As String
    Dim u_name As String, u_pass As String, u_server As String, u_database As String
    
    With thisSystem
        u_name = .USERNAME
        u_pass = .USERPASSWORD
        u_server = .SERVER
        u_database = .DATABASE
    End With
    

    If parameter_id = "" Then
        parameters = ""
    Else
        parameters = """<ReportParameters>" + parameter_id + "<Values>" + parameter_values + "</Values></Parameter></ReportParameters>"""
    End If


    Construct_CMD_CallString = """" + path + """" + " /USER:" + u_name + " /PASSWORD:" + u_pass + " /SERVERNAME:" + u_server + " /DATABASENAME:" + u_database + " /REPORT:" + ReportName + " /OUTPUT:" + save_destination + " /PARAMS:" + parameters
    
    'Should match format '': progdetails = """" + path + """" + " /USER:" + username + " /PASSWORD:" + userpassword + " /SERVERNAME:" + server + " /DATABASENAME:" + database + " /REPORT:" + ReportName + " /OUTPUT:" + save_destination + " /PARAMS:" + date_parameters
    

End Function

Sub Run_Cmd_Reporter(item As QueueItem)
    Dim itemCMDline As String
'   Run command reporter and wait for it to finish
    ACCESS_TYPE = &H400
    STILL_ACTIVE = &H103
'   Application.Cursor = xlWait
  
'   Shell the task
'   windowstyle:
'   0  Hidden
'   1  Normal
'   2  Minimised Icon
'   3  Maximised
    itemCMDline = item.CommandString
    TaskID = Shell(itemCMDline, 1)

'   Get the process handle
    hProc = OpenProcess(ACCESS_TYPE, False, TaskID)

    If Err <> 0 Then
        MsgBox "Cannot start " & itemCMDline, vbCritical, "Error"
        Exit Sub
    End If
    
    
End Sub

Sub Progress()
    Dim time_counter As Integer
    
    
'   Loop continuously
    time_counter = 0
    Do
    
'   Check on the process
    GetExitCodeProcess hProc, lExitCode
    
'   Allow event processing
    DoEvents
    time_counter = time_counter + 1
    Loop While lExitCode = STILL_ACTIVE
    
End Sub



Sub Get_data(report As String)
    'TODO: Support other filetypes than .CSV.
    Dim thepath As String
    thepath = ThisWorkbook.path
    Workbooks.Open Filename:=thepath + "\" + report + ".csv"
    Sheets(report).Activate
    Sheets(report).Move Before:=Workbooks(MASTER_WORKBOOK_NAME).Sheets(1)
End Sub





Private Sub Class_Initialize()
    'Until Initialise_WithSystemRef is run, just acts as a mock/stub (?).
    Set thisSystem = New FraSys
    TESTMODE = True

End Sub
Sub Initialise_WithSystemRef(sys As FraSys)
    Set thisSystem = Nothing 'clear original reference.
    Set thisSystem = sys
    TESTMODE = False
End Sub
