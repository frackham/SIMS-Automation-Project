Attribute VB_Name = "Module1"

Option Explicit

Declare Function OpenProcess Lib "kernel32" _
    (ByVal dwDesiredAccess As Long, _
    ByVal bInheritHandle As Long, _
    ByVal dwProcessId As Long) As Long
    Declare Function GetExitCodeProcess Lib "kernel32" _
    (ByVal hProcess As Long, _
    lpExitCode As Long) As Long
'REFACTOR: All of these beore the autopen are global properties/variables


Dim CMD As CommandLineHandler
Dim OurSystem As FraSys
Dim MessageQueue As QueueManager
    
    

Public MASTER_WORKBOOK_NAME As String
Public DEBUG_MODE As Boolean
Public DONT_SEND_EMAIL_MODE  As Boolean
Public skip As String

Sub auto_open(Optional bRanFromFirstOpen As Boolean = True)
    Dim rRange_ReportName As Range, rRange_ParameterID As Range, rRange_ParameterValues As Range
    Dim i As Integer
    Dim path As String, sNameCell As String, sIDCell As String, sValueCell As String
    Dim ReportName As String
    Dim progdetails As String
    Dim SAVE_PATH As String
    Dim skip As String
    
    
    
    
    If OurSystem Is Nothing Then
        Set OurSystem = New FraSys
    End If
    Call OurSystem.InitialiseSystemConfig 'TODO: Load any config required to system singleton here.
    If CMD Is Nothing Then
        Set CMD = New CommandLineHandler
    End If
    If MessageQueue Is Nothing Then
        Set MessageQueue = New QueueManager
    End If
'    Call CMD.Initialise_WithSystemRef(System)
'    Call MessageQueue.Setup(CMD, System)



    MASTER_WORKBOOK_NAME = "Dashboard (barlow modified).xlsm"
    DEBUG_MODE = Sheets("Set up").Range("J12").Value
    DONT_SEND_EMAIL_MODE = Sheets("Set up").Range("J24").Value
    
    
'   Skips rest of procedure if this is not Dashboard file (ie is Output file)
    If Not ActiveWorkbook.Name = MASTER_WORKBOOK_NAME Then
        'Sheets("Drilldown").Activate
        'Sheets("Dashboard").Activate
        Exit Sub
    End If
        
        
    Dim ws As Worksheet, thepath As String
        
'   Skips running at weekend (if task scheduler set to run daily).
'   Also skips if in setupmode and just opened.
    If Sheets("Set up").Range("K8") = 1 Then
        'Operating Mode.
        If Sheets("Set up").Range("J13") = "Yes" And Weekday(Date, 2) > 5 Then
            Application.DisplayAlerts = False
            If DEBUG_MODE = False Then Application.Quit 'Never quit in debug mode.
            Exit Sub
        End If
    ElseIf Sheets("Set up").Range("K8") = 2 Then
        'Set-up Mode.
        For Each ws In ThisWorkbook.Worksheets
            ws.Visible = True
        Next ws
        
        If bRanFromFirstOpen Then
            'Don't run the auto_open [misnomer] function through if you've just opened it AND it's in setup mode.
            Exit Sub
        End If
    End If
    
    
    
    MessageQueue.PopulateFromValues
    
    
'   ***Will resume on error if report is missing - alternatively comment out any reports that you are not using (2)***
'   Load separate .csv reports

    On Error Resume Next
'
'    Call Get_data("Report1")
'    Call Get_data("Report2")
'    Call Get_data("Report3")
'    Call Get_data("Report4")
'    Call Get_data("Report5")
'    Call Get_data("Report6")
'    Call Get_data("Report7")
'    Call Get_data("Report8")
'    Call Get_data("Report9")
'    Call Get_data("Report10")
'    Call Get_data("Report11")
'    Call Get_data("Report12")
'    Call Get_data("Report13")
'    Call Get_data("Report14")
'    Call Get_data("Report15")
'    Call Get_data("Report16")
'    Call Get_data("Report17")
'    Call Get_data("Report18")
'
'   ////////////////////////Add extra tasks here///////////////////////////
    'Call Per_cent_col
    'Call Beh_weeks
    Call CopyStaffWithNoIncidentsToDashboard
    
     
'   ///////////////////////////////////////////////////////////////////////
    
'   Hide sheets, save, send email

    
    If Sheets("Set up").Range("K8") = 1 Then 'tidy up and save if in operating mode. Will quit at the end of the save.
        Call Tidy_up
        Call Save
        
        If Not OurSystem.EMAIL_PATH = "" Then
    '   (Un)comment email method
        Call send_mail_Outlook(SAVE_PATH)
        'Call send_mail_smtp
        End If
    ElseIf DEBUG_MODE Then 'tidy up and save if in debug mode. Will not quit as debug mode will prevent.
        Call Tidy_up
        Call Save
        
        If Not OurSystem.EMAIL_PATH = "" Then
    '   (Un)comment email method
        Call send_mail_Outlook(SAVE_PATH)
        'Call send_mail_smtp
        End If
    End If
    
    End Sub



Sub Tidy_up()

'   ***Will resume on error if report is missing - alternatively comment out any reports that you are not using (3)***
'   Refresh pivot tables and charts
    
'    On Error Resume Next
'
'    Sheets("Report1_pivot").PivotTables("PivotTable1").PivotCache.Refresh
'    Sheets("Report2_pivot").PivotTables("PivotTable2").PivotCache.Refresh
'    Sheets("Report3_pivot").PivotTables("PivotTable3").PivotCache.Refresh
'    Sheets("Report4_pivot").PivotTables("PivotTable4").PivotCache.Refresh
'    Sheets("Report5_pivot").PivotTables("PivotTable5").PivotCache.Refresh
'    Sheets("Report6_pivot").PivotTables("PivotTable6").PivotCache.Refresh
'    Sheets("Report7_pivot").PivotTables("PivotTable7").PivotCache.Refresh
'    Sheets("Report8_pivot").PivotTables("PivotTable8").PivotCache.Refresh
'    Sheets("Report9_pivot").PivotTables("PivotTable9").PivotCache.Refresh
'
'    For n = 10 To 18
'        Sheets("Report" & n & "_pivot").PivotTables("PivotTable" & n).PivotCache.Refresh
'    Next n

'   ***Will resume on error if report is missing - alternatively comment out any reports that you are not using (4)***
'   Delete report and hide pivot tables sheets
    Application.DisplayAlerts = False
    
'    On Error Resume Next
'
'    Sheets("Report1").Delete
'    Sheets("Report2").Delete
'    Sheets("Report3").Delete
'    Sheets("Report4").Delete
'    Sheets("Report5").Delete
'    Sheets("Report6").Delete
'    Sheets("Report7").Delete
'    Sheets("Report8").Delete
'    Sheets("Report9").Delete
'
'    For n = 10 To 18
'        Sheets("Report" & n).Delete
'    Next n
'
'    Sheets("Report1_pivot").Visible = False
'    Sheets("Report2_pivot").Visible = False
'    Sheets("Report3_pivot").Visible = False
'    Sheets("Report4_pivot").Visible = False
'    Sheets("Report5_pivot").Visible = False
'    Sheets("Report6_pivot").Visible = False
'    Sheets("Report7_pivot").Visible = False
'    Sheets("Report8_pivot").Visible = False
'    Sheets("Report9_pivot").Visible = False
'
'    For n = 10 To 18
'        Sheets("Report" & n & "_pivot").Visible = False
'    Next n

' Hide toolbars, formula bars, gridlines. Comment out if not want to use

'    Sheets("Dashboard").Select
'    Application.DisplayFormulaBar = False
'    ActiveWindow.DisplayGridlines = False
'    ActiveWindow.DisplayHeadings = False
'    Application.DisplayFullScreen = True
'    ActiveWindow.Zoom = Sheets("Set up").Range("B12")
'    Range("H8").Select
    
' Add run date
    Sheets("Dashboard").Select
    Range("A2").Select
    ActiveCell.FormulaR1C1 = Now
    Range("AL2").Select
    ActiveCell.FormulaR1C1 = Now
   
End Sub

Sub send_mail_smtp()

' Sends email notification

    Dim CDO_Mail As Object
    Dim CDO_Config As Object
    Dim SMTP_Config As Variant
    Dim strSubject As String
    Dim strfrom As String
    Dim strto As String
    Dim strCc As String
    Dim strBcc As String
    Dim strBody As String
    Dim strsmtp As String
    Dim EMAILPATH As String
    
    strSubject = "SIMS Dashboard"
    strCc = ""
    strBcc = ""
    'Use this if sending as link
    strBody = "The SIMS dashboard has run successfully and is available here:" & vbCrLf & vbCrLf & EMAILPATH
    'Use these if sending as attachment
    'strBody = "The SIMS dashboard has run successfully and is attached.  It may contain personal information so please be aware of data protection issues." 'Use this if sending as attachment. Edit message as appropriate
       
    Set CDO_Mail = CreateObject("CDO.Message")
    On Error GoTo Error_Handling
    
    Set CDO_Config = CreateObject("CDO.Configuration")
    CDO_Config.Load -1
    
    Set SMTP_Config = CDO_Config.Fields
    
    With SMTP_Config
        .item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
        .item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = strsmtp
        .item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
'        .Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = ""
'        .Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = ""

        .Update
    End With
    
    With CDO_Mail
        Set .Configuration = CDO_Config
    End With
    
    CDO_Mail.Subject = strSubject
    CDO_Mail.From = strfrom
    CDO_Mail.To = strto
    CDO_Mail.TextBody = strBody
    CDO_Mail.CC = strCc
    CDO_Mail.BCC = strBcc
'    CDO_Mail.AddAttachment save_path + "\Todays_Output.xlsm" 'Adds as attachment, including .pdf if change extension and include in Save routine
    CDO_Mail.Send
    
Error_Handling:
    If Err.Description <> "" Then MsgBox Err.Description

End Sub

Sub WRAPPERFORTESTING_SendOutlookEmail()
    Dim SAVE_PATH As String
    If Sheets("Set up").Range("B6") = "" Then
        SAVE_PATH = ThisWorkbook.path
    Else
        SAVE_PATH = Sheets("Set up").Range("B6")
    End If
    Call send_mail_Outlook(SAVE_PATH)

End Sub

Sub send_mail_Outlook(SAVE_PATH As String)
    If DONT_SEND_EMAIL_MODE Then Exit Sub
    
    Dim OutApp As Object
    Dim OutMail As Object
    Dim strCc As String, strBcc As String, strBody As String
    
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)
    
    Dim sOutput As String
    'sOutput = "\Todays_Output.xlsm"
    Dim sYear As String, sMonth As String, sDate As String
    sYear = Year(Date)
    sMonth = Month(Date)
    sDate = Day(Date)
    sOutput = "\Todays_Output (" & sYear & "-" & sMonth & "-" & sDate & ").xlsm"
    
    strCc = ""
    strBcc = ""
    strBody = "The SIMS dashboard has run successfully and is available here:" & vbCrLf & vbCrLf & System

    'check exists:
    If Dir(SAVE_PATH + sOutput) <> "" Then
        If DEBUG_MODE Then MsgBox "File exists."
    Else
        If DEBUG_MODE Then MsgBox "File doesn't exist." 'TODO: Push to log, level error.
        Exit Sub
    End If
    
    'On Error Resume Next
    With OutMail
        .To = OurSystem.MAIL_TO
        .CC = ""
        .BCC = ""
        .Subject = "SIMS Dashboard"
        
        'Use this if sending as link
        '.body = strBody
        
        'Use these if sending as attachment
        .body = "The SIMS dashboard has run successfully and is attached.  It may contain personal information so please be aware of data protection issues." 'Use this if sending as attachment. Edit message as appropriate
        .Attachments.Add (SAVE_PATH + sOutput) 'Adds as attachment, including .pdf if change extension and include in Save routine
        
        .Send
    End With
    'On Error GoTo 0

    Set OutMail = Nothing
    Set OutApp = Nothing
End Sub

Sub Unhide()

' Unhides sheets
    
    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
    ws.Visible = xlSheetVisible
    Next ws

End Sub

Sub Restore_toolbars()
' Restores toolbars
    Application.WindowState = xlMaximized
    Application.DisplayFullScreen = False
    Application.DisplayFormulaBar = True
End Sub

Sub get_params()
' Gets report parameters
Dim thepath As String, save_destination As String, USERPASSWORD As String, USERNAME As String, SERVER As String, DATABASE As String, path As String
Dim reportnumber As String, ReportName As String
Dim paramdetails As String

'TODO: This needs to interact with the ROWs in the ReportList sheet instead.
    
    thepath = OurSystem.ROOT_PATH
    USERNAME = OurSystem.USERNAME
    USERPASSWORD = OurSystem.USERPASSWORD
    
    SERVER = OurSystem.SERVER
    DATABASE = OurSystem.DATABASE
    save_destination = """" + thepath + "\paramdefs.xml" + """"
    
    path = OurSystem.SIMS_PATH + "\CommandReporter.exe"
    reportnumber = InputBox("Report number?")
    Select Case reportnumber
        Case 1
            ReportName = Sheets("Set up").Range("A16")
        Case 2
            ReportName = Sheets("Set up").Range("B16")
        Case 3
            ReportName = Sheets("Set up").Range("C16")
        Case 4
            ReportName = Sheets("Set up").Range("A21")
        Case 5
            ReportName = Sheets("Set up").Range("B21")
        Case 6
            ReportName = Sheets("Set up").Range("C21")
        Case 7
            ReportName = Sheets("Set up").Range("A26")
        Case 8
            ReportName = Sheets("Set up").Range("B26")
        Case 9
            ReportName = Sheets("Set up").Range("C26")
        Case 10
            ReportName = Sheets("Set up").Range("A33")
        Case 11
            ReportName = Sheets("Set up").Range("B33")
        Case 12
            ReportName = Sheets("Set up").Range("C33")
        Case 13
            ReportName = Sheets("Set up").Range("A38")
        Case 14
            ReportName = Sheets("Set up").Range("B38")
        Case 15
            ReportName = Sheets("Set up").Range("C38")
        Case 16
            ReportName = Sheets("Set up").Range("A43")
        Case 17
            ReportName = Sheets("Set up").Range("B43")
        Case 18
            ReportName = Sheets("Set up").Range("C43")
    End Select
     
    ReportName = """" + ReportName + """"
    paramdetails = path + " /USER:" + USERNAME + " /PASSWORD:" + USERPASSWORD + " /SERVERNAME:" + SERVER + " /DATABASENAME:" + DATABASE + " /REPORT:" + ReportName + " /PARAMDEF /OUTPUT:" + save_destination
    
    Dim TaskID As Long
    'TODO: Shift this to commandlinehandler object.
    TaskID = Shell(paramdetails, 1)
End Sub

Sub WRAPPERFORTESTING_Save()
    Dim SAVE_PATH  As String
    If Sheets("Set up").Range("B6") = "" Then
        SAVE_PATH = ThisWorkbook.path
    Else
        SAVE_PATH = Sheets("Set up").Range("B6")
    End If

    Call Save(False) 'False means it won't quit.

End Sub

Sub Save(Optional bReallyQuit As Boolean = True)
    Dim sYear As String, sMonth As String, sDate As String
    
    sYear = Year(Date)
    sMonth = Month(Date)
    sDate = Day(Date)

'   Hides Setup sheet (don't delete, as zoom on title will not be able to detect if Setup or Run mode)
    Sheets("Set up").Visible = False
    Sheets("TODOs").Visible = False
    Sheets("Expected Staff").Visible = False
    
    

'   Save as new workbook
    Application.DisplayAlerts = False
    ActiveWorkbook.SaveCopyAs filename:= _
        OurSystem.ROOT_PATH + "\Todays_Output.xlsm"
    ActiveWorkbook.SaveCopyAs filename:= _
        OurSystem.ROOT_PATH + "\Todays_Output (" & sYear & "-" & sMonth & "-" & sDate & ").xlsm"
        
'   Save as .pdf (first removing page breaks)
        
'    With Worksheets("Dashboard")
'        .PageSetup.Orientation = xlLandscape
'        .Columns("X").PageBreak = xlPageBreakManual
'        .Rows("50").PageBreak = xlPageBreakManual
'    End With
'
'    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
'        save_path + "Dashboard.pdf", Quality:= _
'        xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, _
'        OpenAfterPublish:=False

'   Closes Excel (will not take effect until control passed back to sub from which called has ended)
    If bReallyQuit Then
        If DEBUG_MODE = False Then ' Never quit in debug mode.
            If Sheets("Set up").Range("K8") = 1 Then Application.Quit ' Never quit in setup mode. [only quit if not debug AND operating - regardless of email].
        End If
    End If

End Sub

Sub run_part()
''skip' means that it does't re-run the data.
'   Runs from previously saved .csv files (ie once established Command Reporter works OK)
    skip = "True"
    Call auto_open(False)
    
End Sub
Sub run_full()
''skip' means that it does't re-run the data.
'   Runs from previously saved .csv files (ie once established Command Reporter works OK)
    skip = "False"
    Call auto_open(False)
    
End Sub

Sub SizeAndAlignCharts()

'   Utility to tidy up all charts on a page. Position mouse on top left chart
'   Charts must be named in order (A1, A2 etc)
'   Beware that if you rename a chart, Excel still remembers previous name, so ChartObjects ("Chart 10").name could show as (eg) Chart 15. So ensure names have not been used before.

    If ActiveChart Is Nothing Then
        MsgBox "Select a chart to be used as the base for the sizing"
        Exit Sub
    End If
    
    Dim NumRows As Integer, W As Integer, H As Integer, TopPosition As Integer, LeftPosition As Integer
    Dim i As Integer
    
    MsgBox "Ensure gaps are filled with correctly named place-holder charts"
    NumRows = InputBox("How many rows of charts (numbered 1, 2, 3, etc)?")
    If Err.Number <> 0 Then Exit Sub
    If NumRows < 1 Then Exit Sub

    'Get size of active chart
    W = ActiveChart.Parent.Width
    H = ActiveChart.Parent.Height

    'Change starting positions, if necessary
    TopPosition = 100
    LeftPosition = 20
        
    On Error Resume Next
        
    For i = 1 To NumRows

    With ActiveSheet.ChartObjects("Chart A" + CStr(i) + "")
            .Width = W
            .Height = H
            .Left = LeftPosition
            .Top = TopPosition + ((i - 1) Mod NumRows) * H
        End With
   
    Next i
    
    
    For i = 1 To NumRows

    With ActiveSheet.ChartObjects("Chart B" + CStr(i) + "")
            .Width = W
            .Height = H
            .Left = LeftPosition + (1 * W)
            .Top = TopPosition + (((i - 1) Mod NumRows) * H)
        End With
   
    Next i
    
      
    For i = 1 To NumRows

    With ActiveSheet.ChartObjects("Chart C" + CStr(i) + "")
            .Width = W
            .Height = H
            .Left = LeftPosition + (2 * W)
            .Top = TopPosition + ((i - 1) Mod NumRows) * H
        End With
   
    Next i
    
    
    For i = 1 To NumRows

    With ActiveSheet.ChartObjects("Chart D" + CStr(i) + "")
            .Width = W
            .Height = H
            .Left = LeftPosition + (2 * W)
            .Top = TopPosition + ((i - 1) Mod NumRows) * H
        End With
   
    Next i
    For i = 1 To NumRows

    With ActiveSheet.ChartObjects("Chart E" + CStr(i) + "")
            .Width = W
            .Height = H
            .Left = LeftPosition + (2 * W)
            .Top = TopPosition + ((i - 1) Mod NumRows) * H
        End With
   
    Next i
    
    For i = 1 To NumRows

    With ActiveSheet.ChartObjects("Chart F" + CStr(i) + "")
            .Width = W
            .Height = H
            .Left = LeftPosition + (2 * W)
            .Top = TopPosition + ((i - 1) Mod NumRows) * H
        End With
   
    Next i
On Error GoTo 0

End Sub

Sub SizeAndAlignAndFormatCharts()

    Dim NumRows As Integer, W As Integer, H As Integer, TopPosition As Integer, LeftPosition As Integer
    Dim i As Integer
'   Utility to tidy up all charts on a page. Position mouse on top left chart
'   Charts must be named in order (A1, A2 etc)
'   Beware that if you rename a chart, Excel still remembers previous name, so ChartObjects ("Chart 10").name could show as (eg) Chart 15. So ensure names have not been used before.
'   For more than 9 schools, use different starting positions, or add more code etc
'   For formatting as well, ensure that you have correct details in Sub Format

    
    If ActiveChart Is Nothing Then
        MsgBox "Select a chart to be used as the base for the sizing"
        Exit Sub
    End If
    
    Dim FormatPlease As String
    FormatPlease = InputBox("Enter Yes to format as well? (Ensure you have recorded macro and replaced code in existing macro)")
    
    MsgBox "Ensure gaps are filled with correctly named place-holder charts"
    NumRows = InputBox("How many rows of charts (numbered 1, 2, 3, etc)?")
    If Err.Number <> 0 Then Exit Sub
    If NumRows < 1 Then Exit Sub

    'Get size of active chart
    W = ActiveChart.Parent.Width
    H = ActiveChart.Parent.Height

    'Change starting positions, if necessary
    TopPosition = 100
    LeftPosition = 20
        
    On Error Resume Next
        
    For i = 1 To NumRows

    With ActiveSheet.ChartObjects("Chart A" + CStr(i) + "")
            .Width = W
            .Height = H
            .Left = LeftPosition
            .Top = TopPosition + ((i - 1) Mod NumRows) * H
        End With
        
    If FormatPlease = "Yes" Then
        Call Format("Chart A" + CStr(i) + "")
    End If
    
    Next i
    
    
    For i = 1 To NumRows

    With ActiveSheet.ChartObjects("Chart B" + CStr(i) + "")
            .Width = W
            .Height = H
            .Left = LeftPosition + (1 * W)
            .Top = TopPosition + (((i - 1) Mod NumRows) * H)
        End With
        
    If FormatPlease = "Yes" Then
        Call Format("Chart B" + CStr(i) + "")
    End If
    
    Next i
    
      
    For i = 1 To NumRows

    With ActiveSheet.ChartObjects("Chart C" + CStr(i) + "")
            .Width = W
            .Height = H
            .Left = LeftPosition + (2 * W)
            .Top = TopPosition + ((i - 1) Mod NumRows) * H
        End With
        
    If FormatPlease = "Yes" Then
        Call Format("Chart C" + CStr(i) + "")
        
    End If
   
    Next i
    
On Error GoTo 0

End Sub

Sub Format(ChartName)

'   Called from within SizeAndAlignAndFormat macro.  To change the look of the charts, record a macro as you make...
'   ...changes to a single chart of shape.  Then substitute the code below.
'   Replace occurrences of specific chart or shape names with "Chartname"

With ActiveSheet.Shapes(ChartName).ThreeD
        .SetPresetCamera (msoCameraOrthographicFront)
        .RotationX = 0
        .RotationY = 0
        .RotationZ = 0
        .FieldOfView = 0
        .LightAngle = 0
        .PresetLighting = msoLightRigSoft
        .PresetMaterial = msoMaterialMatte2
        .Depth = 0
        .ContourWidth = 3.5
        .ContourColor.RGB = RGB(255, 255, 255)
        .BevelTopType = msoBevelArtDeco
        .BevelTopInset = 5
        .BevelTopDepth = 5
        .BevelBottomType = msoBevelNone
    End With
    With ActiveSheet.Shapes(ChartName).Shadow
        .Type = msoShadow25
        .Visible = msoTrue
        .Style = msoShadowStyleOuterShadow
        .Blur = 8.5
        .OffsetX = 6.1232339957E-17
        .OffsetY = 1
        .RotateWithShape = msoTrue
        .ForeColor.RGB = RGB(0, 0, 0)
        .Transparency = 0
        .Size = 100
    End With
    ActiveSheet.Shapes(ChartName).Line.Visible = msoFalse
    With ActiveSheet.Shapes(ChartName).Fill
        .Visible = msoTrue
        .ForeColor.RGB = RGB(217, 217, 217)
        .Solid
    End With
    With ActiveSheet.Shapes(ChartName).Fill
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorBackground1
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = -0.150000006
        .Transparency = 0
        .Solid
    End With
    ActiveChart.SetElement (msoElementLegendNone)
    ActiveChart.ShowAllFieldButtons = False
End Sub

'
'Sub TestSingleReport_CLI()
'    Dim sReportName As String
'
'    'TODO:
'    'For now, only tests report 1. See auto_open for this code. Needs refactoring.
'
'    'sReportName = InputBox("Input number of report to test")
'    'Call Get_data("Report" & sReportName)
''__________________
''   Sets up variables
'    thepath = ThisWorkbook.path
'    USERNAME = Sheets("Set up").Range("B1")
'    USERPASSWORD = Sheets("Set up").Range("B2")
'
'    If Sheets("Set up").Range("B6") = "" Then
'        save_path = ThisWorkbook.path
'    Else
'        save_path = Sheets("Set up").Range("B6")
'    End If
'
'    EMAILPATH = Sheets("Set up").Range("B10")
'    strfrom = Sheets("Set up").Range("B11")
'    strto = Sheets("Set up").Range("B8")
'    strsmtp = Sheets("Set up").Range("B9")
'
'    SERVER = Sheets("Set up").Range("B4")
'    DATABASE = Sheets("Set up").Range("B5")
'    path = Sheets("Set up").Range("B3") + "\CommandReporter.exe"
'
''   Replaces double quotes with single in report parameters
'    Sheets("Set up").Range("A17,B17,C17,A22,B22,C22,A27,B27,C27").Replace What:="""", Replacement:="'", LookAt:=xlPart, _
'        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
'        ReplaceFormat:=False
'
'    Sheets("Dashboard").Select
'
'
'
'
'
'
''_______________-
'
'    '   ////////////////////////Report 1//////////////////////////////
'
'    'If Not skip = "True" Then
'
'        If Not Sheets("Set up").Range("A16") = "" Then
'
'            save_destination = """" + thepath + "\Report1.csv" + """"
'            ReportName = """" + Sheets("Set up").Range("A16") + """"
'            parameterid_1 = Sheets("Set up").Range("A17")
'            Dates = Sheets("Set up").Range("A18")
'
'            If Sheets("Set up").Range("A17") = "" Then
'                date_parameters = ""
'            Else
'                date_parameters = """<ReportParameters>" + parameterid_1 + "<Values>" + Dates + "</Values></Parameter></ReportParameters>"""
'            End If
'
'            progdetails = """" + path + """" + " /USER:" + USERNAME + " /PASSWORD:" + USERPASSWORD + " /SERVERNAME:" + SERVER + " /DATABASENAME:" + DATABASE + " /REPORT:" + ReportName + " /OUTPUT:" + save_destination + " /PARAMS:" + date_parameters
'
'            If Sheets("Set up").Range("K8") = 2 Then
'                MsgBox "Remember - Fn/Ctrl/B or Ctrl/Pause or Ctrl/Break to exit macro at any stage"
'                MsgBox progdetails
'                Worksheets("Set up").Range("K4") = progdetails
'            End If
'
'            Call Run_Cmd_Reporter
'            Call Progress
'        End If
'
'End Sub
