VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "QueueManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit
'QueueManager
Private thisCMD As CommandLineHandler
Private thissys As FraSys
Private ThisQueueItems As New Collection
'--------------
    
    'Our Name                        SIMS Report Name                       Parameter Tag                                                           Parameter Values
    'Report1 - Location this week    Automation - Behaviour (DateFromTo)    <Parameter id='61jJQgmWcKF6KBylfgHxlQ==' subreportfilter='True'>        <Start>2016-12-12</Start><End>2016-12-19</End>
     

'--------------


Sub Start()
    Dim thisItem As QueueItem
    'TODO: Iterate through each item, and call .Run.
    
    For Each thisItem In ThisQueueItems
        thisItem.PreProcess
        'thisItem.Run
        Me.RunQueueItem (thisItem)
        thisItem.PostProcess
        'Do something on return with each (log?).
        'Do follow on actions now? or add to Action Queue. Delivery mechanisms (e.g. mail, savefolder accessible to recipient(s)...?).
    Next thisItem
End Sub

Sub PopulateFromValues()
    Dim y As Integer
    Dim sInternalName As String, sReportName As String, sParameterTag As String, sParameterValues As String

    With ThisWorkbook.Worksheets("ReportList")
        y = 1
        While .Range("A1").Offset(y, 0).Value & "" <> ""
            'Our Name    SIMS Report Name    Parameter Tag   Parameter Values
            
            '   Use this line to replace incorrect quote marks in case.
            .Range("A1").Offset(y, 2).Replace What:="""", Replacement:="'", LookAt:=xlPart, _
                SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
                ReplaceFormat:=False
            
            
            sInternalName = .Range("A1").Offset(y, 0).Value
            sReportName = .Range("A1").Offset(y, 1).Value
            sParameterTag = .Range("A1").Offset(y, 2).Value
            sParameterValues = .Range("A1").Offset(y, 3).Value
            
            'Amend/update
            sParameterValues = Me.AmendDateParameters(sParameterValues)
             
            Call Me.Add(thissys, sInternalName, sReportName, sParameterTag, sParameterValues)
            
            
            
            'TODO:
            y = y + 1
        Wend
    End With
    
    
End Sub
Function AmendDateParameters(sParams As String) As String
    'TODO: Add support for daily, monthly, weekly etc.

    AmendDateParameters = sParams
End Function

'Collection class as per notes from:
'http://www.cpearson.com/excel/CollectionClass.aspx

'Core Functionality: CRUD, iteration.
Sub Add(System As FraSys, InternalName As String, ReportName As String, ParameterString As String, ParameterValues As String)
    'CHECK: This needs updating each time the QueueItem is updated.
    'TEST: This needs a test, so that each time there are amendments it flags it up here as well as in QueueItem.
    'create a new QueueItem and add to collection
    Dim p As New QueueItem, cmdString As String
    
    'For each required property on the QueueItem...
    p.InternalName = InternalName
    p.ReportName = ReportName
    p.ParameterString = ParameterString
    p.ParameterValues = ParameterValues
    
    
    '   Create runstring. (rather than at run time, as then we can check each runstring without actually running...
    cmdString = thisCMD.Construct_CMD_CallString(System.ROOT_PATH, ReportName, System.ROOT_PATH, ParameterString, ParameterValues)
    p.CommandString = cmdString
    
    ThisQueueItems.Add p
End Sub
Property Get Count() As Long
    'return the number of people
    Count = ThisQueueItems.Count
End Property
Property Get item(NameOrNumber As Variant) As QueueItem
    'return this particular person
    Set item = ThisQueueItems(NameOrNumber)
End Property
Sub Remove(NameOrNumber As Variant)
    'remove this person from collection
    ThisQueueItems.Remove NameOrNumber
End Sub
Property Get Items() As Collection
    'return the collection of items
    Set Items = ThisQueueItems
End Property

Sub RunQueueItem(msg As QueueItem)

    thisCMD.Run_Cmd_Reporter (msg)

End Sub


Private Sub Class_Initialize()
    'Minimise information here, makes it harder to unit test.


End Sub
Sub Setup(CMD As CommandLineHandler, sys As FraSys)
    Set thisCMD = CMD
    Set thissys = sys
End Sub









