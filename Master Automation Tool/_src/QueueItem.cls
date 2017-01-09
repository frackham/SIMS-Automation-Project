VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "QueueItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit
'QueueItem.
'Generic class for any actions needing to go into the queue.



Public sInternalName As String
Public sReportName As String
Public sParameterString As String
Public sParameterValues As String
Public pCommandString As String


Public dTimeTaken As Date
Public dExpectedTimeTaken As Date

''''''''''''''''''''''
' CommandString property
''''''''''''''''''''''
Public Property Get CommandString() As String
    CommandString = pCommandString
End Property
Public Property Let CommandString(Value As String)
    pCommandString = Value
End Property

''''''''''''''''''''''
' InternalName property
''''''''''''''''''''''
Public Property Get InternalName() As String
    InternalName = sInternalName
End Property
Public Property Let InternalName(Value As String)
    sInternalName = Value
End Property

''''''''''''''''''''''
' ReportName property
''''''''''''''''''''''
Public Property Get ReportName() As String
    ReportName = sReportName
End Property
Public Property Let ReportName(Value As String)
    sReportName = Value
End Property

''''''''''''''''''''''
' ParameterString property
''''''''''''''''''''''
Public Property Get ParameterString() As String
    ParameterString = sParameterString
End Property
Public Property Let ParameterString(Value As String)
    sParameterString = Value
End Property

''''''''''''''''''''''
' ParameterValues property
''''''''''''''''''''''
Public Property Get ParameterValues() As String
    ParameterValues = sParameterValues
End Property
Public Property Let ParameterValues(Value As String)
    sParameterValues = Value
End Property



Private Sub Class_Initialize()
    'TODO: Set parameters.
    'TODO: Set postactions.
    'TODO: Set preactions(?).
    'TODO: optional - add expected time taken.
    
    'TODO: Define fail actions, if any.
    'TODO: Define intended output recipients (do not define delivery mechanism...this should be handled depending on the target and config? E.g. mailing these, or saving in accessible folder for them.
    
    ''For each required property on the QueueItem...
    '    p.ReportName = ReportName
    '    p.ParameterString = ParameterString
End Sub

'Current implementation = Each line in the spreadsheet gets added as a new QueueItem in the QueueManager.
Sub Run()
    'TODO: Check that defined sufficiently.
    'TODO: Get commandlinehandler from global state (singleton pattern).
    'TODO: Pass commandlinehandler the expected time taken. Should quit when exceeds 10x ??
    If Me.ShouldBeRun Then
         
    End If
End Sub

Sub PreProcess()
    'TODO: Define this. What kind of actions might be taken here? Optional run checks (ie don't run if X).

End Sub
Sub PostProcess()
    'TODO: Define this. What kind of actions might be taken here?
    'Examples: Mine-style actions (calculate from ). Define Variables (e.g. sum/counts of columns).
    'Examples: Formatting of outputs.
    'Likely to need a post-processing module that this can call on.

End Sub



Function TimeTaken() As Date
    'Should this return a double?
    'Should this be handled at the collection level? No. This way, we can use this on a single report...

    'TODO: if not run, return 0 (not Null!).

End Function

Function ShouldBeRun() As Boolean
    'TODO: Preprocess rules might determine that should not run.
    'TODO: Date parameters (don't run at weekends, for example) might determine that should not run.1
    ShouldBeRun = True
End Function






