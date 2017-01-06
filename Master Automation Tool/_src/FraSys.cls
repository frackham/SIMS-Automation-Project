VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FraSys"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit
'Class FraSys.
'Horrific personal name for a 'system' class, but at least it's memorable for me.
'Singleton Pattern

Dim str_username As String
Dim str_userpassword As String
Dim str_server As String
Dim str_database As String
Dim str_simspath As String
Dim str_thisworkbookpath As String
Dim str_rootpath As String
Dim str_savepath As String
Dim str_emailpath As String
Dim bIS_INITIALISED As Boolean
Dim str_mailerfrom As String
Dim str_mailerto As String
Dim str_mailersmtp As String

Dim bool_debugmode As Boolean



Private Sub Class_Initialize()
    'Set as blank values on initialise
    str_username = ""
    str_userpassword = ""
    str_server = ""
    str_database = ""
    
    bIS_INITIALISED = False
End Sub

Public Sub InitialiseSystemConfig()
    'CRITICAL: Currently manually set up with account details.
    
    
'   Sets up variables
    str_username = Sheets("Set up").Range("B1")
    str_userpassword = Sheets("Set up").Range("B2")
    
    str_server = Sheets("Set up").Range("B4")
    str_database = Sheets("Set up").Range("B5")
    str_simspath = Sheets("Set up").Range("B3") + "\CommandReporter.exe"
    str_emailpath = Sheets("Set up").Range("B10")
    
    
    str_rootpath = ThisWorkbook.path
    
    If Sheets("Set up").Range("B6") = "" Then 'If we've defined a save location, use it, else save to same path as this workbook.
        str_savepath = ThisWorkbook.path
    Else
        str_savepath = Sheets("Set up").Range("B6")
    End If
    

    str_mailerfrom = Sheets("Set up").Range("B11")
    str_mailerto = Sheets("Set up").Range("B8")
    str_mailersmtp = Sheets("Set up").Range("B9")
    
    
    
    
    str_username = "f.admin" 'TODO: Remove my login details.
    str_userpassword = "peinture"
    str_server = "" 'Correctly blank.
    str_database = "" 'Correctly blank.
    
    bool_debugmode = DEBUG_MODE 'TODO: Yuk. Check this.
    
    bIS_INITIALISED = True
End Sub

Function SIMS_PATH() As String
    SIMS_PATH = str_simspath
End Function
Function EMAIL_PATH() As String
    EMAIL_PATH = str_emailpath
End Function
Function THISWORKBOOK_PATH() As String
    THISWORKBOOK_PATH = ThisWorkbook.path
End Function
Function ROOT_PATH() As String
    ROOT_PATH = str_rootpath
End Function
Function SAVE_PATH() As String
    SAVE_PATH = str_savepath
End Function
Function DATABASE() As String
    DATABASE = str_database
End Function
Function SERVER() As String
    SERVER = str_server
End Function
Function USERNAME() As String
    USERNAME = str_username
End Function
Function USERPASSWORD() As String
    USERPASSWORD = str_userpassword
End Function
Function MAIL_FROM() As String
    MAIL_FROM = str_mailerfrom
End Function
Function MAIL_TO() As String
    MAIL_TO = str_mailerto
End Function
'Function MAIL_FROM() As String
'    MAIL_FROM = str_emailpath
'End Function


Function DEBUG_MODE() As Boolean
    DEBUG_MODE = bool_debugmode
End Function
