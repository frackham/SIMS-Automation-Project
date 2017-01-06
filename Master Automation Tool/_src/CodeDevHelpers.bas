Attribute VB_Name = "CodeDevHelpers"
Option Explicit
'See http://www.cpearson.com/excel/vbe.aspx on programmatic access to VBE object model.


'TODO: Move SRC folder location and logs folder location to Config on the system object.

Public Function CheckWhetherToImportSRCModules() As Boolean
    Dim fso As Scripting.FileSystemObject, f As File, fo As Folder
    Dim dFileUpdated As Date, dThisWorkbookUpdated As Date
    If fso Is Nothing Then
        Set fso = New FileSystemObject
    End If

    dThisWorkbookUpdated = TimeStampThisLastModified()

    'for each .bas or .cls in SRC folder.
    Set fo = fso.GetFolder(ThisWorkbook.path & "\" & "_src") 'TODO: Replace with path variable.
    For Each f In fo.Files
        Select Case UCase(Trim(Replace(Right(f.ShortName, 4), ".", "")))
            Case "CLS", "BAS"
                dFileUpdated = f.DateLastModified
                If GetIsCheckDateBeforeDate(dThisWorkbookUpdated, dFileUpdated) Then 'if this file has been updated since the workbook, update the modules.
                    CheckWhetherToImportSRCModules = True
                    Exit For
                End If
            Case Else
                'Do nothing.
        End Select
    Next f
    Set fso = Nothing
End Function

'See http://www.rondebruin.nl/win/s9/win002.htm for the functions below.
Public Sub ExportModules()
    Dim bExport As Boolean
    Dim wkbSource As Excel.Workbook
    Dim szSourceWorkbook As String
    Dim szExportPath As String
    Dim szFileName As String
    Dim cmpComponent As VBIDE.VBComponent

    ''' The code modules will be exported in a folder named.
    ''' VBAProjectFiles in the Documents folder.
    ''' The code below create this folder if it not exist
    ''' or delete all files in the folder if it exist.
    If FolderWithVBAProjectFiles = "Error" Then
        MsgBox "Export folder for code backup does not exist...create a _src folder in the working directory."
        Exit Sub
    End If
    
    On Error Resume Next
        Kill FolderWithVBAProjectFiles & "\*.*"
    On Error GoTo 0

    ''' NOTE: This workbook must be open in Excel.
    szSourceWorkbook = ActiveWorkbook.Name
    Set wkbSource = Application.Workbooks(szSourceWorkbook)
    
    If wkbSource.VBProject.Protection = 1 Then
    MsgBox "The VBA in this workbook is protected," & _
        "not possible to export the code"
    Exit Sub
    End If
    
    szExportPath = FolderWithVBAProjectFiles & "\"
    
    For Each cmpComponent In wkbSource.VBProject.VBComponents
        
        bExport = True
        szFileName = cmpComponent.Name

        ''' Concatenate the correct filename for export.
        Select Case cmpComponent.Type
            Case vbext_ct_ClassModule
                szFileName = szFileName & ".cls"
            Case vbext_ct_MSForm
                szFileName = szFileName & ".frm"
            Case vbext_ct_StdModule
                szFileName = szFileName & ".bas"
            Case vbext_ct_Document
                ''' This is a worksheet or workbook object.
                ''' Don't try to export.
                bExport = False
        End Select
        
        If bExport Then
            ''' Export the component to a text file.
            cmpComponent.Export szExportPath & szFileName
            'Note that codedevhelpers module (this module) *is* exported each time even though it is not imported.
            
        ''' remove it from the project if you want
        '''wkbSource.VBProject.VBComponents.Remove cmpComponent
        
        End If
   
    Next cmpComponent

    If DEBUG_MODE Then MsgBox "Export is ready"
End Sub

Public Sub ImportModules()
    'Modified from original to import into this current workbook.
    Dim wkbTarget As Excel.Workbook
    Dim objFSO As Scripting.FileSystemObject
    Dim objFile As Scripting.File
    Dim szTargetWorkbook As String
    Dim szImportPath As String
    Dim szFileName As String
    Dim cmpComponents As VBIDE.VBComponents

'    If ActiveWorkbook.Name = ThisWorkbook.Name Then
'        MsgBox "Select another destination workbook" & _
'        "Not possible to import in this workbook "
'        Exit Sub
'    End If

    'Get the path to the folder with modules
    If FolderWithVBAProjectFiles = "Error" Then
        MsgBox "Import Folder not exist"
        Exit Sub
    End If

'    ''' NOTE: This workbook must be open in Excel.
'    szTargetWorkbook = ThisWorkbook.Name
'    Set wkbTarget = Application.Workbooks(szTargetWorkbook)
'
'    If wkbTarget.VBProject.Protection = 1 Then
'    MsgBox "The VBA in this workbook is protected," & _
'        "not possible to Import the code"
'    Exit Sub
'    End If

    ''' NOTE: Path where the code modules are located.
    szImportPath = FolderWithVBAProjectFiles & "\"
        
    Set objFSO = New Scripting.FileSystemObject
    If objFSO.GetFolder(szImportPath).Files.Count = 0 Then
       MsgBox "There are no files to import"
       Exit Sub
    End If

'    'Delete all modules/Userforms from the ActiveWorkbook
'    Call DeleteVBAModulesAndUserForms_ThisWorkbook
'    'Call DeleteVBAModulesAndUserForms

    Set cmpComponents = ThisWorkbook.VBProject.VBComponents
    
    ''' Import all the code modules in the specified path
    ''' to the ActiveWorkbook.
    For Each objFile In objFSO.GetFolder(szImportPath).Files
    
        If (objFSO.GetExtensionName(objFile.Name) = "cls") Or _
            (objFSO.GetExtensionName(objFile.Name) = "frm") Or _
            (objFSO.GetExtensionName(objFile.Name) = "bas") Then
            
            'cmpComponents.Import objFile.path
            Call UpdateCodeInModule(cmpComponents, objFile.path)
            
        End If
        
    Next objFile
    
    If DEBUG_MODE Then MsgBox "Import is ready"
End Sub

Private Sub UpdateCodeInModule(cmpComponents As VBComponents, sImportFilePath As String)
    Dim fso As Scripting.FileSystemObject, f As File, fo As Folder, cmpModule As VBComponent
    Dim dFileUpdated As Date, dThisWorkbookUpdated As Date, sExtension As String, sThisFileName As String, bModuleExists As Boolean, sThisModuleName As String
    Dim sCode As String
    If fso Is Nothing Then
        Set fso = New FileSystemObject
    End If

    dThisWorkbookUpdated = TimeStampThisLastModified()

    'for each .bas or .cls in SRC folder.
    Set fo = fso.GetFolder(ThisWorkbook.path & "\" & "_src") 'TODO: Replace with path variable.
    For Each f In fo.Files
        sExtension = UCase(Trim(Replace(Right(f.Name, 4), ".", "")))
        Select Case sExtension
            Case "CLS", "BAS"
                sThisFileName = Replace(UCase(f.Name), sExtension, "")
                sThisFileName = Replace(sThisFileName, ".", "")
                'get if the module already exists. If so, replace the text (starting from option explicit).
                bModuleExists = False
                For Each cmpModule In cmpComponents
                    If cmpModule.Type = vbext_ct_Document Then
'                       'Thisworkbook or worksheet module
'                       'We do nothing
                    Else
                        sThisModuleName = UCase(cmpModule.Name)
                        
                        If sThisFileName = sThisModuleName Then
                            'UPDATE
                            bModuleExists = True
                            Exit For
                        End If
                        
                    End If
                Next cmpModule
    
                If bModuleExists Then
                    If sThisModuleName = "CODEDEVHELPERS" Then
                        'Skip for now. Don't want to import into this module.
                    Else
                        sCode = GetCodeFromFile_BASorCLS(f.path)
                        If Len(sCode) > 0 Then cmpModule.CodeModule.DeleteLines 1, cmpModule.CodeModule.CountOfLines 'Don't remove unless got something to add in!
                        cmpModule.CodeModule.AddFromString (sCode)
                    End If
                ElseIf bModuleExists = False Then
                    'if it doesn't exist, import
                    cmpComponents.Import sImportFilePath
                End If
                
            Case Else
                'Do nothing.
        End Select
    Next f
    
    Set fso = Nothing
    
End Sub


'
'Public Sub ImportModules()
'    'Modified from original to import into this current workbook.
'    Dim wkbTarget As Excel.Workbook
'    Dim objFSO As Scripting.FileSystemObject
'    Dim objFile As Scripting.File
'    Dim szTargetWorkbook As String
'    Dim szImportPath As String
'    Dim szFileName As String
'    Dim cmpComponents As VBIDE.VBComponents
'
''    If ActiveWorkbook.Name = ThisWorkbook.Name Then
''        MsgBox "Select another destination workbook" & _
''        "Not possible to import in this workbook "
''        Exit Sub
''    End If
'
'    'Get the path to the folder with modules
'    If FolderWithVBAProjectFiles = "Error" Then
'        MsgBox "Import Folder not exist"
'        Exit Sub
'    End If
'
'    ''' NOTE: This workbook must be open in Excel.
'    szTargetWorkbook = ThisWorkbook.Name
'    Set wkbTarget = Application.Workbooks(szTargetWorkbook)
'
'    If wkbTarget.VBProject.Protection = 1 Then
'    MsgBox "The VBA in this workbook is protected," & _
'        "not possible to Import the code"
'    Exit Sub
'    End If
'
'    ''' NOTE: Path where the code modules are located.
'    szImportPath = FolderWithVBAProjectFiles & "\"
'
'    Set objFSO = New Scripting.FileSystemObject
'    If objFSO.GetFolder(szImportPath).Files.Count = 0 Then
'       MsgBox "There are no files to import"
'       Exit Sub
'    End If
'
'    'Delete all modules/Userforms from the ActiveWorkbook
'    Call DeleteVBAModulesAndUserForms_ThisWorkbook
'    'Call DeleteVBAModulesAndUserForms
'
'    Set cmpComponents = wkbTarget.VBProject.VBComponents
'
'    ''' Import all the code modules in the specified path
'    ''' to the ActiveWorkbook.
'    For Each objFile In objFSO.GetFolder(szImportPath).Files
'
'        If (objFSO.GetExtensionName(objFile.Name) = "cls") Or _
'            (objFSO.GetExtensionName(objFile.Name) = "frm") Or _
'            (objFSO.GetExtensionName(objFile.Name) = "bas") Then
'            cmpComponents.Import objFile.path
'        End If
'
'    Next objFile
'
'    MsgBox "Import is ready"
'End Sub

Function FolderWithVBAProjectFiles() As String
    FolderWithVBAProjectFiles = ThisWorkbook.path & "\" & "_src"
End Function

'Function FolderWithVBAProjectFiles() As String
'    Dim WshShell As Object
'    Dim FSO As Object
'    Dim SpecialPath As String
'
'    Set WshShell = CreateObject("WScript.Shell")
'    Set FSO = CreateObject("scripting.filesystemobject")
'
'    SpecialPath = WshShell.SpecialFolders("MyDocuments")
'
'    If Right(SpecialPath, 1) <> "\" Then
'        SpecialPath = SpecialPath & "\"
'    End If
'
'    If FSO.FolderExists(SpecialPath & "VBAProjectFiles") = False Then
'        On Error Resume Next
'        MkDir SpecialPath & "VBAProjectFiles"
'        On Error GoTo 0
'    End If
'
'    If FSO.FolderExists(SpecialPath & "VBAProjectFiles") = True Then
'        FolderWithVBAProjectFiles = SpecialPath & "VBAProjectFiles"
'    Else
'        FolderWithVBAProjectFiles = "Error"
'    End If
'
'End Function

'Function DeleteVBAModulesAndUserForms()
'        Dim VBProj As VBIDE.VBProject
'        Dim VBComp As VBIDE.VBComponent
'
'        Set VBProj = ActiveWorkbook.VBProject
'
'        For Each VBComp In VBProj.VBComponents
'            If VBComp.Type = vbext_ct_Document Then
'                'Thisworkbook or worksheet module
'                'We do nothing
'            Else
'                VBProj.VBComponents.Remove VBComp
'            End If
'        Next VBComp
'End Function
'
'
'Function DeleteVBAModulesAndUserForms_ThisWorkbook()
'        Dim VBProj As VBIDE.VBProject
'        Dim VBComp As VBIDE.VBComponent
'
'        Set VBProj = ActiveWorkbook.VBProject
'
'        For Each VBComp In VBProj.VBComponents
'            If VBComp.Type = vbext_ct_Document Then
'                'Thisworkbook or worksheet module
'                'We do nothing
'            Else
'                VBProj.VBComponents.Remove VBComp
'            End If
'        Next VBComp
'End Function



Private Function GetCodeFromFile_BASorCLS(sFilePath As String) As String
    Dim fso As FileSystemObject, bStillAtTop As Boolean, sWord As String
    If fso Is Nothing Then
        Set fso = New FileSystemObject
    End If
    'the file we're going to read from
    Dim ts As TextStream
    '... we can open a text file with reference to it
    Set ts = fso.OpenTextFile(sFilePath, ForReading)
    
    'keep reading in lines till no more
    Dim ThisLine As String
    Dim i As Integer
    
    i = 0
    bStillAtTop = True 'Skip out as soon as we're not at the top of the actual code.
    Do Until ts.AtEndOfStream
        ThisLine = ts.ReadLine
        i = i + 1
        
        'Debug.Print "Line " & i, ThisLine
        
        If bStillAtTop Then
            If InStr(ThisLine, " ") > 0 Then
                sWord = UCase(Left(ThisLine, InStr(ThisLine, " ")))
            Else
                sWord = UCase(ThisLine)
            End If
            
            Select Case Trim(sWord)
                Case "ATTRIBUTE", "VERSION", "BEGIN", "END", "MULTIUSE"
                    
                    'Do nothing, still at top.
                Case "OPTION"
                    bStillAtTop = False
                Case ""
                    'Do nothing, still (potentially) at top.
                Case Else
                    bStillAtTop = False 'Probably
            End Select
        End If
        If bStillAtTop = False Then GetCodeFromFile_BASorCLS = GetCodeFromFile_BASorCLS & vbCrLf & ThisLine
    Loop
    
    'close down the file
    ts.Close

      
End Function
