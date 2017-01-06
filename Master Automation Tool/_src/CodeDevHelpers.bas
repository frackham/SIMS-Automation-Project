Attribute VB_Name = "CodeDevHelpers"
Option Explicit
'See http://www.cpearson.com/excel/vbe.aspx on programmatic access to VBE object model.


'TODO: Move SRC folder location and logs folder location to Config on the system object.

Public Function CheckWhetherToImportSRCModules() As Boolean
    Dim FSO As Scripting.FileSystemObject, f As File, fo As Folder
    Dim dFileUpdated As Date, dThisWorkbookUpdated As Date
    If FSO Is Nothing Then
        Set FSO = New FileSystemObject
    End If

    dThisWorkbookUpdated = TimeStampThisLastModified()

    'for each .bas or .cls in SRC folder.
    Set fo = FSO.GetFolder(ThisWorkbook.path & "\" & "_src") 'TODO: Replace with path variable.
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
    Set FSO = Nothing
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

    If ActiveWorkbook.Name = ThisWorkbook.Name Then
        MsgBox "Select another destination workbook" & _
        "Not possible to import in this workbook "
        Exit Sub
    End If

    'Get the path to the folder with modules
    If FolderWithVBAProjectFiles = "Error" Then
        MsgBox "Import Folder not exist"
        Exit Sub
    End If

    ''' NOTE: This workbook must be open in Excel.
    szTargetWorkbook = ActiveWorkbook.Name
    Set wkbTarget = Application.Workbooks(szTargetWorkbook)
    
    If wkbTarget.VBProject.Protection = 1 Then
    MsgBox "The VBA in this workbook is protected," & _
        "not possible to Import the code"
    Exit Sub
    End If

    ''' NOTE: Path where the code modules are located.
    szImportPath = FolderWithVBAProjectFiles & "\"
        
    Set objFSO = New Scripting.FileSystemObject
    If objFSO.GetFolder(szImportPath).Files.Count = 0 Then
       MsgBox "There are no files to import"
       Exit Sub
    End If

    'Delete all modules/Userforms from the ActiveWorkbook
    Call DeleteVBAModulesAndUserForms

    Set cmpComponents = wkbTarget.VBProject.VBComponents
    
    ''' Import all the code modules in the specified path
    ''' to the ActiveWorkbook.
    For Each objFile In objFSO.GetFolder(szImportPath).Files
    
        If (objFSO.GetExtensionName(objFile.Name) = "cls") Or _
            (objFSO.GetExtensionName(objFile.Name) = "frm") Or _
            (objFSO.GetExtensionName(objFile.Name) = "bas") Then
            cmpComponents.Import objFile.path
        End If
        
    Next objFile
    
    MsgBox "Import is ready"
End Sub

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

Function DeleteVBAModulesAndUserForms()
        Dim VBProj As VBIDE.VBProject
        Dim VBComp As VBIDE.VBComponent
        
        Set VBProj = ActiveWorkbook.VBProject
        
        For Each VBComp In VBProj.VBComponents
            If VBComp.Type = vbext_ct_Document Then
                'Thisworkbook or worksheet module
                'We do nothing
            Else
                VBProj.VBComponents.Remove VBComp
            End If
        Next VBComp
End Function

