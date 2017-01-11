Attribute VB_Name = "IOHelpers"
Option Explicit

  
  'TODO: Add unit tests. For example, this file should always return 70, shouldn't it?
  
Function IsFileUnavailable(filename As String)
    Dim filenum As Integer, errnum As Integer

    On Error Resume Next   ' Turn error checking off.
    filenum = FreeFile()   ' Get a free file number.
    ' Attempt to open the file and lock it.
    Open filename For Input Lock Read As #filenum
    Close filenum          ' Close the file.
    errnum = Err           ' Save the error number that occurred.
    On Error GoTo 0        ' Turn error checking back on.

    ' Check to see which error occurred.
    Select Case errnum

        ' No error occurred.
        ' File is NOT already open by another user.
        Case 0
         IsFileUnavailable = False

        ' Error number for "Permission Denied."
        ' File is already opened by another user.
        Case 70
            IsFileUnavailable = True


        ' Error number for "File doesn't exist."
        ' File can't be found.
        Case 53
            IsFileUnavailable = False

        ' Another error occurred.
        Case Else
            'Error errnum
            IsFileUnavailable = True
            If DEBUG_MODE Then MsgBox (errnum & ", " & Err.Description)
            Call WriteLineToSystemLog("IsFileUnavailable", "Unknown error on checking file availability for " & filename & ". Err number: " & errnum & ", " & Err.Description, "-")
    End Select

    
End Function

