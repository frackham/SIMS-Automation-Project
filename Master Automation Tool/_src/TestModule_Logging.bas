Attribute VB_Name = "TestModule_Logging"
Option Explicit

Option Private Module

'@TestModule
Private Assert As Object

'@ModuleInitialize
Public Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = CreateObject("Rubberduck.AssertClass")

End Sub

'@ModuleCleanup
Public Sub ModuleCleanup()
    'this method runs once per module.
End Sub

'@TestInitialize
Public Sub TestInitialize()
    'this method runs before every test in the module.
End Sub

'@TestCleanup
Public Sub TestCleanup()
    'this method runs after every test in the module.
End Sub



'@TestMethod
Public Sub TestMethod1()
    On Error GoTo TestFail

    'Arrange

    'Act
        Call WriteLineToSystemLog("Test: System Log Test", "---", "---")
        
    'Assert
        Call Assert.AreEqual(GetLastLog_System, "Test: System Log Test,---,---")

TestExit:
    Exit Sub
TestFail:
    If Err.Number <> 0 Then
        Assert.Fail "Test raised an error: " & Err.Description
    End If
    Resume TestExit
End Sub

