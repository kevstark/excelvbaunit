Attribute VB_Name = "TestRunnerTester"
Option Explicit
Option Private Module
Rem order 10
Rem
Rem =head2
Rem sheetname
Rem
Rem This is part of the test suite for the test harness add-in. Its primary purpose is to test
Rem the functionality of the "TestRunner" class module.
Rem
Rem =head3
Rem sheetname Macros
Rem
Const ksErrMod As String = "TestRunnerTester"

Rem =head4 Function ~
Rem sheetname TestRun
Rem
Rem The "TestRunner" class module has a "Run" method, which is used to process the tests in a
Rem module. This procedure invokes it, using a dummy test module, and then inspects the
Rem results manager for evidence that the tests have been run. The original version overwrote
Rem the Test Results Manager, causing this and all subsequent modules to report zero successes
Rem and zero failures. This has been changed by creating and restoring a backup of the TRM.
Rem
Function TestRun() As Boolean
On Error GoTo ErrorHandler

    Dim tr As TestRunner
    Set tr = New TestRunner
    
    If BackupTRM Then err.Raise knCall, , ksCall
    
    Dim trm As FakeTestResultsManager
    Set trm = New FakeTestResultsManager

    If GetNames Then err.Raise knCall, , ksCall
    If tr.Run(AddInName, trm, "DummyTestModule4") Then err.Raise knCall, , ksCall
    
    Dim reply As String
    reply = trm.FunctionsCalled
    Set trm = Nothing
    If RestoreTRM Then err.Raise knCall, , ksCall

    If AssertEqual(":StartTestCase(DummyTestModule4.Test1):EndTestCase:EndTestSuite", reply) _
        Then err.Raise knCall, , ksCall

Exit Function
ErrorHandler:
ErrTrap ksErrMod, "TestRun"
TestRun = True
End Function
'Public Sub TestRun()
'
'    Dim tr As TestRunner
'    Set tr = New TestRunner
'
'    If BackupTRM Then err.Raise knCall, , ksCall
'
'    Dim trm As FakeTestResultsManager
'    Set trm = New FakeTestResultsManager
'
'    If GetNames Then err.Raise knCall, , ksCall
'    tr.Run AddInName, trm, "DummyTestModule4"
'
'    Dim reply As String
'    reply = trm.FunctionsCalled
'    Set trm = Nothing
'    If RestoreTRM Then err.Raise knCall, , ksCall
'
'    If AssertEqual(":StartTestCase(DummyTestModule4.Test1):EndTestCase:EndTestSuite", reply) _
'        Then err.Raise knCall, , ksCall
'
'End Sub

Rem =head4 Function ~
Rem sheetname TestShouldRunFixture
Rem
Rem The "ShouldRunFixture" returns a boolean indicating whether a module contains tests that
Rem can be run. This procedure checks whether the replies are as expected.
Rem
Function TestShouldRunFixture() As Boolean
On Error GoTo ErrorHandler

    Dim tr As TestRunner
    Set tr = New TestRunner
    Dim tf As TestFixture
    Set tf = New TestFixture

    If GetNames Then err.Raise knCall, , ksCall
    Dim c As VBComponent
    Set c = Application.VBE.VBProjects(AddInName).VBComponents("DummyTestModule")

    tf.ExtractTestCases Application.VBE.VBProjects(AddInName), c

    If AssertTrue(tr.ShouldRunFixture(tf, "")) Then err.Raise knCall, , ksCall
    If AssertTrue(tr.ShouldRunFixture(tf, "DummyTestModule")) Then err.Raise knCall, , ksCall
    If AssertFalse(tr.ShouldRunFixture(tf, "xxx")) Then err.Raise knCall, , ksCall

Exit Function
ErrorHandler:
ErrTrap ksErrMod, "TestShouldRunFixture"
TestShouldRunFixture = True
End Function
'Public Sub TestShouldRunFixture()
'
'    Dim tr As TestRunner
'    Set tr = New TestRunner
'    Dim tf As TestFixture
'    Set tf = New TestFixture
'
'    If GetNames Then err.Raise knCall, , ksCall
'    Dim c As VBComponent
'    Set c = Application.VBE.VBProjects(AddInName).VBComponents("DummyTestModule")
'
'    tf.ExtractTestCases Application.VBE.VBProjects(AddInName), c
'
'    If AssertTrue(tr.ShouldRunFixture(tf, "")) Then err.Raise knCall, , ksCall
'    If AssertTrue(tr.ShouldRunFixture(tf, "DummyTestModule")) Then err.Raise knCall, , ksCall
'    If AssertFalse(tr.ShouldRunFixture(tf, "xxx")) Then err.Raise knCall, , ksCall
'
'End Sub
