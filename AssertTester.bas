Attribute VB_Name = "AssertTester"
Option Explicit
Rem order 4.1
Rem
Rem =head2
Rem sheetname
Rem
Rem Retrofitted by JHD to test Assert commands. Three procedures - SetTestResultsManager,
Rem BackupTRM and RestoreTRM - are not tested, as no obvious means can be seen to do this.
Rem Other procedures are tested, but because of the nature of the procedures being tested, the
Rem number of tests is under-reported. This is because two test managers are needed. It is
Rem necessary to test that assert procedures process failures correctly. Therefore a second
Rem test manager is used to record how many passes and failures there are. These counts are
Rem then passed to the testing procedure which uses the main TRM to test that the counts are as
Rem expected. Thus, in the case of AssertEqual, there are four tests for failure and four for
Rem success. But the main TRM is only asked to check that both these numbers are 4, meaning
Rem that the total of eight tests looks as though there are only two. There will also be reports
Rem of failures in the log. These have been annotated as intentional failures, but may look
Rem strange at first sight.
Rem
Rem =head3
Rem sheetname Macros
Rem
Const ksErrMod As String = "AssertTester"

Rem =head4 Function TestAssertTrue
Rem
Rem 2 tests reported as 2.
Rem
Function TestAssertTrue()
On Error GoTo ErrorHandler

    Dim trmTest As TestResultsManager
    Dim nTRMTrue As Long
    Dim nTRMFalse As Long
    Dim tlTest As DebugTestLogger
    Set tlTest = New DebugTestLogger
    If BackupTRM Then err.Raise knCall, , ksCall
    Set trmTest = New TestResultsManager
    Set trmTest.ITestResultsManager_TestLogger = tlTest
    If SetTestResultsManager(trmTest) Then err.Raise knCall, , ksCall
    If AssertTrue(True) Then err.Raise knCall, , ksCall
    If AssertTrue(False, "Intentional failure") Then err.Raise knCall, , ksCall
    If RestoreTRM Then err.Raise knCall, , ksCall
    nTRMTrue = trmTest.ITestResultsManager_TestCaseSuccessCount
    nTRMFalse = trmTest.ITestResultsManager_TestCaseFailureCount
    If AssertEqual(1, nTRMTrue) Then err.Raise knCall, , ksCall
    If AssertEqual(1, nTRMFalse) Then err.Raise knCall, , ksCall
    Set trmTest = Nothing
    Set tlTest = Nothing
    
Exit Function
ErrorHandler:
ErrTrap ksErrMod, "TestAssertTrue"
TestAssertTrue = True
End Function
Rem =head4 Function TestAssertFalse
Rem
Rem 2 tests reported as 2.
Rem
Function TestAssertFalse()
On Error GoTo ErrorHandler

    Dim trmTest As TestResultsManager
    Dim nTRMTrue As Long
    Dim nTRMFalse As Long
    Dim tlTest As DebugTestLogger
    Set tlTest = New DebugTestLogger
    If BackupTRM Then err.Raise knCall, , ksCall
    Set trmTest = New TestResultsManager
    Set trmTest.ITestResultsManager_TestLogger = tlTest
    If SetTestResultsManager(trmTest) Then err.Raise knCall, , ksCall
    If AssertFalse(False) Then err.Raise knCall, , ksCall
    If AssertFalse(True, "Intentional failure") Then err.Raise knCall, , ksCall
    If RestoreTRM Then err.Raise knCall, , ksCall
    nTRMTrue = trmTest.ITestResultsManager_TestCaseSuccessCount
    nTRMFalse = trmTest.ITestResultsManager_TestCaseFailureCount
    If AssertEqual(1, nTRMTrue) Then err.Raise knCall, , ksCall
    If AssertEqual(1, nTRMFalse) Then err.Raise knCall, , ksCall
    Set trmTest = Nothing
    Set tlTest = Nothing
    
Exit Function
ErrorHandler:
ErrTrap ksErrMod, "TestAssertFalse"
TestAssertFalse = True
End Function
Rem =head4 Function TestAssertEqual
Rem
Rem 8 tests reported as 2.
Rem
Function TestAssertEqual()
On Error GoTo ErrorHandler
    
    Dim trmTest As TestResultsManager
    Dim nTRMTrue As Long
    Dim nTRMFalse As Long
    Dim tlTest As DebugTestLogger
    Set tlTest = New DebugTestLogger
    If BackupTRM Then err.Raise knCall, , ksCall
    Set trmTest = New TestResultsManager
    Set trmTest.ITestResultsManager_TestLogger = tlTest
    If SetTestResultsManager(trmTest) Then err.Raise knCall, , ksCall

    If AssertEqual("a", "a", "String pass") Then err.Raise knCall, , ksCall
    If AssertEqual("a", "ab", "String fail - intentional") Then err.Raise knCall, , ksCall
    If AssertEqual(1, 1, "Int pass") Then err.Raise knCall, , ksCall
    If AssertEqual(1, 2, "Int Fail - intentional") Then err.Raise knCall, , ksCall
    If AssertEqual(1.1, 1.1, "Simple float pass") Then err.Raise knCall, , ksCall
    If AssertEqual(1.1, 1.2, "Float fail - intentional") Then err.Raise knCall, , ksCall
    Dim a As Double, b As Double, c As Double
    a = 1 / 900
    b = 1 + a
    c = b - 1
    If AssertEqual(a, c, "Tough float pass") Then err.Raise knCall, , ksCall
    If AssertEqual("a", a, "Mixed type fail - intentional") Then err.Raise knCall, , ksCall
    
    If RestoreTRM Then err.Raise knCall, , ksCall
    nTRMTrue = trmTest.ITestResultsManager_TestCaseSuccessCount
    nTRMFalse = trmTest.ITestResultsManager_TestCaseFailureCount
    If AssertEqual(4, nTRMTrue) Then err.Raise knCall, , ksCall
    If AssertEqual(4, nTRMFalse) Then err.Raise knCall, , ksCall
    Set trmTest = Nothing
    Set tlTest = Nothing

Exit Function
ErrorHandler:
ErrTrap ksErrMod, "TestAssertEqual"
TestAssertEqual = True
End Function
Rem =head4 Function TestAssertSuccess
Rem
Rem 1 test.
Rem
Function TestAssertSuccess()
On Error GoTo ErrorHandler

    Dim trmTest As TestResultsManager
    Dim nTRMTrue As Long
    Dim tlTest As DebugTestLogger
    Set tlTest = New DebugTestLogger
    If BackupTRM Then err.Raise knCall, , ksCall
    Set trmTest = New TestResultsManager
    Set trmTest.ITestResultsManager_TestLogger = tlTest
    If SetTestResultsManager(trmTest) Then err.Raise knCall, , ksCall
    
    If AssertSuccess Then err.Raise knCall, , ksCall

    If RestoreTRM Then err.Raise knCall, , ksCall
    nTRMTrue = trmTest.ITestResultsManager_TestCaseSuccessCount
    If AssertEqual(1, nTRMTrue) Then err.Raise knCall, , ksCall
    Set trmTest = Nothing
    Set tlTest = Nothing

Exit Function
ErrorHandler:
ErrTrap ksErrMod, "TestAssertSuccess"
TestAssertSuccess = True
End Function
Rem =head4 Function TestAssertFailure
Rem
Rem 1 test.
Rem
Function TestAssertFailure()
On Error GoTo ErrorHandler

    Dim trmTest As TestResultsManager
    Dim nTRMFalse As Long
    Dim tlTest As DebugTestLogger
    Set tlTest = New DebugTestLogger
    If BackupTRM Then err.Raise knCall, , ksCall
    Set trmTest = New TestResultsManager
    Set trmTest.ITestResultsManager_TestLogger = tlTest
    If SetTestResultsManager(trmTest) Then err.Raise knCall, , ksCall

    If AssertFailure("Intentional failure") Then err.Raise knCall, , ksCall
    
    If RestoreTRM Then err.Raise knCall, , ksCall
    nTRMFalse = trmTest.ITestResultsManager_TestCaseFailureCount
    If AssertEqual(1, nTRMFalse) Then err.Raise knCall, , ksCall
    Set trmTest = Nothing
    Set tlTest = Nothing

Exit Function
ErrorHandler:
ErrTrap ksErrMod, "TestAssertFailure"
TestAssertFailure = True
End Function
