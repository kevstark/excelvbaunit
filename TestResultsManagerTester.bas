Attribute VB_Name = "TestResultsManagerTester"
Option Explicit
Option Private Module
Rem order 9
Rem
Rem =head2
Rem sheetname
Rem
Rem This is part of the test suite for the test harness add-in. Its primary purpose is to test
Rem the functionality of the "TestResultsManager" class module.
Rem
Rem =head3
Rem sheetname Macros
Rem
Const ksErrMod As String = "TestResultsManagerTester"

Rem =head4 Function ~
Rem sheetname TestLogSuccess
Rem
Rem This tests the LogSuccess procedure by calling it and LogFailure & checking that the correct
Rem number of successes have been logged. Most of the logging is identical in TestLogFailure,
Rem making this pair a candidate for refactoring to remove duplication. However, it has been
Rem decided not to do this. The reason is that it is possible to write code to check for
Rem coverage and identify untested code. This will rely on every procedure having an
Rem appropriately named test procedure. Refactoring would put the tests for two procedures
Rem into a single test procedure, making coverage appear incomplete.
Rem
Function TestLogSuccess() As Boolean
On Error GoTo ErrorHandler

    Dim trm As ITestResultsManager
    Set trm = New TestResultsManager
    Set trm.testLogger = New FakeDebugTestLogger
    
    If trm.LogSuccess() Then err.Raise knCall, , ksCall
    If trm.LogFailure("Failure 1") Then err.Raise knCall, , ksCall
    If trm.LogFailure("Failure 2") Then err.Raise knCall, , ksCall
    If trm.LogFailure("Failure 3") Then err.Raise knCall, , ksCall
    If trm.LogSuccess() Then err.Raise knCall, , ksCall
    
    If AssertEqual(2, trm.TestCaseSuccessCount) Then err.Raise knCall, , ksCall

Exit Function
ErrorHandler:
ErrTrap ksErrMod, "TestLogSuccess"
TestLogSuccess = True
End Function
'Public Sub TestLogSuccess()
'
'    Dim trm As ITestResultsManager
'    Set trm = New TestResultsManager
'    Set trm.testLogger = New FakeDebugTestLogger
'
'    trm.LogSuccess
'    trm.LogFailure "Failure 1"
'    trm.LogFailure "Failure 2"
'    trm.LogFailure "Failure 3"
'    trm.LogSuccess
'
'    If AssertEqual(2, trm.TestCaseSuccessCount) Then err.Raise knCall, , ksCall
'
'End Sub

Rem =head4 Function ~
Rem sheetname TestLogFailure
Rem
Rem Tests the "LogFailure" procedure. See above for refactoring idea.
Rem
Function TestLogFailure() As Boolean
On Error GoTo ErrorHandler

    Dim trm As ITestResultsManager
    Set trm = New TestResultsManager
    Dim logger As FakeDebugTestLogger
    Set logger = New FakeDebugTestLogger
    Set trm.testLogger = logger
    
    If trm.LogSuccess() Then err.Raise knCall, , ksCall
    If trm.LogFailure("Failure 1") Then err.Raise knCall, , ksCall
    If trm.LogFailure("Failure 2") Then err.Raise knCall, , ksCall
    If trm.LogFailure("Failure 3") Then err.Raise knCall, , ksCall
    If trm.LogSuccess() Then err.Raise knCall, , ksCall
    
    If AssertEqual(3, trm.TestCaseFailureCount) Then err.Raise knCall, , ksCall
    
    If AssertEqual(":Success: Failure 1 Failed: Failure 2 Failed: Failure 3 Failed:Success", _
                   logger.message) Then err.Raise knCall, , ksCall

Exit Function
ErrorHandler:
ErrTrap ksErrMod, "TestLogFailure"
TestLogFailure = True
End Function
'Public Sub TestLogFailure()
'
'    Dim trm As ITestResultsManager
'    Set trm = New TestResultsManager
'    Dim logger As FakeDebugTestLogger
'    Set logger = New FakeDebugTestLogger
'    Set trm.testLogger = logger
'
'    trm.LogSuccess
'    trm.LogFailure "Failure 1"
'    trm.LogFailure "Failure 2"
'    trm.LogFailure "Failure 3"
'    trm.LogSuccess
'
'    If AssertEqual(3, trm.TestCaseFailureCount) Then err.Raise knCall, , ksCall
'
'    If AssertEqual(":Success: Failure 1 Failed: Failure 2 Failed: Failure 3 Failed:Success", _
'                   logger.message) Then err.Raise knCall, , ksCall
'
'End Sub

Rem =head4 Function ~
Rem sheetname TestTestCase
Rem
Rem A "TestCase" is, for the purpose of this system, a set of tests that exist in a single
Rem procedure. The individual tests may pass or fail, and this procedure calls the start and
Rem end of case routines that would be called automatically. Between these calls, the
Rem procedure generates success and failure events. After cases are ended, the results are
Rem inspected to ensure they are correct.
Rem
Function TestTestCase() As Boolean
On Error GoTo ErrorHandler

    Dim trm As ITestResultsManager
    Set trm = New TestResultsManager
    
    Dim logger As FakeDebugTestLogger
    Set logger = New FakeDebugTestLogger
    Set trm.testLogger = logger
    
    If trm.StartTestCase("Test Case 1") Then err.Raise knCall, , ksCall
    If trm.LogFailure("Failure 1") Then err.Raise knCall, , ksCall
    If trm.LogFailure("Failure 2") Then err.Raise knCall, , ksCall
    If trm.LogSuccess Then err.Raise knCall, , ksCall
    If trm.EndTestCase() Then err.Raise knCall, , ksCall
    
    If AssertEqual(1, trm.TestCaseSuccessCount) Then err.Raise knCall, , ksCall
    If AssertEqual(2, trm.TestCaseFailureCount) Then err.Raise knCall, , ksCall
    
    If trm.StartTestCase("Test Case 2") Then err.Raise knCall, , ksCall
    If trm.LogFailure("Failure 3") Then err.Raise knCall, , ksCall
    If trm.LogSuccess Then err.Raise knCall, , ksCall
    If trm.LogSuccess Then err.Raise knCall, , ksCall
    If trm.EndTestCase() Then err.Raise knCall, , ksCall
    
    If AssertEqual(2, trm.TestCaseSuccessCount) Then err.Raise knCall, , ksCall
    If AssertEqual(1, trm.TestCaseFailureCount) Then err.Raise knCall, , ksCall

Exit Function
ErrorHandler:
ErrTrap ksErrMod, "TestTestCase"
TestTestCase = True
End Function
'Public Sub TestTestCase()
'
'    Dim trm As ITestResultsManager
'    Set trm = New TestResultsManager
'
'    Dim logger As FakeDebugTestLogger
'    Set logger = New FakeDebugTestLogger
'    Set trm.testLogger = logger
'
'    trm.StartTestCase "Test Case 1"
'    trm.LogFailure "Failure 1"
'    trm.LogFailure "Failure 2"
'    trm.LogSuccess
'    trm.EndTestCase
'
'    If AssertEqual(1, trm.TestCaseSuccessCount) Then err.Raise knCall, , ksCall
'    If AssertEqual(2, trm.TestCaseFailureCount) Then err.Raise knCall, , ksCall
'
'    trm.StartTestCase "Test Case 2"
'    trm.LogFailure "Failure 3"
'    trm.LogSuccess
'    trm.LogSuccess
'    trm.EndTestCase
'
'    If AssertEqual(2, trm.TestCaseSuccessCount) Then err.Raise knCall, , ksCall
'    If AssertEqual(1, trm.TestCaseFailureCount) Then err.Raise knCall, , ksCall
'
'End Sub

Rem =head4 Function ~
Rem sheetname TestTestFixture
Rem
Rem A "Test Fixture" is, for the purposes of this system, a module. When all tests in a module
Rem have been run, the system produces summary totals. This procedure simulates the calls that
Rem start and end the processing of a module containing tests and inspects the manager object
Rem to ensure it contains the correct summary information.
Rem
Function TestTestFixture() As Boolean
On Error GoTo ErrorHandler

    Dim trm As ITestResultsManager
    Set trm = New TestResultsManager
    
    Dim logger As FakeDebugTestLogger
    Set logger = New FakeDebugTestLogger
    Set trm.testLogger = logger
    
    If trm.StartTestFixture("Fixture 1") Then err.Raise knCall, , ksCall
    If trm.StartTestCase("Test Case 1") Then err.Raise knCall, , ksCall
    If trm.LogSuccess Then err.Raise knCall, , ksCall
    If trm.LogSuccess Then err.Raise knCall, , ksCall
    If trm.EndTestCase() Then err.Raise knCall, , ksCall
    
    If trm.StartTestCase("Test Case 2") Then err.Raise knCall, , ksCall
    If trm.LogFailure("Failure 1") Then err.Raise knCall, , ksCall
    If trm.EndTestCase() Then err.Raise knCall, , ksCall
    If trm.EndTestFixture() Then err.Raise knCall, , ksCall
    
    If AssertEqual(1, trm.FixtureFailureCount) Then err.Raise knCall, , ksCall
    If AssertEqual(2, trm.FixtureSuccessCount) Then err.Raise knCall, , ksCall

Exit Function
ErrorHandler:
ErrTrap ksErrMod, "TestTestFixture"
TestTestFixture = True
End Function
'Public Sub TestTestFixture()
'
'    Dim trm As ITestResultsManager
'    Set trm = New TestResultsManager
'
'    Dim logger As FakeDebugTestLogger
'    Set logger = New FakeDebugTestLogger
'    Set trm.testLogger = logger
'
'    trm.StartTestFixture "Fixture 1"
'    trm.StartTestCase "Test Case 1"
'    trm.LogSuccess
'    trm.LogSuccess
'    trm.EndTestCase
'
'    trm.StartTestCase "Test Case 2"
'    trm.LogFailure "Failure 1"
'    trm.EndTestCase
'    trm.EndTestFixture
'
'    If AssertEqual(1, trm.FixtureFailureCount) Then err.Raise knCall, , ksCall
'    If AssertEqual(2, trm.FixtureSuccessCount) Then err.Raise knCall, , ksCall
'
'End Sub

Rem =head4 Function ~
Rem sheetname TestTestSuite
Rem
Rem A "Test Suite" is all the modules in a project that contain tests. This procedure emulates
Rem the generation of passes and fails and then inspects the manager object to ensure it
Rem contains the correct summary information.
Rem
Function TestTestSuite() As Boolean
On Error GoTo ErrorHandler

    Dim trm As ITestResultsManager
    Set trm = New TestResultsManager
    
    Dim logger As FakeDebugTestLogger
    Set logger = New FakeDebugTestLogger
    Set trm.testLogger = logger
    
    If trm.StartTestFixture("Fixture 1") Then err.Raise knCall, , ksCall
    If trm.StartTestCase("Test Case 1") Then err.Raise knCall, , ksCall
    If trm.LogSuccess Then err.Raise knCall, , ksCall
    If trm.LogSuccess Then err.Raise knCall, , ksCall
    If trm.EndTestCase() Then err.Raise knCall, , ksCall
    
    If trm.StartTestCase("Test Case 2") Then err.Raise knCall, , ksCall
    If trm.LogFailure("Failure 1") Then err.Raise knCall, , ksCall
    If trm.EndTestCase() Then err.Raise knCall, , ksCall
    If trm.EndTestFixture() Then err.Raise knCall, , ksCall
    
    If trm.EndTestSuite() Then err.Raise knCall, , ksCall
    
    If AssertTrue(Right(logger.message, 27) = "Total: 2 passes, 1 failures") _
    Then err.Raise knCall, , ksCall

Exit Function
ErrorHandler:
ErrTrap ksErrMod, "TestTestSuite"
TestTestSuite = True
End Function
'Public Sub TestTestSuite()
'
'    Dim trm As ITestResultsManager
'    Set trm = New TestResultsManager
'
'    Dim logger As FakeDebugTestLogger
'    Set logger = New FakeDebugTestLogger
'    Set trm.testLogger = logger
'
'
'    trm.StartTestFixture "Fixture 1"
'    trm.StartTestCase "Test Case 1"
'    trm.LogSuccess
'    trm.LogSuccess
'    trm.EndTestCase
'
'    trm.StartTestCase "Test Case 2"
'    trm.LogFailure "Failure 1"
'    trm.EndTestCase
'    trm.EndTestFixture
'
'    trm.EndTestSuite
'
'    If AssertTrue(Right(logger.message, 27) = "Total: 2 passes, 1 failures") _
'    Then err.Raise knCall, , ksCall
'
'End Sub
