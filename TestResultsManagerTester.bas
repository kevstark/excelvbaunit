Attribute VB_Name = "TestResultsManagerTester"
Option Explicit

Public Sub TestLogSuccess()

    Dim trm As ITestResultsManager
    Set trm = New TestResultsManager
    Set trm.TestLogger = New FakeDebugTestLogger
    
    trm.LogSuccess
    trm.LogFailure "Failure 1"
    trm.LogFailure "Failure 2"
    trm.LogFailure "Failure 3"
    trm.LogSuccess
    
    AssertEqual 2, trm.TestCaseSuccessCount

End Sub


Public Sub TestLogFailure()

    Dim trm As ITestResultsManager
    Set trm = New TestResultsManager
    
    Dim logger As FakeDebugTestLogger
    Set logger = New FakeDebugTestLogger
    Set trm.TestLogger = logger
    
    trm.LogSuccess
    trm.LogFailure "Failure 1"
    trm.LogFailure "Failure 2"
    trm.LogFailure "Failure 3"
    trm.LogSuccess
    
    AssertEqual 3, trm.TestCaseFailureCount
    
    AssertEqual ": failed. Failure 1: failed. Failure 2: failed. Failure 3", logger.message

End Sub

