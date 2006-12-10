Attribute VB_Name = "TestResultsManagerTester"
Option Explicit

Public Sub TestLogSuccess()

    Dim trm As ITestResultsManager
    Set trm = New TestResultsManager
    Set trm.testLogger = New FakeDebugTestLogger
    
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
    Set trm.testLogger = logger
    
    trm.LogSuccess
    trm.LogFailure "Failure 1"
    trm.LogFailure "Failure 2"
    trm.LogFailure "Failure 3"
    trm.LogSuccess
    
    AssertEqual 3, trm.TestCaseFailureCount
    
    AssertEqual ":Success: Failure 1 Failed: Failure 2 Failed: Failure 3 Failed:Success", logger.message

End Sub

Public Sub TestTestCase()

    Dim trm As ITestResultsManager
    Set trm = New TestResultsManager
    
    Dim logger As FakeDebugTestLogger
    Set logger = New FakeDebugTestLogger
    Set trm.testLogger = logger
    
    trm.StartTestCase "Test Case 1"
    trm.LogFailure "Failure 1"
    trm.LogFailure "Failure 2"
    trm.LogSuccess
    trm.EndTestCase
    
    AssertEqual 1, trm.TestCaseSuccessCount
    AssertEqual 2, trm.TestCaseFailureCount
    
    trm.StartTestCase "Test Case 2"
    trm.LogFailure "Failure 3"
    trm.LogSuccess
    trm.LogSuccess
    trm.EndTestCase
    
    AssertEqual 2, trm.TestCaseSuccessCount
    AssertEqual 1, trm.TestCaseFailureCount
    
End Sub

Public Sub TestTestFixture()

    Dim trm As ITestResultsManager
    Set trm = New TestResultsManager
    
    Dim logger As FakeDebugTestLogger
    Set logger = New FakeDebugTestLogger
    Set trm.testLogger = logger
    
    trm.StartTestFixture "Fixture 1"
    trm.StartTestCase "Test Case 1"
    trm.LogSuccess
    trm.LogSuccess
    trm.EndTestCase
    
    trm.StartTestCase "Test Case 2"
    trm.LogFailure "Failure 1"
    trm.EndTestCase
    trm.EndTestFixture
    
    AssertEqual 1, trm.FixtureFailureCount
    AssertEqual 2, trm.FixtureSuccessCount
    
    
End Sub

Public Sub TestTestSuite()

    Dim trm As ITestResultsManager
    Set trm = New TestResultsManager
    
    Dim logger As FakeDebugTestLogger
    Set logger = New FakeDebugTestLogger
    Set trm.testLogger = logger
    
        
    trm.StartTestFixture "Fixture 1"
    trm.StartTestCase "Test Case 1"
    trm.LogSuccess
    trm.LogSuccess
    trm.EndTestCase
    
    trm.StartTestCase "Test Case 2"
    trm.LogFailure "Failure 1"
    trm.EndTestCase
    trm.EndTestFixture
    
    trm.EndTestSuite
    
    AssertTrue Right(logger.message, 27) = "Total: 2 passes, 1 failures"

End Sub
