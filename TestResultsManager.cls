VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestResultsManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements ITestResultsManager

' Keeps track of the test results for a test run

Private mTotalSuccessCount As Long
Private mTotalFailureCount As Long

Private mFixtureSuccessCount As Long
Private mFixtureFailureCount As Long

Private mTestCaseSuccessCount As Long
Private mTestCaseFailureCount As Long

Private mTestCaseName As String
Private mTestFixtureName As String

Private mTestLogger As ITestLogger


' By default, log progress to the immediate window
Private Sub Class_Initialize()
    Set mTestLogger = New DebugTestLogger
End Sub



Public Property Let ITestResultsManager_ExpectedNumTestCases(count As Long)

    mTestLogger.ExpectedNumTestCases = count
    
End Property


' Override the standard test logger, eg to log to a gui
Property Set ITestResultsManager_TestLogger(logger As ITestLogger)
    Set mTestLogger = logger
End Property



Property Get ITestResultsManager_TotalSuccessCount() As Integer
    ITestResultsManager_TotalSuccessCount = mTotalSuccessCount
End Property



Property Get ITestResultsManager_TotalFailureCount() As Integer
    ITestResultsManager_TotalFailureCount = mTotalFailureCount
End Property



Property Get ITestResultsManager_FixtureSuccessCount() As Integer
    ITestResultsManager_FixtureSuccessCount = mFixtureSuccessCount
End Property



Property Get ITestResultsManager_FixtureFailureCount() As Integer
    ITestResultsManager_FixtureFailureCount = mFixtureFailureCount
End Property



Property Get ITestResultsManager_TestCaseSuccessCount() As Integer
    ITestResultsManager_TestCaseSuccessCount = mTestCaseSuccessCount
End Property



Property Get ITestResultsManager_TestCaseFailureCount() As Integer
    ITestResultsManager_TestCaseFailureCount = mTestCaseFailureCount
End Property



Public Sub ITestResultsManager_LogSuccess()

    mTotalSuccessCount = mTotalSuccessCount + 1
    mFixtureSuccessCount = mFixtureSuccessCount + 1
    mTestCaseSuccessCount = mTestCaseSuccessCount + 1
    
    mTestLogger.LogSuccess
    
End Sub



Public Sub ITestResultsManager_LogFailure(msg As String)

    mTotalFailureCount = mTotalFailureCount + 1
    mFixtureFailureCount = mFixtureFailureCount + 1
    mTestCaseFailureCount = mTestCaseFailureCount + 1
    
    mTestLogger.LogFailure mTestCaseName, msg

End Sub



Public Sub ITestResultsManager_StartTestFixture(name As String)

    mTestLogger.StartTestFixture name
    mTestCaseSuccessCount = 0
    mTestCaseFailureCount = 0
    mFixtureSuccessCount = 0
    mFixtureFailureCount = 0
    

End Sub



Public Sub ITestResultsManager_EndTestFixture()

    mTestLogger.EndTestFixture mFixtureSuccessCount, mFixtureFailureCount

End Sub



Public Sub ITestResultsManager_StartTestCase(name As String)

    mTestLogger.StartTestCase name
    
    mTestCaseName = name
    mTestCaseSuccessCount = 0
    mTestCaseFailureCount = 0
    
End Sub



Public Sub ITestResultsManager_EndTestCase()

    mTestLogger.EndTestCase mTestCaseName, mTestCaseSuccessCount, mTestCaseFailureCount
    
End Sub



Public Sub ITestResultsManager_EndTestSuite()

    mTestLogger.EndTestSuite mTotalSuccessCount, mTotalFailureCount

End Sub

