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

Private mTotalSuccessCount As Integer
Private mTotalFailureCount As Integer

Private mFixtureSuccessCount As Integer
Private mFixtureFailureCount As Integer

Private mTestCaseSuccessCount As Integer
Private mTestCaseFailureCount As Integer

Private mTestCaseName As String
Private mTestFixtureName As String

Private mTestLogger As ITestLogger


Private Sub Class_Initialize()
    Set mTestLogger = New DebugTestLogger
End Sub


' Override the standard test logger to ease testing this class
Property Set TestLogger(logger As ITestLogger)
    Set mTestLogger = logger
End Property

Property Get TotalSuccessCount() As Integer
    TotalSuccessCount = mTotalSuccessCount
End Property

Property Get TotalFailureCount() As Integer
    TotalFailureCount = mTotalFailureCount
End Property

Property Get FixtureSuccessCount() As Integer
    FixtureSuccessCount = mFixtureSuccessCount
End Property

Property Get FixtureFailureCount() As Integer
    FixtureFailureCount = mFixtureFailureCount
End Property

Property Get TestCaseSuccessCount() As Integer
    TestCaseSuccessCount = mTestCaseSuccessCount
End Property

Property Get TestCaseFailureCount() As Integer
    TestCaseFailureCount = mTestCaseFailureCount
End Property

Public Sub LogSuccess()

    mTotalSuccessCount = mTotalSuccessCount + 1
    mFixtureSuccessCount = mFixtureSuccessCount + 1
    mTestCaseSuccessCount = mTestCaseSuccessCount + 1
    
End Sub

Public Sub LogFailure(msg As String)

    mTotalFailureCount = mTotalFailureCount + 1
    mFixtureFailureCount = mFixtureFailureCount + 1
    mTestCaseFailureCount = mTestCaseFailureCount + 1
    
    mTestLogger.LogMsg mTestCaseName & " failed. " & msg

End Sub

Public Sub StartTestFixture(name As String)

    mTestLogger.LogMsg "Start: " & name
    mTestCaseSuccessCount = 0
    mTestCaseFailureCount = 0
    mFixtureSuccessCount = 0
    mFixtureFailureCount = 0
    

End Sub

Public Sub EndTestFixture()

    mTestLogger.LogMsg mFixtureSuccessCount & " passes, " & mFixtureFailureCount & " failures"
    mTestLogger.LogMsg "==================="
    mTestLogger.LogMsg

End Sub

Public Sub StartTestCase(name As String)

    mTestCaseName = name
    mTestCaseSuccessCount = 0
    mTestCaseFailureCount = 0
    
End Sub

Public Sub EndTestCase()

    mTestLogger.LogMsg mTestCaseName & ": " & mTestCaseSuccessCount & " passes, " & mTestCaseFailureCount & " failures"
    
End Sub

Public Sub EndTestSuite()

    mTestLogger.LogMsg "Total: " & mTotalSuccessCount & " passes, " & mTotalFailureCount & " failures"

End Sub

