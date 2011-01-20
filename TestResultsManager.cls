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
Rem order 15
Rem
Rem =head2
Rem sheetname
Rem
Rem rcl ClassLinks
Rem
Rem =head3
Rem sheetname Macros
Rem
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
Const ksErrMod As String = "TestResultsManager"

' By default, log progress to the immediate window
Rem =head4 Sub ~
Rem sheetname Class_Initialize
Rem
Rem Excel requires this to be written as a sub.
Rem
Private Sub Class_Initialize()
    Set mTestLogger = New DebugTestLogger
End Sub

'Rem =head4 Function
'Rem
'Rem
'Rem
'Rem
'Rem rcl ToFn
'Rem
'Public Property Let ITestResultsManager_ExpectedNumTestCases(count As Long)
'
'    mTestLogger.ExpectedNumTestCases = count
'
'End Property

' Override the standard test logger, eg to log to a gui
Rem =head4 Property Set ~
Rem sheetname ITestResultsManager_TestLogger
Rem
Property Set ITestResultsManager_TestLogger(logger As ITestLogger)
    Set mTestLogger = logger
End Property

Rem =head4 Property Get ~
Rem sheetname ITestResultsManager_TotalSuccessCount
Rem
Property Get ITestResultsManager_TotalSuccessCount() As Integer
    ITestResultsManager_TotalSuccessCount = mTotalSuccessCount
End Property

Rem =head4 Property Get ~
Rem sheetname ITestResultsManager_TotalFailureCount
Rem
Property Get ITestResultsManager_TotalFailureCount() As Integer
    ITestResultsManager_TotalFailureCount = mTotalFailureCount
End Property

Rem =head4 Property Get ~
Rem sheetname ITestResultsManager_FixtureSuccessCount
Rem
Property Get ITestResultsManager_FixtureSuccessCount() As Integer
    ITestResultsManager_FixtureSuccessCount = mFixtureSuccessCount
End Property

Rem =head4 Property Get ~
Rem sheetname ITestResultsManager_FixtureFailureCount
Rem
Property Get ITestResultsManager_FixtureFailureCount() As Integer
    ITestResultsManager_FixtureFailureCount = mFixtureFailureCount
End Property

Rem =head4 Property Get ~
Rem sheetname ITestResultsManager_TestCaseSuccessCount
Rem
Property Get ITestResultsManager_TestCaseSuccessCount() As Integer
    ITestResultsManager_TestCaseSuccessCount = mTestCaseSuccessCount
End Property

Rem =head4 Property Get ~
Rem sheetname ITestResultsManager_TestCaseFailureCount
Rem
Property Get ITestResultsManager_TestCaseFailureCount() As Integer
    ITestResultsManager_TestCaseFailureCount = mTestCaseFailureCount
End Property

Rem =head4 Function ~
Rem sheetname ITestResultsManager_LogSuccess
Rem
Rem Increments success counts and calls TestLogger's LogSuccess procedure.
Rem
Function ITestResultsManager_LogSuccess() As Boolean
On Error GoTo ErrorHandler

    mTotalSuccessCount = mTotalSuccessCount + 1
    mFixtureSuccessCount = mFixtureSuccessCount + 1
    mTestCaseSuccessCount = mTestCaseSuccessCount + 1
    
    If mTestLogger.LogSuccess Then err.Raise knCall, , ksCall

Exit Function
ErrorHandler:
ErrTrap ksErrMod, "ITestResultsManager_LogSuccess"
ITestResultsManager_LogSuccess = True
End Function
'Public Sub ITestResultsManager_LogSuccess()
'
'    mTotalSuccessCount = mTotalSuccessCount + 1
'    mFixtureSuccessCount = mFixtureSuccessCount + 1
'    mTestCaseSuccessCount = mTestCaseSuccessCount + 1
'
'    If mTestLogger.LogSuccess Then err.Raise knCall, , ksCall
'
'End Sub

Rem =head4 Function ~
Rem sheetname ITestResultsManager_LogFailure
Rem
Rem Increments failure counts and calls TestLogger's LogFailure procedure.
Rem
Function ITestResultsManager_LogFailure(msg As String) As Boolean
On Error GoTo ErrorHandler

    mTotalFailureCount = mTotalFailureCount + 1
    mFixtureFailureCount = mFixtureFailureCount + 1
    mTestCaseFailureCount = mTestCaseFailureCount + 1
    
    If mTestLogger.LogFailure(mTestCaseName, msg) Then err.Raise knCall, , ksCall

Exit Function
ErrorHandler:
ErrTrap ksErrMod, "ITestResultsManager_LogFailure"
ITestResultsManager_LogFailure = True
End Function
'Public Sub ITestResultsManager_LogFailure(msg As String)
'
'    mTotalFailureCount = mTotalFailureCount + 1
'    mFixtureFailureCount = mFixtureFailureCount + 1
'    mTestCaseFailureCount = mTestCaseFailureCount + 1
'
'    mTestLogger.LogFailure mTestCaseName, msg
'
'End Sub

Rem =head4 Function ~
Rem sheetname ITestResultsManager_StartTestFixture
Rem
Rem Sets counts to zero.
Rem
Function ITestResultsManager_StartTestFixture(name As String) As Boolean
On Error GoTo ErrorHandler

    If mTestLogger.StartTestFixture(name) Then err.Raise knCall, , ksCall
    mTestCaseSuccessCount = 0
    mTestCaseFailureCount = 0
    mFixtureSuccessCount = 0
    mFixtureFailureCount = 0

Exit Function
ErrorHandler:
ErrTrap ksErrMod, "ITestResultsManager_StartTestFixture"
ITestResultsManager_StartTestFixture = True
End Function
'Public Sub ITestResultsManager_StartTestFixture(name As String)
'
'    mTestLogger.StartTestFixture name
'    mTestCaseSuccessCount = 0
'    mTestCaseFailureCount = 0
'    mFixtureSuccessCount = 0
'    mFixtureFailureCount = 0
'
'End Sub

Rem =head4 Function ~
Rem sheetname ITestResultsManager_EndTestFixture
Rem
Rem Calls the logger's EndTestFixture procedure.
Rem
Function ITestResultsManager_EndTestFixture() As Boolean
On Error GoTo ErrorHandler

    mTestLogger.EndTestFixture mFixtureSuccessCount, mFixtureFailureCount

Exit Function
ErrorHandler:
ErrTrap ksErrMod, "ITestResultsManager_EndTestFixture"
ITestResultsManager_EndTestFixture = True
End Function
'Public Sub ITestResultsManager_EndTestFixture()
'
'    mTestLogger.EndTestFixture mFixtureSuccessCount, mFixtureFailureCount
'
'End Sub

Rem =head4 Function ~
Rem sheetname ITestResultsManager_StartTestCase
Rem
Rem Sets all case counts to zero.
Rem
Function ITestResultsManager_StartTestCase(name As String) As Boolean
On Error GoTo ErrorHandler

    mTestCaseName = name
    mTestCaseSuccessCount = 0
    mTestCaseFailureCount = 0

Exit Function
ErrorHandler:
ErrTrap ksErrMod, "ITestResultsManager_StartTestCase"
ITestResultsManager_StartTestCase = True
End Function
'Public Sub ITestResultsManager_StartTestCase(name As String)
'
'    'mTestLogger.StartTestCase name 'JHD: Commented out because all StartTestCase
'                                    'routines are stubs in *TestLogger modules
'    mTestCaseName = name
'    mTestCaseSuccessCount = 0
'    mTestCaseFailureCount = 0
'
'End Sub

Rem =head4 Function ~
Rem sheetname ITestResultsManager_EndTestCase
Rem
Rem Calls the logger's EndTestCase procedure.
Rem
Function ITestResultsManager_EndTestCase() As Boolean
On Error GoTo ErrorHandler

    If mTestLogger.EndTestCase(mTestCaseName, mTestCaseSuccessCount, mTestCaseFailureCount) _
        Then err.Raise knCall, , ksCall

Exit Function
ErrorHandler:
ErrTrap ksErrMod, "ITestResultsManager_EndTestCase"
ITestResultsManager_EndTestCase = True
End Function
'Public Sub ITestResultsManager_EndTestCase()
'
'    mTestLogger.EndTestCase mTestCaseName, mTestCaseSuccessCount, mTestCaseFailureCount
'
'End Sub

Rem =head4 Function ~
Rem sheetname ITestResultsManager_EndTestSuite
Rem
Rem Calls the logger's EndTestSuite procedure.
Rem
Function ITestResultsManager_EndTestSuite() As Boolean
On Error GoTo ErrorHandler

    If mTestLogger.EndTestSuite(mTotalSuccessCount, mTotalFailureCount) _
        Then err.Raise knCall, , ksCall

Exit Function
ErrorHandler:
ErrTrap ksErrMod, "ITestResultsManager_EndTestSuite"
ITestResultsManager_EndTestSuite = True
End Function
'Public Sub ITestResultsManager_EndTestSuite()
'
'    mTestLogger.EndTestSuite mTotalSuccessCount, mTotalFailureCount
'
'End Sub
