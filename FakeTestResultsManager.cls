VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FakeTestResultsManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Rem order 16
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

Public FunctionsCalled As String
Const ksErrMod As String = "FakeTestResultsManager"

Rem =head4 Property Set ~
Rem sheetname ITestResultsManager_TestLogger
Rem
Property Set ITestResultsManager_TestLogger(logger As ITestLogger)
End Property

Rem =head4 Property Get ~
Rem sheetname ITestResultsManager_TotalSuccessCount
Rem
Property Get ITestResultsManager_TotalSuccessCount() As Integer
End Property

Rem =head4 Property Get ~
Rem sheetname ITestResultsManager_TotalFailureCount
Rem
Property Get ITestResultsManager_TotalFailureCount() As Integer
End Property

Rem =head4 Property Get ~
Rem sheetname ITestResultsManager_FixtureSuccessCount
Rem
Property Get ITestResultsManager_FixtureSuccessCount() As Integer
End Property

Rem =head4 Property Get ~
Rem sheetname ITestResultsManager_FixtureFailureCount
Rem
Property Get ITestResultsManager_FixtureFailureCount() As Integer
End Property

Rem =head4 Property Get ~
Rem sheetname ITestResultsManager_TestCaseSuccessCount
Rem
Property Get ITestResultsManager_TestCaseSuccessCount() As Integer
End Property

Rem =head4 Property Get ~
Rem sheetname ITestResultsManager_TestCaseFailureCount
Rem
Property Get ITestResultsManager_TestCaseFailureCount() As Integer
End Property

Rem =head4 Function ~
Rem sheetname ITestResultsManager_LogSuccess
Rem
Function ITestResultsManager_LogSuccess() As Boolean
On Error GoTo ErrorHandler

Exit Function
ErrorHandler:
ErrTrap ksErrMod, "ITestResultsManager_LogSuccess"
ITestResultsManager_LogSuccess = True
End Function
'Public Sub ITestResultsManager_LogSuccess()
'End Sub

Rem =head4 Function ~
Rem sheetname ITestResultsManager_LogFailure
Rem
Function ITestResultsManager_LogFailure(msg As String) As Boolean
On Error GoTo ErrorHandler

Exit Function
ErrorHandler:
ErrTrap ksErrMod, "ITestResultsManager_LogFailure"
ITestResultsManager_LogFailure = True
End Function
'Public Sub ITestResultsManager_LogFailure(msg As String)
'End Sub

Rem =head4 Function ~
Rem sheetname ITestResultsManager_StartTestFixture
Rem
Function ITestResultsManager_StartTestFixture(name As String) As Boolean
On Error GoTo ErrorHandler

Exit Function
ErrorHandler:
ErrTrap ksErrMod, "ITestResultsManager_StartTestFixture"
ITestResultsManager_StartTestFixture = True
End Function
'Public Sub ITestResultsManager_StartTestFixture(name As String)
'End Sub

Rem =head4 Function ~
Rem sheetname ITestResultsManager_EndTestFixture
Rem
Function ITestResultsManager_EndTestFixture() As Boolean
On Error GoTo ErrorHandler

Exit Function
ErrorHandler:
ErrTrap ksErrMod, "ITestResultsManager_EndTestFixture"
ITestResultsManager_EndTestFixture = True
End Function
'Public Sub ITestResultsManager_EndTestFixture()
'End Sub

Rem =head4 Function ~
Rem sheetname ITestResultsManager_StartTestCase
Rem
Rem Appends text to a string.
Rem
Function ITestResultsManager_StartTestCase(name As String) As Boolean
On Error GoTo ErrorHandler

    FunctionsCalled = FunctionsCalled & ":StartTestCase(" & name & ")"

Exit Function
ErrorHandler:
ErrTrap ksErrMod, "ITestResultsManager_StartTestCase"
ITestResultsManager_StartTestCase = True
End Function
'Public Sub ITestResultsManager_StartTestCase(name As String)
'    FunctionsCalled = FunctionsCalled & ":StartTestCase(" & name & ")"
'End Sub

Rem =head4 Function ~
Rem sheetname ITestResultsManager_EndTestCase
Rem
Rem Appends text to a string.
Rem
Function ITestResultsManager_EndTestCase() As Boolean
On Error GoTo ErrorHandler

    FunctionsCalled = FunctionsCalled & ":EndTestCase"

Exit Function
ErrorHandler:
ErrTrap ksErrMod, "ITestResultsManager_EndTestCase"
ITestResultsManager_EndTestCase = True
End Function
'Public Sub ITestResultsManager_EndTestCase()
'    FunctionsCalled = FunctionsCalled & ":EndTestCase"
'End Sub

Rem =head4 Function ~
Rem sheetname ITestResultsManager_EndTestSuite
Rem
Rem Appends text to a string.
Rem
Function ITestResultsManager_EndTestSuite() As Boolean
On Error GoTo ErrorHandler

    FunctionsCalled = FunctionsCalled & ":EndTestSuite"

Exit Function
ErrorHandler:
ErrTrap ksErrMod, "ITestResultsManager_EndTestSuite"
ITestResultsManager_EndTestSuite = True
End Function
'Public Sub ITestResultsManager_EndTestSuite()
'    FunctionsCalled = FunctionsCalled & ":EndTestSuite"
'End Sub

'Rem =head4 Function
'Rem
'Rem
'Rem
'Rem
'Rem rcl ToFn
'Rem
'Public Property Let ITestResultsManager_ExpectedNumTestCases(count As Long)
'End Property
