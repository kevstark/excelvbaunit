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
Implements ITestResultsManager

Public FunctionsCalled As String

Property Set ITestResultsManager_TestLogger(logger As ITestLogger)
End Property

Property Get ITestResultsManager_TotalSuccessCount() As Integer
End Property

Property Get ITestResultsManager_TotalFailureCount() As Integer
End Property

Property Get ITestResultsManager_FixtureSuccessCount() As Integer
End Property

Property Get ITestResultsManager_FixtureFailureCount() As Integer
End Property

Property Get ITestResultsManager_TestCaseSuccessCount() As Integer
End Property

Property Get ITestResultsManager_TestCaseFailureCount() As Integer
End Property

Public Sub ITestResultsManager_LogSuccess()
End Sub

Public Sub ITestResultsManager_LogFailure(msg As String)
End Sub

Public Sub ITestResultsManager_StartTestFixture(name As String)
End Sub

Public Sub ITestResultsManager_EndTestFixture()
End Sub

Public Sub ITestResultsManager_StartTestCase(name As String)
    FunctionsCalled = FunctionsCalled & ":StartTestCase(" & name & ")"
End Sub

Public Sub ITestResultsManager_EndTestCase()
    FunctionsCalled = FunctionsCalled & ":EndTestCase"
End Sub

Public Sub ITestResultsManager_EndTestSuite()
    FunctionsCalled = FunctionsCalled & ":EndTestSuite"
End Sub

