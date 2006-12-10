VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ITestResultsManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Property Set testLogger(logger As ITestLogger)
End Property

Property Get TotalSuccessCount() As Integer
End Property

Property Get TotalFailureCount() As Integer
End Property

Property Get FixtureSuccessCount() As Integer
End Property

Property Get FixtureFailureCount() As Integer
End Property

Property Get TestCaseSuccessCount() As Integer
End Property

Property Get TestCaseFailureCount() As Integer
End Property

Public Sub LogSuccess()
End Sub

Public Sub LogFailure(msg As String)
End Sub

Public Sub StartTestFixture(name As String)
End Sub

Public Sub EndTestFixture()
End Sub

Public Sub StartTestCase(name As String)
End Sub

Public Sub EndTestCase()
End Sub

Public Sub EndTestSuite()
End Sub


Public Property Let ExpectedNumTestCases(count As Long)
End Property
