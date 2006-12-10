VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ITestLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' Interface to logging progress of the test run
Public Sub LogMsg(Optional msg As String = "")
End Sub

Public Sub LogSuccess()
End Sub

Public Sub LogFailure(name As String, msg As String)
End Sub

Public Sub StartTestFixture(name As String)
End Sub

Public Sub EndTestFixture(successCount As Long, failureCount As Long)
End Sub

Public Sub StartTestCase(name As String)
End Sub

Public Sub EndTestCase(name As String, successCount As Long, failureCount As Long)
End Sub

Public Sub EndTestSuite(successCount As Long, failureCount As Long)
End Sub


Property Let ExpectedNumTestCases(count As Long)
End Property

