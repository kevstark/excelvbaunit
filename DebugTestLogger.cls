VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DebugTestLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements ITestLogger

' Logs progress to the immediate window
Public Sub ITestLogger_LogMsg(Optional msg As String = "")

    Debug.Print msg
    
End Sub



Public Sub ITestLogger_LogSuccess()
End Sub



Public Sub ITestLogger_LogFailure(name As String, msg As String)

    ITestLogger_LogMsg name & " failed. " & msg
    
End Sub



Public Sub ITestLogger_StartTestFixture(name As String)

    ITestLogger_LogMsg "Start: " & name
    
End Sub



Public Sub ITestLogger_EndTestFixture(successCount As Long, failureCount As Long)

    ITestLogger_LogMsg successCount & " passes, " & failureCount & " failures"
    ITestLogger_LogMsg "==================="
    ITestLogger_LogMsg
    
End Sub



Public Sub ITestLogger_StartTestCase(name As String)
End Sub



Public Sub ITestLogger_EndTestCase(name As String, successCount As Long, failureCount As Long)

    ITestLogger_LogMsg name & ": " & successCount & " passes, " & failureCount & " failures"

End Sub

Public Sub ITestLogger_EndTestSuite(successCount As Long, failureCount As Long)

    ITestLogger_LogMsg "Total: " & successCount & " passes, " & failureCount & " failures"

End Sub


Property Let ITestLogger_ExpectedNumTestCases(count As Long)

End Property

