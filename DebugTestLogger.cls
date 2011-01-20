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
Rem order 12
Rem
Rem =head2
Rem sheetname
Rem
Rem This is the live implementation of the test logger.
Rem
Rem =head3
Rem sheetname Macros
Rem
Implements ITestLogger
Const ksErrMod As String = "DebugTestLogger"

Rem =head4 Function ~
Rem sheetname ITestLogger_LogMsg
Rem
Rem Outputs the string parameter it receives (if any) to the Immediate pane.
Rem
' Logs progress to the immediate window
Function ITestLogger_LogMsg(Optional msg As String = "") As Boolean
On Error GoTo ErrorHandler

    Debug.Print msg

Exit Function
ErrorHandler:
ErrTrap ksErrMod, "ITestLogger_LogMsg"
ITestLogger_LogMsg = True
End Function
'Public Sub ITestLogger_LogMsg(Optional msg As String = "")
'
'    Debug.Print msg
'
'End Sub

Rem =head4 Function ~
Rem sheetname ITestLogger_LogSuccess
Rem
Rem sto stub
Rem Does nothing. Needed because the Fake equivalent takes action, so there must be a
Rem counterpart here.
Rem sto 0
Rem rcl stub
Rem
Function ITestLogger_LogSuccess() As Boolean
On Error GoTo ErrorHandler

Exit Function
ErrorHandler:
ErrTrap ksErrMod, "ITestLogger_LogSuccess"
ITestLogger_LogSuccess = True
End Function
'Public Sub ITestLogger_LogSuccess()
'End Sub

Rem =head4 Function ~
Rem sheetname ITestLogger_LogFailure
Rem
Rem Passes the calling name, the word "failed" and any message to the LogMsg procedure.
Rem
Function ITestLogger_LogFailure(name As String, msg As String) As Boolean
On Error GoTo ErrorHandler

    If ITestLogger_LogMsg(name & " failed. " & msg) Then err.Raise knCall, , ksCall

Exit Function
ErrorHandler:
ErrTrap ksErrMod, "ITestLogger_LogFailure"
ITestLogger_LogFailure = True
End Function
'Public Sub ITestLogger_LogFailure(name As String, msg As String)
'
'    If ITestLogger_LogMsg(name & " failed. " & msg) Then err.Raise knCall, , ksCall
'
'End Sub

Rem =head4 Function ~
Rem sheetname ITestLogger_StartTestFixture
Rem
Rem Calls LogMsg with a string stating which test is starting.
Rem
Function ITestLogger_StartTestFixture(name As String) As Boolean
On Error GoTo ErrorHandler

    If ITestLogger_LogMsg("Start: " & name) Then err.Raise knCall, , ksCall

Exit Function
ErrorHandler:
ErrTrap ksErrMod, "ITestLogger_StartTestFixture"
ITestLogger_StartTestFixture = True
End Function
'Public Sub ITestLogger_StartTestFixture(name As String)
'
'    If ITestLogger_LogMsg("Start: " & name) Then err.Raise knCall, , ksCall
'
'End Sub

Rem =head4 Function ~
Rem sheetname ITestLogger_EndTestFixture
Rem
Rem Uses LogMsg to output a line detailing successes and failures of the module, a line of
Rem equals signs and a blank line.
Rem
Function ITestLogger_EndTestFixture(successCount As Long, failureCount As Long) As Boolean
On Error GoTo ErrorHandler

    Dim msg As String
    msg = successCount & " passes, " & failureCount & " failures"
    If ITestLogger_LogMsg(msg) Then err.Raise knCall, , ksCall
    If ITestLogger_LogMsg(String(Len(msg), "=")) Then err.Raise knCall, , ksCall
    If ITestLogger_LogMsg() Then err.Raise knCall, , ksCall

Exit Function
ErrorHandler:
ErrTrap ksErrMod, "ITestLogger_EndTestFixture"
ITestLogger_EndTestFixture = True
End Function
'Public Sub ITestLogger_EndTestFixture(successCount As Long, failureCount As Long)
'
'    Dim msg As String
'    msg = successCount & " passes, " & failureCount & " failures"
'    If ITestLogger_LogMsg(msg) Then err.Raise knCall, , ksCall
'    If ITestLogger_LogMsg(String(Len(msg), "=")) Then err.Raise knCall, , ksCall
'    If ITestLogger_LogMsg() Then err.Raise knCall, , ksCall
'
'End Sub

'Rem =head4 Function ITestLogger_StartTestCase
'Rem
'Rem rcl stub
'Rem
'Rem rcl ToFn
'Rem
'Public Sub ITestLogger_StartTestCase(name As String)
'End Sub

Rem =head4 Function ~
Rem sheetname ITestLogger_EndTestCase
Rem
Rem Reports the success and failure count.
Rem
Function ITestLogger_EndTestCase(name As String, _
                                 successCount As Long, _
                                 failureCount As Long) As Boolean
On Error GoTo ErrorHandler

    If ITestLogger_LogMsg(name & ": " & successCount & " passes, " _
                          & failureCount & " failures") Then err.Raise knCall, , ksCall

Exit Function
ErrorHandler:
ErrTrap ksErrMod, "ITestLogger_EndTestCase"
ITestLogger_EndTestCase = True
End Function
'Public Sub ITestLogger_EndTestCase(name As String, successCount As Long, failureCount As Long)
'
'    If ITestLogger_LogMsg(name & ": " & successCount & " passes, " _
'                          & failureCount & " failures") Then err.Raise knCall, , ksCall
'
'End Sub

Rem =head4 Function ~
Rem sheetname ITestLogger_EndTestSuite
Rem
Rem Appends data to a string for later checking against an expected value.
Rem
Function ITestLogger_EndTestSuite(successCount As Long, failureCount As Long) As Boolean
On Error GoTo ErrorHandler

    If ITestLogger_LogMsg("Total: " & successCount & " passes, " & _
                          failureCount & " failures") Then err.Raise knCall, , ksCall

Exit Function
ErrorHandler:
ErrTrap ksErrMod, "ITestLogger_EndTestSuite"
ITestLogger_EndTestSuite = True
End Function
'Public Sub ITestLogger_EndTestSuite(successCount As Long, failureCount As Long)
'
'    If ITestLogger_LogMsg("Total: " & successCount & " passes, " & _
'                          failureCount & " failures") Then err.Raise knCall, , ksCall
'
'End Sub

'Rem =head4 Property Let ITestLogger_ExpectedNumTestCases
'Rem
'Rem
'Rem
'Rem
'Rem rcl ToFn
'Rem
'Property Let ITestLogger_ExpectedNumTestCases(count As Long)
'
'End Property
