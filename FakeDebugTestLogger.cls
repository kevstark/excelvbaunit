VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FakeDebugTestLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Rem order 13
Rem
Rem =head2
Rem sheetname
Rem
Rem This is the fake implementation of "ITestLogger". See L</ITestLogger> for more details.
Rem
Rem =head3
Rem sheetname Macros
Rem

Implements ITestLogger

' Hold onto the last message so we can check it
Public message As String
Const ksErrMod As String = "FakeDebugTestLogger"

Rem =head4 Function ~
Rem sheetname ITestLogger_LogMsg
Rem
Rem Stores the message parameter to a module level variable, delimited by colons. All other
Rem functions in this module append to the variable via this function. Explanations should
Rem be added only when this principle is broken. All procedures are listed for completeness,
Rem but the lack of documentation merely indicates that they work as described here.
Rem
Function ITestLogger_LogMsg(Optional msg As String = "") As Boolean
On Error GoTo ErrorHandler

    message = message & ":" & msg

Exit Function
ErrorHandler:
ErrTrap ksErrMod, "ITestLogger_LogMsg"
ITestLogger_LogMsg = True
End Function
'Public Sub ITestLogger_LogMsg(Optional msg As String = "")
'    message = message & ":" & msg
'End Sub

Rem =head4 Function ~
Rem sheetname ITestLogger_LogSuccess
Rem
Function ITestLogger_LogSuccess() As Boolean
On Error GoTo ErrorHandler

    If ITestLogger_LogMsg("Success") Then err.Raise knCall, , ksCall

Exit Function
ErrorHandler:
ErrTrap ksErrMod, "ITestLogger_LogSuccess"
ITestLogger_LogSuccess = True
End Function
'Public Sub ITestLogger_LogSuccess()
'    If ITestLogger_LogMsg("Success") Then err.Raise knCall, , ksCall
'End Sub

Rem =head4 Function ~
Rem sheetname ITestLogger_LogFailure
Rem
Function ITestLogger_LogFailure(name As String, msg As String) As Boolean
On Error GoTo ErrorHandler

    If ITestLogger_LogMsg(name & " " & msg & " Failed") Then err.Raise knCall, , ksCall

Exit Function
ErrorHandler:
ErrTrap ksErrMod, "ITestLogger_LogFailure"
ITestLogger_LogFailure = True
End Function
'Public Sub ITestLogger_LogFailure(name As String, msg As String)
'    If ITestLogger_LogMsg(name & " " & msg & " Failed") Then err.Raise knCall, , ksCall
'End Sub

Rem =head4 Function ~
Rem sheetname ITestLogger_StartTestFixture
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
'    If ITestLogger_LogMsg("Start: " & name) Then err.Raise knCall, , ksCall
'End Sub

Rem =head4 Function ~
Rem sheetname ITestLogger_EndTestFixture
Rem
Function ITestLogger_EndTestFixture(successCount As Long, failureCount As Long) As Boolean
On Error GoTo ErrorHandler

    If ITestLogger_LogMsg(successCount & " passes, " & failureCount & " failures") _
        Then err.Raise knCall, , ksCall

Exit Function
ErrorHandler:
ErrTrap ksErrMod, "ITestLogger_EndTestFixture"
ITestLogger_EndTestFixture = True
End Function
'Public Sub ITestLogger_EndTestFixture(successCount As Long, failureCount As Long)
'    If ITestLogger_LogMsg(successCount & " passes, " & failureCount & " failures") _
'        Then err.Raise knCall, , ksCall
'End Sub

'Rem =head4 Function ~
'Rem sheetname ITestLogger_StartTestCase
'Rem
'Rem
'Rem
'Rem
'Rem rcl ToFn
'Rem
'Public Sub ITestLogger_StartTestCase(name As String)
'End Sub

Rem =head4 Function ~
Rem sheetname ITestLogger_EndTestCase
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
'    If ITestLogger_LogMsg(name & ": " & successCount & " passes, " _
'                          & failureCount & " failures") Then err.Raise knCall, , ksCall
'End Sub

Rem =head4 Function ~
Rem sheetname ITestLogger_EndTestSuite
Rem
Function ITestLogger_EndTestSuite(successCount As Long, failureCount As Long) As Boolean
On Error GoTo ErrorHandler

    If ITestLogger_LogMsg("Total: " & successCount & " passes, " _
                          & failureCount & " failures") Then err.Raise knCall, , ksCall

Exit Function
ErrorHandler:
ErrTrap ksErrMod, "ITestLogger_EndTestSuite"
ITestLogger_EndTestSuite = True
End Function
'Public Sub ITestLogger_EndTestSuite(successCount As Long, failureCount As Long)
'    If ITestLogger_LogMsg("Total: " & successCount & " passes, " _
'                          & failureCount & " failures") Then err.Raise knCall, , ksCall
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
