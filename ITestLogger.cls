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
Rem order 11
Rem
Rem =head2
Rem sheetname
Rem
Rem This class module is implemented by two other class modules, "DebugTestLogger" and
Rem "FakeDebugTestLogger". This is done so that the real logger can be used when the system is
Rem run live, but when the system is being used to test itself, the fake logger can be used.
Rem An example of this is the LogMsg function. This, in the live version, will print the
Rem input to the Immediate pane. In the fake version, it will append it to a variable so that
Rem it can be checked against the expected result.
Rem
Rem This has the effect of allowing overloading of methods. The calling code does not need to
Rem know whether it is running in live or test mode, as that is determined when the object is
Rem initialised as live or fake. This means that the code getting tested is the live code, not
Rem some special version that had to be written or called differently to be testable.
Rem
Rem The practical implication of this is that all procedures in this class module are stubs,
Rem since the functionality is implemented in other class modules. The procedures are listed,
Rem so that there is a definitive list that can be reviewed in the documentation, but no
Rem description has been provided beyond what appears here.
Rem
Rem The same technique is used in the "*TestResultsManager" class modules. The commentary on
Rem them refers to this section, rather than repeating the description. The documentation groups
Rem the class modules together, rather than retaining the alphabetical order shown in the GUI,
Rem as is more usual.
Rem
Rem =head3
Rem sheetname Macros
Rem
Const ksErrMod As String = "ITestLogger"

Rem =head4 Function ~
Rem sheetname LogMsg
Rem
' Interface to logging progress of the test run
Function LogMsg(Optional msg As String = "") As Boolean
On Error GoTo ErrorHandler

Exit Function
ErrorHandler:
ErrTrap ksErrMod, "LogMsg"
LogMsg = True
End Function
'Public Sub LogMsg(Optional msg As String = "")
'End Sub

Rem =head4 Function ~
Rem sheetname LogSuccess
Rem
Function LogSuccess() As Boolean
On Error GoTo ErrorHandler

Exit Function
ErrorHandler:
ErrTrap ksErrMod, "LogSuccess"
LogSuccess = True
End Function
'Public Sub LogSuccess()
'End Sub

Rem =head4 Function ~
Rem sheetname LogFailure
Rem
Function LogFailure(name As String, msg As String) As Boolean
On Error GoTo ErrorHandler

Exit Function
ErrorHandler:
ErrTrap ksErrMod, "LogFailure"
LogFailure = True
End Function
'Public Sub LogFailure(name As String, msg As String)
'End Sub

Rem =head4 Function ~
Rem sheetname StartTestFixture
Rem
Function StartTestFixture(name As String) As Boolean
On Error GoTo ErrorHandler

Exit Function
ErrorHandler:
ErrTrap ksErrMod, "StartTestFixture"
StartTestFixture = True
End Function
'Public Sub StartTestFixture(name As String)
'End Sub

Rem =head4 Function ~
Rem sheetname EndTestFixture
Rem
Function EndTestFixture(successCount As Long, failureCount As Long) As Boolean
On Error GoTo ErrorHandler

Exit Function
ErrorHandler:
ErrTrap ksErrMod, "EndTestFixture"
EndTestFixture = True
End Function
'Public Sub EndTestFixture(successCount As Long, failureCount As Long)
'End Sub

Rem =head4 Function ~
Rem sheetname StartTestCase
Rem
Rem All TestLogger.StartTestCase procedures are blank, so this has been commented out.
Rem
'Public Sub StartTestCase(name As String)
'End Sub

Rem =head4 Function ~
Rem sheetname EndTestCase
Rem
Function EndTestCase(name As String, successCount As Long, failureCount As Long) As Boolean
On Error GoTo ErrorHandler

Exit Function
ErrorHandler:
ErrTrap ksErrMod, "EndTestCase"
EndTestCase = True
End Function
'Public Sub EndTestCase(name As String, successCount As Long, failureCount As Long)
'End Sub

Rem =head4 Function ~
Rem sheetname EndTestSuite
Rem
Function EndTestSuite(successCount As Long, failureCount As Long) As Boolean
On Error GoTo ErrorHandler

Exit Function
ErrorHandler:
ErrTrap ksErrMod, "EndTestSuite"
EndTestSuite = True
End Function
'Public Sub EndTestSuite(successCount As Long, failureCount As Long)
'End Sub

Rem =head4 Property Let ~
Rem sheetname ExpectedNumTestCases
Rem
Rem This has been commented out of all class modules in which it appears. All such procedures
Rem are blank, except one that passes a variable from TestResultsManager to this empty
Rem procedure, and no class module contains a variable of this name. It is invoked by
Rem TestRunner.SetUpResultsManager (the section doing this has been commented out, too), but
Rem the property is never used. The original (MH) comments in the code say "Results manager
Rem need to know how many tests case to expect to help it logging progress" (sic). However,
Rem commenting everything out does not seem to cause the code any problems, returning exactly
Rem the same results and passing all internal tests. This smells like a relic.
Rem
'Property Let ExpectedNumTestCases(count As Long)
'End Property
