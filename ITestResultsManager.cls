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
Rem order 14
Rem
Rem =head2
Rem sheetname
Rem
Rem sto ClassLinks
Rem See L</ITestLogger> for a description of how these three modules, ITestResultsManager,
Rem TestResultsManager and FakeTestResultsManager work together. This also describes the
Rem need for stub functions. The absence of any documentation merely indicates that the
Rem function is a stub by design or that the code is so trivial it can be inferred from the
Rem name.
Rem sto 0
Rem rcl ClassLinks
Rem
Rem =head3
Rem sheetname Macros
Rem
Const ksErrMod As String = "ITestResultsManager"

Rem =head4 Property Set ~
Rem sheetname testLogger
Rem
Property Set testLogger(logger As ITestLogger)
End Property

Rem =head4 Property Get ~
Rem sheetname TotalSuccessCount
Rem
Property Get TotalSuccessCount() As Integer
End Property

Rem =head4 Property Get ~
Rem sheetname TotalFailureCount
Rem
Property Get TotalFailureCount() As Integer
End Property

Rem =head4 Property Get ~
Rem sheetname FixtureSuccessCount
Rem
Property Get FixtureSuccessCount() As Integer
End Property

Rem =head4 Property Get ~
Rem sheetname FixtureFailureCount
Rem
Property Get FixtureFailureCount() As Integer
End Property

Rem =head4 Property Get ~
Rem sheetname TestCaseSuccessCount
Rem
Property Get TestCaseSuccessCount() As Integer
End Property

Rem =head4 Property Get ~
Rem sheetname TestCaseFailureCount
Rem
Property Get TestCaseFailureCount() As Integer
End Property

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
Function LogFailure(msg As String) As Boolean
On Error GoTo ErrorHandler

Exit Function
ErrorHandler:
ErrTrap ksErrMod, "LogFailure"
LogFailure = True
End Function
'Public Sub LogFailure(msg As String)
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
Function EndTestFixture() As Boolean
On Error GoTo ErrorHandler

Exit Function
ErrorHandler:
ErrTrap ksErrMod, "EndTestFixture"
EndTestFixture = True
End Function
'Public Sub EndTestFixture()
'End Sub

Rem =head4 Function ~
Rem sheetname StartTestCase
Rem
Function StartTestCase(name As String) As Boolean
On Error GoTo ErrorHandler

Exit Function
ErrorHandler:
ErrTrap ksErrMod, "StartTestCase"
StartTestCase = True
End Function
'Public Sub StartTestCase(name As String)
'End Sub

Rem =head4 Function ~
Rem sheetname EndTestCase
Rem
Function EndTestCase() As Boolean
On Error GoTo ErrorHandler

Exit Function
ErrorHandler:
ErrTrap ksErrMod, "EndTestCase"
EndTestCase = True
End Function
'Public Sub EndTestCase()
'End Sub

Rem =head4 Function ~
Rem sheetname EndTestSuite
Rem
Function EndTestSuite() As Boolean
On Error GoTo ErrorHandler

Exit Function
ErrorHandler:
ErrTrap ksErrMod, "EndTestSuite"
EndTestSuite = True
End Function
'Public Sub EndTestSuite()
'End Sub

'Public Property Let ExpectedNumTestCases(count As Long)
'End Property
