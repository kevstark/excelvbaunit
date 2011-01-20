Attribute VB_Name = "DummyTestModule3"
Option Explicit
Option Private Module
Rem order 10.3
Rem
Rem =head2
Rem sheetname
Rem
Rem rcl TestSuite
Rem rcl AsData
Rem It contains code that actually
Rem runs and public variables that record whether the code has been run or not. This means that
Rem developers can verify that tests actually get run by the test harness.
Rem
Rem =head3
Rem sheetname Macros
Rem

Public MeCalled As Boolean
Public Test1Called As Boolean
Public Test2Called As Boolean
Public SetUpCalled As Integer
Public TearDownCalled As Integer
Const ksErrMod As String = "DummyTestModule3"

Rem =head4 Function ~
Rem sheetname Reset
Rem
Rem Reset() As Boolean
Rem
Rem rcl True
Rem
Rem Resets all global variables before tests are run. Will have to be called explicitly, since
Rem it is not called from the SetUp routine as one might expect. This ensures that false or
Rem missing calls to the SetUp procedure will not invalidate other tests.
Rem
Function Reset() As Boolean
On Error GoTo ErrorHandler

    MeCalled = False
    Test1Called = False
    Test2Called = False
    SetUpCalled = 0
    TearDownCalled = 0

Exit Function
ErrorHandler:
ErrTrap ksErrMod, "Reset"
Reset = True
End Function
'Public Sub Reset()
'
'    MeCalled = False
'    Test1Called = False
'    Test2Called = False
'    SetUpCalled = 0
'    TearDownCalled = 0
'
'End Sub

Rem =head4 Function ~
Rem sheetname CallMe
Rem
Rem CallMe() As Boolean
Rem
Rem rcl True
Rem
Rem From its name, this should not be called automatically, and so the boolean should remain
Rem false. However, it is called explicitly at various stages of the test process to prove
Rem that a process invoked by a caller other than VBA is properly called.
Rem
Function CallMe() As Boolean
On Error GoTo ErrorHandler

    MeCalled = True

Exit Function
ErrorHandler:
ErrTrap ksErrMod, "CallMe"
CallMe = True
End Function
'Public Sub CallMe()
'    MeCalled = True
'End Sub

Rem =head4 Function ~
Rem sheetname Test1
Rem
Rem Test1() As Boolean
Rem
Rem rcl True
Rem
Rem This should be called automatically, setting the boolean to true.
Rem
Function Test1() As Boolean
On Error GoTo ErrorHandler

    Test1Called = True

Exit Function
ErrorHandler:
ErrTrap ksErrMod, "Test1"
Test1 = True
End Function
'Public Sub Test1()
'    Test1Called = True
'End Sub

Rem =head4 Function ~
Rem sheetname Test2
Rem
Rem Test2() As Boolean
Rem
Rem rcl True
Rem
Rem This should be called automatically, setting the boolean to true.
Rem
Function Test2() As Boolean
On Error GoTo ErrorHandler

    Test2Called = True

Exit Function
ErrorHandler:
ErrTrap ksErrMod, "Test2"
Test2 = True
End Function
'Public Sub Test2()
'    Test2Called = True
'End Sub

Rem =head4 Function ~
Rem sheetname SetUp
Rem
Rem SetUp() As Boolean
Rem
Rem rcl True
Rem
Rem This should be called before anything else in this module.
Rem
Function SetUp() As Boolean
On Error GoTo ErrorHandler

    SetUpCalled = SetUpCalled + 1

Exit Function
ErrorHandler:
ErrTrap ksErrMod, "SetUp"
SetUp = True
End Function
'Public Sub SetUp()
'    SetUpCalled = SetUpCalled + 1
'End Sub

Rem =head4 Function ~
Rem sheetname TearDown
Rem
Rem TearDown() As Boolean
Rem
Rem rcl True
Rem
Rem This should be called after everything else in this module.
Rem
Function TearDown() As Boolean
On Error GoTo ErrorHandler

    TearDownCalled = TearDownCalled + 1

Exit Function
ErrorHandler:
ErrTrap ksErrMod, "TearDown"
TearDown = True
End Function
'Public Sub TearDown()
'    TearDownCalled = TearDownCalled + 1
'End Sub
