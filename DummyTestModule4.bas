Attribute VB_Name = "DummyTestModule4"
Option Explicit
Option Private Module
Rem order 10.4
Rem
Rem =head2
Rem sheetname
Rem
Rem rcl TestSuite
Rem rcl AsData
Rem One of the procedures contains runnable code.
Rem rcl NoTrap
Rem
Rem =head3
Rem sheetname Macros
Rem
Const ksErrMod As String = "DummyTestModule4"

Public Sub Test1()
    
End Sub

Rem =head4 Function ~
Rem sheetname NotATest
Rem
Rem NotATest() As Boolean
Rem
Rem rcl True
Rem
Rem If this is called, it will throw a failure.
Rem
Function NotATest() As Boolean
On Error GoTo ErrorHandler

    If AssertTrue(False) Then err.Raise knCall, , ksCall

Exit Function
ErrorHandler:
ErrTrap ksErrMod, "NotATest"
NotATest = True
End Function
'Public Sub NotATest()
'    If AssertTrue(False) Then err.Raise knCall, , ksCall
'End Sub
