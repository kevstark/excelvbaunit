Attribute VB_Name = "DummyTestModule2"
Option Explicit
Option Private Module
Rem order 10.2
Rem
Rem =head2
Rem sheetname
Rem
Rem rcl TestSuite
Rem rcl AsData
Rem rcl NoTrap
Rem
Rem =head3
Rem sheetname Macros
Rem
Rem No runnable code
Rem
Const ksErrMod As String = "DummyTestModule2"

Public Sub Test1()

End Sub

Private Sub Test2()

End Sub

Public Function Test3()

End Function

Public Sub NotATest()

End Sub

Public Sub Test4()

End Sub

Public Sub SetUp()

End Sub

Public Sub FixtureSetUp()
    AssertSuccess
End Sub
