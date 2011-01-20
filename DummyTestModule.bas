Attribute VB_Name = "DummyTestModule"
Option Explicit
Option Private Module
Rem order 10.1
Rem
Rem =head2
Rem sheetname
Rem
Rem sto TestSuite
Rem This is part of the test suite for the test harness add-in. It should not be changed
Rem except to support testing of the add-in, no matter how strange or redundant the
Rem procedures seem.
Rem sto AsData
Rem It is treated as data by various testing modules.
Rem sto NoTrap
Rem Procedures without runnable code have not been wrapped in error trapping or documented.
Rem sto 0
Rem rcl TestSuite
Rem rcl AsData
Rem rcl NoTrap
Rem
Rem =head3
Rem sheetname Macros
Rem
Rem No runnable code
Rem
Const ksErrMod As String = "DummyTestModule"

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

Public Sub TearDown()

End Sub
