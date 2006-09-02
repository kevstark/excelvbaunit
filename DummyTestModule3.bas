Attribute VB_Name = "DummyTestModule3"
Option Explicit

Public MeCalled As Boolean
Public Test1Called As Boolean
Public Test2Called As Boolean

Public Sub Reset()

    MeCalled = False
    Test1Called = False
    Test2Called = False
    
End Sub
Public Sub CallMe()
    MeCalled = True
End Sub

Public Sub Test1()
    Test1Called = True
End Sub

Public Sub Test2()
    Test2Called = True
End Sub

