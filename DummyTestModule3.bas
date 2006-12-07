Attribute VB_Name = "DummyTestModule3"
Option Explicit

Public MeCalled As Boolean
Public Test1Called As Boolean
Public Test2Called As Boolean
Public SetUpCalled As Integer
Public TearDownCalled As Integer

Public Sub Reset()

    MeCalled = False
    Test1Called = False
    Test2Called = False
    SetUpCalled = 0
    TearDownCalled = 0
    
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


Public Sub SetUp()

    SetUpCalled = SetUpCalled + 1
End Sub

Public Sub TearDown()
    TearDownCalled = TearDownCalled + 1
End Sub
