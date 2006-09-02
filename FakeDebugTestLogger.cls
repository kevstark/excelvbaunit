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

Implements ITestLogger

' Hold onto the last message so we can check it
Public message As String

Public Sub ITestLogger_LogMsg(Optional msg As String = "")
    message = message & ":" & msg
End Sub

