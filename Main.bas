Attribute VB_Name = "Main"
Option Explicit

Public Sub xRun(projectName As String)

    Dim runner As TestRunner
    Set runner = New TestRunner
    runner.Run projectName

End Sub
