Attribute VB_Name = "Main"
Option Explicit

Public Sub xRun(projectName As String)

    Dim runner As TestRunner
    Set runner = New TestRunner
        
    Dim resultsManager As ITestResultsManager
    Set resultsManager = New TestResultsManager
    runner.Run projectName, resultsManager

End Sub
