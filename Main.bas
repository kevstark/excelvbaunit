Attribute VB_Name = "Main"
Option Explicit

' Runs all tests in a VBA project
Public Sub xRun(projectName As String, Optional fixtureNameToBeRun As String = Empty)

    Dim runner As TestRunner
    Set runner = New TestRunner
        
    Dim resultsManager As ITestResultsManager
    Set resultsManager = New TestResultsManager
    runner.Run projectName, resultsManager, fixtureNameToBeRun

End Sub


