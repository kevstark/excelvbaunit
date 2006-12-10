Attribute VB_Name = "Main"
Option Explicit

' Runs all tests in a VBA project
Public Sub xRun(projectName As String, _
                Optional fixtureNameToBeRun As String = Empty, _
                Optional testLogger As ITestLogger)

    Dim runner As TestRunner
    Set runner = New TestRunner
        
    Dim resultsManager As ITestResultsManager: Set resultsManager = New TestResultsManager
    
    If Not testLogger Is Nothing Then
        Set resultsManager.testLogger = testLogger
    End If
    
    runner.Run projectName, resultsManager, fixtureNameToBeRun

End Sub

