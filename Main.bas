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



' Factory function so can use test manager functions outside the add in. As the class is stateless it doesn't
' matter that we return a new instance with this call
Public Function GetTestManager() As TestManager

    Set GetTestManager = New TestManager
    
End Function


Public Function SafeUbound(var As Variant) As Long

On Error GoTo err
    
    SafeUbound = UBound(var)
    Exit Function
    
err:
    SafeUbound = -1
End Function

