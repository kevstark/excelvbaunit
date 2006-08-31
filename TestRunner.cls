VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestRunner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub Run(projectName As String)

    Dim manager As TestManager
    Set manager = New TestManager
    
    Dim fixtures() As TestFixture
    fixtures = manager.GetTestFixtures(projectName)
    
    Dim resultsManager As TestResultsManager
    Set resultsManager = New TestResultsManager
    Assert.SetTestResultsManager resultsManager
    
    Dim i As Integer
    For i = 0 To UBound(fixtures)
        fixtures(i).RunTests resultsManager
    Next
    
    resultsManager.EndTestSuite

End Sub

