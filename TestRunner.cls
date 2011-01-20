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
Rem order 19
Rem
Rem =head2
Rem sheetname
Rem
Rem The purpose of this class is to organise the running of tests. It is stateless.
Rem
Rem =head3
Rem sheetname Macros
Rem
Const ksErrMod As String = "TestRunner"

Rem =head4 Function ~
Rem sheetname Run
Rem
Rem Given the name of a project as a string and a results manager as an object (remember that
Rem the ITestResultsManager class has two implementation classes), runs all tests within the
Rem project. Optionally, the name of a fixture (module) may be passed as a third argument, in
Rem which case only the tests in that fixture are run.
Rem
Function Run(projectName As String, _
             resultsManager As ITestResultsManager, _
             Optional fixtureNameToBeRun As String = Empty) As Boolean
On Error GoTo ErrorHandler

    Dim manager As TestManager
    Set manager = New TestManager
    
    Dim fixtures() As TestFixture
    fixtures = GetTestFixtures(projectName, manager, fixtureNameToBeRun)

    'JHD: SetUpResultsManager refactored out. Called only from here.
    'SetUpResultsManager resultsManager, fixtures, fixtureNameToBeRun
    If Assert.SetTestResultsManager(resultsManager) Then err.Raise knCall, , ksCall
    
    Dim nTotFix As Long
    nTotFix = SafeUbound(fixtures)
    
    Dim i As Integer
    For i = 0 To nTotFix
        ' If a specific test fixture was named, then run only that fixture
        If (ShouldRunFixture(fixtures(i), fixtureNameToBeRun)) Then
            If fixtures(i).RunTests(resultsManager) Then err.Raise knCall, , ksCall
        End If
    Next
    
    If resultsManager.EndTestSuite() Then err.Raise knCall, , ksCall

Exit Function
ErrorHandler:
ErrTrap ksErrMod, "Run"
Run = True
End Function
'Public Sub Run(projectName As String, _
'               resultsManager As ITestResultsManager, _
'               Optional fixtureNameToBeRun As String = Empty)
'
'    Dim manager As TestManager
'    Set manager = New TestManager
'
'    Dim fixtures() As TestFixture
'    fixtures = GetTestFixtures(projectName, manager, fixtureNameToBeRun)
'
'    'JHD: SetUpResultsManager refactored out. Called only from here.
'    'SetUpResultsManager resultsManager, fixtures, fixtureNameToBeRun
'    If Assert.SetTestResultsManager(resultsManager) Then err.Raise knCall, , ksCall
'
'    Dim nTotFix As Long
'    nTotFix = SafeUbound(fixtures)
'
'    Dim i As Integer
'    For i = 0 To nTotFix
'        ' If a specific test fixture was named, then run only that fixture
'        If (ShouldRunFixture(fixtures(i), fixtureNameToBeRun)) Then
'            If fixtures(i).RunTests(resultsManager) Then err.Raise knCall, , ksCall
'        End If
'    Next
'
'    If resultsManager.EndTestSuite() Then err.Raise knCall, , ksCall
'
'End Sub

'Rem =head4 Function
'Rem
'Rem
'Rem
'Rem
'Rem rcl ToFn
'Rem
'Private Sub SetUpResultsManager(resultsManager As ITestResultsManager, _
'                                fixtures() As TestFixture, _
'                                fixtureNameToBeRun As String)
'
''JHD: Now that the count process has been handled better, this sub has been refactored out.
'
'    ' Assert class need to know the results manager so it can keep track of success and failures
'    If Assert.SetTestResultsManager(resultsManager) Then err.Raise knCall, , ksCall
'
'    ' Results manager need to know how many tests case to expect to help it logging progress
''    Dim i As Long
''    Dim testCaseCount As Long
''    For i = 0 To UBound(fixtures)
''        If (ShouldRunFixture(fixtures(i), fixtureNameToBeRun)) Then
''            testCaseCount = testCaseCount + UBound(fixtures(i).TestProcedures) + 1
''        End If
''    Next
''
''    resultsManager.ExpectedNumTestCases = testCaseCount
'
'End Sub

Rem =head4 Function ~
Rem sheetname ShouldRunFixture
Rem
Rem This function returns a boolean indicating whether the module contains tests for running.
Rem Tests will not be run if a fixture (module) name is specified and does not match that of
Rem the first parameter. This function has been modified to make the fixture name an optional
Rem parameter.
Rem
Friend Function ShouldRunFixture(fixture As TestFixture, _
                                 Optional fixtureNameToBeRun As String = "") As Boolean
On Error GoTo ErrorHandler

    'JHD: added UCase() to avoid any unnecessary typo issues
    If fixtureNameToBeRun = "" Or UCase(fixtureNameToBeRun) = UCase(fixture.fixtureName) Then
        ShouldRunFixture = True
    Else
        ShouldRunFixture = False
    End If

Exit Function
ErrorHandler:
ErrTrap ksErrMod, "ShouldRunFixture"
End Function

Rem =head4 Function ~
Rem sheetname GetTestFixtures
Rem
Rem Given a project name as a string, a test manager as an object and a fixture (module) name
Rem as a string, returns an array of fixtures as objects.
Rem
Private Function GetTestFixtures(projectName As String, _
                                 manager As TestManager, _
                                 fixtureName As String) As TestFixture()
On Error GoTo ErrorHandler

    Dim fixtures() As TestFixture
    If fixtureName = "" Then
        fixtures = manager.GetTestFixtures(projectName)
    Else
        ReDim fixtures(0 To 0)
        Set fixtures(0) = manager.GetTestFixture(projectName, fixtureName)
    End If
    
    GetTestFixtures = fixtures

Exit Function
ErrorHandler:
ErrTrap ksErrMod, "GetTestFixtures"
End Function
