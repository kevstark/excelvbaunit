Attribute VB_Name = "Main"
Option Explicit
Rem order 6
Rem
Rem =head2
Rem sheetname
Rem
Rem This code module contains the core code of the system, including xRun, the procedure that
Rem should be invoked from the Immediate pane to run tests.
Rem
Rem =head3
Rem sheetname Macros
Rem
Public AddInFileName                        As String
Public AddInName                            As String
Public Const ksTestModulePrefix             As String = "Test"
Public Const ksTestModuleSuffix             As String = "Tester"
Public Const ksSetUpFunctionName            As String = "SetUp"
Public Const ksTearDownFunctionName         As String = "TearDown"
Public Const ksFixtureSetUpFunctionName     As String = "FixtureSetUp"
Public Const ksFixtureTearDownFunctionName  As String = "FixtureTearDown"
Const ksErrMod                              As String = "Main"

Rem =head4 Function ~
Rem sheetname xRun
Rem
Rem xRun(Optional projectName As String = Empty, Optional fixtureNameToBeRun As String = Empty,
Rem Optional testLogger As ITestLogger) As Boolean
Rem
Rem rcl True
Rem
Rem This runs all tests in a project, unless the user specifies otherwise. The name of the
Rem project is the first (previously compulsory) parameter, which defaults to the name of the
Rem add-in. There are two further optional parameters. The first is the "Fixture" name - the
Rem Z<>name of a single module containing tests. If this is used, only that module's tests will
Rem be run. The last parameter is the logger. This is unlikely to be wanted except when testing
Rem this project itself.
Rem
Function xRun(Optional projectName As String = Empty, _
              Optional fixtureNameToBeRun As String = Empty, _
              Optional testLogger As ITestLogger) As Boolean
On Error GoTo ErrorHandler

    'Next two lines added by JHD to enable projectName to be optional
    If GetNames Then err.Raise knCall, , ksCall
    If Len(projectName) = 0 Then projectName = AddInName
    
    Dim runner As TestRunner
    Set runner = New TestRunner
        
    Dim resultsManager As ITestResultsManager
    'Set resultsManager = New TestResultsManager
    
    If Not testLogger Is Nothing Then
        Set resultsManager.testLogger = testLogger
    Else    'Added by JHD - "Set" line above taken out to eliminate redundancy
        Set resultsManager = New TestResultsManager
    End If
    
    If runner.Run(projectName, resultsManager, fixtureNameToBeRun) _
        Then err.Raise knCall, , ksCall

Exit Function
ErrorHandler:
ErrTrap ksErrMod, "xRun"
xRun = True
End Function
'Public Sub xRun(Optional projectName As String = Empty, _
'                Optional fixtureNameToBeRun As String = Empty, _
'                Optional testLogger As ITestLogger)
'
'    'Next two lines added by JHD to enable projectName to be optional
'    If GetNames Then err.Raise knCall, , ksCall
'    If Len(projectName) = 0 Then projectName = AddInName
'
'    Dim runner As TestRunner
'    Set runner = New TestRunner
'
'    Dim resultsManager As ITestResultsManager
'    'Set resultsManager = New TestResultsManager
'
'    If Not testLogger Is Nothing Then
'        Set resultsManager.testLogger = testLogger
'    Else    'Added by JHD - "Set" line above taken out to eliminate redundancy
'        Set resultsManager = New TestResultsManager
'    End If
'
'    runner.Run projectName, resultsManager, fixtureNameToBeRun
'
'End Sub

Rem =head4 Function ~
Rem sheetname GetTestManager
Rem
Rem Original documentation by MH:
Rem Factory function so can use test manager functions outside the add in. As the class is
Rem stateless, it doesn't matter that we return a new instance with this call.
Rem
Rem JHD: Unfortunately, this doesn't work. In the absence of the original MainTester, a new
Rem version has been written, but nothing seems to persuade it to return a test manager or do
Rem anything except raise an error in the calling procedure. It has therefore been commented out.
Rem
'Public Function GetTestManager() As TestManager
'On Error GoTo ErrorHandler
'
'    Set GetTestManager = New TestManager
'
'Exit Function
'ErrorHandler:
'ErrTrap ksErrMod, "GetTestManager"
'End Function
'Public Function GetTestManager() As TestManager
'
'    Set GetTestManager = New TestManager
'
'End Function

Rem =head4 Function ~
Rem sheetname SafeUbound
Rem
Rem SafeUbound(var As Variant) As Long
Rem
Rem Accepts a variant, which is expected to be an array. Returns the upper bound of the array,
Rem or if there is an error, returns -1.
Rem
Public Function SafeUbound(var As Variant) As Long

On Error GoTo err
    
    SafeUbound = UBound(var)
    Exit Function
    
err:
    SafeUbound = -1
End Function

Rem =head4 Function ~
Rem sheetname GetNames
Rem
Rem GetNames() As Boolean
Rem
Rem rcl True
Rem
Rem Populates name variables.
Rem
Function GetNames() As Boolean
On Error GoTo ErrorHandler

If Len(AddInFileName) = 0 Then
    AddInFileName = ThisWorkbook.name
    Dim lenName As Long
    lenName = Len(AddInFileName)
    Dim proj As VBProject
    For Each proj In Application.VBE.VBProjects
        If Right(proj.FileName, lenName) = AddInFileName Then
            AddInName = proj.name
            Exit For
        End If 'Right(proj.FileName, lenName) = AddInFileName Then
    Next proj
End If 'Len(AddInFileName) = 0

Exit Function
ErrorHandler:
ErrTrap ksErrMod, "GetNames"
End Function

Rem =head4 Function ~
Rem sheetname DudProj
Rem
Rem DudProj(prjName As String, prj As VBProject) As Boolean
Rem
Rem rcl True
Rem
Rem Traps any reference to a dud project, creating the reference if it's OK.
Rem
Function DudProj(prjName As String, prj As VBProject) As Boolean

    On Error GoTo NoProject
    Set prj = Application.VBE.VBProjects(prjName)

Exit Function

NoProject:
On Error GoTo ErrorHandler
MsgBox "Sorry! I can't find a project called " & prjName & "." & vbCrLf & _
       "Either you have mistyped it or it is not open" & vbCrLf & _
       "or it's in another instance of Excel or there's" & vbCrLf & _
       "something wrong in my code. Whichever the" & vbCrLf & _
       "problem is, it's now yours at no extra charge.", vbCritical, _
       "No project """ & prjName & """"
DudProj = True
Exit Function

ErrorHandler:
ErrTrap ksErrMod, "DudProj"
DudProj = True
End Function
