VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestFixture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const TestMethodPrefix = "Public Sub Test"
Const SetUpFunctionName = "SetUp"
Const TearDownFunctionName = "TearDown"

Private mTestProcedures() As String
Private mFileName As String
Private mFixtureName As String
Private mHasSetUpFunction As Boolean
Private mHasTearDownFunction As Boolean

Property Get TestProcedures() As String()
    TestProcedures = mTestProcedures
End Property

Property Get FileName() As String
    FileName = mFileName
End Property

Property Get FixtureName() As String
    FixtureName = mFixtureName
End Property

Friend Property Get HasSetUpFunction() As Boolean
    HasSetUpFunction = mHasSetUpFunction
End Property

Friend Property Get HasTearDownFunction() As Boolean
    HasTearDownFunction = mHasTearDownFunction
End Property


Public Sub RunTests(resultsManager As ITestResultsManager)

    resultsManager.StartTestFixture mFixtureName
    
    Dim i As Integer
    For i = 0 To UBound(mTestProcedures)
        InvokeProc resultsManager, mFileName, mTestProcedures(i)
    Next
    
    resultsManager.EndTestFixture
    
End Sub



Friend Sub InvokeProc(resultsManager As ITestResultsManager, FileName As String, proc As String)

    resultsManager.StartTestCase proc
    Application.Run (FileName & "!" & proc)
    resultsManager.EndTestCase
    
End Sub


' Extracts the test cases from a test fixture
Public Sub ExtractTestCases(project As VBProject, component As VBComponent)

    Dim count As Integer
    count = GetTestMethodsCount(component)
    
    mTestProcedures = GetTestMethods(component, count)
    
    mFileName = ExtractFileName(project.FileName)
    mFixtureName = component.name
    
    mHasSetUpFunction = DoesMethodExist(SetUpFunctionName, component)
    mHasTearDownFunction = DoesMethodExist(TearDownFunctionName, component)
    
End Sub


Friend Function DoesMethodExist(name As String, component As VBComponent) As Boolean

    Dim line As Integer
    Dim str As String
    Dim searchFor As String
    searchFor = LCase("Public Sub " & name)
    For line = 1 To component.CodeModule.CountOfLines
        str = LCase(component.CodeModule.Lines(line, 1))
        If Left(str, Len(searchFor)) = searchFor Then
            DoesMethodExist = True
            Exit Function
        End If
    Next
    
    DoesMethodExist = False

End Function



Friend Function ExtractFileName(FullPath As String) As String

    Dim parts() As String
    parts = Split(FullPath, "\")
    
    ExtractFileName = parts(UBound(parts))

End Function



Friend Function GetTestMethods(component As VBComponent, count As Integer) As String()
    
    ReDim Results(0 To count - 1) As String
    
    Dim line As Integer
    Dim proc As Integer
    For line = 1 To component.CodeModule.CountOfLines
        If IsTestMethodLine(component.CodeModule.Lines(line, 1)) Then
            Results(proc) = component.name & "." & component.CodeModule.ProcOfLine(line, 0)
            proc = proc + 1
        End If
    Next

    GetTestMethods = Results
End Function


Friend Function GetTestMethodsCount(component As VBComponent) As Integer
    Dim count As Integer
    Dim i As Integer
    For i = 1 To component.CodeModule.CountOfLines
        If IsTestMethodLine(component.CodeModule.Lines(i, 1)) Then
            count = count + 1
        End If
    Next
    
    GetTestMethodsCount = count

End Function



Friend Function IsTestMethodLine(line As String) As Boolean
    IsTestMethodLine = Left(line, Len(TestMethodPrefix)) Like TestMethodPrefix
End Function

