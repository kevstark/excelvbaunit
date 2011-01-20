VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Rem order 18
Rem
Rem =head2
Rem sheetname
Rem
Rem This should not be confused with TestI<Results>Manager, which handles the outputs from
Rem tests. This class manages the tests themselves, or more accurately the modules that contain
Rem tests. It has functions to determine which modules (fixtures) in a project are so named
Rem as to indicate that they might contain tests that should be run automatically. Two private
Rem constants have been moved to "Main" and made public as they are used by modRetroFit.
Rem
Rem =head3
Rem sheetname Macros
Rem

'Private Const TestModulePrefix = "Test"
'Private Const TestModuleSuffix = "Tester"
Const ksErrMod As String = "TestManager"

Rem =head4 Function ~
Rem sheetname GetTestFixtures
Rem
Rem Given a project name as a string, returns an array of modules capable of containing tests.
Rem
Function GetTestFixtures(projectName As String) As TestFixture()
On Error GoTo ErrorHandler
    
    ' Project not found?
    Dim prj As VBProject
    If DudProj(projectName, prj) Then err.Raise knCall, , ksCall
    ' Get list of modules containing test cases
    Dim components() As VBComponent
    components = GetTestingComponents(prj)
    
    ' Each module corresponds to a fixture
    Dim fixtures() As TestFixture
    Dim count As Integer
    count = Main.SafeUbound(components)
    
    If count >= 0 Then
        ReDim fixtures(0 To count) As TestFixture
        
        Dim i As Integer
        For i = 0 To count
            Set fixtures(i) = New TestFixture
            fixtures(i).ExtractTestCases prj, components(i)
        Next
    End If
    
    GetTestFixtures = fixtures
    Exit Function
    
ErrorHandler:
ErrTrap ksErrMod, "GetTestFixtures"
End Function

Rem =head4 Function ~
Rem sheetname GetTestFixture
Rem
Rem Given the names of a project and a fixture (module), returns that fixture as an object.
Rem
Function GetTestFixture(projectName As String, fixtureName As String) As TestFixture
On Error GoTo ErrorHandler

    Dim project As VBProject
    Set project = Application.VBE.VBProjects(projectName)
    ' Project not found?

    ' Get list of modules containing test cases
    Dim component As VBComponent
    Set component = project.VBComponents(fixtureName)

    ' Each module corresponds to a fixture
    Dim fixture As TestFixture
    Set fixture = New TestFixture
    fixture.ExtractTestCases project, component

    Set GetTestFixture = fixture

Exit Function
ErrorHandler:
ErrTrap ksErrMod, "GetTestFixture"
GetTestFixture = True
End Function
'Public Function GetTestFixture(projectName As String, fixtureName As String) As TestFixture
'
'    Dim project As VBProject
'    Set project = Application.VBE.VBProjects(projectName)
'    ' Project not found?
'
'    ' Get list of modules containing test cases
'    Dim component As VBComponent
'    Set component = project.VBComponents(fixtureName)
'
'    ' Each module corresponds to a fixture
'    Dim fixture As TestFixture
'    Set fixture = New TestFixture
'    fixture.ExtractTestCases project, component
'
'    Set GetTestFixture = fixture
'
'End Function

Rem =head4 Function ~
Rem sheetname GetTestingComponents
Rem
Rem Given a project as an object, returns as objects an array of fixtures (modules) that might
Rem contain tests.
Rem
Friend Function GetTestingComponents(project As VBProject) As VBComponent()
On Error GoTo ErrorHandler

    Dim testingComponents() As VBComponent
    
    If Not project Is Nothing Then
        Dim count As Integer
        'count = GetTestingComponentsCount(project)
        count = -1
        
        'If count > 0 Then
        Dim component As VBComponent
        For Each component In project.VBComponents
            If IsTestComponent(component) Then
                count = count + 1
                'ReDim testingComponents(0 To count - 1) As VBComponent
                ReDim Preserve testingComponents(0 To count) As VBComponent
                Set testingComponents(count) = component
    '            Dim component As VBComponent
    '            Dim i As Integer
    '            For Each component In project.VBComponents
    '                If IsTestComponent(component) Then
    '                    Set testingComponents(i) = component
    '                    i = i + 1
    '                End If
    '
    '            Next
            End If 'IsTestComponent(component) Then
        Next component
    End If 'not project is nothing
    
    GetTestingComponents = testingComponents

Exit Function
ErrorHandler:
ErrTrap ksErrMod, "GetTestingComponents"
End Function
Rem =head4 Function GetTestingComponentsCount
Rem
Rem Refactored out to remove the double loop and the need for a count.
Rem
'Rem
'Rem rcl ToFn
'Rem
'Friend Function GetTestingComponentsCount(project As VBProject) As Integer
'
'    Dim count As Integer
'    Dim component As VBComponent
'    For Each component In project.VBComponents
'        If IsTestComponent(component) Then
'            count = count + 1
'        End If
'    Next
'
'    GetTestingComponentsCount = count
'
'End Function
' Counts the number of modules containing tests in a project

Rem =head4 Function ~
Rem sheetname IsTestComponent
Rem
Rem Two of three conditions are needed for this to return True. The object must be a
Rem standard code module and the name must either start or end (or both) with the designated
Rem text. The text is held in constants, and at the time of documenting was "Test" for the
Rem prefix and "Tester" for the suffix. There is no obvious reason for this to need to change.
Rem
Friend Function IsTestComponent(component As VBComponent) As Boolean
On Error GoTo ErrorHandler

    IsTestComponent = False
    If component.Type = vbext_ct_StdModule Then
        If Left(component.name, Len(ksTestModulePrefix)) = ksTestModulePrefix Or _
           Right(component.name, Len(ksTestModuleSuffix)) = ksTestModuleSuffix Then
            IsTestComponent = True
        End If
    End If
    
Exit Function
ErrorHandler:
ErrTrap ksErrMod, "IsTestComponent"
End Function
