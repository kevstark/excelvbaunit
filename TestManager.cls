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

Private Const TestModulePrefix = "Test"
Private Const TestModuleSuffix = "Tester"


Public Function GetTestFixtures(projectName As String) As TestFixture()

    Dim project As VBProject
    Set project = Application.VBE.VBProjects(projectName)
    ' Project not found?
    
    ' Get list of modules containing test cases
    Dim components() As VBComponent
    components = GetTestingComponents(project)
    
    ' Each module corresponds to a fixture
    Dim fixtures() As TestFixture
    Dim count As Integer
    count = Main.SafeUbound(components)
    
    If count > 0 Then
        ReDim fixtures(0 To count) As TestFixture
        
        Dim i As Integer
        For i = 0 To count
            Set fixtures(i) = New TestFixture
            fixtures(i).ExtractTestCases project, components(i)
        Next
    End If
    
    GetTestFixtures = fixtures
End Function




Public Function GetTestFixture(projectName As String, fixtureName As String) As TestFixture

    Dim project As VBProject
    Set project = Application.VBE.VBProjects(projectName)
    ' Project not found?
    
    ' Get list of modules containing test cases
    Dim component As VBComponent
    Set component = project.VBComponents(fixtureName)
            
    ' Each module corresponds to a fixture
    Dim fixture As TestFixture: Set fixture = New TestFixture
    fixture.ExtractTestCases project, component
    
    Set GetTestFixture = fixture

End Function



Friend Function GetTestingComponents(project As VBProject) As VBComponent()

    Dim testingComponents() As VBComponent
    
    Dim count As Integer
    count = GetTestingComponentsCount(project)
    
    If count > 0 Then
        ReDim testingComponents(0 To count - 1) As VBComponent
        Dim component As VBComponent
        Dim i As Integer
        For Each component In project.VBComponents
            If IsTestComponent(component) Then
                Set testingComponents(i) = component
                i = i + 1
            End If
            
        Next
    End If
    
    GetTestingComponents = testingComponents

End Function



Friend Function GetTestingComponentsCount(project As VBProject) As Integer

    Dim count As Integer
    Dim component As VBComponent
    For Each component In project.VBComponents
        If IsTestComponent(component) Then
            count = count + 1
        End If
    Next
    
    GetTestingComponentsCount = count

End Function


' Counts the number of modules containing tests in a project
Friend Function IsTestComponent(component As VBComponent) As Boolean

    IsTestComponent = False
    If component.Type = vbext_ct_StdModule Then
        If Left(component.name, Len(TestModulePrefix)) = TestModulePrefix Or _
           Right(component.name, Len(TestModuleSuffix)) = TestModuleSuffix Then
            IsTestComponent = True
        End If
    End If
    
End Function
