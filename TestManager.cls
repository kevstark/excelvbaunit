VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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
    Dim count As Integer
    count = UBound(components)
    ReDim fixtures(0 To count) As TestFixture
    
    Dim i As Integer
    For i = 0 To count
        Set fixtures(i) = New TestFixture
        fixtures(i).ExtractTestCases project, components(i)
    Next
    
    GetTestFixtures = fixtures
End Function


Private Function GetTestingComponents(project As VBProject)

    Dim count As Integer
    count = GetTestingComponentsCount(project)
    
    ReDim testingComponents(0 To count - 1) As VBComponent
    Dim component As VBComponent
    Dim i As Integer
    For Each component In project.VBComponents
        If IsTestComponent(component) Then
            Set testingComponents(i) = component
            i = i + 1
        End If
        
    Next
    
    GetTestingComponents = testingComponents

End Function



Private Function GetTestingComponentsCount(project As VBProject)

    Dim count As Integer
    Dim component As VBComponent
    For Each component In project.VBComponents
        If IsTestComponent(component) Then
            count = count + 1
        End If
    Next
    
    GetTestingComponentsCount = count

End Function



Private Function IsTestComponent(component As VBComponent)

    IsTestComponent = False
    If component.Type = vbext_ct_StdModule Then
        If Left(component.name, Len(TestModulePrefix)) = TestModulePrefix Or _
           Right(component.name, Len(TestModuleSuffix)) = TestModuleSuffix Then
            IsTestComponent = True
        End If
    End If
    
End Function
