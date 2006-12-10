Attribute VB_Name = "TestManagerTester"
Option Explicit

Public Sub TestIsTestComponent()

    Dim tm As TestManager: Set tm = New TestManager

    Dim c As VBComponent
    Set c = Application.VBE.VBProjects("VbaUnit").VBComponents("DummyTestModule3")
    
    AssertFalse tm.IsTestComponent(c)
    
    Set c = Application.VBE.VBProjects("VbaUnit").VBComponents("TestManager")
    AssertFalse tm.IsTestComponent(c)
    
    Set c = Application.VBE.VBProjects("VbaUnit").VBComponents("TestManagerTester")
    AssertTrue tm.IsTestComponent(c)
    
    
End Sub


Public Sub TestGetTestingComponentsCount()

    Dim tm As TestManager: Set tm = New TestManager

    Dim p As VBProject
    Set p = Application.VBE.VBProjects("VbaUnit")

    ' This test will start to fail if you add new test modules to this add in
    AssertEqual 5, tm.GetTestingComponentsCount(p)

End Sub

Public Sub TestGetTestingComponents()

    Dim tm As TestManager: Set tm = New TestManager

    Dim p As VBProject
    Set p = Application.VBE.VBProjects("VbaUnit")

    Dim cs() As VBComponent
    cs = tm.GetTestingComponents(p)
    
    AssertEqual 4, UBound(cs)
    AssertEqual 0, LBound(cs)
    AssertEqual "MainTester", cs(0).name
    AssertEqual "TestFixtureTester", cs(1).name
    AssertEqual "TestManagerTester", cs(2).name
    AssertEqual "TestResultsManagerTester", cs(3).name
    AssertEqual "TestRunnerTester", cs(4).name

End Sub

Public Sub TestGetTestFixture()

    Dim tm As TestManager: Set tm = New TestManager
    Dim tf As TestFixture
    
    Set tf = tm.GetTestFixture("VbaUnit", "DummyTestModule")
        
    AssertEqual "DummyTestModule", tf.fixtureName

End Sub

Public Sub TestGetTestFixtures()

    Dim tm As TestManager: Set tm = New TestManager
    Dim tfs() As TestFixture
    
    tfs = tm.GetTestFixtures("VbaUnit")
    
    AssertEqual 4, UBound(tfs)
    AssertEqual 0, LBound(tfs)
    
    AssertEqual "MainTester", tfs(0).fixtureName
    AssertEqual "TestFixtureTester", tfs(1).fixtureName
    AssertEqual "TestManagerTester", tfs(2).fixtureName
    AssertEqual "TestResultsManagerTester", tfs(3).fixtureName
    AssertEqual "TestRunnerTester", tfs(4).fixtureName
    
    
End Sub
