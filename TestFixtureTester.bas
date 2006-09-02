Attribute VB_Name = "TestFixtureTester"
Option Explicit

Public Sub TestInvokeProc()

    DummyTestModule3.Reset
    
    Dim f As TestFixture
    Set f = New TestFixture
    
    Dim resultsManager As FakeTestResultsManager
    Set resultsManager = New FakeTestResultsManager
    f.InvokeProc resultsManager, "VbaUnit.xla", "CallMe"

    AssertTrue DummyTestModule3.MeCalled
    AssertEqual ":StartTestCase:EndTestCase", resultsManager.FunctionsCalled
    
End Sub

Public Sub TestRunTests()

    DummyTestModule3.Reset
    
    Dim f As TestFixture
    Set f = New TestFixture
    
    
    Dim c As VBComponent
    Set c = Application.VBE.VBProjects("VbaUnit").VBComponents("DummyTestModule3")
    
    f.ExtractTestCases Application.VBE.VBProjects("VbaUnit"), c
    
    Dim resultsManager As FakeTestResultsManager
    Set resultsManager = New FakeTestResultsManager
    
    f.RunTests resultsManager
    
    AssertFalse DummyTestModule3.MeCalled
    AssertTrue DummyTestModule3.Test1Called
    AssertTrue DummyTestModule3.Test2Called
    AssertEqual ":StartTestCase:EndTestCase:StartTestCase:EndTestCase", resultsManager.FunctionsCalled
    

End Sub


Public Sub TestDoesMethodExist()
    Dim f As TestFixture
    Set f = New TestFixture
    
    Dim c As VBComponent
    Set c = Application.VBE.VBProjects("VbaUnit").VBComponents("DummyTestModule")
    
    AssertFalse f.DoesMethodExist("xxx", c)
    AssertTrue f.DoesMethodExist("NotATest", c)
    AssertTrue f.DoesMethodExist("notatest", c)

End Sub

Public Sub TestExtractTestCases()

    Dim f As TestFixture
    Set f = New TestFixture
    
    Dim c As VBComponent
    Set c = Application.VBE.VBProjects("VbaUnit").VBComponents("DummyTestModule")
    
    f.ExtractTestCases Application.VBE.VBProjects("VbaUnit"), c
    AssertEqual "VbaUnit.xla", f.FileName
    AssertEqual "DummyTestModule", f.FixtureName
        
    Dim s() As String
    s = f.TestProcedures
    
    AssertEqual 1, UBound(s)
    AssertEqual s(0), "DummyTestModule.Test1"
    AssertEqual s(1), "DummyTestModule.Test4"
    
    AssertTrue f.HasSetUpFunction()
    AssertTrue f.HasTearDownFunction()
    
End Sub

Public Sub TestExtractSetUpTearDown()

    Dim f As TestFixture
    Set f = New TestFixture
    
    Dim c As VBComponent
    Set c = Application.VBE.VBProjects("VbaUnit").VBComponents("DummyTestModule2")
    
    f.ExtractTestCases Application.VBE.VBProjects("VbaUnit"), c
    
    AssertTrue f.HasSetUpFunction()
    AssertFalse f.HasTearDownFunction()
    
End Sub


Public Sub TestExtractFileName()

    Dim f As TestFixture
    Set f = New TestFixture
    
    Dim name As String
    name = f.ExtractFileName("c:\test1\test2\test3.xls")
    AssertEqual "test3.xls", name

    name = f.ExtractFileName("test4.xls")
    AssertEqual "test4.xls", name

    name = f.ExtractFileName("test5")
    AssertEqual "test5", name
End Sub



Public Sub TestIsTestMethodLine()

    Dim f As TestFixture
    Set f = New TestFixture
    
    AssertFalse f.IsTestMethodLine("public sub test1")
    AssertTrue f.IsTestMethodLine("Public Sub Test2")
    AssertFalse f.IsTestMethodLine("Public Sub Tst")
    
    
End Sub

Public Sub TestGetTestMethodsCount()

    Dim f As TestFixture
    Set f = New TestFixture
    
    Dim c As VBComponent
    Set c = Application.VBE.VBProjects("VbaUnit").VBComponents("DummyTestModule")
    
    Dim i As Integer
    i = f.GetTestMethodsCount(c)
    
    AssertEqual 2, i
    
End Sub

Public Sub TestGetTestMethods()

    Dim f As TestFixture
    Set f = New TestFixture
    
    Dim c As VBComponent
    Set c = Application.VBE.VBProjects("VbaUnit").VBComponents("DummyTestModule")
    
    Dim s() As String
    s = f.GetTestMethods(c, 2)
    
    AssertEqual 1, UBound(s)
    AssertEqual s(0), "DummyTestModule.Test1"
    AssertEqual s(1), "DummyTestModule.Test4"

End Sub
