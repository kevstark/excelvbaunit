Attribute VB_Name = "TestRunnerTester"
Option Explicit

Public Sub TestRun()

    Dim tr As TestRunner: Set tr = New TestRunner
    Dim trm As FakeTestResultsManager: Set trm = New FakeTestResultsManager
    
    tr.Run "VbaUnit", trm, "DummyTestModule4"
    
    AssertEqual ":StartTestCase(DummyTestModule4.Test1):EndTestCase:EndTestSuite", trm.FunctionsCalled
    
End Sub

Public Sub TestShouldRunFixture()

    Dim tr As TestRunner: Set tr = New TestRunner
    Dim tf As TestFixture: Set tf = New TestFixture
    
    Dim c As VBComponent
    Set c = Application.VBE.VBProjects("VbaUnit").VBComponents("DummyTestModule")
    
    tf.ExtractTestCases Application.VBE.VBProjects("VbaUnit"), c
    
    AssertTrue tr.ShouldRunFixture(tf, "")
    AssertTrue tr.ShouldRunFixture(tf, "DummyTestModule")
    AssertFalse tr.ShouldRunFixture(tf, "xxx")

End Sub
