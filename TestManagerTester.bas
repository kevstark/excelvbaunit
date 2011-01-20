Attribute VB_Name = "TestManagerTester"
Option Explicit
Option Private Module
Rem order 8
Rem
Rem =head2
Rem sheetname
Rem
Rem This is part of the test suite for the test harness add-in. Its primary purpose is to test
Rem the functionality of the "TestManager" class module.
Rem
Rem =head3
Rem sheetname Macros
Rem
Const ksErrMod As String = "TestManagerTester"

Rem =head4 Function ~
Rem sheetname TestIsTestComponent
Rem
Rem Checks the "IsTestComponent" function with three components.
Rem
Function TestIsTestComponent() As Boolean
On Error GoTo ErrorHandler

    Dim tm As TestManager
    Set tm = New TestManager
    Dim c As VBComponent
    
    If GetNames Then err.Raise knCall, , ksCall
    Set c = Application.VBE.VBProjects(AddInName).VBComponents("DummyTestModule3")
    If AssertFalse(tm.IsTestComponent(c)) Then err.Raise knCall, , ksCall
    
    Set c = Application.VBE.VBProjects(AddInName).VBComponents("TestManager")
    If AssertFalse(tm.IsTestComponent(c)) Then err.Raise knCall, , ksCall
    
    Set c = Application.VBE.VBProjects(AddInName).VBComponents("TestManagerTester")
    If AssertTrue(tm.IsTestComponent(c)) Then err.Raise knCall, , ksCall

Exit Function
ErrorHandler:
ErrTrap ksErrMod, "TestIsTestComponent"
TestIsTestComponent = True
End Function
'Public Sub TestIsTestComponent()
'
'    Dim tm As TestManager
'    Set tm = New TestManager
'    Dim c As VBComponent
'
'    If GetNames Then err.Raise knCall, , ksCall
'    Set c = Application.VBE.VBProjects(AddInName).VBComponents("DummyTestModule3")
'    If AssertFalse(tm.IsTestComponent(c)) Then err.Raise knCall, , ksCall
'
'    Set c = Application.VBE.VBProjects(AddInName).VBComponents("TestManager")
'    If AssertFalse(tm.IsTestComponent(c)) Then err.Raise knCall, , ksCall
'
'    Set c = Application.VBE.VBProjects(AddInName).VBComponents("TestManagerTester")
'    If AssertTrue(tm.IsTestComponent(c)) Then err.Raise knCall, , ksCall
'
'End Sub

Rem =head4 Function ~
Rem sheetname TestGetTestingComponentsCount
Rem
Rem Tests the GetTestingComponentsCount. There should be five components of the test suite,
Rem excluding the dummies but including MainTester, which has not been found on the Google
Rem repository. If more modules are added, the expected value of the test will have to be
Rem changed.
Rem
Function TestGetTestingComponentsCount() As Boolean
On Error GoTo ErrorHandler

    Dim tm As TestManager
    Set tm = New TestManager

    If GetNames Then err.Raise knCall, , ksCall
    Dim p As VBProject
    If AssertEqual(-1, SafeUbound(tm.GetTestingComponents(p))) _
        Then err.Raise knCall, , ksCall
        'Test added by JHD. The project has not been set, so should have no components at this
        'stage. This test ensures that the GetTestingComponents routine is returning the
        'right number in more than one situation.
    Set p = Application.VBE.VBProjects(AddInName)
    'This test will start to fail if you add new test modules to this add in
    If AssertEqual(8, SafeUbound(tm.GetTestingComponents(p)) + 1, _
        "To be expected if you have added new test modules") Then err.Raise knCall, , ksCall

Exit Function
ErrorHandler:
ErrTrap ksErrMod, "TestGetTestingComponentsCount"
TestGetTestingComponentsCount = True
End Function
'Public Sub TestGetTestingComponentsCount()
'
'    Dim tm As TestManager
'    Set tm = New TestManager
'
'    If GetNames Then err.Raise knCall, , ksCall
'    Dim p As VBProject
'    Set p = Application.VBE.VBProjects(AddInName)
'
'    ' This test will start to fail if you add new test modules to this add in
'    If AssertEqual(5, tm.GetTestingComponentsCount(p)) Then err.Raise knCall, , ksCall
'
'End Sub

Rem =head4 Function ~
Rem sheetname TestGetTestingComponents
Rem
Rem Tests the "GetTestingComponents" function to ensure that all test modules' names are
Rem returned. This will need changing if additional test models are added.
Rem sto ModOrder
Rem There is an issue. The tests assume the modules will be returned
Rem in a specific order. Since MainTester did not exist when the original add-in was assembled,
Rem it appears last in the list instead of first as the original tests assumed. Anyone
Rem assembling (rather than downloading complete) their own add-in may face the same issue.
Rem Changing the numbers to match the actual order is a perfectly acceptable way of getting
Rem around the problem. An enhancement would be to check if the module exists anywhere.
Rem sto 0
Rem rcl ModOrder
Rem
Function TestGetTestingComponents() As Boolean
On Error GoTo ErrorHandler

    Dim tm As TestManager
    Set tm = New TestManager

    If GetNames Then err.Raise knCall, , ksCall
    Dim p As VBProject
    Set p = Application.VBE.VBProjects(AddInName)

    Dim cs() As VBComponent
    cs = tm.GetTestingComponents(p)
    
    If AssertEqual(7, UBound(cs), _
        "To be expected if you have added new test modules") Then err.Raise knCall, , ksCall
    If AssertEqual(0, LBound(cs)) Then err.Raise knCall, , ksCall
    'AssertEqual "MainTester", cs(0).name
    'AssertEqual "TestFixtureTester", cs(1).name
    'AssertEqual "TestManagerTester", cs(2).name
    'AssertEqual "TestResultsManagerTester", cs(3).name
    'AssertEqual "TestRunnerTester", cs(4).name
    'JHD: Refactored in 2 ways. 1) Uses InArray instead of relying on specific positions that
    'have been observed to change. 2) Sets up an array of names first to be searched.
    Dim ary() As String
    Dim i As Long
    For i = LBound(cs) To UBound(cs)
        ReDim Preserve ary(LBound(cs) To i)
        ary(i) = cs(i).name
    Next i
    If AssertTrue(InArray("TestFixtureTester", ary)) Then err.Raise knCall, , ksCall
    If AssertTrue(InArray("TestManagerTester", ary)) Then err.Raise knCall, , ksCall
    If AssertTrue(InArray("TestResultsManagerTester", ary)) Then err.Raise knCall, , ksCall
    If AssertTrue(InArray("TestRunnerTester", ary)) Then err.Raise knCall, , ksCall
    If AssertTrue(InArray("MainTester", ary)) Then err.Raise knCall, , ksCall

Exit Function
ErrorHandler:
ErrTrap ksErrMod, "TestGetTestingComponents"
TestGetTestingComponents = True
End Function
'Public Sub TestGetTestingComponents()
'
'    Dim tm As TestManager
'    Set tm = New TestManager
'
'    If GetNames Then err.Raise knCall, , ksCall
'    Dim p As VBProject
'    Set p = Application.VBE.VBProjects(AddInName)
'
'    Dim cs() As VBComponent
'    cs = tm.GetTestingComponents(p)
'
'    If AssertEqual(4, UBound(cs)) Then err.Raise knCall, , ksCall
'    If AssertEqual(0, LBound(cs)) Then err.Raise knCall, , ksCall
'    'AssertEqual "MainTester", cs(0).name
'    'AssertEqual "TestFixtureTester", cs(1).name
'    'AssertEqual "TestManagerTester", cs(2).name
'    'AssertEqual "TestResultsManagerTester", cs(3).name
'    'AssertEqual "TestRunnerTester", cs(4).name
'    If AssertEqual("TestFixtureTester", cs(0).name) Then err.Raise knCall, , ksCall
'    If AssertEqual("TestManagerTester", cs(1).name) Then err.Raise knCall, , ksCall
'    If AssertEqual("TestResultsManagerTester", cs(2).name) Then err.Raise knCall, , ksCall
'    If AssertEqual("TestRunnerTester", cs(3).name) Then err.Raise knCall, , ksCall
'    If AssertEqual("MainTester", cs(4).name) Then err.Raise knCall, , ksCall
'
'End Sub

Rem =head4 Function ~
Rem sheetname TestGetTestFixture
Rem
Rem This tests the "GetTestFixture" procedure by passing a known module to it and comparing
Rem the name of the module returned to the name passed.
Rem
Function TestGetTestFixture() As Boolean
On Error GoTo ErrorHandler

    Dim tm As TestManager
    Set tm = New TestManager
    Dim tf As TestFixture
    
    If GetNames Then err.Raise knCall, , ksCall
    Set tf = tm.GetTestFixture(AddInName, "DummyTestModule")
        
    If AssertEqual("DummyTestModule", tf.fixtureName) Then err.Raise knCall, , ksCall

Exit Function
ErrorHandler:
ErrTrap ksErrMod, "TestGetTestFixture"
TestGetTestFixture = True
End Function
'Public Sub TestGetTestFixture()
'
'    Dim tm As TestManager
'    Set tm = New TestManager
'    Dim tf As TestFixture
'
'    If GetNames Then err.Raise knCall, , ksCall
'    Set tf = tm.GetTestFixture(AddInName, "DummyTestModule")
'
'    If AssertEqual("DummyTestModule", tf.fixtureName) Then err.Raise knCall, , ksCall
'
'End Sub

Rem =head4 Function ~
Rem sheetname TestGetTestFixtures
Rem
Rem This tests the "GetTestFixtures" by ensuring that the five test modules of the add-in's own
Rem test suite are present.
Rem rcl ModOrder
Rem
Function TestGetTestFixtures() As Boolean
On Error GoTo ErrorHandler

    Dim tm As TestManager
    Set tm = New TestManager
    Dim tfs() As TestFixture
    
    If GetNames Then err.Raise knCall, , ksCall
    tfs = tm.GetTestFixtures(AddInName)
    
    If AssertEqual(7, UBound(tfs), _
        "To be expected if you have added new test modules") Then err.Raise knCall, , ksCall
    If AssertEqual(0, LBound(tfs)) Then err.Raise knCall, , ksCall
    
    'AssertEqual "MainTester", tfs(0).fixtureName
    'AssertEqual "TestFixtureTester", tfs(1).fixtureName
    'AssertEqual "TestManagerTester", tfs(2).fixtureName
    'AssertEqual "TestResultsManagerTester", tfs(3).fixtureName
    'AssertEqual "TestRunnerTester", tfs(4).fixtureName
    'JHD: Refactored in 2 ways. 1) Uses InArray instead of relying on specific positions that
    'have been observed to change. 2) Sets up an array of names first to be searched.
    Dim ary() As String
    Dim i As Long
    For i = LBound(tfs) To UBound(tfs)
        ReDim Preserve ary(LBound(tfs) To i)
        ary(i) = tfs(i).fixtureName
    Next i
    If AssertTrue(InArray("TestFixtureTester", ary)) Then err.Raise knCall, , ksCall
    If AssertTrue(InArray("TestManagerTester", ary)) Then err.Raise knCall, , ksCall
    If AssertTrue(InArray("TestResultsManagerTester", ary)) Then err.Raise knCall, , ksCall
    If AssertTrue(InArray("TestRunnerTester", ary)) Then err.Raise knCall, , ksCall
    If AssertTrue(InArray("MainTester", ary)) Then err.Raise knCall, , ksCall

Exit Function
ErrorHandler:
ErrTrap ksErrMod, "TestGetTestFixtures"
TestGetTestFixtures = True
End Function
'Public Sub TestGetTestFixtures()
'
'    Dim tm As TestManager: Set tm = New TestManager
'    Dim tfs() As TestFixture
'
'    If GetNames Then err.Raise knCall, , ksCall
'    tfs = tm.GetTestFixtures(AddInName)
'
'    If AssertEqual(4, UBound(tfs)) Then err.Raise knCall, , ksCall
'    If AssertEqual(0, LBound(tfs)) Then err.Raise knCall, , ksCall
'
'    'AssertEqual "MainTester", tfs(0).fixtureName
'    'AssertEqual "TestFixtureTester", tfs(1).fixtureName
'    'AssertEqual "TestManagerTester", tfs(2).fixtureName
'    'AssertEqual "TestResultsManagerTester", tfs(3).fixtureName
'    'AssertEqual "TestRunnerTester", tfs(4).fixtureName
'    If AssertEqual("TestFixtureTester", tfs(0).fixtureName) Then err.Raise knCall, , ksCall
'    If AssertEqual("TestManagerTester", tfs(1).fixtureName) Then err.Raise knCall, , ksCall
'    If AssertEqual("TestResultsManagerTester", tfs(2).fixtureName) _
'        Then err.Raise knCall, , ksCall
'    If AssertEqual("TestRunnerTester", tfs(3).fixtureName) Then err.Raise knCall, , ksCall
'    If AssertEqual("MainTester", tfs(4).fixtureName) Then err.Raise knCall, , ksCall
'
'End Sub
