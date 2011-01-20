Attribute VB_Name = "TestFixtureTester"
Option Explicit
Option Private Module
Rem order 7
Rem
Rem =head2
Rem sheetname
Rem
Rem This is part of the test suite for the test harness add-in. Its primary purpose is to test
Rem the functionality of the "TestFixture" class module.
Rem
Rem =head3
Rem sheetname Macros
Rem
Const ksErrMod As String = "TestFixtureTester"

Rem =head4 Function ~
Rem sheetname TestInvokeProc
Rem
Rem Uses DummyTestModule3 to test that the macros in a test module were indeed called and that
Rem the right functions were called.
Rem
'Public Sub TestInvokeProc()
'
'    DummyTestModule3.Reset
'
'    Dim f As TestFixture
'    Set f = New TestFixture
'
'    Dim resultsManager As FakeTestResultsManager
'    Set resultsManager = New FakeTestResultsManager
'    If GetNames Then err.Raise knCall, , ksCall
'    f.InvokeProc resultsManager, AddInFileName, "CallMe"
'
'    AssertTrue DummyTestModule3.MeCalled
'    AssertEqual ":StartTestCase(CallMe):EndTestCase", resultsManager.FunctionsCalled
'
'End Sub

Function TestInvokeProc() As Boolean
On Error GoTo ErrorHandler

    If DummyTestModule3.Reset Then err.Raise knCall, , ksCall
    
    Dim f As TestFixture
    Set f = New TestFixture
    
    Dim resultsManager As FakeTestResultsManager
    Set resultsManager = New FakeTestResultsManager
    If GetNames Then err.Raise knCall, , ksCall
    If f.InvokeProc(resultsManager, AddInFileName, "CallMe") Then err.Raise knCall, , ksCall

    If AssertTrue(DummyTestModule3.MeCalled) Then err.Raise knCall, , ksCall
    If AssertEqual(":StartTestCase(CallMe):EndTestCase", resultsManager.FunctionsCalled) _
        Then err.Raise knCall, , ksCall

Exit Function
ErrorHandler:
ErrTrap ksErrMod, "TestInvokeProc"
TestInvokeProc = True
End Function

Rem =head4 Function ~
Rem sheetname TestRunTests
Rem
Rem Uses DummyTestModule3 to test that the Setup, TearDown and Test procedures are called and
Rem that others that are not tests are not called.
Rem
Function TestRunTests() As Boolean
On Error GoTo ErrorHandler

    If DummyTestModule3.Reset Then err.Raise knCall, , ksCall
    
    Dim f As TestFixture
    Set f = New TestFixture
    
    If GetNames Then err.Raise knCall, , ksCall
    Dim c As VBComponent
    Set c = Application.VBE.VBProjects(AddInName).VBComponents("DummyTestModule3")
    
    f.ExtractTestCases Application.VBE.VBProjects(AddInName), c
    
    Dim resultsManager As FakeTestResultsManager
    Set resultsManager = New FakeTestResultsManager
    
    If f.RunTests(resultsManager) Then err.Raise knCall, , ksCall
    
    If AssertFalse(DummyTestModule3.MeCalled) Then err.Raise knCall, , ksCall
    If AssertTrue(DummyTestModule3.Test1Called) Then err.Raise knCall, , ksCall
    If AssertTrue(DummyTestModule3.Test2Called) Then err.Raise knCall, , ksCall
    If AssertEqual(":StartTestCase(DummyTestModule3.Test1):EndTestCase:StartTestCase(DummyTestModule3.Test2):EndTestCase", _
        resultsManager.FunctionsCalled) Then err.Raise knCall, , ksCall
    
    If AssertEqual(1, DummyTestModule3.SetUpCalled) Then err.Raise knCall, , ksCall
    If AssertEqual(1, DummyTestModule3.TearDownCalled) Then err.Raise knCall, , ksCall

Exit Function
ErrorHandler:
ErrTrap ksErrMod, "TestRunTests"
TestRunTests = True
End Function
'Rem =head4 Function TestRunTests
'Rem
'Rem Uses DummyTestModule3 to test that the Setup, TearDown and Test procedures are called and
'Rem that others that are not tests are not called.
'Rem
'Rem rcl ToFn
'Rem
'Public Sub TestRunTests()
'
'    DummyTestModule3.Reset
'
'    Dim f As TestFixture
'    Set f = New TestFixture
'
'    If GetNames Then err.Raise knCall, , ksCall
'    Dim c As VBComponent
'    Set c = Application.VBE.VBProjects(AddInName).VBComponents("DummyTestModule3")
'
'    f.ExtractTestCases Application.VBE.VBProjects(AddInName), c
'
'    Dim resultsManager As FakeTestResultsManager
'    Set resultsManager = New FakeTestResultsManager
'
'    f.RunTests resultsManager
'
'    AssertFalse DummyTestModule3.MeCalled
'    AssertTrue DummyTestModule3.Test1Called
'    AssertTrue DummyTestModule3.Test2Called
'    AssertEqual ":StartTestCase(DummyTestModule3.Test1):EndTestCase:StartTestCase(DummyTestModule3.Test2):EndTestCase", resultsManager.FunctionsCalled
'
'    AssertEqual 2, DummyTestModule3.SetUpCalled
'    AssertEqual 2, DummyTestModule3.TearDownCalled
'
'End Sub

Rem =head4 Function ~
Rem sheetname TestDoesMethodExist
Rem
Rem Uses DummyTestModule to test for the presence and absence of various procedures.
Rem
Function TestDoesMethodExist() As Boolean
On Error GoTo ErrorHandler

    Dim f As TestFixture
    Set f = New TestFixture
    
    If GetNames Then err.Raise knCall, , ksCall
    Dim c As VBComponent
    Set c = Application.VBE.VBProjects(AddInName).VBComponents("DummyTestModule")
    
    If AssertFalse(f.DoesMethodExist("xxx", c)) Then err.Raise knCall, , ksCall
    If AssertTrue(f.DoesMethodExist("NotATest", c)) Then err.Raise knCall, , ksCall
    If AssertTrue(f.DoesMethodExist("notatest", c)) Then err.Raise knCall, , ksCall

Exit Function
ErrorHandler:
ErrTrap ksErrMod, "TestDoesMethodExist"
TestDoesMethodExist = True
End Function
'Public Sub TestDoesMethodExist()
'    Dim f As TestFixture
'    Set f = New TestFixture
'
'    If GetNames Then err.Raise knCall, , ksCall
'    Dim c As VBComponent
'    Set c = Application.VBE.VBProjects(AddInName).VBComponents("DummyTestModule")
'
'    If AssertFalse(f.DoesMethodExist("xxx", c)) Then err.Raise knCall, , ksCall
'    If AssertTrue(f.DoesMethodExist("NotATest", c)) Then err.Raise knCall, , ksCall
'    If AssertTrue(f.DoesMethodExist("notatest", c)) Then err.Raise knCall, , ksCall
'
'End Sub

Rem =head4 Function ~
Rem sheetname TestExtractTestCases
Rem
Rem Uses DummyTestModule to test that the ExtractTestCases code returns the correct number and
Rem names of test routines.
Rem
Function TestExtractTestCases() As Boolean
On Error GoTo ErrorHandler

    Dim f As TestFixture
    Set f = New TestFixture
    
    If GetNames Then err.Raise knCall, , ksCall
    Dim c As VBComponent
    Set c = Application.VBE.VBProjects(AddInName).VBComponents("DummyTestModule")
    
    f.ExtractTestCases Application.VBE.VBProjects(AddInName), c
    If AssertEqual(AddInFileName, f.FileName) Then err.Raise knCall, , ksCall
    If AssertEqual("DummyTestModule", f.fixtureName) Then err.Raise knCall, , ksCall
        
    Dim s() As String
    s = f.TestProcedures
    
    'AssertEqual 1, UBound(s) 'JHD: Amended because functions are now allowed.
    If AssertEqual(2, UBound(s)) Then err.Raise knCall, , ksCall
    'AssertEqual s(0), "DummyTestModule.Test1" JHD: Wrong way round in original
    If AssertEqual("DummyTestModule.Test1", s(0)) Then err.Raise knCall, , ksCall
    'AssertEqual s(1), "DummyTestModule.Test4" JHD: Wrong way round in original
    If AssertEqual("DummyTestModule.Test3", s(1)) Then err.Raise knCall, , ksCall
    
    If AssertTrue(f.HasSetUpFunction()) Then err.Raise knCall, , ksCall
    If AssertTrue(f.HasTearDownFunction()) Then err.Raise knCall, , ksCall

Exit Function
ErrorHandler:
ErrTrap ksErrMod, "TestExtractTestCases"
TestExtractTestCases = True
End Function
'Public Sub TestExtractTestCases()
'
'    Dim f As TestFixture
'    Set f = New TestFixture
'
'    If GetNames Then err.Raise knCall, , ksCall
'    Dim c As VBComponent
'    Set c = Application.VBE.VBProjects(AddInName).VBComponents("DummyTestModule")
'
'    f.ExtractTestCases Application.VBE.VBProjects(AddInName), c
'    If AssertEqual(AddInFileName, f.FileName) Then err.Raise knCall, , ksCall
'    If AssertEqual("DummyTestModule", f.fixtureName) Then err.Raise knCall, , ksCall
'
'    Dim s() As String
'    s = f.TestProcedures
'
'    'AssertEqual 1, UBound(s) 'JHD: Amended because functions are now allowed.
'    If AssertEqual(2, UBound(s)) Then err.Raise knCall, , ksCall
'    'AssertEqual s(0), "DummyTestModule.Test1" JHD: Wrong way round in original
'    If AssertEqual("DummyTestModule.Test1", s(0)) Then err.Raise knCall, , ksCall
'    'AssertEqual s(1), "DummyTestModule.Test4" JHD: Wrong way round in original
'    If AssertEqual("DummyTestModule.Test3", s(1)) Then err.Raise knCall, , ksCall
'
'    If AssertTrue(f.HasSetUpFunction()) Then err.Raise knCall, , ksCall
'    If AssertTrue(f.HasTearDownFunction()) Then err.Raise knCall, , ksCall
'
'End Sub

Rem =head4 Function ~
Rem sheetname TestExtractSetUpTearDown
Rem
Rem Original documentation per MH: Tests for checking test case and test fixture level
Rem set up/tear down functions.
Rem
Rem JHD: This procedure does not have a single direct equivalent in TestFixture. It is used to
Rem test several of the Property Let / Get procedures that cannot be tested by their own
Rem test procedures. This will therefore be reported as unmatched by the coverage analysis.
Rem
' Tests for checking test case and test fixture level set up/tear down functions
Function TestExtractSetUpTearDown() As Boolean
On Error GoTo ErrorHandler

    Dim f As TestFixture
    Set f = New TestFixture
    
    If GetNames Then err.Raise knCall, , ksCall
    Dim c As VBComponent
    Set c = Application.VBE.VBProjects(AddInName).VBComponents("DummyTestModule2")
    
    f.ExtractTestCases Application.VBE.VBProjects(AddInName), c
    
    If AssertTrue(f.HasSetUpFunction()) Then err.Raise knCall, , ksCall
    If AssertFalse(f.HasTearDownFunction()) Then err.Raise knCall, , ksCall
    
    If AssertTrue(f.HasFixtureSetUpFunction()) Then err.Raise knCall, , ksCall
    If AssertFalse(f.HasFixtureTearDownFunction()) Then err.Raise knCall, , ksCall

Exit Function
ErrorHandler:
ErrTrap ksErrMod, "TestExtractSetUpTearDown"
TestExtractSetUpTearDown = True
End Function
'Public Sub TestExtractSetUpTearDown()
'
'    Dim f As TestFixture
'    Set f = New TestFixture
'
'    If GetNames Then err.Raise knCall, , ksCall
'    Dim c As VBComponent
'    Set c = Application.VBE.VBProjects(AddInName).VBComponents("DummyTestModule2")
'
'    f.ExtractTestCases Application.VBE.VBProjects(AddInName), c
'
'    If AssertTrue(f.HasSetUpFunction()) Then err.Raise knCall, , ksCall
'    If AssertFalse(f.HasTearDownFunction()) Then err.Raise knCall, , ksCall
'
'    If AssertTrue(f.HasFixtureSetUpFunction()) Then err.Raise knCall, , ksCall
'    If AssertFalse(f.HasFixtureTearDownFunction()) Then err.Raise knCall, , ksCall
'
'End Sub

Rem =head4 Function ~
Rem sheetname TestExtractFileName
Rem
Rem Tests the "ExtractFileName" code with various formats of file name, including and excluding
Rem paths and extensions.
Rem
Function TestExtractFileName() As Boolean
On Error GoTo ErrorHandler

    Dim f As TestFixture
    Set f = New TestFixture
    
    Dim name As String
    name = f.ExtractFileName("c:\test1\test2\test3.xls")
    If AssertEqual("test3.xls", name) Then err.Raise knCall, , ksCall

    name = f.ExtractFileName("test4.xls")
    If AssertEqual("test4.xls", name) Then err.Raise knCall, , ksCall

    name = f.ExtractFileName("test5")
    If AssertEqual("test5", name) Then err.Raise knCall, , ksCall

Exit Function
ErrorHandler:
ErrTrap ksErrMod, "TestExtractFileName"
TestExtractFileName = True
End Function
'Public Sub TestExtractFileName()
'
'    Dim f As TestFixture
'    Set f = New TestFixture
'
'    Dim name As String
'    name = f.ExtractFileName("c:\test1\test2\test3.xls")
'    If AssertEqual("test3.xls", name) Then err.Raise knCall, , ksCall
'
'    name = f.ExtractFileName("test4.xls")
'    If AssertEqual("test4.xls", name) Then err.Raise knCall, , ksCall
'
'    name = f.ExtractFileName("test5")
'    If AssertEqual("test5", name) Then err.Raise knCall, , ksCall
'End Sub

Rem =head4 Function ~
Rem sheetname TestIsTestMethodLine
Rem
Rem Tests the "IsTestMethodLine" function with three pseudo-lines of code.
Rem
Function TestIsTestMethodLine() As Boolean
On Error GoTo ErrorHandler

    Dim f As TestFixture
    Set f = New TestFixture
    
    If AssertFalse(f.IsTestMethodLine("public sub test1")) Then err.Raise knCall, , ksCall
    If AssertTrue(f.IsTestMethodLine("Public Sub Test2")) Then err.Raise knCall, , ksCall
    If AssertFalse(f.IsTestMethodLine("Public Sub Tst")) Then err.Raise knCall, , ksCall

Exit Function
ErrorHandler:
ErrTrap ksErrMod, "TestIsTestMethodLine"
TestIsTestMethodLine = True
End Function
'Public Sub TestIsTestMethodLine()
'
'    Dim f As TestFixture
'    Set f = New TestFixture
'
'    If AssertFalse(f.IsTestMethodLine("public sub test1")) Then err.Raise knCall, , ksCall
'    If AssertTrue(f.IsTestMethodLine("Public Sub Test2")) Then err.Raise knCall, , ksCall
'    If AssertFalse(f.IsTestMethodLine("Public Sub Tst")) Then err.Raise knCall, , ksCall
'
'
'End Sub

'Public Sub TestGetTestMethodsCount() - commented out because the GetCount routine has been
'                                       refactored out
'
'    Dim f As TestFixture
'    Set f = New TestFixture
'
'    Dim c As VBComponent
'    Set c = Application.VBE.VBProjects(AddInName).VBComponents("DummyTestModule")
'
'    Dim i As Integer
'    i = f.GetTestMethodsCount(c)
'
'    AssertEqual 3, i 'JHD: Amended because functions are now allowed.
'
'End Sub

Rem =head4 Function ~
Rem sheetname TestGetTestMethods
Rem
Rem Tests the GetTestMethods function using DummyTestModule. The correct names and numbers of
Rem procedures should be returned. Refactored to take account of changes to the system.
Rem
Function TestGetTestMethods() As Boolean
On Error GoTo ErrorHandler

    Dim f As TestFixture
    Set f = New TestFixture
    
    If GetNames Then err.Raise knCall, , ksCall
    Dim c As VBComponent
    Set c = Application.VBE.VBProjects(AddInName).VBComponents("DummyTestModule")
    
    Dim s() As String
    's = f.GetTestMethods(c, 2)
    'JHD: The previous line caused crashes because there are three sub tests in DummyTestModule.
    'Since there is also a Function Test* and the code has been modified to allow this, the
    'test will return four anyway. Since the call attempts to set the count manually, this
    'results in a subscript out of range error when three or four are found. However, the code
    'has now been refactored to avoid the use of a count function and to get all methods. This
    'will prevent the bug without invalidating the tests below (although the count has changed
    'now that functions are allowed).
    's = f.GetTestMethods(c, 3)
    s = f.GetTestMethods(c)
    
    'AssertEqual 1, UBound(s)
    'AssertEqual s(0), "DummyTestModule.Test1"
    'AssertEqual s(1), "DummyTestModule.Test4"
    If AssertEqual(2, UBound(s)) Then err.Raise knCall, , ksCall
    'If AssertEqual("DummyTestModule.Test1", s(0)) Then err.Raise knCall, , ksCall
    If AssertTrue(InArray("DummyTestModule.Test1", s)) Then err.Raise knCall, , ksCall
    'If AssertEqual("DummyTestModule.Test4", s(2)) Then err.Raise knCall, , ksCall
    If AssertTrue(InArray("DummyTestModule.Test4", s)) Then err.Raise knCall, , ksCall

Exit Function
ErrorHandler:
ErrTrap ksErrMod, "TestGetTestMethods"
TestGetTestMethods = True
End Function
'Public Sub TestGetTestMethods()
'
'    Dim f As TestFixture
'    Set f = New TestFixture
'
'    If GetNames Then err.Raise knCall, , ksCall
'    Dim c As VBComponent
'    Set c = Application.VBE.VBProjects(AddInName).VBComponents("DummyTestModule")
'
'    Dim s() As String
'    's = f.GetTestMethods(c, 2)
'    'JHD: The previous line caused crashes because there are three sub tests in DummyTestModule.
'    'Since there is also a Function Test* and the code has been modified to allow this, the
'    'test will return four anyway. Since the call attempts to set the count manually, this
'    'results in a subscript out of range error when three or four are found. However, the code
'    'has now been refactored to avoid the use of a count function and to get all methods. This
'    'will prevent the bug without invalidating the tests below (although the count has changed
'    'now that functions are allowed).
'    's = f.GetTestMethods(c, 3)
'    s = f.GetTestMethods(c)
'
'    'AssertEqual 1, UBound(s)
'    'AssertEqual s(0), "DummyTestModule.Test1"
'    'AssertEqual s(1), "DummyTestModule.Test4"
'    If AssertEqual(2, UBound(s)) Then err.Raise knCall, , ksCall
'    If AssertEqual("DummyTestModule.Test1", s(0)) Then err.Raise knCall, , ksCall
'    If AssertEqual("DummyTestModule.Test4", s(2)) Then err.Raise knCall, , ksCall
'
'End Sub
