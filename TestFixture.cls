VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestFixture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Rem order 17
Rem
Rem =head2
Rem sheetname
Rem
Rem This class module contains variables indicating what procedures a module contains
Rem and code to calculate their values as well as to run the tests in that module.
Rem
Rem =head3
Rem sheetname Macros
Rem

' Control the list of tests within a test fixtures, ie with a test module
'Const TestMethodPrefix = "Public Sub Test" Changed by JHD to allow functions or subs not
                                           'explicitly declared as Public.
Const TestMethodPrefix              As String = "Test"

Private mTestProcedures()           As String
Private mFileName                   As String
Private mFixtureName                As String
Private mHasSetUpFunction           As Boolean
Private mHasTearDownFunction        As Boolean
Private mHasFixtureSetUpFunction    As Boolean
Private mHasFixtureTearDownFunction As Boolean
Const ksErrMod                      As String = "TestFixture"

Rem =head4 Property Get ~
Rem sheetname TestProcedures
Rem
Property Get TestProcedures() As String()
    TestProcedures = mTestProcedures
End Property

Rem =head4 Property Get ~
Rem sheetname FileName
Rem
Property Get FileName() As String
    FileName = mFileName
End Property

Rem =head4 Property Get ~
Rem sheetname fixtureName
Rem
Property Get fixtureName() As String
    fixtureName = mFixtureName
End Property

Rem =head4 Property Get ~
Rem sheetname HasSetUpFunction
Rem
Friend Property Get HasSetUpFunction() As Boolean
    HasSetUpFunction = mHasSetUpFunction
End Property

Rem =head4 Property Get ~
Rem sheetname HasTearDownFunction
Rem
Friend Property Get HasTearDownFunction() As Boolean
    HasTearDownFunction = mHasTearDownFunction
End Property

Rem =head4 Property Get ~
Rem sheetname HasFixtureSetUpFunction
Rem
Friend Property Get HasFixtureSetUpFunction() As Boolean
    HasFixtureSetUpFunction = mHasFixtureSetUpFunction
End Property

Rem =head4 Property Get ~
Rem sheetname HasFixtureTearDownFunction
Rem
Friend Property Get HasFixtureTearDownFunction() As Boolean
    HasFixtureTearDownFunction = mHasFixtureTearDownFunction
End Property

Rem =head4 Function ~
Rem sheetname RunTests
Rem
Rem Runs all tests in a fixture (module), running the SetUp and TearDown functions before
Rem and after. Changed by JHD to run SetUp and TearDown once only, instead of for every
Rem iteration. FixtureSetUp and FixtureTearDown are not run - perhaps these can be
Rem refactored out. A search reveals no part of the code where they might be run, although
Rem stubs exist in the dummy modules. Inserting "AssertSuccess" or "AssertFailure" into
Rem DummytestModule2.FixtureSetUp generates no failures in the tests. Putting a breakpoint on
Rem the relevant line indicates that the line is never called. This is the only instance of
Rem FixtureSetUp in the system. There are no instances of FixtureTearDown. Added feature:
Rem warns if there is a SetUp but no TearDown, as the original system might be
Rem damaged by the process of testing if repeated changes are made without being undone.
Rem
Function RunTests(resultsManager As ITestResultsManager) As Boolean
On Error GoTo ErrorHandler

    If resultsManager.StartTestFixture(mFixtureName) Then err.Raise knCall, , ksCall
    If HasSetUpFunction Then   'JHD: moved outside loop so that it isn't run on every iteration
                               'Moving it required a change in the expected values in
                               'TestFixtureTester.TestRunTests, since this and the TearDown
                               'function are now run once instead of twice (or whatever number
                               'of tests are in DummyTestModule3).
        Dim bRunSetUp As Boolean
        bRunSetUp = False
        If HasTearDownFunction Then
            bRunSetUp = True
        Else
            If (Not bRunSetUp And MsgBox(mFixtureName & " has a SetUp function but no TearDown." _
                & vbCrLf & "Are you sure you want to run it?", vbYesNo) = vbYes) _
                Then bRunSetUp = True
        End If
        If bRunSetUp Then _
            Application.Run ("'" & FileName & "'!" & mFixtureName & "." & ksSetUpFunctionName)
    End If
    
    Dim i As Integer
    For i = 0 To SafeUbound(mTestProcedures)
'        If HasSetUpFunction Then
'            Application.Run ("'" & FileName & "'!" & mFixtureName & "." & SetUpFunctionName)
'        End If
        
        If InvokeProc(resultsManager, mFileName, mTestProcedures(i)) Then err.Raise knCall, , ksCall
        
'        If HasTearDownFunction Then
'            Application.Run ("'" & FileName & "'!" & mFixtureName & "." & TearDownFunctionName)
'        End If
        
    Next
    
    If HasTearDownFunction Then 'JHD: moved outside loop so that it isn't run on every iteration
        Application.Run ("'" & FileName & "'!" & mFixtureName & "." & ksTearDownFunctionName)
    End If
    
    If resultsManager.EndTestFixture() Then err.Raise knCall, , ksCall

Exit Function
ErrorHandler:
ErrTrap ksErrMod, "RunTests"
RunTests = True
End Function
'Public Sub RunTests(resultsManager As ITestResultsManager)
'
'    If resultsManager.StartTestFixture(mFixtureName) Then err.Raise knCall, , ksCall
'
'    Dim i As Integer
'    For i = 0 To SafeUbound(mTestProcedures)
'        If HasSetUpFunction Then
'            Application.Run ("'" & FileName & "'!" & mFixtureName & "." & SetUpFunctionName)
'        End If
'
'        InvokeProc resultsManager, mFileName, mTestProcedures(i)
'
'        If HasTearDownFunction Then
'            Application.Run ("'" & FileName & "'!" & mFixtureName & "." & TearDownFunctionName)
'        End If
'
'    Next
'
'    If resultsManager.EndTestFixture() Then err.Raise knCall, , ksCall
'
'End Sub

Rem =head4 Function ~
Rem sheetname InvokeProc
Rem
Rem This will invoke any procedure, function or sub, but wraps in a results manager's
Rem StartTestCase and EndTestCase routines. No parameters are passed to the procedure
Rem unless they are in the string passed as a parameter to this routine that gives
Rem the parameters to be passed. In other words, parameters are passed ByVal and are
Rem the responsibility of the calling routine. However, the routine actually called - the test
Rem case - can pass whatever parameters it chooses to whatever procedures it chooses without
Rem restriction. The idea is that this routine will call the Test* procedures in a module
Rem and that these will call procedures to be tested with the necessary parameters. It's
Rem quite difficult to see why one would want a test procedure to receive parameters, but
Rem sooner or later, someone will find an application.
Rem
Friend Function InvokeProc(resultsManager As ITestResultsManager, _
                           FileName As String, _
                           proc As String) As Boolean
On Error GoTo ErrorHandler

    If resultsManager.StartTestCase(proc) Then err.Raise knCall, , ksCall
    Application.Run ("'" & FileName & "'!" & proc)
    If resultsManager.EndTestCase() Then err.Raise knCall, , ksCall

Exit Function
ErrorHandler:
ErrTrap ksErrMod, "InvokeProc"
InvokeProc = True
End Function
'Friend Sub InvokeProc(resultsManager As ITestResultsManager, FileName As String, proc As String)
'
'    If resultsManager.StartTestCase(proc) Then err.Raise knCall, , ksCall
'    Application.Run ("'" & FileName & "'!" & proc)
'    If resultsManager.EndTestCase() Then err.Raise knCall, , ksCall
'
'End Sub

' Extracts the test cases from a test fixture
Rem =head4 Function ~
Rem sheetname ExtractTestCases
Rem
Rem Extracts the test cases from a test fixture (a procedure), populating them into the
Rem mTestProcedures array of this class. Also populates the booleans that indicate whether
Rem certain procedures exist.
Rem
Function ExtractTestCases(project As VBProject, component As VBComponent) As Boolean
On Error GoTo ErrorHandler

    mTestProcedures = GetTestMethods(component)
    
    mFileName = ExtractFileName(project.FileName)
    mFixtureName = component.name
    
    mHasSetUpFunction = DoesMethodExist(ksSetUpFunctionName, component)
    mHasTearDownFunction = DoesMethodExist(ksTearDownFunctionName, component)
    
    mHasFixtureSetUpFunction = DoesMethodExist(ksFixtureSetUpFunctionName, component)
    mHasFixtureTearDownFunction = DoesMethodExist(ksFixtureTearDownFunctionName, component)

Exit Function
ErrorHandler:
ErrTrap ksErrMod, "ExtractTestCases"
ExtractTestCases = True
End Function
'Public Sub ExtractTestCases(project As VBProject, component As VBComponent)
'
'    'Dim count As Integer
'    'count = GetTestMethodsCount(component)
'
'    'mTestProcedures = GetTestMethods(component, count)
'    mTestProcedures = GetTestMethods(component)
'
'    mFileName = ExtractFileName(project.FileName)
'    mFixtureName = component.name
'
'    mHasSetUpFunction = DoesMethodExist(SetUpFunctionName, component)
'    mHasTearDownFunction = DoesMethodExist(TearDownFunctionName, component)
'
'    mHasFixtureSetUpFunction = DoesMethodExist(FixtureSetUpFunctionName, component)
'    mHasFixtureTearDownFunction = DoesMethodExist(FixtureTearDownFunctionName, component)
'
'End Sub

Rem =head4 Function ~
Rem sheetname DoesMethodExist
Rem
Rem Identifies whether a procedure exists. It must be public, but since public is the
Rem default, it does not have to be stated explicitly. It must be a sub or function, i.e. not
Rem a "Property *" procedure. It can't be a "Friend" or "Private" function or sub. The reason
Rem for this strange set of restrictions is that the purpose is to find the SetUp and
Rem TearDown functions that will always adhere to these restrictions. If you are planning to
Rem use it for any other purpose, you are advised to write your own function under a different
Rem name. Base it on this, by all means. But changing this, unless MH and JHD have created a
Rem monster, is a Bad Idea(tm).
Rem
Friend Function DoesMethodExist(ByVal name As String, component As VBComponent) As Boolean
On Error GoTo ErrorHandler

    Dim line As Integer
    Dim str As String
    Dim searchFor As String
    name = LCase(name)
    For line = 1 To component.CodeModule.CountOfLines
        str = LCase(component.CodeModule.Lines(line, 1))
        searchFor = ""
        Select Case Left(str, 1)
        Case "p" 'might be public
            Select Case Mid(str, 8, 1)
            Case "s"
                searchFor = "public sub "
            Case "f"
                searchFor = "public function "
            End Select
        Case "s" 'might be sub
            searchFor = "sub "
        Case "f" 'might be function
            searchFor = "function "
        End Select
        If Len(searchFor) > 0 Then
            searchFor = searchFor & name
            If Left(str, Len(searchFor)) = searchFor Then
                DoesMethodExist = True
                Exit Function
            End If
        End If
        
    Next line
    
    DoesMethodExist = False

Exit Function
ErrorHandler:
ErrTrap ksErrMod, "DoesMethodExist"
End Function
'Friend Function DoesMethodExist(ByVal name As String, component As VBComponent) As Boolean
'
'    Dim line As Integer
'    Dim str As String
'    Dim searchFor As String
'    name = LCase(name) 'JHD: avoids converting case for every line
''    searchFor = LCase("Public Sub " & name)
'    For line = 1 To component.CodeModule.CountOfLines
'        str = LCase(component.CodeModule.Lines(line, 1))
'        'JHD: We can't continue with the simple "searchFor" approach, as that precludes both
'        'functions and subs that aren't specifically declared as Public, which is the default.
'        searchFor = ""
'        Select Case Left(str, 1)
'        Case "p" 'might be public
'            Select Case Mid(str, 8, 1)
'            Case "s"
'                searchFor = "public sub "
'            Case "f"
'                searchFor = "public function "
'            End Select
'        Case "s" 'might be sub
'            searchFor = "sub "
'        Case "f" 'might be function
'            searchFor = "function "
'        End Select
'        If Len(searchFor) > 0 Then
'            searchFor = searchFor & name
'            If Left(str, Len(searchFor)) = searchFor Then
'                DoesMethodExist = True
'                Exit Function
'            End If
'        End If
'
'    Next line
'
'    DoesMethodExist = False
'
'End Function

Rem =head4 Function ~
Rem sheetname ExtractFileName
Rem
Rem Given a fully qualifies file name, extracts solely the file name.
Rem
Friend Function ExtractFileName(FullPath As String) As String
On Error GoTo ErrorHandler

    Dim parts() As String
    parts = Split(FullPath, "\")
    
    ExtractFileName = parts(UBound(parts))

Exit Function
ErrorHandler:
ErrTrap ksErrMod, "ExtractFileName"
End Function

'Friend Function GetTestMethods(component As VBComponent, count As Integer) As String()
'
'    ReDim Results(0 To count - 1) As String
'
'    Dim line As Integer
'    Dim proc As Integer
'    For line = 1 To component.CodeModule.CountOfLines
'        If IsTestMethodLine(component.CodeModule.Lines(line, 1)) Then
'            Results(proc) = component.name & "." & component.CodeModule.ProcOfLine(line, 0)
'            proc = proc + 1
'        End If
'    Next
'
'    GetTestMethods = Results
'End Function

Rem =head4 Function ~
Rem sheetname GetTestMethods
Rem
Rem Rewritten by JHD to avoid the need for a count.
Rem

Friend Function GetTestMethods(component As VBComponent) As String()
On Error GoTo ErrorHandler

Dim sResults() As String
Dim line As Long
Dim proc As Long
proc = -1
For line = 1 To component.CodeModule.CountOfLines
    If IsTestMethodLine(component.CodeModule.Lines(line, 1)) Then
        proc = proc + 1
        ReDim Preserve sResults(0 To proc)
        sResults(proc) = component.name & "." & component.CodeModule.ProcOfLine(line, 0)
    End If
Next

GetTestMethods = sResults

Exit Function
ErrorHandler:
ErrTrap ksErrMod, "GetTestMethods"
End Function

'Friend Function GetTestMethodsCount(component As VBComponent) As Integer
'    Dim count As Integer
'    Dim i As Integer
'    For i = 1 To component.CodeModule.CountOfLines
'        If IsTestMethodLine(component.CodeModule.Lines(i, 1)) Then
'            count = count + 1
'        End If
'    Next
'
'    GetTestMethodsCount = count
'
'End Function

Rem
Rem =head4 Function ~
Rem sheetname IsTestMethodLine
Rem
Rem Originally, this compared the start of a line to a single string. It has been modified by
Rem JHD to strip out "Public" to allow the running of code not explicitly declared as Public
Rem (the default) and to identify functions as well as subs, to allow
Rem comprehensive error trapping.
Rem

Friend Function IsTestMethodLine(ByVal line As String) As Boolean
On Error GoTo ErrorHandler

    line = Replace(line, "Public ", "", 1, 1)   'Strip out leading Public, if it exists
    
    IsTestMethodLine = Left(line, Len(TestMethodPrefix) + 4) Like "Sub " & TestMethodPrefix Or _
                       Left(line, Len(TestMethodPrefix) + 9) Like "Function " & TestMethodPrefix

Exit Function
ErrorHandler:
ErrTrap ksErrMod, "IsTestMethodLine"
IsTestMethodLine = False
End Function
