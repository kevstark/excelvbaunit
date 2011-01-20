Attribute VB_Name = "modRetroFitTester"
Option Explicit

Rem order 5.51
Rem
Rem =head2
Rem sheetname
Rem
Rem 7 out of 8 routines tested.
Rem
Rem =head3
Rem sheetname Macros
Rem
Dim mdlTester As VBComponent
Dim nLines As Long
Const ksProc As String = "sub"
Const ksName As String = "TestName"
Const ksErrMod As String = "modRetroFitTester"

Rem =head4 Function TestCorrectMod
Rem
Rem 3 tests.
Rem
Function TestCorrectMod()
On Error GoTo ErrorHandler

    Dim sMod As String
    sMod = "a"
    If CorrectMod(sMod) Then err.Raise knCall, , ksCall
    If AssertEqual(AddInName & ".a", sMod) Then err.Raise knCall, , ksCall
    sMod = ".a"
    If CorrectMod(sMod) Then err.Raise knCall, , ksCall
    If AssertEqual(AddInName & ".a", sMod) Then err.Raise knCall, , ksCall
    sMod = "a.a"
    If CorrectMod(sMod) Then err.Raise knCall, , ksCall
    If AssertEqual("a.a", sMod) Then err.Raise knCall, , ksCall

Exit Function
ErrorHandler:
ErrTrap ksErrMod, "TestCorrectMod"
TestCorrectMod = True
End Function

Rem =head4 Function TestProjFound
Rem
Rem 3 tests.
Rem
Function TestProjFound()
On Error GoTo ErrorHandler

    Dim prj As VBProject
    Set prj = Application.VBE.VBProjects(AddInName)
    If AssertTrue(ProjFound(AddInName & ".modRetroFit", prj)) Then err.Raise knCall, , ksCall
    If AssertFalse(ProjFound("a.modRetroFit", prj)) Then err.Raise knCall, , ksCall
    If AssertTrue(ProjFound(AddInName & ".a", prj)) Then err.Raise knCall, , ksCall

Exit Function
ErrorHandler:
ErrTrap ksErrMod, "TestProjFound"
TestProjFound = True
End Function

Rem =head4 Function TestModFound
Rem
Rem 2 tests.
Rem
Function TestModFound()
On Error GoTo ErrorHandler

    Dim mdl As VBComponent
    Dim prj As VBProject
    Set prj = Application.VBE.VBProjects(AddInName)
    If AssertTrue(ModFound(ksErrMod, prj, mdl)) Then err.Raise knCall, , ksCall
    If AssertFalse(ModFound("asdf", prj, mdl)) Then err.Raise knCall, , ksCall

Exit Function
ErrorHandler:
ErrTrap ksErrMod, "TestModFound"
TestModFound = True
End Function

Rem =head4 Function TestAddStubs
Rem
Rem 1 test.
Rem
Function TestAddStubs()
On Error GoTo ErrorHandler

    nLines = mdlTester.CodeModule.CountOfLines
    Dim mdl As VBComponent
    Set mdl = Application.VBE.VBProjects(AddInName).VBComponents("DummyTestModule")
    If AddStubs(mdl, mdlTester, False, False, False) Then err.Raise knCall, , ksCall
    If AssertEqual(21 + nLines, mdlTester.CodeModule.CountOfLines) _
        Then err.Raise knCall, , ksCall

Exit Function
ErrorHandler:
ErrTrap ksErrMod, "TestAddStubs"
TestAddStubs = True
End Function

Rem =head4 Function TestAddPODHead
Rem
Rem 1 test.
Rem
Function TestAddPODHead()
On Error GoTo ErrorHandler

    nLines = mdlTester.CodeModule.CountOfLines
    If AddPODHead(mdlTester) Then err.Raise knCall, , ksCall
    If AssertEqual(10 + nLines, mdlTester.CodeModule.CountOfLines) _
        Then err.Raise knCall, , ksCall

Exit Function
ErrorHandler:
ErrTrap ksErrMod, "TestAddPODHead"
TestAddPODHead = True
End Function

Rem =head4 Function TestAddPODProc
Rem
Rem 1 test.
Rem
Function TestAddPODProc()
On Error GoTo ErrorHandler

    nLines = mdlTester.CodeModule.CountOfLines
    If AddPODProc(mdlTester, "TestProc", "TestName") Then err.Raise knCall, , ksCall
    If AssertEqual(4 + nLines, mdlTester.CodeModule.CountOfLines) _
        Then err.Raise knCall, , ksCall

Exit Function
ErrorHandler:
ErrTrap ksErrMod, "TestAddPODProc"
TestAddPODProc = True
End Function

Rem =head4 Function TestInsTrap
Rem
Rem 1 test.
Rem
Function TestInsTrap()
On Error GoTo ErrorHandler

    nLines = mdlTester.CodeModule.CountOfLines
    mdlTester.CodeModule.InsertLines nLines + 1, ksProc & " " & ksName
    nLines = mdlTester.CodeModule.CountOfLines
    If InsTrap(mdlTester, ksProc, ksName) Then err.Raise knCall, , ksCall
    If AssertEqual(6 + nLines, mdlTester.CodeModule.CountOfLines) _
        Then err.Raise knCall, , ksCall

Exit Function
ErrorHandler:
ErrTrap ksErrMod, "TestInsTrap"
TestInsTrap = True
End Function

Rem =head4 Function SetUp
Rem
Rem Creates a new module to which data will be added.
Rem
Function SetUp() As Boolean
On Error GoTo ErrorHandler

    If GetNames Then err.Raise knCall, , ksCall
    Set mdlTester = Application.VBE.VBProjects(AddInName).VBComponents.Add(vbext_ct_StdModule)
    mdlTester.name = "ToDelete"

Exit Function
ErrorHandler:
ErrTrap ksErrMod, "SetUp"
SetUp = True
End Function

Rem =head4 Function TearDown
Rem
Rem Deletes the module created by SetUp.
Rem
Function TearDown() As Boolean
On Error GoTo ErrorHandler

    Application.VBE.VBProjects(AddInName).VBComponents.Remove mdlTester

Exit Function
ErrorHandler:
ErrTrap ksErrMod, "TearDown"
TearDown = True
End Function
