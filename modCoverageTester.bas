Attribute VB_Name = "modCoverageTester"
Option Explicit

Rem order 5.65
Rem
Rem =head2
Rem sheetname
Rem
Rem
Rem
Rem =head3
Rem sheetname Macros
Rem
Const ksErrMod As String = "modCoverageTester"

Rem =head4 Function ~
Rem sheetname TestSetUpArrays
Rem
Rem 8 tests that arrays are properly populated.
Rem
Function TestSetUpArrays()
On Error GoTo ErrorHandler

    If SetUpArrays Then err.Raise knCall, , ksCall
    If AssertEqual(gAllCalls(0), "Sub", "Can't set up modCoverage.gAllCalls") _
        Then err.Raise knCall, , ksCall
    If AssertEqual(gAllCalls(1), "Public Sub", "Can't set up modCoverage.gAllCalls") _
        Then err.Raise knCall, , ksCall
    If AssertEqual(gAllCalls(2), "Private Sub", "Can't set up modCoverage.gAllCalls") _
        Then err.Raise knCall, , ksCall
    If AssertEqual(gAllCalls(3), "Friend Sub", "Can't set up modCoverage.gAllCalls") _
        Then err.Raise knCall, , ksCall
    If AssertEqual(gAllCalls(4), "Function", "Can't set up modCoverage.gAllCalls") _
        Then err.Raise knCall, , ksCall
    If AssertEqual(gAllCalls(5), "Public Function", "Can't set up modCoverage.gAllCalls") _
        Then err.Raise knCall, , ksCall
    If AssertEqual(gAllCalls(6), "Private Function", "Can't set up modCoverage.gAllCalls") _
        Then err.Raise knCall, , ksCall
    If AssertEqual(gAllCalls(7), "Friend Function", "Can't set up modCoverage.gAllCalls") _
        Then err.Raise knCall, , ksCall

Exit Function
ErrorHandler:
ErrTrap ksErrMod, "TestSetUpArrays"
TestSetUpArrays = True
End Function

Rem =head4 Function ~
Rem sheetname TestInArray
Rem
Rem 3 tests, 1 that arrays are properly populated, 1 that InArray returns false on a mismatch,
Rem 1 that InArray can handle an empty array.
Rem
Function TestInArray() As Boolean
On Error GoTo ErrorHandler

    If SetUpArrays Then err.Raise knCall, , ksCall
    If AssertTrue(InArray("Friend Sub", gAllCalls)) Then err.Raise knCall, , ksCall
    If AssertFalse(InArray("qwerty", gAllCalls)) Then err.Raise knCall, , ksCall
    Dim ary() As String
    If AssertFalse(InArray("", ary)) Then err.Raise knCall, , ksCall

Exit Function
ErrorHandler:
ErrTrap ksErrMod, "TestInArray"
TestInArray = True
End Function

Rem =head4 Function ~
Rem sheetname TestExtractMods
Rem
Rem 3 tests, clearing and populating.
Rem

Function TestExtractMods() As Boolean
On Error GoTo ErrorHandler

    Erase gMods
    If AssertEqual(-1, SafeUbound(gMods), "Can't clear gMods") Then err.Raise knCall, , ksCall
    If GetNames Then err.Raise knCall, , ksCall
    Dim prj As VBProject
    Set prj = Application.VBE.VBProjects(AddInName)
    If ExtractMods(prj) Then err.Raise knCall, , ksCall
    If AssertTrue(SafeUbound(gMods) > 0, "No modules extracted") Then err.Raise knCall, , ksCall
    If AssertTrue(InArray(ksErrMod, gMods)) Then err.Raise knCall, , ksCall

Exit Function
ErrorHandler:
ErrTrap ksErrMod, "TestExtractMods"
TestExtractMods = True
End Function

Rem =head4 Function ~
Rem sheetname TestMatchMods
Rem
Rem 3 tests, 1 for this module and its live equivalent, 1 for an imaginary module and 1 for a
Rem non-tester module.
Rem
Function TestMatchMods() As Boolean
On Error GoTo ErrorHandler

    If GetNames Then err.Raise knCall, , ksCall
    Dim prj As VBProject
    Set prj = Application.VBE.VBProjects(AddInName)
    If ExtractMods(prj) Then err.Raise knCall, , ksCall
    If AssertEqual(Replace(ksErrMod, ksTestModuleSuffix, ""), MatchMods(ksErrMod)) _
        Then err.Raise knCall, , ksCall
    If AssertFalse(MatchMods("xTester") = "x") Then err.Raise knCall, , ksCall
    If AssertEqual(MatchMods(Replace(ksErrMod, ksTestModuleSuffix, "")), "") _
        Then err.Raise knCall, , ksCall

Exit Function
ErrorHandler:
ErrTrap ksErrMod, "TestMatchMods"
TestMatchMods = True
End Function

Rem =head4 Function ~
Rem sheetname TestGetTesters
Rem
Rem 3 tests for presence of this module and absence of other modules.
Rem
Function TestGetTesters() As Boolean
On Error GoTo ErrorHandler

    If GetNames Then err.Raise knCall, , ksCall
    Dim prj As VBProject
    Set prj = Application.VBE.VBProjects(AddInName)
    If ExtractMods(prj) Then err.Raise knCall, , ksCall
    Dim sTesters() As String
    sTesters = GetTesters
    If AssertTrue(InArray(ksErrMod, sTesters)) Then err.Raise knCall, , ksCall
    If AssertFalse(InArray(Replace(ksErrMod, ksTestModuleSuffix, ""), sTesters)) _
        Then err.Raise knCall, , ksCall
    If AssertFalse(InArray("xTester", sTesters)) Then err.Raise knCall, , ksCall

Exit Function
ErrorHandler:
ErrTrap ksErrMod, "TestGetTesters"
TestGetTesters = True
End Function

Rem =head4 Function ~
Rem sheetname TestGetProcs
Rem
Rem 2 tests:
Rem
Rem =over
Rem
Rem =item 1) This module has some procedures.
Rem
Rem =item 2) DummyTestModule has 7.
Rem
Rem =back
Rem
Function TestGetProcs() As Boolean
On Error GoTo ErrorHandler

    Dim prj As VBProject
    Set prj = Application.VBE.VBProjects(AddInName)
    Dim sProcs() As String
    If GetProcs(prj, ksErrMod, sProcs) Then err.Raise knCall, , ksCall
    If AssertTrue(SafeUbound(sProcs) > -1) Then err.Raise knCall, , ksCall
    
    If GetProcs(prj, "DummyTestModule", sProcs) Then err.Raise knCall, , ksCall
    If AssertEqual(7, SafeUbound(sProcs)) Then err.Raise knCall, , ksCall

Exit Function
ErrorHandler:
ErrTrap ksErrMod, "TestGetProcs"
TestGetProcs = True
End Function

Rem =head4 Function ~
Rem sheetname TestProcNames
Rem
Rem Unknown number of tests. Depends on code path, number of procedures etc.
Rem
Function TestProcNames() As Boolean
On Error GoTo ErrorHandler

    Dim prj As VBProject
    Set prj = Application.VBE.VBProjects(AddInName)
    Dim sProcs() As String
    If GetProcs(prj, "DummyTestModule", sProcs) Then err.Raise knCall, , ksCall
    If ProcNames(sProcs) Then err.Raise knCall, , ksCall
    If SafeUbound(sProcs) > -1 Then
        Dim i As Long
        For i = LBound(sProcs) To UBound(sProcs)
            If AssertEqual(0, InStr(sProcs(i), " ")) Then err.Raise knCall, , ksCall
            If AssertEqual(0, InStr(sProcs(i), ")")) Then err.Raise knCall, , ksCall
        Next i
    Else
        If AssertFailure("sProcs not populated") Then err.Raise knCall, , ksCall
    End If 'SafeUbound(sProcs) > -1
    
Exit Function
ErrorHandler:
ErrTrap ksErrMod, "TestProcNames"
TestProcNames = True
End Function

Rem =head4 Function ~
Rem sheetname TestMapNames
Rem
Rem 1 test that this module is totally matched.
Rem
Function TestMapNames() As Boolean
On Error GoTo ErrorHandler

    Dim prj As VBProject
    Set prj = Application.VBE.VBProjects(AddInName)
    Dim sLiveProcs() As String
    If GetProcs(prj, Replace(ksErrMod, ksTestModuleSuffix, ""), sLiveProcs) _
        Then err.Raise knCall, , ksCall
    If ProcNames(sLiveProcs) Then err.Raise knCall, , ksCall
    Dim sTestProcs() As String
    If GetProcs(prj, ksErrMod, sTestProcs) Then err.Raise knCall, , ksCall
    If ProcNames(sTestProcs) Then err.Raise knCall, , ksCall
    If MapNames(sLiveProcs, sTestProcs) Then err.Raise knCall, , ksCall
    If AssertEqual(-1, SafeUbound(sTestProcs)) Then err.Raise knCall, , ksCall

Exit Function
ErrorHandler:
ErrTrap ksErrMod, "TestMapNames"
TestMapNames = True
End Function

Rem =head4 Function ~
Rem sheetname TestMapOneName
Rem
Rem 2 tests, one on each array included in a match.
Rem
Function TestMapOneName() As Boolean
On Error GoTo ErrorHandler

    Dim prj As VBProject
    Set prj = Application.VBE.VBProjects(AddInName)
    Dim sLiveProcs() As String
    If GetProcs(prj, Replace(ksErrMod, ksTestModuleSuffix, ""), sLiveProcs) _
        Then err.Raise knCall, , ksCall
    Dim sTestProcs() As String
    If GetProcs(prj, ksErrMod, sTestProcs) Then err.Raise knCall, , ksCall
    If ProcNames(sLiveProcs) Then err.Raise knCall, , ksCall
    If ProcNames(sTestProcs) Then err.Raise knCall, , ksCall
    Dim nLive As Long
    nLive = UBound(sLiveProcs)
    Dim nTest As Long
    nTest = UBound(sTestProcs)
    If MapOneName("MapOneName", sLiveProcs, sTestProcs) Then err.Raise knCall, , ksCall
    If AssertEqual(nLive - 1, UBound(sLiveProcs)) Then err.Raise knCall, , ksCall
    If AssertEqual(nTest - 1, UBound(sTestProcs)) Then err.Raise knCall, , ksCall

Exit Function
ErrorHandler:
ErrTrap ksErrMod, "TestMapOneName"
TestMapOneName = True
End Function

Rem =head4 Function ~
Rem sheetname TestStripSetUp
Rem
Rem 3 tests. 1 test ensures something is deleted by StripSetUp. 1 ensures that StripSetUp can
Rem accept an empty array. 1 ensures that StripSetUp can cope with emptying an array.
Rem
Function TestStripSetUp() As Boolean
On Error GoTo ErrorHandler

    Dim prj As VBProject
    Set prj = Application.VBE.VBProjects(AddInName)
    Dim sProcs() As String
    If GetProcs(prj, "DummyTestModule", sProcs) Then err.Raise knCall, , ksCall
    If ProcNames(sProcs) Then err.Raise knCall, , ksCall
    Dim nBefore As Long
    nBefore = SafeUbound(sProcs)
    If StripSetUp(sProcs) Then err.Raise knCall, , ksCall
    If AssertTrue(SafeUbound(sProcs) < nBefore) Then err.Raise knCall, , ksCall
    Erase sProcs
    If StripSetUp(sProcs) Then err.Raise knCall, , ksCall
    If AssertEqual(-1, SafeUbound(sProcs)) Then err.Raise knCall, , ksCall
    ReDim sProcs(1 To 1)
    sProcs(1) = ksSetUpFunctionName
    If StripSetUp(sProcs) Then err.Raise knCall, , ksCall
    If AssertEqual(-1, SafeUbound(sProcs)) Then err.Raise knCall, , ksCall
    
Exit Function
ErrorHandler:
ErrTrap ksErrMod, "TestStripSetUp"
TestStripSetUp = True
End Function

Rem =head4 Function ~
Rem sheetname TestPop
Rem
Rem 2 tests, one to make sure that an array is reduced, one to make sure a zero length array
Rem can be handled. Note that this works only with single dimension arrays.
Rem

Function TestPop() As Boolean
On Error GoTo ErrorHandler

    Dim ary() As String
    ReDim ary(1 To 2)
    If Pop(ary) Then err.Raise knCall, , ksCall
    If AssertEqual(1, UBound(ary)) Then err.Raise knCall, , ksCall
    If Pop(ary()) Then err.Raise knCall, , ksCall
    If AssertEqual(-1, SafeUbound(ary)) Then err.Raise knCall, , ksCall

Exit Function
ErrorHandler:
ErrTrap ksErrMod, "TestPop"
TestPop = True
End Function
