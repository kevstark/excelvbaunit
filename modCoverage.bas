Attribute VB_Name = "modCoverage"
Option Explicit
Rem order 5.6
Rem
Rem =head2
Rem sheetname
Rem
Rem This standard code module contains code to calculate the coverage of tests in all modules
Rem that have a related tester module. Modules that have no such tester module are ignored,
Rem as the statistics may be useless. Examples from this add-in are the groups of class modules
Rem where three modules are linked, two providing implementations of the third. While it would
Rem be possible to calculate the coverage of the test module against its intended target,
Rem suggesting that coverage of the other two modules is zero would be pointless and misleading.
Rem Similarly, the DummyTestModules should not be tested. They are used for another purpose.
Rem
Rem =head3
Rem sheetname Macros
Rem
Public gMods() As String
Public gAllCalls As Variant
Dim mModYes As Long
Dim mModNo As Long
Dim mModWhy As Long
Dim mTotYes As Long
Dim mTotNo As Long
Dim mTotWhy As Long
Const ksErrMod As String = "modCoverage"

Rem =head4 Function ~
Rem sheetname SetUpArrays
Rem
Rem SetUpArrays() As Boolean
Rem
Rem rcl True
Rem
Rem Populates the gAllCalls array.
Rem
Function SetUpArrays() As Boolean
On Error GoTo ErrorHandler

    If SafeUbound(gAllCalls) = -1 Then
        gAllCalls = Array("Sub", "Public Sub", "Private Sub", "Friend Sub", _
            "Function", "Public Function", "Private Function", "Friend Function")
    End If 'SafeUbound(mAllCalls) = -1

Exit Function
ErrorHandler:
ErrTrap ksErrMod, "SetUpArrays"
SetUpArrays = True
End Function

Rem =head4 Function ~
Rem sheetname InArray
Rem
Rem InArray(ByVal val As Variant, arr As Variant, Optional bTrunc As Boolean = False) As Boolean
Rem
Rem rcl True
Rem
Rem Given a value and an array (a variant), returns True if the value is in the array. An
Rem optional boolean can be passed indicating that strings should be truncated to the length
Rem of the array elements. If False (the default), strings of different lengths will not match.
Rem The boolean will be ignored unless the value is a string.
Rem
Function InArray(ByVal val As Variant, _
                 arr As Variant, _
                 Optional bTrunc As Boolean = False) As Boolean
On Error GoTo ErrorHandler

    If IsEmpty(arr) Then Exit Function
    If Len(Join(arr, "")) = 0 Then Exit Function
    Dim i As Long
    If TypeName(val) = "String" Then
        For i = LBound(arr) To UBound(arr)
            Dim sVal As String
            sVal = UCase(val)
            If bTrunc Then sVal = Left(sVal, Len(arr(i)))
            If UCase(arr(i)) = sVal Then
                InArray = True
                Exit Function
            End If
        Next i
    Else
        For i = LBound(arr) To UBound(arr)
            If arr(i) = val Then
                InArray = True
                Exit Function
            End If
        Next i
    End If

Exit Function
ErrorHandler:
ErrTrap ksErrMod, "InArray"
End Function

Rem =head4 Function ~
Rem sheetname ExtractMods
Rem
Rem ExtractMods(prj As VBProject) As Boolean
Rem
Rem rcl True
Rem
Rem Populates gMods with all current module names.
Rem
Function ExtractMods(prj As VBProject) As Boolean
On Error GoTo ErrorHandler

    Erase gMods
    Dim nMods As Long
    nMods = 0
    Dim mdl As VBComponent
    For Each mdl In prj.VBComponents
        If mdl.Type = vbext_ct_StdModule Then
            nMods = nMods + 1
            ReDim Preserve gMods(1 To nMods)
            gMods(nMods) = mdl.name
        End If 'mdl.Type = vbext_ct_StdModule
    Next mdl

Exit Function
ErrorHandler:
ErrTrap ksErrMod, "ExtractMods"
ExtractMods = True
End Function

Rem =head4 Function ~
Rem sheetname MatchMods
Rem
Rem MatchMods(sTester As String) As String
Rem
Rem rcl True
Rem
Rem Given a string purporting to be the name of a tester module, returns the equivalent
Rem live module. Returns an empty string if the argument is not in the right form or the
Rem live equivalent does not exist.
Rem
Function MatchMods(sTester As String) As String
On Error GoTo ErrorHandler

    Dim sTarget As String
    sTarget = Replace(sTester, ksTestModuleSuffix, "")
    If Len(sTarget) + Len(ksTestModuleSuffix) = Len(sTester) _
        Then If InArray(sTarget, gMods) Then MatchMods = sTarget

Exit Function
ErrorHandler:
ErrTrap ksErrMod, "MatchMods"
End Function

Rem =head4 Function ~
Rem sheetname GetTesters
Rem
Rem GetTesters() As String()
Rem
Rem rcl True
Rem
Rem This accepts no parameters, but expects gMods to contain an array of module names. The
Rem function iterates through this, returning the names of any modules that end with the
Rem test module suffix constant.
Rem
Function GetTesters() As String()
On Error GoTo ErrorHandler

    If Len(Join(gMods, "")) = 0 Then Exit Function
    Dim sTesters() As String
    Dim nRight As Long
    nRight = Len(ksTestModuleSuffix)
    Dim nTesters As Long
    Dim i As Long
    For i = LBound(gMods) To UBound(gMods)
        If Right(gMods(i), nRight) = ksTestModuleSuffix Then
            nTesters = nTesters + 1
            ReDim Preserve sTesters(1 To nTesters)
            sTesters(nTesters) = gMods(i)
        End If 'Right(gMods(i), nRight) = ksTestModuleSuffix
    Next i
    GetTesters = sTesters

Exit Function
ErrorHandler:
ErrTrap ksErrMod, "GetTesters"
End Function

Rem =head4 Function ~
Rem sheetname GetProcs
Rem
Rem GetProcs(prj As VBProject, sMod As String, sProcs() As String) As Boolean
Rem
Rem rcl True
Rem
Rem Given a project and the name of a module, returns an array of strings (the last parameter)
Rem containing all procedures in the array.
Rem
Function GetProcs(prj As VBProject, sMod As String, sProcs() As String) As Boolean
On Error GoTo ErrorHandler

    Erase sProcs
    Dim nProcs As Long
    nProcs = 0
    If SetUpArrays Then err.Raise knCall, , ksCall
    Dim mdl As VBComponent
    Set mdl = prj.VBComponents(sMod)
    With mdl.CodeModule
        Dim nLine As Long
        For nLine = 1 To .CountOfLines
            Dim sLine As String
            sLine = .Lines(nLine, 1)
            If InArray(sLine, gAllCalls, True) Then
                nProcs = nProcs + 1
                ReDim Preserve sProcs(1 To nProcs)
                sProcs(nProcs) = sLine
            End If
        Next nLine
    End With

Exit Function
ErrorHandler:
ErrTrap ksErrMod, "GetProcs"
GetProcs = True
End Function

Rem =head4 Function ~
Rem sheetname ProcNames
Rem
Rem ProcNames(ByRef sProcs() As String) As Boolean
Rem
Rem rcl True
Rem
Rem Takes an array of strings of lines containing procedure names and strips them down to the
Rem bare procedure names.
Rem
Function ProcNames(ByRef sProcs() As String) As Boolean
On Error GoTo ErrorHandler

    Dim nProc As Long
    For nProc = LBound(sProcs) To UBound(sProcs)
        sProcs(nProc) = Left(sProcs(nProc), InStr(sProcs(nProc), "(") - 1)
        Dim nCall As Long
        For nCall = LBound(gAllCalls) To UBound(gAllCalls)
            If Left(sProcs(nProc), Len(gAllCalls(nCall))) = gAllCalls(nCall) Then _
                sProcs(nProc) = Replace(sProcs(nProc), gAllCalls(nCall) & " ", "", , 1)
        Next nCall
    Next nProc

Exit Function
ErrorHandler:
ErrTrap ksErrMod, "ProcNames"
ProcNames = True
End Function

Rem =head4 Function ~
Rem sheetname MapOneName
Rem
Rem MapOneName(ByVal sProc As String, sLive() As String, sTest() As String) As Boolean
Rem
Rem rcl True
Rem
Rem Takes a string representing a procedure name and two arrays of procedure names. If the
Rem string exists in the first array and, preceded by ksTestModulePrefix, in the second array,
Rem it is deleted from both.
Rem
Function MapOneName(ByVal sProc As String, sLive() As String, sTest() As String) As Boolean
On Error GoTo ErrorHandler

    If Not InArray(sProc, sLive) Then Exit Function
    If Not InArray(ksTestModulePrefix & sProc, sTest) Then Exit Function
    Dim i As Long
    For i = LBound(sLive) To UBound(sLive)
        If sLive(i) = sProc Then
            sLive(i) = sLive(UBound(sLive))
            If Pop(sLive) Then err.Raise knCall, , ksCall
            Exit For
        End If
    Next i
    sProc = ksTestModulePrefix & sProc
    For i = LBound(sTest) To UBound(sTest)
        If sTest(i) = sProc Then
            sTest(i) = sTest(UBound(sTest))
            If Pop(sTest) Then err.Raise knCall, , ksCall
            Exit For
        End If
    Next i
    mModYes = mModYes + 1

Exit Function
ErrorHandler:
ErrTrap ksErrMod, "MapOneName"
MapOneName = True
End Function

Rem =head4 Function ~
Rem sheetname StartCoverAll
Rem
Rem StartCoverAll() As Boolean
Rem
Rem rcl True
Rem
Rem Sets grand total variables to zero.
Rem sto ModVars
Rem This uses module level variables. These cannot be seen outside the module, making its
Rem actions untestable. Maintenance programmers planning anything more sophisticated than
Rem is currently here are advised to put the code in another procedure that can be tested.
Rem sto 0
Rem rcl ModVars
Rem
Function StartCoverAll() As Boolean
On Error GoTo ErrorHandler

    mTotYes = 0
    mTotNo = 0
    mTotWhy = 0
    If StartCoverMod Then err.Raise knCall, , ksCall

Exit Function
ErrorHandler:
ErrTrap ksErrMod, "StartCoverAll"
StartCoverAll = True
End Function

Rem =head4 Function ~
Rem sheetname StartCoverMod
Rem
Rem StartCoverMod() As Boolean
Rem
Rem rcl True
Rem
Rem Sets module totals to zero.
Rem rcl ModVars
Rem
Function StartCoverMod() As Boolean
On Error GoTo ErrorHandler

    mModYes = 0
    mModNo = 0
    mModWhy = 0

Exit Function
ErrorHandler:
ErrTrap ksErrMod, "StartCoverMod"
StartCoverMod = True
End Function

Rem =head4 Function ~
Rem sheetname EndCoverMod
Rem
Rem EndCoverMod(sMod As String) As Boolean
Rem
Rem rcl True
Rem
Rem Takes a module name as a string argument. Increments totals and prints module results.
Rem rcl ModVars
Rem
Function EndCoverMod(sMod As String) As Boolean
On Error GoTo ErrorHandler

    mTotYes = mTotYes + mModYes
    mTotNo = mTotNo + mModNo
    mTotWhy = mTotWhy + mModWhy
    Debug.Print "Module """ & sMod & """: " & mModYes & " tested, " & mModNo & " untested, " _
        & mModWhy & " unmatched in " & sMod & ksTestModuleSuffix

Exit Function
ErrorHandler:
ErrTrap ksErrMod, "EndCoverMod"
EndCoverMod = True
End Function

Rem =head4 Function ~
Rem sheetname EndCoverAll
Rem
Rem EndCoverAll() As Boolean
Rem
Rem rcl True
Rem
Rem Prints total results.
Rem rcl ModVars
Rem
Function EndCoverAll() As Boolean
On Error GoTo ErrorHandler

    Debug.Print "Total: " & mTotYes & " tested, " & mTotNo & " untested, " _
        & mTotWhy & " unmatched in all """ & ksTestModuleSuffix & """ modules"

Exit Function
ErrorHandler:
ErrTrap ksErrMod, "EndCoverAll"
EndCoverAll = True
End Function

Rem =head4 Function ~
Rem sheetname StripSetUp
Rem
Rem StripSetUp(sProcs() As String) As Boolean
Rem
Rem rcl True
Rem
Rem Removes not only SetUp but also TearDown and their fixture equivalents (see
Rem L<here|Function TestFixture RunTests>) from an array of strings containing procedure names.
Rem
Function StripSetUp(sProcs() As String) As Boolean
On Error GoTo ErrorHandler

    If SafeUbound(sProcs) = -1 Then Exit Function
    Dim i As Long
    For i = UBound(sProcs) To LBound(sProcs) Step -1
        If sProcs(i) = ksSetUpFunctionName _
        Or sProcs(i) = ksTearDownFunctionName _
        Or sProcs(i) = ksFixtureSetUpFunctionName _
        Or sProcs(i) = ksFixtureTearDownFunctionName Then
            If i < UBound(sProcs) Then
                sProcs(i) = sProcs(UBound(sProcs))
            End If 'i < UBound(sProcs)
            If Pop(sProcs) Then err.Raise knCall, , ksCall
        End If
    Next i

Exit Function
ErrorHandler:
ErrTrap ksErrMod, "StripSetUp"
StripSetUp = True
End Function

Rem =head4 Function ~
Rem sheetname Coverage
Rem doc 1
Rem =head2 Coverage
Rem doc 1 2
Rem
Rem Coverage(Optional sPrj As String, Optional sMod As String) As Boolean
Rem
Rem rcl True
Rem
Rem Takes two optional string parameters. The first is the name of the project, assumed to
Rem be the add-in if empty. The second is the name of the module to be analysed. If no
Rem module is specified, every module in the project that has a "Tester" module will be
Rem analysed. To run this, type C<coverage [project][,module]> in the immediate pane.
Rem
Rem doc 2
Rem This procedure is not tested.
Rem
Rem doc 1 2
Rem The coverage analysis will return three statistics for each pair of modules. These are
Rem "tested", "untested" and "unmatched". The "tested" statistic gives the number of procedures
Rem (always excluding Property* procedures in class modules) that have a matching Test
Rem procedure. The match is independent of the procedure type, so a "Friend Sub"'s matching
Rem test can be a Function. The "untested" statistic is the number of procedures in the live
Rem module that have no direct equivalent in the test module. The "unmatched" statistic is the
Rem number of procedures in the test module that have no equivalent in the live module.
Rem
Rem These statistics should not be used blindly. They are merely a test of procedure names, not
Rem of the quantity or quality of the tests in testing procedures. Simply running RetroFit
Rem will create statistics implying 100% coverage, even though no tests have been written.
Rem Conversely, class modules that implement other class modules will result in statistics
Rem suggesting that nothing is tested. This is because the procedure names in the implementing
Rem class module will have names that point to the implemented class module, while the testing
Rem module will not have the name of the implemented class module in the name of the
Rem procedure. It is conceivable that this will be circumvented in a later version.
Rem
Rem doc 2
Function Coverage(Optional sPrj As String, Optional sMod As String) As Boolean
On Error GoTo ErrorHandler

    Dim nRight As Long
    nRight = Len(ksTestModuleSuffix)
    If StartCoverAll Then err.Raise knCall, , ksCall
    If GetNames Then err.Raise knCall, , ksCall
    If Len(sPrj) = 0 Then sPrj = AddInName
    Dim prj As VBProject
    If DudProj(sPrj, prj) Then err.Raise knCall, , ksCall
    Dim sTest As String
    If Len(sMod) = 0 Then
        Dim mdl1 As VBComponent
        For Each mdl1 In prj.VBComponents
            If Right(mdl1.name, nRight) = ksTestModuleSuffix Then
                sTest = mdl1.name
                sMod = Left(sTest, Len(sTest) - nRight)
                Dim mdl2 As VBComponent
                For Each mdl2 In prj.VBComponents
                    If mdl2.name = sMod Then
                        If CoverSingle(prj, sMod) Then err.Raise knCall, , ksCall
                    End If
                Next mdl2
            End If 'Right(mdl1.name, nRight) = ksTestModuleSuffix
        Next mdl1
    Else
        If Right(sMod, nRight) = ksTestModuleSuffix Then
            sTest = sMod
            sMod = Left(sMod, Len(sMod) - nRight)
        Else
            sTest = sMod & ksTestModuleSuffix
        End If 'Right(sMod, Len(ksTestModuleSuffix)) = ksTestModuleSuffix
        If CoverSingle(prj, sMod) Then err.Raise knCall, , ksCall
    End If 'Len(sMod) = 0
    If EndCoverAll Then err.Raise knCall, , ksCall

Exit Function
ErrorHandler:
ErrTrap ksErrMod, "Coverage"
Coverage = True
End Function

Rem =head4 Function ~
Rem sheetname MapNames
Rem
Rem MapNames(sLive() As String, sTest() As String) As Boolean
Rem
Rem rcl True
Rem
Rem Takes two arrays of strings and strips out those that have test equivalents.
Rem
Function MapNames(sLive() As String, sTest() As String) As Boolean
On Error GoTo ErrorHandler

    Dim i As Long
    For i = UBound(sTest) To LBound(sTest) Step -1
        Dim sProc As String
        sProc = Right(sTest(i), Len(sTest(i)) - Len(ksTestModulePrefix))
        If MapOneName(sProc, sLive, sTest) Then err.Raise knCall, , ksCall
    Next i

Exit Function
ErrorHandler:
ErrTrap ksErrMod, "MapNames"
MapNames = True
End Function

Rem =head4 Function ~
Rem sheetname Pop
Rem
Rem Pop(ary As Variant) As Boolean
Rem
Rem rcl True
Rem
Rem Takes an array and strips the last element. Note that this works only for single dimension
Rem arrays.
Rem
Function Pop(ary As Variant) As Boolean
On Error GoTo ErrorHandler

    If VarType(ary) < vbArray Then err.Raise knCall - 1, , "Trying to pass a scalar to Pop."
    If UBound(ary) = LBound(ary) Then
        Erase ary
    Else
        ReDim Preserve ary(LBound(ary) To UBound(ary) - 1)
    End If 'UBound(ary) = LBound(ary)

Exit Function
ErrorHandler:
ErrTrap ksErrMod, "Pop"
Pop = True
End Function

Rem =head4 Function ~
Rem sheetname CoverSingle
Rem
Rem CoverSingle(prj As VBProject, sMod As String) As Boolean
Rem
Rem rcl True
Rem
Rem Calculates coverage statistics for a single module.
Rem
Function CoverSingle(prj As VBProject, sMod As String) As Boolean
On Error GoTo ErrorHandler

    If StartCoverMod Then err.Raise knCall, , ksCall
    Dim sTest As String
    sTest = sMod & ksTestModuleSuffix
    Dim sLiveProcs() As String
    If GetProcs(prj, sMod, sLiveProcs) Then err.Raise knCall, , ksCall
    Dim sTestProcs() As String
    If GetProcs(prj, sTest, sTestProcs) Then err.Raise knCall, , ksCall
    If ProcNames(sLiveProcs) Then err.Raise knCall, , ksCall
    If ProcNames(sTestProcs) Then err.Raise knCall, , ksCall
    If StripSetUp(sTestProcs) Then err.Raise knCall, , ksCall
    If MapNames(sLiveProcs, sTestProcs) Then err.Raise knCall, , ksCall
    If SafeUbound(sTestProcs) > -1 Then mModWhy = UBound(sTestProcs)
    If SafeUbound(sLiveProcs) > -1 Then mModNo = UBound(sLiveProcs)
    If EndCoverMod(sMod) Then err.Raise knCall, , ksCall

Exit Function
ErrorHandler:
ErrTrap ksErrMod, "CoverSingle"
CoverSingle = True
End Function
