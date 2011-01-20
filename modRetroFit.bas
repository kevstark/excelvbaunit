Attribute VB_Name = "modRetroFit"
Option Explicit
Rem order 5.5
Rem
Rem =head2
Rem sheetname
Rem
Rem This code module contains procedures to add a fixture (code module) to match an existing
Rem code module and to populate the fixture with stub functions to test every existing
Rem sub or function in the given module. This module's main function cannot
Rem be tested, but it should be possible to test many of the routines it calls.
Rem
Rem =head3
Rem sheetname Macros
Rem
Const ksErrMod As String = "modRetroFit"

Rem =head4 Function ~
Rem sheetname RetroFit
Rem doc 1
Rem =head2 RetroFit
Rem doc 1 2
Rem
Rem RetroFit(sMod As String, Optional bFn As Boolean = True,
Rem Optional bAddErrTrap As Boolean = True, Optional bPOD As Boolean = True) As Boolean
Rem
Rem rcl True
Rem
Rem This function is called from the Immediate pane. Given a string parameter of the name of a
Rem module, it will create a new module to test it. The name should be in the fully qualified
Rem form of project.module. If the project is omitted, the add-in itself will be assumed.
Rem There are three optional boolean parameters. The first indicates whether functions are
Rem to be used. This defaults to True, but some users may prefer subs. The second indicates
Rem whether an error trap is to be added. Again, this defaults to True. The last indicates
Rem whether POD is to be added. Obviously, this is the framework only, but if it is True
Rem (again, this is the default), the framework will be added both to the top of the new
Rem module and to the top of each generated procedure stub.
Rem
Rem doc 1
Rem If you are reading this without any idea of what POD is, it is the comments in the code
Rem that are processed into the document you are currently reading. It makes documentation
Rem less work and is a Good Thing(tm). It has been snarfed shamelessly from Perl (indeed, the
Rem ExcelPOD code that turns comments into documentation is written in Perl) and stands for
Rem Plain Old Documentation.
Rem
Rem doc 2
Function RetroFit(sMod As String, _
                  Optional bFn As Boolean = True, _
                  Optional bAddErrTrap As Boolean = True, _
                  Optional bPOD As Boolean = True) As Boolean
On Error GoTo ErrorHandler

    If CorrectMod(sMod) Then err.Raise knCall, , ksCall
    Dim prj As VBProject
    If Not ProjFound(sMod, prj) Then
        MsgBox "Sorry! Can't find your project " & sMod, vbCritical, "Aborting"
        Exit Function
    End If 'Not ProjFound
    Dim mdl As VBComponent
    If ModFound(sMod & ksTestModuleSuffix, prj, mdl) Then
        MsgBox "Sorry! Can't duplicate module " & sMod & ksTestModuleSuffix, _
            vbCritical, "Aborting"
        Exit Function
    End If 'ModFound
    If Not ModFound(sMod, prj, mdl) Then
        MsgBox "Sorry! Can't find your module " & sMod, vbCritical, "Aborting"
        Exit Function
    End If 'not ModFound
    Dim mdlTester As VBComponent
    Set mdlTester = prj.VBComponents.Add(vbext_ct_StdModule)
    mdlTester.name = mdl.name & ksTestModuleSuffix
    If bPOD Then If AddPODHead(mdlTester) Then err.Raise knCall, , ksCall
    If AddStubs(mdl, mdlTester, bFn, bAddErrTrap, bPOD) Then err.Raise knCall, , ksCall

Exit Function
ErrorHandler:
ErrTrap ksErrMod, "RetroFit"
RetroFit = True
End Function

Rem =head4 Function ~
Rem sheetname CorrectMod
Rem
Rem CorrectMod(ByRef sMod As String) As Boolean
Rem
Rem rcl True
Rem
Rem Given a module name, tests it for a project name. If no project name is found, prepends
Rem the name of the add-in.
Rem
Function CorrectMod(ByRef sMod As String) As Boolean
On Error GoTo ErrorHandler

    If InStr(sMod, ".") < 2 Then
        If GetNames Then err.Raise knCall, , ksCall
        sMod = AddInName & "." & Replace(sMod, ".", "", , 1) 'Strips out any leading .
    End If 'InStr(sMod, ".") = 0

Exit Function
ErrorHandler:
ErrTrap ksErrMod, "CorrectMod"
CorrectMod = True
End Function

Rem =head4 Function ~
Rem sheetname ProjFound
Rem
Rem ProjFound(sMod As String, prj As VBProject) As Boolean
Rem
Rem Given a qualified module name, tests whether the project (the qualifier) can be found.
Rem True indicates the project has been found. The prj variable will be set to the project.
Rem
Function ProjFound(sMod As String, prj As VBProject) As Boolean
On Error GoTo ErrorHandler

    Dim sPrj As String
    sPrj = UCase(Left(sMod, InStr(sMod, ".") - 1))
    For Each prj In Application.VBE.VBProjects
        If UCase(prj.name) = sPrj Then
            ProjFound = True
            Exit Function
        End If 'UCase(prj.name) = sPrj
    Next prj
    Set prj = Nothing
    
Exit Function
ErrorHandler:
ErrTrap ksErrMod, "ProjFound"
End Function

Rem =head4 Function ~
Rem sheetname ModFound
Rem
Rem ModFound(ByVal sMod As String, prj As VBProject, mdl As VBComponent) As Boolean
Rem
Rem Returns True if the module is found in the project. Also sets the third parameter to the
Rem module.
Rem
Function ModFound(ByVal sMod As String, prj As VBProject, mdl As VBComponent) As Boolean
On Error GoTo ErrorHandler

    sMod = UCase(sMod)
    If InStr(sMod, ".") Then sMod = Right(sMod, Len(sMod) - InStrRev(sMod, "."))
    Dim cmp As VBComponent
    For Each cmp In prj.VBComponents
        If sMod = UCase(cmp.name) Then
            Set mdl = cmp
            ModFound = True
            Exit Function
        End If 'UCase(cmp.name) = sMod
    Next cmp

Exit Function
ErrorHandler:
ErrTrap ksErrMod, "ModFound"
End Function

Rem =head4 Function ~
Rem sheetname AddStubs
Rem
Rem AddStubs(mdl As VBComponent, mdlTester As VBComponent, Optional bFn As Boolean = True,
Rem Optional bAddErrTrap As Boolean = True, Optional bPOD As Boolean = True) As Boolean
Rem
Rem rcl True
Rem
Rem Adds the stubs to the Tester module based on the source module.
Rem

Function AddStubs(mdl As VBComponent, _
                  mdlTester As VBComponent, _
                  Optional bFn As Boolean = True, _
                  Optional bAddErrTrap As Boolean = True, _
                  Optional bPOD As Boolean = True) As Boolean
On Error GoTo ErrorHandler

    Dim sCall As Variant
    sCall = Array("Sub", "Public Sub", "Private Sub", "Friend Sub", _
                  "Function", "Public Function", "Private Function", "Friend Function")
    Dim sProc As String
    sProc = IIf(bFn, "Function", "Sub")
    Dim nLine As Long
    For nLine = 1 To mdl.CodeModule.CountOfLines
        Dim sLine As String
        sLine = mdl.CodeModule.Lines(nLine, 1)
        Dim nCall As Long
        For nCall = LBound(sCall) To UBound(sCall)
            If Left(sLine, Len(sCall(nCall))) = sCall(nCall) Then
                Dim sCode As String
                sCode = sProc & " "
                Dim sName As String
                sName = Replace(Replace(sLine, sCall(nCall), ""), " ", "", 1, 1)
                sName = ksTestModulePrefix & Left(sName, InStr(sName, "("))
                sCode = sCode & sName & ")"
                sName = Replace(sName, "(", "")
                If bPOD Then If AddPODProc(mdlTester, sProc, sName) _
                    Then err.Raise knCall, , ksCall
                With mdlTester.CodeModule
                    .InsertLines .CountOfLines + 1, sCode
                    If bAddErrTrap Then
                        If InsTrap(mdlTester, sProc, sName) Then err.Raise knCall, , ksCall
                    Else
                        .InsertLines .CountOfLines + 1, ""
                    End If 'bAddErrTrap
                    sCode = "End "
                    sCode = sCode & sProc
                    .InsertLines .CountOfLines + 1, sCode
                End With
                Exit For
            End If 'Left(sLine, Len(sCall(nCall))) = sCall(nCall)
        Next nCall
    Next nLine

Exit Function
ErrorHandler:
ErrTrap ksErrMod, "AddStubs"
AddStubs = True
End Function

Rem =head4 Function ~
Rem sheetname AddPODHead
Rem
Rem AddPODHead(mdl As VBComponent) As Boolean
Rem
Rem rcl True
Rem
Rem Adds POD to the head of a module.
Rem
Function AddPODHead(mdl As VBComponent) As Boolean
On Error GoTo ErrorHandler

    With mdl.CodeModule
        .InsertLines .CountOfLines + 1, "Rem order"
        .InsertLines .CountOfLines + 1, "Rem"
        .InsertLines .CountOfLines + 1, "Rem =head2"
        .InsertLines .CountOfLines + 1, "Rem sheetname"
        .InsertLines .CountOfLines + 1, "Rem"
        .InsertLines .CountOfLines + 1, "Rem"
        .InsertLines .CountOfLines + 1, "Rem"
        .InsertLines .CountOfLines + 1, "Rem =head3"
        .InsertLines .CountOfLines + 1, "Rem sheetname Macros"
        .InsertLines .CountOfLines + 1, "Rem"
    End With

Exit Function
ErrorHandler:
ErrTrap ksErrMod, "AddPODHead"
AddPODHead = True
End Function

Rem =head4 Function ~
Rem sheetname AddPODProc
Rem
Rem AddPODProc(mdl As VBComponent, sProc As String, sName As String) As Boolean
Rem
Rem rcl True
Rem
Rem Adds POD to a procedure.
Rem
Function AddPODProc(mdl As VBComponent, sProc As String, sName As String) As Boolean
On Error GoTo ErrorHandler

    With mdl.CodeModule
        .InsertLines .CountOfLines + 1, "Rem =head4 " & sProc & " " & Replace(sName, "(", "")
        .InsertLines .CountOfLines + 1, "Rem"
        .InsertLines .CountOfLines + 1, "Rem"
        .InsertLines .CountOfLines + 1, "Rem"
    End With

Exit Function
ErrorHandler:
ErrTrap ksErrMod, "AddPODProc"
AddPODProc = True
End Function

Rem =head4 Function ~
Rem sheetname InsTrap
Rem
Rem InsTrap(mdl As VBComponent, sProc As String, sName As String) As Boolean
Rem
Rem rcl True
Rem
Rem Adds an error trap to a sub or function.
Rem
Function InsTrap(mdl As VBComponent, sProc As String, sName As String) As Boolean
On Error GoTo ErrorHandler

    With mdl.CodeModule
        .InsertLines .CountOfLines + 1, "On Error GoTo ErrorHandler"
        .InsertLines .CountOfLines + 1, ""
        .InsertLines .CountOfLines + 1, "Exit " & sProc
        .InsertLines .CountOfLines + 1, "ErrorHandler:"
        .InsertLines .CountOfLines + 1, "ErrTrap ksErrMod, """ & sName & """"
        .InsertLines .CountOfLines + 1, sName & " = True"
    End With

Exit Function
ErrorHandler:
ErrTrap ksErrMod, "InsTrap"
InsTrap = True
End Function
