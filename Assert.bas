Attribute VB_Name = "Assert"
Option Explicit

Rem order 4
Rem doc 1
Rem =head2
Rem Available "Assert" commands
Rem
Rem AssertTrue, AssertFalse, AssertEqual and AssertFailure have an optional
Rem string parameter. This will appear in the log if the test fails, which
Rem in the case of AssertFailure is always.
Rem
Rem doc 2
Rem =head2
Rem sheetname
Rem
Rem This code module contains subs that determine whether tests have passed or failed. It also
Rem contains a test result manager object and a sub to create an instance of it.
Rem
Rem =head3
Rem sheetname Macros
Rem

Const ksErrMod As String = "Assert"
Private mTestResultManager As ITestResultsManager
Private trmBackup As ITestResultsManager

Rem =head4 Function ~
Rem sheetname SetTestResultsManager
Rem
Rem SetTestResultsManager(manager As ITestResultsManager) As Boolean
Rem
Rem sto True
Rem Returns True if an error occurs.
Rem sto 0
Rem rcl True
Rem
Rem Creates an instance of the manager using a module level variable.
Rem

Public Function SetTestResultsManager(manager As ITestResultsManager) As Boolean
On Error GoTo ErrorHandler

    Set mTestResultManager = manager

Exit Function
ErrorHandler:
ErrTrap ksErrMod, "SetTestResultsManager"
SetTestResultsManager = True
End Function
'Public Sub SetTestResultsManager(manager As ITestResultsManager)
'
'    Set mTestResultManager = manager
'
'End Sub

Rem =head4 Function ~
Rem sheetname AssertTrue
Rem doc 1
Rem =head3 AssertTrue
Rem doc 1 2
Rem
Rem AssertTrue(test As Boolean, Optional msg As String = "") As Boolean
Rem
Rem rcl True
Rem
Rem Accepts a boolean which will be true if the test has passed. Logs the result via the
Rem test result manager.
Rem
Rem doc 2
Rem sto ToFn
Rem This needs to be converted to a function with error trapping.
Rem sto 0
Rem
Function AssertTrue(test As Boolean, Optional msg As String = "") As Boolean
On Error GoTo ErrorHandler

    If test Then
        If mTestResultManager.LogSuccess Then err.Raise knCall, , ksCall
    Else
        If mTestResultManager.LogFailure(msg) Then err.Raise knCall, , ksCall
    End If

Exit Function
ErrorHandler:
ErrTrap ksErrMod, "AssertTrue"
AssertTrue = True
End Function
'Public Sub AssertTrue(test As Boolean, Optional msg As String = "")
'
'    If test Then
'        mTestResultManager.LogSuccess
'    Else
'        mTestResultManager.LogFailure msg
'    End If
'
'End Sub

Rem =head4 Function ~
Rem sheetname AssertFalse
Rem doc 1
Rem =head3 AssertFalse
Rem doc 1 2
Rem
Rem AssertFalse(test As Boolean, Optional msg As String = "") As Boolean
Rem
Rem rcl True
Rem
Rem Accepts a boolean which will be false if the test has passed. Logs the result via the
Rem test result manager.
Rem
Rem doc 2
Rem
Function AssertFalse(test As Boolean, Optional msg As String = "") As Boolean
On Error GoTo ErrorHandler

    If Not test Then
        If mTestResultManager.LogSuccess Then err.Raise knCall, , ksCall
    Else
        If mTestResultManager.LogFailure(msg) Then err.Raise knCall, , ksCall
    End If

Exit Function
ErrorHandler:
ErrTrap ksErrMod, "AssertFalse"
AssertFalse = True
End Function
'Public Sub AssertFalse(test As Boolean, Optional msg As String = "")
'
'    If Not test Then
'        mTestResultManager.LogSuccess
'    Else
'        mTestResultManager.LogFailure msg
'    End If
'
'End Sub

Rem =head4 Function ~
Rem sheetname AssertEqual
Rem doc 1
Rem =head3 AssertEqual
Rem doc 1 2
Rem
Rem AssertEqual(expected As Variant, actual As Variant, Optional msg As String = "") As Boolean
Rem
Rem rcl True
Rem
Rem Accepts two values that will be equal if the test has passed. In the case of floating point
Rem numbers, "Equal" is taken to mean that the ratio of the two is 1 if rounded to five decimal
Rem places. If this is not suitable for your application, you may need to write your own, more
Rem accurate test and use L<AssertSuccess|/"Function Assert AssertSuccess"> and
Rem L<AssertFailure|/"Function Assert AssertFailure"> to get the results you need. In this
Rem situation, you should know enough about floating point accuracy issues at least to find out
Rem more on them using a search engine. They are widely documented.
Rem
Rem doc 2
Rem
Function AssertEqual(expected As Variant, _
                     actual As Variant, _
                     Optional msg As String = "") As Boolean
On Error GoTo ErrorHandler

    If expected = actual Then
        If mTestResultManager.LogSuccess Then err.Raise knCall, , ksCall
    ElseIf actual = 0 Then
        If mTestResultManager.LogFailure("Expected '" & expected & _
                            "', got 0. " & msg) Then err.Raise knCall, , ksCall
    ElseIf IsNumeric(expected) _
           And IsNumeric(actual) Then
        If (VarType(expected) = vbSingle _
        Or VarType(expected) = vbDouble _
        Or VarType(actual) = vbSingle _
        Or VarType(actual) = vbDouble) _
        And Round(expected / actual, 5) = 1 Then
            If mTestResultManager.LogSuccess Then err.Raise knCall, , ksCall
        Else
            If mTestResultManager.LogFailure("Expected '" & expected & _
                            "', got '" & actual & "'. " & msg) Then err.Raise knCall, , ksCall
        End If
    Else
        If mTestResultManager.LogFailure("Expected '" & expected & _
                            "', got '" & actual & "'. " & msg) Then err.Raise knCall, , ksCall
    End If

Exit Function
ErrorHandler:
ErrTrap ksErrMod, "AssertEqual"
AssertEqual = True
End Function
'Public Sub AssertEqual(expected As Variant, actual As Variant, Optional msg As String = "")
'
'    If expected = actual Then
'        mTestResultManager.LogSuccess
'    Else
'        mTestResultManager.LogFailure "Expected '" & expected & "', got '" & actual & "'. " & msg
'    End If
'
'End Sub

Rem =head4 Function ~
Rem sheetname AssertSuccess
Rem doc 1
Rem =head3 AssertSuccess
Rem doc 1 2
Rem
Rem rcl True
Rem
Rem This function takes no parameters. It can be used if the previous "Assert" tests do not
Rem provide the needed functionality when a test has met the user's definition of success.
Rem
Rem doc 2
Rem
Function AssertSuccess() As Boolean
On Error GoTo ErrorHandler

    If mTestResultManager.LogSuccess Then err.Raise knCall, , ksCall

Exit Function
ErrorHandler:
ErrTrap ksErrMod, "AssertSuccess"
AssertSuccess = True
End Function
'Public Sub AssertSuccess()
'
'    mTestResultManager.LogSuccess
'
'End Sub

Rem =head4 Function ~
Rem sheetname AssertFailure
Rem doc 1
Rem =head3 AssertFailure
Rem doc 1 2
Rem
Rem AssertFailure(Optional msg As String = "") As Boolean
Rem
Rem rcl True
Rem
Rem This is the converse of AssertSuccess.
Rem
Rem doc 2
Rem
Function AssertFailure(Optional msg As String = "") As Boolean
On Error GoTo ErrorHandler

    If mTestResultManager.LogFailure("Failure. " & msg) Then err.Raise knCall, , ksCall

Exit Function
ErrorHandler:
ErrTrap ksErrMod, "AssertFailure"
AssertFailure = True
End Function
'Public Sub AssertFailure(Optional msg As String = "")
'
'    mTestResultManager.LogFailure "Failure. " & msg
'
'End Sub

Rem =head4 Function ~
Rem sheetname BackupTRM
Rem
Rem BackupTRM() As Boolean
Rem
Rem rcl True
Rem
Rem Creates a backup of the Test Result Manager. This is needed in self-testing, when testing
Rem the TRM handling procedures themselves. The "real" TRM needs to be overwritten to enable
Rem a fake one to be used for testing. However, the results of testing the fake need to be
Rem logged to the real TRM, otherwise the system reports zero passes and zero failures,
Rem regardless of the actual results. The backup is stored in a module level variable.
Rem
Function BackupTRM() As Boolean
On Error GoTo ErrorHandler

    Set trmBackup = mTestResultManager

Exit Function
ErrorHandler:
ErrTrap ksErrMod, "BackupTRM"
BackupTRM = True
End Function

Rem =head4 Function ~
Rem sheetname RestoreTRM
Rem
Rem RestoreTRM() As Boolean
Rem
Rem rcl True
Rem
Rem Reverses the polarity of the neutron flow^W^W^W BackupTRM above.
Rem
Function RestoreTRM() As Boolean
On Error GoTo ErrorHandler

    Set mTestResultManager = trmBackup
    Set trmBackup = Nothing

Exit Function
ErrorHandler:
ErrTrap ksErrMod, "RestoreTRM"
RestoreTRM = True
End Function
