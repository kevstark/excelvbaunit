Attribute VB_Name = "Assert"
Option Explicit

Private mTestResultManager As ITestResultsManager

Public Sub SetTestResultsManager(manager As ITestResultsManager)

    Set mTestResultManager = manager
    
End Sub



Public Sub AssertTrue(test As Boolean, Optional msg As String = "")

    If test Then
        mTestResultManager.LogSuccess
    Else
        mTestResultManager.LogFailure msg
    End If

End Sub

Public Sub AssertFalse(test As Boolean, Optional msg As String = "")

    If Not test Then
        mTestResultManager.LogSuccess
    Else
        mTestResultManager.LogFailure msg
    End If

End Sub

Public Sub AssertEqual(expected As Variant, actual As Variant, Optional msg As String = "")

    If expected = actual Then
        mTestResultManager.LogSuccess
    Else
        mTestResultManager.LogFailure "Expected '" & expected & "', got '" & actual & "'. " & msg
    End If

End Sub

Public Sub AssertSuccess()

    mTestResultManager.LogSuccess

End Sub

Public Sub AssertFailure(Optional msg As String = "")

    mTestResultManager.LogFailure "Failure. " & msg

End Sub

