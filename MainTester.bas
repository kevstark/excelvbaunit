Attribute VB_Name = "MainTester"
Option Explicit
Rem order 6.5
Rem
Rem =head2
Rem sheetname
Rem
Rem This is a retrofitted module to test Main. The code implies that one was created as part
Rem of someone's development, but it was not on the Google source repository. The main
Rem procedure, "xRun", cannot be tested.
Rem
Rem =head3
Rem sheetname Macros
Rem
Const ksErrMod As String = "MainTester"

Rem =head4 Function ~
Rem sheetname TestGetTestManager
Rem
Rem Everything JHD has tried has raised errors. Both the test routine - so far as it has been
Rem written - and the procedure in Main have been commented out.
Rem
'Function TestGetTestManager()
'On Error GoTo ErrorHandler
'
'    Dim tm As TestManager
'    Dim bTest As Boolean
'    bTest = tm Is Nothing
'    If AssertTrue(bTest, "Test Manager autoinstanciated") Then err.Raise knCall, , ksCall
'    tm = GetTestManager
'
'Exit Function
'ErrorHandler:
'ErrTrap ksErrMod, "TestGetTestManager"
'TestGetTestManager = True
'End Function

Rem =head4 Function ~
Rem sheetname TestSafeUbound
Rem
Rem 4 tests on 2 arrays and 2 scalars.
Rem
Function TestSafeUbound()
On Error GoTo ErrorHandler

    Dim test() As Long
    If AssertEqual(-1, SafeUbound(test()), "Returning a value for an undimensioned array") _
         Then err.Raise knCall, , ksCall
    If AssertEqual(-1, SafeUbound(""), "Returning a value for a string") _
         Then err.Raise knCall, , ksCall
    If AssertEqual(-1, SafeUbound(1), "Returning a value for an integer") _
         Then err.Raise knCall, , ksCall
    ReDim test(0 To 0)
    If AssertEqual(0, SafeUbound(test), "Can't handle 0 to 0") _
         Then err.Raise knCall, , ksCall

Exit Function
ErrorHandler:
ErrTrap ksErrMod, "TestSafeUbound"
TestSafeUbound = True
End Function

Rem =head4 Function ~
Rem sheetname TestGetNames
Rem
Rem 4 tests to set and reset the names.
Rem
Function TestGetNames()
On Error GoTo ErrorHandler
    
    Dim tmpName As String
    Dim tmpFileName As String
    tmpName = AddInName
    tmpFileName = AddInFileName
    AddInName = ""
    AddInFileName = ""
    If AssertEqual("", AddInName, "Can't clear addin name") Then err.Raise knCall, , ksCall
    If AssertEqual("", AddInFileName, "Can't clear addin filename") _
        Then err.Raise knCall, , ksCall
    If GetNames Then err.Raise knCall, , ksCall
    If AssertEqual(tmpName, AddInName, "Can't reproduce addin name") _
        Then err.Raise knCall, , ksCall
    If AssertEqual(tmpFileName, AddInFileName, "Can't reproduce addin filename") _
        Then err.Raise knCall, , ksCall
    AddInName = tmpName
    AddInFileName = tmpFileName
    
Exit Function
ErrorHandler:
ErrTrap ksErrMod, "TestGetNames"
TestGetNames = True
End Function
