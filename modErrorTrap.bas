Attribute VB_Name = "modErrorTrap"
Option Explicit
Option Private Module
Rem order 20
Rem
Rem =head2
Rem sheetname
Rem
Rem This code module contains three public constants, two public variables and one procedure.
Rem The public variables are usually populated by an initialisation routine, but are not used
Rem outside this module.
Rem
Rem It should be possible to use the module in almost any file. It is possible to rewrite it
Rem so that the two global variables are passed as parameters, but as they are unlikely to
Rem change, it is felt that globals are better as they avoid the repetition of passing the
Rem same parameters from every procedure. For this application, the globals have been
Rem replaced with constants.
Rem
Rem =head3
Rem sheetname Macros
Rem

Public Const knCall As Long = 9999
Public Const ksCall As String = "Call to previous error"
Const ksReportTo As String = "http://code.google.com/p/excelvbaunit/issues/list"
Const ksLogFile As String = ""
'Public sLogFile As String
'Public sReportTo As String
Public Const bFullTrace As Boolean = False

Rem
Rem =head4 Sub ErrTrap
Rem
Rem This is the central error handling routine. Every procedure in the file has error traps
Rem that point to it. It accepts two to four parameters. The first two, compulsory, parameters
Rem form the name of the procedure where the error occurred. The third is a flag indicating
Rem whether to log errors to a file. It is a boolean, so if it is missing, it will
Rem default to FALSE. The fourth is a flag indicating whether to show message boxes when
Rem errors occur. The definition includes the default of "True", as showing message boxes
Rem is strongly recommended.
Rem
Rem The procedure's first action is to record the
Rem state of the error object. Once this has been done, C<On Error Resume Next> is used to
Rem prevent any infinite recursive calls to the error trap. The C<On Error Resume Next>
Rem statement has the potential to reset the error object, so it cannot be invoked until
Rem the contents of the error object have been recorded.
Rem
Rem At last, the processing of the error can begin. Errors can be reported in either or both
Rem of two ways. They can be logged in a file and a messagebox can be shown. The first test
Rem is for the log file. The boolean parameter is tested, but the length of the file name
Rem is also tested to make sure that it is non-zero, i.e. that there is a file name. It is
Rem the developer's responsibility to make sure that the file name is kosher and that it
Rem will not conflict with any existing file. No checking of this has been done - at some
Rem level, the developer has to be trusted to get things right. A line of text is appended to
Rem the file, which is then closed.
Rem
Rem The next section of code deals with messageboxes. The box contains four lines:
Rem
Rem =over
Rem
Rem =item 1 The error number
Rem
Rem =item 2 The VBA description of the error
Rem
Rem =item 3 The routine in which the error occurred
Rem
Rem =item 4 The name or function to whom to report the error
Rem
Rem =back
Rem

Sub ErrTrap(sMod As String, _
            ByVal sProc As String, _
            Optional bLog As Boolean, _
            Optional bMsg As Boolean = True)

Dim nErrNo As Long
Dim sErrDesc As String
nErrNo = err.Number
sErrDesc = err.Description
On Error Resume Next
sProc = sMod & "." & sProc

If bLog And Len(ksLogFile) > 0 Then
    Dim nFile As Long
    nFile = FreeFile()
    Open ThisWorkbook.Path & "\" & ksLogFile For Append As #nFile
    Print #nFile, Format$(Now(), "dd mmm yy hh:mm:ss") & " " & _
                  ThisWorkbook.name & " " & "Error " & nErrNo & _
                  sErrDesc & vbCrLf & " occurred in " & sProc
    Close #nFile
End If 'bLog And Len(sLogFile) > 0

If bMsg And (nErrNo <> knCall Or bFullTrace) Then
    MsgBox "Error " & nErrNo & vbCrLf & sErrDesc & vbCrLf & _
           "occurred in " & sProc & vbCrLf & "Please report to " & ksReportTo & "." & vbCrLf & _
           "The report should include a log file output as well as steps to reproduce", _
           vbCritical, ThisWorkbook.name & " aborting"
End If 'bMsg

End Sub
