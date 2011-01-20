Attribute VB_Name = "Export"
Option Explicit
Rem order 5
Rem
Rem =head2
Rem sheetname
Rem
Rem This standard code module contains procedures to delete all code modules from an add-in
Rem named AddInName, i.e. itself. It is unlikely to be useful to many people. It is not
Rem tested by the self-testing process since it deletes the code that would be tested and
Rem does the testing.
Rem
Rem =head3
Rem sheetname Macros
Rem
Const ksErrMod As String = "Export"

' Functions to export and remove vba code from an add in.
Rem =head4 Function ~
Rem sheetname ExportAllCode
Rem
Rem ExportAllCode(dir As String, Optional projectName As String) As Boolean
Rem
Rem rcl True
Rem
Rem Originally, this was called by ExportThisCode, which has now been deleted. It has been
Rem refactored by means of the AddInName variable to enable both its original purpose (to
Rem export code from any named project) and the purpose of ExportThisCode (to export the
Rem add-in). If a project is supplied, it is exported. If not, the add-in is exported. A
Rem directory, to which the code will be exported, must be specified.
Rem
Function ExportAllCode(dir As String, Optional projectName As String) As Boolean
On Error GoTo ErrorHandler

    If GetNames Then err.Raise knCall, , ksCall
    If projectName = "" Then projectName = AddInName
    
    Dim components As VBComponents
    Set components = Application.VBE.VBProjects(projectName).VBComponents
        
    Dim c As VBComponent
    For Each c In components
        If IsCodeModule(c) Then
            c.Export GetExportFileName(dir, c)
        End If
    Next c

Exit Function
ErrorHandler:
ErrTrap ksErrMod, "ExportAllCode"
ExportAllCode = True
End Function
'Public Sub ExportAllCode(dir As String, Optional projectName As String)
'
'    If GetNames Then err.Raise knCall, , ksCall
'    If projectName = "" Then projectName = AddInName
'
'    Dim components As VBComponents
'    Set components = Application.VBE.VBProjects(projectName).VBComponents
'
'    Dim c As VBComponent
'    For Each c In components
'        If IsCodeModule(c) Then
'            c.Export GetExportFileName(dir, c)
'        End If
'    Next c
'
'End Sub

Rem =head4 Function ~
Rem sheetname DeleteAllCode
Rem
Rem DeleteAllCode(Optional projectName As String) As Boolean
Rem
Rem rcl True
Rem
Rem This procedure has been refactored to take the add-in name from the initialisation routine
Rem if it is missing. It must be called from the Immediate pane and will check before it
Rem deletes code. God speed, brave programmer. I hope you backed everything up before you
Rem mistyped or forgot the project name.
Rem
Function DeleteAllCode(Optional projectName As String) As Boolean
On Error GoTo ErrorHandler

    If GetNames Then err.Raise knCall, , ksCall
    If projectName = "" Then projectName = AddInName
 
    Dim components As VBComponents
    Set components = Application.VBE.VBProjects(projectName).VBComponents
    
    Dim CodeModules() As VBComponent
    CodeModules = GetCodeModules(components)
    
    MsgBox "About to delete all code from " & projectName & ". Do you really want to?", _
    vbYesNo, "Last appeal to sanity"
    
    Dim i As Integer
    For i = 0 To UBound(CodeModules)
        components.Remove CodeModules(i)
    Next i

Exit Function
ErrorHandler:
ErrTrap ksErrMod, "DeleteAllCode"
DeleteAllCode = True
End Function
'Public Sub DeleteAllCode(Optional projectName As String)
'
'    If GetNames Then err.Raise knCall, , ksCall
'    If projectName = "" Then projectName = AddInName
'
'    Dim components As VBComponents
'    Set components = Application.VBE.VBProjects(projectName).VBComponents
'
'    Dim CodeModules() As VBComponent
'    CodeModules = GetCodeModules(components)
'
'    MsgBox "About to delete all code from " & projectName & ". Do you really want to?", _
'    vbYesNo, "Last appeal to sanity"
'
'    Dim i As Integer
'    For i = 0 To UBound(CodeModules)
'        components.Remove CodeModules(i)
'    Next i
'
'End Sub

'Rem =head4 Function GetCodeModules
'Rem
'Rem Refactored to obviate the need for CountCodeModules.
'Rem
'Private Function GetCodeModules(components As VBComponents) As VBComponent()
'
'    Dim count As Integer
'    count = CountCodeModules(components)
'
'    ReDim cs(0 To count - 1) As VBComponent
'    Dim i As Integer
'    i = 0
'    Dim c As VBComponent
'    For Each c In components
'        If IsCodeModule(c) Then
'            Set cs(i) = c
'            i = i + 1
'        End If
'    Next c
'
'
'End Function
Rem =head4 Function ~
Rem sheetname GetCodeModules
Rem
Rem GetCodeModules(components As VBComponents) As VBComponent()
Rem
Rem Returns an array of code modules from the components collection passed.
Rem
Function GetCodeModules(components As VBComponents) As VBComponent()
On Error GoTo ErrorHandler

    Dim comp As VBComponent
    Dim count As Long
    count = -1
    Dim compReturn() As VBComponent
    For Each comp In components
        If IsCodeModule(comp) Then
            count = count + 1
            ReDim Preserve compReturn(0 To count)
            Set compReturn(count) = comp
        End If
    Next comp
    
    GetCodeModules = compReturn

Exit Function
ErrorHandler:
ErrTrap ksErrMod, "GetCodeModules"
End Function

'Rem =head4 Function CountCodeModules
'Rem
'Rem This seems to be redundant. It is called only from GetCodeModules, which repeats the loop
'Rem and call to IsCodeModule. It can be refactored out by using ReDim Preserve.
'Rem
'Private Function CountCodeModules(components As VBComponents) As Integer
'
'    Dim count As Integer
'    count = 0
'
'    Dim c As VBComponent
'    For Each c In components
'        If IsCodeModule(c) Then
'            count = count + 1
'        End If
'    Next c
'
'    CountCodeModules = count
'
'End Function

Rem =head4 Function ~
Rem sheetname GetExportFileName
Rem
Rem GetExportFileName(dir As String, comp As VBComponent) As String
Rem
Rem Takes a directory and a component and returns the fully qualified file name with the
Rem necessary suffix.
Rem
Private Function GetExportFileName(dir As String, comp As VBComponent) As String

    If Right(dir, 1) <> "\" Then
        dir = dir & "\"
    End If

    GetExportFileName = dir & comp.name & GetExportFileNameSuffix(comp)

End Function

Rem =head4 Function ~
Rem sheetname GetExportFileNameSuffix
Rem
Rem GetExportFileNameSuffix(c As VBComponent) As String
Rem
Rem Given a component, returns a suffix (or conceivably three question marks if no definitive
Rem suffix is available).
Rem
Private Function GetExportFileNameSuffix(c As VBComponent) As String

    Select Case c.Type
        Case vbext_ct_ActiveXDesigner
            GetExportFileNameSuffix = ".???"
        Case vbext_ct_ClassModule
            GetExportFileNameSuffix = ".cls"
        Case vbext_ct_Document
            GetExportFileNameSuffix = ".???"
        Case vbext_ct_MSForm
            GetExportFileNameSuffix = ".???"
        Case vbext_ct_StdModule
            GetExportFileNameSuffix = ".bas"
        Case Else
            err.Raise knCall - 1, , "Unknown component type"
    End Select
    
End Function

Rem =head4 Function ~
Rem sheetname IsCodeModule
Rem
Rem Given a component, returns a boolean indicating whether it is a code module.
Rem
Private Function IsCodeModule(c As VBComponent) As Boolean

    Select Case c.Type
        Case vbext_ct_ActiveXDesigner
            IsCodeModule = False
        Case vbext_ct_ClassModule
            IsCodeModule = True
        Case vbext_ct_Document
            IsCodeModule = False
        Case vbext_ct_MSForm
            IsCodeModule = False
        Case vbext_ct_StdModule
            IsCodeModule = True
        Case Else
            err.Raise knCall - 1, , "Unknown component type"
    End Select
    
End Function
