Attribute VB_Name = "Export"
Option Explicit

Public Sub ExportThisCode(dir As String)
    ExportAllCode dir, "VbaUnit"
End Sub

' Functions to export and remove vba code from an add in.
Public Sub ExportAllCode(dir As String, projectName As String)

    Dim components As VBComponents
    Set components = Application.VBE.VBProjects("VbaUnit").VBComponents
        
    Dim c As VBComponent
    For Each c In components
        If IsCodeModule(c) Then
            c.Export GetExportFileName(dir, c)
        End If
    Next
    
End Sub

Public Sub DeleteAllCode(projectName As String)

    Dim components As VBComponents
    Set components = Application.VBE.VBProjects("VbaUnit").VBComponents
    
    Dim CodeModules() As VBComponent
    CodeModules = GetCodeModules(components)
    
    Dim i As Integer
    For i = 0 To UBound(CodeModules)
        components.Remove CodeModules(i)
    Next

End Sub



Private Function GetCodeModules(components As VBComponents) As VBComponent()

    Dim count As Integer
    count = CountCodeModules(components)
    
    ReDim cs(0 To count - 1) As VBComponent
    Dim i As Integer
    i = 0
    Dim c As VBComponent
    For Each c In components
        If IsCodeModule(c) Then
            Set cs(i) = c
            i = i + 1
        End If
    Next


End Function

Private Function CountCodeModules(components As VBComponents) As Integer

    Dim count As Integer
    count = 0
    
    Dim c As VBComponent
    For Each c In components
        If IsCodeModule(c) Then
            count = count + 1
        End If
    Next

    CountCodeModules = count

End Function



Private Function GetExportFileName(dir As String, c As VBComponent) As String

    If Right(dir, 1) <> "\" Then
        dir = dir & "\"
    End If
    
    GetExportFileName = dir & c.name & GetExportFileNameSuffix(c)

End Function

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
    End Select
    
End Function

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
    End Select
    
End Function

