Attribute VB_Name = "ExportFunction"
Option Explicit
'Remember to add a reference to Microsoft Visual Basic for Applications Extensibility

'Exports all VBA project components containing code to a folder in the same directory as this spreadsheet.
Public Sub ExportAllComponents()
Attribute ExportAllComponents.VB_Description = "Exports all VBA project components containing code to a folder in the same directory as this spreadsheet."
Attribute ExportAllComponents.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim VBComp As VBIDE.VBComponent
    Dim destDir As String, fName As String, ext As String

    'Create the directory where code will be created.
    'Alternatively, you could change this so that the user is prompted
    If ActiveWorkbook.Path = "" Then
        MsgBox "You must first save this workbook somewhere so that it has a path.", , "Error"
        Exit Sub
    End If
    destDir = ActiveWorkbook.Path & "\" & ActiveWorkbook.Name & " Modules"
    If Dir(destDir, vbDirectory) = vbNullString Then MkDir destDir
    
    'Export all non-blank components to the directory
    For Each VBComp In ActiveWorkbook.VBProject.VBComponents
        If VBComp.CodeModule.CountOfLines > 0 Then
            'Determine the standard extention of the exported file.
            'These can be anything, but for re-importing, should be the following:
            Select Case VBComp.Type
                Case vbext_ct_ClassModule: ext = ".cls"
                Case vbext_ct_Document: ext = ".cls"
                Case vbext_ct_StdModule: ext = ".bas"
                Case vbext_ct_MSForm: ext = ".frm"
                Case Else: ext = vbNullString
            End Select
            If ext <> vbNullString Then
                fName = destDir & "\" & VBComp.Name & ext
                'Overwrite the existing file
                'Alternatively, you can prompt the user before killing the file.
                If Dir(fName, vbNormal) <> vbNullString Then Kill (fName)
                VBComp.Export (fName)
            End If
        End If
    Next VBComp
End Sub
