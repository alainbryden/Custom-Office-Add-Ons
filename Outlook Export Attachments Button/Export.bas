Attribute VB_Name = "Export"
Option Explicit

Public Sub ExportAttachments()
    Dim objOL As Outlook.Application
    Dim objMsg As Object
    Dim objAttachments As Outlook.Attachments
    Dim objSelection As Outlook.Selection
    Dim i As Long, lngCount As Long
    Dim filesRemoved As String, fName As String, strFolder As String, saveFolder As String, savePath As String
    Dim alterEmails As Boolean, overwrite As Boolean
    Dim result
    
    saveFolder = BrowseForFolder("Select the folder to save attachments to.")
    If saveFolder = vbNullString Then Exit Sub
    
    result = MsgBox("Do you want to remove attachments from selected file(s)? " & vbNewLine & _
    "(Clicking no will export attachments but leave the emails alone)", vbYesNo + vbQuestion)
    alterEmails = (result = vbYes)
    
    Set objOL = CreateObject("Outlook.Application")
    Set objSelection = objOL.ActiveExplorer.Selection
    
    For Each objMsg In objSelection
        If objMsg.Class = olMail Then
            Set objAttachments = objMsg.Attachments
            lngCount = objAttachments.Count
            If lngCount > 0 Then
                filesRemoved = ""
                For i = lngCount To 1 Step -1
                    On Error GoTo errorSkipFile
                    fName = objAttachments.Item(i).FileName
                    savePath = saveFolder & "\" & fName
                    overwrite = False
                    While Dir(savePath) <> vbNullString And Not overwrite
                        Dim newFName As String
                        newFName = InputBox("The file '" & fName & _
                            "' already exists. Please enter a new file name, or just hit OK overwrite.", _
                            "Confirm File Name", fName)
                        If newFName = vbNullString Then GoTo skipfile
                        If newFName = fName Then overwrite = True Else fName = newFName
                        savePath = saveFolder & "\" & fName
                    Wend
                    
                    objAttachments.Item(i).SaveAsFile savePath
                    
                    If alterEmails Then
                        filesRemoved = filesRemoved & "<br>""" & objAttachments.Item(i).FileName & """ (" & _
                                                                formatSize(objAttachments.Item(i).size) & ") " & _
                            "<a href=""" & savePath & """>[Location Saved]</a>"
                        objAttachments.Item(i).Delete
                    End If
skipfile:
                Next i
                
                If alterEmails Then
                    filesRemoved = "<b>Attachments removed</b>: " & filesRemoved & "<br><br>"
                    
                    Dim objDoc As Object
                    Dim objInsp As Outlook.Inspector
                    Set objInsp = objMsg.GetInspector
                    Set objDoc = objInsp.WordEditor

                    objMsg.HTMLBody = filesRemoved + objMsg.HTMLBody
                    objMsg.Save
                End If
            End If
        End If
    Next
    
ExitSub:
    Set objAttachments = Nothing
    Set objMsg = Nothing
    Set objSelection = Nothing
    Set objOL = Nothing
    Exit Sub
errorSkipFile:
    Resume skipfile
End Sub

Function formatSize(size As Long) As String
    Dim val As Double, newVal As Double
    Dim unit As String
    
    val = size
    unit = "bytes"
    
    newVal = Round(val / 1024, 1)
    If newVal > 0 Then
        val = newVal
        unit = "KB"
    End If
    newVal = Round(val / 1024, 1)
    If newVal > 0 Then
        val = newVal
        unit = "MB"
    End If
    newVal = Round(val / 1024, 1)
    If newVal > 0 Then
        val = newVal
        unit = "GB"
    End If
    
    formatSize = val & " " & unit
End Function

'Function purpose:  To Browser for a user selected folder.
'If the "OpenAt" path is provided, open the browser at that directory
'NOTE:  If invalid, it will open at the Desktop level
Function BrowseForFolder(Optional Prompt As String, Optional OpenAt As Variant) As String
    Dim ShellApp As Object
    Set ShellApp = CreateObject("Shell.Application").BrowseForFolder(0, Prompt, 0, OpenAt)

    On Error Resume Next
    BrowseForFolder = ShellApp.self.Path
    On Error GoTo 0
    Set ShellApp = Nothing
     
    'Check for invalid or non-entries and send to the Invalid error handler if found
    'Valid selections can begin L: (where L is a letter) or \\ (as in \\servername\sharename.  All others are invalid
    Select Case Mid(BrowseForFolder, 2, 1)
        Case Is = ":": If Left(BrowseForFolder, 1) = ":" Then GoTo Invalid
        Case Is = "\": If Not Left(BrowseForFolder, 1) = "\" Then GoTo Invalid
        Case Else: GoTo Invalid
    End Select
     
    Exit Function
Invalid:
     'If it was determined that the selection was invalid, set to False
    BrowseForFolder = vbNullString
End Function

Function BrowseForFile(Optional Prompt As String, Optional OpenAt As Variant) As String
    Dim ShellApp As Object
    Set ShellApp = CreateObject("Shell.Application").BrowseForFolder(0, Prompt, 16 + 16384, OpenAt)
    
    On Error Resume Next
    BrowseForFile = ShellApp.self.Path
    On Error GoTo 0
    Set ShellApp = Nothing
     
    'Check for invalid or non-entries and send to the Invalid error handler if found
    'Valid selections can begin L: (where L is a letter) or \\ (as in \\servername\sharename.  All others are invalid
    Select Case Mid(BrowseForFolder, 2, 1)
        Case Is = ":": If Left(BrowseForFolder, 1) = ":" Then GoTo Invalid
        Case Is = "\": If Not Left(BrowseForFolder, 1) = "\" Then GoTo Invalid
        Case Else: GoTo Invalid
    End Select
     
    Exit Function
Invalid:
     'If it was determined that the selection was invalid, set to False
    BrowseForFile = vbNullString
End Function

