Option Explicit
Public Sub Archive()
    MoveSelectedItemsToFolder "Archive", False
End Sub
Public Sub ArchiveAndMarkAsRead()
    MoveSelectedItemsToFolder "Archive", True
End Sub
Private Sub MoveSelectedItemsToFolder(FolderName As String, MarkAsRead As Boolean)
    On Error GoTo ErrorHandler
    
    Dim Namespace As Outlook.Namespace
    Set Namespace = Application.GetNamespace("MAPI")
    
    Dim Inbox As Outlook.MAPIFolder
    Set Inbox = Namespace.GetDefaultFolder(olFolderInbox)
    
    Dim Folder As Outlook.MAPIFolder
    Set Folder = Inbox.Folders(FolderName)
    If Folder Is Nothing Then
        MsgBox "The '" & FolderName & "' folder doesn't exist!", _
            vbOKOnly + vbExclamation, "Invalid Folder"
    End If
    Dim Message As Object
    For Each Message In Application.ActiveExplorer.Selection
        If MarkAsRead Then If Message.UnRead Then Message.UnRead = False
        If Message.Parent.Name = "Inbox" Then 'This button only works in the inbox folder
            Select Case (Message.SenderEmailAddress)
                Case "tfsservice@flagstonere.com", "TFServer@flagstonere.com":
                    Message.Move Inbox.Folders("Automated Messages").Folders("TFS Notice")
                    
                Case "quartzsupport@flagstonere.bm", "HPC-Help@flagstonere.com", "hfxpdtss004@flagstonere.com":
                    Message.Move Inbox.Folders("Automated Messages").Folders("Tickets")
            
                Case "paprd@hfxpdhpc017.flagstonere.local", "HFXPDANS002_DBMail@flagstonere.com", _
                     "HFXPRDMAS01_DBMail@flagstonere.com", "ANVPDMAS001_DBMail@flagstonere.com", _
                     "ANVPDANS001_DBMAIL@flagstonere.com", "SQLAdminHFX@flagstonere.bm":
                    Message.Move Inbox.Folders("Automated Messages").Folders("MOSAIC Support Junk")
            
                Case Else:
                    Message.Move Folder
            End Select
        End If
    Next Message
    
    Exit Sub
ErrorHandler:
    MsgBox Error(Err)
End Sub

Sub DeleteDuplicates()
    Dim i As Long, j As Long
    For i = 1 To Application.ActiveExplorer.Selection.Count
        For j = Application.ActiveExplorer.Selection.Count To i + 1 Step -1
            Dim Message As Object, Message2 As Object
            Set Message = Application.ActiveExplorer.Selection.Item(i)
            Set Message2 = Application.ActiveExplorer.Selection.Item(j)
            If Message = Message2 Then
                Message2.Delete
            End If
        Next j
    Next i
End Sub

'Expects the messages to be sorted by date received
Sub DeleteDuplicatesFast()
    Dim i As Long, j As Long
    For i = 1 To Application.ActiveExplorer.Selection.Count - 1
        If i >= Application.ActiveExplorer.Selection.Count Then Exit For
        If Application.ActiveExplorer.Selection.Item(i) = Application.ActiveExplorer.Selection.Item(i + 1) Then
            Application.ActiveExplorer.Selection.Item(i + 1).Delete
        End If
    Next i
End Sub

Sub DeleteEmptyMessages()
    Dim i As Long
    For i = Application.ActiveExplorer.Selection.Count - 1 To 1 Step -1
        Dim Message As Object
        Set Message = Application.ActiveExplorer.Selection.Item(i)
        If Message.Body = vbNullString Then
            Message.Delete
        End If
    Next i
End Sub
