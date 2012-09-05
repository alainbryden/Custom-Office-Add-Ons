For a formatted version of this document with images, see the wiki.

 Intro

I've seen a few requests floating around for scripts that will export or strip attachments out of emails, and so far have never encountered any high quality implementations of this feature, so I created this method with a little flavor. Whether you just want an easy way to save all attachments somewhere on your computer, or you're doing a bulk reduction of your inbox size, this script should satisfy your needs.

All you need to do is select the emails you want to export or remove attachments from, and click your new "Export Attachments" button:

    Export Attachments

Export Attachments


Specify your file path:

    Specify File Path

Specify File Path


Choose whether or not you want to remove attachments from the original message(s):

    Option to remove attachments or just export

Option to remove attachments or just export


(Don't worry about overwriting same-named files, I have you covered):

    Prevents Unwanted Overwriting

Prevents Unwanted Overwriting


And presto! Your attachments have been saved to the desired location:

    Attachments Exported

Attachments Exported


And if you chose to strip attachments, your emails will now be conveniently prepended with the files it used to contain, their size, and a link to the path where they were saved:

    Attachments Stripped

Attachments Stripped




Here's how you make yourself an Export Attachments feature in Outlook

1
    Insert a New Module

    You're going to be making a VBA macro for going through messages and attachments. To do this, you need to open up the VBA project for your Outlook (by pressing Alt+F11). When it opens, your window will look similar to the one in the screen shot below. Then, right click your project on the left, and click Insert, Module, as shown:



2
    Paste in the Macro Code

    The hard work has all been done for you :)
    The code below takes care of all the features describe above (and more). Just copy paste it into the window, and you're almost ready to go! The code to copy in is below:

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


When you've pasted it in, everything should look like this. The code is pretty straightforward, so those of you who like to dig in should be able to understand and customize it at will:

    Code Pasted into VBA Project

Code Pasted into VBA Project



3
    Add a Button for Your Macro

    Now you can close down the Microsoft Visual Basic window. You've done the hard part. Next we want to add the macro to the toolbar so that you can use it conveniently. Right click on a blank area of the toolbar as shown, and click "Customize...". This will bring up the Outlook Customize Toolbars dialogue.



4
    Drag your Macro Onto the Toolbar

    You have to find your macro and drag it onto the Toolbar now. Switch to the "Commands" tab in the dialogue, and select "Macros" from the list on the left. You should see "ExportAttachments" macro there. Select it and drag it to wherever you want it on your toolbar:

    Add the Macro to the Toolbar

Add the Macro to the Toolbar



5
    Rename the Button

    You probably don't like the ugly name Outlook has given your button, so go ahead and rename it. To do this, click the "Rearrange Commands" button at the bottom of the dialogue we have open (I know - pretty unintuitive). A new dialogue will open. Click the "Toolbar" option button (instead of "Menu Bar") and find the new button you just created. When you find it, select it and click the "Modify Selection" button. Here, you can rename it to whatever you want. Here's an illustration:

    Rename the Button

Rename the Button



If you want, you can give your button a shortcut by inserting an ampersand (&) in front of the letter you want to be the shortcut. Then when you press Alt+'That Letter', it will trigger the button. Careful, if you chose a letter that is already a shortcut (like E for the 'Edit' menu item) then you'll have to press Alt+E+E(again) to cycle through to your button.


6
    Test it out!

    Now test it out! Select one or more emails, click the button, and save out your attachments. Isn't that sharp?

    Export Attachments

Export Attachments


    Attachments Stripped

Attachments Stripped




A few extra notes

    If you're about to overwrite a file, the dialogue pops up to confirm the overwrite or let you change the file name. If you keep changing the name to that of a file that already exists, it will keep prompting you  :)

    The code puts a nicely formatted message at the top of your emails, including all the names of removed files, and their sizes, and even the path where they were saved (as a url). You can customize this message if you edit the code.

    If you move the file after you've saved it, the link in the emails doesn't get updated (unless you do it manually), so be aware if you plan to make use of this feature to track down attachments in the future - save it where you want to store it!

    If you click cancel in the initial folder chooser dialogue, the whole routine will end. If you click cancel when picking a new name for a file that already exists, it will just skip that file and leave it attached to the email.

    If you want a button that will always strip attachments, or one that will always just save all attachments, but leave the email as is, you can easily modify this macro to not prompt you and always perform the desired action.
    For instance, replacing the lines 

result = MsgBox("Do you want to remove attachments from selected file(s)? " & vbNewLine & _
    "(Clicking no will export attachments but leave the emails alone)", vbYesNo + vbQuestion)
alterEmails = (result = vbYes)

       

with just

alterEmails = false

                                    
will cause the macro to just save attachments but not strip them out.


One other thing you might want to consider is changing the icon of your button. I'm sure you've noticed that that nice icon on the first page doesn't appear on its own. I created that in the same place as I renamed the button. After clicking "Modify Selection", click "Modify button icon..." and you will get a window where you can specify your own icon. This is how I designed mine:

    Design your own Button Icon

Design your own Button Icon



